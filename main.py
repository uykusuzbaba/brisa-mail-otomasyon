"""
Brisa Mail Otomasyon — v24 FINAL
Gmail IMAP (App Password) + Sheets Service Account = Sonsuz token

Tüm özellikler dahil:
  - Türkçe karakter normalizasyonu
  - Rakip firma listesi (CABOT, MITSUI + SOUTHLAND/GTR/IOI/RHODIA/SAPH vb.)
  - Yükleyici/Gönderici Brisa kontrolü
  - Virgüllü konşimento parse + House AWB/HAWB
  - _norm() ile eşleştirme normalizasyonu
  - Mükerrer kayıt kontrolü (Sheets K kolonu)
  - page.inner_text("body") — düz metin (HTML değil)
  - Bekleyenler için 3 aşamalı eleme
  - sheets_bekleyeni_aitdegil_guncelle()
  - Kullanıcı (L kolonu) bilgisi
  - Sheets retry mekanizması (503/429)
  - fatura_url bekleyenlerden taşınıyor
  - YENİ v24: Bize ait değil çapraz eleme
    (Ait değil faturanın konşimento/konteyner/beyanname numarası
     bekleyen faturalarda da geçiyorsa → otomatik ait değil yap)
"""

import email
import imaplib
import json
import logging
import os
import re
import sqlite3
import time
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from playwright.sync_api import sync_playwright

# ── Loglama ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Ayarlar (GitHub Secrets'tan gelir) ───────────────────────
GOOGLE_API_KEY       = os.environ.get("GOOGLE_API_KEY", "")
GMAIL_APP_PASSWORD   = os.environ["GMAIL_APP_PASSWORD"]
GMAIL_EMAIL          = os.environ.get("GMAIL_EMAIL", "ramsesium.md@gmail.com")
SERVICE_ACCOUNT_JSON = os.environ["SERVICE_ACCOUNT_JSON"]
SHEETS_ID            = os.environ["SHEETS_ID"]
BILDIRIM_ALICISI     = os.environ["BILDIRIM_ALICISI"]
GONDEREN_LISTESI     = os.environ.get("GONDEREN_LISTESI", "m.hizmet@brisa.com.tr").split(",")
SORGU_PENCERESI_GUN  = int(os.environ.get("SORGU_PENCERESI_GUN", "60"))

# ── Sabitler ─────────────────────────────────────────────────
DB_PATH = Path("brisa_mail.db")
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

RAKIP_FIRMALAR = [
    # Önceki firmalar
    "CABOT", "MITSUI", "TANAKA", "ZEON CHEMICAL",
    "KOLON", "THAI TOKAI", "HS HYOSUNG", "PYRAMID",
    # v24'te eklenen firmalar
    "SOUTHLAND",          # SOUTHLAND KATI COTE D'IVOIRE (SKCI) ve SOUTHLAND RUBBER CO.
    "SKCI",               # SOUTHLAND KATI kısa adı
    "G T RUBBER",         # G T RUBBER CO.,LTD
    "IOI ACIDCHEM",       # IOI ACIDCHEM SDN. BHD.
    "RHODIA",             # RHODIA OPERATIONS
    "SAPH",               # SOCIETE AFRICAINE DE PLANTATIONS D'HEVEAS (her iki yazımı kapsar)
    "SOCIETE AFRICAINE",  # Tam adıyla gelenler için
]


# ════════════════════════════════════════════════════════════
#  VERİTABANI
# ════════════════════════════════════════════════════════════

def db_init():
    conn = sqlite3.connect(DB_PATH)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS processed (
            email_id TEXT PRIMARY KEY,
            dosya_no TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS pending (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email_id TEXT UNIQUE,
            gonderen TEXT,
            konu TEXT,
            tarih TEXT,
            numaralar TEXT,
            retry_count INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now'))
        );
    """)
    conn.commit()
    return conn


def db_islendi_mi(conn, email_id):
    r = conn.execute("SELECT 1 FROM processed WHERE email_id=?", (email_id,)).fetchone()
    if r:
        return True
    r = conn.execute("SELECT 1 FROM pending WHERE email_id=?", (email_id,)).fetchone()
    return r is not None


def db_islendi_ekle(conn, email_id, dosya_no):
    conn.execute(
        "INSERT OR IGNORE INTO processed (email_id, dosya_no) VALUES (?,?)",
        (email_id, dosya_no)
    )
    conn.commit()


def db_bekleyen_ekle(conn, email_id, gonderen, konu, tarih, numaralar):
    conn.execute(
        """INSERT OR IGNORE INTO pending
           (email_id, gonderen, konu, tarih, numaralar)
           VALUES (?,?,?,?,?)""",
        (email_id, gonderen, konu, tarih, json.dumps(numaralar))
    )
    conn.commit()


# ════════════════════════════════════════════════════════════
#  GMAIL IMAP (App Password — sonsuz token)
# ════════════════════════════════════════════════════════════

def gmail_imap_baglanti():
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
        log.info(f"✅ IMAP bağlantısı başarılı: {GMAIL_EMAIL}")
        return mail
    except Exception as e:
        log.error(f"❌ IMAP bağlantı hatası: {e}")
        raise


def gmail_okunmamis_mailler_imap():
    mail = gmail_imap_baglanti()
    mail.select("inbox")
    detaylar = []

    for gonderen in GONDEREN_LISTESI:
        try:
            status, messages = mail.search(None, 'UNSEEN', f'FROM "{gonderen.strip()}"')
            if status != "OK":
                continue

            msg_nums = messages[0].split()
            log.info(f"📬 {gonderen} → {len(msg_nums)} okunmamış mail")

            for num in msg_nums:
                try:
                    status, msg_data = mail.fetch(num, "(RFC822)")
                    if status != "OK":
                        continue

                    msg = email.message_from_bytes(msg_data[0][1])

                    subject_header = msg["Subject"] or ""
                    decoded = decode_header(subject_header)
                    subject = ""
                    for content, encoding in decoded:
                        if isinstance(content, bytes):
                            subject += content.decode(encoding or "utf-8", errors="replace")
                        else:
                            subject += content

                    html_body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/html":
                                payload = part.get_payload(decode=True)
                                if payload:
                                    html_body = payload.decode(
                                        part.get_content_charset() or "utf-8", errors="replace"
                                    )
                                    break
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload:
                            html_body = payload.decode(
                                msg.get_content_charset() or "utf-8", errors="replace"
                            )

                    detaylar.append({
                        "id":        msg["Message-ID"] or num.decode(),
                        "imap_num":  num.decode(),
                        "gonderen":  msg["From"] or "",
                        "konu":      subject,
                        "tarih":     msg["Date"] or "",
                        "html_body": html_body,
                    })

                except Exception as e:
                    log.warning(f"Mail parse hatası (#{num}): {e}")

        except Exception as e:
            log.error(f"Gönderen search hatası ({gonderen}): {e}")

    mail.close()
    mail.logout()
    log.info(f"📨 Toplam {len(detaylar)} okunmamış mail bulundu")
    return detaylar


def gmail_okundu_isaretle_imap(imap_num):
    try:
        mail = gmail_imap_baglanti()
        mail.select("inbox")
        mail.store(imap_num.encode(), '+FLAGS', '\\Seen')
        mail.close()
        mail.logout()
    except Exception as e:
        log.warning(f"IMAP okundu işareti hatası: {e}")


# ════════════════════════════════════════════════════════════
#  SHEETS (Service Account — sonsuz token)
# ════════════════════════════════════════════════════════════

def _sheets_creds():
    return service_account.Credentials.from_service_account_info(
        json.loads(SERVICE_ACCOUNT_JSON), scopes=SHEETS_SCOPES
    )


def sheets_servis():
    return build("sheets", "v4", credentials=_sheets_creds())


def sheets_referans_veri():
    """
    📋 Referans sekmesi: A=Konşimento, B=Konteyner, C=Beyanname, D=Dosya No, E=Kullanıcı
    503/429 hatalarında 3 kez retry.
    """
    servis = sheets_servis()

    for deneme in range(3):
        try:
            result = servis.spreadsheets().values().get(
                spreadsheetId=SHEETS_ID,
                range="📋 Referans!A:E"
            ).execute()
            break
        except HttpError as e:
            if e.resp.status in [503, 429]:
                bekleme = 5 * (deneme + 1)
                log.warning(f"Sheets geçici hata ({e.resp.status}) — {bekleme}s ({deneme+1}/3)")
                if deneme < 2:
                    time.sleep(bekleme)
                    continue
            log.error(f"Sheets okuma hatası: {e}")
            return []
    else:
        log.error("Sheets 3 denemede okunamadı")
        return []

    rows = result.get("values", [])
    if len(rows) < 2:
        return []

    kayitlar = []
    for row in rows[1:]:
        padded = row + [""] * (5 - len(row))
        dosya = str(padded[3]).strip() if padded[3] else None
        if not dosya:
            continue

        kayitlar.append({
            "dosya_no":           dosya,
            "konsimento_listesi": [_norm(k) for k in str(padded[0]).split(",") if k.strip()],
            "konteyner_listesi":  [_norm(k) for k in str(padded[1]).split(",") if k.strip()],
            "beyanname_listesi":  [_norm(k) for k in str(padded[2]).split(",") if k.strip()],
            "kullanici":          str(padded[4]).strip() if padded[4] else "",
        })

    log.info(f"📋 {len(kayitlar)} referans kaydı yüklendi")
    return kayitlar


def sheets_fatura_islendi_mi(servis, email_id):
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID, range="📑 Fatura Listesi!K:K"
        ).execute()
        for row in result.get("values", []):
            if row and str(row[0]).strip() == str(email_id).strip():
                return True
    except HttpError:
        pass
    return False


def sheets_faturaListesiYaz(gonderen, tarih, numaralar, durum,
                             dosya_no="", fatura_url="", email_id="", kullanici=""):
    """
    A=Geliş, B=Mail Tarihi, C=Gönderen, D=Konşimento, E=Konteyner,
    F=Beyanname, G=Durum, H=Dosya No, I=Fatura Linki,
    J=İşlemi Yapan, K=Email ID, L=Kullanıcı
    """
    try:
        servis = sheets_servis()

        if email_id and sheets_fatura_islendi_mi(servis, email_id):
            log.info(f"Fatura listesinde zaten var, atlanıyor: {email_id}")
            return

        konsimento = ", ".join(numaralar.get("konsimento_list", [])) if numaralar else ""
        konteyner  = ", ".join(numaralar.get("konteyner_list",  [])) if numaralar else ""
        beyanname  = ", ".join(numaralar.get("beyanname_list",  [])) if numaralar else ""

        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A:L",
            valueInputOption="RAW",
            body={"values": [[
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                tarih, gonderen, konsimento, konteyner, beyanname,
                durum, dosya_no, fatura_url, "", email_id, kullanici,
            ]]}
        ).execute()
        log.info(f"✅ Fatura Listesi güncellendi: {durum}")

    except HttpError as e:
        log.error(f"Sheets yazma hatası: {e}")


def sheets_bekleyenleri_getir(servis):
    """
    🟡 Bekliyor satırlarını oku. fatura_url da dahil.
    503/429 hatalarında retry.
    """
    for deneme in range(3):
        try:
            result = servis.spreadsheets().values().get(
                spreadsheetId=SHEETS_ID,
                range="📑 Fatura Listesi!A:L"
            ).execute()
            break
        except HttpError as e:
            if e.resp.status in [503, 429]:
                bekleme = 5 * (deneme + 1)
                log.warning(f"Bekleyen okuma geçici hata — {bekleme}s ({deneme+1}/3)")
                if deneme < 2:
                    time.sleep(bekleme)
                    continue
            log.error(f"Bekleyen okuma hatası: {e}")
            return []
    else:
        return []

    rows = result.get("values", [])
    if len(rows) < 2:
        return []

    bekleyenler = []
    for i, row in enumerate(rows[1:], start=2):
        padded = row + [""] * (12 - len(row))
        durum_hucre = padded[6].strip()

        if "Bekliyor" not in durum_hucre and "🟡" not in durum_hucre:
            continue

        numaralar = {
            "konsimento_list": [v.strip() for v in padded[3].split(",") if v.strip()],
            "konteyner_list":  [v.strip() for v in padded[4].split(",") if v.strip()],
            "beyanname_list":  [v.strip() for v in padded[5].split(",") if v.strip()],
        }
        if not any(numaralar.values()):
            continue

        bekleyenler.append({
            "satir_no":   i,
            "gonderen":   padded[2].strip(),
            "tarih":      padded[1].strip(),
            "email_id":   padded[10].strip(),
            "fatura_url": padded[8].strip(),
            "numaralar":  numaralar,
        })

    log.info(f"🟡 {len(bekleyenler)} bekleyen fatura bulundu")
    return bekleyenler


def sheets_bekleyeni_guncelle(servis, satir_no, dosya_no, kriter, deger):
    try:
        servis.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"📑 Fatura Listesi!G{satir_no}:H{satir_no}",
            valueInputOption="RAW",
            body={"values": [["✅ Eşleşti", dosya_no]]}
        ).execute()
        log.info(f"Satır {satir_no} güncellendi → {dosya_no} ({kriter}: {deger})")
    except HttpError as e:
        log.error(f"Bekleyen güncelleme hatası (satır {satir_no}): {e}")


def sheets_bekleyeni_aitdegil_guncelle(servis, satir_no, sebep):
    try:
        servis.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"📑 Fatura Listesi!G{satir_no}",
            valueInputOption="RAW",
            body={"values": [[f"🔴 Bize Ait Değil | {sebep}"]]}
        ).execute()
        log.info(f"Satır {satir_no} → Bize Ait Değil ({sebep})")
    except HttpError as e:
        log.error(f"Bekleyen eleme hatası (satır {satir_no}): {e}")


def sheets_ait_degil_numaralari_getir(servis):
    """
    Fatura Listesi'nden 'Bize Ait Değil' satırlarının konşimento,
    konteyner ve beyanname numaralarını okur.
    Bekleyenleri çapraz eleme için kullanılır.
    Döndürür: {
        "konsimento": {norm_no, norm_no, ...},
        "konteyner":  {norm_no, ...},
        "beyanname":  {son6_rakam, ...}
    }
    """
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A:L"
        ).execute()
    except HttpError as e:
        log.error(f"Ait değil numaraları okuma hatası: {e}")
        return {"konsimento": set(), "konteyner": set(), "beyanname": set()}

    rows = result.get("values", [])
    ait_degil = {"konsimento": set(), "konteyner": set(), "beyanname": set()}

    for row in rows[1:]:
        padded = row + [""] * (12 - len(row))
        durum = padded[6].strip()
        if "Bize Ait Değil" not in durum and "🔴" not in durum:
            continue

        # Konşimento (D kolonu)
        for k in padded[3].split(","):
            k = k.strip()
            if k:
                ait_degil["konsimento"].add(_norm(k))

        # Konteyner (E kolonu)
        for k in padded[4].split(","):
            k = k.strip()
            if k:
                ait_degil["konteyner"].add(_norm(k))

        # Beyanname (F kolonu) — son 6 rakam
        for b in padded[5].split(","):
            b = b.strip()
            if b:
                son6 = ''.join(c for c in b if c.isdigit())[-6:]
                if son6:
                    ait_degil["beyanname"].add(son6)

    log.info(
        f"🔴 Ait değil havuzu: "
        f"Konşimento={len(ait_degil['konsimento'])} "
        f"Konteyner={len(ait_degil['konteyner'])} "
        f"Beyanname={len(ait_degil['beyanname'])}"
    )
    return ait_degil


def capraz_eleme_kontrol(numaralar, ait_degil_havuzu):
    """
    Bekleyen faturanın numaralarını, daha önce 'Bize Ait Değil' olarak
    işaretlenmiş faturaların numara havuzuyla karşılaştırır.
    Eşleşme varsa (False, sebep) döner.
    """
    for k in numaralar.get("konsimento_list", []):
        if k and _norm(k) in ait_degil_havuzu["konsimento"]:
            return False, f"Konşimento çapraz eşleşme: {k}"

    for k in numaralar.get("konteyner_list", []):
        if k and _norm(k) in ait_degil_havuzu["konteyner"]:
            return False, f"Konteyner çapraz eşleşme: {k}"

    for b in numaralar.get("beyanname_list", []):
        son6 = ''.join(c for c in b if c.isdigit())[-6:]
        if son6 and son6 in ait_degil_havuzu["beyanname"]:
            return False, f"Beyanname çapraz eşleşme: {b}"

    return True, ""


# ════════════════════════════════════════════════════════════
#  PLAYWRIGHT
# ════════════════════════════════════════════════════════════

def sayfayi_playwright_ile_oku(url):
    """inner_text("body") — düz metin, HTML değil."""
    log.info(f"Playwright ile sayfa açılıyor: {url[:80]}...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"]
            )
            page = browser.new_page()
            page.goto(url, wait_until="networkidle", timeout=30000)
            page.wait_for_timeout(3000)
            icerik = page.inner_text("body")
            browser.close()
            log.info(f"Sayfa okundu: {len(icerik)} karakter")
            return icerik[:15000]
    except Exception as e:
        log.error(f"Playwright hatası: {e}")
        return None


# ════════════════════════════════════════════════════════════
#  YARDIMCI FONKSİYONLAR
# ════════════════════════════════════════════════════════════

def turkce_normalize(metin):
    if not metin:
        return ""
    metin = metin.upper()
    for tr, eng in {'Ç':'C','Ğ':'G','İ':'I','Ö':'O','Ş':'S','Ü':'U',
                    'ç':'C','ğ':'G','ı':'I','i':'I','ö':'O','ş':'S','ü':'U'}.items():
        metin = metin.replace(tr, eng)
    return metin


def _norm(v):
    if not v:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(v).upper())


def linkten_url_bul(html_govde):
    match = re.search(r'https?://[^\s"\'<>]*edoksis[^\s"\'<>]*', html_govde, re.IGNORECASE)
    if match:
        url = match.group(0).replace("&amp;", "&")
        log.info(f"🔗 Link bulundu: {url[:80]}...")
        return url
    return None


def firma_adi_cek(icerik):
    if not icerik:
        return ""
    try:
        match = re.search(r'G[\xf6o]nderen[^\n]{0,5}([A-Z][A-Z\s\.&,]{5,79})', icerik, re.IGNORECASE)
        if match:
            firma = match.group(1).strip()
            if not re.match(r'^[\d\-]+$', firma):
                return firma[:80].strip()
    except Exception:
        pass
    return ""


def bize_ait_mi_kontrol(icerik, html_govde):
    metin = turkce_normalize((icerik or "") + " " + (html_govde or ""))

    if "DLK" in metin:
        return True, ""

    for firma in RAKIP_FIRMALAR:
        if firma in metin:
            return False, f"{firma} müşteri faturası"

    if "IHRACAT LIMAN" in metin and "OPERASYONEL HIZMET" in metin:
        return False, "İhracat Liman ve Operasyonel Hizmetler"
    if "DENIZ IHRACAT NAVLUNU" in metin or "DENIZ IHRACAT" in metin:
        return False, "Deniz ihracat navlunu"
    if "KONTEYNER VGM" in metin or "VGM HIZMET" in metin:
        return False, "Konteyner VGM Hizmeti"
    if re.search(r'(YUKLEYICI|GONDERICI|SHIPPER|CONSIGNOR)[:\s]*BRISA', metin):
        return False, "Yükleyici/Gönderici Brisa (ihracat faturası)"

    metin_email_haric = metin.replace("ITHALAT.BRISA@SUBASI.NET", "").replace("@SUBASI.NET", "")
    if "SOLMAZ GUMRUK" in metin_email_haric:
        return False, "Solmaz Gümrük Müşavirliği faturası"
    if "SUBASI GUMRUK" in metin_email_haric:
        return False, "Subaşı Gümrük Müşavirliği faturası"

    return True, ""


# ════════════════════════════════════════════════════════════
#  NUMARA ÇIKARMA (SADECE REGEX)
# ════════════════════════════════════════════════════════════

def numaralari_regex_ile_cek(metin):
    if not metin:
        return None

    mu = metin.upper()

    # Konteyner
    konteynerler = [k for k in set(re.findall(r'\b([A-Z]{4}[0-9]{7})\b', mu))
                    if not k.startswith('YLP')]

    # Beyanname
    tip_kodlari = ['IM', 'AN', 'EX', 'IH', 'TR', 'TI', 'AB', 'AT', 'EI']
    beyannameler = []
    for b in re.findall(r'\b(\d{5}[A-Z]{2}\d{8})\b', mu):
        if any(b[5:7] == t for t in tip_kodlari):
            beyannameler.append(b)
    for b in re.findall(r'\b(\d{2}/[A-Z]{2}\d{7,8})\b', mu):
        if any(f'/{t}' in b for t in tip_kodlari):
            beyannameler.append(b)
    for m in re.finditer(r'BEY\.?\s*NO[:\s]*(\d{6})\b', mu):
        beyannameler.append(m.group(1))
    beyannameler = list(set(beyannameler))

    # Konşimento
    konsimentolar = []
    kon_etiket = (
        r'(?:KON[Ss]IMENTO|B/?L|BILL OF LADING|BL NO'
        r'|AWB|HOUSE AWB|HAWB|MASTER AWB|MAWB)'
        r'[^A-Z0-9]{0,15}([A-Z0-9\-,\s]{6,80})'
    )
    for m in re.finditer(kon_etiket, mu):
        for kod in re.split(r'[,\s]+', m.group(1).strip()):
            kod = kod.strip()
            if len(kod) >= 6 and not kod.startswith(('YLP', 'TR1', 'TK')):
                konsimentolar.append(kod)

    konsimentolar.extend(re.findall(r'\b(\d{3,11}-\d{5,11})\b', mu))

    bilinen_prefix = ['MEDU', 'SPE', 'HLCU', 'MSK', 'ONE', 'CMA']
    for h in re.findall(r'\b([A-Z]{3,6}\d{6,10})\b', mu):
        if any(h.startswith(p) for p in bilinen_prefix):
            konsimentolar.append(h)

    for m in re.finditer(r'AWB\s*NO[:\s]*(\d{10})\b', mu):
        konsimentolar.append(m.group(1))
    for m in re.finditer(r'AWB[/\s]*H[II]ZMET[:\s]*(\d{10})\b', mu):
        konsimentolar.append(m.group(1))
    for m in re.finditer(r'AWB\s*NO[:\s]*(\d{3}-\d{8})\b', mu):
        konsimentolar.append(m.group(1))
    for m in re.finditer(r'(?:HOUSE\s*AWB|HAWB)[:\s]*([A-Z0-9\-,\s]{6,80})', mu):
        for kod in re.split(r'[,\s]+', m.group(1).strip()):
            if len(kod.strip()) >= 6:
                konsimentolar.append(kod.strip())

    # DHL tipi fatura — 3 varyasyon:
    # A) Boşluksuz birleşik tablo: "22/04/20269782128510TKU..." (tarih+AWB yapışık)
    for m in re.finditer(r'\d{2}/\d{2}/\d{4}(\d{10})(?=\D)', mu):
        konsimentolar.append(m.group(1))
    # B) Tab/boşluk ayrık tablo: "22/04/2026\t1291072031\tKSF"
    for m in re.finditer(r'\d{2}/\d{2}/\d{4}[\t ]+(\d{10})(?!\d)', mu):
        konsimentolar.append(m.group(1))
    # C) AWB/HIZMET başlığı hemen ardından: "AWB/HIZMET  9782123761  KSF"
    for m in re.finditer(r'AWB[/\s]*H[II]ZMET[\t ]+(\d{10})(?!\d)', mu):
        konsimentolar.append(m.group(1))

    konsimentolar = list(set(konsimentolar))

    sonuc = {
        "konsimento_list": konsimentolar,
        "konteyner_list":  konteynerler,
        "beyanname_list":  beyannameler,
    }
    return None if not any(sonuc.values()) else sonuc


# ════════════════════════════════════════════════════════════
#  EŞLEŞTİRME
# ════════════════════════════════════════════════════════════

def eslestir(numaralar, referans):
    f_kon = [_norm(v) for v in numaralar.get("konsimento_list", []) if v]
    f_knt = [_norm(v) for v in numaralar.get("konteyner_list",  []) if v]
    f_bey = [_norm(v) for v in numaralar.get("beyanname_list",  []) if v]

    eslesmeler = []
    gorulmus = set()

    for row in referans:
        dosya = row.get("dosya_no")
        kullanici = row.get("kullanici", "")
        if not dosya or dosya in gorulmus:
            continue

        for r_kon in row.get("konsimento_listesi", []):
            if r_kon and r_kon in f_kon:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konşimento No",
                                   "deger": r_kon, "kullanici": kullanici})
                gorulmus.add(dosya)
                break

        if dosya in gorulmus:
            continue

        for r_knt in row.get("konteyner_listesi", []):
            if r_knt and r_knt in f_knt:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konteyner No",
                                   "deger": r_knt, "kullanici": kullanici})
                gorulmus.add(dosya)
                break

        if dosya in gorulmus:
            continue

        for r_b in row.get("beyanname_listesi", []):
            if not r_b:
                continue
            r_son6 = ''.join(c for c in r_b if c.isdigit())[-6:]
            for f_b in f_bey:
                f_son6 = ''.join(c for c in f_b if c.isdigit())[-6:]
                if r_son6 and f_son6 and r_son6 == f_son6:
                    eslesmeler.append({"dosya_no": dosya, "kriter": "Beyanname No",
                                       "deger": f"{f_b} ≈ {r_b}", "kullanici": kullanici})
                    gorulmus.add(dosya)
                    break
            if dosya in gorulmus:
                break

    return eslesmeler if eslesmeler else None


# ════════════════════════════════════════════════════════════
#  ANA PIPELINE
# ════════════════════════════════════════════════════════════

def main():
    log.info("═══ Sistem başladı (IMAP + Service Account) v24 ═══")

    conn = db_init()
    referans = sheets_referans_veri()
    if not referans:
        log.warning("Referans verisi boş — işlem yapılamaz.")
        return

    mailler = gmail_okunmamis_mailler_imap()

    MAX_MAIL_PER_RUN = 50
    if len(mailler) > MAX_MAIL_PER_RUN:
        log.info(f"⚠️ {len(mailler)} mail bulundu, ilk {MAX_MAIL_PER_RUN} işlenecek")
        mailler = mailler[:MAX_MAIL_PER_RUN]

    eslesti = beklemeye = okunamadi = 0

    for mail_item in mailler:
        email_id = mail_item["id"]
        imap_num = mail_item["imap_num"]
        gonderen = mail_item["gonderen"]
        konu     = mail_item["konu"]
        tarih    = mail_item["tarih"]
        govde    = mail_item["html_body"]

        if db_islendi_mi(conn, email_id):
            log.info(f"Daha önce işlendi, atlanıyor: {email_id}")
            gmail_okundu_isaretle_imap(imap_num)
            continue

        log.info(f"Yeni mail işleniyor: {konu} | {gonderen}")

        url = linkten_url_bul(govde)
        if not url:
            log.info(f"Link bulunamadı: {konu}")
            gmail_okundu_isaretle_imap(imap_num)
            continue

        icerik = sayfayi_playwright_ile_oku(url)
        if not icerik:
            log.error(f"Sayfa okunamadı: {url[:80]}")
            sheets_faturaListesiYaz(gonderen, tarih, None, "❌ Okunamadı",
                                    fatura_url=url, email_id=email_id)
            gmail_okundu_isaretle_imap(imap_num)
            okunamadi += 1
            continue

        firma_adi = firma_adi_cek(icerik)

        bize_ait, sahip_olmama_sebebi = bize_ait_mi_kontrol(icerik, govde)
        if not bize_ait:
            log.info(f"Bize ait değil ({sahip_olmama_sebebi}): {konu}")
            sheets_faturaListesiYaz(
                firma_adi or gonderen, tarih, None,
                "🔴 Bize Ait Değil | " + sahip_olmama_sebebi,
                fatura_url=url, email_id=email_id
            )
            gmail_okundu_isaretle_imap(imap_num)
            continue

        numaralar = numaralari_regex_ile_cek(icerik)
        if not numaralar:
            log.info(f"Numara bulunamadı: {konu}")
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, None,
                                    "❌ Okunamadı", fatura_url=url, email_id=email_id)
            gmail_okundu_isaretle_imap(imap_num)
            okunamadi += 1
            continue

        log.info(
            f"Numaralar → Konşimento: {len(numaralar['konsimento_list'])} "
            f"| Konteyner: {len(numaralar['konteyner_list'])} "
            f"| Beyanname: {len(numaralar['beyanname_list'])}"
        )

        eslesmeler = eslestir(numaralar, referans)

        if eslesmeler:
            dosyalar_str = ", ".join(e["dosya_no"] for e in eslesmeler)
            kullanicilar_str = ", ".join(
                set(e.get("kullanici", "") for e in eslesmeler if e.get("kullanici"))
            )
            db_islendi_ekle(conn, email_id, dosyalar_str)
            sheets_faturaListesiYaz(
                firma_adi or gonderen, tarih, numaralar, "✅ Eşleşti",
                dosya_no=dosyalar_str, fatura_url=url,
                email_id=email_id, kullanici=kullanicilar_str
            )
            eslesti += 1
        else:
            db_bekleyen_ekle(conn, email_id, gonderen, konu, tarih, numaralar)
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar,
                                    "🟡 Bekliyor", fatura_url=url, email_id=email_id)
            beklemeye += 1

        gmail_okundu_isaretle_imap(imap_num)
        time.sleep(3)

    # ── Bekleyenleri yeniden dene — 3 Aşamalı + Çapraz Eleme ────
    servis = sheets_servis()
    bekleyenler = sheets_bekleyenleri_getir(servis)

    # YENİ v24: Ait değil numara havuzunu bir kere oku (tüm bekleyenler için kullan)
    ait_degil_havuzu = sheets_ait_degil_numaralari_getir(servis)

    yeniden_eslesti = 0
    bekleyen_elenen = 0

    for item in bekleyenler:
        item_gonderen   = item.get("gonderen", "")
        item_tarih      = item.get("tarih", "")
        item_email_id   = item.get("email_id", "")
        item_fatura_url = item.get("fatura_url", "")
        satir_no        = item["satir_no"]
        numaralar_item  = item["numaralar"]

        # AŞAMA 0 (YENİ v24): Çapraz eleme — numaralar ait değil havuzunda mı?
        capraz_gecti, capraz_sebep = capraz_eleme_kontrol(numaralar_item, ait_degil_havuzu)
        if not capraz_gecti:
            log.info(f"Bekleyen satır {satir_no} çapraz eleme ile elendi: {capraz_sebep}")
            sheets_bekleyeni_aitdegil_guncelle(servis, satir_no, capraz_sebep)
            bekleyen_elenen += 1
            continue

        # AŞAMA 1: Gönderen firma adından hızlı eleme
        gonderen_norm = turkce_normalize(item_gonderen)
        elenmis = False
        eleme_sebebi = ""
        for firma in RAKIP_FIRMALAR:
            if firma in gonderen_norm:
                elenmis = True
                eleme_sebebi = f"{firma} müşteri faturası"
                break

        if elenmis:
            log.info(f"Bekleyen satır {satir_no} elendi (firma adı): {eleme_sebebi}")
            sheets_bekleyeni_aitdegil_guncelle(servis, satir_no, eleme_sebebi)
            bekleyen_elenen += 1
            continue

        # AŞAMA 2: Fatura içeriği kontrolü (Playwright)
        if item_fatura_url:
            log.info(f"Bekleyen satır {satir_no}: İçerik kontrol ediliyor...")
            icerik = sayfayi_playwright_ile_oku(item_fatura_url)
            if icerik:
                bize_ait, sahip_olmama_sebebi = bize_ait_mi_kontrol(icerik, "")
                if not bize_ait:
                    log.info(f"Bekleyen satır {satir_no} elendi (içerik): {sahip_olmama_sebebi}")
                    sheets_bekleyeni_aitdegil_guncelle(servis, satir_no, sahip_olmama_sebebi)
                    bekleyen_elenen += 1
                    continue
            else:
                log.warning(f"Bekleyen satır {satir_no}: Sayfa okunamadı, eşleştirmeye devam")

        # AŞAMA 3: Eşleştirme
        log.info(
            f"Bekleyen satır {satir_no}: "
            f"Kon={numaralar_item.get('konsimento_list', [])} | "
            f"Knt={numaralar_item.get('konteyner_list', [])} | "
            f"Bey={numaralar_item.get('beyanname_list', [])}"
        )

        eslesmeler_b = eslestir(numaralar_item, referans)
        if not eslesmeler_b:
            log.info(f"Satır {satir_no}: Eşleşme bulunamadı")
            continue

        dosyalar_str_b = ", ".join(e["dosya_no"] for e in eslesmeler_b)
        ilk_e = eslesmeler_b[0]
        sheets_bekleyeni_guncelle(servis, satir_no,
                                   dosyalar_str_b, ilk_e["kriter"], ilk_e["deger"])

        if item_email_id:
            db_islendi_ekle(conn, item_email_id, dosyalar_str_b)

        yeniden_eslesti += 1

    log.info(
        f"═══ Bitti → Eşleşti: {eslesti} | "
        f"Beklemeye: {beklemeye} | "
        f"Okunamadı: {okunamadi} | "
        f"Bekleyenden eşleşti: {yeniden_eslesti} | "
        f"Bekleyenden elenen: {bekleyen_elenen} ═══"
    )
    conn.close()


if __name__ == "__main__":
    main()
