"""
Brisa Mail Otomasyon — v23 FINAL
Gmail IMAP (App Password) + Sheets Service Account = Sonsuz token
v22'nin tüm iyileştirmeleri korundu:
  - Türkçe karakter normalizasyonu
  - Rakip firma listesi (CABOT, MITSUI vb.)
  - Yükleyici/Gönderici Brisa kontrolü
  - Virgüllü konşimento parse
  - House AWB / HAWB desteği
  - _norm() ile eşleştirme normalizasyonu
  - Mükerrer kayıt kontrolü (Sheets)
  - page.inner_text("body") — düz metin (HTML değil)
"""

import base64
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

import requests
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
    r = conn.execute(
        "SELECT 1 FROM processed WHERE email_id=?", (email_id,)
    ).fetchone()
    if r:
        return True
    r = conn.execute(
        "SELECT 1 FROM pending WHERE email_id=?", (email_id,)
    ).fetchone()
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


def db_bekleyenleri_getir(conn):
    sinir = (datetime.utcnow() - timedelta(days=SORGU_PENCERESI_GUN)).isoformat()
    rows = conn.execute(
        "SELECT * FROM pending WHERE created_at >= ? ORDER BY created_at",
        (sinir,)
    ).fetchall()
    return [dict(zip([c[0] for c in conn.execute(
        "PRAGMA table_info(pending)"
    ).fetchall()], r)) for r in rows]


def db_bekleyeni_sil(conn, pending_id):
    conn.execute("DELETE FROM pending WHERE id=?", (pending_id,))
    conn.commit()


# ════════════════════════════════════════════════════════════
#  GMAIL IMAP (App Password — sonsuz token)
# ════════════════════════════════════════════════════════════

def gmail_imap_baglanti():
    """IMAP ile Gmail'e bağlan."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
        log.info(f"✅ IMAP bağlantısı başarılı: {GMAIL_EMAIL}")
        return mail
    except Exception as e:
        log.error(f"❌ IMAP bağlantı hatası: {e}")
        raise


def gmail_okunmamis_mailler_imap():
    """IMAP ile okunmamış Brisa maillerini çek."""
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

                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    # Subject decode
                    subject_header = msg["Subject"] or ""
                    decoded = decode_header(subject_header)
                    subject = ""
                    for content, encoding in decoded:
                        if isinstance(content, bytes):
                            subject += content.decode(encoding or "utf-8", errors="replace")
                        else:
                            subject += content

                    from_header = msg["From"] or ""
                    date_header = msg["Date"] or ""

                    # HTML gövde çek
                    html_body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/html":
                                payload = part.get_payload(decode=True)
                                if payload:
                                    charset = part.get_content_charset() or "utf-8"
                                    html_body = payload.decode(charset, errors="replace")
                                    break
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload:
                            charset = msg.get_content_charset() or "utf-8"
                            html_body = payload.decode(charset, errors="replace")

                    # Email ID — Message-ID header tercih edilir (stabil)
                    email_id = msg["Message-ID"] or num.decode()

                    detaylar.append({
                        "id": email_id,
                        "imap_num": num.decode(),
                        "gonderen": from_header,
                        "konu": subject,
                        "tarih": date_header,
                        "html_body": html_body,
                    })

                except Exception as e:
                    log.warning(f"Mail parse hatası (#{num}): {e}")
                    continue

        except Exception as e:
            log.error(f"Gönderen search hatası ({gonderen}): {e}")
            continue

    mail.close()
    mail.logout()

    log.info(f"📨 Toplam {len(detaylar)} okunmamış mail bulundu")
    return detaylar


def gmail_okundu_isaretle_imap(imap_num):
    """IMAP ile maili okundu işaretle."""
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
    """Sheets için Service Account credentials."""
    service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=SHEETS_SCOPES
    )
    return creds


def sheets_servis():
    return build("sheets", "v4", credentials=_sheets_creds())


def sheets_referans_veri():
    """
    📋 Referans sekmesini oku.
    Gerçek kolon sırası: A=Konşimento, B=Konteyner, C=Beyanname, D=Dosya No
    Birden fazla değer virgülle ayrılmış olabilir.
    """
    try:
        servis = sheets_servis()
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📋 Referans!A2:D"
        ).execute()

        rows = result.get("values", [])
        if not rows:
            log.warning("Referans sekmesi boş")
            return []

        kayitlar = []
        for row in rows:
            padded = row + [""] * (4 - len(row))
            dosya = str(padded[3]).strip() if padded[3] else None  # D=Dosya No
            if not dosya:
                continue

            # Virgülle ayrılmış birden fazla değer desteği
            konsimento_listesi = [
                _norm(k) for k in str(padded[0]).split(",") if k.strip()  # A=Konşimento
            ]
            konteyner_listesi = [
                _norm(k) for k in str(padded[1]).split(",") if k.strip()  # B=Konteyner
            ]
            beyanname_listesi = [
                _norm(k) for k in str(padded[2]).split(",") if k.strip()  # C=Beyanname
            ]

            kayitlar.append({
                "dosya_no":           dosya,
                "konsimento_listesi": konsimento_listesi,
                "konteyner_listesi":  konteyner_listesi,
                "beyanname_listesi":  beyanname_listesi,
            })

        log.info(f"📋 {len(kayitlar)} referans kaydı yüklendi")
        return kayitlar

    except HttpError as e:
        log.error(f"Sheets okuma hatası: {e}")
        return []


def sheets_fatura_islendi_mi(servis, email_id):
    """Fatura Listesi K kolonunda bu email_id var mı kontrol et (mükerrer önleme)."""
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!K:K"
        ).execute()
        rows = result.get("values", [])
        for row in rows:
            if row and str(row[0]).strip() == str(email_id).strip():
                return True
    except HttpError:
        pass
    return False


def sheets_faturaListesiYaz(gonderen, tarih, numaralar, durum,
                             dosya_no="", fatura_url="", email_id=""):
    """
    📑 Fatura Listesi sekmesine yaz.
    Kolonlar: A=Geliş Tarihi, B=Mail Tarihi, C=Gönderen, D=Konşimento,
              E=Konteyner, F=Beyanname, G=Durum, H=Dosya No,
              I=Fatura Linki, J=İşlemi Yapan, K=Email ID
    """
    try:
        servis = sheets_servis()

        # Mükerrer kontrol
        if email_id and sheets_fatura_islendi_mi(servis, email_id):
            log.info(f"Fatura listesinde zaten var, atlanıyor: {email_id}")
            return

        konsimento = ", ".join(numaralar.get("konsimento_list", [])) if numaralar else ""
        konteyner  = ", ".join(numaralar.get("konteyner_list",  [])) if numaralar else ""
        beyanname  = ", ".join(numaralar.get("beyanname_list",  [])) if numaralar else ""

        simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A:K",
            valueInputOption="RAW",
            body={"values": [[
                simdi,       # A: Geliş Tarihi
                tarih,       # B: Mail Tarihi
                gonderen,    # C: Gönderen Firma
                konsimento,  # D: Konşimento
                konteyner,   # E: Konteyner
                beyanname,   # F: Beyanname
                durum,       # G: Durum
                dosya_no,    # H: Dosya No
                fatura_url,  # I: Fatura Linki
                "",          # J: İşlemi Yapan (manuel)
                email_id,    # K: Email ID (mükerrer kontrol)
            ]]}
        ).execute()

        log.info(f"✅ Fatura Listesi güncellendi: {durum}")

    except HttpError as e:
        log.error(f"Sheets yazma hatası: {e}")


def sheets_eslesmeyiKaydet(gonderen, konu, tarih, dosya_no, kriter, deger):
    """✅ Eşleşenler sekmesine kaydet."""
    try:
        servis = sheets_servis()
        simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="✅ Eşleşenler!A:G",
            valueInputOption="RAW",
            body={"values": [[
                simdi, tarih, gonderen, konu, dosya_no, kriter, deger
            ]]}
        ).execute()

        log.info(f"✅ Eşleşme kaydedildi: {dosya_no}")

    except HttpError as e:
        log.error(f"Sheets eşleşme kayıt hatası: {e}")


def sheets_okunamayanEkle(gonderen, konu, tarih, sebep, fatura_url=""):
    """❌ Okunamayanlar sekmesine ekle."""
    try:
        servis = sheets_servis()
        simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="❌ Okunamayanlar!A:F",
            valueInputOption="RAW",
            body={"values": [[
                simdi, tarih, gonderen, konu, sebep, fatura_url
            ]]}
        ).execute()

        log.info(f"❌ Okunamayan kaydedildi: {sebep}")

    except HttpError as e:
        log.error(f"Sheets okunamayan kayıt hatası: {e}")


def sheets_bekleyenleri_getir(servis):
    """
    📑 Fatura Listesi'nden 🟡 Bekliyor durumundaki satırları oku.
    SQLite'a bağımlılık yok — Sheets kalıcı kaynak.
    """
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A2:K"
        ).execute()

        rows = result.get("values", [])
        if len(rows) < 1:
            return []

        bekleyenler = []
        for i, row in enumerate(rows, start=2):  # header=1, data start=2
            padded = row + [""] * (11 - len(row))
            durum_hucre = padded[6].strip()  # G: Durum

            if "Bekliyor" not in durum_hucre and "🟡" not in durum_hucre:
                continue

            numaralar = {
                "konsimento_list": [v.strip() for v in padded[3].split(",") if v.strip()],
                "konteyner_list":  [v.strip() for v in padded[4].split(",") if v.strip()],
                "beyanname_list":  [v.strip() for v in padded[5].split(",") if v.strip()],
            }

            # Hiç numara yoksa atla (okunamayan satır)
            if not any(numaralar.values()):
                continue

            bekleyenler.append({
                "satir_no":  i,
                "gonderen":  padded[2].strip(),   # C
                "tarih":     padded[1].strip(),   # B
                "email_id":  padded[10].strip(),  # K
                "numaralar": numaralar,
            })

        log.info(f"🟡 {len(bekleyenler)} bekleyen fatura bulundu")
        return bekleyenler

    except HttpError as e:
        log.error(f"Bekleyenler okuma hatası: {e}")
        return []


def sheets_bekleyeni_guncelle(servis, satir_no, dosya_no, kriter, deger):
    """Bekleyen satırın Durum (G) ve Dosya No (H) kolonlarını güncelle."""
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


# ════════════════════════════════════════════════════════════
#  PLAYWRIGHT — JS RENDER
# ════════════════════════════════════════════════════════════

def sayfayi_playwright_ile_oku(url):
    """
    Playwright ile sayfayı tam render et.
    inner_text("body") ile düz metin döndür — regex HTML'de çalışmaz!
    """
    log.info(f"Playwright ile sayfa açılıyor: {url[:80]}...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"]
            )
            page = browser.new_page()
            page.goto(url, wait_until="networkidle", timeout=30000)
            page.wait_for_timeout(3000)  # edoksis için ek bekleme

            # ÖNEMLİ: inner_text() — düz metin, HTML değil
            # page.content() HTML döndürür, regex pattern'ları bozulur
            icerik = page.inner_text("body")
            browser.close()

            log.info(f"Sayfa okundu: {len(icerik)} karakter")
            return icerik[:15000]  # Çok büyük sayfalarda kırp

    except Exception as e:
        log.error(f"Playwright hatası: {e}")
        return None


# ════════════════════════════════════════════════════════════
#  YARDIMCI FONKSİYONLAR
# ════════════════════════════════════════════════════════════

def turkce_normalize(metin):
    """
    Türkçe ve İngilizce karakterleri normalize eder.
    Büyük/küçük harf + Türkçe karakter farklarını ortadan kaldırır.
    Örnek: "Subaşı Gümrük" → "SUBASI GUMRUK"
    """
    if not metin:
        return ""
    metin = metin.upper()
    tr_map = {
        'Ç': 'C', 'Ğ': 'G', 'İ': 'I', 'Ö': 'O', 'Ş': 'S', 'Ü': 'U',
        'ç': 'C', 'ğ': 'G', 'ı': 'I', 'i': 'I', 'ö': 'O', 'ş': 'S', 'ü': 'U'
    }
    for tr_char, eng_char in tr_map.items():
        metin = metin.replace(tr_char, eng_char)
    return metin


def _norm(v):
    """Eşleştirme normalizasyonu: büyük harf + sadece alfanümerik."""
    if not v:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(v).upper())


def linkten_url_bul(html_govde):
    """Mail gövdesinden edoksis fatura linkini çıkar."""
    pattern = r'https?://[^\s"\'<>]*edoksis[^\s"\'<>]*'
    match = re.search(pattern, html_govde, re.IGNORECASE)
    if match:
        url = match.group(0).replace("&amp;", "&")
        log.info(f"🔗 Link bulundu: {url[:80]}...")
        return url
    return None


def firma_adi_cek(icerik):
    """
    Fatura sayfasının düz metninden gönderici firma adını çeker.
    'Gönderen' kelimesinin hemen ardından gelen büyük harfli şirket adı.
    """
    if not icerik:
        return ""
    try:
        pattern = r'G[\xf6o]nderen[^\n]{0,5}([A-Z][A-Z\s\.&,]{5,79})'
        match = re.search(pattern, icerik, re.IGNORECASE)
        if match:
            firma = match.group(1).strip()
            if re.match(r'^[\d\-]+$', firma):
                return ""
            return firma[:80].strip()
    except Exception:
        pass
    return ""


def bize_ait_mi_kontrol(icerik, html_govde):
    """
    Faturanın bize ait olup olmadığını kontrol eder.

    Mantık:
    - DLK geçiyorsa → kesinlikle bizim
    - Rakip müşteri firmaları → bize ait değil
    - İhracat hizmetleri → bize ait değil
    - Brisa gönderici/yükleyici → bize ait değil (ihracat)
    - Rakip gümrükçüler → bize ait değil

    Döndürür: (bize_ait: bool, sebep: str)
    """
    metin_ham = (icerik or "") + " " + (html_govde or "")
    metin = turkce_normalize(metin_ham)

    # DLK geçiyorsa kesinlikle bizim — diğer kontrollere gerek yok
    if "DLK" in metin:
        return True, ""

    # ── RAKİP/MÜŞTERİ FİRMALARI ───────────────────────────────
    rakip_firmalar = [
        "CABOT",
        "MITSUI",
        "TANAKA",
        "ZEON CHEMICAL",
        "KOLON",
        "THAI TOKAI",
        "HS HYOSUNG",
        "PYRAMID"
    ]
    for firma in rakip_firmalar:
        if firma in metin:
            return False, f"{firma} müşteri faturası"

    # ── İHRACAT HİZMETLERİ ─────────────────────────────────────
    if "IHRACAT LIMAN" in metin and "OPERASYONEL HIZMET" in metin:
        return False, "İhracat Liman ve Operasyonel Hizmetler"

    if "DENIZ IHRACAT NAVLUNU" in metin or "DENIZ IHRACAT" in metin:
        return False, "Deniz ihracat navlunu"

    if "KONTEYNER VGM" in metin or "VGM HIZMET" in metin:
        return False, "Konteyner VGM Hizmeti"

    # ── YÜKLEYİCİ/GÖNDERİCİ BRISA ──────────────────────────────
    yukleyici_pattern = r'(YUKLEYICI|GONDERICI|SHIPPER|CONSIGNOR)[:\s]*BRISA'
    if re.search(yukleyici_pattern, metin):
        return False, "Yükleyici/Gönderici Brisa (ihracat faturası)"

    # ── RAKİP GÜMRÜKÇÜLER ─────────────────────────────────────
    # ithalat.brisa@subasi.net adresini hariç tut (Brisa'nın kendi adresi)
    metin_email_haric = metin.replace("ITHALAT.BRISA@SUBASI.NET", "").replace("@SUBASI.NET", "")

    if "SOLMAZ GUMRUK" in metin_email_haric:
        return False, "Solmaz Gümrük Müşavirliği faturası"
    if "SUBASI GUMRUK" in metin_email_haric:
        return False, "Subaşı Gümrük Müşavirliği faturası"

    return True, ""


# ════════════════════════════════════════════════════════════
#  NUMARA ÇIKARMA (SADECE REGEX — Gemini kapalı)
# ════════════════════════════════════════════════════════════

def numaralari_regex_ile_cek(metin):
    """
    Sadece regex ile numara çıkarımı.
    Düz metin (inner_text) üzerinde çalışır.

    Konteyner : ISO 6346 — 4 büyük harf + 7 rakam
    Beyanname : Türk Gümrük — standart / slash / kısa format
    Konşimento: Etiket yanında / tire / harfli prefix + AWB/HAWB
    """
    if not metin:
        return None

    metin_upper = metin.upper()

    # ── Konteyner (ISO 6346) ──────────────────────────────────
    konteyner_pattern = r'\b([A-Z]{4}[0-9]{7})\b'
    konteynerler = list(set(re.findall(konteyner_pattern, metin_upper)))
    # Fatura no gibi YLP... ile başlayanları ele
    konteynerler = [k for k in konteynerler if not k.startswith('YLP')]

    # ── Beyanname (Türk Gümrük) ───────────────────────────────
    beyannameler = []
    tip_kodlari = ['IM', 'AN', 'EX', 'IH', 'TR', 'TI', 'AB', 'AT', 'EI']

    # 1. Standart: 26410500IM00045784
    beyanname_pattern = r'\b(\d{5}[A-Z]{2}\d{8})\b'
    beyler = re.findall(beyanname_pattern, metin_upper)
    beyannameler.extend([b for b in beyler if any(b[5:7] == tip for tip in tip_kodlari)])

    # 2. Slash: 26/IM0226949
    slash_pattern = r'\b(\d{2}/[A-Z]{2}\d{7,8})\b'
    slash_beyler = re.findall(slash_pattern, metin_upper)
    beyannameler.extend([b for b in slash_beyler if any(f'/{tip}' in b for tip in tip_kodlari)])

    # 3. Kısa (Bey.No yanında): 6 rakam
    bey_kisa = r'BEY\.?\s*NO[:\s]*(\d{6})\b'
    for m in re.finditer(bey_kisa, metin_upper):
        beyannameler.append(m.group(1))

    beyannameler = list(set(beyannameler))

    # ── Konşimento ────────────────────────────────────────────
    konsimentolar = []

    # 1. Etiket yanındaki alfanümerik kod (AWB, HAWB, HOUSE AWB dahil)
    kon_etiket = (
        r'(?:KON[Ss]IMENTO|B/?L|BILL OF LADING|BL NO'
        r'|AWB|HOUSE AWB|HAWB|MASTER AWB|MAWB)'
        r'[^A-Z0-9]{0,15}([A-Z0-9\-,\s]{6,80})'
    )
    for m in re.finditer(kon_etiket, metin_upper):
        kod_ham = m.group(1).strip()
        # Virgülle ayrılmış olabilir: "SPE041901781, SPE041901782"
        parcalar = [k.strip() for k in re.split(r'[,\s]+', kod_ham) if k.strip()]
        for kod in parcalar:
            if (len(kod) >= 6
                    and not kod.startswith('YLP')
                    and not kod.startswith('TR1')
                    and not kod.startswith('TK')):
                konsimentolar.append(kod)

    # 2. Tire ile: 205-75999114, 61598712294-11
    kon_tire = r'\b(\d{3,11}-\d{5,11})\b'
    konsimentolar.extend(re.findall(kon_tire, metin_upper))

    # 3. Harfli prefix: MEDUKC378982, HLCUIST2501XXXXX
    bilinen_prefix = ['MEDU', 'SPE', 'HLCU', 'MSK', 'ONE', 'CMA']
    kon_harfli = r'\b([A-Z]{3,6}\d{6,10})\b'
    for h in re.findall(kon_harfli, metin_upper):
        if any(h.startswith(p) for p in bilinen_prefix):
            konsimentolar.append(h)

    # ── AWB (Hava Konşimentosu) ───────────────────────────────
    # AWB NO: 10 rakam
    awb_10 = r'AWB\s*NO[:\s]*(\d{10})\b'
    for m in re.finditer(awb_10, metin_upper):
        konsimentolar.append(m.group(1))

    # AWB/HİZMET: 10 rakam (Türkçe I normalize)
    awb_hizmet = r'AWB[/\s]*H[II]ZMET[:\s]*(\d{10})\b'
    for m in re.finditer(awb_hizmet, metin_upper):
        konsimentolar.append(m.group(1))

    # Airline AWB: 235-87235831
    awb_tire = r'AWB\s*NO[:\s]*(\d{3}-\d{8})\b'
    for m in re.finditer(awb_tire, metin_upper):
        konsimentolar.append(m.group(1))

    # House AWB (virgüllü)
    house_awb = r'(?:HOUSE\s*AWB|HAWB)[:\s]*([A-Z0-9\-,\s]{6,80})'
    for m in re.finditer(house_awb, metin_upper):
        kod_ham = m.group(1).strip()
        parcalar = [k.strip() for k in re.split(r'[,\s]+', kod_ham) if k.strip()]
        for kod in parcalar:
            if len(kod) >= 6:
                konsimentolar.append(kod)

    konsimentolar = list(set(konsimentolar))

    sonuc = {
        "konsimento_list": konsimentolar,
        "konteyner_list":  konteynerler,
        "beyanname_list":  beyannameler,
    }

    hic_yok = not any([konsimentolar, konteynerler, beyannameler])
    return None if hic_yok else sonuc


# ════════════════════════════════════════════════════════════
#  EŞLEŞTİRME
# ════════════════════════════════════════════════════════════

def eslestir(numaralar, referans):
    """
    Faturadan çıkarılan numaraları Sheets referans listesiyle karşılaştırır.
    TÜM eşleşmeleri döndürür — aynı konşimento için birden fazla dosya olabilir.

    Normalizasyon: _norm() ile harf/rakam dışı karakterler temizlenir.
    Beyanname: sadece rakamların son 6'sı karşılaştırılır (esnek eşleştirme).
    """
    f_konsimentolar = [_norm(v) for v in numaralar.get("konsimento_list", []) if v]
    f_konteynerlar  = [_norm(v) for v in numaralar.get("konteyner_list",  []) if v]
    f_beyannameler  = [_norm(v) for v in numaralar.get("beyanname_list",  []) if v]

    eslesmeler = []
    gorulmus_dosyalar = set()

    for row in referans:
        dosya = row.get("dosya_no")
        if not dosya or dosya in gorulmus_dosyalar:
            continue

        # Konşimento eşleşmesi (tam)
        for r_kon in row.get("konsimento_listesi", []):
            if r_kon and r_kon in f_konsimentolar:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konşimento No", "deger": r_kon})
                gorulmus_dosyalar.add(dosya)
                break

        if dosya in gorulmus_dosyalar:
            continue

        # Konteyner eşleşmesi (tam)
        for r_knt in row.get("konteyner_listesi", []):
            if r_knt and r_knt in f_konteynerlar:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konteyner No", "deger": r_knt})
                gorulmus_dosyalar.add(dosya)
                break

        if dosya in gorulmus_dosyalar:
            continue

        # Beyanname eşleşmesi (son 6 rakam — esnek)
        for r_bey in row.get("beyanname_listesi", []):
            if not r_bey:
                continue
            r_bey_rakamlar = ''.join(c for c in r_bey if c.isdigit())
            r_son6 = r_bey_rakamlar[-6:] if len(r_bey_rakamlar) >= 6 else r_bey_rakamlar

            for f_bey in f_beyannameler:
                f_bey_rakamlar = ''.join(c for c in f_bey if c.isdigit())
                f_son6 = f_bey_rakamlar[-6:] if len(f_bey_rakamlar) >= 6 else f_bey_rakamlar

                if (r_son6 and f_son6 and r_son6 == f_son6) or r_bey == f_bey:
                    eslesmeler.append({
                        "dosya_no": dosya,
                        "kriter":   "Beyanname No",
                        "deger":    f"{f_bey} ≈ {r_bey}"
                    })
                    gorulmus_dosyalar.add(dosya)
                    break

            if dosya in gorulmus_dosyalar:
                break

    return eslesmeler if eslesmeler else None


# ════════════════════════════════════════════════════════════
#  ANA PIPELINE
# ════════════════════════════════════════════════════════════

def main():
    log.info("═══ Sistem başladı (IMAP + Service Account) v23 ═══")

    conn = db_init()
    referans = sheets_referans_veri()

    if not referans:
        log.warning("Referans verisi boş — işlem yapılamaz.")
        return

    mailler = gmail_okunmamis_mailler_imap()

    # Her çalışmada max 50 mail (GitHub Actions timeout önleme)
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

        # Mail gövdesinden edoksis linkini çek
        url = linkten_url_bul(govde)

        if not url:
            log.info(f"Link bulunamadı: {konu}")
            gmail_okundu_isaretle_imap(imap_num)
            continue

        # Playwright ile sayfayı oku (düz metin)
        icerik = sayfayi_playwright_ile_oku(url)

        if not icerik:
            log.error(f"Sayfa okunamadı: {url[:80]}")
            sheets_okunamayanEkle(gonderen, konu, tarih, "Sayfa açılamadı", fatura_url=url)
            sheets_faturaListesiYaz(gonderen, tarih, None, "❌ Okunamadı",
                                    fatura_url=url, email_id=email_id)
            gmail_okundu_isaretle_imap(imap_num)
            okunamadi += 1
            continue

        # Firma adını fatura sayfasından çek
        firma_adi = firma_adi_cek(icerik)

        # Bize ait mi kontrol et
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

        # Sadece Regex ile numaraları çıkar
        numaralar = numaralari_regex_ile_cek(icerik)

        if not numaralar:
            log.info(f"Numara bulunamadı: {konu}")
            sheets_okunamayanEkle(
                gonderen, konu, tarih,
                "Konşimento/Konteyner/Beyanname bulunamadı",
                fatura_url=url
            )
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

        # Eşleştir
        eslesmeler = eslestir(numaralar, referans)

        if eslesmeler:
            for e in eslesmeler:
                sheets_eslesmeyiKaydet(
                    gonderen, konu, tarih,
                    e["dosya_no"], e["kriter"], e["deger"]
                )
            dosyalar_str = ", ".join(e["dosya_no"] for e in eslesmeler)
            db_islendi_ekle(conn, email_id, dosyalar_str)
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "✅ Eşleşti",
                                    dosya_no=dosyalar_str, fatura_url=url, email_id=email_id)
            eslesti += 1
        else:
            db_bekleyen_ekle(conn, email_id, gonderen, konu, tarih, numaralar)
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "🟡 Bekliyor",
                                    fatura_url=url, email_id=email_id)
            beklemeye += 1

        gmail_okundu_isaretle_imap(imap_num)
        time.sleep(3)  # edoksis rate limit önleme

    # ── Bekleyenleri Sheets'ten yeniden dene ─────────────────
    servis = sheets_servis()
    bekleyenler = sheets_bekleyenleri_getir(servis)
    yeniden_eslesti = 0

    for item in bekleyenler:
        numaralar_item = item["numaralar"]
        eslesmeler_b   = eslestir(numaralar_item, referans)
        if not eslesmeler_b:
            continue

        item_gonderen = item.get("gonderen", "")
        item_tarih    = item.get("tarih", "")
        item_email_id = item.get("email_id", "")
        satir_no      = item["satir_no"]

        dosyalar_str_b = ", ".join(e["dosya_no"] for e in eslesmeler_b)
        ilk_e = eslesmeler_b[0]

        sheets_bekleyeni_guncelle(
            servis, satir_no,
            dosyalar_str_b, ilk_e["kriter"], ilk_e["deger"]
        )

        for e in eslesmeler_b:
            sheets_eslesmeyiKaydet(
                item_gonderen, "", item_tarih,
                e["dosya_no"], e["kriter"], e["deger"]
            )

        if item_email_id:
            db_islendi_ekle(conn, item_email_id, dosyalar_str_b)

        yeniden_eslesti += 1

    log.info(
        f"═══ Bitti → Eşleşti: {eslesti} | "
        f"Beklemeye: {beklemeye} | "
        f"Okunamadı: {okunamadi} | "
        f"Bekleyenden eşleşti: {yeniden_eslesti} ═══"
    )
    conn.close()


if __name__ == "__main__":
    main()
