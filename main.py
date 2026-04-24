"""
Brisa Mail Otomasyon — IMAP Versiyonu (Token Sorunu Yok)
Gmail IMAP ile mail okuma (OAuth token yenileme derdi olmadan)
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
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
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
GOOGLE_API_KEY       = os.environ["GOOGLE_API_KEY"]
GMAIL_APP_PASSWORD   = os.environ["GMAIL_APP_PASSWORD"]       # IMAP şifresi (yeni)
GMAIL_EMAIL          = os.environ.get("GMAIL_EMAIL", "murat.dagli@dlklogistics.com")
SHEETS_CREDENTIALS   = os.environ.get("SHEETS_CREDENTIALS", "")  # Sheets için OAuth (ayrı)
SHEETS_ID            = os.environ["SHEETS_ID"]
BILDIRIM_ALICISI     = os.environ["BILDIRIM_ALICISI"]
GONDEREN_LISTESI     = os.environ.get("GONDEREN_LISTESI", "m.hizmet@brisa.com.tr").split(",")
SORGU_PENCERESI_GUN  = int(os.environ.get("SORGU_PENCERESI_GUN", "60"))

# ── Sabitler ─────────────────────────────────────────────────
DB_PATH = Path("brisa_mail.db")
GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.0-flash:generateContent?key=" + GOOGLE_API_KEY
)

SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# ════════════════════════════════════════════════════════════
#  VERİTABANI (DEĞİŞMEDİ)
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
#  GMAIL IMAP (YENİ - OAuth yerine App Password)
# ════════════════════════════════════════════════════════════

def gmail_imap_baglanti():
    """IMAP ile Gmail'e bağlan"""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
        log.info(f"✅ IMAP bağlantısı başarılı: {GMAIL_EMAIL}")
        return mail
    except Exception as e:
        log.error(f"❌ IMAP bağlantı hatası: {e}")
        raise


def gmail_okunmamis_mailler_imap():
    """IMAP ile okunmamış Brisa maillerini çek"""
    mail = gmail_imap_baglanti()
    mail.select("inbox")
    
    # Gönderenler için IMAP search query oluştur
    search_parts = []
    for gonderen in GONDEREN_LISTESI:
        search_parts.append(f'FROM "{gonderen.strip()}"')
    
    # IMAP UNSEEN (okunmamış) + FROM filtresi
    # Not: IMAP OR operatörü karmaşık, basit tutuyoruz
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
                    subject_header = msg["Subject"]
                    if subject_header:
                        decoded = decode_header(subject_header)
                        subject = ""
                        for content, encoding in decoded:
                            if isinstance(content, bytes):
                                subject += content.decode(encoding or "utf-8", errors="replace")
                            else:
                                subject += content
                    else:
                        subject = ""
                    
                    # From
                    from_header = msg["From"] or ""
                    
                    # Date
                    date_header = msg["Date"] or ""
                    
                    # Body (HTML)
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
                    
                    # Email ID (Message-ID header kullan)
                    email_id = msg["Message-ID"] or num.decode()
                    
                    detaylar.append({
                        "id": email_id,
                        "imap_num": num.decode(),  # IMAP işaretlemek için
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
    """IMAP ile maili okundu işaretle"""
    try:
        mail = gmail_imap_baglanti()
        mail.select("inbox")
        mail.store(imap_num.encode(), '+FLAGS', '\\Seen')
        mail.close()
        mail.logout()
    except Exception as e:
        log.warning(f"IMAP okundu işareti hatası: {e}")


# ════════════════════════════════════════════════════════════
#  SHEETS (OAuth ile - değişmedi ama ayrı credentials)
# ════════════════════════════════════════════════════════════

def _sheets_creds():
    """Sheets için Service Account credentials (TOKEN YOK!)"""
    service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=SHEETS_SCOPES
    )
    log.info("✅ Service Account ile Sheets erişimi sağlandı")
    return creds


def sheets_servis():
    return build("sheets", "v4", credentials=_sheets_creds())


def sheets_referans_veri():
    """📋 Referans sekmesini oku"""
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
        
        referans = []
        for row in rows:
            if len(row) < 4:
                continue
            referans.append({
                "dosya_no": row[0].strip(),
                "konsimento": row[1].strip() if len(row) > 1 else "",
                "konteyner": row[2].strip() if len(row) > 2 else "",
                "beyanname": row[3].strip() if len(row) > 3 else "",
            })
        
        log.info(f"📋 {len(referans)} referans kaydı yüklendi")
        return referans
    
    except HttpError as e:
        log.error(f"Sheets okuma hatası: {e}")
        return []


def sheets_faturaListesiYaz(gonderen, tarih, numaralar, durum, dosya_no="", fatura_url="", email_id=""):
    """📑 Fatura Listesi sekmesine yaz"""
    try:
        servis = sheets_servis()
        
        # Numaraları formatla
        konsimento = ", ".join(numaralar.get("konsimento_list", [])) if numaralar else ""
        konteyner = ", ".join(numaralar.get("konteyner_list", [])) if numaralar else ""
        beyanname = ", ".join(numaralar.get("beyanname_list", [])) if numaralar else ""
        
        yeni_satir = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Geliş Tarihi
            tarih,           # Mail Tarihi
            gonderen,        # Gönderen Firma
            konsimento,      # Konşimento
            konteyner,       # Konteyner
            beyanname,       # Beyanname
            durum,           # Durum
            dosya_no,        # Dosya No
            fatura_url,      # Fatura Linki
            "",              # İşlemi Yapan (manuel)
            email_id         # Email ID (gizli)
        ]
        
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A:K",
            valueInputOption="USER_ENTERED",
            body={"values": [yeni_satir]}
        ).execute()
        
        log.info(f"✅ Fatura Listesi güncellendi: {durum}")
    
    except HttpError as e:
        log.error(f"Sheets yazma hatası: {e}")


def sheets_eslesmeyiKaydet(gonderen, konu, tarih, dosya_no, kriter, deger):
    """✅ Eşleşenler sekmesine kaydet"""
    try:
        servis = sheets_servis()
        
        yeni_satir = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            gonderen,
            konu,
            tarih,
            dosya_no,
            kriter,
            deger
        ]
        
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="✅ Eşleşenler!A:G",
            valueInputOption="USER_ENTERED",
            body={"values": [yeni_satir]}
        ).execute()
        
        log.info(f"✅ Eşleşme kaydedildi: {dosya_no}")
    
    except HttpError as e:
        log.error(f"Sheets eşleşme kayıt hatası: {e}")


def sheets_okunamayanEkle(gonderen, konu, tarih, sebep, fatura_url=""):
    """❌ Okunamayanlar sekmesine ekle"""
    try:
        servis = sheets_servis()
        
        yeni_satir = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            gonderen,
            konu,
            tarih,
            sebep,
            fatura_url
        ]
        
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="❌ Okunamayanlar!A:F",
            valueInputOption="USER_ENTERED",
            body={"values": [yeni_satir]}
        ).execute()
        
        log.info(f"❌ Okunamayan kaydedildi: {sebep}")
    
    except HttpError as e:
        log.error(f"Sheets okunamayan kayıt hatası: {e}")


def sheets_bekleyenleri_getir(servis):
    """🟡 Bekleyen faturalarını Sheets'ten çek"""
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A2:K"
        ).execute()
        
        rows = result.get("values", [])
        bekleyenler = []
        
        for i, row in enumerate(rows, start=2):
            if len(row) < 7:
                continue
            
            durum = row[6] if len(row) > 6 else ""
            if "🟡" not in durum and "Bekli" not in durum:
                continue
            
            # Numaraları parse et
            numaralar = {
                "konsimento_list": [x.strip() for x in row[3].split(",") if x.strip()] if len(row) > 3 else [],
                "konteyner_list": [x.strip() for x in row[4].split(",") if x.strip()] if len(row) > 4 else [],
                "beyanname_list": [x.strip() for x in row[5].split(",") if x.strip()] if len(row) > 5 else [],
            }
            
            bekleyenler.append({
                "satir_no": i,
                "gonderen": row[2] if len(row) > 2 else "",
                "tarih": row[1] if len(row) > 1 else "",
                "numaralar": numaralar,
                "email_id": row[10] if len(row) > 10 else "",
            })
        
        log.info(f"🟡 {len(bekleyenler)} bekleyen fatura bulundu")
        return bekleyenler
    
    except HttpError as e:
        log.error(f"Bekleyenler okuma hatası: {e}")
        return []


def sheets_bekleyeni_guncelle(servis, satir_no, dosya_no, kriter, deger):
    """Bekleyen satırını güncelle (✅ Eşleşti)"""
    try:
        durum_yeni = f"✅ Eşleşti ({kriter}: {deger})"
        
        servis.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"📑 Fatura Listesi!G{satir_no}:H{satir_no}",
            valueInputOption="USER_ENTERED",
            body={"values": [[durum_yeni, dosya_no]]}
        ).execute()
        
        log.info(f"✅ Bekleyen güncellendi (satır {satir_no}): {dosya_no}")
    
    except HttpError as e:
        log.error(f"Bekleyen güncelleme hatası: {e}")


# ════════════════════════════════════════════════════════════
#  PLAYWRIGHT — JS RENDER (DEĞİŞMEDİ)
# ════════════════════════════════════════════════════════════

def sayfayi_playwright_ile_oku(url):
    """Playwright ile JS render edilen sayfayı oku"""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            
            # edoksis rate limit önleme
            page.goto(url, wait_until="networkidle", timeout=30000)
            time.sleep(3)
            
            icerik = page.content()
            browser.close()
            
            return icerik
    
    except Exception as e:
        log.error(f"Playwright hatası: {e}")
        return None


# ════════════════════════════════════════════════════════════
#  NUMARA ÇIKARMA (SADECE REGEX - Gemini Kapalı)
# ════════════════════════════════════════════════════════════

def linkten_url_bul(html_govde):
    """Mail içinden edoksis linkini çıkar"""
    match = re.search(r'https://edoksis\.com/[^\s"<>]+', html_govde, re.IGNORECASE)
    return match.group(0) if match else None


def firma_adi_cek(icerik):
    """Fatura sayfasından firma adını regex ile çek"""
    patterns = [
        r'<div[^>]*class="[^"]*company-name[^"]*"[^>]*>(.*?)</div>',
        r'<span[^>]*class="[^"]*firm[^"]*"[^>]*>(.*?)</span>',
        r'<h[1-3][^>]*>(.*?GÜMRÜK.*?)</h[1-3]>',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, icerik, re.IGNORECASE | re.DOTALL)
        if match:
            firma = re.sub(r'<[^>]+>', '', match.group(1)).strip()
            return firma[:100] if firma else None
    
    return None


def bize_ait_mi_kontrol(icerik, govde):
    """Eleme kriterleri: Solmaz/Subaşı/VGM/Deniz İhracat kontrolü"""
    
    # Email temizleme (ithalat.brisa@subasi.net geçer)
    if "ithalat.brisa@subasi.net" in govde.lower():
        return True, ""
    
    metin = (icerik + " " + govde).lower()
    
    # DLK varsa kesinlikle bizim
    if "dlk" in metin:
        return True, ""
    
    # Eleme kriterleri
    if "solmaz" in metin and ("gumruk" in metin or "gümrük" in metin):
        return False, "Solmaz Gümrük"
    
    if "subasi" in metin or "subaşı" in metin:
        if "gumruk" in metin or "gümrük" in metin:
            return False, "Subaşı Gümrük"
    
    if "deniz ihracat navlunu" in metin:
        return False, "Deniz İhracat Navlunu"
    
    if "konteyner vgm hizmet" in metin or "vgm hizmet" in metin:
        return False, "Konteyner VGM Hizmeti"
    
    return True, ""


def numaralari_regex_ile_cek(icerik):
    """Sadece regex ile numaraları çıkar (Gemini kapalı)"""
    
    # Konteyner (ISO 6346)
    konteyner_pattern = r'\b[A-Z]{4}[0-9]{7}\b'
    konteyner_list = list(set(re.findall(konteyner_pattern, icerik)))
    
    # Beyanname (Türk Gümrük)
    beyanname_patterns = [
        r'\b\d{5}(IM|AN|EX|IH|TR|TI|AB|AT|EI)\d{8}\b',  # 26410500IM00045784
        r'\b\d{2}/(IM|AN|EX|IH|TR|TI|AB|AT|EI)\d{7,8}\b',  # 26/IM0226949
        r'BEY\.?\s*NO[:\s]*(\d{6})\b',  # Bey.No 228183
    ]
    beyanname_list = []
    for pattern in beyanname_patterns:
        matches = re.findall(pattern, icerik, re.IGNORECASE)
        for m in matches:
            if isinstance(m, tuple):
                beyanname_list.append(m[0] if len(m) == 1 else "".join(m))
            else:
                beyanname_list.append(m)
    beyanname_list = list(set(beyanname_list))
    
    # Konşimento
    konsimento_patterns = [
        r'\b(MEDU|SPE|HLCU|MSK|ONE|CMA)[A-Z0-9]{6,12}\b',  # Harfli prefix
        r'\b\d{3,11}-\d{5,11}\b',  # Tire ile: 205-75999114
        r'Konşimento\s+([A-Z0-9]{8,15})\b',  # Etiket yanında
    ]
    konsimento_list = []
    for pattern in konsimento_patterns:
        matches = re.findall(pattern, icerik, re.IGNORECASE)
        konsimento_list.extend(matches)
    konsimento_list = list(set(konsimento_list))
    
    # AWB (Hava)
    awb_patterns = [
        r'AWB\s*NO\s*(\d{10})\b',  # AWB NO 4613755226
        r'\b\d{3}-\d{8}\b',  # Airline: 235-87235831
    ]
    awb_list = []
    for pattern in awb_patterns:
        matches = re.findall(pattern, icerik, re.IGNORECASE)
        awb_list.extend(matches)
    
    # AWB'leri konşimento listesine ekle
    konsimento_list.extend(awb_list)
    konsimento_list = list(set(konsimento_list))
    
    return {
        "konsimento_list": konsimento_list,
        "konteyner_list": konteyner_list,
        "beyanname_list": beyanname_list,
    }


# ════════════════════════════════════════════════════════════
#  EŞLEŞTIRME (DEĞİŞMEDİ)
# ════════════════════════════════════════════════════════════

def eslestir(numaralar, referans):
    """Numaraları referans ile eşleştir"""
    eslesmeler = []
    
    konsimento_list = numaralar.get("konsimento_list", [])
    konteyner_list = numaralar.get("konteyner_list", [])
    beyanname_list = numaralar.get("beyanname_list", [])
    
    for ref in referans:
        dosya_no = ref["dosya_no"]
        
        # Konteyner eşleşmesi (tam)
        for k in konteyner_list:
            if k.upper() == ref["konteyner"].upper():
                eslesmeler.append({
                    "dosya_no": dosya_no,
                    "kriter": "Konteyner",
                    "deger": k
                })
                break
        
        # Konşimento eşleşmesi (tam)
        for ks in konsimento_list:
            if ks.upper() == ref["konsimento"].upper():
                eslesmeler.append({
                    "dosya_no": dosya_no,
                    "kriter": "Konşimento",
                    "deger": ks
                })
                break
        
        # Beyanname eşleşmesi (son 6 rakam)
        for b in beyanname_list:
            ref_bey = ref["beyanname"]
            if ref_bey and len(ref_bey) >= 6 and len(b) >= 6:
                if ref_bey[-6:] == b[-6:]:
                    eslesmeler.append({
                        "dosya_no": dosya_no,
                        "kriter": "Beyanname (son 6)",
                        "deger": b
                    })
                    break
    
    return eslesmeler


# ════════════════════════════════════════════════════════════
#  ANA PIPELINE
# ════════════════════════════════════════════════════════════

def main():
    log.info("═══ Sistem başladı (IMAP versiyonu) ═══")
    
    conn = db_init()
    referans = sheets_referans_veri()
    
    if not referans:
        log.warning("Referans verisi boş — işlem yapılamaz.")
        return
    
    mailler = gmail_okunmamis_mailler_imap()
    eslesti = beklemeye = okunamadi = 0
    
    for mail in mailler:
        email_id = mail["id"]
        imap_num = mail["imap_num"]
        gonderen = mail["gonderen"]
        konu     = mail["konu"]
        tarih    = mail["tarih"]
        govde    = mail["html_body"]
        
        if db_islendi_mi(conn, email_id):
            log.info(f"Daha önce işlendi, atlanıyor: {email_id}")
            gmail_okundu_isaretle_imap(imap_num)
            continue
        
        log.info(f"Yeni mail işleniyor: {konu} | {gonderen}")
        
        # Mail gövdesinden linki çek
        url = linkten_url_bul(govde)
        
        if not url:
            log.info(f"Link bulunamadı: {konu}")
            gmail_okundu_isaretle_imap(imap_num)
            continue
        
        # Playwright ile sayfayı oku
        icerik = sayfayi_playwright_ile_oku(url)
        
        if not icerik:
            log.error(f"Sayfa okunamadı: {url[:80]}")
            sheets_okunamayanEkle(gonderen, konu, tarih, "Sayfa açılamadı", fatura_url=url)
            sheets_faturaListesiYaz(gonderen, tarih, None, "❌ Okunamadı", fatura_url=url, email_id=email_id)
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
        
        # SADECE REGEX ile numaraları çıkar (Gemini kapalı)
        numaralar = numaralari_regex_ile_cek(icerik)
        
        if not numaralar or not any([
            numaralar["konsimento_list"],
            numaralar["konteyner_list"],
            numaralar["beyanname_list"]
        ]):
            log.info(f"Numara bulunamadı: {konu}")
            sheets_okunamayanEkle(
                gonderen, konu, tarih,
                "Konşimento/Konteyner/Beyanname bulunamadı",
                fatura_url=url
            )
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "❌ Okunamadı", fatura_url=url, email_id=email_id)
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
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "🟡 Bekliyor", fatura_url=url, email_id=email_id)
            beklemeye += 1
        
        gmail_okundu_isaretle_imap(imap_num)
        
        # Her fatura arasında 3 saniye bekle
        time.sleep(3)
    
    # ── Bekleyenleri Sheets'ten yeniden dene ─────────────────
    sheets_serv = sheets_servis()
    bekleyenler = sheets_bekleyenleri_getir(sheets_serv)
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
            sheets_serv, satir_no,
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
