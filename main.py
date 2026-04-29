"""
Brisa Mail Otomasyon — v27 LABEL + TIMEOUT + YENİ FİRMALAR + LİMAN KONTROLÜ
Gmail IMAP (App Password) + Sheets Service Account = Sonsuz token

v27 değişiklikleri:
  - Yeni rakip firmalar: ZHEJIANG HAILIDE, KAMIN, BIRLA CARBON
  - Liman kontrolü: SAN PEDRO, ABIDJAN (POL)
  - Gmail label: "brisa-tedarikci-faturalari"
  - YÜKLEYICI MANTIK DÜZELTMESİ: SHIPPER = Brisa ise → İHRACAT → BİZE AİT DEĞİL
  
v26 özellikleri:
  - Gmail LABEL bazlı arama (inbox yerine)
  - Label yoksa otomatik inbox'a döner
  
v25 özellikleri:
  - httplib2 timeout ayarları (60s)
  - Socket timeout (45s)
  - Exponential backoff retry (3 deneme)
  - Timeout'ta otomatik retry
  
Tüm özellikler dahil (v24'ten):
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
  - Bize ait değil çapraz eleme
"""

import email
import imaplib
import json
import logging
import os
import re
import socket
import sqlite3
import time
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import httplib2
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
GMAIL_LABEL          = os.environ.get("GMAIL_LABEL", "brisa-tedarikci-faturalari")  # Gmail etiket adı
SERVICE_ACCOUNT_JSON = os.environ["SERVICE_ACCOUNT_JSON"]
SHEETS_ID            = os.environ["SHEETS_ID"]
BILDIRIM_ALICISI     = os.environ["BILDIRIM_ALICISI"]
GONDEREN_LISTESI     = os.environ.get("GONDEREN_LISTESI", "m.hizmet@brisa.com.tr").split(",")
SORGU_PENCERESI_GUN  = int(os.environ.get("SORGU_PENCERESI_GUN", "60"))

# ── Sabitler ─────────────────────────────────────────────────
DB_PATH = Path("brisa_mail.db")
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Timeout ayarları (saniye)
SOCKET_TIMEOUT = 45
HTTP_TIMEOUT = 60
MAX_RETRY_ATTEMPTS = 3

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
    # v27'de eklenen firmalar
    "ZHEJIANG HAILIDE",   # ZHEJIANG HAILIDE NEW MATERIAL CO.,LTD
    "HAILIDE",            # Kısa adı
    "KAMIN",              # KAMIN
    "BIRLA CARBON",       # BIRLA CARBON EGYPT S.AE
]

# Eleme limanları (POL = Port of Loading)
ELEME_LIMANLARI = [
    "SAN PEDRO",          # Fildişi Sahili (SAPH, SOUTHLAND için)
    "ABIDJAN",            # Fildişi Sahili (SAPH, SOUTHLAND için)
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
    """
    IMAP ile okunmamış Brisa maillerini çek
    ✅ v26: Label bazlı arama (inbox yerine)
    """
    mail = gmail_imap_baglanti()
    
    # Label'ı seç (Gmail'de "Brisa/Fatura" → IMAP'te "Brisa.Fatura" olarak görünür)
    label_imap = GMAIL_LABEL.replace("/", ".")
    
    try:
        # Önce label'ı seç
        status, _ = mail.select(f'"{label_imap}"')
        if status != "OK":
            # Label yoksa inbox'a dön
            log.warning(f"⚠️ Label '{label_imap}' bulunamadı, inbox kullanılıyor")
            mail.select("inbox")
    except Exception as e:
        log.warning(f"⚠️ Label seçme hatası: {e}, inbox kullanılıyor")
        mail.select("inbox")
    
    detaylar = []

    for gonderen in GONDEREN_LISTESI:
        try:
            status, messages = mail.search(None, 'UNSEEN', f'FROM "{gonderen.strip()}"')
            if status != "OK":
                continue

            msg_nums = messages[0].split()
            log.info(f"📬 {gonderen} → {len(msg_nums)} okunmamış mail (label: {label_imap})")

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
#  SHEETS (Service Account — sonsuz token + TIMEOUT FIX)
# ════════════════════════════════════════════════════════════

def _sheets_creds():
    return service_account.Credentials.from_service_account_info(
        json.loads(SERVICE_ACCOUNT_JSON), scopes=SHEETS_SCOPES
    )


def sheets_servis():
    """
    ✅ v25: httplib2 timeout + socket timeout ayarları
    """
    # Socket seviyesi timeout
    socket.setdefaulttimeout(SOCKET_TIMEOUT)
    
    # HTTP client timeout
    http = httplib2.Http(timeout=HTTP_TIMEOUT)
    
    creds = _sheets_creds()
    return build("sheets", "v4", credentials=creds, http=http)


def _sheets_retry_wrapper(func, *args, **kwargs):
    """
    Exponential backoff ile retry wrapper
    Timeout, 503, 429 hatalarında tekrar dener
    """
    for attempt in range(MAX_RETRY_ATTEMPTS):
        try:
            return func(*args, **kwargs)
        
        except (socket.timeout, TimeoutError) as e:
            wait_time = 2 ** attempt * 5  # 5s, 10s, 20s
            log.warning(f"⏱️ Timeout hatası (deneme {attempt+1}/{MAX_RETRY_ATTEMPTS}): {e}")
            if attempt < MAX_RETRY_ATTEMPTS - 1:
                log.info(f"🔄 {wait_time}s bekleyip yeniden deneniyor...")
                time.sleep(wait_time)
                continue
            else:
                log.error(f"❌ Maksimum deneme sayısına ulaşıldı")
                raise
        
        except HttpError as e:
            if e.resp.status in [503, 429]:
                wait_time = 2 ** attempt * 5
                log.warning(f"⚠️ Sheets geçici hata {e.resp.status} (deneme {attempt+1}/{MAX_RETRY_ATTEMPTS})")
                if attempt < MAX_RETRY_ATTEMPTS - 1:
                    log.info(f"🔄 {wait_time}s bekleyip yeniden deneniyor...")
                    time.sleep(wait_time)
                    continue
            log.error(f"❌ Sheets HTTP hatası: {e}")
            raise
        
        except Exception as e:
            log.error(f"❌ Beklenmeyen hata: {type(e).__name__}: {e}")
            raise
    
    # Buraya asla gelmemeli ama güvenlik için
    raise Exception(f"Retry wrapper başarısız oldu ({MAX_RETRY_ATTEMPTS} deneme)")


def sheets_referans_veri():
    """
    📋 Referans sekmesi: A=Konşimento, B=Konteyner, C=Beyanname, D=Dosya No, E=Kullanıcı
    v25: Timeout korumalı retry mekanizması
    """
    def _read_sheet():
        servis = sheets_servis()
        return servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📋 Referans!A:E"
        ).execute()
    
    try:
        result = _sheets_retry_wrapper(_read_sheet)
    except Exception as e:
        log.error(f"Sheets referans okunamadı: {e}")
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
    """v25: Timeout korumalı"""
    def _read_column():
        return servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID, range="📑 Fatura Listesi!K:K"
        ).execute()
    
    try:
        result = _sheets_retry_wrapper(_read_column)
        for row in result.get("values", []):
            if row and str(row[0]).strip() == str(email_id).strip():
                return True
    except Exception:
        pass
    return False


def sheets_yeni_fatura_ekle(servis, satir_verisi):
    """v25: Timeout korumalı append"""
    def _append_row():
        return servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="📑 Fatura Listesi!A:L",
            valueInputOption="RAW",
            body={"values": [satir_verisi]}
        ).execute()
    
    try:
        _sheets_retry_wrapper(_append_row)
        log.info(f"✅ Fatura eklendi: {satir_verisi[6]}")
    except Exception as e:
        log.error(f"❌ Fatura eklenemedi: {e}")


def sheets_bekleyeni_guncelle(servis, email_id, yeni_durum, yeni_dosya_no=""):
    """v25: Timeout korumalı update"""
    def _read_list():
        return servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID, range="📑 Fatura Listesi!A:L"
        ).execute()
    
    try:
        result = _sheets_retry_wrapper(_read_list)
    except Exception as e:
        log.error(f"Bekleyen güncelleme için okuma hatası: {e}")
        return

    rows = result.get("values", [])
    for idx, row in enumerate(rows[1:], start=2):
        padded = row + [""] * 12
        if str(padded[10]).strip() == str(email_id).strip():
            def _update_row():
                if yeni_dosya_no:
                    padded[7] = yeni_dosya_no
                padded[6] = yeni_durum
                
                return servis.spreadsheets().values().update(
                    spreadsheetId=SHEETS_ID,
                    range=f"📑 Fatura Listesi!A{idx}:L{idx}",
                    valueInputOption="RAW",
                    body={"values": [padded]}
                ).execute()
            
            try:
                _sheets_retry_wrapper(_update_row)
                log.info(f"✅ Bekleyen güncellendi: {email_id} → {yeni_durum}")
            except Exception as e:
                log.error(f"Bekleyen güncelleme hatası: {e}")
            return


def sheets_bekleyeni_aitdegil_guncelle(servis, email_id):
    """v25: Timeout korumalı"""
    sheets_bekleyeni_guncelle(servis, email_id, "🔴 Bize Ait Değil")


def sheets_bekleyen_satirlar():
    """
    📑 Fatura Listesi'nden "🟡 Bekleyen" satırları döndürür
    v25: Timeout korumalı
    """
    def _read_list():
        servis = sheets_servis()
        return servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID, range="📑 Fatura Listesi!A:L"
        ).execute()
    
    try:
        result = _sheets_retry_wrapper(_read_list)
    except Exception as e:
        log.error(f"Bekleyen satırlar okunamadı: {e}")
        return []

    rows = result.get("values", [])
    bekleyenler = []
    
    for idx, row in enumerate(rows[1:], start=2):
        padded = row + [""] * 12
        durum = str(padded[6]).strip()
        
        if durum == "🟡 Bekleyen":
            bekleyenler.append({
                "satir_no":   idx,
                "email_id":   str(padded[10]).strip(),
                "konsimento": str(padded[3]).strip(),
                "konteyner":  str(padded[4]).strip(),
                "beyanname":  str(padded[5]).strip(),
                "fatura_url": str(padded[8]).strip() if len(padded) > 8 else "",
            })
    
    return bekleyenler


def sheets_aitdegil_numaralar():
    """
    📑 Fatura Listesi'nden "🔴 Bize Ait Değil" satırlarının numara setini döndürür
    v25: Timeout korumalı
    """
    def _read_list():
        servis = sheets_servis()
        return servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID, range="📑 Fatura Listesi!A:L"
        ).execute()
    
    try:
        result = _sheets_retry_wrapper(_read_list)
    except Exception as e:
        log.error(f"Ait değil numaralar okunamadı: {e}")
        return set()

    rows = result.get("values", [])
    numaralar = set()
    
    for row in rows[1:]:
        padded = row + [""] * 12
        durum = str(padded[6]).strip()
        
        if durum == "🔴 Bize Ait Değil":
            for col_idx in [3, 4, 5]:  # D=Konşimento, E=Konteyner, F=Beyanname
                val = str(padded[col_idx]).strip()
                if val:
                    numaralar.add(_norm(val))
    
    return numaralar


# ════════════════════════════════════════════════════════════
#  YARDIMCI FONKSİYONLAR
# ════════════════════════════════════════════════════════════

def _norm(text):
    """Türkçe karakter normalizasyonu + büyük harf"""
    tr_map = str.maketrans("ıİöÖüÜşŞğĞçÇ", "iIOOUUSsGgCC")
    return str(text).strip().translate(tr_map).upper()


def _beyanname_son_6(numara):
    """Beyanname son 6 rakamı çıkar"""
    rakamlar = re.findall(r'\d', numara)
    return "".join(rakamlar[-6:]) if len(rakamlar) >= 6 else ""


def edoksis_link_mi(url):
    """edoksis linki mi kontrol et"""
    return "edoksis.com.tr" in url.lower()


def edoksis_icerik_cek(url):
    """
    Playwright ile edoksis linkini render et
    page.inner_text("body") → düz metin (HTML tag'leri olmadan)
    """
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            time.sleep(3)  # Rate limit önleme
            
            # HTML yerine düz metin
            metin = page.inner_text("body")
            
            browser.close()
            return metin
    except Exception as e:
        log.error(f"edoksis render hatası: {e}")
        return ""


def rakip_firma_mi(metin):
    """Rakip firma kontrolü (case-insensitive)"""
    if not metin:
        return False
    ust = _norm(metin)
    for firma in RAKIP_FIRMALAR:
        if _norm(firma) in ust:
            return True
    return False


def brisa_mi(metin):
    """
    Metin içinde BRISA BRIDGESTONE SABUNCİ LASTİK geçiyor mu?
    Gerekirse şirket adı güncellenebilir.
    """
    if not metin:
        return False
    ust = _norm(metin)
    return "BRISA" in ust or "BRIDGESTONE" in ust or "SABUNCI" in ust or "LASTIK" in ust


def bize_ait_degil_mi(fatura_metni):
    """
    Eleme kriterleri (v27 — DÜZELTİLMİŞ):
      1. Email temizleme: ithalat.brisa@subasi.net hariç
      2. Eğer DLK geçiyorsa → kesinlikle bizim
      3. Yoksa: Solmaz/Subaşı/VGM/Deniz İhracat kontrol
      4. Rakip firma kontrolü (CABOT/MITSUI/SOUTHLAND/HAILIDE/KAMIN/BIRLA CARBON)
      5. Liman kontrolü (POL: SAN PEDRO, ABIDJAN)
      6. Yükleyici/Gönderici kontrolü (SHIPPER = Brisa ise → İHRACAT → Bize ait değil)
    """
    if not fatura_metni:
        return False

    # Email temizleme (case-insensitive)
    if "ithalat.brisa@subasi.net".lower() in fatura_metni.lower():
        log.info("📧 Brisa email adresi bulundu → bizim")
        return False

    # DLK geçiyorsa kesinlikle bizim
    if "DLK" in _norm(fatura_metni):
        log.info("🏢 DLK bulundu → kesinlikle bizim")
        return False

    ust = _norm(fatura_metni)

    # Gümrük firmaları
    if any(x in ust for x in ["SOLMAZ GUMRUK", "SOLMAZ GÜMRÜK", "SUBASI GUMRUK", "SUBASI GÜMRÜK"]):
        log.info("🚫 Rakip gümrük firması tespit edildi")
        return True

    # VGM / Deniz İhracat
    if "KONTEYNER VGM HIZMETI" in ust or "DENIZ IHRACAT NAVLUNU" in ust:
        log.info("🚫 VGM/Deniz İhracat navlunu tespit edildi")
        return True

    # Rakip firma kontrolü (v27 genişletildi: HAILIDE, KAMIN, BIRLA CARBON)
    if rakip_firma_mi(fatura_metni):
        log.info("🚫 Rakip firma tespit edildi")
        return True

    # Liman kontrolü (POL - Port of Loading)
    # "Yükleme Limanı - POL: SAN PEDRO" veya "POL: ABIDJAN"
    for liman in ELEME_LIMANLARI:
        liman_norm = _norm(liman)
        # POL yanında veya "Yükleme Limanı" yanında geçiyor mu?
        if re.search(rf"(?:POL|YUKLEME\s*LIMANI)[:\s-]*{re.escape(liman_norm)}", ust):
            log.info(f"🚫 Eleme limanı tespit edildi: {liman}")
            return True

    # Yükleyici/Gönderici kontrolü (İHRACAT ELEMESİ)
    # Eğer SHIPPER/CONSIGNOR = Brisa ise → İHRACAT faturası → BİZE AİT DEĞİL
    shipper_pattern = r"(?:SHIPPER|CONSIGNOR|YUKLEYICI|GONDERICI)[:\s]*([^\n]+)"
    shipper_match = re.search(shipper_pattern, ust, re.IGNORECASE)
    if shipper_match:
        shipper_text = shipper_match.group(1).strip()
        if brisa_mi(shipper_text):
            log.info(f"🚫 Yükleyici Brisa (İhracat): {shipper_text[:50]}")
            return True
        else:
            log.info(f"✅ Yükleyici Brisa değil (İthalat): {shipper_text[:50]}")

    return False


def numara_cikart(text):
    """
    SADECE REGEX ile numara çıkarma (Gemini kapalı)
    Desteklenen formatlar:
      - Konteyner (ISO 6346): ARKU2438358
      - Beyanname: 26410500IM00045784, 26/IM0226949, Bey.No 228183
      - Konşimento: Konşimento SPE041901781, 205-75999114, MEDUKC378982
      - AWB: AWB NO 4613755226, 235-87235831
    """
    konteyner = set()
    beyanname = set()
    konsimento = set()

    # ── KONTEYNER (ISO 6346) ──
    # 4 harf + 7 rakam
    for m in re.finditer(r'\b([A-Z]{4}[0-9]{7})\b', text):
        konteyner.add(m.group(1))

    # ── BEYANNAME ──
    # Standart: 26410500IM00045784
    for m in re.finditer(r'\b(\d{5}(?:IM|AN|EX|IH|TR|TI|AB|AT|EI)\d{8})\b', text):
        beyanname.add(m.group(1))
    
    # Slash: 26/IM0226949
    for m in re.finditer(r'\b(\d{2}/[A-Z]{2}\d{7,8})\b', text):
        beyanname.add(m.group(1))
    
    # Kısa format: Bey.No 228183
    for m in re.finditer(r'BEY\.?\s*NO[:\s]*(\d{6})', text, re.IGNORECASE):
        beyanname.add(m.group(1))

    # ── KONŞIMENTO ──
    # Etiket yanında: "Konşimento SPE041901781"
    for m in re.finditer(r'KON[SŞ]IMENTO[:\s]*([A-Z0-9]+)', text, re.IGNORECASE):
        konsimento.add(m.group(1))
    
    # Tire formatı: 205-75999114, 61598712294-11
    for m in re.finditer(r'\b(\d{3,11}-\d{5,11})\b', text):
        konsimento.add(m.group(1))
    
    # Harfli prefix: MEDUKC378982, 26HU121000
    for m in re.finditer(r'\b((?:MEDU|SPE|HLCU|MSK|ONE|CMA)[A-Z0-9]{6,12})\b', text):
        konsimento.add(m.group(1))

    # ── AWB (Hava) ──
    # AWB NO 4613755226 (10 rakam)
    for m in re.finditer(r'AWB[/\s]*(?:NO|HIZMET)?[:\s]*(\d{10})', text, re.IGNORECASE):
        konsimento.add(m.group(1))
    
    # Airline format: 235-87235831
    for m in re.finditer(r'\b(\d{3}-\d{8})\b', text):
        konsimento.add(m.group(1))

    # ── Virgüllü konşimento parse ──
    # "61598712294-11, 61598712294-12, ..." → her biri ayrı
    for k in list(konsimento):
        if "," in k:
            konsimento.remove(k)
            for part in k.split(","):
                part = part.strip()
                if part:
                    konsimento.add(part)

    # ── HAWB / House AWB ──
    # "HAWB 850600038" → AWB benzeri
    for m in re.finditer(r'(?:HAWB|HOUSE\s*AWB)[:\s]*([A-Z0-9]+)', text, re.IGNORECASE):
        konsimento.add(m.group(1))

    return {
        "konteyner": sorted(konteyner),
        "beyanname": sorted(beyanname),
        "konsimento": sorted(konsimento),
    }


def numaralarla_eslesme_bul(numaralar, referanslar):
    """
    Beyanname için esnek eşleştirme: son 6 rakam
    Konteyner/Konşimento: tam eşleşme
    """
    def _ref_match(ref_list, fatura_val, is_beyanname=False):
        fatura_norm = _norm(fatura_val)
        
        if is_beyanname:
            # Son 6 rakam eşleştirme
            fatura_son_6 = _beyanname_son_6(fatura_norm)
            if not fatura_son_6:
                return False
            
            for ref in ref_list:
                ref_norm = _norm(ref)
                ref_son_6 = _beyanname_son_6(ref_norm)
                if ref_son_6 and fatura_son_6 == ref_son_6:
                    return True
        else:
            # Tam eşleşme (konteyner/konşimento)
            if fatura_norm in [_norm(r) for r in ref_list]:
                return True
        
        return False

    for ref in referanslar:
        # Beyanname kontrolü (esnek)
        for bey in numaralar.get("beyanname", []):
            if _ref_match(ref["beyanname_listesi"], bey, is_beyanname=True):
                return ref["dosya_no"], ref["kullanici"]
        
        # Konteyner kontrolü (tam)
        for kon in numaralar.get("konteyner", []):
            if _ref_match(ref["konteyner_listesi"], kon):
                return ref["dosya_no"], ref["kullanici"]
        
        # Konşimento kontrolü (tam)
        for kons in numaralar.get("konsimento", []):
            if _ref_match(ref["konsimento_listesi"], kons):
                return ref["dosya_no"], ref["kullanici"]
    
    return None, ""


# ════════════════════════════════════════════════════════════
#  ANA İŞLEMLER
# ════════════════════════════════════════════════════════════

def yeni_mailleri_isle(conn, referanslar):
    """Yeni gelen mailleri işle"""
    log.info("═══ Yeni mailler kontrol ediliyor ═══")
    
    mailler = gmail_okunmamis_mailler_imap()
    if not mailler:
        log.info("Yeni mail yok")
        return

    servis = sheets_servis()
    
    # v24: Bize ait değil numaraları çek
    aitdegil_set = sheets_aitdegil_numaralar()
    log.info(f"📋 {len(aitdegil_set)} adet 'Bize Ait Değil' numara yüklendi")

    for mail in mailler:
        email_id = mail["id"]
        
        # Mükerrer kontrolü
        if db_islendi_mi(conn, email_id):
            log.info(f"⏭️ Zaten işlendi: {email_id}")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue
        
        if sheets_fatura_islendi_mi(servis, email_id):
            log.info(f"⏭️ Sheets'te var: {email_id}")
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue

        log.info(f"\n📧 İşleniyor: {mail['konu'][:60]}")
        
        # edoksis linki bul
        link_match = re.search(r'(https?://[^\s"<>]+edoksis\.com\.tr[^\s"<>]*)', mail["html_body"])
        if not link_match:
            log.warning("❌ edoksis linki bulunamadı")
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue
        
        fatura_url = link_match.group(1)
        log.info(f"🔗 Link: {fatura_url[:80]}")
        
        # Render et
        fatura_metin = edoksis_icerik_cek(fatura_url)
        if not fatura_metin:
            log.warning("❌ İçerik çekilemedi")
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue

        # Bize ait değil kontrolü
        if bize_ait_degil_mi(fatura_metin):
            log.info("🔴 Eleme: Bize ait değil")
            satir = [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                mail["tarih"][:16] if mail["tarih"] else "",
                mail["gonderen"][:50],
                "", "", "",  # Konşimento, Konteyner, Beyanname
                "🔴 Bize Ait Değil",
                "",  # Dosya No
                fatura_url,
                "",  # İşlemi Yapan
                email_id,
                ""   # Kullanıcı
            ]
            sheets_yeni_fatura_ekle(servis, satir)
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue

        # Numara çıkart
        numaralar = numara_cikart(fatura_metin)
        log.info(f"🔢 Çıkarılan: Kon={len(numaralar['konsimento'])}, "
                 f"Kont={len(numaralar['konteyner'])}, Bey={len(numaralar['beyanname'])}")

        if not any([numaralar["konsimento"], numaralar["konteyner"], numaralar["beyanname"]]):
            log.warning("❌ Hiç numara bulunamadı → Okunamayan")
            satir = [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                mail["tarih"][:16] if mail["tarih"] else "",
                mail["gonderen"][:50],
                "", "", "",
                "❌ Okunamayan",
                "",
                fatura_url,
                "",
                email_id,
                ""
            ]
            sheets_yeni_fatura_ekle(servis, satir)
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue

        # v24: Çapraz eleme kontrolü
        capraz_eslesti = False
        for num_list in [numaralar["konsimento"], numaralar["konteyner"], numaralar["beyanname"]]:
            for num in num_list:
                if _norm(num) in aitdegil_set:
                    log.info(f"🔴 Çapraz eleme: {num} → Bize ait değil listesinde")
                    capraz_eslesti = True
                    break
            if capraz_eslesti:
                break

        if capraz_eslesti:
            satir = [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                mail["tarih"][:16] if mail["tarih"] else "",
                mail["gonderen"][:50],
                ", ".join(numaralar["konsimento"][:3]),
                ", ".join(numaralar["konteyner"][:3]),
                ", ".join(numaralar["beyanname"][:3]),
                "🔴 Bize Ait Değil",
                "",
                fatura_url,
                "",
                email_id,
                ""
            ]
            sheets_yeni_fatura_ekle(servis, satir)
            db_islendi_ekle(conn, email_id, "")
            gmail_okundu_isaretle_imap(mail.get("imap_num", ""))
            continue

        # Eşleştirme dene
        dosya_no, kullanici = numaralarla_eslesme_bul(numaralar, referanslar)

        if dosya_no:
            log.info(f"✅ Eşleşti: {dosya_no} ({kullanici})")
            durum = "✅ Eşleşenler"
        else:
            log.info("🟡 Eşleşmedi → Bekleyen")
            durum = "🟡 Bekleyen"
            db_bekleyen_ekle(conn, email_id, mail["gonderen"], mail["konu"], mail["tarih"], numaralar)

        satir = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            mail["tarih"][:16] if mail["tarih"] else "",
            mail["gonderen"][:50],
            ", ".join(numaralar["konsimento"][:3]),
            ", ".join(numaralar["konteyner"][:3]),
            ", ".join(numaralar["beyanname"][:3]),
            durum,
            dosya_no or "",
            fatura_url,
            "",
            email_id,
            kullanici
        ]
        
        sheets_yeni_fatura_ekle(servis, satir)
        db_islendi_ekle(conn, email_id, dosya_no or "")
        gmail_okundu_isaretle_imap(mail.get("imap_num", ""))


def bekleyenleri_kontrol_et(conn, referanslar):
    """
    Bekleyen faturaları yeniden kontrol et
    v24: 3 aşamalı eleme
      1. Referans eşleştirme
      2. Bize ait değil çapraz eleme
      3. 6 saatten eski bekleyenler → bize ait değil
    """
    log.info("\n═══ Bekleyenler kontrol ediliyor ═══")
    
    bekleyenler = sheets_bekleyen_satirlar()
    if not bekleyenler:
        log.info("Bekleyen yok")
        return

    log.info(f"📋 {len(bekleyenler)} bekleyen satır bulundu")
    
    servis = sheets_servis()
    aitdegil_set = sheets_aitdegil_numaralar()
    
    simdi = datetime.now()
    
    for satir in bekleyenler:
        email_id = satir["email_id"]
        log.info(f"\n🔄 Bekleyen kontrol: {email_id}")
        
        # DB'den detayları al
        db_row = conn.execute(
            "SELECT numaralar, created_at FROM pending WHERE email_id=?",
            (email_id,)
        ).fetchone()
        
        if not db_row:
            log.warning(f"⚠️ DB'de bulunamadı: {email_id}")
            continue
        
        numaralar = json.loads(db_row[0])
        created_at = datetime.fromisoformat(db_row[1])
        
        # 1. AŞAMA: Referans eşleştirme
        dosya_no, kullanici = numaralarla_eslesme_bul(numaralar, referanslar)
        if dosya_no:
            log.info(f"✅ Eşleşti: {dosya_no} ({kullanici})")
            sheets_bekleyeni_guncelle(servis, email_id, "✅ Eşleşenler", dosya_no)
            
            # DB'den sil
            conn.execute("DELETE FROM pending WHERE email_id=?", (email_id,))
            conn.commit()
            continue
        
        # 2. AŞAMA: Çapraz eleme
        capraz_eslesti = False
        for num_list in [numaralar.get("konsimento", []), numaralar.get("konteyner", []), numaralar.get("beyanname", [])]:
            for num in num_list:
                if _norm(num) in aitdegil_set:
                    log.info(f"🔴 Çapraz eleme: {num} → Bize ait değil listesinde")
                    capraz_eslesti = True
                    break
            if capraz_eslesti:
                break
        
        if capraz_eslesti:
            sheets_bekleyeni_aitdegil_guncelle(servis, email_id)
            conn.execute("DELETE FROM pending WHERE email_id=?", (email_id,))
            conn.commit()
            continue
        
        # 3. AŞAMA: 6 saatten eski bekleyenler
        bekleyen_sure = simdi - created_at
        if bekleyen_sure > timedelta(hours=6):
            log.info(f"⏰ 6 saatten eski bekleyen ({bekleyen_sure}) → Bize ait değil")
            sheets_bekleyeni_aitdegil_guncelle(servis, email_id)
            conn.execute("DELETE FROM pending WHERE email_id=?", (email_id,))
            conn.commit()
            continue
        
        log.info(f"🟡 Hala bekliyor ({bekleyen_sure})")


# ════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════

def main():
    log.info("═══ Sistem başladı (IMAP + Service Account) v27 — YENİ FİRMALAR + LİMAN + LABEL ═══")
    
    try:
        # Socket timeout ayarla
        socket.setdefaulttimeout(SOCKET_TIMEOUT)
        log.info(f"⏱️ Socket timeout: {SOCKET_TIMEOUT}s, HTTP timeout: {HTTP_TIMEOUT}s")
        
        # DB
        conn = db_init()
        log.info("✅ SQLite hazır")
        
        # Referans yükle
        referans = sheets_referans_veri()
        if not referans:
            log.warning("⚠️ Referans listesi boş, sadece bekleyenleri kontrol ediyorum")
        
        # Yeni mailler
        yeni_mailleri_isle(conn, referans)
        
        # Bekleyenler
        bekleyenleri_kontrol_et(conn, referans)
        
        conn.close()
        log.info("✅ İşlemler tamamlandı")
    
    except Exception as e:
        log.error(f"❌ Kritik hata: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
