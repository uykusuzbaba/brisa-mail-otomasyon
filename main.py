"""
Brisa Mail Otomasyon — GitHub Actions üzerinde çalışan ana script
Playwright ile JS render edilen fatura sayfalarını okur
"""

import base64
import json
import logging
import os
import re
import sqlite3
import time
from datetime import datetime, timedelta
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
GMAIL_TOKEN_JSON     = os.environ["GMAIL_TOKEN_JSON"]       # token.json içeriği
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

GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets",
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
#  GMAIL
# ════════════════════════════════════════════════════════════

def gmail_servis():
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)


def gmail_okunmamis_mailler(servis):
    gonderen_sorgu = " OR ".join(f"from:{g.strip()}" for g in GONDEREN_LISTESI)
    query = f"is:unread ({gonderen_sorgu})"

    try:
        result = servis.users().messages().list(
            userId="me", q=query, maxResults=50
        ).execute()
    except HttpError as e:
        log.error(f"Gmail liste hatası: {e}")
        return []

    messages = result.get("messages", [])
    detaylar = []

    for msg in messages:
        try:
            detail = servis.users().messages().get(
                userId="me", id=msg["id"], format="full"
            ).execute()
            headers = {
                h["name"]: h["value"]
                for h in detail["payload"].get("headers", [])
            }
            detaylar.append({
                "id": msg["id"],
                "gonderen": headers.get("From", ""),
                "konu": headers.get("Subject", ""),
                "tarih": headers.get("Date", ""),
                "payload": detail["payload"],
            })
        except HttpError as e:
            log.error(f"Mail detay hatası: {e}")

    log.info(f"{len(detaylar)} okunmamış mail bulundu.")
    return detaylar


def gmail_govde_al(payload):
    """Mail gövdesini (HTML) çıkar."""
    def _bul(part):
        mime = part.get("mimeType", "")
        if mime == "text/html":
            data = part.get("body", {}).get("data", "")
            if data:
                return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
        for sub in part.get("parts", []):
            result = _bul(sub)
            if result:
                return result
        return None

    return _bul(payload) or ""


def gmail_okundu_isaretle(servis, message_id):
    try:
        servis.users().messages().modify(
            userId="me", id=message_id,
            body={"removeLabelIds": ["UNREAD"]}
        ).execute()
    except HttpError as e:
        log.warning(f"Okundu işareti hatası: {e}")


def gmail_mail_gonder(servis, konu, html_govde):
    """Bildirim mailini gönder."""
    import email.mime.multipart
    import email.mime.text

    msg = email.mime.multipart.MIMEMultipart("alternative")
    msg["to"] = BILDIRIM_ALICISI
    msg["subject"] = konu
    msg.attach(email.mime.text.MIMEText(html_govde, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        servis.users().messages().send(
            userId="me", body={"raw": raw}
        ).execute()
        log.info(f"Bildirim gönderildi → {BILDIRIM_ALICISI}")
    except HttpError as e:
        log.error(f"Mail gönderme hatası: {e}")


# ════════════════════════════════════════════════════════════
#  PLAYWRIGHT — JS RENDER EDİLEN SAYFA OKUMA
# ════════════════════════════════════════════════════════════

def sayfayi_playwright_ile_oku(url):
    """
    Playwright ile sayfayı tam render et, düz metni döndür.
    JS çalıştıktan sonraki içeriği okur.
    """
    log.info(f"Playwright ile sayfa açılıyor: {url[:80]}...")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"]
            )
            page = browser.new_page()

            # Sayfayı aç, JS yüklenene kadar bekle
            page.goto(url, wait_until="networkidle", timeout=30000)

            # Ek bekleme — dinamik içerik için
            page.wait_for_timeout(2000)

            # Tam sayfa metnini al
            icerik = page.inner_text("body")
            browser.close()

            log.info(f"Sayfa okundu: {len(icerik)} karakter")
            return icerik[:15000]

    except Exception as e:
        log.error(f"Playwright hatası: {e}")
        return None


def firma_adi_cek(icerik):
    """
    Fatura sayfasının düz metninden gönderici firma adını çeker.
    Sayfa metninde "Gönderen" kelimesinin hemen ardından gelen
    büyük harfli şirket adını alır.
    Örnek: "Gönderen YILPORT KONTEYNER TERMİNALİ..."
    """
    if not icerik:
        return ""
    try:
        # Düz metin içinde "Gönderen" kelimesini bul
        pattern = r'G[\xf6o]nderen[^\n]{0,5}([A-Z][A-Z\s\.&,]{5,79})'
        match = re.search(pattern, icerik, re.IGNORECASE)
        if match:
            firma = match.group(1).strip()
            # Vergi no, GUID gibi sayısal şeyleri ele
            # Eğer sadece rakam ve tire ise firma adı değildir
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
    - DLK geçiyorsa: kesinlikle bizim, işleme devam
    - Solmaz veya Subaşı geçiyorsa: kesinlikle bizim değil
    - DENİZ İHRACAT NAVLUNU geçiyorsa: bizim değil
    - Hiçbiri geçmiyorsa: belirsiz, işleme devam (bekleyene düşer)

    Döndürür: (bize_ait: bool, sebep: str)
    """
    metin = (icerik or "").upper() + " " + (html_govde or "").upper()

    # DLK geçiyorsa kesinlikle bizim — diğer kontrollere gerek yok
    if "DLK" in metin:
        return True, ""

    # Deniz ihracat navlunu → bizim değil
    if "DENIZ IHRACAT NAVLUNU" in metin or "DENIZ IHRACAT" in metin:
        return False, "Deniz ihracat navlunu"

    # Konteyner VGM Hizmeti → bizim değil
    if "KONTEYNER VGM" in metin or "VGM HIZMET" in metin:
        return False, "Konteyner VGM Hizmeti"

    # Rakip gümrükçü firmaları → kesinlikle bizim değil
    # Ama önce Brisa'nın kendi email adresini (ithalat.brisa@subasi.net) hariç tut
    metin_email_haric = metin.replace("ITHALAT.BRISA@SUBASI.NET", "").replace("@SUBASI.NET", "")
    
    if "SOLMAZ GUMRUK" in metin_email_haric or "SOLMAZ GÜMRÜK" in metin_email_haric:
        return False, "Solmaz Gümrük Müşavirliği faturası"
    if "SUBASI GUMRUK" in metin_email_haric or "SUBAŞI GÜMRÜK" in metin_email_haric:
        return False, "Subaşı Gümrük Müşavirliği faturası"

    # Hiçbiri geçmiyorsa → belirsiz, normal işleme devam et
    return True, ""


def linkten_url_bul(html_govde):
    """Mail gövdesinden edoksis fatura linkini çıkar."""
    pattern = r'https?://[^\s"\'<>]*edoksis[^\s"\'<>]*'
    match = re.search(pattern, html_govde, re.IGNORECASE)
    if match:
        url = match.group(0).replace("&amp;", "&")
        return url
    return None


# ════════════════════════════════════════════════════════════
#  GEMİNİ — NUMARA ÇIKARIMI
# ════════════════════════════════════════════════════════════

def regex_ile_numaralari_cek(metin):
    """
    Uluslararası standart formatlarla regex tabanlı numara çıkarımı.
    Claude'a gitmeden önce çalışır — hızlı ve ücretsiz.

    Konteyner: ISO 6346 — 4 büyük harf + 7 rakam (ör: MSMU7499454)
    Beyanname: Türk Gümrük — 5 rakam + 2 harf (IM/AN/EX..) + 8 rakam
               mutlaka 00IM0 veya 00AN0 veya 00EX0 gibi bir pattern içerir
    Konşimento: Değişken format — harf+rakam karışımı 8-20 karakter
                "Konşimento", "B/L", "BL" etiketlerinin yanında aranır
    """
    if not metin:
        return None

    metin_upper = metin.upper()

    # ── Konteyner (ISO 6346): 4 büyük harf + 7 rakam ────────
    konteyner_pattern = r'\b([A-Z]{4}[0-9]{7})\b'
    konteynerler = list(set(re.findall(konteyner_pattern, metin_upper)))
    # Fatura no gibi YLP... ile başlayanları ele
    konteynerler = [k for k in konteynerler if not k.startswith('YLP')]

    # ── Beyanname (Türk Gümrük): rakam+IM/AN/EX/IH+rakam ────
    beyanname_pattern = r'\b(\d{5}[A-Z]{2}\d{8})\b'
    beyannameler = list(set(re.findall(beyanname_pattern, metin_upper)))
    # Sadece bilinen Türk gümrük tip kodlarını al
    tip_kodlari = ['IM', 'AN', 'EX', 'IH', 'TR', 'TI', 'AB', 'AT']
    beyannameler = [b for b in beyannameler
                    if any(b[5:7] == tip for tip in tip_kodlari)]

    # ── Konşimento: etiket yanındaki alfanümerik kod ──────────
    konsimentolar = []
    kon_etiket = r'(?:KON[Şs]IMENTO|B/?L|BILL OF LADING|BL NO)[^A-Z0-9]{0,15}([A-Z0-9]{6,25})'
    for m in re.finditer(kon_etiket, metin_upper):
        kod = m.group(1).strip()
        if kod and not kod.startswith('YLP') and not kod.startswith('TR1') and not kod.startswith('TK'):
            konsimentolar.append(kod)
    konsimentolar = list(set(konsimentolar))

    sonuc = {
        "konsimento_list": konsimentolar,
        "konteyner_list":  konteynerler,
        "beyanname_list":  beyannameler,
    }

    hic_yok = not any([konsimentolar, konteynerler, beyannameler])
    return None if hic_yok else sonuc


def claude_ile_numaralari_cek(metin):
    """
    Gemini ile numara çıkarımı.
    Sadece regex başarısız olduğunda çağrılır.
    """
    prompt = (
        "Aşağıdaki metin bir Türk lojistik/gümrük e-faturasına ait sayfa içeriğidir.\n"
        "Bu metinden 3 tür numara çıkar. Numaralar Not satırlarında yazıyor olabilir.\n\n"

        "1. KONŞİMENTO NUMARASI\n"
        "   Etiketler: Konsimento, Konsimento No, B/L, BL, Bill of Lading\n"
        "   Not satirlarinda da olabilir: Not 11: Konsimento :MEDUFB089573\n"
        "   Ornek: SPE041901781, MEDUFB089573, HLCUIST2501XXXXX\n\n"

        "2. KONTEYNER NUMARASI — ISO 6346\n"
        "   Format: 4 buyuk harf + 7 rakam\n"
        "   Virgülle ayrilmis birden fazla olabilir\n"
        "   Ornek: ARKU2438358, MSMU7499454\n\n"

        "3. BEYANNAME NUMARASI — Turk Gümrük\n"
        "   Format: 5 rakam + 2 harf (IM/AN/EX/IH) + 8 rakam\n"
        "   Not satirlarinda: Not 3: Beyanname No: 26410500IM00045602\n"
        "   Ornek: 26410500IM00045784\n\n"

        "KURALLAR:\n"
        "- Fatura no (YLP...), ETTN, GUID, vergi no ALMA\n"
        "- Not satirlarina ozellikle dikkat et\n"
        "- SADECE JSON dondur:\n"
        '{"konsimento_list":[],"konteyner_list":[],"beyanname_list":[]}\n\n'
        f"--- FATURA ---\n{metin[:12000]}\n--- BITIS ---"
    )

    try:
        response = requests.post(
            GEMINI_URL,
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0, "maxOutputTokens": 500}
            },
            timeout=30
        )
        if response.status_code != 200:
            log.error(f"Gemini HTTP {response.status_code}")
            return None

        data = response.json()
        raw = data["candidates"][0]["content"]["parts"][0]["text"]
        temiz = re.sub(r"```json|```", "", raw).strip()
        sonuc = json.loads(temiz)
        sonuc.setdefault("konsimento_list", [])
        sonuc.setdefault("konteyner_list", [])
        sonuc.setdefault("beyanname_list", [])

        hic_yok = not any([
            sonuc["konsimento_list"],
            sonuc["konteyner_list"],
            sonuc["beyanname_list"]
        ])
        return None if hic_yok else sonuc

    except Exception as e:
        log.error(f"Gemini hatası: {e}")
        return None


def gemini_numaralari_kadir(metin):
    """
    Hibrit numara çıkarım motoru:
    1. Önce regex ile standart formatlarda ara (hızlı, ücretsiz)
    2. Bulamazsa Claude Haiku'ya gönder (yavaş, ücretli ama güvenilir)
    """
    # Adım 1: Regex
    sonuc = regex_ile_numaralari_cek(metin)
    if sonuc:
        log.info(f"Regex ile bulundu → "
                 f"Konşimento: {len(sonuc['konsimento_list'])} "
                 f"| Konteyner: {len(sonuc['konteyner_list'])} "
                 f"| Beyanname: {len(sonuc['beyanname_list'])}")
        return sonuc

    # Adım 2: Claude
    log.info("Regex bulamadı, Claude'a gönderiliyor...")
    return claude_ile_numaralari_cek(metin)


# ════════════════════════════════════════════════════════════
#  GOOGLE SHEETS — REFERANS VERİSİ
# ════════════════════════════════════════════════════════════

def sheets_referans_veri():
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    servis = build("sheets", "v4", credentials=creds)
    try:
        result = servis.spreadsheets().values().get(
            spreadsheetId=SHEETS_ID,
            range="📋 Referans!A:D"
        ).execute()
    except HttpError as e:
        log.error(f"Sheets okuma hatası: {e}")
        return []

    rows = result.get("values", [])
    if len(rows) < 2:
        return []

    kayitlar = []
    for row in rows[1:]:
        padded = row + [None] * (4 - len(row))
        dosya = str(padded[3]).strip() if padded[3] else None
        if not dosya:
            continue

        # Konteyner sütununda virgülle ayrılmış birden fazla numara olabilir
        # Örnek: "ARKU2438358, TCKU1234567, MSCU9876543"
        konteyner_ham = str(padded[1]) if padded[1] else ""
        konteyner_listesi = [
            _norm(k) for k in konteyner_ham.split(",") if k.strip()
        ]

        # Konşimento ve Beyanname da virgülle gelebilir
        konsimento_ham = str(padded[0]) if padded[0] else ""
        konsimento_listesi = [
            _norm(k) for k in konsimento_ham.split(",") if k.strip()
        ]

        beyanname_ham = str(padded[2]) if padded[2] else ""
        beyanname_listesi = [
            _norm(k) for k in beyanname_ham.split(",") if k.strip()
        ]

        kayitlar.append({
            "konsimento_listesi": konsimento_listesi,
            "konteyner_listesi":  konteyner_listesi,
            "beyanname_listesi":  beyanname_listesi,
            "dosya_no":           dosya,
        })

    log.info(f"Sheets'ten {len(kayitlar)} kayıt okundu.")
    return kayitlar


def sheets_eslesmeyiKaydet(gonderen, konu, tarih, dosya_no, kriter, deger):
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    servis = build("sheets", "v4", credentials=creds)
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

    try:
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="✅ Eşleşenler!A:I",
            valueInputOption="RAW",
            body={"values": [[
                simdi, tarih, gonderen, konu,
                "fatura-link", dosya_no, kriter, deger
            ]]}
        ).execute()
    except HttpError as e:
        log.error(f"Sheets yazma hatası: {e}")


def sheets_okunamayanEkle(gonderen, konu, tarih, sebep, fatura_url=None):
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    servis = build("sheets", "v4", credentials=creds)
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

    try:
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="⚠️ Okunamayanlar!A:H",
            valueInputOption="RAW",
            body={"values": [[
                simdi, tarih, gonderen, konu,
                "fatura-link", sebep, fatura_url or ""
            ]]}
        ).execute()
    except HttpError as e:
        log.error(f"Sheets okunamayan yazma hatası: {e}")


def sheets_fatura_islendi_mi(servis, email_id):
    """Fatura Listesi'nde bu email_id daha önce yazılmış mı kontrol et."""
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
    Tüm faturaları tek bir listede tutar.
    Email ID kontrolü ile mükerrer kayıt engellenir.
    Durum: "✅ Eşleşti" | "🟡 Bekliyor" | "❌ Okunamadı" | "🔴 Bize Ait Değil"
    """
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    servis = build("sheets", "v4", credentials=creds)
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Mükerrer kontrol
    if email_id and sheets_fatura_islendi_mi(servis, email_id):
        log.info(f"Fatura listesinde zaten var, atlanıyor: {email_id}")
        return

    # Numaraları düz metin olarak yaz
    konsimento = ", ".join(numaralar.get("konsimento_list", [])) if numaralar else ""
    konteyner  = ", ".join(numaralar.get("konteyner_list",  [])) if numaralar else ""
    beyanname  = ", ".join(numaralar.get("beyanname_list",  [])) if numaralar else ""

    try:
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
                "",          # J: İşlemi Yapan
                email_id,    # K: Email ID (mükerrer kontrol için)
            ]]}
        ).execute()
    except HttpError as e:
        log.error(f"Fatura listesi yazma hatası: {e}")


# ════════════════════════════════════════════════════════════
#  EŞLEŞTİRME
# ════════════════════════════════════════════════════════════

def _norm(v):
    if not v:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(v).upper())


def eslestir(numaralar, referans):
    """
    Faturadan çıkarılan numaraları Sheets referans listesiyle karşılaştırır.
    TÜM eşleşmeleri döndürür — aynı konşimento/konteyner için birden fazla
    dosya (beyanname) olabilir, hepsi tek mailde gösterilir.
    """
    f_konsimentolar = [_norm(v) for v in numaralar.get("konsimento_list", []) if v]
    f_konteynerlar  = [_norm(v) for v in numaralar.get("konteyner_list",  []) if v]
    f_beyannameler  = [_norm(v) for v in numaralar.get("beyanname_list",  []) if v]

    eslesmeler = []
    gorulmus_dosyalar = set()  # Aynı dosyayı iki kez eklememek için

    for row in referans:
        dosya = row.get("dosya_no")
        if not dosya or dosya in gorulmus_dosyalar:
            continue

        # Konşimento eşleşmesi
        for r_kon in row.get("konsimento_listesi", []):
            if r_kon and r_kon in f_konsimentolar:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konşimento No", "deger": r_kon})
                gorulmus_dosyalar.add(dosya)
                break

        if dosya in gorulmus_dosyalar:
            continue

        # Konteyner eşleşmesi
        for r_knt in row.get("konteyner_listesi", []):
            if r_knt and r_knt in f_konteynerlar:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Konteyner No", "deger": r_knt})
                gorulmus_dosyalar.add(dosya)
                break

        if dosya in gorulmus_dosyalar:
            continue

        # Beyanname eşleşmesi
        for r_bey in row.get("beyanname_listesi", []):
            if r_bey and r_bey in f_beyannameler:
                eslesmeler.append({"dosya_no": dosya, "kriter": "Beyanname No", "deger": r_bey})
                gorulmus_dosyalar.add(dosya)
                break

    return eslesmeler if eslesmeler else None


# ════════════════════════════════════════════════════════════
#  BİLDİRİM MAİLİ
# ════════════════════════════════════════════════════════════

def bildirim_html(gonderen, konu, tarih, eslesmeler, fatura_url=None):
    """
    eslesmeler: [{"dosya_no": ..., "kriter": ..., "deger": ...}, ...]
    Birden fazla eşleşme tek mailde gösterilir.
    """
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Fatura linki satırı
    url_satiri = ""
    if fatura_url:
        url_satiri = (
            "<tr><td style=\"padding:10px 14px;color:#666;width:40%\">Fatura Linki</td>"
            "<td style=\"padding:10px 14px\">"
            f"<a href=\"{fatura_url}\" style=\"color:#1a5276;font-weight:500\">"
            "🔗 Faturayı Görüntüle →</a></td></tr>"
        )

    # Eşleşen dosyalar bölümü
    dosya_satirlari = ""
    for i, e in enumerate(eslesmeler, 1):
        baslik = f"{i}. EŞLEŞEN DOSYA" if len(eslesmeler) > 1 else "EŞLEŞme DETAYI"
        dosya_satirlari += f"""
    <tr style="background:#eaf0fb"><th colspan="2" style="padding:10px 14px;
        text-align:left;color:#1a5276;font-size:12px">{baslik}</th></tr>
    <tr><td style="padding:10px 14px;color:#666;width:40%">Dosya No</td>
        <td style="padding:10px 14px;font-weight:700;font-size:15px;
            color:#1a5276">{e["dosya_no"]}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Eşleşen Kriter</td>
        <td style="padding:10px 14px;font-weight:700;color:#1a7a4a">{e["kriter"]}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Eşleşen Değer</td>
        <td style="padding:10px 14px;font-weight:700;color:#1a7a4a">{e["deger"]}</td></tr>"""

    # Başlık için dosya no özeti
    if len(eslesmeler) == 1:
        baslik_dosya = f"📂 Dosya No: {eslesmeler[0]['dosya_no']}"
    else:
        dosyalar = ", ".join(e["dosya_no"] for e in eslesmeler)
        baslik_dosya = f"📂 {len(eslesmeler)} Dosya Eşleşti: {dosyalar}"

    return f"""
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
<div style="background:#1a5276;color:white;padding:20px;border-radius:8px 8px 0 0">
  <h2 style="margin:0;font-size:18px">✅ Fatura Eşleşmesi Bulundu</h2>
  <p style="margin:6px 0 0;font-size:13px;opacity:.8">{simdi}</p>
</div>
<div style="background:#f8f9fa;padding:24px;border:1px solid #dee2e6;border-top:none">
  <div style="background:#1a5276;color:white;font-size:20px;font-weight:700;
       padding:18px;border-radius:8px;text-align:center;margin-bottom:20px">
    {baslik_dosya}
  </div>
  <table style="width:100%;border-collapse:collapse;background:white;border-radius:6px">
    <tr style="background:#eaf0fb"><th colspan="2" style="padding:10px 14px;
        text-align:left;color:#1a5276;font-size:12px">FATURA BİLGİLERİ</th></tr>
    <tr><td style="padding:10px 14px;color:#666;width:40%">Gönderen</td>
        <td style="padding:10px 14px;font-weight:500">{gonderen}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Mail Konusu</td>
        <td style="padding:10px 14px;font-weight:500">{konu}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Mail Tarihi</td>
        <td style="padding:10px 14px;font-weight:500">{tarih}</td></tr>
    {url_satiri}
    {dosya_satirlari}
  </table>
</div>
<div style="padding:12px;text-align:center;font-size:12px;color:#888;
     background:#f0f0f0;border:1px solid #dee2e6;border-top:none;
     border-radius:0 0 8px 8px">Brisa Mail Otomasyon Sistemi</div>
</div>""".strip()


# ════════════════════════════════════════════════════════════
#  ANA PIPELINE
# ════════════════════════════════════════════════════════════

def main():
    log.info("═══ Sistem başladı ═══")

    conn = db_init()
    gmail = gmail_servis()
    referans = sheets_referans_veri()

    if not referans:
        log.warning("Referans verisi boş — işlem yapılamaz.")
        return

    mailler = gmail_okunmamis_mailler(gmail)
    eslesti = beklemeye = okunamadi = 0

    for mail in mailler:
        email_id = mail["id"]
        gonderen = mail["gonderen"]
        konu     = mail["konu"]
        tarih    = mail["tarih"]

        if db_islendi_mi(conn, email_id):
            log.info(f"Daha önce işlendi, atlanıyor: {email_id}")
            gmail_okundu_isaretle(gmail, email_id)
            continue

        log.info(f"Yeni mail işleniyor: {konu} | {gonderen}")

        # Mail gövdesinden linki çek
        govde = gmail_govde_al(mail["payload"])
        url   = linkten_url_bul(govde)

        if not url:
            log.info(f"Link bulunamadı: {konu}")
            gmail_okundu_isaretle(gmail, email_id)
            continue

        # Playwright ile sayfayı oku
        icerik = sayfayi_playwright_ile_oku(url)

        if not icerik:
            log.error(f"Sayfa okunamadı: {url[:80]}")
            sheets_okunamayanEkle(gonderen, konu, tarih, "Sayfa açılamadı", fatura_url=url)
            sheets_faturaListesiYaz(gonderen, tarih, None, "❌ Okunamadı", fatura_url=url, email_id=email_id)
            gmail_okundu_isaretle(gmail, email_id)
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
            gmail_okundu_isaretle(gmail, email_id)
            continue

        # Gemini ile numaraları çıkar
        numaralar = gemini_numaralari_kadir(icerik)

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
            gmail_okundu_isaretle(gmail, email_id)
            okunamadi += 1
            continue

        log.info(
            f"Numaralar → Konşimento: {len(numaralar['konsimento_list'])} "
            f"| Konteyner: {len(numaralar['konteyner_list'])} "
            f"| Beyanname: {len(numaralar['beyanname_list'])}"
        )

        # Eşleştir — birden fazla dosya eşleşebilir
        eslesmeler = eslestir(numaralar, referans)

        if eslesmeler:
            # Her eşleşmeyi Sheets'e kaydet
            for e in eslesmeler:
                sheets_eslesmeyiKaydet(
                    gonderen, konu, tarih,
                    e["dosya_no"], e["kriter"], e["deger"]
                )

            # Tüm eşleşmeleri tek mailde gönder
            dosyalar_str = ", ".join(e["dosya_no"] for e in eslesmeler)
            html = bildirim_html(gonderen, konu, tarih, eslesmeler, fatura_url=url)
            gmail_mail_gonder(
                gmail,
                f"✅ Fatura Eşleşmesi — {len(eslesmeler)} Dosya: {dosyalar_str}",
                html
            )
            db_islendi_ekle(conn, email_id, dosyalar_str)
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "✅ Eşleşti",
                                    dosya_no=dosyalar_str, fatura_url=url, email_id=email_id)
            eslesti += 1
        else:
            db_bekleyen_ekle(conn, email_id, gonderen, konu, tarih, numaralar)
            sheets_faturaListesiYaz(firma_adi or gonderen, tarih, numaralar, "🟡 Bekliyor", fatura_url=url, email_id=email_id)
            beklemeye += 1

        gmail_okundu_isaretle(gmail, email_id)

    # Bekleyenleri yeniden dene
    bekleyenler = db_bekleyenleri_getir(conn)
    yeniden_eslesti = 0

    for item in bekleyenler:
        try:
            numaralar_json = item.get("numaralar") or "{}"
            numaralar_item = json.loads(numaralar_json)
        except Exception:
            continue

        eslesmeler_b = eslestir(numaralar_item, referans)
        if not eslesmeler_b:
            continue

        item_gonderen = item.get("gonderen", "")
        item_konu     = item.get("konu", "")
        item_tarih    = item.get("tarih", "")
        item_email_id = item.get("email_id", "")
        item_id       = item.get("id", 0)

        for e in eslesmeler_b:
            sheets_eslesmeyiKaydet(
                item_gonderen, item_konu, item_tarih,
                e["dosya_no"], e["kriter"], e["deger"]
            )

        dosyalar_str_b = ", ".join(e["dosya_no"] for e in eslesmeler_b)
        html = bildirim_html(item_gonderen, item_konu, item_tarih, eslesmeler_b)
        gmail_mail_gonder(
            gmail,
            f"✅ Fatura Eşleşmesi — {len(eslesmeler_b)} Dosya: {dosyalar_str_b}",
            html
        )
        db_islendi_ekle(conn, item_email_id, dosyalar_str_b)
        db_bekleyeni_sil(conn, item_id)
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
