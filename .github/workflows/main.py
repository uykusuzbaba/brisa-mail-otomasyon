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

import gspread
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
            userId="me", q=query, maxResults=30
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

def gemini_numaralari_kadir(metin):
    prompt = (
        "Aşağıdaki metin bir lojistik/gümrük faturasına ait sayfa içeriğidir.\n"
        "Bu metinden 3 tür numara çıkarmanı istiyorum.\n\n"

        "1. KONŞİMENTO NUMARASI:\n"
        "   Konşimento, Konşimento No, B/L, BL No etiketlerinin yanındaki kodlar\n"
        "   Örnek: SPE041901781, HLCUIST2501XXXXX\n\n"

        "2. KONTEYNER NUMARASI:\n"
        "   Tam olarak 4 büyük harf + 7 rakam formatındaki kodlar\n"
        "   Etiket olmadan sadece numara da yazabilir\n"
        "   Örnek: ARKU2438358, TCKU1234567\n\n"

        "3. BEYANNAME NUMARASI:\n"
        "   Beyanname, Beyanname No, Gümrük Beyannamesi etiketlerinin yanındaki kodlar\n"
        "   Türk gümrük formatı: rakam+harf karışımı\n"
        "   Örnek: 26410500IM00045784\n\n"

        "ÖNEMLİ KURALLAR:\n"
        "- Etiket olmasa bile 4 harf+7 rakam formatındaki her kodu konteyner olarak al\n"
        "- Fatura no, vergi no, GUID gibi alakasız numaraları alma\n"
        "- Aynı numara birden fazla geçiyorsa bir kez yaz\n\n"

        "YANIT FORMATI — sadece bu JSON, başka hiçbir şey:\n"
        '{"konsimento_list":["SPE041901781"],'
        '"konteyner_list":["ARKU2438358"],'
        '"beyanname_list":["26410500IM00045784"]}\n\n'
        f"--- METİN BAŞLANGIÇ ---\n{metin}\n--- METİN BİTİŞ ---"
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
            log.error(f"Gemini HTTP {response.status_code}: {response.text[:200]}")
            return None

        data = response.json()
        raw = data["candidates"][0]["content"]["parts"][0]["text"]
        temiz = re.sub(r"```json|```", "", raw).strip()
        sonuc = json.loads(temiz)

        sonuc.setdefault("konsimento_list", [])
        sonuc.setdefault("konteyner_list", [])
        sonuc.setdefault("beyanname_list", [])
        return sonuc

    except Exception as e:
        log.error(f"Gemini hatası: {e}")
        return None


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
        kayitlar.append({
            "konsimento": _norm(padded[0]),
            "konteyner":  _norm(padded[1]),
            "beyanname":  _norm(padded[2]),
            "dosya_no":   str(padded[3]).strip() if padded[3] else None,
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


def sheets_okunamayanEkle(gonderen, konu, tarih, sebep):
    token_data = json.loads(GMAIL_TOKEN_JSON)
    creds = Credentials.from_authorized_user_info(token_data, GMAIL_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    servis = build("sheets", "v4", credentials=creds)
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")

    try:
        servis.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID,
            range="⚠️ Okunamayanlar!A:G",
            valueInputOption="RAW",
            body={"values": [[
                simdi, tarih, gonderen, konu,
                "fatura-link", sebep
            ]]}
        ).execute()
    except HttpError as e:
        log.error(f"Sheets okunamayan yazma hatası: {e}")


# ════════════════════════════════════════════════════════════
#  EŞLEŞTİRME
# ════════════════════════════════════════════════════════════

def _norm(v):
    if not v:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(v).upper())


def eslestir(numaralar, referans):
    konsimentolar = [_norm(v) for v in numaralar.get("konsimento_list", []) if v]
    konteynerlar  = [_norm(v) for v in numaralar.get("konteyner_list",  []) if v]
    beyannameler  = [_norm(v) for v in numaralar.get("beyanname_list",  []) if v]

    for row in referans:
        dosya = row.get("dosya_no")
        if not dosya:
            continue

        if row["konsimento"] and row["konsimento"] in konsimentolar:
            return {"dosya_no": dosya, "kriter": "Konşimento No", "deger": row["konsimento"]}
        if row["konteyner"] and row["konteyner"] in konteynerlar:
            return {"dosya_no": dosya, "kriter": "Konteyner No",  "deger": row["konteyner"]}
        if row["beyanname"] and row["beyanname"] in beyannameler:
            return {"dosya_no": dosya, "kriter": "Beyanname No",  "deger": row["beyanname"]}

    return None


# ════════════════════════════════════════════════════════════
#  BİLDİRİM MAİLİ
# ════════════════════════════════════════════════════════════

def bildirim_html(gonderen, konu, tarih, dosya_no, kriter, deger):
    simdi = datetime.now().strftime("%d.%m.%Y %H:%M")
    return f"""
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
<div style="background:#1a5276;color:white;padding:20px;border-radius:8px 8px 0 0">
  <h2 style="margin:0;font-size:18px">✅ Fatura Eşleşmesi Bulundu</h2>
  <p style="margin:6px 0 0;font-size:13px;opacity:.8">{simdi}</p>
</div>
<div style="background:#f8f9fa;padding:24px;border:1px solid #dee2e6;border-top:none">
  <div style="background:#1a5276;color:white;font-size:24px;font-weight:700;
       padding:18px;border-radius:8px;text-align:center;margin-bottom:20px">
    📂 Dosya No: {dosya_no}
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
    <tr style="background:#eaf0fb"><th colspan="2" style="padding:10px 14px;
        text-align:left;color:#1a5276;font-size:12px">EŞLEŞme DETAYI</th></tr>
    <tr><td style="padding:10px 14px;color:#666">Eşleşen Kriter</td>
        <td style="padding:10px 14px;font-weight:700;color:#1a7a4a">{kriter}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Eşleşen Değer</td>
        <td style="padding:10px 14px;font-weight:700;color:#1a7a4a">{deger}</td></tr>
    <tr><td style="padding:10px 14px;color:#666">Dosya No</td>
        <td style="padding:10px 14px;font-weight:700;font-size:16px;
            color:#1a5276">{dosya_no}</td></tr>
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
            gmail_okundu_isaretle(gmail, email_id)
            continue

        # Mail gövdesinden linki bul
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
            sheets_okunamayanEkle(gonderen, konu, tarih, "Sayfa açılamadı")
            gmail_okundu_isaretle(gmail, email_id)
            okunamadi += 1
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
                "Konşimento/Konteyner/Beyanname bulunamadı"
            )
            gmail_okundu_isaretle(gmail, email_id)
            okunamadi += 1
            continue

        log.info(
            f"Numaralar → Konşimento: {len(numaralar['konsimento_list'])} "
            f"| Konteyner: {len(numaralar['konteyner_list'])} "
            f"| Beyanname: {len(numaralar['beyanname_list'])}"
        )

        # Eşleştir
        eslesme = eslestir(numaralar, referans)

        if eslesme:
            sheets_eslesmeyiKaydet(
                gonderen, konu, tarih,
                eslesme["dosya_no"], eslesme["kriter"], eslesme["deger"]
            )
            html = bildirim_html(
                gonderen, konu, tarih,
                eslesme["dosya_no"], eslesme["kriter"], eslesme["deger"]
            )
            gmail_mail_gonder(
                gmail,
                f"✅ Fatura Eşleşmesi — Dosya No: {eslesme['dosya_no']}",
                html
            )
            db_islendi_ekle(conn, email_id, eslesme["dosya_no"])
            eslesti += 1
        else:
            db_bekleyen_ekle(conn, email_id, gonderen, konu, tarih, numaralar)
            beklemeye += 1

        gmail_okundu_isaretle(gmail, email_id)

    # Bekleyenleri yeniden dene
    bekleyenler = db_bekleyenleri_getir(conn)
    yeniden_eslesti = 0

    for item in bekleyenler:
        numaralar = json.loads(item["numaralar"])
        eslesme   = eslestir(numaralar, referans)

        if eslesme:
            sheets_eslesmeyiKaydet(
                item["gonderen"], item["konu"], item["tarih"],
                eslesme["dosya_no"], eslesme["kriter"], eslesme["deger"]
            )
            html = bildirim_html(
                item["gonderen"], item["konu"], item["tarih"],
                eslesme["dosya_no"], eslesme["kriter"], eslesme["deger"]
            )
            gmail_mail_gonder(
                gmail,
                f"✅ Fatura Eşleşmesi — Dosya No: {eslesme['dosya_no']}",
                html
            )
            db_islendi_ekle(conn, item["email_id"], eslesme["dosya_no"])
            db_bekleyeni_sil(conn, item["id"])
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
