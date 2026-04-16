"""
Bu scripti BİR KEZ kendi bilgisayarında çalıştır.
Gmail OAuth token'ını üretir ve ekrana basar.
O çıktıyı GitHub Secret olarak kaydet.

Gereksinimler:
    pip install google-auth-oauthlib google-auth-httplib2

Kullanım:
    python token_uret.py
"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from pathlib import Path

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets",
]

def main():
    creds = None

    if Path("token.json").exists():
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as f:
            f.write(creds.to_json())

    print("\n" + "="*60)
    print("GMAIL_TOKEN_JSON değeri (GitHub Secret'a yapıştır):")
    print("="*60)
    token_data = json.loads(Path("token.json").read_text())
    print(json.dumps(token_data))
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
