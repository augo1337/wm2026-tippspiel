#!/usr/bin/env python3
"""
Einmalige Einrichtung von Google Drive für das WM 2026 Tippspiel.
Führe dieses Script einmal aus, bevor du --publish in auswertung.py verwendest.
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
import json

def main():
    script_dir  = Path(__file__).parent
    creds_file  = script_dir / "credentials.json"
    token_file  = script_dir / "token.json"
    config_file = script_dir / "google_drive_config.txt"

    print("=" * 55)
    print("  WM 2026 Tippspiel – Google Drive Einrichtung")
    print("=" * 55)

    # 1. Credentials prüfen
    if not creds_file.exists():
        print("""
SCHRITT 1: credentials.json herunterladen
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Öffne: https://console.cloud.google.com/
2. Erstelle ein neues Projekt (z.B. "WM2026")
3. Klicke auf "APIs & Dienste" → "Bibliothek"
4. Suche nach "Google Drive API" und aktiviere sie
5. Klicke auf "APIs & Dienste" → "Anmeldedaten"
6. Klicke "+ Anmeldedaten erstellen" → "OAuth-Client-ID"
7. Anwendungstyp: "Desktop-App", Name: "WM2026"
8. Lade die JSON-Datei herunter
9. Benenne sie um in: credentials.json
10. Lege sie in diesen Ordner:
    """ + str(script_dir) + """

Dann dieses Script erneut ausführen.
""")
        return

    print("\n✓ credentials.json gefunden.")

    # 2. OAuth-Login
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        print("""
FEHLER: Google-Bibliotheken nicht installiert.
Bitte ausführen:
  pip install google-api-python-client google-auth-oauthlib
""")
        return

    SCOPES = ["https://www.googleapis.com/auth/drive.file"]

    creds = None
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("\nSCHRITT 2: Google-Login im Browser ...")
            print("(Es öffnet sich ein Browser-Fenster – mit deinem Google-Konto anmelden)\n")
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_file), SCOPES)
            creds = flow.run_local_server(port=0)
        token_file.write_text(creds.to_json(), encoding="utf-8")
        print("✓ Login erfolgreich, Token gespeichert.")

    service = build("drive", "v3", credentials=creds)

    # 3. Test-Datei anlegen
    print("\nSCHRITT 3: Erstelle Ranglisten-Datei auf Google Drive ...")
    test_html = script_dir / "_setup_test.html"
    test_html.write_text(
        "<html><body><h1>WM 2026 Tippspiel</h1><p>Rangliste wird hier erscheinen.</p></body></html>",
        encoding="utf-8"
    )

    meta  = {"name": "WM2026_Rangliste.html", "mimeType": "text/html"}
    media = MediaFileUpload(str(test_html), mimetype="text/html")

    if config_file.exists():
        existing_id = config_file.read_text(encoding="utf-8").strip()
        print(f"  Datei-ID bereits vorhanden: {existing_id}")
        print("  Überschreibe bestehende Datei ...")
        service.files().update(fileId=existing_id, media_body=media).execute()
        file_id = existing_id
    else:
        file = service.files().create(body=meta, media_body=media, fields="id").execute()
        file_id = file.get("id")
        config_file.write_text(file_id, encoding="utf-8")

        # Öffentlich machen
        service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()

    try:
        test_html.unlink()
    except PermissionError:
        pass  # OneDrive hält die Datei kurz fest – kein Problem

    url = f"https://drive.google.com/file/d/{file_id}/view"

    print(f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✓ Einrichtung abgeschlossen!

FIXER LINK (jetzt in die WhatsApp-Gruppe senden):
  {url}

Alle Teilnehmer speichern diesen Link.
Danach einfach nur den Browser neu laden – immer aktuell!

Zukünftige Nutzung:
  python auswertung.py Tipps/ Ergebnisse.xlsx --publish
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
""")

if __name__ == "__main__":
    main()
