# Google Drive Einrichtung – WM 2026 Tippspiel

Mit dieser Einrichtung bekommt jeder Teilnehmer **einen fixen Link**,
der immer die aktuelle Rangliste zeigt – einfach im Browser neu laden.

---

## Einmalige Einrichtung (ca. 10 Minuten)

### Schritt 1: Google Drive API aktivieren

1. Öffne [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. Oben links auf **"Projekt auswählen"** → **"Neues Projekt"**
   - Name: `WM2026` → Erstellen
3. Im Menü: **"APIs & Dienste"** → **"Bibliothek"**
4. Suche nach `Google Drive API` → Klicke darauf → **"Aktivieren"**

### Schritt 2: OAuth-Zugangsdaten erstellen

1. Im Menü: **"APIs & Dienste"** → **"Anmeldedaten"**
2. Klicke **"+ Anmeldedaten erstellen"** → **"OAuth-Client-ID"**
3. Falls gefragt: Klicke zuerst auf **"Zustimmungsbildschirm konfigurieren"**
   - User-Typ: **"Extern"** → Erstellen
   - App-Name: `WM2026 Tippspiel`, deine E-Mail eingeben → Speichern
   - Unter **"Testnutzer"**: Deine E-Mail-Adresse hinzufügen → Speichern
   - Zurück zu "Anmeldedaten"
4. **"+ Anmeldedaten erstellen"** → **"OAuth-Client-ID"**
   - Anwendungstyp: **"Desktop-App"**
   - Name: `WM2026`
   - **"Erstellen"**
5. Im Dialog: **"JSON herunterladen"**
6. Datei umbenennen in: `credentials.json`
7. Datei in den Ordner `Fußball WM 2026/` legen

### Schritt 3: Setup-Script ausführen

```bash
pip install google-api-python-client google-auth-oauthlib
python setup_googledrive.py
```

- Es öffnet sich ein Browser-Fenster → mit deinem Google-Konto anmelden
- Das Script erstellt automatisch die Datei auf Google Drive
- Am Ende siehst du den **fixen Link** → diesen in die WhatsApp-Gruppe senden!

---

## Tägliche Nutzung

```bash
python auswertung.py Tipps/ Ergebnisse.xlsx --publish
```

Das war's. Die Datei in Google Drive wird automatisch aktualisiert.
Teilnehmer laden den Link im Browser neu – fertig.

---

## Was wird erstellt?

| Datei | Beschreibung |
|-------|-------------|
| `credentials.json` | Von dir heruntergeladen (nicht löschen!) |
| `token.json` | Automatisch erstellt (Login-Token) |
| `google_drive_config.txt` | Automatisch erstellt (Datei-ID) |

---

## Troubleshooting

**"credentials.json nicht gefunden"**
→ Schritt 1–2 wiederholen, Datei in den richtigen Ordner legen.

**"Zugriff verweigert" im Browser**
→ Unter Zustimmungsbildschirm → Testnutzer → deine E-Mail hinzufügen.

**Token abgelaufen nach Monaten**
→ `token.json` löschen und `setup_googledrive.py` erneut ausführen.
