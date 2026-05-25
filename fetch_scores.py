#!/usr/bin/env python3
"""
fetch_scores.py – Holt WM 2026 Ergebnisse von football-data.org,
aktualisiert Ergebnisse.xlsx und ruft auswertung.py auf.

Lokale Nutzung:
    python fetch_scores.py             # nur lokal
    python fetch_scores.py --publish   # + Google Drive Update

Auf GitHub Actions läuft es automatisch mit --publish via Cron.

API Key: Umgebungsvariable FOOTBALL_DATA_API_KEY setzen,
         oder in dieser Datei direkt eintragen (Zeile API_KEY).
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import os
import subprocess
import requests
import openpyxl
from pathlib import Path

BASE         = Path(__file__).parent
ERGEBNISSE   = BASE / "Ergebnisse.xlsx"
TIPPS_ORDNER = BASE / "Tipps"

# ──────────────────────────────────────────────────────────────
# API KONFIGURATION
# ──────────────────────────────────────────────────────────────
API_KEY  = os.environ.get("FOOTBALL_DATA_API_KEY", "f64e2692b1d34d1fb6287cbc2b6659fe")
BASE_URL = "https://api.football-data.org/v4"

# ──────────────────────────────────────────────────────────────
# TEAM-NAMEN MAPPING  (football-data.org Englisch → Deutsch)
# ──────────────────────────────────────────────────────────────
TEAM_MAP = {
    "Mexico":                       "Mexiko",
    "South Africa":                 "Südafrika",
    "Korea Republic":               "Südkorea",
    "South Korea":                  "Südkorea",
    "Republic of Korea":            "Südkorea",
    "Czechia":                      "Tschechien",
    "Czech Republic":               "Tschechien",
    "Canada":                       "Kanada",
    "Bosnia and Herzegovina":       "Bosnien-Herzegowina",
    "Bosnia-Herzegovina":           "Bosnien-Herzegowina",
    "Qatar":                        "Katar",
    "Switzerland":                  "Schweiz",
    "Brazil":                       "Brasilien",
    "Morocco":                      "Marokko",
    "Haiti":                        "Haiti",
    "Scotland":                     "Schottland",
    "United States":                "USA",
    "USA":                          "USA",
    "Paraguay":                     "Paraguay",
    "Australia":                    "Australien",
    "Turkey":                       "Türkei",
    "Türkiye":                      "Türkei",
    "Germany":                      "Deutschland",
    "Curaçao":                      "Curaçao",
    "Curacao":                      "Curaçao",
    "Côte d'Ivoire":                "Elfenbeinküste",
    "Ivory Coast":                  "Elfenbeinküste",
    "Ecuador":                      "Ecuador",
    "Netherlands":                  "Niederlande",
    "Japan":                        "Japan",
    "Sweden":                       "Schweden",
    "Tunisia":                      "Tunesien",
    "Belgium":                      "Belgien",
    "Egypt":                        "Ägypten",
    "Iran":                         "Iran",
    "New Zealand":                  "Neuseeland",
    "Spain":                        "Spanien",
    "Cape Verde":                   "Kap Verde",
    "Cabo Verde":                   "Kap Verde",
    "Saudi Arabia":                 "Saudi-Arabien",
    "Uruguay":                      "Uruguay",
    "France":                       "Frankreich",
    "Senegal":                      "Senegal",
    "Iraq":                         "Irak",
    "Norway":                       "Norwegen",
    "Argentina":                    "Argentinien",
    "Algeria":                      "Algerien",
    "Austria":                      "Österreich",
    "Jordan":                       "Jordanien",
    "Portugal":                     "Portugal",
    "DR Congo":                     "DR Kongo",
    "Congo DR":                     "DR Kongo",
    "Democratic Republic of Congo": "DR Kongo",
    "Uzbekistan":                   "Usbekistan",
    "Colombia":                     "Kolumbien",
    "England":                      "England",
    "Croatia":                      "Kroatien",
    "Ghana":                        "Ghana",
    "Panama":                       "Panama",
}

def to_de(name):
    """Englischen API-Namen → Deutschen Namen."""
    return TEAM_MAP.get(name, name)

# ──────────────────────────────────────────────────────────────
# SPIELPLAN  (muss mit auswertung.py übereinstimmen)
# ──────────────────────────────────────────────────────────────
sys.path.insert(0, str(BASE))
from auswertung import GRUPPENSPIELE

# ──────────────────────────────────────────────────────────────
# API ABRUF
# ──────────────────────────────────────────────────────────────
def fetch_matches():
    url = f"{BASE_URL}/competitions/WC/matches"
    headers = {"X-Auth-Token": API_KEY}
    r = requests.get(url, headers=headers, timeout=15)
    if r.status_code == 403:
        print("FEHLER: API Key ungültig oder kein Zugriff auf WC-Daten.")
        sys.exit(1)
    if r.status_code == 404:
        print("HINWEIS: WM 2026 noch nicht in der API verfügbar.")
        return []
    r.raise_for_status()
    return r.json().get("matches", [])

# ──────────────────────────────────────────────────────────────
# ERGEBNISSE AUSWERTEN
# WM 2026 Rundenstruktur (48 Teams):
#   ROUND_OF_32   → unser S16  (32 spielen, 16 kommen weiter)
#   ROUND_OF_16   → unser S8
#   QUARTER_FINALS→ unser VF
#   SEMI_FINALS   → unser HF
#   FINAL         → unser F + WM
# ──────────────────────────────────────────────────────────────
STAGE_MAP = {
    "ROUND_OF_32":    "S16",
    "ROUND_OF_16":    "S8",
    "QUARTER_FINALS": "VF",
    "SEMI_FINALS":    "HF",
    "FINAL":          "F",
}

def parse_matches(api_matches):
    # Lookup: (heim_de, gast_de) → match_id
    gs_lookup = {(m["heim"], m["gast"]): m["id"] for m in GRUPPENSPIELE}

    gs_results = {}
    ko_results = {k: [] for k in ["S16", "S8", "VF", "HF", "F", "WM"]}

    unmatched = []

    for m in api_matches:
        stage  = m.get("stage", "")
        status = m.get("status", "")
        heim_de = to_de(m["homeTeam"]["name"])
        gast_de = to_de(m["awayTeam"]["name"])

        # ── Gruppenphase ──
        if stage == "GROUP_STAGE":
            if status == "FINISHED":
                match_id = gs_lookup.get((heim_de, gast_de))
                if match_id:
                    h = m["score"]["fullTime"]["home"]
                    a = m["score"]["fullTime"]["away"]
                    gs_results[match_id] = f"{h}:{a}"
                else:
                    unmatched.append(f"{heim_de} vs {gast_de}")

        # ── KO-Runden ──
        elif stage in STAGE_MAP and status == "FINISHED":
            runde = STAGE_MAP[stage]
            winner_side = m["score"].get("winner")
            winner = heim_de if winner_side == "HOME_TEAM" else (
                     gast_de if winner_side == "AWAY_TEAM" else None)

            if stage == "FINAL":
                # Beide Finalisten speichern (für Finale-Bonus)
                ko_results["F"] = [heim_de, gast_de]
                if winner:
                    ko_results["WM"] = [winner]
            elif winner:
                ko_results[runde].append(winner)

    if unmatched:
        print(f"  Warnung: {len(unmatched)} Spiele nicht zugeordnet: {unmatched[:3]}")

    return gs_results, ko_results

# ──────────────────────────────────────────────────────────────
# ERGEBNISSE.XLSX SCHREIBEN
# ──────────────────────────────────────────────────────────────
def write_ergebnisse(gs_results, ko_results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gruppenphase"
    ws.append(["ID", "Heim", "Gast", "Datum", "Ergebnis"])
    for m in GRUPPENSPIELE:
        ws.append([m["id"], m["heim"], m["gast"], m["datum"],
                   gs_results.get(m["id"], "")])

    for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
        ws2 = wb.create_sheet(runde)
        ws2.append(["Slot", "Team"])
        for i, team in enumerate(ko_results.get(runde, []), 1):
            ws2.append([i, team])

    wb.save(ERGEBNISSE)

# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  WM 2026 – Automatisches Score-Update")
    print("=" * 50)

    # Tipps-Ordner prüfen
    if not TIPPS_ORDNER.exists() or not list(TIPPS_ORDNER.glob("*.json")):
        print(f"\nHINWEIS: Keine JSON-Tipps in '{TIPPS_ORDNER}' gefunden.")
        print("Workflow wartet auf Tipp-Dateien – kein Fehler.")
        sys.exit(0)

    # Ergebnisse abrufen
    print("\nHole Ergebnisse von football-data.org ...")
    api_matches = fetch_matches()

    if not api_matches:
        print("Keine Spieldaten verfügbar – Abbruch.")
        sys.exit(0)

    print(f"  {len(api_matches)} Spiele in der API gefunden")
    gs_results, ko_results = parse_matches(api_matches)

    finished_gs = sum(1 for v in gs_results.values() if v)
    print(f"  Gruppenphase abgeschlossen: {finished_gs}/72")
    ko_info = {k: len(v) for k, v in ko_results.items() if v}
    if ko_info:
        print(f"  KO-Teams bekannt: {ko_info}")

    # Ergebnisse.xlsx aktualisieren
    write_ergebnisse(gs_results, ko_results)
    print(f"\n✓ Ergebnisse.xlsx aktualisiert")

    # Rangliste erstellen
    print("\nErstelle Rangliste ...")
    result = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(BASE / "auswertung.py"),
         str(TIPPS_ORDNER),
         str(ERGEBNISSE)],
    )
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
