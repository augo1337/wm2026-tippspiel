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
from datetime import datetime, timezone, timedelta
from pathlib import Path

BASE         = Path(__file__).parent
ERGEBNISSE   = BASE / "Ergebnisse.xlsx"
TIPPS_ORDNER = BASE / "Tipps"

# ──────────────────────────────────────────────────────────────
# SPIELPLAN (Anstoß-Zeiten in UTC; Alle Zeiten CEST = UTC+2)
# Gruppenphase + KO bis Finale – wird für Zeitfenster-Check genutzt
# ──────────────────────────────────────────────────────────────
def _u(mo, d, h, mi=0):
    return datetime(2026, mo, d, h, mi, tzinfo=timezone.utc)

MATCH_KICKOFFS_UTC = [
    # Gruppenphase (11.06.–28.06.)
    _u(6,11,19),                                              # 11.06
    _u(6,12, 2), _u(6,12,19),                                # 12.06
    _u(6,13, 1), _u(6,13,19), _u(6,13,22),                  # 13.06
    _u(6,14, 1), _u(6,14, 4), _u(6,14,17), _u(6,14,20), _u(6,14,23),  # 14.06
    _u(6,15, 2), _u(6,15,16), _u(6,15,19), _u(6,15,22),    # 15.06
    _u(6,16, 1), _u(6,16,19), _u(6,16,22),                  # 16.06
    _u(6,17, 1), _u(6,17, 4), _u(6,17,17), _u(6,17,20), _u(6,17,23),  # 17.06
    _u(6,18, 2), _u(6,18,16), _u(6,18,19), _u(6,18,22),    # 18.06
    _u(6,19, 1), _u(6,19,19), _u(6,19,22),                  # 19.06
    _u(6,20, 0,30), _u(6,20, 3), _u(6,20,17), _u(6,20,20), # 20.06
    _u(6,21, 0), _u(6,21, 4), _u(6,21,16), _u(6,21,19), _u(6,21,22),  # 21.06
    _u(6,22, 1), _u(6,22,17), _u(6,22,21),                  # 22.06
    _u(6,23, 0), _u(6,23, 3), _u(6,23,17), _u(6,23,20), _u(6,23,23),  # 23.06
    _u(6,24, 2), _u(6,24,19), _u(6,24,22),                  # 24.06
    _u(6,25, 1), _u(6,25,20), _u(6,25,23),                  # 25.06
    _u(6,26, 2), _u(6,26,19),                                # 26.06
    _u(6,27, 0), _u(6,27, 3), _u(6,27,21), _u(6,27,23,30), # 27.06
    _u(6,28, 2),                                             # 28.06 letzte Gruppenspiele
    # KO-Runden (28.06.–19.07.)
    _u(6,28,19),                                             # S16 Tag 1
    _u(6,29,17), _u(6,29,20,30),                            # S16
    _u(6,30, 1), _u(6,30,17), _u(6,30,21),                  # S16
    _u(7, 1, 1), _u(7, 1,16), _u(7, 1,20),                  # S16
    _u(7, 2, 0), _u(7, 2,19), _u(7, 2,23),                  # S16
    _u(7, 3, 3), _u(7, 3,18), _u(7, 3,22),                  # S16
    _u(7, 4, 1,30), _u(7, 4,17), _u(7, 4,21),               # S16 Ende / S8 Beginn
    _u(7, 5,20),                                             # S8
    _u(7, 6, 0), _u(7, 6,19),                                # S8
    _u(7, 7, 0), _u(7, 7,16), _u(7, 7,20),                  # S8 Ende
    _u(7, 9,20),                                             # VF
    _u(7,10,19),                                             # VF
    _u(7,11,21),                                             # VF
    _u(7,12, 1),                                             # VF Ende
    _u(7,14,19), _u(7,15,19),                                # HF
    _u(7,18,19),                                             # Platz 3
    _u(7,19,19),                                             # Finale
]

# Zeitfenster: wie viele Minuten nach Anstoß darf ein Run noch als "relevant" gelten?
# 210 min = 90 min Spielzeit + 30 min Verlängerung + 30 min Elfmeter + 30 min Puffer
MATCH_WINDOW_MIN = 210

def is_match_window():
    """True wenn gerade ein WM-Spiel läuft oder erst kürzlich geendet hat."""
    now = datetime.now(timezone.utc)
    window = timedelta(minutes=MATCH_WINDOW_MIN)
    return any(kickoff <= now <= kickoff + window for kickoff in MATCH_KICKOFFS_UTC)

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
                    if h is not None and a is not None:
                        gs_results[match_id] = f"{h}:{a}"
                    else:
                        print(f"  Warnung: Kein Ergebnis für {heim_de} vs {gast_de} (fullTime null)")
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
    # Vorhandene valide Ergebnisse laden und mit neuen mergen
    existing_gs = {}
    existing_ko = {k: [] for k in ["S16", "S8", "VF", "HF", "F", "WM"]}
    if ERGEBNISSE.exists():
        try:
            wb_old = openpyxl.load_workbook(ERGEBNISSE)
            if "Gruppenphase" in wb_old.sheetnames:
                for row in wb_old["Gruppenphase"].iter_rows(min_row=2, values_only=True):
                    mid, _, _, _, score = row[0], row[1], row[2], row[3], row[4]
                    if mid and score:
                        import re as _re
                        if _re.match(r"^\d+:\d+$", str(score).strip()):
                            existing_gs[str(mid)] = str(score).strip()
            for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
                if runde in wb_old.sheetnames:
                    teams = []
                    for row in wb_old[runde].iter_rows(min_row=2, values_only=True):
                        val = row[1] if len(row) > 1 else row[0]
                        if val:
                            teams.append(str(val).strip())
                    existing_ko[runde] = teams
        except Exception as e:
            print(f"  Hinweis: Konnte alte Ergebnisse nicht lesen: {e}")

    # Neue Ergebnisse über alte mergen (neue überschreiben alte)
    merged_gs = {**existing_gs, **gs_results}
    merged_ko = {}
    for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
        merged_ko[runde] = ko_results.get(runde) if ko_results.get(runde) else existing_ko.get(runde, [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gruppenphase"
    ws.append(["ID", "Heim", "Gast", "Datum", "Ergebnis"])
    for m in GRUPPENSPIELE:
        ws.append([m["id"], m["heim"], m["gast"], m["datum"],
                   merged_gs.get(m["id"], "")])

    for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
        ws2 = wb.create_sheet(runde)
        ws2.append(["Slot", "Team"])
        for i, team in enumerate(merged_ko.get(runde, []), 1):
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

    # Fehlende Ergebnisse ermitteln (aus vorhandener xlsx)
    existing_gs = {}
    if ERGEBNISSE.exists():
        try:
            import re as _re2
            _wb = openpyxl.load_workbook(ERGEBNISSE)
            if "Gruppenphase" in _wb.sheetnames:
                for _row in _wb["Gruppenphase"].iter_rows(min_row=2, values_only=True):
                    _mid, _score = _row[0], _row[4]
                    if _mid and _score and _re2.match(r"^\d+:\d+$", str(_score).strip()):
                        existing_gs[str(_mid)] = str(_score).strip()
        except Exception:
            pass
    missing_results = len(GRUPPENSPIELE) - len(existing_gs)

    # API abrufen wenn: Spiel im aktiven Zeitfenster ODER noch fehlende Ergebnisse
    if not is_match_window() and missing_results == 0:
        print("\nKein WM-Spiel im aktiven Zeitfenster und alle Ergebnisse vorhanden – überspringe API-Abruf.")
    else:
        if is_match_window():
            print("\nSpiel im aktiven Zeitfenster – hole Ergebnisse ...")
        else:
            print(f"\n{missing_results} fehlende Ergebnisse – hole Ergebnisse trotz inaktivem Zeitfenster ...")

        api_matches = fetch_matches()

        if not api_matches:
            print("Keine Spieldaten verfügbar – überspringe API-Update.")
        else:
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

    # Rangliste immer neu erstellen (auch vor WM-Start, damit neue Tipps sichtbar werden)
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
