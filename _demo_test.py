#!/usr/bin/env python3
"""
Demo: 3 fiktive Teilnehmer nach Spieltag 2 der Vorrunde.
Erstellt Testdaten und führt auswertung.py aus.
"""
import json, sys, subprocess
from pathlib import Path
import openpyxl

BASE = Path(__file__).parent

# ──────────────────────────────────────────────────────────────
# TATSÄCHLICHE ERGEBNISSE (48 Spiele = Spieltag 1+2 pro Gruppe)
# ──────────────────────────────────────────────────────────────
ACTUAL = {
    # Spieltag 1
    "A1":"2:1", "A2":"1:1",
    "B1":"1:0", "B2":"2:2",
    "C1":"3:1", "C2":"0:2",
    "D1":"2:0", "D2":"1:2",
    "E1":"5:0", "E2":"2:1",
    "F1":"1:1", "F2":"2:1",
    "G1":"1:0", "G2":"2:1",
    "H1":"2:0", "H2":"1:1",
    "I1":"3:0", "I2":"1:0",
    "J1":"2:1", "J2":"1:0",
    "K1":"1:0", "K2":"0:1",
    "L1":"2:0", "L2":"3:1",
    # Spieltag 2
    "A3":"1:0", "A4":"3:0",
    "B3":"1:1", "B4":"2:1",
    "C3":"1:1", "C4":"4:0",
    "D3":"1:0", "D4":"0:0",
    "E3":"2:0", "E4":"1:0",
    "F3":"3:0", "F4":"0:1",
    "G3":"0:1", "G4":"1:2",
    "H3":"2:1", "H4":"2:0",
    "I3":"2:1", "I4":"1:0",
    "J3":"3:1", "J4":"2:0",
    "K3":"2:2", "K4":"1:0",
    "L3":"1:1", "L4":"0:2",
}

# ──────────────────────────────────────────────────────────────
# TIPPS PRO PERSON (48 gespielte Spiele)
# Format: (anna, thomas, lisa)
# ──────────────────────────────────────────────────────────────
# Legende: pts in () sind zum Testen – NICHT im Code verwendet
TIPS_PLAYED = {
    # GROUP A — actual: 2:1 / 1:1 / 1:0 / 3:0
    "A1": ("2:1",  "1:0",  "2:0"),   # A:4 T:2 L:2
    "A2": ("2:1",  "1:1",  "0:0"),   # A:0 T:4 L:2
    "A3": ("1:0",  "2:1",  "0:1"),   # A:4 T:2 L:0
    "A4": ("3:1",  "3:0",  "2:0"),   # A:2 T:4 L:2
    # GROUP B — actual: 1:0 / 2:2 / 1:1 / 2:1
    "B1": ("2:0",  "1:0",  "0:1"),   # A:2 T:4 L:0
    "B2": ("1:1",  "2:2",  "3:1"),   # A:2 T:4 L:0
    "B3": ("1:1",  "2:1",  "1:1"),   # A:4 T:0 L:4
    "B4": ("2:1",  "2:0",  "1:0"),   # A:4 T:2 L:2
    # GROUP C — actual: 3:1 / 0:2 / 1:1 / 4:0
    "C1": ("3:1",  "2:0",  "2:1"),   # A:4 T:2 L:2
    "C2": ("0:1",  "0:2",  "1:1"),   # A:2 T:4 L:0
    "C3": ("2:1",  "1:1",  "0:0"),   # A:0 T:4 L:2
    "C4": ("4:0",  "3:0",  "2:0"),   # A:4 T:2 L:2
    # GROUP D — actual: 2:0 / 1:2 / 1:0 / 0:0
    "D1": ("2:0",  "1:0",  "2:1"),   # A:4 T:2 L:2
    "D2": ("2:1",  "0:1",  "1:2"),   # A:0 T:2 L:4
    "D3": ("1:0",  "1:0",  "2:1"),   # A:4 T:4 L:2
    "D4": ("1:1",  "0:0",  "1:0"),   # A:2 T:4 L:0
    # GROUP E — actual: 5:0 / 2:1 / 2:0 / 1:0
    "E1": ("4:0",  "3:0",  "5:0"),   # A:2 T:2 L:4
    "E2": ("2:1",  "1:0",  "1:2"),   # A:4 T:2 L:0
    "E3": ("2:0",  "2:1",  "1:0"),   # A:4 T:2 L:2
    "E4": ("1:0",  "2:0",  "0:1"),   # A:4 T:2 L:0
    # GROUP F — actual: 1:1 / 2:1 / 3:0 / 0:1
    "F1": ("1:1",  "2:1",  "1:1"),   # A:4 T:0 L:4
    "F2": ("3:1",  "2:1",  "2:2"),   # A:2 T:4 L:0
    "F3": ("2:0",  "3:0",  "2:1"),   # A:2 T:4 L:2
    "F4": ("1:1",  "0:1",  "0:2"),   # A:0 T:4 L:2
    # GROUP G — actual: 1:0 / 2:1 / 0:1 / 1:2
    "G1": ("1:0",  "2:0",  "1:1"),   # A:4 T:2 L:0
    "G2": ("2:1",  "2:1",  "1:0"),   # A:4 T:4 L:2
    "G3": ("1:0",  "0:1",  "0:2"),   # A:0 T:4 L:2
    "G4": ("1:1",  "1:2",  "0:1"),   # A:0 T:4 L:2
    # GROUP H — actual: 2:0 / 1:1 / 2:1 / 2:0
    "H1": ("2:0",  "1:0",  "2:1"),   # A:4 T:2 L:2
    "H2": ("1:1",  "1:1",  "2:1"),   # A:4 T:4 L:0
    "H3": ("2:1",  "2:0",  "1:0"),   # A:4 T:2 L:2
    "H4": ("3:0",  "2:0",  "1:0"),   # A:2 T:4 L:2
    # GROUP I — actual: 3:0 / 1:0 / 2:1 / 1:0
    "I1": ("3:0",  "2:0",  "3:1"),   # A:4 T:2 L:2
    "I2": ("0:1",  "1:0",  "0:0"),   # A:0 T:4 L:0
    "I3": ("2:1",  "3:1",  "2:0"),   # A:4 T:2 L:2
    "I4": ("1:0",  "1:1",  "2:0"),   # A:4 T:0 L:2
    # GROUP J — actual: 2:1 / 1:0 / 3:1 / 2:0
    "J1": ("2:1",  "2:1",  "1:0"),   # A:4 T:4 L:2
    "J2": ("1:0",  "1:0",  "2:1"),   # A:4 T:4 L:2
    "J3": ("3:0",  "3:1",  "2:1"),   # A:2 T:4 L:2
    "J4": ("1:0",  "2:0",  "1:1"),   # A:2 T:2 L:0
    # GROUP K — actual: 1:0 / 0:1 / 2:2 / 1:0
    "K1": ("1:0",  "2:1",  "0:1"),   # A:4 T:2 L:0
    "K2": ("0:1",  "1:1",  "1:2"),   # A:4 T:0 L:2
    "K3": ("1:1",  "2:2",  "2:1"),   # A:2 T:4 L:0
    "K4": ("2:0",  "1:0",  "2:1"),   # A:2 T:4 L:2
    # GROUP L — actual: 2:0 / 3:1 / 1:1 / 0:2
    "L1": ("2:0",  "1:0",  "2:1"),   # A:4 T:2 L:2
    "L2": ("3:1",  "2:0",  "3:2"),   # A:4 T:2 L:2
    "L3": ("2:0",  "1:1",  "1:1"),   # A:0 T:4 L:4
    "L4": ("0:2",  "0:1",  "1:1"),   # A:4 T:2 L:0
}

# Ungespielte Spiele (Spieltag 3) – kein Einfluss auf aktuelle Wertung
TIPS_UNPLAYED = {
    "A5": ("2:1","1:0","0:1"), "A6": ("1:0","1:1","2:1"),
    "B5": ("2:1","1:1","2:0"), "B6": ("1:0","2:1","1:2"),
    "C5": ("1:1","2:1","1:0"), "C6": ("2:0","1:0","3:1"),
    "D5": ("2:1","1:0","0:1"), "D6": ("2:0","1:1","1:0"),
    "E5": ("1:1","0:1","1:0"), "E6": ("3:0","2:0","1:1"),
    "F5": ("2:1","1:1","0:1"), "F6": ("1:2","0:1","1:0"),
    "G5": ("1:2","2:1","1:1"), "G6": ("2:1","1:0","2:0"),
    "H5": ("0:2","0:1","1:1"), "H6": ("2:0","3:1","1:0"),
    "I5": ("0:1","1:2","1:1"), "I6": ("2:1","1:0","0:1"),
    "J5": ("1:2","0:2","1:1"), "J6": ("1:0","0:1","2:1"),
    "K5": ("3:0","2:1","1:0"), "K6": ("0:1","1:0","1:2"),
    "L5": ("1:2","0:1","1:1"), "L6": ("1:0","2:1","0:1"),
}

# ──────────────────────────────────────────────────────────────
# KO-TIPPS (vor dem Turnier abgegeben, noch nicht gewertet)
# ──────────────────────────────────────────────────────────────
KO_TIPPS = {
    "anna": {
        "S16": ["Deutschland","Frankreich","Brasilien","Argentinien",
                "Spanien","Portugal","England","Niederlande",
                "USA","Mexiko","Uruguay","Belgien",
                "Japan","Türkei","Österreich","Kolumbien"],
        "S8":  ["Deutschland","Frankreich","Brasilien","Argentinien",
                "Spanien","Portugal","England","Niederlande"],
        "VF":  ["Deutschland","Frankreich","Brasilien","Spanien"],
        "HF":  ["Deutschland","Brasilien"],
        "F":   ["Deutschland","Brasilien"],
        "WM":  "Deutschland",
    },
    "thomas": {
        "S16": ["Frankreich","Argentinien","Brasilien","Deutschland",
                "Spanien","Portugal","England","Belgien",
                "Niederlande","USA","Uruguay","Schweiz",
                "Japan","Kroatien","Türkei","Österreich"],
        "S8":  ["Frankreich","Argentinien","Brasilien","Deutschland",
                "Spanien","Portugal","England","Belgien"],
        "VF":  ["Frankreich","Argentinien","Brasilien","Spanien"],
        "HF":  ["Frankreich","Argentinien"],
        "F":   ["Frankreich","Argentinien"],
        "WM":  "Frankreich",
    },
    "lisa": {
        "S16": ["Brasilien","Deutschland","Argentinien","Frankreich",
                "Spanien","Niederlande","Portugal","USA",
                "England","Mexiko","Schweiz","Uruguay",
                "Japan","Ghana","Australien","Norwegen"],
        "S8":  ["Brasilien","Deutschland","Argentinien","Frankreich",
                "Spanien","Niederlande","Portugal","USA"],
        "VF":  ["Brasilien","Argentinien","Frankreich","Spanien"],
        "HF":  ["Brasilien","Frankreich"],
        "F":   ["Brasilien","Frankreich"],
        "WM":  "Brasilien",
    },
}

# ──────────────────────────────────────────────────────────────
# GRUPPENSPIELE (kopiert aus auswertung.py für Spaltennamen)
# ──────────────────────────────────────────────────────────────
GRUPPENSPIELE = [
    {"id":"A1","datum":"11.06.","heim":"Mexiko","gast":"Südafrika","gruppe":"A"},
    {"id":"A2","datum":"11.06.","heim":"Südkorea","gast":"Tschechien","gruppe":"A"},
    {"id":"A3","datum":"18.06.","heim":"Tschechien","gast":"Südafrika","gruppe":"A"},
    {"id":"A4","datum":"18.06.","heim":"Mexiko","gast":"Südkorea","gruppe":"A"},
    {"id":"A5","datum":"24.06.","heim":"Tschechien","gast":"Mexiko","gruppe":"A"},
    {"id":"A6","datum":"24.06.","heim":"Südafrika","gast":"Südkorea","gruppe":"A"},
    {"id":"B1","datum":"12.06.","heim":"Kanada","gast":"Bosnien-Herzegowina","gruppe":"B"},
    {"id":"B2","datum":"13.06.","heim":"Katar","gast":"Schweiz","gruppe":"B"},
    {"id":"B3","datum":"18.06.","heim":"Schweiz","gast":"Bosnien-Herzegowina","gruppe":"B"},
    {"id":"B4","datum":"18.06.","heim":"Kanada","gast":"Katar","gruppe":"B"},
    {"id":"B5","datum":"24.06.","heim":"Schweiz","gast":"Kanada","gruppe":"B"},
    {"id":"B6","datum":"24.06.","heim":"Bosnien-Herzegowina","gast":"Katar","gruppe":"B"},
    {"id":"C1","datum":"13.06.","heim":"Brasilien","gast":"Marokko","gruppe":"C"},
    {"id":"C2","datum":"13.06.","heim":"Haiti","gast":"Schottland","gruppe":"C"},
    {"id":"C3","datum":"19.06.","heim":"Schottland","gast":"Marokko","gruppe":"C"},
    {"id":"C4","datum":"19.06.","heim":"Brasilien","gast":"Haiti","gruppe":"C"},
    {"id":"C5","datum":"24.06.","heim":"Schottland","gast":"Brasilien","gruppe":"C"},
    {"id":"C6","datum":"24.06.","heim":"Marokko","gast":"Haiti","gruppe":"C"},
    {"id":"D1","datum":"12.06.","heim":"USA","gast":"Paraguay","gruppe":"D"},
    {"id":"D2","datum":"13.06.","heim":"Australien","gast":"Türkei","gruppe":"D"},
    {"id":"D3","datum":"19.06.","heim":"USA","gast":"Australien","gruppe":"D"},
    {"id":"D4","datum":"19.06.","heim":"Türkei","gast":"Paraguay","gruppe":"D"},
    {"id":"D5","datum":"25.06.","heim":"Türkei","gast":"USA","gruppe":"D"},
    {"id":"D6","datum":"25.06.","heim":"Paraguay","gast":"Australien","gruppe":"D"},
    {"id":"E1","datum":"14.06.","heim":"Deutschland","gast":"Curaçao","gruppe":"E"},
    {"id":"E2","datum":"14.06.","heim":"Elfenbeinküste","gast":"Ecuador","gruppe":"E"},
    {"id":"E3","datum":"20.06.","heim":"Deutschland","gast":"Elfenbeinküste","gruppe":"E"},
    {"id":"E4","datum":"20.06.","heim":"Ecuador","gast":"Curaçao","gruppe":"E"},
    {"id":"E5","datum":"25.06.","heim":"Curaçao","gast":"Elfenbeinküste","gruppe":"E"},
    {"id":"E6","datum":"25.06.","heim":"Ecuador","gast":"Deutschland","gruppe":"E"},
    {"id":"F1","datum":"14.06.","heim":"Niederlande","gast":"Japan","gruppe":"F"},
    {"id":"F2","datum":"14.06.","heim":"Schweden","gast":"Tunesien","gruppe":"F"},
    {"id":"F3","datum":"20.06.","heim":"Niederlande","gast":"Schweden","gruppe":"F"},
    {"id":"F4","datum":"20.06.","heim":"Tunesien","gast":"Japan","gruppe":"F"},
    {"id":"F5","datum":"25.06.","heim":"Japan","gast":"Schweden","gruppe":"F"},
    {"id":"F6","datum":"25.06.","heim":"Tunesien","gast":"Niederlande","gruppe":"F"},
    {"id":"G1","datum":"15.06.","heim":"Belgien","gast":"Ägypten","gruppe":"G"},
    {"id":"G2","datum":"15.06.","heim":"Iran","gast":"Neuseeland","gruppe":"G"},
    {"id":"G3","datum":"21.06.","heim":"Belgien","gast":"Iran","gruppe":"G"},
    {"id":"G4","datum":"21.06.","heim":"Neuseeland","gast":"Ägypten","gruppe":"G"},
    {"id":"G5","datum":"26.06.","heim":"Ägypten","gast":"Iran","gruppe":"G"},
    {"id":"G6","datum":"26.06.","heim":"Neuseeland","gast":"Belgien","gruppe":"G"},
    {"id":"H1","datum":"15.06.","heim":"Spanien","gast":"Kap Verde","gruppe":"H"},
    {"id":"H2","datum":"15.06.","heim":"Saudi-Arabien","gast":"Uruguay","gruppe":"H"},
    {"id":"H3","datum":"21.06.","heim":"Spanien","gast":"Saudi-Arabien","gruppe":"H"},
    {"id":"H4","datum":"21.06.","heim":"Uruguay","gast":"Kap Verde","gruppe":"H"},
    {"id":"H5","datum":"26.06.","heim":"Kap Verde","gast":"Saudi-Arabien","gruppe":"H"},
    {"id":"H6","datum":"26.06.","heim":"Uruguay","gast":"Spanien","gruppe":"H"},
    {"id":"I1","datum":"16.06.","heim":"Frankreich","gast":"Senegal","gruppe":"I"},
    {"id":"I2","datum":"16.06.","heim":"Irak","gast":"Norwegen","gruppe":"I"},
    {"id":"I3","datum":"22.06.","heim":"Frankreich","gast":"Irak","gruppe":"I"},
    {"id":"I4","datum":"22.06.","heim":"Norwegen","gast":"Senegal","gruppe":"I"},
    {"id":"I5","datum":"26.06.","heim":"Norwegen","gast":"Frankreich","gruppe":"I"},
    {"id":"I6","datum":"26.06.","heim":"Senegal","gast":"Irak","gruppe":"I"},
    {"id":"J1","datum":"16.06.","heim":"Argentinien","gast":"Algerien","gruppe":"J"},
    {"id":"J2","datum":"16.06.","heim":"Österreich","gast":"Jordanien","gruppe":"J"},
    {"id":"J3","datum":"22.06.","heim":"Argentinien","gast":"Österreich","gruppe":"J"},
    {"id":"J4","datum":"22.06.","heim":"Jordanien","gast":"Algerien","gruppe":"J"},
    {"id":"J5","datum":"27.06.","heim":"Algerien","gast":"Österreich","gruppe":"J"},
    {"id":"J6","datum":"27.06.","heim":"Jordanien","gast":"Argentinien","gruppe":"J"},
    {"id":"K1","datum":"17.06.","heim":"Portugal","gast":"DR Kongo","gruppe":"K"},
    {"id":"K2","datum":"17.06.","heim":"Usbekistan","gast":"Kolumbien","gruppe":"K"},
    {"id":"K3","datum":"23.06.","heim":"Portugal","gast":"Usbekistan","gruppe":"K"},
    {"id":"K4","datum":"23.06.","heim":"Kolumbien","gast":"DR Kongo","gruppe":"K"},
    {"id":"K5","datum":"27.06.","heim":"Kolumbien","gast":"Portugal","gruppe":"K"},
    {"id":"K6","datum":"27.06.","heim":"DR Kongo","gast":"Usbekistan","gruppe":"K"},
    {"id":"L1","datum":"17.06.","heim":"England","gast":"Kroatien","gruppe":"L"},
    {"id":"L2","datum":"17.06.","heim":"Ghana","gast":"Panama","gruppe":"L"},
    {"id":"L3","datum":"23.06.","heim":"England","gast":"Ghana","gruppe":"L"},
    {"id":"L4","datum":"23.06.","heim":"Panama","gast":"Kroatien","gruppe":"L"},
    {"id":"L5","datum":"27.06.","heim":"Panama","gast":"England","gruppe":"L"},
    {"id":"L6","datum":"27.06.","heim":"Kroatien","gast":"Ghana","gruppe":"L"},
]

def match_col(m):
    return f"[Gr.{m['gruppe']}] {m['heim']} vs {m['gast']} ({m['datum']})"

# ──────────────────────────────────────────────────────────────
# DATEIEN ERSTELLEN
# ──────────────────────────────────────────────────────────────

def build_gruppenphase_tips():
    tips = {}
    for m in GRUPPENSPIELE:
        col = match_col(m)
        played = TIPS_PLAYED.get(m["id"])
        unplayed = TIPS_UNPLAYED.get(m["id"])
        anna   = played[0] if played else (unplayed[0] if unplayed else "2:1")
        thomas = played[1] if played else (unplayed[1] if unplayed else "1:0")
        lisa   = played[2] if played else (unplayed[2] if unplayed else "1:1")
        tips[col] = (anna, thomas, lisa)
    return tips

def create_json_files():
    folder = BASE / "Tipps_Demo"
    folder.mkdir(exist_ok=True)

    gt = build_gruppenphase_tips()
    people = [
        ("anna",   "Anna Müller"),
        ("thomas", "Thomas Schmidt"),
        ("lisa",   "Lisa Weber"),
    ]

    for key, name in people:
        idx = {"anna": 0, "thomas": 1, "lisa": 2}[key]
        gruppenphase = {col: tips[idx] for col, tips in gt.items()}
        ko = KO_TIPPS[key]
        data = {
            "name": name,
            "exported": "2026-06-01T08:00:00.000Z",
            "gruppenphase": gruppenphase,
            "S16": ko["S16"],
            "S8":  ko["S8"],
            "VF":  ko["VF"],
            "HF":  ko["HF"],
            "F":   ko["F"],
            "WM":  ko["WM"],
        }
        path = folder / f"{key}.json"
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Tipp-Dateien erstellt: {folder}")
    return folder

def create_ergebnisse():
    path = BASE / "Ergebnisse_Demo.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gruppenphase"
    ws.append(["ID", "Heim", "Gast", "Datum", "Ergebnis"])

    for m in GRUPPENSPIELE:
        ergebnis = ACTUAL.get(m["id"], "")
        ws.append([m["id"], m["heim"], m["gast"], m["datum"], ergebnis])

    for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
        ws2 = wb.create_sheet(runde)
        ws2.append(["Slot", "Team"])

    wb.save(path)
    print(f"Ergebnisdatei erstellt: {path.name}")
    return path

# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 55)
    print("  WM 2026 Demo – 3 Teilnehmer nach Spieltag 2")
    print(f"  Gespielte Spiele: {len(ACTUAL)}/72 (je Spieltag 1+2)")
    print("=" * 55)
    print()

    tipps_folder = create_json_files()
    ergebnisse   = create_ergebnisse()
    print()

    result = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(BASE / "auswertung.py"),
         str(tipps_folder), str(ergebnisse)],
    )
    sys.stdout.flush()
