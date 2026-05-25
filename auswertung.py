#!/usr/bin/env python3
"""
WM 2026 Tippspiel – Auswertung
===============================
Liest Tipp-Dateien (JSON vom tippzettel.html oder Ordner mit JSONs)
+ Ergebnisse.xlsx und erstellt Rangliste.xlsx + Rangliste.html.

Verwendung:
    python auswertung.py <tipps_ordner_oder_datei> <ergebnisse.xlsx>

Beispiele:
    python auswertung.py Tipps/        Ergebnisse.xlsx
"""

import sys
import re
import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from pathlib import Path

# ============================================================
# SPIELPLAN
# ============================================================

ALLE_TEAMS = [
    "Mexiko", "Südafrika", "Südkorea", "Tschechien",
    "Kanada", "Bosnien-Herzegowina", "Katar", "Schweiz",
    "Brasilien", "Marokko", "Haiti", "Schottland",
    "USA", "Paraguay", "Australien", "Türkei",
    "Deutschland", "Curaçao", "Elfenbeinküste", "Ecuador",
    "Niederlande", "Japan", "Schweden", "Tunesien",
    "Belgien", "Ägypten", "Iran", "Neuseeland",
    "Spanien", "Kap Verde", "Saudi-Arabien", "Uruguay",
    "Frankreich", "Senegal", "Irak", "Norwegen",
    "Argentinien", "Algerien", "Österreich", "Jordanien",
    "Portugal", "DR Kongo", "Usbekistan", "Kolumbien",
    "England", "Kroatien", "Ghana", "Panama",
]

GRUPPENSPIELE = [
    {"id": "A1", "datum": "11.06.", "uhrzeit": "21:00", "heim": "Mexiko",         "gast": "Südafrika",           "gruppe": "A"},
    {"id": "A2", "datum": "12.06.", "uhrzeit": "04:00", "heim": "Südkorea",       "gast": "Tschechien",          "gruppe": "A"},
    {"id": "A3", "datum": "18.06.", "uhrzeit": "18:00", "heim": "Tschechien",     "gast": "Südafrika",           "gruppe": "A"},
    {"id": "A4", "datum": "19.06.", "uhrzeit": "03:00", "heim": "Mexiko",         "gast": "Südkorea",            "gruppe": "A"},
    {"id": "A5", "datum": "25.06.", "uhrzeit": "03:00", "heim": "Tschechien",     "gast": "Mexiko",              "gruppe": "A"},
    {"id": "A6", "datum": "25.06.", "uhrzeit": "03:00", "heim": "Südafrika",      "gast": "Südkorea",            "gruppe": "A"},
    {"id": "B1", "datum": "12.06.", "uhrzeit": "21:00", "heim": "Kanada",         "gast": "Bosnien-Herzegowina", "gruppe": "B"},
    {"id": "B2", "datum": "13.06.", "uhrzeit": "21:00", "heim": "Katar",          "gast": "Schweiz",             "gruppe": "B"},
    {"id": "B3", "datum": "18.06.", "uhrzeit": "21:00", "heim": "Schweiz",        "gast": "Bosnien-Herzegowina", "gruppe": "B"},
    {"id": "B4", "datum": "19.06.", "uhrzeit": "00:00", "heim": "Kanada",         "gast": "Katar",               "gruppe": "B"},
    {"id": "B5", "datum": "24.06.", "uhrzeit": "21:00", "heim": "Schweiz",        "gast": "Kanada",              "gruppe": "B"},
    {"id": "B6", "datum": "24.06.", "uhrzeit": "21:00", "heim": "Bosnien-Herzegowina", "gast": "Katar",          "gruppe": "B"},
    {"id": "C1", "datum": "14.06.", "uhrzeit": "00:00", "heim": "Brasilien",      "gast": "Marokko",             "gruppe": "C"},
    {"id": "C2", "datum": "14.06.", "uhrzeit": "03:00", "heim": "Haiti",          "gast": "Schottland",          "gruppe": "C"},
    {"id": "C3", "datum": "20.06.", "uhrzeit": "00:00", "heim": "Schottland",     "gast": "Marokko",             "gruppe": "C"},
    {"id": "C4", "datum": "20.06.", "uhrzeit": "02:30", "heim": "Brasilien",      "gast": "Haiti",               "gruppe": "C"},
    {"id": "C5", "datum": "25.06.", "uhrzeit": "00:00", "heim": "Schottland",     "gast": "Brasilien",           "gruppe": "C"},
    {"id": "C6", "datum": "25.06.", "uhrzeit": "00:00", "heim": "Marokko",        "gast": "Haiti",               "gruppe": "C"},
    {"id": "D1", "datum": "13.06.", "uhrzeit": "03:00", "heim": "USA",            "gast": "Paraguay",            "gruppe": "D"},
    {"id": "D2", "datum": "14.06.", "uhrzeit": "06:00", "heim": "Australien",     "gast": "Türkei",              "gruppe": "D"},
    {"id": "D3", "datum": "19.06.", "uhrzeit": "21:00", "heim": "USA",            "gast": "Australien",          "gruppe": "D"},
    {"id": "D4", "datum": "20.06.", "uhrzeit": "05:00", "heim": "Türkei",         "gast": "Paraguay",            "gruppe": "D"},
    {"id": "D5", "datum": "26.06.", "uhrzeit": "04:00", "heim": "Türkei",         "gast": "USA",                 "gruppe": "D"},
    {"id": "D6", "datum": "26.06.", "uhrzeit": "04:00", "heim": "Paraguay",       "gast": "Australien",          "gruppe": "D"},
    {"id": "E1", "datum": "14.06.", "uhrzeit": "19:00", "heim": "Deutschland",    "gast": "Curaçao",             "gruppe": "E"},
    {"id": "E2", "datum": "15.06.", "uhrzeit": "01:00", "heim": "Elfenbeinküste", "gast": "Ecuador",             "gruppe": "E"},
    {"id": "E3", "datum": "20.06.", "uhrzeit": "22:00", "heim": "Deutschland",    "gast": "Elfenbeinküste",      "gruppe": "E"},
    {"id": "E4", "datum": "21.06.", "uhrzeit": "02:00", "heim": "Ecuador",        "gast": "Curaçao",             "gruppe": "E"},
    {"id": "E5", "datum": "25.06.", "uhrzeit": "22:00", "heim": "Curaçao",        "gast": "Elfenbeinküste",      "gruppe": "E"},
    {"id": "E6", "datum": "25.06.", "uhrzeit": "22:00", "heim": "Ecuador",        "gast": "Deutschland",         "gruppe": "E"},
    {"id": "F1", "datum": "14.06.", "uhrzeit": "22:00", "heim": "Niederlande",    "gast": "Japan",               "gruppe": "F"},
    {"id": "F2", "datum": "15.06.", "uhrzeit": "04:00", "heim": "Schweden",       "gast": "Tunesien",            "gruppe": "F"},
    {"id": "F3", "datum": "20.06.", "uhrzeit": "19:00", "heim": "Niederlande",    "gast": "Schweden",            "gruppe": "F"},
    {"id": "F4", "datum": "21.06.", "uhrzeit": "06:00", "heim": "Tunesien",       "gast": "Japan",               "gruppe": "F"},
    {"id": "F5", "datum": "26.06.", "uhrzeit": "01:00", "heim": "Japan",          "gast": "Schweden",            "gruppe": "F"},
    {"id": "F6", "datum": "26.06.", "uhrzeit": "01:00", "heim": "Tunesien",       "gast": "Niederlande",         "gruppe": "F"},
    {"id": "G1", "datum": "15.06.", "uhrzeit": "21:00", "heim": "Belgien",        "gast": "Ägypten",             "gruppe": "G"},
    {"id": "G2", "datum": "16.06.", "uhrzeit": "03:00", "heim": "Iran",           "gast": "Neuseeland",          "gruppe": "G"},
    {"id": "G3", "datum": "21.06.", "uhrzeit": "21:00", "heim": "Belgien",        "gast": "Iran",                "gruppe": "G"},
    {"id": "G4", "datum": "22.06.", "uhrzeit": "03:00", "heim": "Neuseeland",     "gast": "Ägypten",             "gruppe": "G"},
    {"id": "G5", "datum": "27.06.", "uhrzeit": "05:00", "heim": "Ägypten",        "gast": "Iran",                "gruppe": "G"},
    {"id": "G6", "datum": "27.06.", "uhrzeit": "05:00", "heim": "Neuseeland",     "gast": "Belgien",             "gruppe": "G"},
    {"id": "H1", "datum": "15.06.", "uhrzeit": "18:00", "heim": "Spanien",        "gast": "Kap Verde",           "gruppe": "H"},
    {"id": "H2", "datum": "16.06.", "uhrzeit": "00:00", "heim": "Saudi-Arabien",  "gast": "Uruguay",             "gruppe": "H"},
    {"id": "H3", "datum": "21.06.", "uhrzeit": "18:00", "heim": "Spanien",        "gast": "Saudi-Arabien",       "gruppe": "H"},
    {"id": "H4", "datum": "22.06.", "uhrzeit": "00:00", "heim": "Uruguay",        "gast": "Kap Verde",           "gruppe": "H"},
    {"id": "H5", "datum": "27.06.", "uhrzeit": "02:00", "heim": "Kap Verde",      "gast": "Saudi-Arabien",       "gruppe": "H"},
    {"id": "H6", "datum": "27.06.", "uhrzeit": "02:00", "heim": "Uruguay",        "gast": "Spanien",             "gruppe": "H"},
    {"id": "I1", "datum": "16.06.", "uhrzeit": "21:00", "heim": "Frankreich",     "gast": "Senegal",             "gruppe": "I"},
    {"id": "I2", "datum": "17.06.", "uhrzeit": "00:00", "heim": "Irak",           "gast": "Norwegen",            "gruppe": "I"},
    {"id": "I3", "datum": "22.06.", "uhrzeit": "23:00", "heim": "Frankreich",     "gast": "Irak",                "gruppe": "I"},
    {"id": "I4", "datum": "23.06.", "uhrzeit": "02:00", "heim": "Norwegen",       "gast": "Senegal",             "gruppe": "I"},
    {"id": "I5", "datum": "26.06.", "uhrzeit": "21:00", "heim": "Norwegen",       "gast": "Frankreich",          "gruppe": "I"},
    {"id": "I6", "datum": "26.06.", "uhrzeit": "21:00", "heim": "Senegal",        "gast": "Irak",                "gruppe": "I"},
    {"id": "J1", "datum": "17.06.", "uhrzeit": "03:00", "heim": "Argentinien",    "gast": "Algerien",            "gruppe": "J"},
    {"id": "J2", "datum": "17.06.", "uhrzeit": "06:00", "heim": "Österreich",     "gast": "Jordanien",           "gruppe": "J"},
    {"id": "J3", "datum": "22.06.", "uhrzeit": "19:00", "heim": "Argentinien",    "gast": "Österreich",          "gruppe": "J"},
    {"id": "J4", "datum": "23.06.", "uhrzeit": "05:00", "heim": "Jordanien",      "gast": "Algerien",            "gruppe": "J"},
    {"id": "J5", "datum": "28.06.", "uhrzeit": "04:00", "heim": "Algerien",       "gast": "Österreich",          "gruppe": "J"},
    {"id": "J6", "datum": "28.06.", "uhrzeit": "04:00", "heim": "Jordanien",      "gast": "Argentinien",         "gruppe": "J"},
    {"id": "K1", "datum": "17.06.", "uhrzeit": "19:00", "heim": "Portugal",       "gast": "DR Kongo",            "gruppe": "K"},
    {"id": "K2", "datum": "18.06.", "uhrzeit": "04:00", "heim": "Usbekistan",     "gast": "Kolumbien",           "gruppe": "K"},
    {"id": "K3", "datum": "23.06.", "uhrzeit": "19:00", "heim": "Portugal",       "gast": "Usbekistan",          "gruppe": "K"},
    {"id": "K4", "datum": "24.06.", "uhrzeit": "04:00", "heim": "Kolumbien",      "gast": "DR Kongo",            "gruppe": "K"},
    {"id": "K5", "datum": "28.06.", "uhrzeit": "01:30", "heim": "Kolumbien",      "gast": "Portugal",            "gruppe": "K"},
    {"id": "K6", "datum": "28.06.", "uhrzeit": "01:30", "heim": "DR Kongo",       "gast": "Usbekistan",          "gruppe": "K"},
    {"id": "L1", "datum": "17.06.", "uhrzeit": "22:00", "heim": "England",        "gast": "Kroatien",            "gruppe": "L"},
    {"id": "L2", "datum": "18.06.", "uhrzeit": "01:00", "heim": "Ghana",          "gast": "Panama",              "gruppe": "L"},
    {"id": "L3", "datum": "23.06.", "uhrzeit": "22:00", "heim": "England",        "gast": "Ghana",               "gruppe": "L"},
    {"id": "L4", "datum": "24.06.", "uhrzeit": "01:00", "heim": "Panama",         "gast": "Kroatien",            "gruppe": "L"},
    {"id": "L5", "datum": "27.06.", "uhrzeit": "23:00", "heim": "Panama",         "gast": "England",             "gruppe": "L"},
    {"id": "L6", "datum": "27.06.", "uhrzeit": "23:00", "heim": "Kroatien",       "gast": "Ghana",               "gruppe": "L"},
]

# ============================================================
# SPALTENNAMEN  (müssen exakt mit den Forms-Fragen übereinstimmen)
# ============================================================

NAME_COL = "Name (Vorname Nachname)"

def match_col(m):
    return f"[Gr.{m['gruppe']}] {m['heim']} vs {m['gast']} ({m['datum']})"

KO_COLS = {
    "S16": [f"[S16] Weiterkommer {i}" for i in range(1, 17)],
    "S8":  [f"[S8] Weiterkommer {i}"  for i in range(1, 9)],
    "VF":  [f"[VF] Weiterkommer {i}"  for i in range(1, 5)],
    "HF":  [f"[HF] Weiterkommer {i}"  for i in range(1, 3)],
    "F":   ["[F] Finalist 1", "[F] Finalist 2"],
    "WM":  ["[WM] Weltmeister"],
}

KO_PUNKTE = {"S16": 3, "S8": 5, "VF": 10, "HF": 15}

# ============================================================
# PUNKTE BERECHNUNG
# ============================================================

def parse_score(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    m = re.match(r"^\s*(\d+)\s*[:–\-]\s*(\d+)\s*$", str(s).strip())
    return (int(m.group(1)), int(m.group(2))) if m else None

def calc_group_pts(tipp, result):
    t, r = parse_score(tipp), parse_score(result)
    if t is None or r is None:
        return 0
    th, tg = t
    rh, rg = r
    if th == rh and tg == rg:
        return 4
    if (th - tg) == (rh - rg):
        return 3
    if (th > tg) == (rh > rg) and not (th == tg and rh != rg) and not (th != tg and rh == rg):
        return 2
    if th == tg and rh == rg:
        return 2
    return 0

def norm(name):
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    return str(name).strip().lower()

def calc_ko_pts(pred_list, actual_set, pts_each):
    pred = {norm(t) for t in pred_list if norm(t)}
    actual = {norm(t) for t in actual_set if norm(t)}
    return len(pred & actual) * pts_each

def calc_finale_pts(pred_list, actual_set):
    pred = {norm(t) for t in pred_list if norm(t)}
    actual = {norm(t) for t in actual_set if norm(t)}
    return 20 if len(actual) == 2 and pred == actual else 0

def calc_wm_pts(pred, actual):
    return 25 if norm(pred) and norm(actual) and norm(pred) == norm(actual) else 0

# ============================================================
# ERGEBNISSE EINLESEN
# ============================================================

def read_ergebnisse(path):
    wb = openpyxl.load_workbook(path)

    # Gruppenphase
    gs = wb["Gruppenphase"]
    results = {}
    for row in gs.iter_rows(min_row=2, values_only=True):
        match_id, _, _, _, score = row[0], row[1], row[2], row[3], row[4]
        if match_id and score:
            results[str(match_id)] = str(score).strip()

    # KO-Runden
    ko_results = {}
    for runde in ["S16", "S8", "VF", "HF", "F", "WM"]:
        if runde in wb.sheetnames:
            sheet = wb[runde]
            teams = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # Column A = slot number, Column B = team name
                val = row[1] if len(row) > 1 else row[0]
                if val:
                    teams.append(str(val).strip())
            ko_results[runde] = teams

    return results, ko_results

# ============================================================
# JSON TIPPS EINLESEN  (neues Format: tippzettel.html)
# ============================================================

def json_to_row(data):
    row = {NAME_COL: data.get("name", "Unbekannt")}
    for key, val in data.get("gruppenphase", {}).items():
        row[key] = val
    for runde_id, cols in KO_COLS.items():
        if runde_id == "WM":
            row[cols[0]] = data.get("WM", "")
        else:
            teams = data.get(runde_id, [])
            if not isinstance(teams, list):
                teams = []
            for i, col in enumerate(cols):
                row[col] = teams[i] if i < len(teams) else ""
    return row

def read_json_tipps(path):
    p = Path(path)
    if p.is_dir():
        files = sorted(p.glob("*.json"))
        if not files:
            raise FileNotFoundError(f"Keine JSON-Dateien in {p} gefunden.")
        rows = [json_to_row(json.loads(f.read_text(encoding="utf-8"))) for f in files]
        print(f"  JSON-Dateien gelesen: {len(rows)}")
    else:
        rows = [json_to_row(json.loads(p.read_text(encoding="utf-8")))]
        print(f"  JSON-Datei gelesen: {p.name}")
    return pd.DataFrame(rows)

# ============================================================
# FORMS-RESPONSES EINLESEN
# ============================================================

def read_responses(path):
    df = pd.read_excel(path)
    # Microsoft Forms fügt Metaspalten vorne hinzu – Name-Spalte suchen
    if NAME_COL not in df.columns:
        print(f"WARNUNG: Spalte '{NAME_COL}' nicht gefunden.")
        print(f"Vorhandene Spalten: {list(df.columns[:10])}")
    return df

# ============================================================
# AUSWERTUNG
# ============================================================

def auswerten(df, gs_results, ko_results):
    rows = []
    for _, resp in df.iterrows():
        name = str(resp.get(NAME_COL, "Unbekannt")).strip()
        if not name or name == "nan":
            continue

        pts_gruppe = 0
        detail_gruppe = {}
        for m in GRUPPENSPIELE:
            col = match_col(m)
            tipp = resp.get(col)
            result = gs_results.get(m["id"])
            p = calc_group_pts(tipp, result)
            pts_gruppe += p
            detail_gruppe[m["id"]] = {
                "tipp": tipp, "result": result, "punkte": p,
                "heim": m["heim"], "gast": m["gast"], "datum": m["datum"]
            }

        pts_ko = {}
        for runde, pts_each in KO_PUNKTE.items():
            pred = [resp.get(c) for c in KO_COLS[runde]]
            actual = ko_results.get(runde, [])
            pts_ko[runde] = calc_ko_pts(pred, actual, pts_each)

        pred_finale = [resp.get(c) for c in KO_COLS["F"]]
        actual_finale = ko_results.get("F", [])
        pts_ko["F"] = calc_finale_pts(pred_finale, actual_finale)

        pts_ko["WM"] = calc_wm_pts(resp.get(KO_COLS["WM"][0]), ko_results.get("WM", [None])[0] if ko_results.get("WM") else None)

        total = pts_gruppe + sum(pts_ko.values())
        rows.append({
            "Name": name,
            "Gesamt": total,
            "Gruppenphase": pts_gruppe,
            "S16": pts_ko.get("S16", 0),
            "S8":  pts_ko.get("S8",  0),
            "VF":  pts_ko.get("VF",  0),
            "HF":  pts_ko.get("HF",  0),
            "Finale": pts_ko.get("F", 0),
            "Weltmeister": pts_ko.get("WM", 0),
            "_detail": detail_gruppe,
            "_resp": resp,
        })

    rows.sort(key=lambda x: x["Gesamt"], reverse=True)
    for i, r in enumerate(rows):
        r["Platz"] = i + 1
    return rows

# ============================================================
# EXCEL OUTPUT
# ============================================================

# Farben
CLR_HEADER   = "1A3A5C"   # Dunkelblau
CLR_GOLD     = "FFD700"
CLR_SILVER   = "C0C0C0"
CLR_BRONZE   = "CD7F32"
CLR_GREEN    = "D4EDDA"
CLR_YELLOW   = "FFF3CD"
CLR_RED      = "F8D7DA"
CLR_LIGHT    = "F0F4F8"
CLR_WHITE    = "FFFFFF"

def header_font():    return Font(bold=True, color="FFFFFF", name="Calibri", size=11)
def cell_font():      return Font(name="Calibri", size=10)
def bold_font():      return Font(bold=True, name="Calibri", size=10)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def write_rangliste(wb, rows):
    ws = wb.create_sheet("Rangliste")
    ws.sheet_view.showGridLines = False

    headers = ["Platz", "Name", "Gesamt", "Gruppenphase",
               "Sechzehntelfinale", "Achtelfinale", "Viertelfinale",
               "Halbfinale", "Finale", "Weltmeister"]
    col_widths = [8, 22, 10, 14, 18, 14, 14, 12, 10, 14]

    # Header
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = header_font()
        cell.fill = fill(CLR_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border()
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22

    # Daten
    for r in rows:
        row_num = r["Platz"] + 1
        values = [r["Platz"], r["Name"], r["Gesamt"], r["Gruppenphase"],
                  r["S16"], r["S8"], r["VF"], r["HF"], r["Finale"], r["Weltmeister"]]

        platz_fill = CLR_LIGHT
        if r["Platz"] == 1: platz_fill = CLR_GOLD
        elif r["Platz"] == 2: platz_fill = CLR_SILVER
        elif r["Platz"] == 3: platz_fill = CLR_BRONZE

        for c, v in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=c, value=v)
            cell.font = bold_font() if c <= 3 else cell_font()
            cell.fill = fill(platz_fill if c <= 3 else (CLR_LIGHT if row_num % 2 == 0 else CLR_WHITE))
            cell.alignment = Alignment(horizontal="center" if c != 2 else "left", vertical="center")
            cell.border = thin_border()

    ws.freeze_panes = "A2"

def write_details(wb, rows, gs_results):
    ws = wb.create_sheet("Details Gruppenphase")
    ws.sheet_view.showGridLines = False

    # Kopfzeile
    ws.cell(1, 1, "Spiel").font = header_font()
    ws.cell(1, 1).fill = fill(CLR_HEADER)
    ws.cell(1, 1).border = thin_border()
    ws.column_dimensions["A"].width = 30

    ws.cell(1, 2, "Ergebnis").font = header_font()
    ws.cell(1, 2).fill = fill(CLR_HEADER)
    ws.cell(1, 2).border = thin_border()
    ws.column_dimensions["B"].width = 12

    for i, r in enumerate(rows):
        col = i + 3
        ws.cell(1, col, r["Name"]).font = header_font()
        ws.cell(1, col).fill = fill(CLR_HEADER)
        ws.cell(1, col).alignment = Alignment(horizontal="center")
        ws.cell(1, col).border = thin_border()
        ws.column_dimensions[get_column_letter(col)].width = 14

    gruppe = None
    data_row = 2
    for m in GRUPPENSPIELE:
        if m["gruppe"] != gruppe:
            gruppe = m["gruppe"]
            ws.cell(data_row, 1, f"── Gruppe {gruppe} ──").font = Font(bold=True, name="Calibri", size=10, color=CLR_HEADER)
            ws.cell(data_row, 1).fill = fill("E8EFF7")
            ws.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=len(rows)+2)
            data_row += 1

        label = f"{m['heim']} vs {m['gast']} ({m['datum']})"
        ws.cell(data_row, 1, label).font = cell_font()
        ws.cell(data_row, 1).border = thin_border()

        actual = gs_results.get(m["id"], "–")
        ws.cell(data_row, 2, actual).font = bold_font()
        ws.cell(data_row, 2).alignment = Alignment(horizontal="center")
        ws.cell(data_row, 2).border = thin_border()

        for i, r in enumerate(rows):
            col = i + 3
            detail = r["_detail"].get(m["id"], {})
            tipp = detail.get("tipp", "–") or "–"
            pts  = detail.get("punkte", 0)
            cell_val = f"{tipp}  (+{pts})" if actual != "–" else str(tipp)
            c = ws.cell(data_row, col, cell_val)
            c.font = cell_font()
            c.alignment = Alignment(horizontal="center")
            c.border = thin_border()
            if actual != "–":
                if pts == 4:   c.fill = fill(CLR_GREEN)
                elif pts == 3: c.fill = fill("C8E6C9")
                elif pts == 2: c.fill = fill(CLR_YELLOW)
                else:          c.fill = fill(CLR_RED)

        data_row += 1

    ws.freeze_panes = "C2"

def write_ko_details(wb, rows, ko_results):
    ws = wb.create_sheet("Details KO-Runden")
    ws.sheet_view.showGridLines = False

    runden_info = [
        ("S16", "Sechzehntelfinale", 16, 5),
        ("S8",  "Achtelfinale",       8, 10),
        ("VF",  "Viertelfinale",      4, 15),
        ("HF",  "Halbfinale",         2, 20),
        ("F",   "Finale (beide Finalisten)", 2, 25),
        ("WM",  "Weltmeister",        1, 30),
    ]

    ws.cell(1, 1, "KO-Runde / Slot").font = header_font()
    ws.cell(1, 1).fill = fill(CLR_HEADER)
    ws.cell(1, 1).border = thin_border()
    ws.column_dimensions["A"].width = 32

    ws.cell(1, 2, "Tatsächlich").font = header_font()
    ws.cell(1, 2).fill = fill(CLR_HEADER)
    ws.cell(1, 2).border = thin_border()
    ws.column_dimensions["B"].width = 18

    for i, r in enumerate(rows):
        col = i + 3
        ws.cell(1, col, r["Name"]).font = header_font()
        ws.cell(1, col).fill = fill(CLR_HEADER)
        ws.cell(1, col).alignment = Alignment(horizontal="center")
        ws.cell(1, col).border = thin_border()
        ws.column_dimensions[get_column_letter(col)].width = 16

    data_row = 2
    for runde_id, runde_label, slots, pts in runden_info:
        # Sektionszeile
        ws.cell(data_row, 1, f"── {runde_label} (je {pts} Pkt.) ──").font = Font(bold=True, name="Calibri", size=10, color=CLR_HEADER)
        ws.cell(data_row, 1).fill = fill("E8EFF7")
        ws.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=len(rows)+2)
        data_row += 1

        actual_teams = {norm(t) for t in ko_results.get(runde_id, []) if norm(t)}

        for i, r in enumerate(rows):
            resp = r["_resp"]
            cols = KO_COLS[runde_id]
            pred = [resp.get(c) for c in cols]

            # Tipps als komma-getrennte Liste in einer Zeile pro Teilnehmer
            # Schreibe nur eine Übersichtszeile pro Runde (alle Tipps zusammen)
            pass

        # Eine Zeile pro Slot
        for slot_idx in range(slots):
            ws.cell(data_row, 1, f"  Weiterkommer {slot_idx + 1}").font = cell_font()
            ws.cell(data_row, 1).border = thin_border()

            actual_at_slot = list(ko_results.get(runde_id, []))
            actual_val = actual_at_slot[slot_idx] if slot_idx < len(actual_at_slot) else "–"
            ws.cell(data_row, 2, actual_val).font = bold_font()
            ws.cell(data_row, 2).alignment = Alignment(horizontal="center")
            ws.cell(data_row, 2).border = thin_border()

            for i, r in enumerate(rows):
                resp = r["_resp"]
                col_key = KO_COLS[runde_id][slot_idx] if slot_idx < len(KO_COLS[runde_id]) else None
                pred_val = resp.get(col_key) if col_key else None
                pred_norm = norm(pred_val)
                is_correct = pred_norm and pred_norm in actual_teams

                c = ws.cell(data_row, i + 3, str(pred_val) if pred_val and not pd.isna(pred_val) else "–")
                c.font = cell_font()
                c.alignment = Alignment(horizontal="center")
                c.border = thin_border()
                if actual_teams:
                    c.fill = fill(CLR_GREEN if is_correct else CLR_RED)

            data_row += 1

    ws.freeze_panes = "C2"

# ============================================================
# HTML RANGLISTE  – interaktiv mit Spielerfilter & Detailansicht
# ============================================================

def write_html_rangliste(rows, gs_results, ko_results, out_path):
    import json as _json

    def _clean(val):
        if val is None:
            return ""
        try:
            if pd.isna(val):
                return ""
        except (TypeError, ValueError):
            pass
        return str(val).strip()

    # ── Upcoming-Match-Erkennung ──────────────────────────
    _now  = datetime.now()
    _soon = (_now + timedelta(days=2)).date()
    def _is_kommend(datum_str, has_result):
        if has_result:
            return False
        try:
            return datetime.strptime(datum_str.rstrip('.') + '.2026', '%d.%m.%Y').date() <= _soon
        except Exception:
            return False

    # ── Daten für JavaScript aufbereiten ──────────────────
    players_js = []
    for r in rows:
        spiele = []
        for m in GRUPPENSPIELE:
            d = r["_detail"].get(m["id"], {})
            ergebnis = _clean(d.get("result"))
            tipp     = _clean(d.get("tipp"))
            pts      = d.get("punkte", 0) if ergebnis else None
            spiele.append({
                "id": m["id"], "gr": m["gruppe"],
                "heim": m["heim"], "gast": m["gast"], "datum": m["datum"], "uhrzeit": m.get("uhrzeit", ""),
                "tipp": tipp, "ergebnis": ergebnis, "punkte": pts,
                "kommend": _is_kommend(m["datum"], bool(ergebnis)),
            })

        ko_tipps = {}
        for runde, cols in KO_COLS.items():
            if runde == "WM":
                ko_tipps[runde] = _clean(r["_resp"].get(cols[0]))
            else:
                ko_tipps[runde] = [_clean(r["_resp"].get(c)) for c in cols]

        players_js.append({
            "platz": r["Platz"], "name": r["Name"],
            "gesamt": r["Gesamt"], "gruppe": r["Gruppenphase"],
            "s16": r["S16"], "s8": r["S8"],
            "vf": r["VF"], "hf": r["HF"],
            "finale": r["Finale"], "wm": r["Weltmeister"],
            "spiele": spiele, "ko_tipps": ko_tipps,
        })

    ko_res_js = {
        k: [str(t).strip().lower() for t in v if t and str(t).strip()]
        for k, v in ko_results.items()
    }

    filled  = sum(1 for v in gs_results.values() if v)
    ts      = datetime.now().strftime("%d.%m.%Y %H:%M Uhr")
    leader  = rows[0]["Name"] if rows else "–"
    lpts    = rows[0]["Gesamt"] if rows else 0
    n       = len(rows)
    data_json = _json.dumps({"players": players_js, "ko": ko_res_js}, ensure_ascii=False)

    # ── CSS (plain string – keine f-string-Escapes nötig) ─
    CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#152438;color:#eef5fd;min-height:100vh;font-size:15px}
.hdr{background:linear-gradient(135deg,#1e3f60,#0e2848);border-bottom:2px solid #c8a200;padding:16px 24px;display:flex;align-items:center;gap:14px}
.hdr-logo{height:60px;width:auto;filter:drop-shadow(0 2px 8px rgba(0,0,0,.6));flex-shrink:0}.hdr-title{font-size:1.5rem;font-weight:800;color:#f5c518}.hdr-sub{font-size:.82rem;color:#a0c0de;margin-top:3px}
.stats{display:flex;gap:12px;flex-wrap:wrap;padding:14px 24px;background:#182d45;border-bottom:1px solid #2e4e72}
.stat-box{background:#1c3450;border:1px solid #2e4e72;border-radius:8px;padding:10px 16px;min-width:130px}
.stat-val{font-size:1.4rem;font-weight:800;color:#f5c518}.stat-lbl{font-size:.75rem;color:#a0c0de;margin-top:2px}
.filter-bar{padding:12px 24px;background:#182d45;border-bottom:1px solid #2e4e72;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.filter-lbl{color:#a0c0de;font-size:.82rem;font-weight:600;text-transform:uppercase;letter-spacing:.05em}
.drop-wrap{position:relative}
.drop-btn{background:#1c3450;border:1px solid #2e4e72;color:#eef5fd;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:.82rem;min-width:210px;text-align:left;display:flex;align-items:center;justify-content:space-between;gap:10px;transition:all .2s}
.drop-btn:hover,.drop-btn.open{background:#2a4e78;border-color:#4fc3f7}
.drop-panel{display:none;position:absolute;top:calc(100% + 5px);left:0;min-width:230px;max-height:320px;background:#1c3450;border:1px solid #2e4e72;border-radius:8px;z-index:200;box-shadow:0 8px 28px rgba(0,0,0,.5);overflow:hidden}
.drop-panel.open{display:flex;flex-direction:column}
.drop-hdr{display:flex;border-bottom:1px solid #2e4e72;flex-shrink:0}
.drop-hdr-btn{flex:1;padding:7px;font-size:.75rem;color:#a0c0de;background:none;border:none;border-right:1px solid #2e4e72;cursor:pointer;transition:all .15s}
.drop-hdr-btn:last-child{border-right:none}.drop-hdr-btn:hover{color:#f5c518;background:#2a4e78}
.drop-opts{overflow-y:auto;flex:1}
.dlabel{display:flex;align-items:center;gap:9px;padding:8px 14px;cursor:pointer;font-size:.83rem;color:#eef5fd;transition:background .12s;user-select:none}
.dlabel:hover{background:#2a4e78}.dlabel input[type=checkbox]{accent-color:#4fc3f7;width:14px;height:14px;cursor:pointer;flex-shrink:0}
.hint{font-size:.75rem;color:#546e7a}
.section{padding:20px 24px}
.sec-title{font-size:.88rem;font-weight:700;color:#f5c518;text-transform:uppercase;letter-spacing:.08em;margin-bottom:14px;display:flex;align-items:center;gap:10px}
.sec-title::after{content:'';flex:1;height:1px;background:linear-gradient(to right,#1c3350,transparent)}
.wrap{overflow-x:auto}
.rtbl{width:100%;border-collapse:collapse;min-width:500px}
.rtbl th{background:#2a4e78;color:#f5c518;padding:9px 10px;text-align:right;font-size:.78rem;font-weight:600;white-space:nowrap}
.rtbl th:first-child,.rtbl th:nth-child(2){text-align:left}
.rtbl td{padding:8px 10px;text-align:right;border-bottom:1px solid #182d45;font-size:.88rem;transition:background .15s}
.rtbl td.pl,.rtbl td.nm{text-align:left}.rtbl td.pts{font-weight:700;font-size:1rem;color:#4fc3f7}
.rtbl tr{cursor:pointer}.rtbl tr:hover td{background:#1e3a5c!important}
.rtbl tr.sel td{box-shadow:inset 3px 0 0 #4fc3f7}
.rtbl tr.sel td.nm{color:#f5c518;font-weight:700}
.rtbl tr.r1 td{background:#3a2a00}.rtbl tr.r2 td{background:#1c2638}.rtbl tr.r3 td{background:#182a3e}
.no-sel{text-align:center;padding:44px 16px;color:#546e7a}
.no-sel-icon{font-size:2.4rem;margin-bottom:10px}.no-sel-txt{font-size:.9rem;color:#7a9bbe}
.badge{display:inline-block;padding:2px 9px;border-radius:12px;font-size:.72rem;font-weight:700;white-space:nowrap}
.badge.ex{background:#1b5e20;color:#a5d6a7;border:1px solid #2e7d32}
.badge.df{background:#004d40;color:#80cbc4;border:1px solid #00695c}
.badge.td{background:#bf360c;color:#ffccbc;border:1px solid #e64a19}
.badge.ms{background:#3e0000;color:#ef9a9a;border:1px solid #b71c1c}
.badge.op{background:#1c3048;color:#7a9bbe;border:1px solid #2e4a60}
.pcard{background:#182d45;border:1px solid #2e4e72;border-radius:10px;margin-bottom:18px;overflow:hidden}
.pcard.hidden{display:none}
.phdr{background:linear-gradient(90deg,#1e3f60,#182d45);padding:14px 18px;display:flex;align-items:center;gap:12px;cursor:pointer;user-select:none;border-bottom:1px solid #2e4e72}
.phdr:hover{background:#2a4e78}
.pmedal{font-size:1.6rem;flex-shrink:0}.pinfo{flex:1}.pname{font-size:1.05rem;font-weight:700;color:#f5c518}
.pbreakdown{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px}
.pbitem{font-size:.72rem;color:#a0c0de}.pbval{color:#eef5fd;font-weight:600}
.ptotal{font-size:1.4rem;font-weight:800;color:#4fc3f7;flex-shrink:0}
.pbody{padding:16px 18px}
.gr-tabs{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:12px}
.gtab{padding:3px 10px;background:#1c3450;border:1px solid #2e4e72;border-radius:5px;font-size:.75rem;cursor:pointer;color:#a0c0de;transition:all .15s}
.gtab:hover{background:#2a4e78}.gtab.active{background:#2a4e78;border-color:#4fc3f7;color:#f5c518;font-weight:700}
.gtbl{width:100%;border-collapse:collapse;min-width:450px}
.gtbl th{background:#1c3450;color:#7a9bbe;padding:7px 10px;font-size:.72rem;text-align:left;font-weight:600;text-transform:uppercase;white-space:nowrap}
.gtbl td{padding:7px 10px;border-bottom:1px solid #182d45;font-size:.84rem;vertical-align:middle}
.gtbl tr:hover td{background:#1e3a5c}.gtbl tr.gr-hide{display:none}
.sc{font-family:monospace;font-size:.9rem}.sc-t{color:#a0b8c8}.sc-r{color:#4fc3f7;font-weight:700}.arrow{color:#546e7a;margin:0 3px}
.ko-wrap{margin-top:16px;border-top:1px solid #2e4e72;padding-top:14px}
.ko-sec-title{font-size:.75rem;font-weight:700;color:#a0c0de;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px}
.ko-row{display:flex;align-items:flex-start;gap:10px;margin-bottom:8px}
.ko-rname{width:135px;font-size:.75rem;color:#7a9bbe;font-weight:600;text-transform:uppercase;padding-top:3px;flex-shrink:0}
.ko-rdate{font-size:.68rem;color:#546e7a;font-weight:400;text-transform:none}
.de-row td{background:rgba(245,197,24,.18)!important;border-left:4px solid #f5c518}
.de-row{box-shadow:inset 0 0 16px rgba(245,197,24,.10)}
.ko-teams{display:flex;flex-wrap:wrap;gap:5px}
.kt{padding:2px 10px;border-radius:12px;font-size:.75rem;background:#1c3450;border:1px solid #2e4e72;color:#eef5fd}
.kt.hit{background:#1b5e20;border-color:#2e7d32;color:#a5d6a7}.kt.miss{background:#3e0000;border-color:#b71c1c;color:#ef9a9a}
.foot{text-align:center;padding:18px;color:#546e7a;font-size:.78rem;border-top:1px solid #182d45}
@media(max-width:600px){.section{padding:12px}.filter-bar{padding:10px 12px}.pbody{padding:12px}}
.gt-bar{display:flex;flex-wrap:wrap;gap:5px;padding:10px 0 8px}
.mtx-wrap{overflow-x:auto}
.mtx{width:100%;border-collapse:collapse;font-size:.82rem}
.mtx th{background:#1c3450;color:#a0c0de;padding:6px 8px;border:1px solid #2e4e72;text-align:center;white-space:nowrap;font-weight:600}
.mtx th.mh-l{text-align:left}
.mtx td{padding:5px 7px;border:1px solid #1e3550;vertical-align:middle;text-align:center}
.mtx td.mi{text-align:left;color:#7a9bbe;font-size:.78rem;white-space:nowrap}
.mtx td.mi-t{text-align:left;color:#c8d8e8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px}
.mtx td.mi-r{text-align:center;font-weight:700;color:#f5c518;min-width:52px;max-width:52px}
.mtx tr.de-row td{background:rgba(245,197,24,.05)!important}
.mtx tr.kommend>td:first-child{border-left:3px solid #f5c518}
.mtx tr.kommend .mi{color:#f5c518}
.mcel{font-size:.82rem}
.mcel-ex{background:rgba(39,174,96,.22);color:#5dde8a}
.mcel-df{background:rgba(79,195,247,.15);color:#4fc3f7}
.mcel-td{background:rgba(255,167,38,.15);color:#ffa726}
.mcel-ms{background:rgba(180,40,40,.1);color:#7a9bbe}
.mcel-op{color:#c8d8e8}
.mcel-sub{display:block;font-size:.68rem;opacity:.8;margin-top:2px}
.mtx tfoot td{background:#182d45;border-top:2px solid #2e4e72;font-weight:700;color:#f5c518;text-align:center;padding:6px 8px}
.mtx tfoot td.sub-lbl{text-align:left;color:#7a9bbe;font-weight:400;font-size:.78rem}
.mpl-name{font-size:.8rem;font-weight:600}
.mpl-pts{font-size:.72rem;color:#f5c518;display:block;margin-top:2px}
.mtx-grp th:nth-child(-n+3){position:sticky;z-index:3;background:#1c3450}
.mtx-grp td:nth-child(-n+3){position:sticky;z-index:2;background:#152438}
.mtx-grp th:nth-child(1),.mtx-grp td:nth-child(1){left:0;min-width:95px;max-width:95px}
.mtx-grp th:nth-child(2),.mtx-grp td:nth-child(2){left:95px;min-width:180px;max-width:180px}
.mtx-grp th:nth-child(3),.mtx-grp td:nth-child(3){left:275px;min-width:52px;max-width:52px;box-shadow:4px 0 8px rgba(0,0,0,.45)}
.mtx-grp tr.de-row td:nth-child(-n+3){background:#1a2d42!important}
.mtx-grp tr.kommend td:nth-child(-n+3){background:#16202e!important}
.mtx-ko th:nth-child(-n+2){position:sticky;z-index:3;background:#1c3450}
.mtx-ko td:nth-child(-n+2){position:sticky;z-index:2;background:#152438}
.mtx-ko th:nth-child(1),.mtx-ko td:nth-child(1){left:0;min-width:36px;max-width:36px}
.mtx-ko th:nth-child(2),.mtx-ko td:nth-child(2){left:36px;box-shadow:4px 0 8px rgba(0,0,0,.45)}
"""

    # ── JavaScript (plain string – kein f-string-Escaping) ─
    JS = """
const MEDALS=["🥇","🥈","🥉"];
const GROUPS=["A","B","C","D","E","F","G","H","I","J","K","L"];
const KO_ROUNDS=[["S16","Sechzehntelfinale","28.06.–04.07."],["S8","Achtelfinale","04.–07.07."],
  ["VF","Viertelfinale","09.–12.07."],["HF","Halbfinale","14.–15.07."],["F","Finale","19.07."],["WM","Weltmeister",""]];
let selected=new Set(DATA.players.map(p=>p.name));

function badge(pts){
  if(pts===null||pts===undefined) return '<span class="badge op">· Offen</span>';
  const cfg={4:["ex","⚽ Exakt +4"],3:["df","✓ Tordiff +3"],2:["td","≈ Tendenz +2"],0:["ms","✗ Daneben"]};
  const [cls,lbl]=cfg[pts]||["op","?"];
  return `<span class="badge ${cls}">${lbl}</span>`;
}

function renderDropdown(){
  const sorted=[...DATA.players].sort((a,b)=>a.name.localeCompare(b.name,"de"));
  const cnt=selected.size;
  const btn=document.getElementById("dropBtn");
  if(cnt===0)
    btn.innerHTML='Spieler auswählen <span style="opacity:.5;margin-left:4px">&#9662;</span>';
  else if(cnt===DATA.players.length)
    btn.innerHTML=`Alle (${cnt}) <span style="color:#4fc3f7;margin-left:4px">&#9662;</span>`;
  else
    btn.innerHTML=`<span style="color:#f5c518">${cnt} ausgewählt</span><span style="opacity:.5;margin-left:4px">&#9662;</span>`;
  document.getElementById("dropList").innerHTML=sorted.map(p=>
    `<label class="dlabel">
      <input type="checkbox" ${selected.has(p.name)?"checked":""}
        onchange="togglePlayerDrop('${p.name.replace(/'/g,"\\\\'")}',this.checked)">
      ${p.name}
    </label>`
  ).join("");
}

function toggleDropdown(e){
  e.stopPropagation();
  const panel=document.getElementById("dropPanel");
  const btn=document.getElementById("dropBtn");
  const open=panel.classList.toggle("open");
  btn.classList.toggle("open",open);
}

document.addEventListener("click",e=>{
  const wrap=document.getElementById("dropWrap");
  if(wrap&&!wrap.contains(e.target)){
    document.getElementById("dropPanel").classList.remove("open");
    document.getElementById("dropBtn").classList.remove("open");
  }
});

function togglePlayerDrop(name,checked){
  if(checked) selected.add(name); else selected.delete(name);
  renderAll();
}

function selectAll(){DATA.players.forEach(p=>selected.add(p.name));renderAll();}
function selectNone(){selected.clear();renderAll();}

function renderRanking(){
  document.getElementById("rankBody").innerHTML=DATA.players.map(p=>{
    const med=p.platz<=3?MEDALS[p.platz-1]:"";
    const cls=[p.platz<=3?`r${p.platz}`:"",selected.has(p.name)?"sel":""].filter(Boolean).join(" ");
    return `<tr class="${cls}" onclick="togglePlayerRow('${p.name.replace(/'/g,"\\\\'")}')">
      <td class="pl">${med} ${p.platz}.</td><td class="nm">${p.name}</td>
      <td class="pts">${p.gesamt}</td><td>${p.gruppe}</td>
      <td>${p.s16}</td><td>${p.s8}</td><td>${p.vf}</td>
      <td>${p.hf}</td><td>${p.finale}</td><td>${p.wm}</td>
    </tr>`;
  }).join("");
}

function togglePlayerRow(name){
  if(selected.has(name)) selected.delete(name); else selected.add(name);
  renderAll();
}

function findBestInitTab(){
  for(const g of GROUPS){if(DATA.players[0]?.spiele.some(s=>s.gr===g&&s.kommend))return g;}
  for(const g of GROUPS){if(DATA.players[0]?.spiele.some(s=>s.gr===g&&!s.ergebnis))return g;}
  return 'A';
}
let activeTab=findBestInitTab();

function renderDetails(){
  const players=DATA.players.filter(p=>selected.has(p.name));
  const cont=document.getElementById("detailContainer");
  if(players.length===0){
    cont.innerHTML=`<div class="no-sel"><div class="no-sel-icon">👆</div>
      <div class="no-sel-txt">Spieler über das Dropdown auswählen</div></div>`;
    return;
  }
  const plHdrs=players.map(p=>{
    const med=p.platz<=3?MEDALS[p.platz-1]:`${p.platz}.`;
    return `<th class="mh-pl" style="min-width:100px"><span class="mpl-name">${med} ${p.name}</span><span class="mpl-pts">${p.gesamt} Pkt</span></th>`;
  }).join('');
  const tabsHtml=[
    ...GROUPS.map(g=>`<span class="gtab${g===activeTab?' active':''}" onclick="setTab('${g}')">Gr.${g}</span>`),
    `<span style="color:#2e4e72;padding:0 4px">│</span>`,
    ...KO_ROUNDS.map(([k,lbl])=>{
      const short={S16:'S16',S8:'S8',VF:'VF',HF:'HF',F:'Finale',WM:'WM'}[k]||k;
      return `<span class="gtab${k===activeTab?' active':''}" onclick="setTab('${k}')">${short}</span>`;
    })
  ].join('');
  let tableHtml='';
  if(GROUPS.includes(activeTab)){
    const matches=players[0].spiele.filter(s=>s.gr===activeTab);
    const rows=matches.map(s=>{
      const hasRes=s.ergebnis&&s.ergebnis!=='';
      const isDE=s.heim==='Deutschland'||s.gast==='Deutschland';
      const trCls=[isDE?'de-row':'',s.kommend?'kommend':''].filter(Boolean).join(' ');
      const plCells=players.map(p=>{
        const ps=p.spiele.find(x=>x.id===s.id)||{};
        const tipp=ps.tipp||'–';
        const pts=hasRes?ps.punkte:null;
        const cls=pts===null?'mcel-op':pts===4?'mcel-ex':pts===3?'mcel-df':pts===2?'mcel-td':'mcel-ms';
        const sub=pts!==null?`<span class="mcel-sub">${pts>0?'+'+pts:'0'}</span>`:'';
        return `<td class="mcel ${cls}">${tipp}${sub}</td>`;
      }).join('');
      return `<tr${trCls?` class="${trCls}"`:''}><td class="mi">${s.datum}${s.uhrzeit?' '+s.uhrzeit:''}</td><td class="mi-t">${s.heim} <span style="color:#3a5f84">vs</span> ${s.gast}</td><td class="mi-r">${hasRes?s.ergebnis:s.kommend?'▶':'–'}</td>${plCells}</tr>`;
    }).join('');
    const subs=players.map(p=>{
      const pts=p.spiele.filter(s=>s.gr===activeTab&&s.punkte!==null&&s.punkte!==undefined).reduce((a,s)=>a+(s.punkte||0),0);
      const n=p.spiele.filter(s=>s.gr===activeTab&&s.ergebnis).length;
      return `<td><b>${pts}</b><span style="opacity:.5;font-size:.73rem;margin-left:4px">(${n}/6)</span></td>`;
    }).join('');
    tableHtml=`<div class="mtx-wrap"><table class="mtx mtx-grp"><thead><tr><th class="mh-l">Datum</th><th class="mh-l">Spiel</th><th>Erg.</th>${plHdrs}</tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="3" class="sub-lbl">Gruppe ${activeTab} gesamt:</td>${subs}</tr></tfoot></table></div>`;
  } else {
    const actual=DATA.ko[activeTab]||[];
    const norm=actual.map(t=>(t||'').toLowerCase());
    if(activeTab==='WM'){
      const plCells=players.map(p=>{
        const tip=p.ko_tipps['WM']||'–';
        const hit=norm.length>0&&tip.toLowerCase()===norm[0];
        const miss=norm.length>0&&!hit;
        return `<td class="mcel ${hit?'mcel-ex':miss?'mcel-ms':'mcel-op'}">${tip}</td>`;
      }).join('');
      tableHtml=`<div class="mtx-wrap"><table class="mtx mtx-ko"><thead><tr><th class="mh-l">Weltmeister</th>${plHdrs}</tr></thead><tbody><tr><td class="mi-t">${actual[0]||'?'}</td>${plCells}</tr></tbody></table></div>`;
    } else {
      const tips=players.map(p=>Array.isArray(p.ko_tipps[activeTab])?p.ko_tipps[activeTab]:[]);
      const maxLen=Math.max(...tips.map(t=>t.length),0);
      const rows=maxLen===0
        ?`<tr><td colspan="${players.length+2}" style="text-align:center;padding:24px;color:#546e7a">Runde noch nicht begonnen</td></tr>`
        :Array.from({length:maxLen},(_,i)=>{
          const plCells=players.map((p,pi)=>{
            const tip=tips[pi][i]||'–';
            const hit=norm.length>i&&tip.toLowerCase()===norm[i];
            const miss=norm.length>i&&!hit;
            return `<td class="mcel ${hit?'mcel-ex':miss?'mcel-ms':'mcel-op'}">${tip}</td>`;
          }).join('');
          return `<tr><td class="mi" style="text-align:center;width:32px">${i+1}.</td><td class="mi-t">${actual[i]||'–'}</td>${plCells}</tr>`;
        }).join('');
      tableHtml=`<div class="mtx-wrap"><table class="mtx mtx-ko"><thead><tr><th style="width:32px">#</th><th class="mh-l">Weitergekommen</th>${plHdrs}</tr></thead><tbody>${rows}</tbody></table></div>`;
    }
  }
  cont.innerHTML=`<div class="gt-bar">${tabsHtml}</div>${tableHtml}`;
}

function setTab(tab){activeTab=tab;renderDetails();}

function renderAll(){renderRanking();renderDropdown();renderDetails();}
renderAll();
"""

    # ── HTML-Gerüst (f-string nur für dynamische Werte) ───
    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>WM 2026 Tippspiel</title>
<style>{CSS}</style>
</head>
<body>
<div class="hdr">
  <img class="hdr-logo" src="https://upload.wikimedia.org/wikipedia/en/thumb/1/17/2026_FIFA_World_Cup_emblem.svg/120px-2026_FIFA_World_Cup_emblem.svg.png" alt="WM 2026 Logo">
  <div><div class="hdr-title">WM 2026 Tippspiel</div>
  <div class="hdr-sub">Stand: {ts} · {filled}/72 Gruppenspiele</div></div>
</div>
</div>
<div class="stats">
  <div class="stat-box"><div class="stat-val">{n}</div><div class="stat-lbl">Teilnehmer</div></div>
  <div class="stat-box"><div class="stat-val">{leader}</div><div class="stat-lbl">🏆 Führend</div></div>
  <div class="stat-box"><div class="stat-val">{lpts} Pkt</div><div class="stat-lbl">Höchstpunktzahl</div></div>
  <div class="stat-box"><div class="stat-val">{filled}/72</div><div class="stat-lbl">Spiele ausgewertet</div></div>
</div>
<div class="filter-bar">
  <span class="filter-lbl">Detailansicht:</span>
  <div class="drop-wrap" id="dropWrap">
    <button class="drop-btn" id="dropBtn" onclick="toggleDropdown(event)">Spieler auswählen &#9662;</button>
    <div class="drop-panel" id="dropPanel">
      <div class="drop-hdr">
        <button class="drop-hdr-btn" onclick="selectAll()">Alle ausw.</button>
        <button class="drop-hdr-btn" onclick="selectNone()">Keine</button>
      </div>
      <div class="drop-opts" id="dropList"></div>
    </div>
  </div>
  <span class="hint">&#183; Zeile anklicken f&#252;r Details / Vergleich</span>
</div>
<div class="section">
  <div class="sec-title">Gesamtrangliste</div>
  <div class="wrap">
    <table class="rtbl">
      <thead><tr>
        <th>Platz</th><th>Name</th><th>Gesamt</th><th>Gruppe</th>
        <th>S16</th><th>AF</th><th>VF</th><th>HF</th><th>Finale</th><th>WM</th>
      </tr></thead>
      <tbody id="rankBody"></tbody>
    </table>
  </div>
</div>
<div class="section">
  <div class="sec-title">Detailauswertung</div>
  <div id="detailContainer"></div>
</div>
<div class="foot">Generiert am {ts} · WM 2026 Tippspiel</div>
<script>
const DATA={data_json};
{JS}
</script>
</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 3:
        print("Verwendung: python auswertung.py <tipps_ordner_oder_datei> <ergebnisse.xlsx>")
        print("Beispiel:   python auswertung.py Tipps/  Ergebnisse.xlsx")
        sys.exit(1)

    tipps_path      = Path(sys.argv[1])
    ergebnisse_path = Path(sys.argv[2])

    print(f"Lese Tipps aus:      {tipps_path}")
    print(f"Lese Ergebnisse aus: {ergebnisse_path}")

    if tipps_path.is_dir() or str(tipps_path).lower().endswith(".json"):
        df = read_json_tipps(tipps_path)
    else:
        df = read_responses(tipps_path)

    gs_results, ko_results = read_ergebnisse(ergebnisse_path)

    print(f"Teilnehmer gefunden: {len(df)}")
    print(f"Gruppenspiel-Ergebnisse eingetragen: {sum(1 for v in gs_results.values() if v)}/72")

    rows = auswerten(df, gs_results, ko_results)

    out_path = Path("Rangliste.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_rangliste(wb, rows)
    write_details(wb, rows, gs_results)
    write_ko_details(wb, rows, ko_results)

    wb.save(out_path)

    html_path = Path("Rangliste.html")
    write_html_rangliste(rows, gs_results, ko_results, html_path)

    print(f"\n✓ Excel gespeichert:  {out_path}")
    print(f"✓ HTML gespeichert:   {html_path}")
    print("\nAktuelle Rangliste:")
    print(f"{'Platz':>5}  {'Name':<22}  {'Punkte':>6}")
    print("-" * 40)
    for r in rows:
        print(f"{r['Platz']:>5}  {r['Name']:<22}  {r['Gesamt']:>6}")

if __name__ == "__main__":
    main()
