# -*- coding: utf-8 -*-
"""Carnet de soudure pipeline — professional QC tracking workbook (VT / RT / repairs / RT programme).

Usage: python welding/build_carnet.py <source_carnet.xlsx> <output.xlsx>
The source is the legacy "CARNET DE SOUDURE" workbook; its joints are imported as-is.
"""
import sys
from datetime import date, datetime

import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC, OUT = sys.argv[1], sys.argv[2]

# ------------------------------------------------------------------ styles
NAVY, BLUE, TEAL, GOLD, GREY = "1F3864", "2E75B6", "0F766E", "C9A227", "F2F2F2"
GREEN_BG, GREEN_FG = "C6EFCE", "006100"
RED_BG, RED_FG = "FFC7CE", "9C0006"
ORANGE_BG, ORANGE_FG = "FFEB9C", "9C5700"
PURPLE_BG, PURPLE_FG = "E4DFEC", "5B2C83"
BLUE_BG, BLUE_FG = "DDEBF7", "1F4E78"
FONT = "Arial"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
PCT = '0.0%;-0.0%;"-"'
DATE = "dd/mm/yyyy"


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


INPUT_FONT = font(color="0000FF")
INPUT_FILL = fill("FFF2CC")


def style(cell, fmt=None, bold=False, color="000000", bg=None, align=CENTER, size=10):
    cell.font = font(size, bold, color)
    cell.alignment = align
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    if bg:
        cell.fill = fill(bg)


def banner(ws, rng, text, size=16, bg=NAVY, color="FFFFFF", height=34):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.font = font(size, True, color)
    c.fill = fill(bg)
    c.alignment = CENTER
    ws.row_dimensions[c.row].height = height


def head(ws, row, col, text, bg=NAVY):
    c = ws.cell(row, col, text)
    style(c, bold=True, color="FFFFFF", bg=bg)
    return c


def section(ws, rng, text):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.font = font(12, True, NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[c.row].height = 24
    first, last = rng.split(":")
    r = c.row
    for col in range(c.column, ws[last].column + 1):
        ws.cell(r, col).border = Border(bottom=Side("medium", GOLD))


# ------------------------------------------------------------------ constants
ST_TO_WELD = "À SOUDER"
ST_VT_WAIT = "VT EN ATTENTE"
ST_VT_REJ = "VT REJETÉ – À REPRENDRE"
ST_TO_RT = "À RADIOGRAPHIER"
ST_ACC = "ACCEPTÉ"
ST_REP = "À RÉPARER"
ST_CO = "COUPE – À RESOUDER"
ST_REP_DONE = "RÉPARÉ – À RADIOGRAPHIER"
ST_NX = "NX – RT À REFAIRE"
STATUSES = [ST_ACC, ST_TO_RT, ST_REP_DONE, ST_NX, ST_REP, ST_CO, ST_VT_REJ, ST_VT_WAIT, ST_TO_WELD]
NEEDS_RT = [ST_TO_RT, ST_REP_DONE, ST_NX]
NEEDS_REP = [ST_REP, ST_CO, ST_VT_REJ]

LINES = [  # name, description, diameter, estimated joints
    ("ISB-302", "Puits ISB-302 → réseau de collecte", '6"', 200),
    ("ISB-303", "Puits ISB-303 → réseau de collecte", '6"', 200),
    ("ISB-304", "Puits ISB-304 → réseau de collecte", '6"', 200),
    ("TFT 716 → TL7-MMFW2", "Ligne TFT 716 vers TL7-MMFW2 (données importées)", '6"', 204),
]
MAX_LINES = 8
WELDERS = [  # WR, name, qualifications, status
    ("WR-01", "BRAHIMI Boumediene", "WQT-001 / WQT-002", "Qualifié"),
    ("WR-02", "MEHANGUEF Med Amine", "WQT-003 / 004 / 006 / 014", "Qualifié"),
    ("WR-03", "BEN KREDDA Mohammed Amine", "WQT-005", "Qualifié"),
    ("WR-04", "DEHKEL Ammar", "—", "Non qualifié"),
    ("WR-05", "KHOCHNI Sofiane", "WQT-007 / WQT-015", "Qualifié"),
    ("WR-06", "ACHOUR Walid", "WQT-008 / WQT-016", "Qualifié"),
    ("WR-07", "", "", ""),
    ("WR-08", "", "", ""),
    ("WR-09", "LALA Lazhar", "WQT-011 / WQT-012", "Qualifié"),
    ("WR-10", "BOUTAIBA Nour Eddine", "WQT-009 / 010 / 013", "Qualifié"),
    ("WR-11", "HAMOUCHE Zahir", "—", "Non qualifié"),
    ("WR-12", "", "", ""), ("WR-13", "", "", ""), ("WR-14", "", "", ""), ("WR-15", "", "", ""),
]

# ------------------------------------------------------------------ read legacy data
src = openpyxl.load_workbook(SRC, data_only=True)["CARNET DE SOUDURE"]


def v(r, col):
    x = src[f"{col}{r}"].value
    if isinstance(x, str):
        x = x.strip()
        return x or None
    return x


def res(x):
    if x is None:
        return None
    x = str(x).strip().upper()
    return {"NR": "R", "ACCEPTE": "A"}.get(x, x)


joints = []
for r in range(8, src.max_row + 1):
    if not v(r, "E"):
        continue
    line = v(r, "B")
    if line and "TFT 716" in line:
        line = "TFT 716 → TL7-MMFW2"
    rt1 = res(v(r, "AB"))
    rep_wr = v(r, "AD")
    joints.append({
        "line": line, "troncon": v(r, "C"), "pk": v(r, "D"), "joint": v(r, "E"), "zone": v(r, "L"),
        "d": v(r, "M"), "ep": v(r, "N"), "nuance": v(r, "O"), "wps": v(r, "P"),
        "t1": v(r, "F"), "t2": v(r, "I"),
        "wr1": v(r, "Q"), "wr2": v(r, "R"), "wr3": v(r, "S"), "dsoud": v(r, "T"),
        "vtd": v(r, "V"), "vtpv": v(r, "W"), "vt": res(v(r, "X")),
        "rt1d": v(r, "Z"), "rt1pv": v(r, "AA"), "rt1": rt1,
        "r1wr": rep_wr, "r1vt": res(v(r, "AG")), "r1d": v(r, "AH"), "r1pv": v(r, "AI"), "r1": res(v(r, "AJ")),
        "rem": v(r, "AN"),
    })
# pre-list the planned joints of each ISB line (numbering JN°001..N) so the field team only fills results
for name, _desc, diam, est in LINES:
    if name.startswith("ISB") and est:
        for k in range(1, est + 1):
            joints.append({"line": name, "joint": f"JN°{k:03d}", "d": diam})
LAST_DATE = max(j["rt1d"] for j in joints if isinstance(j.get("rt1d"), datetime))

wb = openpyxl.Workbook()

# ================================================================== PARAMÈTRES
pa = wb.active
pa.title = "PARAMÈTRES"
pa.sheet_properties.tabColor = "7F7F7F"
pa.sheet_view.showGridLines = False
banner(pa, "B2:H2", "PARAMÈTRES DU PROJET")
info = [("Projet", "Extension du réseau de collecte de GTFT (11 nouveaux puits)"),
        ("Affaire N°", "8546"), ("Direction", "Direction Régionale Hassi R'Mel"),
        ("Entreprise", "ENGTP"), ("Code de référence", "API 1104")]
for i, (k, val) in enumerate(info):
    r = 4 + i
    style(pa.cell(r, 2, k), bold=True, bg=GREY, align=LEFT)
    pa.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    c = pa.cell(r, 3, val)
    style(c, color="0000FF", bg="FFF2CC", align=LEFT)
PROJ, AFF = "PARAMÈTRES!$C$4", "PARAMÈTRES!$C$5"

section(pa, "B10:F10", "LIGNES (une ligne par puits) — cellules jaunes à renseigner")
for i, h in enumerate(["Ligne", "Description", "Ø", "Joints estimés"]):
    head(pa, 11, 2 + i, h, BLUE)
L_FIRST = 12
L_LAST = L_FIRST + MAX_LINES - 1
for i in range(MAX_LINES):
    r = L_FIRST + i
    row = LINES[i] if i < len(LINES) else ("", "", "", None)
    for j, val in enumerate(row):
        c = pa.cell(r, 2 + j, val if val != "" else None)
        style(c, color="0000FF", bg="FFF2CC" if j == 3 else None, align=LEFT if j == 1 else CENTER)
    pa.cell(r, 5).number_format = "#,##0"
pa.cell(L_FIRST, 5).comment = Comment("Nombre total de joints prévus sur la ligne (sert au calcul de l'avancement).", "QC")
LINE_NAMES = f"PARAMÈTRES!$B${L_FIRST}:$B${L_LAST}"
LINE_EST = f"PARAMÈTRES!$E${L_FIRST}:$E${L_LAST}"

section(pa, "B22:F22", "CODES DE RÉSULTAT")
for i, h in enumerate(["Code", "Contrôle", "Signification", "Action"]):
    head(pa, 23, 2 + i, h, BLUE)
CODES = [("A", "VT / RT", "Accepté", "Aucune — joint conforme"),
         ("R", "VT / RT", "Rejeté — défaut à réparer", "Réparer (WPS REP) → VT → RT de contrôle"),
         ("NX", "RT", "Reprise de film (film à refaire)", "Refaire une radiographie (sans réparation)"),
         ("CO", "RT", "À couper", "Couper le joint, resouder → VT → RT")]
for i, row in enumerate(CODES):
    for j, val in enumerate(row):
        style(pa.cell(24 + i, 2 + j, val), bold=(j == 0), align=LEFT if j >= 2 else CENTER)
pa.cell(24, 2).fill = fill(GREEN_BG); pa.cell(25, 2).fill = fill(RED_BG)
pa.cell(26, 2).fill = fill(ORANGE_BG); pa.cell(27, 2).fill = fill(PURPLE_BG)

section(pa, "B30:F30", "LISTES (menus déroulants)")
head(pa, 31, 2, "Filtre ligne", BLUE)
head(pa, 31, 3, "VT", BLUE)
head(pa, 31, 4, "RT", BLUE)
head(pa, 31, 5, "Statuts", BLUE)
style(pa.cell(32, 2, "TOUTES"), bold=True)
for i in range(MAX_LINES):
    c = pa.cell(33 + i, 2, f'=IF(B{L_FIRST + i}="","",B{L_FIRST + i})')
    style(c)
for i, x in enumerate(["A", "R"]):
    style(pa.cell(32 + i, 3, x))
for i, x in enumerate(["A", "R", "NX", "CO"]):
    style(pa.cell(32 + i, 4, x))
for i, x in enumerate(STATUSES):
    style(pa.cell(32 + i, 5, x), align=LEFT)
FILTER_LIST = f"PARAMÈTRES!$B$32:$B${32 + MAX_LINES}"
VT_LIST = "PARAMÈTRES!$C$32:$C$33"
RT_LIST = "PARAMÈTRES!$D$32:$D$35"
for col, w in {"A": 2, "B": 24, "C": 44, "D": 40, "E": 34, "F": 4}.items():
    pa.column_dimensions[col].width = w

# ================================================================== CARNET
cs = wb.create_sheet("CARNET DE SOUDURE", 0)
cs.sheet_properties.tabColor = BLUE
FIRST = 7
CAP = 1500
LAST = FIRST + CAP - 1
COLS = [  # header, width, kind (in=input, f=formula, h=hidden helper), group
    ("N°", 6, "f", "ID"), ("Ligne", 20, "in", "ID"), ("Tronçon", 8, "in", "ID"), ("PK", 7, "in", "ID"),
    ("N° Joint", 13, "in", "ID"), ("Zone", 6, "in", "ID"), ("Ø", 5, "in", "ID"), ("Ép. (mm)", 7, "in", "ID"),
    ("Nuance", 11, "in", "ID"), ("WPS", 12, "in", "ID"), ("Tube 1", 9, "in", "ID"), ("Tube 2", 9, "in", "ID"),
    ("WR 1ère passe", 9, "in", "SOUD"), ("WR 2ème passe", 9, "in", "SOUD"), ("WR B+F", 9, "in", "SOUD"),
    ("Date soudage", 11, "in", "SOUD"),
    ("VT Date", 11, "in", "VT"), ("VT PV N°", 9, "in", "VT"), ("VT Rés.", 6, "in", "VT"),
    ("RT1 Date", 11, "in", "RT1"), ("RT1 PV N°", 10, "in", "RT1"), ("RT1 Rés.", 6, "in", "RT1"),
    ("R1 WR", 8, "in", "R1"), ("R1 VT Rés.", 7, "in", "R1"), ("R1 RT Date", 11, "in", "R1"),
    ("R1 RT PV N°", 10, "in", "R1"), ("R1 RT Rés.", 7, "in", "R1"),
    ("R2 WR", 8, "in", "R2"), ("R2 VT Rés.", 7, "in", "R2"), ("R2 RT Date", 11, "in", "R2"),
    ("R2 RT PV N°", 10, "in", "R2"), ("R2 RT Rés.", 7, "in", "R2"),
    ("RT programmé le", 11, "in", "SUIVI"), ("Remarques", 22, "in", "SUIVI"),
    ("Nb RT", 6, "f", "SUIVI"), ("Dernier RT", 8, "f", "SUIVI"), ("STATUT", 24, "f", "SUIVI"),
    ("Attente (j)", 8, "f", "SUIVI"),
    ("hRT", 4, "h", "H"), ("hRTn", 4, "h", "H"), ("hREP", 4, "h", "H"), ("hREPn", 4, "h", "H"),
]
C = {h: CL(i + 1) for i, (h, *_rest) in enumerate(COLS)}
LASTCOL = CL(len(COLS))
GROUPS = {"ID": ("IDENTIFICATION DU JOINT", NAVY), "SOUD": ("SOUDAGE", TEAL), "VT": ("CONTRÔLE VISUEL (VT)", "548235"),
          "RT1": ("RADIOGRAPHIE RT-01", BLUE), "R1": ("RÉPARATION 1 → VT → RT", "C55A11"),
          "R2": ("RÉPARATION 2 → VT → RT", "843C0C"), "SUIVI": ("PROGRAMME & SUIVI", "7030A0"), "H": ("", "808080")}

banner(cs, f"A1:{C['Attente (j)']}1", "CARNET DE SOUDURE — PIPELINE", 18)
cs.merge_cells(f"A2:{C['Attente (j)']}2")
cs["A2"] = f'={PROJ}&"   •   Affaire N° "&{AFF}&"   •   Saisir uniquement les colonnes en bleu — les colonnes N°, Nb RT, Dernier RT, STATUT et Attente sont automatiques"'
cs["A2"].font = font(9, False, "595959", True)
cs["A2"].alignment = CENTER
# quick counters row 3
cs["B3"] = "Joints :"; cs["C3"] = f'=COUNTA({C["N° Joint"]}{FIRST}:{C["N° Joint"]}{LAST})'
cs["E3"] = "Soudés :"; cs["F3"] = f'=COUNT({C["Date soudage"]}{FIRST}:{C["Date soudage"]}{LAST})'
cs["I3"] = "Acceptés :"; cs["J3"] = f'=COUNTIF({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST},"{ST_ACC}")'
cs["M3"] = "À radiographier :"; cs["P3"] = f'=SUMPRODUCT(({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_TO_RT}")+({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_REP_DONE}")+({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_NX}"))'
cs["S3"] = "À réparer :"; cs["V3"] = f'=SUMPRODUCT(({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_REP}")+({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_CO}")+({C["STATUT"]}{FIRST}:{C["STATUT"]}{LAST}="{ST_VT_REJ}"))'
for a, b in (("B3", "C3"), ("E3", "F3"), ("I3", "J3"), ("M3", "P3"), ("S3", "V3")):
    cs[a].font = font(10, True, NAVY); cs[a].alignment = Alignment(horizontal="right")
    cs[b].font = font(12, True, "C00000"); cs[b].alignment = CENTER
cs.merge_cells("M3:O3"); cs.merge_cells("S3:U3")
cs.row_dimensions[3].height = 22

# group header row 5, column header row 6
start = 1
while start <= len(COLS):
    g = COLS[start - 1][3]
    end = start
    while end < len(COLS) and COLS[end][3] == g:
        end += 1
    label, color = GROUPS[g]
    if end > start:
        cs.merge_cells(start_row=5, start_column=start, end_row=5, end_column=end)
    c = cs.cell(5, start, label)
    c.font = font(10, True, "FFFFFF"); c.alignment = CENTER
    for col in range(start, end + 1):
        cs.cell(5, col).fill = fill(color)
        cs.cell(5, col).border = BORDER
    start = end + 1
cs.row_dimensions[5].height = 22
for i, (h, w, kind, g) in enumerate(COLS, 1):
    c = cs.cell(6, i, h)
    c.font = font(9, True, "FFFFFF"); c.alignment = CENTER
    c.fill = fill(GROUPS[g][1]); c.border = BORDER
    cs.column_dimensions[CL(i)].width = w
    if kind == "h":
        cs.column_dimensions[CL(i)].hidden = True
cs.row_dimensions[6].height = 36

FILT_RT = "'PROGRAMME RT'!$D$5"
REF_DATE = "'PROGRAMME RT'!$C$4"
FILT_REP = "'RÉPARATIONS'!$D$5"
KEYS = ["line", "troncon", "pk", "joint", "zone", "d", "ep", "nuance", "wps", "t1", "t2", "wr1", "wr2", "wr3",
        "dsoud", "vtd", "vtpv", "vt", "rt1d", "rt1pv", "rt1", "r1wr", "r1vt", "r1d", "r1pv", "r1",
        None, None, None, None, None, None, "rem"]
KEY_COLS = [C[h] for h, *_ in COLS[1:34]]  # Ligne .. Remarques (33 cols)
KEY_COLS = KEY_COLS[:32] + [C["Remarques"]]
key_map = dict(zip(KEYS[:26], [C[h] for h, *_ in COLS[1:27]]))
key_map["rem"] = C["Remarques"]
DATE_COLS = [C[h] for h in ("Date soudage", "VT Date", "RT1 Date", "R1 RT Date", "R2 RT Date", "RT programmé le")]
RES_COLS = [C[h] for h in ("VT Rés.", "RT1 Rés.", "R1 VT Rés.", "R1 RT Rés.", "R2 VT Rés.", "R2 RT Rés.")]

E, P, S = C["N° Joint"], C["Date soudage"], C["VT Rés."]
V1, V2, V3 = C["RT1 Rés."], C["R1 RT Rés."], C["R2 RT Rés."]
X1, X2 = C["R1 VT Rés."], C["R2 VT Rés."]
NB, LR, ST, AT = C["Nb RT"], C["Dernier RT"], C["STATUT"], C["Attente (j)"]
HRT, HRTN, HREP, HREPN = C["hRT"], C["hRTn"], C["hREP"], C["hREPn"]
B = C["Ligne"]

for idx in range(CAP):
    r = FIRST + idx
    if idx < len(joints):
        j = joints[idx]
        for k, col in key_map.items():
            if j.get(k) is not None:
                cs[f"{col}{r}"] = j[k]
    cs[f"A{r}"] = f'=IF({E}{r}="","",ROW()-{FIRST - 1})'
    cs[f"{NB}{r}"] = f'=IF({E}{r}="","",({V1}{r}<>"")+({V2}{r}<>"")+({V3}{r}<>""))'
    cs[f"{LR}{r}"] = f'=IF({V3}{r}<>"",{V3}{r},IF({V2}{r}<>"",{V2}{r},IF({V1}{r}<>"",{V1}{r},"")))'
    next_vt = f'IF({NB}{r}=1,{X1}{r},IF({NB}{r}=2,{X2}{r},""))'
    cs[f"{ST}{r}"] = (
        f'=IF({E}{r}="","",IF({P}{r}="","{ST_TO_WELD}",IF({S}{r}="","{ST_VT_WAIT}",'
        f'IF({S}{r}="R","{ST_VT_REJ}",IF({LR}{r}="","{ST_TO_RT}",IF({LR}{r}="A","{ST_ACC}",'
        f'IF({LR}{r}="NX","{ST_NX}",IF({next_vt}="A","{ST_REP_DONE}",IF({next_vt}="R","{ST_VT_REJ}",'
        f'IF({LR}{r}="CO","{ST_CO}","{ST_REP}"))))))))))'
    )
    active = f'OR({ST}{r}="{ST_TO_RT}",{ST}{r}="{ST_REP_DONE}",{ST}{r}="{ST_NX}",{ST}{r}="{ST_REP}",{ST}{r}="{ST_CO}",{ST}{r}="{ST_VT_REJ}")'
    cs[f"{AT}{r}"] = f'=IF(AND(ISNUMBER({P}{r}),{active}),MAX(0,{REF_DATE}-{P}{r}),"")'
    needs_rt = f'OR({ST}{r}="{ST_TO_RT}",{ST}{r}="{ST_REP_DONE}",{ST}{r}="{ST_NX}")'
    needs_rep = f'OR({ST}{r}="{ST_REP}",{ST}{r}="{ST_CO}",{ST}{r}="{ST_VT_REJ}")'
    cs[f"{HRT}{r}"] = f'=IF(AND({needs_rt},OR({FILT_RT}="TOUTES",{B}{r}={FILT_RT})),1,0)'
    cs[f"{HRTN}{r}"] = f'=IF({HRT}{r}=1,SUM({HRT}${FIRST}:{HRT}{r}),"")'
    cs[f"{HREP}{r}"] = f'=IF(AND({needs_rep},OR({FILT_REP}="TOUTES",{B}{r}={FILT_REP})),1,0)'
    cs[f"{HREPN}{r}"] = f'=IF({HREP}{r}=1,SUM({HREP}${FIRST}:{HREP}{r}),"")'
    for i, (h, w, kind, g) in enumerate(COLS, 1):
        c = cs.cell(r, i)
        c.font = font(9, kind == "f" and h == "STATUT", "0000FF" if kind == "in" else "000000")
        c.alignment = LEFT if h in ("Remarques", "Ligne") else CENTER
    for col in DATE_COLS:
        cs[f"{col}{r}"].number_format = DATE

tbl = Table(displayName="Carnet", ref=f"A6:{LASTCOL}{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
cs.add_table(tbl)
cs.freeze_panes = f"F{FIRST}"

# validations
def add_dv(ws, formula, cells, msg, kind="list"):
    d = DataValidation(type=kind, formula1=formula, allow_blank=True, showErrorMessage=True, showInputMessage=True)
    d.errorTitle, d.error, d.prompt = "Valeur non valide", msg, msg
    ws.add_data_validation(d)
    for c in cells:
        d.add(c)

rows = f"{FIRST}:{{}}{LAST}"
add_dv(cs, LINE_NAMES, [f"{B}{FIRST}:{B}{LAST}"], "Choisir la ligne dans la liste (PARAMÈTRES)")
WR_LIST = "SOUDEURS!$B$6:$B$20"
add_dv(cs, WR_LIST, [f"{C[h]}{FIRST}:{C[h]}{LAST}" for h in ("WR 1ère passe", "WR 2ème passe", "WR B+F", "R1 WR", "R2 WR")],
       "Choisir le repère soudeur (WR) — liste SOUDEURS")
add_dv(cs, VT_LIST, [f"{C[h]}{FIRST}:{C[h]}{LAST}" for h in ("VT Rés.", "R1 VT Rés.", "R2 VT Rés.")], "VT : A = Accepté, R = Rejeté")
add_dv(cs, RT_LIST, [f"{C[h]}{FIRST}:{C[h]}{LAST}" for h in ("RT1 Rés.", "R1 RT Rés.", "R2 RT Rés.")],
       "RT : A = Accepté, R = À réparer, NX = Reprise de film, CO = À couper")
dd = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                    allow_blank=True, showErrorMessage=True)
dd.error, dd.errorTitle = "Saisir une date valide (jj/mm/aaaa)", "Date"
cs.add_data_validation(dd)
for col in DATE_COLS:
    dd.add(f"{col}{FIRST}:{col}{LAST}")

# conditional formatting
def res_colors(ws, rng):
    for code, bg, fg in (("A", GREEN_BG, GREEN_FG), ("R", RED_BG, RED_FG), ("NX", ORANGE_BG, ORANGE_FG), ("CO", PURPLE_BG, PURPLE_FG)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{code}"'], fill=fill(bg),
                                                      font=Font(name=FONT, bold=True, color=fg)))

for col in RES_COLS + [LR]:
    res_colors(cs, f"{col}{FIRST}:{col}{LAST}")
STATUS_STYLE = {ST_ACC: (GREEN_BG, GREEN_FG), ST_TO_RT: (BLUE_BG, BLUE_FG), ST_REP_DONE: (BLUE_BG, BLUE_FG),
                ST_NX: (ORANGE_BG, ORANGE_FG), ST_REP: (RED_BG, RED_FG), ST_CO: (PURPLE_BG, PURPLE_FG),
                ST_VT_REJ: (RED_BG, RED_FG), ST_VT_WAIT: ("FFF2CC", "7F6000"), ST_TO_WELD: ("EDEDED", "595959")}


def status_colors(ws, rng):
    for s, (bg, fg) in STATUS_STYLE.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{s}"'], fill=fill(bg),
                                                      font=Font(name=FONT, bold=True, color=fg)))

status_colors(cs, f"{ST}{FIRST}:{ST}{LAST}")
cs.conditional_formatting.add(f"{AT}{FIRST}:{AT}{LAST}", FormulaRule(formula=[f'AND(ISNUMBER({AT}{FIRST}),{AT}{FIRST}>3)'],
                              fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
cs.conditional_formatting.add(f"A{FIRST}:{C['N° Joint']}{LAST}", FormulaRule(formula=[f'${ST}{FIRST}="{ST_ACC}"'],
                              font=Font(name=FONT, bold=True, color=GREEN_FG)))
cs.conditional_formatting.add(f"{E}{FIRST}:{E}{LAST}", FormulaRule(formula=[f'AND({E}{FIRST}<>"",COUNTIFS(${B}${FIRST}:${B}${LAST},{B}{FIRST},${E}${FIRST}:${E}${LAST},{E}{FIRST})>1)'],
                              fill=fill("FF0000"), font=Font(name=FONT, bold=True, color="FFFFFF")))
cs[f"{E}6"].comment = Comment("Un N° de joint en double sur la même ligne s'affiche en ROUGE.", "QC")
cs[f"{AT}6"].comment = Comment("Jours depuis le soudage (par rapport à la date de PROGRAMME RT) pour un joint non encore accepté. Rouge si > 3 jours.", "QC")
cs[f"{ST}6"].comment = Comment("Calculé automatiquement à partir des résultats VT / RT et des réparations.", "QC")
cs[f"{V1}6"].comment = Comment("A = Accepté · R = À réparer · NX = Reprise de film · CO = À couper", "QC")
cs.page_setup.orientation = "landscape"; cs.page_setup.paperSize = cs.PAPERSIZE_A3
cs.page_setup.fitToWidth = 1; cs.page_setup.fitToHeight = 0
cs.sheet_properties.pageSetUpPr.fitToPage = True
cs.print_title_rows = "5:6"

RNG = lambda col: f"'CARNET DE SOUDURE'!${col}${FIRST}:${col}${LAST}"

# ================================================================== SOUDEURS
so = wb.create_sheet("SOUDEURS")
so.sheet_properties.tabColor = TEAL
so.sheet_view.showGridLines = False
banner(so, "B2:K2", "SOUDEURS — QUALIFICATION & PERFORMANCE")
so.merge_cells("B3:K3")
so["B3"] = "Taux de réparation = joints rejetés (R ou CO) au RT-01 ÷ joints radiographiés (hors NX). Un soudeur est compté s'il a réalisé au moins une passe du joint."
so["B3"].font = font(9, False, "595959", True); so["B3"].alignment = CENTER
HDR = ["WR", "Nom & Prénom", "Qualifications (WQT)", "Statut", "Joints soudés", "Radiographiés RT-01",
       "Rejetés (R/CO)", "Taux de réparation", "Réparations effectuées", "Évaluation"]
for i, h in enumerate(HDR):
    head(so, 5, 2 + i, h, TEAL)
so.row_dimensions[5].height = 34
M_, N_, O_ = C["WR 1ère passe"], C["WR 2ème passe"], C["WR B+F"]
for i, (wr, name, q, stt) in enumerate(WELDERS):
    r = 6 + i
    so.cell(r, 2, wr); so.cell(r, 3, name or None); so.cell(r, 4, q or None); so.cell(r, 5, stt or None)
    anyp = f"((({RNG(M_)}=B{r})+({RNG(N_)}=B{r})+({RNG(O_)}=B{r}))>0)"
    so.cell(r, 6, f"=SUMPRODUCT(--{anyp})")
    so.cell(r, 7, f'=SUMPRODUCT({anyp}*(({RNG(V1)}="A")+({RNG(V1)}="R")+({RNG(V1)}="CO")))')
    so.cell(r, 8, f'=SUMPRODUCT({anyp}*(({RNG(V1)}="R")+({RNG(V1)}="CO")))')
    so.cell(r, 9, f"=IFERROR(H{r}/G{r},0)")
    so.cell(r, 10, f'=COUNTIF({RNG(C["R1 WR"])},B{r})+COUNTIF({RNG(C["R2 WR"])},B{r})')
    so.cell(r, 11, f'=IF(G{r}=0,"—",IF(I{r}<=0.02,"Excellent",IF(I{r}<=0.05,"Bon","À surveiller")))')
    for c in range(2, 12):
        style(so.cell(r, c), align=LEFT if c in (3, 4) else CENTER, color="0000FF" if c <= 5 else "000000", bold=(c == 2))
    so.cell(r, 9).number_format = PCT
tr = 6 + len(WELDERS)
so.cell(tr, 2, "TOTAL")
so.cell(tr, 6, f'=SUMPRODUCT(--((({RNG(M_)}<>"")+({RNG(N_)}<>"")+({RNG(O_)}<>""))>0))')
so.cell(tr, 7, f'=COUNTIF({RNG(V1)},"A")+COUNTIF({RNG(V1)},"R")+COUNTIF({RNG(V1)},"CO")')
so.cell(tr, 8, f'=COUNTIF({RNG(V1)},"R")+COUNTIF({RNG(V1)},"CO")')
so.cell(tr, 9, f"=IFERROR(H{tr}/G{tr},0)")
so.cell(tr, 10, f"=SUM(J6:J{tr - 1})")
for c in range(2, 12):
    style(so.cell(tr, c), bold=True, color="FFFFFF", bg=TEAL)
so.cell(tr, 9).number_format = PCT
for col, w in zip("ABCDEFGHIJK", [2, 9, 28, 26, 13, 12, 14, 12, 13, 14, 14]):
    so.column_dimensions[col].width = w
so.conditional_formatting.add(f"I6:I{tr - 1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.2, color="F4B183"))
so.conditional_formatting.add(f"K6:K{tr - 1}", CellIsRule(operator="equal", formula=['"Excellent"'], fill=fill(GREEN_BG), font=Font(name=FONT, bold=True, color=GREEN_FG)))
so.conditional_formatting.add(f"K6:K{tr - 1}", CellIsRule(operator="equal", formula=['"À surveiller"'], fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
so.conditional_formatting.add(f"E6:E{tr - 1}", CellIsRule(operator="equal", formula=['"Non qualifié"'], fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
so.freeze_panes = "C6"

# ================================================================== list sheets (programme RT / repairs)
def list_sheet(name, title_txt, color, flag_n, rows_n, motif_hdr, extra_note):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color
    ws.sheet_view.showGridLines = False
    banner(ws, "B2:M2", title_txt, bg=color)
    ws["B4"] = "Date :"; ws["C4"] = "=TODAY()" if name == "PROGRAMME RT" else "='PROGRAMME RT'!C4"
    ws["B5"] = "Ligne :"; ws["D5"] = "TOUTES"
    ws.merge_cells("D5:E5")
    for a in ("B4", "B5"):
        ws[a].font = font(11, True, NAVY)
    ws["C4"].number_format = DATE
    for a in ("C4", "D5"):
        ws[a].font = Font(name=FONT, size=11, bold=True, color="0000FF")
        ws[a].fill = INPUT_FILL
        ws[a].alignment = CENTER
        ws[a].border = BORDER
    ws["C4"].comment = Comment("Date du jour automatique. Pour voir un autre jour, tapez une date (ex. 17/09/2025).", "QC")
    add_dv(ws, FILTER_LIST, ["D5"], "Choisir une ligne ou TOUTES")
    ws["G4"] = "Nombre de joints :"
    ws["G4"].font = font(11, True, NAVY)
    ws["I4"] = f"=SUM('CARNET DE SOUDURE'!${flag_n}${FIRST}:${flag_n}${LAST})"
    ws["I4"].font = font(16, True, "C00000"); ws["I4"].alignment = CENTER
    ws.merge_cells("G5:L5")
    ws["G5"] = extra_note
    ws["G5"].font = font(9, False, "595959", True)
    return ws


PROG_ROWS = 150
pg = list_sheet("PROGRAMME RT", "PROGRAMME RADIOGRAPHIE (RT) DU JOUR", BLUE, HRT, PROG_ROWS, "Motif",
                "Liste automatique des joints soudés et acceptés au VT, pas encore radiographiés, réparés, ou NX à refaire.")
rp = list_sheet("RÉPARATIONS", "SUIVI DES RÉPARATIONS", "C00000", HREP, PROG_ROWS, "Motif",
                "Joints rejetés au RT (R), à couper (CO) ou rejetés au VT : réparer puis refaire VT + RT.")


def fill_list(ws, rank_col, headers, getters, fmts, start=12):
    for i, h in enumerate(headers):
        head(ws, start - 1, 2 + i, h, NAVY)
    ws.row_dimensions[start - 1].height = 30
    for k in range(1, PROG_ROWS + 1):
        r = start + k - 1
        m = f"MATCH({k},'CARNET DE SOUDURE'!${rank_col}${FIRST}:${rank_col}${LAST},0)"
        ws.cell(r, 2, f'=IF({k}>$I$4,"",{k})')
        for i, g in enumerate(getters):
            if g is None:
                c = ws.cell(r, 3 + i)  # manual column
                style(c, color="0000FF")
                continue
            if g.startswith("="):
                expr = g[1:].replace("{m}", m).replace("{r}", str(r))
            else:
                expr = f"INDEX({RNG(g)},{m})"
            ws.cell(r, 3 + i, f'=IF($B{r}="","",{expr})')
        for i in range(len(headers)):
            c = ws.cell(r, 2 + i)
            if getters[i - 1] is None and i > 0:
                continue
            style(c, fmts[i])
        if k % 2 == 0:
            for i in range(len(headers)):
                if i == 0 or getters[i - 1] is not None:
                    ws.cell(r, 2 + i).fill = fill("F7F9FC")
    return start, start + PROG_ROWS - 1


WR_JOIN = f'=INDEX({RNG(M_)},{{m}})&" / "&INDEX({RNG(N_)},{{m}})&" / "&INDEX({RNG(O_)},{{m}})'
p_first, p_last = fill_list(
    pg, HRTN,
    ["#", "Ligne", "N° Joint", "PK", "Soudeurs (WR)", "Date soudage", "Motif", "RT réalisés", "Attente (j)", "Fait ✓", "PV RT / Observations"],
    [B, E, C["PK"], WR_JOIN, P, ST, NB, AT, None, None],
    [None, None, None, None, None, DATE, None, "0", "0", None, None])
r_first, r_last = fill_list(
    rp, HREPN,
    ["#", "Ligne", "N° Joint", "Soudeurs (WR)", "Dernier RT", "Date dernier RT", "PV", "Statut", "Attente (j)", "Soudeur réparation", "Observations"],
    [B, E, WR_JOIN,
     LR,
     f'=IF(INDEX({RNG(V3)},{{m}})<>"",INDEX({RNG(C["R2 RT Date"])},{{m}}),IF(INDEX({RNG(V2)},{{m}})<>"",INDEX({RNG(C["R1 RT Date"])},{{m}}),INDEX({RNG(C["RT1 Date"])},{{m}})))',
     f'=IF(INDEX({RNG(V3)},{{m}})<>"",INDEX({RNG(C["R2 RT PV N°"])},{{m}}),IF(INDEX({RNG(V2)},{{m}})<>"",INDEX({RNG(C["R1 RT PV N°"])},{{m}}),INDEX({RNG(C["RT1 PV N°"])},{{m}})))&""',
     ST, AT, None, None],
    [None, None, None, None, None, DATE, None, None, "0", None, None])
for ws, (a, b), stcol in ((pg, (p_first, p_last), "H"), (rp, (r_first, r_last), "I")):
    status_colors(ws, f"{stcol}{a}:{stcol}{b}")
    atc = "J"
    ws.conditional_formatting.add(f"{atc}{a}:{atc}{b}", FormulaRule(formula=[f'AND(ISNUMBER({atc}{a}),{atc}{a}>3)'],
                                  fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
    for col, w in zip("ABCDEFGHIJKL", [2, 5, 20, 12, 7, 20, 12, 24, 10, 10, 12, 26]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"C{a}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{a - 1}:{a - 1}"
res_colors(rp, f"F{r_first}:F{r_last}")
rp.column_dimensions["F"].width = 10; rp.column_dimensions["G"].width = 12; rp.column_dimensions["E"].width = 20
rp.column_dimensions["H"].width = 11; rp.column_dimensions["I"].width = 24; rp.column_dimensions["J"].width = 10
rp.column_dimensions["K"].width = 14; rp["D5"].value = "TOUTES"
pg.column_dimensions["F"].width = 20

# daily report block on programme sheet (rows 7-9)
pg.merge_cells("B7:M7")
pg["B7"] = '="RAPPORT DU "&TEXT($C$4,"dd/mm/yyyy")'
pg["B7"].font = font(11, True, "FFFFFF"); pg["B7"].fill = fill(NAVY); pg["B7"].alignment = CENTER
day_items = [
    ("Joints soudés", f"COUNTIF({RNG(P)},$C$4)"),
    ("VT réalisés", f"COUNTIF({RNG(C['VT Date'])},$C$4)"),
    ("RT réalisés", f"COUNTIF({RNG(C['RT1 Date'])},$C$4)+COUNTIF({RNG(C['R1 RT Date'])},$C$4)+COUNTIF({RNG(C['R2 RT Date'])},$C$4)"),
    ("RT acceptés", f'COUNTIFS({RNG(C["RT1 Date"])},$C$4,{RNG(V1)},"A")+COUNTIFS({RNG(C["R1 RT Date"])},$C$4,{RNG(V2)},"A")+COUNTIFS({RNG(C["R2 RT Date"])},$C$4,{RNG(V3)},"A")'),
    ("RT à réparer (R/CO)", f'COUNTIFS({RNG(C["RT1 Date"])},$C$4,{RNG(V1)},"R")+COUNTIFS({RNG(C["RT1 Date"])},$C$4,{RNG(V1)},"CO")+COUNTIFS({RNG(C["R1 RT Date"])},$C$4,{RNG(V2)},"R")+COUNTIFS({RNG(C["R1 RT Date"])},$C$4,{RNG(V2)},"CO")+COUNTIFS({RNG(C["R2 RT Date"])},$C$4,{RNG(V3)},"R")+COUNTIFS({RNG(C["R2 RT Date"])},$C$4,{RNG(V3)},"CO")'),
    ("RT NX", f'COUNTIFS({RNG(C["RT1 Date"])},$C$4,{RNG(V1)},"NX")+COUNTIFS({RNG(C["R1 RT Date"])},$C$4,{RNG(V2)},"NX")+COUNTIFS({RNG(C["R2 RT Date"])},$C$4,{RNG(V3)},"NX")'),
]
col = 2
for label, f in day_items:
    pg.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col + 1) if col < 12 else None
    c = pg.cell(8, col, label); style(c, bold=True, color=NAVY, bg="DDEBF7")
    pg.cell(8, col + 1).border = BORDER; pg.cell(8, col + 1).fill = fill("DDEBF7")
    pg.merge_cells(start_row=9, start_column=col, end_row=9, end_column=col + 1)
    c = pg.cell(9, col, f"={f}"); style(c, bold=True, size=14)
    pg.cell(9, col + 1).border = BORDER
    col += 2
pg.row_dimensions[9].height = 26

# ================================================================== DASHBOARD
db = wb.create_sheet("TABLEAU DE BORD", 0)
db.sheet_properties.tabColor = GOLD
db.sheet_view.showGridLines = False
db.column_dimensions["A"].width = 2
for i in range(2, 14):
    db.column_dimensions[CL(i)].width = 13.5
db.column_dimensions["B"].width = 22
banner(db, "B2:M2", "TABLEAU DE BORD — SOUDAGE & CONTRÔLE QUALITÉ PIPELINE", 18, height=40)
db.merge_cells("B3:M3")
db["B3"] = f'={PROJ}&"   •   Affaire N° "&{AFF}&"   •   Mis à jour le "&TEXT(TODAY(),"dd/mm/yyyy")'
db["B3"].font = font(10, False, "595959", True); db["B3"].alignment = CENTER
db["B5"] = "Ligne / Puits :"
db["B5"].font = font(12, True, NAVY); db["B5"].alignment = Alignment(horizontal="right", vertical="center")
db.merge_cells("C5:D5")
db["C5"] = "TOUTES"
for cc in ("C5", "D5"):
    db[cc].fill = INPUT_FILL
    db[cc].border = Border(left=Side("medium", GOLD), right=Side("medium", GOLD), top=Side("medium", GOLD), bottom=Side("medium", GOLD))
db["C5"].font = Font(name=FONT, size=12, bold=True, color="0000FF"); db["C5"].alignment = CENTER
add_dv(db, FILTER_LIST, ["C5"], "Choisir une ligne ou TOUTES")
db.merge_cells("E5:J5")
db["E5"] = "◄ choisir une ligne pour filtrer tous les indicateurs"
db["E5"].font = font(9, False, "7F7F7F", True)
db["O5"] = '=IF(C5="TOUTES","*",C5)'   # criteria helper
db.column_dimensions["O"].hidden = True
CR = "$O$5"
LN = RNG(B)
cnt = lambda *crit: "COUNTIFS(" + ",".join([LN, CR] + list(crit)) + ")"
st_cnt = lambda s: cnt(RNG(ST), f'"{s}"')
EST = f'IF($C$5="TOUTES",SUM({LINE_EST}),SUMIF({LINE_NAMES},$C$5,{LINE_EST}))'
WELDED = cnt(RNG(P), '"<>"')
RT_DONE = cnt(RNG(V1), '"<>"')
REJ = f'({cnt(RNG(V1), chr(34) + "R" + chr(34))}+{cnt(RNG(V1), chr(34) + "CO" + chr(34))})'
RT_INT = f'({cnt(RNG(V1), chr(34) + "A" + chr(34))}+{REJ})'

cards1 = [
    ("JOINTS SOUDÉS", f"={WELDED}", "#,##0", NAVY, f'="sur "&TEXT({EST},"#,##0")&" estimés"'),
    ("AVANCEMENT SOUDAGE", f"=IFERROR({WELDED}/{EST},0)", PCT, TEAL, f'="Reste : "&TEXT(MAX(0,{EST}-{WELDED}),"#,##0")&" joints"'),
    ("JOINTS RADIOGRAPHIÉS", f"={RT_DONE}", "#,##0", BLUE, f'=TEXT(IFERROR({RT_DONE}/{WELDED},0),"0.0%")&" des joints soudés"'),
    ("JOINTS ACCEPTÉS", f"={st_cnt(ST_ACC)}", "#,##0", "548235", f'=TEXT(IFERROR({st_cnt(ST_ACC)}/{WELDED},0),"0.0%")&" des joints soudés"'),
    ("EN ATTENTE RT", f"={st_cnt(ST_TO_RT)}+{st_cnt(ST_REP_DONE)}+{st_cnt(ST_NX)}", "#,##0", "7030A0", '="voir PROGRAMME RT"'),
    ("TAUX DE RÉPARATION", f"=IFERROR({REJ}/{RT_INT},0)", PCT, "C00000", f'=TEXT({REJ},"0")&" rejet(s) au RT-01"'),
]
cards2 = [
    ("À RÉPARER (R)", f"={st_cnt(ST_REP)}", RED_BG, RED_FG),
    ("COUPE (CO)", f"={st_cnt(ST_CO)}", PURPLE_BG, PURPLE_FG),
    ("NX – RT À REFAIRE", f"={st_cnt(ST_NX)}", ORANGE_BG, ORANGE_FG),
    ("RÉPARÉS → RT", f"={st_cnt(ST_REP_DONE)}", BLUE_BG, BLUE_FG),
    ("VT EN ATTENTE / REJETÉ", f"={st_cnt(ST_VT_WAIT)}+{st_cnt(ST_VT_REJ)}", "FFF2CC", "7F6000"),
    ("À SOUDER (listés)", f"={st_cnt(ST_TO_WELD)}", "EDEDED", "595959"),
]
for i, (lab, f, fmt, color, note) in enumerate(cards1):
    c1, c2 = CL(2 + 2 * i), CL(3 + 2 * i)
    for rr, rng_ in ((7, f"{c1}7:{c2}7"), (8, f"{c1}8:{c2}9"), (10, f"{c1}10:{c2}10")):
        db.merge_cells(rng_)
    db[f"{c1}7"] = lab; db[f"{c1}8"] = f; db[f"{c1}10"] = note
    db[f"{c1}7"].font = font(9, True, "FFFFFF"); db[f"{c1}8"].font = font(22, True, "FFFFFF")
    db[f"{c1}10"].font = font(8, False, "F2F2F2", True); db[f"{c1}8"].number_format = fmt
    for rr in (7, 8, 9, 10):
        for cc in (c1, c2):
            db[f"{cc}{rr}"].fill = fill(color); db[f"{cc}{rr}"].alignment = CENTER
            db[f"{cc}{rr}"].border = Border(left=Side("thick", "FFFFFF"), right=Side("thick", "FFFFFF"))
db.row_dimensions[7].height = 22; db.row_dimensions[8].height = 22; db.row_dimensions[9].height = 22
for i, (lab, f, bg, fg) in enumerate(cards2):
    c1, c2 = CL(2 + 2 * i), CL(3 + 2 * i)
    db.merge_cells(f"{c1}12:{c2}12"); db.merge_cells(f"{c1}13:{c2}13")
    db[f"{c1}12"] = lab; db[f"{c1}13"] = f
    db[f"{c1}12"].font = font(9, True, fg); db[f"{c1}13"].font = font(18, True, fg)
    for rr in (12, 13):
        for cc in (c1, c2):
            db[f"{cc}{rr}"].fill = fill(bg); db[f"{cc}{rr}"].alignment = CENTER
            db[f"{cc}{rr}"].border = Border(left=Side("thick", "FFFFFF"), right=Side("thick", "FFFFFF"))
db.row_dimensions[13].height = 30

# ---- situation par ligne
section(db, "B16:M16", "SITUATION PAR LIGNE")
SH = ["Ligne", "Joints estimés", "Joints soudés", "Reste à souder", "% Soudage", "Radiographiés", "% RT",
      "Acceptés", "% Accepté", "En réparation", "Taux de réparation", "En attente RT"]
for i, h in enumerate(SH):
    head(db, 17, 2 + i, h, NAVY)
db.row_dimensions[17].height = 30
S_FIRST = 18
for i in range(MAX_LINES):
    r = S_FIRST + i
    pr_ = L_FIRST + i
    ln = f"PARAMÈTRES!$B${pr_}"
    c_ = lambda *crit: "COUNTIFS(" + ",".join([LN, f"$B{r}"] + list(crit)) + ")"
    stc = lambda s: c_(RNG(ST), f'"{s}"')
    db[f"B{r}"] = f'=IF({ln}="","",{ln})'
    g = lambda expr: f'=IF($B{r}="","",{expr})'
    db[f"C{r}"] = g(f"N(PARAMÈTRES!$E${pr_})")
    db[f"D{r}"] = g(c_(RNG(P), '"<>"'))
    db[f"E{r}"] = g(f"MAX(0,C{r}-D{r})")
    db[f"F{r}"] = g(f"IFERROR(D{r}/C{r},0)")
    db[f"G{r}"] = g(c_(RNG(V1), '"<>"'))
    db[f"H{r}"] = g(f"IFERROR(G{r}/D{r},0)")
    db[f"I{r}"] = g(stc(ST_ACC))
    db[f"J{r}"] = g(f"IFERROR(I{r}/D{r},0)")
    db[f"K{r}"] = g(f"{stc(ST_REP)}+{stc(ST_CO)}+{stc(ST_VT_REJ)}")
    rej = f'({c_(RNG(V1), chr(34) + "R" + chr(34))}+{c_(RNG(V1), chr(34) + "CO" + chr(34))})'
    db[f"L{r}"] = g(f'IFERROR({rej}/({c_(RNG(V1), chr(34) + "A" + chr(34))}+{rej}),0)')
    db[f"M{r}"] = g(f"{stc(ST_TO_RT)}+{stc(ST_REP_DONE)}+{stc(ST_NX)}")
    for i2, fmt in enumerate([None, "#,##0", "#,##0", "#,##0", PCT, "#,##0", PCT, "#,##0", PCT, "#,##0", PCT, "#,##0"]):
        style(db.cell(r, 2 + i2), fmt, bold=(i2 == 0), bg="F7F9FC" if i % 2 else None)
S_LAST = S_FIRST + MAX_LINES - 1
tr = S_LAST + 1
db[f"B{tr}"] = "TOTAL"
for colx in "CDEGIKM":
    db[f"{colx}{tr}"] = f"=SUM({colx}{S_FIRST}:{colx}{S_LAST})"
db[f"F{tr}"] = f"=IFERROR(D{tr}/C{tr},0)"; db[f"H{tr}"] = f"=IFERROR(G{tr}/D{tr},0)"; db[f"J{tr}"] = f"=IFERROR(I{tr}/D{tr},0)"
db[f"L{tr}"] = f'=IFERROR((COUNTIF({RNG(V1)},"R")+COUNTIF({RNG(V1)},"CO"))/(COUNTIF({RNG(V1)},"A")+COUNTIF({RNG(V1)},"R")+COUNTIF({RNG(V1)},"CO")),0)'
for i2, fmt in enumerate([None, "#,##0", "#,##0", "#,##0", PCT, "#,##0", PCT, "#,##0", PCT, "#,##0", PCT, "#,##0"]):
    style(db.cell(tr, 2 + i2), fmt, bold=True, color="FFFFFF", bg=GOLD)
for colx in "FHJ":
    db.conditional_formatting.add(f"{colx}{S_FIRST}:{colx}{S_LAST}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
db.conditional_formatting.add(f"L{S_FIRST}:L{S_LAST}", CellIsRule(operator="greaterThan", formula=["0.05"], fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
db.conditional_formatting.add(f"K{S_FIRST}:K{S_LAST}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))

ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.gapWidth = 60
ch.title = "Avancement par ligne"
ch.add_data(Reference(db, min_col=3, max_col=4, min_row=17, max_row=S_LAST), titles_from_data=True)
ch.add_data(Reference(db, min_col=9, max_col=9, min_row=17, max_row=S_LAST), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=S_FIRST, max_row=S_FIRST + len(LINES) - 1))
for s_, colr in zip(ch.series, ["BFBFBF", NAVY, "548235"]):
    s_.graphicalProperties.solidFill = colr
    s_.graphicalProperties.line.solidFill = colr
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.majorGridlines = None
ch.legend.position = "b"; ch.width, ch.height = 15.5, 8
db.add_chart(ch, f"B{tr + 2}")

# ---- statut breakdown (filtered)
SB = tr + 2
db[f"J{SB}"] = "Statut"; db[f"L{SB}"] = "Joints"; db[f"M{SB}"] = "%"
db.merge_cells(f"J{SB}:K{SB}")
for cc in ("J", "K", "L", "M"):
    style(db[f"{cc}{SB}"], bold=True, color="FFFFFF", bg=NAVY)
for i, s in enumerate(STATUSES):
    r = SB + 1 + i
    db.merge_cells(f"J{r}:K{r}")
    db[f"J{r}"] = s
    db[f"L{r}"] = f"={st_cnt(s)}"
    db[f"M{r}"] = f"=IFERROR(L{r}/SUM($L${SB + 1}:$L${SB + len(STATUSES)}),0)"
    bg, fg = STATUS_STYLE[s]
    style(db[f"J{r}"], bold=True, color=fg, bg=bg, align=LEFT); db[f"K{r}"].border = BORDER; db[f"K{r}"].fill = fill(bg)
    style(db[f"L{r}"], "#,##0"); style(db[f"M{r}"], PCT)
pie = PieChart(); pie.title = "Répartition des statuts"
pie.add_data(Reference(db, min_col=12, min_row=SB, max_row=SB + len(STATUSES)), titles_from_data=True)
pie.set_categories(Reference(db, min_col=10, min_row=SB + 1, max_row=SB + len(STATUSES)))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False; pie.dataLabels.showCatName = False; pie.dataLabels.showSerName = False
pie.dataLabels.showLeaderLines = False
pie.legend.position = "r"; pie.width, pie.height = 12.5, 6.5
db.add_chart(pie, f"J{SB + len(STATUSES) + 2}")

# ---- welders
WS = SB + len(STATUSES) + 18
section(db, f"B{WS}:M{WS}", "PERFORMANCE DES SOUDEURS (toutes lignes)")
for i, h in enumerate(["WR", "Nom", "", "Joints soudés", "Radiographiés", "Rejetés", "Taux de réparation", "Évaluation"]):
    head(db, WS + 1, 2 + i, h, TEAL)
db.merge_cells(f"C{WS + 1}:D{WS + 1}")
for i in range(len(WELDERS)):
    r, s_ = WS + 2 + i, 6 + i
    db[f"B{r}"] = f"=SOUDEURS!B{s_}"
    db.merge_cells(f"C{r}:D{r}")
    db[f"C{r}"] = f'=IF(SOUDEURS!C{s_}="","",SOUDEURS!C{s_})'
    db[f"E{r}"] = f"=SOUDEURS!F{s_}"; db[f"F{r}"] = f"=SOUDEURS!G{s_}"; db[f"G{r}"] = f"=SOUDEURS!H{s_}"
    db[f"H{r}"] = f"=SOUDEURS!I{s_}"; db[f"I{r}"] = f"=SOUDEURS!K{s_}"
    for cc, fmt in zip("BCDEFGHI", [None, None, None, "#,##0", "#,##0", "#,##0", PCT, None]):
        style(db[f"{cc}{r}"], fmt, bold=cc == "B", align=LEFT if cc == "C" else CENTER, color="008000")
WL = WS + 1 + len(WELDERS)
db.conditional_formatting.add(f"H{WS + 2}:H{WL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.2, color="F4B183"))
db.conditional_formatting.add(f"I{WS + 2}:I{WL}", CellIsRule(operator="equal", formula=['"Excellent"'], fill=fill(GREEN_BG), font=Font(name=FONT, bold=True, color=GREEN_FG)))
db.conditional_formatting.add(f"I{WS + 2}:I{WL}", CellIsRule(operator="equal", formula=['"À surveiller"'], fill=fill(RED_BG), font=Font(name=FONT, bold=True, color=RED_FG)))
bar = BarChart(); bar.type = "bar"; bar.title = "Joints soudés par soudeur"
bar.add_data(Reference(db, min_col=5, min_row=WS + 1, max_row=WL), titles_from_data=True)
bar.set_categories(Reference(db, min_col=2, min_row=WS + 2, max_row=WL))
bar.series[0].graphicalProperties.solidFill = TEAL
bar.x_axis.delete = False; bar.y_axis.delete = False; bar.y_axis.majorGridlines = None
bar.legend = None; bar.width, bar.height = 9.5, 9
db.add_chart(bar, f"J{WS + 1}")
db.page_setup.orientation = "portrait"; db.page_setup.fitToWidth = 1; db.page_setup.fitToHeight = 0
db.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== GUIDE (Arabic)
gd = wb.create_sheet("دليل الاستخدام")
gd.sheet_view.rightToLeft = True
gd.sheet_view.showGridLines = False
gd.sheet_properties.tabColor = "7F7F7F"
banner(gd, "B2:C2", "دليل استخدام سجل اللحام")
gd.column_dimensions["A"].width = 2; gd.column_dimensions["B"].width = 30; gd.column_dimensions["C"].width = 100
G = [
    ("1. PARAMÈTRES", "اكتب معلومات المشروع، وأسماء الخطوط (ISB-302، ISB-303، ISB-304…) وعدد الوصلات المتوقع لكل خط في الخلايا الصفراء."),
    ("2. SOUDEURS", "قائمة اللحامين: الرمز WR والاسم والتأهيل. الأداء (عدد الوصلات، الرفض، نسبة الإصلاح) يُحسب تلقائياً."),
    ("3. CARNET DE SOUDURE", "سطر واحد لكل وصلة (joint). املأ الأعمدة الزرقاء فقط: الخط، رقم الوصلة، اللحامين WR لكل تمريرة، تاريخ اللحام، نتيجة VT ثم RT-01."),
    ("   إذا كانت نتيجة RT = R أو CO", "املأ قسم «RÉPARATION 1»: لحام الإصلاح WR، نتيجة VT بعد الإصلاح، ثم تاريخ ونتيجة RT الجديد. وإذا رُفض مرة أخرى استعمل «RÉPARATION 2»."),
    ("   إذا كانت النتيجة NX", "إعادة الفيلم: تُجرى صورة إشعاعية جديدة فقط دون إصلاح. سجّل نتيجة RT الجديدة في «R1 RT» واترك «R1 VT» فارغاً."),
    ("   عمود STATUT", "يُحسب تلقائياً: À SOUDER ← VT EN ATTENTE ← À RADIOGRAPHIER ← ACCEPTÉ، أو À RÉPARER / COUPE / NX حسب النتائج، حتى تصبح الوصلة مقبولة."),
    ("4. PROGRAMME RT", "كل يوم: اكتب التاريخ واختر الخط، فتظهر تلقائياً قائمة الوصلات التي تحتاج صورة إشعاعية (جديدة، أو بعد إصلاح، أو NX)، مع ملخص أعمال ذلك اليوم."),
    ("5. RÉPARATIONS", "قائمة تلقائية بكل الوصلات المرفوضة التي تنتظر الإصلاح، مع عدد أيام الانتظار (أحمر إذا تجاوز 3 أيام)."),
    ("6. TABLEAU DE BORD", "لوحة التحكم: اختر خطاً أو TOUTES، فتظهر نسبة التقدم وعدد الوصلات المقبولة ونسبة الإصلاح وأداء اللحامين والرسوم البيانية."),
    ("", ""),
    ("رموز النتائج", "A = مقبول   •   R = مرفوض، يجب الإصلاح   •   NX = إعادة الفيلم، تُجرى صورة إشعاعية جديدة   •   CO = قطع الوصلة وإعادة لحامها"),
    ("الألوان", "أزرق = بيانات تُدخلها أنت   •   أسود = صيغ تلقائية (لا تكتب فوقها)   •   أصفر = إعدادات   •   رقم وصلة بالأحمر = مكرر على نفس الخط"),
    ("ملاحظة", "بيانات الخط TFT 716 منقولة من ملفك الأصلي (205 وصلات). في الملف الأصلي كان الرمز NR يعني «للإصلاح»، فاستُبدل بالرمز R. خطوط ISB جاهزة لإدخال بياناتها."),
]
for i, (a, b) in enumerate(G):
    r = 4 + i
    gd[f"B{r}"], gd[f"C{r}"] = a, b
    gd[f"B{r}"].font = font(11, True, NAVY)
    gd[f"C{r}"].font = font(11)
    gd[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="center", readingOrder=2)
    gd[f"B{r}"].alignment = Alignment(vertical="center", readingOrder=2)
    gd.row_dimensions[r].height = 36 if b else 10

wb.active = 0
wb.save(OUT)
print("joints listed:", len(joints), "| rt date:", LAST_DATE.date(), "->", OUT)
