"""Erzeugt Wirbelstrasse.xlsx: Karmansche Wirbelstrasse (Wirbelstaerke-Stromfunktion) nur mit Excel-Formeln.

Ablauf pro Neuberechnung (F9) = ein Zeitschritt. Die Rechenbloecke stehen untereinander auf dem
Blatt "Rechnung", weil Excel Zirkelbezuege zeilenweise von oben nach unten durchrechnet:
  W0  = alte Wirbelstaerke (Kopie von W1 aus dem letzten Durchlauf)
  U,V = Geschwindigkeit aus der Stromfunktion
  W1  = neue Wirbelstaerke  (Advektion 2. Ordnung upwind + Diffusion, Euler explizit)
  PSI = Stromfunktion        (ein SOR-Sweep  nabla^2 psi = -omega pro Schritt)
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

NX, NY = 130, 50          # Gitter (Zellen)
OUT = sys.argv[1] if len(sys.argv) > 1 else "Wirbelstrasse.xlsx"

F = "Arial"
GREEN, GOLD, YELLOW = "1F4D3A", "F2A900", "FFE699"
font = lambda **k: Font(name=F, **k)

wb = Workbook()
ws = wb.active
ws.title = "Strömung"
rc = wb.create_sheet("Rechnung")
info = wb.create_sheet("Anleitung")

# ---------------------------------------------------------------- Parameter
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.row_dimensions[2].height = 30
ws["B2"] = "Wirbelstraße"; ws["B2"].font = font(size=22, bold=True, color=GREEN)
ws["B3"] = "∂ω/∂t + u·∇ω = ν∇²ω     ∇²ψ = −ω"; ws["B3"].font = font(size=12, italic=True, color="555555")

params = [  # (Symbol, Text, Name, Wert/Formel, Einheit, Eingabe?, Kommentar)
    ("U∞", "Anströmgeschwindigkeit", "Uinf", 1.0, "Zelle/t", True, "Geschwindigkeit am Einlass (links)."),
    ("D", "Zylinderdurchmesser", "Dzyl", 10, "Zellen", True, "Durchmesser des Zylinders in Gitterzellen."),
    ("Re", "Reynoldszahl", "Reynolds", 150, "–", True, "Ab ca. Re ≈ 47 lösen sich Wirbel ab. Bei Re < 40 bleibt der Nachlauf ruhig."),
    ("Δt", "Zeitschritt", "dtime", 0.1, "t", True, "Stabilitätsgrenze grob: Δt < 2ν/U∞² und Δt < 0.5. Zu groß → die Rechnung explodiert."),
    ("ν", "Viskosität  = U∞·D / Re", "visc", "=Uinf*Dzyl/Reynolds", "Zelle²/t", False, "Kinematische Viskosität, wird berechnet."),
    ("β", "SOR-Faktor (Poisson)", "beta_sor", 1.5, "–", True, "Überrelaxation für die Stromfunktion, 1 … 1.9. Ein Sweep pro Zeitschritt."),
    ("x₀", "Zylindermitte x", "xzyl", 25, "Zelle", True, f"Spalte der Zylindermitte (0 … {NX-1})."),
    ("y₀", "Zylindermitte y", "yzyl", 25.5, "Zelle", True, f"Zeile der Zylindermitte (0 … {NY-1}). Leicht außermittig (Kanalmitte = {(NY-1)/2}), damit die Symmetrie bricht und die Wirbel schneller starten."),
    ("R", "Radius  = D / 2", "Rzyl", "=Dzyl/2", "Zellen", False, "Radius."),
    ("ψ_K", "Stromfunktion am Körper", "psiBody", "=Uinf*yzyl", "", False, "ψ ist auf der Zylinderoberfläche konstant."),
    ("", "Reset (1 = Anfangszustand)", "Reset", 1, "", True, "1 = alles auf Anfang. Zum Rechnen auf 0 setzen und F9 drücken (oder Makro 'Starten')."),
    ("n", "Schritt", "Schritt", "=IF(Reset,0,Schritt+1)", "", False, "Zählt die Zeitschritte (Zirkelbezug auf sich selbst)."),
    ("", "Schritte pro Bild (Makro)", "Schritte_pro_Bild", 5, "", True, "Wie viele Zeitschritte das Makro rechnet, bevor der Bildschirm neu gezeichnet wird."),
]
thin = Side(style="thin", color="BBBBBB")
# Alle Spalten ab B sind schmal (Pixel der Anzeige); die Parametertabelle nutzt verbundene Zellen.
CW = 0.6
for cidx in range(2, 2 + NX + 2):
    ws.column_dimensions[L(cidx)].width = CW
SYM, TXT, VAL, UNI = (2, 5), (6, 40), (41, 52), (53, 61)
def put(r, span, value):
    ws.merge_cells(start_row=r, start_column=span[0], end_row=r, end_column=span[1])
    return ws.cell(r, span[0], value)
for k, (sym, txt, name, val, unit, inp, com) in enumerate(params):
    r = 5 + k
    ws.row_dimensions[r].height = 16
    put(r, SYM, sym).font = font(size=11, italic=True)
    put(r, TXT, txt).font = font(size=10, color="444444")
    c = put(r, VAL, val)
    c.font = font(size=11, bold=True, color="0000FF" if inp else ("FFFFFF" if name == "Schritt" else "000000"))
    c.alignment = Alignment(horizontal="right")
    for cc in range(VAL[0], VAL[1] + 1):
        ws.cell(r, cc).border = Border(top=thin, bottom=thin, left=thin if cc == VAL[0] else None, right=thin if cc == VAL[1] else None)
        if inp:
            ws.cell(r, cc).fill = PatternFill("solid", fgColor=YELLOW)
        if name == "Schritt":
            ws.cell(r, cc).fill = PatternFill("solid", fgColor=GREEN)
    if name == "Schritt":
        c.number_format = "#,##0"
    elif name == "visc":
        c.number_format = "0.000"
    elif isinstance(val, float):
        c.number_format = "0.00"
    c.comment = Comment(com, "Wirbelstraße")
    put(r, UNI, unit).font = font(size=9, color="777777")
    wb.defined_names[name] = DefinedName(name, attr_text=f"'Strömung'!${L(VAL[0])}${r}")
for k, t in enumerate([
        "Gelb = Eingabe (blaue Zahl).  Fahrt die Maus über einen Wert für Erklärungen.",
        "Ohne Makro: Reset auf 0 setzen, dann F9 gedrückt halten (1 × F9 = 1 Zeitschritt).",
        "Mit Makro (siehe Blatt 'Anleitung'): Starten / Anhalten / Zuruecksetzen / Einzelschritt.",
        "Farbe = Wirbelstärke ω:  grün = rechtsdrehend (ω < 0),  gold = linksdrehend (ω > 0),  schwarz = Zylinder."]):
    ws.cell(5 + k, 65, t).font = font(size=10, color="444444")

# ---------------------------------------------------------------- Rechnung
rc.sheet_view.zoomScale = 70
rc.column_dimensions["A"].width = 12
rc.column_dimensions["B"].width = 4
for i in range(NX):
    rc.column_dimensions[L(3 + i)].width = 5.5
col = lambda i: L(3 + i)

blocks = ["MASK", "NF", "W0", "U", "V", "W1", "PSI"]
titles = {
    "MASK": "Zylinder: 1 = fest, 0 = Fluid",
    "NF": "Anzahl Fluid-Nachbarn (für Wandwirbelstärke)",
    "W0": "ω alt  (Kopie von ω neu aus dem letzten Schritt)",
    "U": "u = ∂ψ/∂y",
    "V": "v = −∂ψ/∂x",
    "W1": "ω neu = ω alt + Δt·(−u·∇ω + ν∇²ω)",
    "PSI": "ψ  (ein SOR-Sweep von ∇²ψ = −ω)",
}
top, hdr = {}, {}
r = 1
for b in blocks:
    rc.cell(r, 1, b).font = font(bold=True, color="FFFFFF")
    rc.cell(r, 1).fill = PatternFill("solid", fgColor=GREEN)
    rc.cell(r, 3, titles[b]).font = font(bold=True, color=GREEN)
    hdr[b] = r + 1
    rc.cell(r + 1, 2, "j \\ i").font = font(size=8, color="777777")
    for i in range(NX):
        rc.cell(r + 1, 3 + i, i).font = font(size=8, color="777777")
    top[b] = r + 2
    for j in range(NY):
        rc.cell(top[b] + NY - 1 - j, 2, j).font = font(size=8, color="777777")
    r = top[b] + NY + 1

def ref(b, i, j, absrow=False):
    return f"{col(i)}{top[b] + NY - 1 - j}"
jref = lambda b, j: f"$B{top[b] + NY - 1 - j}"
iref = lambda b, i: f"{col(i)}${hdr[b]}"

f8 = font(size=8)
for j in range(NY):
    for i in range(NX):
        R = lambda b, di=0, dj=0: ref(b, i + di, j + dj)
        # Zylinder-Maske (statisch)
        rc[R("MASK")] = f"=IF(({iref('MASK', i)}-xzyl)^2+({jref('MASK', j)}-yzyl)^2<=Rzyl^2,1,0)"
        # Fluid-Nachbarn
        if 0 < i < NX - 1 and 0 < j < NY - 1:
            rc[R("NF")] = (f"=4-{R('MASK',1)}-{R('MASK',-1)}-{R('MASK',0,1)}-{R('MASK',0,-1)}")
        else:
            rc[R("NF")] = 0
        # alte Wirbelstaerke
        rc[R("W0")] = f"=IF(Reset,0,{R('W1')})"
        # Geschwindigkeit
        if 0 < i < NX - 1 and 0 < j < NY - 1:
            rc[R("U")] = f"=({R('PSI',0,1)}-{R('PSI',0,-1)})/2"
            rc[R("V")] = f"=-({R('PSI',1)}-{R('PSI',-1)})/2"
        else:
            rc[R("U")] = "=Uinf"
            rc[R("V")] = 0
        # neue Wirbelstaerke
        if i < 2 or j < 2 or j > NY - 3:
            rc[R("W1")] = 0                                   # Einlass / Ränder: ω = 0
        elif i > NX - 3:
            rc[R("W1")] = f"={ref('W1', NX - 3, j)}"          # Auslass: ∂ω/∂x = 0
        else:
            w, u, v = R("W0"), R("U"), R("V")
            wall = (f"2*(psiBody*{R('NF')}-((1-{R('MASK',1)})*{R('PSI',1)}+(1-{R('MASK',-1)})*{R('PSI',-1)}"
                    f"+(1-{R('MASK',0,1)})*{R('PSI',0,1)}+(1-{R('MASK',0,-1)})*{R('PSI',0,-1)}))/MAX({R('NF')},1)")
            adv = (f"MAX({u},0)*(3*{w}-4*{R('W0',-1)}+{R('W0',-2)})/2"
                   f"+MIN({u},0)*(-3*{w}+4*{R('W0',1)}-{R('W0',2)})/2"
                   f"+MAX({v},0)*(3*{w}-4*{R('W0',0,-1)}+{R('W0',0,-2)})/2"
                   f"+MIN({v},0)*(-3*{w}+4*{R('W0',0,1)}-{R('W0',0,2)})/2")
            lap = f"{R('W0',1)}+{R('W0',-1)}+{R('W0',0,1)}+{R('W0',0,-1)}-4*{w}"
            rc[R("W1")] = f"=IF(Reset,0,IF({R('MASK')},{wall},{w}+dtime*(-({adv})+visc*({lap}))))"
        # Stromfunktion
        y = jref("PSI", j)
        if j == 0 or j == NY - 1 or i == 0:
            rc[R("PSI")] = f"=Uinf*{y}"                       # ψ = U∞·y  (Einlass, oben, unten)
        elif i == NX - 1:
            rc[R("PSI")] = f"=IF(Reset,Uinf*{y},{R('PSI',-1)})"   # Auslass: ∂ψ/∂x = 0
        else:
            p = R("PSI")
            rc[R("PSI")] = (f"=IF(Reset,IF({R('MASK')},psiBody,Uinf*{y}),IF({R('MASK')},psiBody,"
                            f"(1-beta_sor)*{p}+beta_sor*({R('PSI',1)}+{R('PSI',-1)}+{R('PSI',0,1)}+{R('PSI',0,-1)}+{R('W1')})/4))")
        for b in blocks:
            c = rc[R(b)]
            c.font = f8
            c.number_format = "0" if b in ("MASK", "NF") else "0.00"

# Farbskalen auf den Rechenbloecken
for b in ("W0", "W1"):
    rng = f"{col(0)}{top[b]}:{col(NX-1)}{top[b]+NY-1}"
    rc.conditional_formatting.add(rng, ColorScaleRule(start_type="num", start_value=-1, start_color="1F6B4A",
        mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=1, end_color="F2A900"))
rng = f"{col(0)}{top['PSI']}:{col(NX-1)}{top['PSI']+NY-1}"
rc.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="F5EFC8", end_type="max", end_color="4F8F6F"))
rng = f"{col(0)}{top['MASK']}:{col(NX-1)}{top['MASK']+NY-1}"
rc.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", start_color="10231A", end_color="10231A"), font=Font(name=F, color="10231A")))
rc.freeze_panes = "C1"

# ---------------------------------------------------------------- Anzeige
D0R, D0C = 20, 2
ws.cell(D0R - 2, 2, "ω  Wirbelstärke").font = font(bold=True, color=GREEN, size=12)
put(D0R - 2, (41, 75), "=\"Schritt \"&TEXT(Schritt,\"#,##0\")&\"   t = \"&TEXT(Schritt*dtime,\"0.0\")").font = font(size=10, color="444444")
for jj in range(NY):
    ws.row_dimensions[D0R + jj].height = 5.25
for j in range(NY):
    for i in range(NX):
        c = ws.cell(D0R + NY - 1 - j, D0C + i,
                    f"=IF(Rechnung!{ref('MASK', i, j)},\"\",Rechnung!{ref('W1', i, j)})")
        c.number_format = ";;;"
rng = f"{L(D0C)}{D0R}:{L(D0C+NX-1)}{D0R+NY-1}"
ws.conditional_formatting.add(rng, FormulaRule(formula=[f"ISTEXT({L(D0C)}{D0R})"], fill=PatternFill("solid", start_color="10231A", end_color="10231A"), stopIfTrue=True))
ws.conditional_formatting.add(rng, ColorScaleRule(start_type="num", start_value=-0.6, start_color="1F6B4A",
    mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=0.6, end_color="F2A900"))
ws.cell(D0R + NY + 1, 2, "Einlass →  (links, ψ = U∞·y)        oben/unten: reibungsfreie Wand        Auslass rechts: ∂/∂x = 0").font = font(size=9, color="777777")

# ---------------------------------------------------------------- Anleitung
info.column_dimensions["A"].width = 110
lines = [
    ("So benutzt du die Datei", True),
    ("", False),
    ("Ohne Makros (reines .xlsx):", True),
    ("1. Die Datei steht auf 'Manuell berechnen' mit 'Iterative Berechnung' (max. 1 Iteration). Prüfen unter Datei → Optionen → Formeln.", False),
    ("   Wichtig: Excel übernimmt diese Einstellung von der ZUERST geöffneten Mappe – am besten Excel frisch nur mit dieser Datei öffnen.", False),
    ("2. Auf Blatt 'Strömung' die Zelle 'Reset' von 1 auf 0 setzen.", False),
    ("3. F9 drücken = 1 Zeitschritt. F9 gedrückt halten = Film. Nach ca. 600–800 Schritten beginnt der Nachlauf zu pendeln und Wirbel lösen sich ab.", False),
    ("4. Neustart: Reset auf 1, einmal F9, Reset wieder auf 0.", False),
    ("", False),
    ("Mit Makros (Knöpfe Starten / Anhalten / Reset / Einzelschritt):", True),
    ("1. Datei → Speichern unter → 'Excel-Arbeitsmappe mit Makros (*.xlsm)'.", False),
    ("2. Alt+F11 → Datei → Datei importieren → 'Wirbelstrasse.bas' (liegt neben dieser Datei).", False),
    ("3. Entwicklertools → Einfügen → Schaltfläche (Formularsteuerelement) → Makro zuweisen: Starten, Anhalten, Zuruecksetzen, Einzelschritt.", False),
    ("   (Ohne Knöpfe geht's auch: Alt+F8 → Makro auswählen → Ausführen.)", False),
    ("", False),
    ("Was passiert in jeder Zelle? (Blatt 'Rechnung', Blöcke von oben nach unten)", True),
    ("MASK  =WENN((i−x₀)²+(j−y₀)² ≤ R²;1;0) – zeichnet den Zylinder.", False),
    ("W0    ω alt = Kopie von ω neu aus dem letzten Durchlauf.", False),
    ("U, V  Geschwindigkeit aus der Stromfunktion: u = (ψ oben − ψ unten)/2,  v = −(ψ rechts − ψ links)/2.", False),
    ("W1    ω neu = ω alt + Δt·(−Transport + ν·Diffusion). Transport mit Upwind 2. Ordnung (je nach Vorzeichen von u, v),", False),
    ("      Diffusion = Summe der 4 Nachbarn − 4·Mitte. Am Zylinder: Wandwirbelstärke nach Thom, ω = 2·(ψ_K − ψ_Nachbar).", False),
    ("PSI   ψ = Mittelwert der 4 Nachbarn + ω/4 (überrelaxiert mit β). Ein Sweep pro Zeitschritt reicht, weil sich ψ nur langsam ändert.", False),
    ("", False),
    ("Warum klappt das in Excel?", True),
    ("Mit iterativer Berechnung darf eine Zelle sich selbst (oder ein späterer Block) referenzieren. Excel rechnet die Zirkelbezüge Zeile", False),
    ("für Zeile von oben nach unten – deshalb stehen die Blöcke genau in der Reihenfolge eines Zeitschritts untereinander.", False),
    ("Das Makro verlässt sich nicht darauf: Es rechnet die Blöcke per Range.Calculate explizit nacheinander (W0 → U → V → W1 → PSI).", False),
    ("", False),
    ("Spielwiese", True),
    ("Re = 30 → ruhiger, symmetrischer Nachlauf.  Re = 100…200 → schöne Wirbelstraße.  D kleiner/größer, Zylinder verschieben …", False),
    ("Wenn alles knallbunt wird / #ZAHL! erscheint: Δt war zu groß → Δt verkleinern, Reset.", False),
]
for k, (t, bold) in enumerate(lines):
    info.cell(1 + k, 1, t).font = font(size=12 if bold else 10, bold=bold, color=GREEN if bold else "222222")

for b in blocks:
    wb.defined_names[f"blk_{b}"] = DefinedName(f"blk_{b}", attr_text=f"Rechnung!${col(0)}${top[b]}:${col(NX-1)}${top[b]+NY-1}")
wb.defined_names["blk_Anzeige"] = DefinedName("blk_Anzeige", attr_text=f"'Strömung'!${L(D0C)}${D0R}:${L(D0C+NX-1)}${D0R+NY-1}")
wb.defined_names["Schrittanzeige"] = DefinedName("Schrittanzeige", attr_text=f"'Strömung'!${L(41)}${D0R-2}")
wb.calculation = CalcProperties(calcMode="manual", iterate=True, iterateCount=1, iterateDelta=0.000001, fullCalcOnLoad=True)
wb.save(OUT)
print("saved", OUT, {b: top[b] for b in blocks})
