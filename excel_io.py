"""Import des Raumprogramm-Rev8-Raumbuchs und Export als Excel/CSV."""
from __future__ import annotations

from typing import Any, Dict, List

import os

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import calc
from model import Project, new_consumer, new_product, new_room

# Spaltenmapping des Rev8-Raumbuchs (Header-Zeile 18, Daten ab Zeile 19)
ROOM_MAP = {
    "H": "bauteil", "I": "ebene", "J": "versorgung", "K": "baustufe",
    "M": "r_nr", "N": "r_name", "O": "flaeche", "P": "hoehe",
    "R": "t_winter", "S": "t_soll",
    "W": "verbrauchsart_k", "X": "q0_ber", "Y": "q0_zusch", "Z": "q0_produkt",
    "AD": "kreis_k",
    "AF": "verbrauchsart_h", "AG": "qh_ber", "AH": "qh_zusch", "AI": "qh_produkt",
    "AM": "kreis_h",
}
CONSUMER_MAP = {
    "H": "baut", "I": "gesch", "J": "versorgung", "K": "baustufe",
    "M": "r_nr", "N": "bezeichnung", "O": "ma_nr", "P": "ma_typ",
    "S": "verbrauchsart_k", "U": "q0", "V": "glz_k", "W": "kreis_k",
    "Y": "verbrauchsart_h", "AA": "qh", "AB": "glz_h", "AC": "kreis_h",
}
PRODUCT_MAP = {
    "H": "baut", "I": "gesch", "J": "versorgung", "K": "baustufe",
    "M": "r_nr", "N": "produkt", "O": "verbrauchsart", "P": "prozess",
    "Q": "prod_zeit", "S": "ton_p_tag", "T": "kg_batch", "U": "verweilzeit",
    "W": "t1", "X": "t2", "Z": "sf", "AA": "kreis",
}


def _read_rows(ws, col_map: Dict[str, str], start_row: int = 19,
               required: tuple = ("N",)) -> List[Dict[str, Any]]:
    rows = []
    for r in range(start_row, ws.max_row + 1):
        row = {}
        empty = True
        for col, field in col_map.items():
            v = ws[f"{col}{r}"].value
            if isinstance(v, str) and v.startswith("="):
                v = None  # Formeln nicht übernehmen – wird neu berechnet
            if col in required and v not in (None, ""):
                empty = False
            row[field] = v
        if not empty:
            name = str(row.get("r_nr") or row.get("r_name") or row.get("bezeichnung") or row.get("produkt") or "")
            if "summe" in name.lower() or "total" in name.lower():
                continue  # Summenzeilen des Rev8-Blatts überspringen
            rows.append(row)
    return rows


def import_rev8(path: str) -> Project:
    """Liest Räume, Verbraucher und Produkte aus der Rev8-Datei."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    project = Project()
    try:
        ws = wb["Room"]
        for row in _read_rows(ws, ROOM_MAP):
            project.rooms.append(new_room(**row))
        ws = wb["Consumer"]
        for row in _read_rows(ws, CONSUMER_MAP):
            project.consumers.append(new_consumer(**row))
        ws = wb["Produkt"]
        for row in _read_rows(ws, PRODUCT_MAP):
            project.products.append(new_product(**row))
    finally:
        wb.close()
    return project


# ------------------------------------------------------------------
# Export
# ------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _sheet_from_rows(ws, headers: List[str], rows: List[List[Any]], widths: List[int]) -> None:
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if isinstance(v, float):
                cell.number_format = "0.00"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def export_excel(project: Project, path: str) -> None:
    wb = openpyxl.Workbook()

    # Projekt (mit Branding: Logo + Kopfzeile wie im Raumbuch)
    ws = wb.active
    ws.title = "Projekt"
    logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_full.png")
    if os.path.exists(logo):
        img = XLImage(logo)
        img.width, img.height = 240, 71
        ws.add_image(img, "B2")
    ws["G2"] = project.meta.get("firma", "")
    ws["G2"].font = Font(bold=True, size=12)
    ws["G3"] = " | ".join(str(project.meta.get(k, "")) for k in
                          ("projekt_nr", "projekt_name", "projekt_phase", "datum", "projekt_status")
                          if project.meta.get(k))
    header_row = 6
    for c, h in enumerate(["Feld", "Wert"], 1):
        cell = ws.cell(header_row, c, h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
    for i, (k, v) in enumerate(project.meta.items()):
        ws.cell(header_row + 1 + i, 1, k).border = BORDER
        ws.cell(header_row + 1 + i, 2, v).border = BORDER
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 50

    # Räume
    ws = wb.create_sheet("Räume")
    headers = ["R.-Nr.", "Raumname", "Baut.", "Ebene", "Baustufe", "Fläche [m²]",
               "Höhe [m]", "Volumen [m³]", "tR Soll [°C]",
               "Verbrauchsart K", "q0 [W/m²]", "Q0 [kW]", "q1 [W/m³]", "Kreis K",
               "Verbrauchsart H", "qh [W/m²]", "Qh [kW]", "Kreis H", "Bemerkung"]
    rows = []
    for rr in calc.room_rows(project):
        rm = rr["room"]
        rows.append([
            rm.get("r_nr"), rm.get("r_name"), rm.get("bauteil"), rm.get("ebene"),
            rm.get("baustufe"), rm.get("flaeche"), rm.get("hoehe"), rr["vol"],
            rm.get("t_soll"), rm.get("verbrauchsart_k"),
            round(rr["q0"], 2), round(rr["q0_kw"], 3),
            round(rr["q1"], 2) if rr["q1"] is not None else None,
            rm.get("kreis_k"), rm.get("verbrauchsart_h"),
            round(rr["qh"], 2), round(rr["qh_kw"], 3), rm.get("kreis_h"),
            rm.get("bemerkung"),
        ])
    _sheet_from_rows(ws, headers, rows, [8, 26, 6, 8, 9, 9, 8, 10, 9, 14, 10, 9, 10, 8, 14, 10, 9, 8, 24])

    # Verbraucher
    ws = wb.create_sheet("Verbraucher")
    headers = ["Baut.", "Ebene", "R.-Nr.", "Bezeichnung", "Ma.-Nr.", "Ma.-Typ",
               "Verbrauchsart K", "Q0 [kW]", "GLZ", "Q0 eff. [kW]", "Kreis K",
               "Verbrauchsart H", "Qh [kW]", "GLZ", "Qh eff. [kW]", "Kreis H", "Bemerkung"]
    rows = []
    for c in project.consumers:
        rows.append([
            c.get("baut"), c.get("gesch"), c.get("r_nr"), c.get("bezeichnung"),
            c.get("ma_nr"), c.get("ma_typ"), c.get("verbrauchsart_k"),
            c.get("q0"), c.get("glz_k"), round(calc.consumer_q0_kw(c), 3), c.get("kreis_k"),
            c.get("verbrauchsart_h"), c.get("qh"), c.get("glz_h"),
            round(calc.consumer_qh_kw(c), 3), c.get("kreis_h"), c.get("bemerkung"),
        ])
    _sheet_from_rows(ws, headers, rows, [6, 8, 8, 26, 8, 10, 16, 9, 6, 10, 8, 14, 9, 6, 10, 8, 24])

    # Produkte
    ws = wb.create_sheet("Produkte")
    headers = ["R.-Nr.", "Produkt", "Verbrauchsart", "Prozess", "Prod.-Zeit [h/d]",
               "t/Tag", "kg/h", "t1 [°C]", "t2 [°C]", "SF", "Q0_P [kW]", "Kreis", "Bemerkung"]
    rows = []
    for p in project.products:
        m = calc.product_mass_flow(p)
        q = calc.product_q0_kw(p)
        rows.append([
            p.get("r_nr"), p.get("produkt"), p.get("verbrauchsart"), p.get("prozess"),
            p.get("prod_zeit"), p.get("ton_p_tag"),
            round(m, 1) if m else None, p.get("t1"), p.get("t2"), p.get("sf"),
            round(q, 3) if q else None, p.get("kreis"), p.get("bemerkung"),
        ])
    _sheet_from_rows(ws, headers, rows, [8, 22, 16, 8, 12, 8, 9, 8, 8, 6, 10, 8, 24])

    # Kreise
    ws = wb.create_sheet("Kreise")
    headers = ["Kreis", "Medium", "t0 [°C]", "ts VL [°C]", "ts RL [°C]", "ΔP [K]",
               "k [W/m²K]", "Q0 [kW]", "Qh [kW]"]
    totals = calc.compute_totals(project)
    rows = []
    for c in project.circuits:
        t = totals.get(c["id"], {"q0": 0, "qh": 0})
        rows.append([c["id"], c.get("medium"), c.get("t0"), c.get("ts_vl"),
                     c.get("ts_rl"), c.get("dp"), c.get("k"),
                     round(t["q0"], 3), round(t["qh"], 3)])
    rows.append(["SUMME", "", "", "", "", "", "",
                 round(totals["SUMME"]["q0"], 3), round(totals["SUMME"]["qh"], 3)])
    _sheet_from_rows(ws, headers, rows, [8, 8, 8, 9, 9, 7, 10, 10, 10])

    wb.save(path)


def export_csv(project: Project, path: str) -> None:
    import csv
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["Kreis", "Medium", "t0 [°C]", "Q0 [kW]", "Qh [kW]"])
        totals = calc.compute_totals(project)
        for c in project.circuits:
            t = totals.get(c["id"], {"q0": 0, "qh": 0})
            w.writerow([c["id"], c.get("medium"), c.get("t0"),
                        round(t["q0"], 3), round(t["qh"], 3)])
        w.writerow(["SUMME", "", "", round(totals["SUMME"]["q0"], 3),
                    round(totals["SUMME"]["qh"], 3)])
