"""Berechnungs-Engine – bildet die Formeln des Raumprogramms Rev8 nach.

Raum (Raumbuch):
    V  = A * H                                  [m³]
    q0 = q0_ber. + q0_zusch. + q0_produkt       [W/m²]
    Q0 = A * q0 / 1000                          [kW]
    q1 = Q0 * 1000 / V                          [W/m³]
analog Heizen (qh, Qh).

Verbraucher:
    Q0_eff = Q0 * GLZ                            [kW]

Produkt:
    m  = t/Tag * 1000 / Prod.-Zeit               [kg/h]
    CLG: Q0_P = m * cp1 * (t1 - t2) / 3600
    FRZ: Q0_P = m * (cp1*(t1 - tF) + r + cp2*(tF - t2)) / 3600
    DEF: Q0_P = m * (cp2*(t1 - tF) + r + cp1*(tF - t2)) / 3600
    Q0_eff = Q0_P * (1 + SF)

Kreissummen / Gesamtbedarf:
    Summe der effektiven Kühllasten je Kreis,
    getrennt Kühlen (Q0) und Heizen (Qh).
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from model import PRODUCT_DATA, Project


def _num(v: Any) -> Optional[float]:
    """Wandelt in float um; None/''/Text -> 0.0 bzw. None."""
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def room_volume(room: Dict[str, Any]) -> Optional[float]:
    a, h = _num(room.get("flaeche")), _num(room.get("hoehe"))
    if a and h:
        return a * h
    return None


def room_q0(room: Dict[str, Any]) -> float:
    return _num(room.get("q0_ber")) + _num(room.get("q0_zusch")) + _num(room.get("q0_produkt"))


def room_qh(room: Dict[str, Any]) -> float:
    return _num(room.get("qh_ber")) + _num(room.get("qh_zusch")) + _num(room.get("qh_produkt"))


def room_q0_kw(room: Dict[str, Any]) -> float:
    return _num(room.get("flaeche")) * room_q0(room) / 1000.0


def room_qh_kw(room: Dict[str, Any]) -> float:
    return _num(room.get("flaeche")) * room_qh(room) / 1000.0


def room_q1(room: Dict[str, Any]) -> Optional[float]:
    vol = room_volume(room)
    q0_kw = room_q0_kw(room)
    if vol:
        return q0_kw * 1000.0 / vol
    return None


def consumer_q0_kw(c: Dict[str, Any]) -> float:
    return _num(c.get("q0")) * _num(c.get("glz_k"))


def consumer_qh_kw(c: Dict[str, Any]) -> float:
    return _num(c.get("qh")) * _num(c.get("glz_h"))


def product_mass_flow(p: Dict[str, Any]) -> Optional[float]:
    t_tag, zeit = _num(p.get("ton_p_tag")), _num(p.get("prod_zeit"))
    if t_tag and zeit:
        return t_tag * 1000.0 / zeit
    return None


def product_q0_kw(p: Dict[str, Any]) -> Optional[float]:
    m = product_mass_flow(p)
    if not m:
        return None
    t1, t2 = _num(p.get("t1")), _num(p.get("t2"))
    cp1, cp2, r, tf = PRODUCT_DATA["cp1"], PRODUCT_DATA["cp2"], PRODUCT_DATA["r"], PRODUCT_DATA["t_freeze"]
    prozess = (p.get("prozess") or "").upper()
    if prozess == "CLG":
        q = m * cp1 * (t1 - t2) / 3600.0
    elif prozess == "FRZ":
        q = m * (cp1 * (t1 - tf) + r + cp2 * (tf - t2)) / 3600.0
    elif prozess == "DEF":
        q = m * (cp2 * (t1 - tf) + r + cp1 * (tf - t2)) / 3600.0
    else:
        return None
    return q * (1.0 + _num(p.get("sf")))


def compute_totals(project: Project) -> Dict[str, Dict[str, float]]:
    """Liefert je Kreis: {'q0': kW, 'qh': kW} plus 'SUMME'."""
    totals: Dict[str, Dict[str, float]] = {}
    for cid in project.circuit_ids():
        totals[cid] = {"q0": 0.0, "qh": 0.0}

    def add(cid: Optional[str], key: str, value: float) -> None:
        if cid and cid in totals:
            totals[cid][key] += value

    for room in project.rooms:
        add(room.get("kreis_k"), "q0", room_q0_kw(room))
        add(room.get("kreis_h"), "qh", room_qh_kw(room))
    for c in project.consumers:
        add(c.get("kreis_k"), "q0", consumer_q0_kw(c))
        add(c.get("kreis_h"), "qh", consumer_qh_kw(c))
    for p in project.products:
        q = product_q0_kw(p)
        if q:
            add(p.get("kreis"), "q0", q)

    totals["SUMME"] = {
        "q0": sum(v["q0"] for k, v in totals.items() if k != "SUMME"),
        "qh": sum(v["qh"] for k, v in totals.items() if k != "SUMME"),
    }
    return totals


def room_rows(project: Project) -> List[Dict[str, Any]]:
    """Berechnete Zeilen für die Raum-Tabelle (Anzeige/Export)."""
    rows = []
    for room in project.rooms:
        rows.append({
            "room": room,
            "vol": room_volume(room),
            "q0": room_q0(room),
            "q0_kw": room_q0_kw(room),
            "q1": room_q1(room),
            "qh": room_qh(room),
            "qh_kw": room_qh_kw(room),
        })
    return rows
