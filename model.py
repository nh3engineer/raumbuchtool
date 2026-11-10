"""Datenmodell des Raumprogramms (Raumbuch) für Kältetechnik.

Alle Objekte sind als einfache dicts mit JSON-Persistenz gehalten,
damit eigene Spalten (custom fields) jederzeit ergänzt werden können,
ohne den Code anzufassen.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Dict, List

VERSION = 1

DEFAULT_CIRCUITS = [
    {"id": "LP1", "medium": "NH3", "t0": -42.0, "ts_vl": 0.0, "ts_rl": 0.0, "dp": 1.0, "k": 18.0},
    {"id": "LP2", "medium": "NH3", "t0": -34.0, "ts_vl": 0.0, "ts_rl": 0.0, "dp": 1.0, "k": 18.0},
    {"id": "MP1", "medium": "PG", "t0": -10.0, "ts_vl": -8.0, "ts_rl": -4.0, "dp": 1.0, "k": 18.0},
    {"id": "MP2", "medium": "PG", "t0": -10.0, "ts_vl": 1.0, "ts_rl": 5.0, "dp": 0.0, "k": 18.0},
    {"id": "AC1", "medium": "H2O", "t0": -10.0, "ts_vl": 6.0, "ts_rl": 12.0, "dp": 0.0, "k": 18.0},
    {"id": "Qh", "medium": "PG", "t0": 30.0, "ts_vl": 30.0, "ts_rl": 20.0, "dp": 0.0, "k": 18.0},
]

DEFAULT_CONSUMER_TYPES = [
    "Raumkühlung", "TK-Lager", "NK-Lager", "Lüftung", "Klimatisierung",
    "Klimatisierung-Tech.", "Maschinen", "Eiswasser (Erz.)", "Produkt",
    "Bestand", "Erweiterung", "Reserve", "Option", "KEINE KÜHLUNG",
    "siehe Verbraucher", "siehe Räume", "-",
]

DEFAULT_HEATING_TYPES = [
    "Raumheizung", "Frostfrei", "Auftauen", "Türen/Tore", "-",
]

DEFAULT_PROCESSES = ["CLG", "FRZ", "DEF"]

ROOM_BASE_FIELDS = [
    ("r_nr", "R.-Nr.", "text"),
    ("r_name", "Raumname", "text"),
    ("bauteil", "Baut.", "text"),
    ("ebene", "Ebene", "text"),
    ("versorgung", "Versorgung", "text"),
    ("baustufe", "Baustufe", "text"),
    ("flaeche", "Fläche [m²]", "float"),
    ("hoehe", "Höhe [m]", "float"),
    ("t_winter", "tR Winter [°C]", "float"),
    ("t_soll", "tR Soll [°C]", "float"),
    ("verbrauchsart_k", "Verbrauchsart Kühlen", "choice"),
    ("q0_ber", "q0 ber. [W/m²]", "float"),
    ("q0_zusch", "q0 Zuschlag [W/m²]", "float"),
    ("q0_produkt", "q0 Produkt [W/m²]", "float"),
    ("kreis_k", "Kreis Kühlen", "text"),
    ("verbrauchsart_h", "Verbrauchsart Heizen", "choice_h"),
    ("qh_ber", "qh ber. [W/m²]", "float"),
    ("qh_zusch", "qh Zuschlag [W/m²]", "float"),
    ("qh_produkt", "qh Produkt [W/m²]", "float"),
    ("kreis_h", "Kreis Heizen", "text"),
    ("bemerkung", "Bemerkung", "text"),
]

CONSUMER_BASE_FIELDS = [
    ("baut", "Baut.", "text"),
    ("gesch", "Ebene", "text"),
    ("versorgung", "Versorgung", "text"),
    ("baustufe", "Baustufe", "text"),
    ("r_nr", "R.-Nr.", "text"),
    ("bezeichnung", "Bezeichnung", "text"),
    ("ma_nr", "Ma.-Nr.", "text"),
    ("ma_typ", "Ma.-Typ", "text"),
    ("verbrauchsart_k", "Verbrauchsart Kühlen", "choice"),
    ("q0", "Q0 [kW]", "float"),
    ("glz_k", "GLZ Kühlen", "float"),
    ("kreis_k", "Kreis Kühlen", "text"),
    ("verbrauchsart_h", "Verbrauchsart Heizen", "choice_h"),
    ("qh", "Qh [kW]", "float"),
    ("glz_h", "GLZ Heizen", "float"),
    ("kreis_h", "Kreis Heizen", "text"),
    ("bemerkung", "Bemerkung", "text"),
]

PRODUCT_BASE_FIELDS = [
    ("baut", "Baut.", "text"),
    ("gesch", "Ebene", "text"),
    ("versorgung", "Versorgung", "text"),
    ("baustufe", "Baustufe", "text"),
    ("r_nr", "R.-Nr.", "text"),
    ("produkt", "Produkt", "text"),
    ("verbrauchsart", "Verbrauchsart", "choice"),
    ("prozess", "Prozess", "process"),
    ("prod_zeit", "Prod.-Zeit [h/d]", "float"),
    ("ton_p_tag", "t/Tag", "float"),
    ("kg_batch", "kg/Batch", "float"),
    ("verweilzeit", "Verweilzeit [h]", "float"),
    ("t1", "t1 [°C]", "float"),
    ("t2", "t2 [°C]", "float"),
    ("sf", "SF", "float"),
    ("kreis", "Kreis", "text"),
    ("bemerkung", "Bemerkung", "text"),
]

# Produktdaten (cp1 oberhalb Gefrierpunkt, cp2 unterhalb, r = Kristallisationswärme)
PRODUCT_DATA = {
    "cp1": 3.5,      # kJ/(kg·K) oberhalb 0 °C (Fleisch ~3.5)
    "cp2": 1.8,      # kJ/(kg·K) unterhalb Gefrierpunkt
    "r": 250.0,      # kJ/kg Gefrierenthalpie
    "t_freeze": -2.0,
}


@dataclass
class Project:
    meta: Dict[str, Any] = field(default_factory=lambda: {
        "projekt_nr": "",
        "projekt_name": "",
        "projekt_phase": "Masterplan",
        "projekt_status": "in Bearbeitung",
        "standort": "",
        "firma": "Eckert & Partner beratende Ingenieure mbB",
        "datum": "",
        "bearbeiter": "",
    })
    rooms: List[Dict[str, Any]] = field(default_factory=list)
    consumers: List[Dict[str, Any]] = field(default_factory=list)
    products: List[Dict[str, Any]] = field(default_factory=list)
    circuits: List[Dict[str, Any]] = field(default_factory=lambda: json.loads(json.dumps(DEFAULT_CIRCUITS)))
    lists: Dict[str, List[str]] = field(default_factory=lambda: {
        "verbrauchsarten": list(DEFAULT_CONSUMER_TYPES),
        "heizarten": list(DEFAULT_HEATING_TYPES),
        "prozesse": list(DEFAULT_PROCESSES),
    })
    # Eigene Spalten: pro Tab Key -> Liste (name, label, typ)
    custom_fields: Dict[str, List[List[str]]] = field(default_factory=lambda: {
        "rooms": [], "consumers": [], "products": [],
    })

    # ---- Persistenz -------------------------------------------------
    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": VERSION,
            "meta": self.meta,
            "rooms": self.rooms,
            "consumers": self.consumers,
            "products": self.products,
            "circuits": self.circuits,
            "lists": self.lists,
            "custom_fields": self.custom_fields,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Project":
        p = cls()
        p.meta.update(data.get("meta", {}))
        p.rooms = data.get("rooms", [])
        p.consumers = data.get("consumers", [])
        p.products = data.get("products", [])
        circuits = data.get("circuits")
        if circuits:
            p.circuits = circuits
        p.lists.update(data.get("lists", {}))
        p.custom_fields.update(data.get("custom_fields", {}))
        return p

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self.to_dict(), fh, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "Project":
        with open(path, "r", encoding="utf-8") as fh:
            return cls.from_dict(json.load(fh))

    # ---- Erweiterbarkeit --------------------------------------------
    def add_custom_field(self, tab: str, name: str, label: str, typ: str) -> None:
        if not name or any(f[0] == name for f in self.custom_fields.get(tab, [])):
            raise ValueError("Spaltenname fehlt oder existiert bereits.")
        self.custom_fields.setdefault(tab, []).append([name, label, typ])
        for row in getattr(self, {"rooms": "rooms", "consumers": "consumers", "products": "products"}[tab]):
            row.setdefault(name, None)

    def remove_custom_field(self, tab: str, name: str) -> None:
        self.custom_fields[tab] = [f for f in self.custom_fields.get(tab, []) if f[0] != name]
        rows = {"rooms": self.rooms, "consumers": self.consumers, "products": self.products}[tab]
        for row in rows:
            row.pop(name, None)

    def circuit_ids(self) -> List[str]:
        return [c["id"] for c in self.circuits]


def new_room(**kwargs) -> Dict[str, Any]:
    room = {name: None for name, _, _ in ROOM_BASE_FIELDS}
    room.update(kwargs)
    return room


def new_consumer(**kwargs) -> Dict[str, Any]:
    c = {name: None for name, _, _ in CONSUMER_BASE_FIELDS}
    c.update(kwargs)
    return c


def new_product(**kwargs) -> Dict[str, Any]:
    p = {name: None for name, _, _ in PRODUCT_BASE_FIELDS}
    p.update(kwargs)
    return p
