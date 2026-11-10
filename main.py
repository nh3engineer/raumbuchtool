"""Raumbuch-Tool für kältetechnische Berechnungen.

Tkinter-Desktop-Oberfläche zum Raumprogramm (Rev8-Logik):
  * Projekt-, Raum-, Verbraucher-, Produkt- und Kreisverwaltung
  * Automatische Berechnung q0/Q0/q1/Qh, GLZ, Prozesskühlung, Kreissummen
  * Individuell erweiterbar: eigene Spalten, eigene Kreise,
    eigene Verbrauchsarten/Prozesse
  * Import aus dem Rev8-Raumbuch, Export nach Excel/CSV, Speichern als JSON

Start:  python main.py
"""
from __future__ import annotations

import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional

import calc
import excel_io
from model import (
    CONSUMER_BASE_FIELDS,
    PRODUCT_BASE_FIELDS,
    ROOM_BASE_FIELDS,
    Project,
    new_consumer,
    new_product,
    new_room,
)

FIELD_TYPES = {"text", "float", "choice", "choice_h", "process"}


class EditableTable(ttk.Frame):
    """Baumtabelle mit Doppelklick-Bearbeitung und berechneten Spalten."""

    def __init__(self, master, app: "App", tab_key: str,
                 base_fields, calc_columns: Optional[List] = None,
                 new_row_factory: Optional[Callable] = None):
        super().__init__(master)
        self.app = app
        self.tab_key = tab_key
        self.base_fields = base_fields
        self.calc_columns = calc_columns or []
        self.new_row_factory = new_row_factory
        self._sort_state: Dict[str, bool] = {}

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(toolbar, text="+ Zeile", command=self.add_row).pack(side="left", padx=2)
        ttk.Button(toolbar, text="– Zeile", command=self.delete_row).pack(side="left", padx=2)
        ttk.Button(toolbar, text="⧉ Duplizieren", command=self.duplicate_row).pack(side="left", padx=2)
        self.info_var = tk.StringVar()
        ttk.Label(toolbar, textvariable=self.info_var).pack(side="right")

        cols = [f[0] for f in self.fields] + [c[0] for c in self.calc_columns]
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        for name, label, typ in self.fields:
            self.tree.heading(name, text=label,
                              command=lambda n=name: self.sort_by(n))
            self.tree.column(name, width=90, anchor="w")
        for name, label, _fmt in self.calc_columns:
            self.tree.heading(name, text=label,
                              command=lambda n=name: self.sort_by(n))
            self.tree.column(name, width=90, anchor="e")
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Return>", lambda e: self.on_double_click(e))
        self._entry: Optional[tk.Widget] = None

    @property
    def fields(self):
        return self.base_fields + self.app.project.custom_fields.get(self.tab_key, [])

    def rows(self) -> List[Dict[str, Any]]:
        return getattr(self.app.project, self.tab_key)

    def refresh(self) -> None:
        # Spalten können sich geändert haben (eigene Spalten)
        cols = [f[0] for f in self.fields] + [c[0] for c in self.calc_columns]
        if list(self.tree["columns"]) != cols:
            self.tree["columns"] = cols
            for name, label, typ in self.fields:
                self.tree.heading(name, text=label, command=lambda n=name: self.sort_by(n))
                self.tree.column(name, width=90, anchor="w")
            for name, label, _f in self.calc_columns:
                self.tree.heading(name, text=label, command=lambda n=name: self.sort_by(n))
                self.tree.column(name, width=90, anchor="e")
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows()):
            vals = []
            for name, _l, typ in self.fields:
                vals.append("" if row.get(name) is None else row.get(name))
            for name, _l, fmt in self.calc_columns:
                vals.append(fmt(row))
            self.tree.insert("", "end", iid=str(i), values=vals)
        self.info_var.set(f"{len(self.rows())} Zeilen")

    def sort_by(self, col: str) -> None:
        reverse = self._sort_state.get(col, False)
        def key(i):
            v = self.rows()[int(i)].get(col)
            try:
                return (0, float(v))
            except (TypeError, ValueError):
                return (1, str(v or ""))
        rows = self.rows()
        order = sorted(range(len(rows)), key=lambda i: key(i), reverse=reverse)
        new_rows = [rows[i] for i in order]
        # Zugehörige Liste im Projekt ersetzen
        setattr(self.app.project, self.tab_key, new_rows)
        self._sort_state[col] = not reverse
        self.refresh()

    def add_row(self) -> None:
        if self.new_row_factory:
            self.rows().append(self.new_row_factory())
        else:
            self.rows().append({})
        self.refresh()
        self.tree.selection_set(str(len(self.rows()) - 1))
        self.tree.see(str(len(self.rows()) - 1))

    def delete_row(self) -> None:
        sel = self.tree.selection()
        if sel and messagebox.askyesno("Löschen", "Zeile wirklich löschen?"):
            for iid in sel:
                del self.rows()[int(iid)]
            self.refresh()

    def duplicate_row(self) -> None:
        sel = self.tree.selection()
        if sel:
            import copy
            for iid in sel:
                self.rows().append(copy.deepcopy(self.rows()[int(iid)]))
            self.refresh()

    # ---- Zellbearbeitung --------------------------------------------
    def on_double_click(self, event) -> None:
        if self._entry:
            self._commit()
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        col_name = self.tree["columns"][int(col.replace("#", "")) - 1]
        field = next((f for f in self.fields if f[0] == col_name), None)
        if not field:
            return  # berechnete Spalten sind nicht editierbar
        x, y, width, height = self.tree.bbox(iid, col)
        typ = field[2]
        current = self.rows()[int(iid)].get(col_name) or ""
        if typ == "choice":
            values = self.app.project.lists["verbrauchsarten"]
            widget: tk.Widget = ttk.Combobox(self.tree, values=values)
        elif typ == "choice_h":
            widget = ttk.Combobox(self.tree, values=self.app.project.lists["heizarten"])
        elif typ == "process":
            widget = ttk.Combobox(self.tree, values=self.app.project.lists["prozesse"])
        elif typ == "float":
            widget = ttk.Entry(self.tree)
            widget.insert(0, "" if current in (None, "") else str(current))
        else:
            widget = ttk.Entry(self.tree)
            widget.insert(0, str(current))
        self._entry = widget
        self._edit = (int(iid), col_name, typ)
        widget.place(x=x, y=y, width=max(width, 60), height=height)
        widget.focus_set()
        if isinstance(widget, ttk.Entry):
            widget.select_range(0, "end")
        widget.bind("<FocusOut>", lambda e: self._commit())
        widget.bind("<Return>", lambda e: self._commit())
        widget.bind("<Escape>", lambda e: self._destroy_entry())

    def _destroy_entry(self) -> None:
        if self._entry:
            self._entry.destroy()
            self._entry = None

    def _commit(self) -> None:
        if not self._entry:
            return
        row_i, col_name, typ = self._edit
        raw = self._entry.get()
        self._destroy_entry()
        value: Any = raw
        if typ == "float":
            try:
                value = float(raw.replace(",", ".")) if raw.strip() else None
            except ValueError:
                messagebox.showwarning("Eingabe", f"„{raw}“ ist keine Zahl.")
                return
        self.rows()[row_i][col_name] = value
        self.refresh()
        self.app.mark_dirty()


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Raumbuch-Tool Kältetechnik")
        self.geometry("1400x800")
        self.project = Project()
        self.project_path: Optional[str] = None
        self._dirty = False

        self._build_menu()
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=6, pady=6)
        # Branding-Kopfzeile (wie im Raumbuch: Logo + Firmenname + Projektzeile)
        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 6))
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")
        if os.path.exists(logo_path):
            self._logo_img = tk.PhotoImage(file=logo_path)
            try:
                self._logo_img = self._logo_img.subsample(4)
            except tk.TclError:
                pass
            ttk.Label(header, image=self._logo_img).pack(side="left")
        right = ttk.Frame(header)
        right.pack(side="right")
        self.header_firm = tk.StringVar()
        self.header_line = tk.StringVar()
        ttk.Label(right, textvariable=self.header_firm,
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="e")
        ttk.Label(right, textvariable=self.header_line).pack(anchor="e")
        nb = ttk.Notebook(main)
        nb.pack(fill="both", expand=True)

        self.tab_project = self._build_project_tab(nb)
        self.tab_rooms = self._build_rooms_tab(nb)
        self.tab_consumers = self._build_consumers_tab(nb)
        self.tab_products = self._build_products_tab(nb)
        self.tab_circuits = self._build_circuits_tab(nb)
        self.tab_total = self._build_total_tab(nb)
        self.tab_settings = self._build_settings_tab(nb)

        self.refresh_all()

    # ---- Menü --------------------------------------------------------
    def _build_menu(self) -> None:
        m = tk.Menu(self)
        self.config(menu=m)
        fm = tk.Menu(m, tearoff=0)
        m.add_cascade(label="Datei", menu=fm)
        fm.add_command(label="Neu", command=self.new_project)
        fm.add_command(label="Öffnen (JSON)…", command=self.load_project)
        fm.add_command(label="Speichern", command=self.save_project, accelerator="Strg+S")
        fm.add_command(label="Speichern unter…", command=self.save_project_as)
        fm.add_separator()
        fm.add_command(label="Raumbuch Rev8 importieren…", command=self.import_rev8)
        fm.add_separator()
        fm.add_command(label="Beenden", command=self.on_close)
        em = tk.Menu(m, tearoff=0)
        m.add_cascade(label="Export", menu=em)
        em.add_command(label="Excel-Export…", command=self.export_excel)
        em.add_command(label="CSV-Export (Kreissummen)…", command=self.export_csv)
        self.bind("<Control-s>", lambda e: self.save_project())

    # ---- Tabs --------------------------------------------------------
    def _build_project_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text="Projekt")
        self.proj_vars: Dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate([
            ("projekt_nr", "Projekt-Nr."), ("projekt_name", "Projektname"),
            ("projekt_phase", "Projektphase"), ("projekt_status", "Projektstatus"),
            ("standort", "Standort"), ("firma", "Firma"),
            ("datum", "Datum"), ("bearbeiter", "Bearbeiter"),
        ]):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(self.project.meta.get(key, "")))
            var.trace_add("write", self._on_meta_changed)
            ttk.Entry(f, textvariable=var, width=50).grid(row=i, column=1, sticky="w", pady=3)
            self.proj_vars[key] = var
        return f

    def _build_rooms_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb)
        nb.add(f, text="Räume")
        calc_cols = [
            ("c_vol", "V [m³]", lambda r: calc.room_volume(r) or ""),
            ("c_q0", "q0 [W/m²]", lambda r: round(calc.room_q0(r), 2)),
            ("c_q0kw", "Q0 [kW]", lambda r: round(calc.room_q0_kw(r), 3)),
            ("c_q1", "q1 [W/m³]", lambda r: round(calc.room_q1(r), 2) if calc.room_q1(r) is not None else ""),
            ("c_qh", "qh [W/m²]", lambda r: round(calc.room_qh(r), 2)),
            ("c_qhkw", "Qh [kW]", lambda r: round(calc.room_qh_kw(r), 3)),
        ]
        self.table_rooms = EditableTable(f, self, "rooms", ROOM_BASE_FIELDS, calc_cols, new_room)
        self.table_rooms.pack(fill="both", expand=True)
        return f

    def _build_consumers_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb)
        nb.add(f, text="Verbraucher")
        calc_cols = [
            ("c_q0e", "Q0 eff. [kW]", lambda r: round(calc.consumer_q0_kw(r), 3)),
            ("c_qhe", "Qh eff. [kW]", lambda r: round(calc.consumer_qh_kw(r), 3)),
        ]
        self.table_consumers = EditableTable(f, self, "consumers", CONSUMER_BASE_FIELDS, calc_cols, new_consumer)
        self.table_consumers.pack(fill="both", expand=True)
        return f

    def _build_products_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb)
        nb.add(f, text="Produkte")
        calc_cols = [
            ("c_m", "kg/h", lambda r: round(calc.product_mass_flow(r) or 0, 1)),
            ("c_q0p", "Q0_P [kW]", lambda r: round(calc.product_q0_kw(r) or 0, 3)),
        ]
        self.table_products = EditableTable(f, self, "products", PRODUCT_BASE_FIELDS, calc_cols, new_product)
        self.table_products.pack(fill="both", expand=True)
        return f

    def _build_circuits_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb)
        nb.add(f, text="Kreise")
        cols = ("id", "medium", "t0", "ts_vl", "ts_rl", "dp", "k", "q0", "qh")
        headers = ["Kreis", "Medium", "t0 [°C]", "ts VL [°C]", "ts RL [°C]",
                   "ΔP [K]", "k [W/m²K]", "Q0 [kW]", "Qh [kW]"]
        toolbar = ttk.Frame(f)
        toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(toolbar, text="+ Kreis", command=self.add_circuit).pack(side="left", padx=2)
        ttk.Button(toolbar, text="– Kreis", command=self.delete_circuit).pack(side="left", padx=2)
        self.tree_circuits = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c, h in zip(cols, headers):
            self.tree_circuits.heading(c, text=h)
            self.tree_circuits.column(c, width=80, anchor="center")
        self.tree_circuits.pack(fill="both", expand=True)
        self.tree_circuits.bind("<Double-1>", self.edit_circuit)
        self.tree_circuits.bind("<Return>", self.edit_circuit)
        return f

    def _build_total_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb)
        nb.add(f, text="Gesamtbedarf")
        cols = ("id", "medium", "t0", "q0", "qh", "share")
        headers = ["Kreis", "Medium", "t0 [°C]", "Q0 [kW]", "Qh [kW]", "Anteil Q0 [%]"]
        self.tree_total = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c, h, w in zip(cols, headers, (70, 70, 70, 100, 100, 90)):
            self.tree_total.heading(c, text=h)
            self.tree_total.column(c, width=w, anchor="center")
        self.tree_total.pack(fill="both", expand=True)
        ttk.Button(f, text="Aktualisieren", command=self.refresh_all).pack(pady=6)
        return f

    def _build_settings_tab(self, nb) -> ttk.Frame:
        f = ttk.Frame(nb, padding=8)
        nb.add(f, text="Erweiterung")
        # Eigene Spalten
        g1 = ttk.LabelFrame(f, text="Eigene Spalten", padding=8)
        g1.pack(fill="x", pady=4)
        ttk.Label(g1, text="Tabelle:").grid(row=0, column=0)
        self.cf_tab = tk.StringVar(value="rooms")
        ttk.Combobox(g1, textvariable=self.cf_tab, state="readonly",
                     values=[("rooms", "Räume"), ("consumers", "Verbraucher"), ("products", "Produkte")],
                     width=14).grid(row=0, column=1)
        self.cf_tab.trace_add("write", lambda *_: self.refresh_custom_fields())
        ttk.Label(g1, text="Name (key):").grid(row=1, column=0, sticky="w")
        self.cf_name = ttk.Entry(g1, width=16); self.cf_name.grid(row=1, column=1)
        ttk.Label(g1, text="Beschriftung:").grid(row=2, column=0, sticky="w")
        self.cf_label = ttk.Entry(g1, width=16); self.cf_label.grid(row=2, column=1)
        ttk.Label(g1, text="Typ:").grid(row=3, column=0, sticky="w")
        self.cf_type = tk.StringVar(value="text")
        ttk.Combobox(g1, textvariable=self.cf_type, state="readonly",
                     values=["text", "float"], width=14).grid(row=3, column=1)
        ttk.Button(g1, text="Spalte hinzufügen", command=self.add_custom_field).grid(
            row=4, column=0, columnspan=2, pady=(6, 0))
        self.tree_cf = ttk.Treeview(g1, columns=("name", "label", "type"),
                                    show="headings", height=4)
        for c, h in zip(("name", "label", "type"), ("Name", "Beschriftung", "Typ")):
            self.tree_cf.heading(c, text=h)
            self.tree_cf.column(c, width=110)
        self.tree_cf.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(g1, text="Spalte entfernen", command=self.remove_custom_field).grid(
            row=6, column=0, columnspan=2, pady=4)

        # Listen (Verbrauchsarten etc.)
        g2 = ttk.LabelFrame(f, text="Wertelisten (Dropdowns)", padding=8)
        g2.pack(fill="both", expand=True, pady=4)
        ttk.Label(g2, text="Liste:").grid(row=0, column=0, sticky="w")
        self.list_key = tk.StringVar(value="verbrauchsarten")
        ttk.Combobox(g2, textvariable=self.list_key, state="readonly",
                     values=["verbrauchsarten", "heizarten", "prozesse"],
                     width=18).grid(row=0, column=1)
        self.list_key.trace_add("write", lambda *_: self.refresh_lists())
        self.list_entry = ttk.Entry(g2, width=24)
        self.list_entry.grid(row=1, column=0, pady=(6, 0))
        ttk.Button(g2, text="Hinzufügen", command=self.add_list_value).grid(row=1, column=1, sticky="w", padx=4)
        self.list_box = tk.Listbox(g2, height=8)
        self.list_box.grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Button(g2, text="Entfernen", command=self.remove_list_value).grid(row=2, column=1, sticky="w")
        return f

    # ---- Aktionen ----------------------------------------------------
    def _on_meta_changed(self, *_args) -> None:
        self._write_meta()
        self.mark_dirty()
        self._refresh_header()

    def _refresh_header(self) -> None:
        if not hasattr(self, "header_firm"):
            return
        self.header_firm.set(str(self.project.meta.get("firma", "")))
        parts = [self.project.meta.get(k, "") for k in
                 ("projekt_nr", "projekt_name", "projekt_phase", "datum", "projekt_status")]
        self.header_line.set(" | ".join(str(p) for p in parts if p))

    def mark_dirty(self) -> None:
        self._dirty = True
        self.title("Raumbuch-Tool Kältetechnik – *" + (os.path.basename(self.project_path) if self.project_path else "ungespeichert"))
        self.refresh_totals()

    def refresh_all(self) -> None:
        for key, var in self.proj_vars.items():
            var.set(str(self.project.meta.get(key, "")))
        self._refresh_header()
        self.table_rooms.refresh()
        self.table_consumers.refresh()
        self.table_products.refresh()
        self.refresh_circuits()
        self.refresh_totals()
        self.refresh_custom_fields()
        self.refresh_lists()

    def refresh_circuits(self) -> None:
        totals = calc.compute_totals(self.project)
        self.tree_circuits.delete(*self.tree_circuits.get_children())
        for c in self.project.circuits:
            t = totals.get(c["id"], {"q0": 0, "qh": 0})
            self.tree_circuits.insert("", "end", values=(
                c["id"], c.get("medium"), c.get("t0"), c.get("ts_vl"),
                c.get("ts_rl"), c.get("dp"), c.get("k"),
                round(t["q0"], 3), round(t["qh"], 3)))

    def refresh_totals(self) -> None:
        if not hasattr(self, "tree_total"):
            return
        totals = calc.compute_totals(self.project)
        total_q0 = totals["SUMME"]["q0"] or 0.0
        self.tree_total.delete(*self.tree_total.get_children())
        for c in self.project.circuits:
            t = totals.get(c["id"], {"q0": 0, "qh": 0})
            share = round(t["q0"] / total_q0 * 100, 1) if total_q0 else 0
            self.tree_total.insert("", "end", values=(
                c["id"], c.get("medium"), c.get("t0"),
                round(t["q0"], 3), round(t["qh"], 3), share))
        self.tree_total.insert("", "end", values=(
            "SUMME", "", "", round(totals["SUMME"]["q0"], 3),
            round(totals["SUMME"]["qh"], 3), 100.0))

    def add_circuit(self) -> None:
        win = tk.Toplevel(self)
        win.title("Kreis hinzufügen")
        fields = [("id", "Kreis (z. B. LP3)"), ("medium", "Medium"),
                  ("t0", "t0 [°C]"), ("ts_vl", "ts VL [°C]"), ("ts_rl", "ts RL [°C]"),
                  ("dp", "ΔP [K]"), ("k", "k [W/m²K]")]
        vars_: Dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate(fields):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="w", padx=6, pady=2)
            v = tk.StringVar(value="")
            ttk.Entry(win, textvariable=v).grid(row=i, column=1, padx=6, pady=2)
            vars_[key] = v
        def ok():
            cid = vars_["id"].get().strip()
            if not cid or cid in self.project.circuit_ids():
                messagebox.showwarning("Kreis", "Kreis-ID fehlt oder existiert bereits.")
                return
            def num(k, d=None):
                try:
                    return float(vars_[k].get().replace(",", "."))
                except ValueError:
                    return d
            self.project.circuits.append({
                "id": cid, "medium": vars_["medium"].get() or "NH3",
                "t0": num("t0"), "ts_vl": num("ts_vl"), "ts_rl": num("ts_rl"),
                "dp": num("dp", 0.0), "k": num("k", 18.0)})
            win.destroy()
            self.mark_dirty()
            self.refresh_circuits()
        ttk.Button(win, text="Übernehmen", command=ok).grid(row=len(fields), column=0, columnspan=2, pady=6)

    def delete_circuit(self) -> None:
        sel = self.tree_circuits.selection()
        if sel:
            vals = self.tree_circuits.item(sel[0], "values")
            if messagebox.askyesno("Kreis", f"Kreis {vals[0]} löschen?"):
                self.project.circuits = [c for c in self.project.circuits if c["id"] != vals[0]]
                self.mark_dirty()
                self.refresh_circuits()

    def edit_circuit(self, event) -> None:
        sel = self.tree_circuits.selection()
        if not sel:
            return
        idx = self.tree_circuits.index(sel[0])
        circuit = self.project.circuits[idx]
        col = self.tree_circuits.identify_column(event.x)
        if not col:
            return
        key = ("id", "medium", "t0", "ts_vl", "ts_rl", "dp", "k")[int(col.replace("#", "")) - 1]
        x, y, w, h = self.tree_circuits.bbox(sel[0], col)
        entry = ttk.Entry(self.tree_circuits)
        entry.insert(0, str(circuit.get(key) or ""))
        entry.place(x=x, y=y, width=max(w, 60), height=h)
        entry.focus_set(); entry.select_range(0, "end")
        def commit(_e=None):
            raw = entry.get()
            if key in ("t0", "ts_vl", "ts_rl", "dp", "k"):
                try:
                    circuit[key] = float(raw.replace(",", ".")) if raw else None
                except ValueError:
                    messagebox.showwarning("Eingabe", f"„{raw}“ ist keine Zahl.")
            else:
                circuit[key] = raw
            entry.destroy()
            self.mark_dirty()
            self.refresh_circuits()
        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    def refresh_custom_fields(self) -> None:
        if not hasattr(self, "tree_cf"):
            return
        self.tree_cf.delete(*self.tree_cf.get_children())
        for name, label, typ in self.project.custom_fields.get(self.cf_tab.get(), []):
            self.tree_cf.insert("", "end", values=(name, label, typ))

    def add_custom_field(self) -> None:
        try:
            self.project.add_custom_field(self.cf_tab.get(), self.cf_name.get().strip(),
                                          self.cf_label.get().strip() or self.cf_name.get().strip(),
                                          self.cf_type.get())
        except ValueError as e:
            messagebox.showwarning("Eigene Spalte", str(e))
            return
        self.cf_name.delete(0, "end"); self.cf_label.delete(0, "end")
        self.mark_dirty()
        self.refresh_custom_fields()
        self.table_rooms.refresh(); self.table_consumers.refresh(); self.table_products.refresh()

    def remove_custom_field(self) -> None:
        sel = self.tree_cf.selection()
        if sel:
            name = self.tree_cf.item(sel[0], "values")[0]
            self.project.remove_custom_field(self.cf_tab.get(), name)
            self.mark_dirty()
            self.refresh_custom_fields()
            self.table_rooms.refresh(); self.table_consumers.refresh(); self.table_products.refresh()

    def refresh_lists(self) -> None:
        if not hasattr(self, "list_box"):
            return
        self.list_box.delete(0, "end")
        for v in self.project.lists.get(self.list_key.get(), []):
            self.list_box.insert("end", v)

    def add_list_value(self) -> None:
        v = self.list_entry.get().strip()
        if v:
            lst = self.project.lists.setdefault(self.list_key.get(), [])
            if v not in lst:
                lst.append(v)
            self.list_entry.delete(0, "end")
            self.mark_dirty(); self.refresh_lists()

    def remove_list_value(self) -> None:
        sel = self.list_box.curselection()
        if sel:
            del self.project.lists[self.list_key.get()][sel[0]]
            self.mark_dirty(); self.refresh_lists()

    # ---- Datei-Operationen -------------------------------------------
    def new_project(self) -> None:
        if self._dirty and not messagebox.askyesno("Neu", "Ungespeicherte Änderungen verwerfen?"):
            return
        self.project = Project()
        self.project_path = None
        self._dirty = False
        self.refresh_all()

    def load_project(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Raumbuch-Projekt", "*.json")])
        if not path:
            return
        self.project = Project.load(path)
        self.project_path = path
        self._dirty = False
        self.refresh_all()

    def save_project(self) -> None:
        if self.project_path:
            self._write_meta()
            self.project.save(self.project_path)
            self._dirty = False
            self.title("Raumbuch-Tool Kältetechnik")
        else:
            self.save_project_as()

    def save_project_as(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".json",
                                            filetypes=[("Raumbuch-Projekt", "*.json")])
        if path:
            self.project_path = path
            self.save_project()

    def _write_meta(self) -> None:
        for key, var in self.proj_vars.items():
            self.project.meta[key] = var.get()

    def import_rev8(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            imported = excel_io.import_rev8(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Import", f"Import fehlgeschlagen:\n{e}")
            return
        if not (imported.rooms or imported.consumers or imported.products):
            messagebox.showwarning("Import", "Keine Datenzeilen gefunden.")
            return
        if not self.project.rooms and not self.project.consumers and not self.project.products:
            self.project.rooms = imported.rooms
            self.project.consumers = imported.consumers
            self.project.products = imported.products
        else:
            if messagebox.askyesno("Import", "Vorhandene Daten ersetzen? (Nein = anhängen)"):
                self.project.rooms = imported.rooms
                self.project.consumers = imported.consumers
                self.project.products = imported.products
            else:
                self.project.rooms += imported.rooms
                self.project.consumers += imported.consumers
                self.project.products += imported.products
        messagebox.showinfo(
            "Import",
            f"Importiert: {len(imported.rooms)} Räume, "
            f"{len(imported.consumers)} Verbraucher, {len(imported.products)} Produkte")
        self.mark_dirty()
        self.refresh_all()

    def export_excel(self) -> None:
        self._write_meta()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if path:
            try:
                excel_io.export_excel(self.project, path)
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Export", f"Export fehlgeschlagen:\n{e}")
                return
            messagebox.showinfo("Export", f"Exportiert nach:\n{path}")

    def export_csv(self) -> None:
        self._write_meta()
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")])
        if path:
            excel_io.export_csv(self.project, path)
            messagebox.showinfo("Export", f"Exportiert nach:\n{path}")

    def on_close(self) -> None:
        if self._dirty and not messagebox.askyesno("Beenden", "Ungespeicherte Änderungen verwerfen?"):
            return
        self.destroy()



def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
