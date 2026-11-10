# Raumbuch-Tool Kältetechnik

Desktop-Oberfläche für kältetechnische Berechnungen nach der Logik des
Raumprogramms (Rev8): Räume, weitere Verbraucher, Produkte (Prozesskühlung),
Kreise und Gesamt-Kältebedarf – inklusive Eckert-&-Partner-Branding.

## Start

```bash
pip install openpyxl          # einzige Abhängigkeit
python main.py
```

## Funktionen

- **Projekt**: Projektkopf (Nr., Name, Phase, Status, Standort, Firma, Datum, Bearbeiter)
- **Räume (Raumbuch)**: R.-Nr., Name, Fläche/Höhe → Volumen, Solltemperaturen,
  Verbrauchsart, q0 berechnet/Zuschlag/Produkt → automatische Berechnung von
  q0 [W/m²], Q0 [kW], q1 [W/m³] sowie qh/Qh (Heizen) und Kreiszordnung
- **Verbraucher**: Q0/Qh mit Gleichzeitigkeitsfaktor (GLZ) je Kreis
- **Produkte**: Prozesskühlung CLG (Kühlen) / FRZ (Einfrieren) / DEF (Auftauen)
  aus t/Tag, Prod.-Zeit, t1/t2 und Sicherheitsfaktor SF
- **Kreise**: Medien (NH3, PG, H2O …), t0, ts VL/RL, ΔP, k – frei ergänzbar
- **Gesamtbedarf**: Summen Q0/Qh je Kreis und gesamt, Anteil je Kreis
- **Erweiterung**: eigene Spalten je Tabelle, eigene Kreise, eigene
  Verbrauchsarten/Heizarten/Prozesse für die Dropdowns

## Datei

- **Import**: „Datei → Raumbuch Rev8 importieren…“ liest Room/Consumer/Produkt
- **Speichern**: Projekt als JSON (alle Daten, eigene Spalten, Listen)
- **Export**: Excel (Projekt, Räume, Verbraucher, Produkte, Kreise) und
  CSV (Kreissummen); Excel-Export mit Logo und Projektzeile im Kopf

## Berechnungsformeln

```
Raum:   V = A·H | q0 = q0_ber+q0_zusch+q0_prod | Q0 = A·q0/1000 | q1 = Q0·1000/V
Verbraucher: Q0_eff = Q0 · GLZ
Produkt: m = t/Tag·1000/Prod.-Zeit
         CLG: Q0 = m·cp1·(t1−t2)/3600 · (1+SF)
         FRZ: Q0 = m·(cp1·(t1−tF)+r+cp2·(tF−t2))/3600 · (1+SF)
         DEF: Q0 = m·(cp2·(t1−tF)+r+cp1·(tF−t2))/3600 · (1+SF)
```

Produktdaten (cp1, cp2, r, Gefrierpunkt) stehen in `model.py` (PRODUCT_DATA)
und können je Produktart angepasst werden.

## Dateien

| Datei | Inhalt |
|---|---|
| `main.py` | Tkinter-Oberfläche (Tabs, Bearbeitung, Import/Export) |
| `model.py` | Datenmodell, Kreise, Wertelisten, eigene Spalten |
| `calc.py` | Berechnungs-Engine |
| `excel_io.py` | Rev8-Import, Excel-/CSV-Export |
| `assets/` | Logo (GUI + Excel-Export) |
