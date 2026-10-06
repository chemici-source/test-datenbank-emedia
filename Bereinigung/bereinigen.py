"""
Bereinigen einer E-Medien-Statistik-Datei
==========================================

Anwendung (im Terminal, im Ordner mit dieser Datei):

    python bereinigen.py "Pfad\\zur\\Original-Datei.xlsx"

Ergebnis:
    Im selben Ordner entsteht "<Dateiname>_bereinigt.xlsx". Pro verarbeitetem
    Blatt gibt es ein Blatt "Bereinigt_<Blattname>", dazu ein gemeinsames
    Blatt "Protokoll" mit allen Aenderungen/Maskierungen/Kategorisierungen.

Welche Konfiguration verwendet wird, entscheidet das Skript anhand des
Dateinamens (siehe CONFIGS unten). Fuer jedes neue Dateiformat kommt ein
neuer Eintrag dazu. Nur die dort unter "sheets" gelisteten Blaetter werden
ueberhaupt eingelesen, alle anderen Blaetter der Originaldatei (z.B. ein
Kontakt-Blatt mit Personendaten) werden nie geoeffnet.

Prinzip pro Spalte:
- Exakter Spaltenname in "whitelist"      -> Regel wird direkt angewendet.
- Sonst: Name gegen "muster_regeln" (Liste von Regex-Mustern) geprueft,
  erstes Match gewinnt. Damit lassen sich Spalten wie "Preis 2025 in CHF"
  oder "Kosten CHF 2025" erfassen, ohne jedes Jahr einzeln aufzufuehren.
- Sonst: Spalte gilt als unbekannt -> wird komplett maskiert und geloggt.
- Freitext-Spalten (z.B. Notizen/Kommentare) werden nie inhaltlich
  geprueft, sondern unveraendert durchgelassen und im Protokoll als
  "manuell pruefen" markiert.

Aktionen, die eine Regel ausloesen kann:
- "durchlassen":     Wert bleibt, sofern er zum erwarteten Typ passt.
- "kategorisieren":  Betrag wird durch eine Preis-Bandbreite ersetzt
                      (inkl. Waehrung, negative Werte -> "Gutschrift/Korrektur").
"""

import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd

PLATZHALTER = "bitte in Originaldatei nachschauen"

# Bekannte "eigentlich leer"-Werte in den Originaldateien (z.B. ein einzelnes
# Leerzeichen statt einer leeren Zelle, oder "?" fuer "Preis unbekannt").
# Werden wie eine leere Zelle behandelt: still uebersprungen, keine
# Typ-Abweichung, kein Protokoll-Eintrag, keine monatliche Nachkontrolle noetig.
# Bei Bedarf hier ergaenzen, sobald weitere solche Platzhalter auftauchen.
BEKANNTE_LEERWERTE = {"", "?", "-", "--", "n/a", "n.a", "n.a.", "k.a.", "k.a", "kein treffer"}


def ist_bekannter_leerwert(wert) -> bool:
    return isinstance(wert, str) and wert.strip().lower() in BEKANNTE_LEERWERTE

# Bandbreiten, empirisch anhand aller bereits vorhandenen Langformat-Dateien
# hergeleitet (Kennzahl "Kosten CHF", n=1207, ohne negative Werte). Ergibt eine
# ausgewogene Verteilung: <500 CHF 28%, 500-2000 CHF 34%, 2000-10000 CHF 26%,
# 10000-50000 CHF 9%, >50000 CHF 3%.
PREIS_BANDBREITEN = [
    (500,          "unter 500"),
    (2000,         "500-2000"),
    (10000,        "2000-10000"),
    (50000,        "10000-50000"),
    (float("inf"), "über 50000"),
]

DURCHLASSEN_NUM = {"typ": "numerisch", "aktion": "durchlassen"}
DURCHLASSEN_TEXT = {"typ": "text_kategorie", "aktion": "durchlassen"}
# Fuer Spalten, die strukturell (nicht nur gelegentlich) Personendaten Dritter
# enthalten (z.B. Name/E-Mail/Telefon von Verlagskontakten) - anders als
# Freitext-Spalten wird hier IMMER maskiert, unabhaengig vom Zeileninhalt,
# da eine manuelle Zeile-fuer-Zeile-Kontrolle hier nicht ausreicht.
MASKIEREN = {"typ": None, "aktion": "maskieren"}


def kategorisieren_chf():
    return {"typ": "numerisch", "aktion": "kategorisieren", "waehrung": "CHF"}


def kategorisieren_auto():
    """Fuer Zellen, die den Waehrungscode direkt im Wert selbst tragen
    (z.B. '€ 818,67', '$ 850.00', 'Fr. 2\\'885.30'), statt in einer eigenen
    Spalte oder Gruppen-Kopfzeile. Waehrung wird pro Zelle erkannt."""
    return {"typ": "numerisch", "aktion": "kategorisieren_auto"}


def erkenne_waehrung_aus_text(wert) -> str:
    text = str(wert)
    if "€" in text:
        return "EUR"
    if "$" in text:
        return "USD"
    if "£" in text:
        return "GBP"
    if re.search(r"(?i)\bfr\.?", text):
        return "CHF"
    if re.search(r"(?i)\bchf\b", text):
        return "CHF"
    if re.search(r"(?i)\beur\b", text):
        return "EUR"
    if re.search(r"(?i)\busd\b", text):
        return "USD"
    if re.search(r"(?i)\bgbp\b", text):
        return "GBP"
    return "?"


def waehrung_aus_zahlenformat(format_code) -> str:
    """Manche Excel-Zellen zeigen die Waehrung nur ueber ein benutzerdefiniertes
    Zahlenformat an (z.B. '"CHF" #,##0.00'), der eigentliche Zellwert ist dann
    eine reine Zahl ohne jedes Waehrungszeichen. Pandas liest nur den Wert,
    nicht das Format - dafuer wird das Format separat per openpyxl gelesen."""
    if not format_code:
        return "?"
    fc = str(format_code).upper()
    if "CHF" in fc:
        return "CHF"
    if "EUR" in fc or "€" in fc:
        return "EUR"
    if "USD" in fc or "$" in fc:
        return "USD"
    if "GBP" in fc or "£" in fc:
        return "GBP"
    return "?"


def bestimme_waehrung(wert, zahlenformat=None) -> str:
    """Kombiniert Erkennung aus dem Zellwert selbst (Text mit Waehrungszeichen)
    mit Erkennung aus dem Excel-Zahlenformat (falls der Wert eine reine Zahl
    ohne eingebettetes Waehrungszeichen ist, z.B. per Zellformat als 'CHF'
    dargestellt)."""
    waehrung = erkenne_waehrung_aus_text(wert)
    if waehrung != "?":
        return waehrung
    return waehrung_aus_zahlenformat(zahlenformat)


def lade_zahlenformate(pfad, blattname: str, max_datenzeilen: int) -> pd.DataFrame:
    """Liest fuer ein Blatt die Excel-Zahlenformate der Datenzellen ein,
    positionell ausgerichtet wie pd.read_excel(...) es tun wuerde (Spalte
    0..n, Zeile 0..n ab der ersten Datenzeile).

    'max_datenzeilen' begrenzt das Lesen auf die von pandas tatsaechlich
    erkannte Zeilenzahl (+ etwas Puffer). Manche Excel-Dateien haben ein
    aufgeblaehtes 'benutztes' Blatt (Formatierungs-Reste weit ueber die
    echten Daten hinaus, z.B. bis Zeile 1'048'458) - ohne diese Grenze
    wuerde openpyxl Millionen leerer Zeilen einzeln einlesen und das Skript
    praktisch einfrieren."""
    wb = openpyxl.load_workbook(pfad, data_only=True)
    ws = wb[blattname]
    # Zeile 1 ist die Kopfzeile, danach maximal max_datenzeilen + Puffer lesen.
    max_zeile = 1 + max_datenzeilen + 20
    zeilen = list(ws.iter_rows(min_row=2, max_row=max_zeile))
    formate = [[zelle.number_format for zelle in zeile] for zeile in zeilen]
    return pd.DataFrame(formate)


# ---------------------------------------------------------------
# Konfigurationen pro Dateiformat. "muster" ist ein Teilstring, der im
# Dateinamen gesucht wird, um die passende Konfiguration automatisch
# auszuwaehlen (Gross/Kleinschreibung egal).
# ---------------------------------------------------------------

CONFIGS = {
    # ACHTUNG Namenskollision: "Nutzungsstatistiken Datenbanken NAT.xlsx" und
    # die neuere, eigenstaendige "Datenbanken NAT.xlsx" (Sachkonto 310203,
    # eigener Config-Eintrag weiter unten) sind ZWEI VERSCHIEDENE Dateien mit
    # unterschiedlichem Inhalt, die nur einen aehnlichen Namen teilen. Der
    # Config-Key und das "muster" sind hier bewusst auf den vollen Dateinamen
    # ausgeschrieben, damit "Datenbanken NAT.xlsx" (ohne "Nutzungsstatistiken"-
    # Praefix) nicht versehentlich hier landet.
    "Nutzungsstatistiken Datenbanken NAT": {
        "muster": "nutzungsstatistiken datenbanken nat",
        "sheets": {
            # Blatt "Kontakte" (echte E-Mail-Adressen von Verlagskontakten,
            # Personendaten) ist bewusst NICHT gelistet -> wird nie eingelesen.
            "Nutzungsdaten": {
                "whitelist": {
                    "Titel Datenbank": DURCHLASSEN_TEXT,
                    "Fach":            DURCHLASSEN_TEXT,
                    "Lieferant":       DURCHLASSEN_TEXT,
                    "Verlag":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen"},
                "muster_regeln": [
                    # Kosten CHF / Preis CHF: bekanntes Synonym (Datenglossar)
                    (r"^(Kosten CHF|Preis CHF) \d{4}$", kategorisieren_chf()),
                    (r"^(CPU UIR|CPU TIR|Searches|UIR|TIR|TII) \d{4}$", DURCHLASSEN_NUM),
                ],
                # Mitten im Blatt eingefuegte Abkuerzungs-Legende ("Legende
                # Metric_Type", "UIR= Unique_Item_Requests", ...) steht direkt
                # in der Titel-Spalte, keine echten Eintraege (gleiches Muster
                # wie bei "Nutzungsstatistiken EBook-Pakete NAT").
                "summenzeilen_entfernen": {
                    "spalte": "Titel Datenbank",
                    "praefixe": ["Legende", "UIR=", "TIR=", "UII=", "TII=", "Searches_Regular"],
                },
            },
        },
    },
    # Eigenstaendige, neuere Datei (Sachkonto 310203) - siehe Hinweis oben,
    # nicht zu verwechseln mit "Nutzungsstatistiken Datenbanken NAT.xlsx".
    "Datenbanken NAT": {
        "muster": "datenbanken nat",
        "sheets": {
            "310203": {
                "whitelist": {
                    "Datenbank":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis FW \d{4}\d*$", kategorisieren_auto()),
                    # "Preis in  CHF ..." (unregelmaessige Leerzeichen/NBSP,
                    # Fussnoten-Ziffern) und die einfache "Preis <jahr>"-
                    # Variante ohne "CHF" im Namen - Waehrung wird ueberall
                    # werteseitig erkannt, nicht ueber die Kopfzeile.
                    (r"^Preis\s+in\s+CHF\s*\d{4}\d*$", kategorisieren_auto()),
                    (r"^Preis \d{4}\d*$", kategorisieren_auto()),
                ],
                # Am Blattende angehaengter "Sachkonto/Gesamtkosten"-Block
                # (wie bei E-Book-Pakete NAT) ohne Eintrag in der Titel-Spalte.
                "titel_leer_entfernen": {"spalte": "Datenbank"},
            },
        },
    },
    "E-Journals B452": {
        "muster": "b452",
        "sheets": {
            "Einzeltitel": {
                "whitelist": {
                    "Journal": DURCHLASSEN_TEXT,
                    "ISSN":    DURCHLASSEN_TEXT,
                },
                # "Comments Michelle" nennt eine Person beim Vornamen (Kollegin) -
                # wird wie Freitext behandelt (unveraendert, manuell pruefen)
                # UND fuer die Weitergabe in einen neutralen Namen umbenannt.
                "freitext_spalten": {"Notizen B452", "Comments E-Library", "Comments Michelle"},
                "spalten_umbenennen": {"Comments Michelle": "Comments BBL"},
                "muster_regeln": [
                    (r"^Preis \d{4} in CHF$", kategorisieren_chf()),
                    (r"^(CPU|Uses) \d{4}$", DURCHLASSEN_NUM),
                ],
                # Zeilen 0-1 (nach Header) sind Teil des mehrzeiligen Original-Headers
                # ("tr_J1"-Marker, TIR/UIR-Unterzeile), keine echten Datenzeilen.
                "datenzeilen_ab": 2,
                # Zeile 1 (0-basiert, vor dem Abschneiden) enthaelt die TIR/UIR-
                # Unterbeschriftung: massgeblich dafuer, ob eine Spalte trotz
                # scheinbar anderem Spaltennamen eine Kennzahl (nicht Preis) ist.
                "tir_uir_zeile": 1,
            },
            "Pakete": {
                "whitelist": {
                    "Paket": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen B452"},
                "muster_regeln": [
                    # Hier steht im Original kein Jahr im Label selbst
                    # ("Preis in CHF" wiederholt sich pro Jahresblock).
                    (r"^Preis in CHF$", kategorisieren_chf()),
                    (r"^Uses \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 2,
                # Einige Spalten haben nur in dieser Unterzeile ein Label
                # (TIR/UIR), in der obersten Kopfzeile bleiben sie leer und
                # erben sonst faelschlich das Label der vorherigen Preis-Spalte.
                "tir_uir_zeile": 1,
            },
        },
    },
    "E-Journals Geographie": {
        "muster": "geographie",
        "sheets": {
            # Blatt "Legende E-Journals" ist nur eine Abkuerzungs-Erklaerung
            # (TIR/UIR), keine Daten -> bewusst nicht gelistet, wird nie eingelesen.
            "Geographie 2024-2021": {
                "whitelist": {
                    "Titel":     DURCHLASSEN_TEXT,
                    "ISSN":      DURCHLASSEN_TEXT,
                    "Verlag":    DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen", "Entscheid FR"},
                "muster_regeln": [
                    # Kein "CHF" im Spaltennamen, aber gemaess Datenglossar
                    # bereits als CHF bestaetigt (Kennzahl "Kosten CHF").
                    (r"^Kosten \d{4}$", kategorisieren_chf()),
                    (r"^(CPU|Nutzung) \d{4}$", DURCHLASSEN_NUM),
                ],
                # Zeile 0 (nach Header) ist die TIR/UIR-Unterzeile, keine Datenzeile.
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
            },
        },
    },
    "Nutzungsstatistiken E-Journals BMM": {
        "muster": "e-journals bmm",
        "sheets": {
            # Blatt "Legende E-Journals" ist nur eine Abkuerzungs-Erklaerung,
            # keine Daten -> bewusst nicht gelistet, wird nie eingelesen.
            "Muesmatt 2025-2021": {
                "whitelist": {
                    "Titel":      DURCHLASSEN_TEXT,
                    "ISSN":       DURCHLASSEN_TEXT,
                    "Fachgebiet": DURCHLASSEN_TEXT,
                    "Verlag":     DURCHLASSEN_TEXT,
                    "Lieferant":  DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen", "Entscheid FR"},
                "muster_regeln": [
                    (r"^Kosten \d{4}$", kategorisieren_chf()),
                    (r"^(CPU|Nutzung) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
            },
        },
    },
    "Nutzungsstatistiken E-Journals Bewi": {
        "muster": "e-journals bewi",
        "sheets": {
            # Blatt "Legende E-Journals" ist nur eine Abkuerzungs-Erklaerung,
            # keine Daten -> bewusst nicht gelistet, wird nie eingelesen.
            "Astro-Physik 2021-2024": {
                "whitelist": {
                    "Titel":     DURCHLASSEN_TEXT,
                    "ISSN":      DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                    "Verlag":    DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen", "Entscheid FR"},
                "muster_regeln": [
                    (r"^Kosten \d{4}$", kategorisieren_chf()),
                    (r"^(CPU|Nutzung) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
            },
            "Stat 2021-2024": {
                "whitelist": {
                    "Titel":     DURCHLASSEN_TEXT,
                    "ISSN":      DURCHLASSEN_TEXT,
                    "Verlag":    DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen", "Entscheid FR"},
                "muster_regeln": [
                    (r"^Kosten \d{4}$", kategorisieren_chf()),
                    (r"^(CPU|Nutzung) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
            },
            "Math 2021-2024": {
                "whitelist": {
                    "Titel":     DURCHLASSEN_TEXT,
                    "ISSN":      DURCHLASSEN_TEXT,
                    "Verlag":    DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen", "Entscheid FR"},
                "muster_regeln": [
                    (r"^Kosten \d{4}$", kategorisieren_chf()),
                    (r"^(CPU|Nutzung) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
            },
        },
    },
    "Stats-Lizenzen": {
        "muster": "stats-lizenzen",
        "sheets": {
            sheet: {
                "whitelist": {
                    "package":  DURCHLASSEN_TEXT,
                    "year":     {"typ": "jahr", "aktion": "durchlassen"},
                    "usage":    DURCHLASSEN_NUM,
                    "CPU CHF":  DURCHLASSEN_NUM,
                    "CPU O":    DURCHLASSEN_NUM,
                    "N° of titles": DURCHLASSEN_NUM,
                    "Price CHF excl. MwsT": kategorisieren_chf(),
                    # Fremdwaehrungs-Betrag: Waehrung kommt zeilenweise aus der
                    # per reiche_waehrung_weiter() erzeugten Hilfsspalte.
                    "Price Originalwährung": {
                        "typ": "numerisch", "aktion": "kategorisieren",
                        "waehrungsspalte": "Erkannte Fremdwaehrung",
                    },
                    # Wird durch die Waehrungs-Weiterreichung erzeugt, ist selbst
                    # nicht sensibel (nur ein Waehrungscode) -> einfach durchlassen.
                    "Erkannte Fremdwaehrung": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Comment"},
                "muster_regeln": [],
                # "Summe"/"davon ..."-Zeilen sind Subtotale, keine einzelnen
                # Lizenzen - werden komplett entfernt (nicht kategorisiert).
                "summenzeilen_entfernen": {"spalte": "package", "praefixe": ["Summe", "davon"]},
                # Waehrung steht nur einmal pro Verlags-/Paketgruppe in einer
                # Kopfzeile (Jahr leer); mehrdeutige ("GBP, EUR") oder fehlende
                # Angaben (z.B. "Covidence") werden zu "?" (unbekannt).
                "waehrung_weiterreichen": {
                    "waehrungs_spalte": "Price Originalwährung",
                    "jahr_spalte": "year",
                    "gruppen_spalte": "package",
                    "hilfsspalte": "Erkannte Fremdwaehrung",
                },
            }
            for sheet in ["E-Books MED", "Apps & DB MED", "Fortsetzungen MED"]
        },
    },
    "E-Book-Pakete NAT": {
        "muster": "e-book-pakete nat",
        "sheets": {
            "310204": {
                "whitelist": {
                    "E-Book Pakete":      DURCHLASSEN_TEXT,
                    "Bestellposten Alma": DURCHLASSEN_TEXT,
                    "Erwerbunsgart":      DURCHLASSEN_TEXT,
                    "Etat Code":          DURCHLASSEN_TEXT,
                    "Lieferant":          DURCHLASSEN_TEXT,
                    "Verlag":             DURCHLASSEN_TEXT,
                    "Abo-Nr":             DURCHLASSEN_TEXT,
                    "Fach":               DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    # Waehrung steht direkt im Wert selbst (z.B. "€ 818,67")
                    (r"^Preis FW \d{4}$", kategorisieren_auto()),
                    # Spaltentext leicht uneinheitlich (Leerzeichen, Fussnoten-
                    # Ziffern wie "inkl. Mwst3", "inkl. Mwst24") - daher tolerant.
                    (r"^Preis in CHF\s+inkl\.\s*Mwst\d*$", kategorisieren_chf()),
                ],
                # Am Blattende angehaengter "Gesamtkosten"-Aufschluesselungsblock
                # nach Fachbereich (Sachkonto/Gesamtkosten/Natwiss allg./...) hat
                # keinen Eintrag in der Titel-Spalte - keine echten Zeilen.
                "titel_leer_entfernen": {"spalte": "E-Book Pakete"},
            },
            "310205": {
                "whitelist": {
                    "Ebook Pakete (Forsetzungen)": DURCHLASSEN_TEXT,
                    "Bestellposten Alma":          DURCHLASSEN_TEXT,
                    "Erwerbunsgart":                DURCHLASSEN_TEXT,
                    "Etat Code":                    DURCHLASSEN_TEXT,
                    "Lieferant":                     DURCHLASSEN_TEXT,
                    "Verlag":                        DURCHLASSEN_TEXT,
                    "Abo-Nr":                        DURCHLASSEN_TEXT,
                    "Fach":                          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    (r"^Preis FW \d{4}$", kategorisieren_auto()),
                    (r"^Preis in CHF\s+inkl\.\s*Mwst\d*$", kategorisieren_chf()),
                ],
                # Gleicher angehaengter Gesamtkosten-Block wie in "310204".
                # Achtung: Spaltenname hat im Original ein geschuetztes
                # Leerzeichen (\xa0) am Ende - hier exakt der rohe Name noetig.
                "titel_leer_entfernen": {"spalte": "Ebook Pakete (Forsetzungen)\xa0"},
            },
            "B552": {
                "whitelist": {
                    "Ebook Pakete (Forsetzungen)": DURCHLASSEN_TEXT,
                    "Bestellnummer Alma":          DURCHLASSEN_TEXT,
                    "Erwerbunsgart":                DURCHLASSEN_TEXT,
                    "Etat Code":                     DURCHLASSEN_TEXT,
                    "Lieferant":                     DURCHLASSEN_TEXT,
                    "Verlag":                        DURCHLASSEN_TEXT,
                    "Abo-Nr":                        DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    (r"^Preis FW \d{4}$", kategorisieren_auto()),
                    (r"^Preis in CHF\s+inkl\.\s*Mwst\d*$", kategorisieren_chf()),
                ],
            },
        },
    },
    # ACHTUNG Namenskollision: NICHT zu verwechseln mit "E-Book-Pakete NAT.xlsx"
    # (Muster "e-book-pakete nat", mit Bindestrich). Diese Datei heisst im
    # Original ohne Bindestrich ("EBook-Pakete") und enthaelt trotz gleicher
    # Blattnamen (310204/310205) einen KOMPLETT ANDEREN Inhalt: Nutzungszahlen
    # (CPU/Nutzung/TIR/UTR) statt Abo-Preise (R'datum/Preis FW/inkl. Mwst).
    "Nutzungsstatistiken EBook-Pakete NAT": {
        "muster": "ebook-pakete nat",
        # Blatt "Kontaktdaten" (echte Verlagskontakte, Personendaten) ist
        # bewusst NICHT gelistet -> wird nie eingelesen.
        "sheets": {
            "310204": {
                "whitelist": {
                    "Titel":    DURCHLASSEN_TEXT,
                    "Ref,":     DURCHLASSEN_TEXT,
                    "Verlag":   DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen"},
                "muster_regeln": [
                    (r"^CPU \d{4}$", DURCHLASSEN_NUM),
                    (r"^Nutzung \d{4}$", DURCHLASSEN_NUM),
                    (r"^Preis CHF \d{4}$", kategorisieren_auto()),
                ],
                "datenzeilen_ab": 1,
                # Unterzeile enthaelt TIR/UTR (nicht UIR!) - UTR = Unique Title
                # Request, eine von UIR (Unique Item Request) verschiedene
                # COUNTER-Kennzahl, wird bewusst nicht gleichgesetzt.
                "tir_uir_zeile": 0,
                # Mitten im Blatt eingefuegte Abkuerzungs-Legende ("Legende:",
                # "TIR Total Item Requests", "UTR Unique Title Requests",
                # "Springer Metric: Chapter Request", "CPU Cost per Use") steht
                # direkt in der Titel-Spalte, keine echten Eintraege.
                "summenzeilen_entfernen": {
                    "spalte": "Titel",
                    "praefixe": ["Legende", "TIR ", "UTR ", "CPU ", "Springer Metric"],
                },
            },
            "310205": {
                "whitelist": {
                    "Titel":    DURCHLASSEN_TEXT,
                    "Ref.":     DURCHLASSEN_TEXT,
                    "Verlag":   DURCHLASSEN_TEXT,
                    "Lieferant": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Bemerkungen"},
                "muster_regeln": [
                    (r"^CPU \d{4}$", DURCHLASSEN_NUM),
                    (r"^Nutzung \d{4}$", DURCHLASSEN_NUM),
                    # "Klosten" ist ein Tippfehler im Original (2024-Spalte),
                    # beide Schreibweisen tolerieren.
                    (r"^Kl?osten CHF \d{4}$", kategorisieren_auto()),
                ],
                "datenzeilen_ab": 1,
                "tir_uir_zeile": 0,
                # Gleiche eingefuegte Legende wie in "310204".
                "summenzeilen_entfernen": {
                    "spalte": "Titel",
                    "praefixe": ["Legende", "TIR ", "UTR ", "CPU ", "Springer Metric"],
                },
            },
        },
    },
    "E-Journals B521": {
        "muster": "b521",
        "sheets": {
            "Einzeltitel": {
                # "Journal"/"Journal.1": zwei parallele Titel-Spalten (siehe
                # Hinweis im Chat) - fuer die Bereinigung irrelevant, beide
                # sind unbedenklicher Text.
                "whitelist": {
                    "Journal": DURCHLASSEN_TEXT,
                    "ISSN":    DURCHLASSEN_TEXT,
                },
                # War faelschlich "Notizen B540" (aus der B540-Vorlage kopiert) -
                # die echte Spalte in dieser Datei heisst "Notizen B521".
                "freitext_spalten": {"Notizen B521", "Comments E-Library"},
                "muster_regeln": [
                    (r"^Preis \d{4} in CHF$", kategorisieren_chf()),
                    (r"^(CPU|Uses) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 2,
                "tir_uir_zeile": 1,
            },
            "Pakete": {
                # Kein tr_J1/TIR-UIR-Kopfzeilen-Unterbau hier, Daten beginnen
                # sofort - anders als bei B452/Pakete.
                "whitelist": {
                    "Paket": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen B521"},
                "muster_regeln": [
                    (r"^Preis in CHF$", kategorisieren_chf()),
                    # "CPU" stand urspruenglich ohne Jahr im Header; Jahreszahl
                    # wurde inzwischen ergaenzt (fuer eindeutige Spaltennamen
                    # nach der Bereinigung) - beide Schreibweisen tolerieren.
                    (r"^CPU(\s+\d{4})?$", DURCHLASSEN_NUM),
                    (r"^Grand Total Usage \d{4}$", DURCHLASSEN_NUM),
                ],
            },
        },
    },
    "E-Journals B540 (2014-19)": {
        "muster": "b540-2014",
        "sheets": {
            # Alte, einfache Vorlage: ein Header, keine tr_J1/TIR-UIR-Unterzeilen.
            "VET Zs": {
                "whitelist": {
                    "Einzeltitel Vetsuisse":                 DURCHLASSEN_TEXT,
                    "ISSN":                                  DURCHLASSEN_TEXT,
                    "Erwerbungsart (Print und/oder Online)": DURCHLASSEN_TEXT,
                    "Lieferant":                              DURCHLASSEN_TEXT,
                    "Total uses":                             DURCHLASSEN_NUM,
                    # Verhaeltniszahl (Kosten/Nutzung), kein Rohpreis -> nicht sensibel
                    "Preis pro Click in CHF":                 DURCHLASSEN_NUM,
                },
                "freitext_spalten": {"Kommentar"},
                "muster_regeln": [
                    (r"^Preise \d{4} in CHF inkl\. Mwst$", kategorisieren_chf()),
                    (r"^\d{4} uses$", DURCHLASSEN_NUM),
                ],
            },
        },
    },
    "E-Journals B540 (2024)": {
        "muster": "b540-2024",
        # Identisch zur B521-Vorlage (gleiche Kopfzeilen-Struktur/-Spaltennamen).
        "sheets": {
            "Einzeltitel": {
                "whitelist": {
                    "Journal": DURCHLASSEN_TEXT,
                    "ISSN":    DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen B540", "Comments E-Library"},
                "muster_regeln": [
                    (r"^Preis \d{4} in CHF$", kategorisieren_chf()),
                    (r"^(CPU|Uses) \d{4}$", DURCHLASSEN_NUM),
                ],
                "datenzeilen_ab": 2,
                "tir_uir_zeile": 1,
            },
            "Pakete": {
                "whitelist": {
                    "Paket": DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen B540"},
                "muster_regeln": [
                    (r"^Preis in CHF$", kategorisieren_chf()),
                    # "CPU" stand urspruenglich ohne Jahr im Header; Jahreszahl
                    # wurde inzwischen ergaenzt (fuer eindeutige Spaltennamen
                    # nach der Bereinigung) - beide Schreibweisen tolerieren.
                    (r"^CPU(\s+\d{4})?$", DURCHLASSEN_NUM),
                    (r"^Grand Total Usage \d{4}$", DURCHLASSEN_NUM),
                ],
            },
        },
    },
    "ZS Masterliste BMM GEO": {
        "muster": "bmm geo",
        "sheets": {
            "Muesmatt": {
                "whitelist": {
                    "Einzeltitel Zeitschriften Bibliothek Muesmatt 087-22": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "Fachgebiet":          DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    # Waehrung steht direkt im Wert (CHF/€ vor der Zahl), nicht
                    # an der Anzahl Leerzeichen im Spaltennamen erkennbar.
                    (r"^Preis\s+(?:CHF\s+)?\d{4}$", kategorisieren_auto()),
                ],
                # Am Blattende angehaengte Summen-/Prognose-Zeilen ohne Titel
                # (z.B. "Gesamtausgaben 2026 B554") sind keine echten Eintraege.
                "titel_leer_entfernen": {"spalte": "Einzeltitel Zeitschriften Bibliothek Muesmatt 087-22"},
            },
            "Geographie": {
                "whitelist": {
                    "Einzeltitel Zeitschriften Bibliothek Geographie 087-34": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Einzeltitel Zeitschriften Bibliothek Geographie 087-34"},
            },
            "Abbestellt BMM": {
                "whitelist": {
                    "Einzeltitel Zeitschriften Bibliothek Muesmatt 087-22 \nAbgeschlossene Abos": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "Fachgebiet":          DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}$", kategorisieren_auto()),
                    # Alte, in der Originaldatei mittlerweile korrigierte Spalte
                    # (Name sagte CHF, Inhalt war Fremdwaehrung) - Erkennung ist
                    # ohnehin werteseitig, daher hier zur Sicherheit belassen.
                    (r"^Preis in CHF\s+inkl\.\s*Mwst\d*$", kategorisieren_auto()),
                ],
            },
            "Abbestellt Geo": {
                "whitelist": {
                    "Einzeltitel Zeitschriften Bibliothek Geographie 087-34 Abgeschlossene Abos": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"^R'datum \d{4}$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}$", kategorisieren_auto()),
                ],
            },
        },
    },
    "ZS Masterliste EXWI": {
        "muster": "exwi",
        # Blatt "Uebersicht 2026" ist eine handgemachte Zusammenfassung
        # (Institute nebeneinander, andere Form), keine Einzeldaten -> bewusst
        # nicht gelistet, wird nie eingelesen.
        "sheets": {
            "Astro": {
                "whitelist": {
                    "Zeitschriften Astronomie (087-33/13)": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo Nr.":             DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
                # Am Blattende angehaengte Total-/Prognose-Zeilen ohne Titel
                # (z.B. "Tot. 2026") sind keine echten Eintraege.
                "titel_leer_entfernen": {"spalte": "Zeitschriften Astronomie (087-33/13)"},
            },
            "Phys": {
                "whitelist": {
                    "Zeitschriften Physik 087-33/14": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo Nr.":             DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Zeitschriften Physik 087-33/14"},
            },
            "Math": {
                "whitelist": {
                    "Zeitschriften Mathematik 087-33/11": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten\xa0Alma": DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo Nr":              DURCHLASSEN_TEXT,
                    "Spalte1":             DURCHLASSEN_TEXT,
                    "Preisentwicklung":    DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Zeitschriften Mathematik 087-33/11"},
            },
            "Stat": {
                "whitelist": {
                    "Titel":               DURCHLASSEN_TEXT,
                    "Zeitschriften\xa0Statistik (087-33/12)": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten\xa0Alma": DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo Nr.":             DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen AS"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Zeitschriften\xa0Statistik (087-33/12)"},
            },
            "INF": {
                "whitelist": {
                    "Zeitschriften Informatik 087-32": DURCHLASSEN_TEXT,
                    "ISSN online":         DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten":       DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo Nr.":             DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Zeitschriften Informatik 087-32"},
            },
            "B400": {
                "whitelist": {
                    "Einzeltitel Zeitschriften DeZeMB": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "Fachgebiet":          DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Hrsg.":               DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    # Spaltenname behauptet CHF, das war bei "Abbestellt BMM"
                    # nachweislich schon mal falsch beschriftet -> sicherheitshalber
                    # werteseitige Erkennung statt fixer Waehrung.
                    (r"^Preis in CHF \d{4} inkl\.\s*Mwst\d*$", kategorisieren_auto()),
                    (r"^Preis\s+\d{4}\d*$", kategorisieren_auto()),
                ],
                "titel_leer_entfernen": {"spalte": "Einzeltitel Zeitschriften DeZeMB"},
            },
            "AIP Journals R&P": {
                "whitelist": {
                    "Zeitschriften Physik 087-33/14": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+(?:CHF\s+)?\d{4}\d*$", kategorisieren_auto()),
                ],
            },
            "Abbestellt": {
                "whitelist": {
                    "Mathematik (087.3311)": DURCHLASSEN_TEXT,
                    "ISSN Print":          DURCHLASSEN_TEXT,
                    "ISSN Online":         DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Agenturref":          DURCHLASSEN_TEXT,
                    "PubArt (2019)":       DURCHLASSEN_TEXT,
                    "Abbestellung":        DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [],
            },
        },
    },
    "ZS Online Abos B552": {
        "muster": "online abos b552",
        "sheets": {
            "Pflanzenwissenschaften": {
                "whitelist": {
                    "Einzeltitel Zeitschriften Bibliothek Pflanzenwissenschaften": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Hrsg.":               DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [
                    (r"(?i)^R'datum \d{4}\d*$", DURCHLASSEN_TEXT),
                    (r"^Preis\s+\d{4}\d*$", kategorisieren_auto()),
                ],
            },
            "Alte Liste": {
                "whitelist": {
                    "Alte Bestellungen\xa0(Liste von Jan Dirk)": DURCHLASSEN_TEXT,
                    "ISSN":                DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Erwerbunsgart":       DURCHLASSEN_TEXT,
                    "Bezugsart":           DURCHLASSEN_TEXT,
                    "REF-Nr.":             DURCHLASSEN_TEXT,
                    "Hrsg.":               DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr":              DURCHLASSEN_TEXT,
                    "Herkunft":            DURCHLASSEN_TEXT,
                    "Verwaltung":          DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen"},
                "muster_regeln": [],
                # Personenname im Original-Spaltentitel ("Liste von Jan Dirk")
                # wird in der bereinigten Ausgabe entfernt.
                "spalten_umbenennen": {
                    "Alte Bestellungen\xa0(Liste von Jan Dirk)": "Alte Bestellungen",
                },
            },
        },
    },
    "Übersicht-B452-B521-B540-Lizenzen": {
        # Anders als alle bisherigen Dateien: fast nur Freitext, "Contact
        # details" enthaelt praktisch durchgehend echte Personendaten
        # externer Verlagskontakte (Name/E-Mail/Telefon) -> wird IMMER
        # maskiert (siehe MASKIEREN), nicht wie Notizen nur durchgelassen.
        "muster": "übersicht-b452-b521-b540-lizenzen",
        "sheets": {
            "MED": {
                "whitelist": {
                    "Datenbank Medizin":   DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr.":             DURCHLASSEN_TEXT,
                    "Contact details":     MASKIEREN,
                    "Angebot 2027":        DURCHLASSEN_TEXT,
                    "Rechnung":            DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen", "Deals & Discounts", "Important information"},
                "muster_regeln": [],
                "titel_leer_entfernen": {"spalte": "Datenbank Medizin"},
            },
            "PHM": {
                "whitelist": {
                    "Datenbak Pharmazie":  DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr.":             DURCHLASSEN_TEXT,
                    "Contact details":     MASKIEREN,
                    "Preis 2027":          DURCHLASSEN_TEXT,
                    "Rechnung":            DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen", "Important information"},
                "muster_regeln": [],
                "titel_leer_entfernen": {"spalte": "Datenbak Pharmazie"},
            },
            "B400": {
                "whitelist": {
                    "Datenbak B400":       DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr.":             DURCHLASSEN_TEXT,
                    "Contact details":     MASKIEREN,
                    "Angebot 2027":        DURCHLASSEN_TEXT,
                    "Rechnung":            DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen", "Important information"},
                "muster_regeln": [],
                "titel_leer_entfernen": {"spalte": "Datenbak B400"},
            },
            "VET": {
                "whitelist": {
                    "Datenbank Bibliothek Vetsuisse": DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr.":             DURCHLASSEN_TEXT,
                    "Contact details":     MASKIEREN,
                    "Erneuerung 2027":     DURCHLASSEN_TEXT,
                    "Rechnung":            DURCHLASSEN_TEXT,
                },
                "freitext_spalten": {"Notizen", "Important information"},
                "muster_regeln": [],
                "titel_leer_entfernen": {"spalte": "Datenbank Bibliothek Vetsuisse"},
            },
            "ZMK": {
                "whitelist": {
                    "Datenbank Bibliothek Zahnmedizin": DURCHLASSEN_TEXT,
                    "Erwerbungsart":       DURCHLASSEN_TEXT,
                    "Etat Code":           DURCHLASSEN_TEXT,
                    "Bestellposten Alma":  DURCHLASSEN_TEXT,
                    "Lieferant":           DURCHLASSEN_TEXT,
                    "Verlag":              DURCHLASSEN_TEXT,
                    "Abo-Nr.":             DURCHLASSEN_TEXT,
                    "Contact details":     MASKIEREN,
                },
                "freitext_spalten": {"Notizen", "Deals & Discount", "Important information"},
                "muster_regeln": [],
                "titel_leer_entfernen": {"spalte": "Datenbank Bibliothek Zahnmedizin"},
            },
        },
    },
}


def reiche_waehrung_weiter(df: pd.DataFrame, waehrungs_spalte: str, jahr_spalte: str,
                            gruppen_spalte: str, hilfsspalte: str) -> pd.DataFrame:
    """Fuer Dateien, in denen die Fremdwaehrung nur einmal in einer Gruppen-
    Kopfzeile steht (Jahr leer, Gruppen-Spalte z.B. Verlags-/Paketname gefuellt,
    Waehrungsspalte enthaelt den Waehrungscode statt eines Betrags) und alle
    Datenzeilen darunter (mit echtem Jahr) den Betrag in Fremdwaehrung ohne
    erneute Waehrungsangabe fuehren. Reicht die zuletzt bekannte Waehrung nach
    unten weiter und setzt sie in einer neuen Hilfsspalte ein; mehrdeutige
    ('GBP, EUR') oder fehlende Angaben werden zu '?' (unbekannt), wie
    abgesprochen. Die Kopfzeilen selbst werden in der Waehrungsspalte auf
    leer gesetzt, damit der Waehrungscode dort nicht faelschlich als
    Zahlen-Fehler auffaellt."""
    df = df.copy()
    aktuelle_waehrung = []
    tracker = "?"
    for i in df.index:
        jahr = df.at[i, jahr_spalte]
        gruppe = df.at[i, gruppen_spalte]
        wert = df.at[i, waehrungs_spalte]
        if pd.isna(jahr) and not pd.isna(gruppe):
            # Neue Gruppen-Kopfzeile: Waehrungscode (falls eindeutig) uebernehmen
            if isinstance(wert, str) and re.fullmatch(r"[A-Za-z]{3}", wert.strip()):
                tracker = wert.strip().upper()
            else:
                tracker = "?"
            df.at[i, waehrungs_spalte] = None
        aktuelle_waehrung.append(tracker)
    df[hilfsspalte] = aktuelle_waehrung
    return df


def entferne_summenzeilen(df: pd.DataFrame, spalte: str, praefixe: list,
                           zahlenformate: pd.DataFrame = None):
    """Entfernt Zeilen, deren Wert in 'spalte' mit einem der Praefixe beginnt
    (z.B. 'Summe', 'davon ...'), da das Aggregat-/Subtotal-Zeilen sind, die
    nicht als einzelne Lizenz kategorisiert werden sollen."""
    werte = df[spalte].astype(str).str.strip().str.lower()
    praefixe_klein = tuple(p.lower() for p in praefixe)
    maske = werte.str.startswith(praefixe_klein)
    df = df[~maske].reset_index(drop=True)
    if zahlenformate is not None:
        zahlenformate = zahlenformate[~maske.values].reset_index(drop=True)
    return df, zahlenformate


def entferne_zeilen_ohne_titel(df: pd.DataFrame, spalte: str, zahlenformate: pd.DataFrame = None):
    """Entfernt Zeilen ohne Eintrag in der Titel-Spalte - bei diesen Dateien
    (eine Zeile pro Zeitschrift/Datenbank) ein zuverlaessiges Zeichen fuer
    angehaengte Summen-/Fussnoten-Zeilen am Blattende, keine echten Eintraege."""
    maske = df[spalte].isna()
    df = df[~maske].reset_index(drop=True)
    if zahlenformate is not None:
        zahlenformate = zahlenformate[~maske.values].reset_index(drop=True)
    return df, zahlenformate


def normalisiere_tir_uir_wert(wert: str):
    """Erkennt COUNTER-Unterzeilen-Marker (TIR/UIR/UTR), auch mit Fussnoten-
    Zusatz (z.B. "TIR/Chap Req*" bei Springer-E-Book-Paketen). TIR = Total
    Item Requests, UIR = Unique Item Requests, UTR = Unique TITLE Requests -
    UTR ist gemaess COUNTER eine eigene, von UIR verschiedene Kennzahl und
    wird bewusst NICHT mit UIR gleichgesetzt. Gibt "TIR"/"UIR"/"UTR" zurueck,
    oder None wenn der Wert zu keinem der drei passt."""
    wert = wert.strip().upper()
    if wert.startswith("UTR"):
        return "UTR"
    if wert.startswith("UIR"):
        return "UIR"
    if wert.startswith("TIR"):
        return "TIR"
    return None


def braucht_zahlenformat(sheet_config: dict) -> bool:
    """Prueft, ob eine Konfiguration irgendwo 'kategorisieren_auto' verwendet -
    nur dann lohnt sich der zusaetzliche Aufwand, die Excel-Zahlenformate
    separat per openpyxl einzulesen."""
    regeln = list(sheet_config.get("whitelist", {}).values())
    regeln += [regel for _, regel in sheet_config.get("muster_regeln", [])]
    return any(regel.get("aktion") == "kategorisieren_auto" for regel in regeln)


def waehle_konfiguration(dateiname: str) -> dict:
    name_klein = dateiname.lower()
    # Nicht das erste Muster nehmen, sondern das laengste (spezifischste)
    # Match - sonst wuerde z.B. "nat" in "E-Book-Pakete NAT.xlsx" faelschlich
    # vor "e-book-pakete nat" zuschlagen, nur weil es zuerst im Dict steht.
    treffer = [(label, config) for label, config in CONFIGS.items() if config["muster"] in name_klein]
    if not treffer:
        raise ValueError(
            f"Keine passende Konfiguration fuer '{dateiname}' gefunden. "
            f"Bitte in CONFIGS eintragen oder Dateinamen anpassen."
        )
    label, config = max(treffer, key=lambda lc: len(lc[1]["muster"]))
    print(f"Konfiguration erkannt: {label}")
    return config


def effektive_spaltenlabels(spalten):
    """Baut aus den rohen Pandas-Spaltennamen die eigentlichen Header-Labels:
    - Verschmolzene Header-Zellen (mehrzeilige/mehrspaltige Jahres-Header)
      erscheinen bei Pandas als 'Unnamed: N', werden hier mit dem letzten
      echten Label nach rechts aufgefuellt.
    - Von Pandas automatisch umbenannte Duplikate (z.B. zweimal 'Preis in
      CHF' -> 'Preis in CHF', 'Preis in CHF.1') werden zurueckgestutzt."""
    labels = []
    letztes = None
    for c in spalten:
        c_str = str(c).strip()
        if c_str.startswith("Unnamed"):
            eff = letztes
        else:
            eff = re.sub(r"\.\d+$", "", c_str).strip()
        labels.append(eff)
        letztes = eff
    return labels


def finde_regel(label: str, whitelist: dict, muster_regeln: list):
    if label in whitelist:
        return whitelist[label]
    for muster, regel in muster_regeln:
        if re.match(muster, label):
            return regel
    return None


def bereinige(df: pd.DataFrame, whitelist: dict, freitext_spalten: set, muster_regeln: list = None,
              spalten_umbenennen: dict = None, erzwinge_durchlassen: dict = None,
              zahlenformate: pd.DataFrame = None):
    muster_regeln = muster_regeln or []
    spalten_umbenennen = spalten_umbenennen or {}
    # erzwinge_durchlassen: Spalte (roher Pandas-Name) -> Zielname (TIR/UIR).
    erzwinge_durchlassen = erzwinge_durchlassen or {}
    # astype(object): Spalten koennen gemischt Zahlen/Text (Kategorie-Label,
    # Platzhalter) enthalten, das verhindert Dtype-Warnungen/-Fehler von pandas.
    bereinigt = df.copy().astype(object)
    log = []
    labels = effektive_spaltenlabels(df.columns)

    for j, (spalte, label) in enumerate(zip(df.columns, labels)):
        if label is None:
            continue  # ganz vorne stehende unbenannte Spalte, kein Bezugslabel vorhanden

        if label in freitext_spalten:
            log.append(f"[MANUELL PRUEFEN] Spalte '{spalte}' (Label '{label}') unveraendert "
                        f"durchgelassen, bitte Zeile fuer Zeile pruefen.")
            continue

        if spalte in erzwinge_durchlassen:
            # Anhand einer Unterzeile (z.B. TIR/UIR/UTR) als Kennzahl-Spalte
            # erkannt, nicht als Preis - unabhaengig davon, was der Spaltenname
            # vermuten liesse.
            regel = DURCHLASSEN_NUM
            log.append(f"[UNTERZEILE ERKANNT] Spalte '{spalte}' (Label '{label}') anhand TIR/UIR/UTR-Unterzeile "
                        f"als Kennzahl behandelt, nicht als Preis.")
        else:
            regel = finde_regel(label, whitelist, muster_regeln)
            if regel is None:
                bereinigt[spalte] = PLATZHALTER
                log.append(f"[NEUE SPALTE] '{spalte}' (Label '{label}') passt zu keiner Regel, "
                            f"komplett maskiert. Einmalig pruefen und ggf. Konfiguration ergaenzen.")
                continue

        if regel["aktion"] == "maskieren":
            bereinigt[spalte] = PLATZHALTER
            log.append(f"[BEWUSST MASKIERT] Spalte '{spalte}' (Label '{label}') enthaelt strukturell "
                        f"Personendaten Dritter, wird gemaess Konfiguration immer maskiert.")
            continue

        erwarteter_typ = regel["typ"]
        aktion = regel["aktion"]

        for i, wert in df[spalte].items():
            if pd.isna(wert) or ist_bekannter_leerwert(wert):
                continue
            if not pruefe_typ(wert, erwarteter_typ):
                bereinigt.at[i, spalte] = PLATZHALTER
                log.append(f"[TYP-ABWEICHUNG] Zeile {i + 2}, Spalte '{spalte}' (Label '{label}'): "
                            f"Inhalt '{wert}' passt nicht zu erwartetem Typ '{erwarteter_typ}', maskiert.")
                continue

            if aktion == "kategorisieren":
                waehrung = regel.get("waehrung")
                if waehrung is None:
                    waehrungsspalte = regel.get("waehrungsspalte")
                    waehrung = df.at[i, waehrungsspalte] if waehrungsspalte else "?"
                    if pd.isna(waehrung):
                        waehrung = "?"
                bereinigt.at[i, spalte] = kategorisiere_preis(wert, waehrung)
            elif aktion == "kategorisieren_auto":
                zahlenformat = None
                if zahlenformate is not None and i in zahlenformate.index and j in zahlenformate.columns:
                    zahlenformat = zahlenformate.at[i, j]
                waehrung = bestimme_waehrung(wert, zahlenformat)
                bereinigt.at[i, spalte] = kategorisiere_preis(wert, waehrung)

    # Spalten umbenennen (z.B. Personennamen aus Header entfernen), anhand des
    # effektiven Labels, nicht des rohen (evtl. dedup-suffigierten) Spaltennamens.
    umbenennungen = {}
    for spalte, label in zip(df.columns, labels):
        if label in spalten_umbenennen:
            neuer_name = spalten_umbenennen[label]
            umbenennungen[spalte] = neuer_name
            log.append(f"[SPALTE UMBENANNT] '{spalte}' (Label '{label}') -> '{neuer_name}'.")

    # TIR/UIR-Unterzeile: eigene, spaltengenaue Umbenennung (nicht ueber das
    # geerbte Label, da die beiden betroffenen Spalten dieses oft teilen und
    # sonst nicht unterscheidbar benannt wuerden - z.B. beide "CPU 2025").
    for spalte, neuer_name in erzwinge_durchlassen.items():
        if spalte in umbenennungen:
            continue  # explizit konfigurierte Umbenennung hat Vorrang
        umbenennungen[spalte] = neuer_name
        log.append(f"[SPALTE UMBENANNT] '{spalte}' (anhand TIR/UIR-Unterzeile) -> '{neuer_name}'.")

    if umbenennungen:
        bereinigt = bereinigt.rename(columns=umbenennungen)

    return bereinigt, log


# Waehrungs-Praefixe/-Suffixe, die vor der Zahlen-Erkennung entfernt werden.
# "Fr." ist gleichbedeutend mit CHF (Schweizer Franken), kein Fehler, nur eine
# andere Schreibweise - wird also als gueltige Zahl akzeptiert, nicht maskiert.
WAEHRUNGS_TOKEN = re.compile(r"(?i)\bfr\.?|\b(?:chf|eur|usd|gbp)\b|[€$£]")


def parse_zahl(wert) -> float:
    """Wandelt einen Zellwert in eine Zahl um, auch wenn er als Text mit
    Waehrungsangabe und/oder uneinheitlicher Zahlenschreibweise vorliegt,
    z.B. 'Fr. 82,80', 'CHF  4’331.79', '€ 32.433.61', '- $ 265.73'.
    Wirft ValueError, wenn nach dem Bereinigen keine gueltige Zahl mehr
    uebrig bleibt."""
    text = WAEHRUNGS_TOKEN.sub("", str(wert))
    # Tausendertrennzeichen (Schweizer Apostroph, verschiedene Unicode-Varianten) entfernen
    text = text.replace("’", "").replace("‘", "").replace("'", "")
    text = re.sub(r"\s+", "", text)  # auch Leerzeichen mitten im Wert (z.B. "-  265.73")

    if "," in text and "." in text:
        # Das zuletzt vorkommende Zeichen ist das Dezimaltrennzeichen,
        # das jeweils andere gilt als Tausendertrennzeichen.
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") > 1:
        # Mehrere Kommas ohne Punkt (z.B. "1,071,71", vermutlich ein Tippfehler
        # im Original): alle bis auf das letzte sind Tausendertrennzeichen.
        teile = text.split(",")
        text = "".join(teile[:-1]) + "." + teile[-1]
    elif "," in text:
        text = text.replace(",", ".")
    elif text.count(".") > 1:
        # Mehrere Punkte ohne Komma (z.B. "32.433.61"): alle bis auf den
        # letzten sind Tausendertrennzeichen.
        teile = text.split(".")
        text = "".join(teile[:-1]) + "." + teile[-1]

    return float(text)


def pruefe_typ(wert, erwarteter_typ) -> bool:
    if erwarteter_typ == "numerisch":
        try:
            parse_zahl(wert)
            return True
        except ValueError:
            return False
    if erwarteter_typ == "jahr":
        try:
            return 1900 <= int(float(wert)) <= 2100
        except (ValueError, TypeError):
            return False
    if erwarteter_typ == "text_kategorie":
        return True
    return False


def kategorisiere_preis(wert, waehrung: str = "CHF") -> str:
    zahl = parse_zahl(wert)
    if zahl < 0:
        # Gutschriften/Rabattkorrekturen: eigene Kategorie statt Betrags-Einordnung
        return f"Gutschrift/Korrektur ({waehrung})"
    for obergrenze, label in PREIS_BANDBREITEN:
        if zahl < obergrenze:
            return f"{label} {waehrung}"
    return f"{PREIS_BANDBREITEN[-1][1]} {waehrung}"


def main():
    if len(sys.argv) < 2:
        print("Bitte den Pfad zur Excel-Datei als Argument angeben.")
        print(r'Beispiel: python bereinigen.py "C:\Users\...\Datenbanken_NAT_2026.xlsx"')
        sys.exit(1)

    # .resolve(): macht den Pfad absolut, damit "parent.parent" (Ordner ueber
    # "Stats_original") auch bei relativer Eingabe (z.B. Drag-and-Drop, nur
    # Dateiname ohne Ordnerpfad) korrekt zwei Ebenen hochgeht statt bei "."
    # (dem aktuellen Ordner) stehenzubleiben.
    eingabe_pfad = Path(sys.argv[1]).resolve()
    if not eingabe_pfad.exists():
        print(f"Datei nicht gefunden: {eingabe_pfad}")
        sys.exit(1)

    config = waehle_konfiguration(eingabe_pfad.name)

    alle_blaetter = pd.ExcelFile(eingabe_pfad).sheet_names
    zu_verarbeitende_blaetter = list(config["sheets"].keys())
    nicht_eingelesen = [b for b in alle_blaetter if b not in zu_verarbeitende_blaetter]
    if nicht_eingelesen:
        print(f"Hinweis: Blatt/Blaetter {nicht_eingelesen} werden NICHT eingelesen "
              f"(gemaess Konfiguration nicht Teil der Weitergabe).")

    gesamtes_protokoll = []
    ergebnisse = {}

    for blattname in zu_verarbeitende_blaetter:
        if blattname not in alle_blaetter:
            print(f"Warnung: Blatt '{blattname}' aus der Konfiguration existiert nicht in dieser Datei, wird uebersprungen.")
            continue
        sheet_config = config["sheets"][blattname]
        df = pd.read_excel(eingabe_pfad, sheet_name=blattname)

        zahlenformate = None
        if braucht_zahlenformat(sheet_config):
            zahlenformate = lade_zahlenformate(eingabe_pfad, blattname, len(df))
            # openpyxl liest manchmal mehr (leere, nur formatierte) Zeilen als
            # pandas, das nachtraeglich leere Zeilen am Ende ignoriert -
            # auf die von pandas erkannte Zeilenzahl kuerzen, sonst Versatz.
            zahlenformate = zahlenformate.iloc[:len(df)].reset_index(drop=True)

        summenzeilen_filter = sheet_config.get("summenzeilen_entfernen")
        if summenzeilen_filter:
            df, zahlenformate = entferne_summenzeilen(
                df, summenzeilen_filter["spalte"], summenzeilen_filter["praefixe"], zahlenformate
            )

        titel_filter = sheet_config.get("titel_leer_entfernen")
        if titel_filter:
            df, zahlenformate = entferne_zeilen_ohne_titel(df, titel_filter["spalte"], zahlenformate)

        waehrung_config = sheet_config.get("waehrung_weiterreichen")
        if waehrung_config:
            df = reiche_waehrung_weiter(
                df,
                waehrung_config["waehrungs_spalte"],
                waehrung_config["jahr_spalte"],
                waehrung_config["gruppen_spalte"],
                waehrung_config["hilfsspalte"],
            )

        # Vor dem Abschneiden der Header-Zeilen: TIR/UIR-Unterzeile auswerten,
        # falls konfiguriert, um Spalten zu erkennen, die im obersten Header
        # keinen (oder einen irrefuehrenden, per Vorwaertsauffuellung geerbten)
        # Namen haben. Dict statt Set: haelt zugleich fest, wie die Spalte in
        # der Ausgabe heissen soll (TIR/UIR), da beide Spalten oft denselben
        # geerbten obersten Header-Namen teilen und sonst nicht unterscheidbar
        # benannt wuerden.
        erzwinge_durchlassen = {}
        tir_uir_zeile = sheet_config.get("tir_uir_zeile")
        if tir_uir_zeile is not None:
            unterzeile = df.iloc[tir_uir_zeile]
            # Geerbtes oberstes Label (z.B. "CPU 2025", "Uses 2025") mit TIR/UIR
            # kombinieren -> "CPU 2025 TIR"/"CPU 2025 UIR" usw. Wichtig: pro Jahr
            # gibt es oft ZWEI TIR/UIR-Paare fuer unterschiedliche Kennzahlen
            # (z.B. CPU und Uses) - nur den Jahreszahl zu uebernehmen wuerde
            # diese beiden Kennzahlen unter demselben Namen zusammenwerfen.
            # Nur verwenden, wenn das Label auch wirklich eine Jahreszahl
            # enthaelt (Bestaetigung, dass es sich um ein echtes geerbtes
            # Kennzahl-Label handelt und nicht zufaellig um eine vorherige,
            # thematisch unpassende Text-/Freitextspalte).
            labels_fuer_tir_uir = effektive_spaltenlabels(df.columns)
            for spalte, label in zip(df.columns, labels_fuer_tir_uir):
                marker = normalisiere_tir_uir_wert(str(unterzeile[spalte]))
                if marker is not None:
                    hat_jahr = re.search(r"\d{4}", label or "")
                    neuer_name = f"{label} {marker}" if hat_jahr else marker
                    erzwinge_durchlassen[spalte] = neuer_name

        ab_zeile = sheet_config.get("datenzeilen_ab", 0)
        if ab_zeile:
            df = df.iloc[ab_zeile:].reset_index(drop=True)
            if zahlenformate is not None:
                zahlenformate = zahlenformate.iloc[ab_zeile:].reset_index(drop=True)

        bereinigt, log = bereinige(
            df,
            sheet_config["whitelist"],
            sheet_config["freitext_spalten"],
            sheet_config.get("muster_regeln"),
            sheet_config.get("spalten_umbenennen"),
            erzwinge_durchlassen,
            zahlenformate,
        )
        ergebnisse[blattname] = bereinigt
        gesamtes_protokoll.extend(f"[{blattname}] {eintrag}" for eintrag in log)

    # Bereinigte Dateien landen gesammelt in "Stats_bereinigt" - auf derselben
    # Ebene wie "Stats_original" (also eine Stufe ueber dem Ordner, aus dem die
    # Originaldatei stammt), wird bei Bedarf automatisch angelegt.
    ausgabe_ordner = eingabe_pfad.parent.parent / "Stats_bereinigt"
    ausgabe_ordner.mkdir(exist_ok=True)
    ausgabe_pfad = ausgabe_ordner / (eingabe_pfad.stem + "_bereinigt.xlsx")
    with pd.ExcelWriter(ausgabe_pfad, engine="openpyxl") as writer:
        for blattname, bereinigt in ergebnisse.items():
            tab_name = f"Bereinigt_{blattname}"[:31]  # Excel-Limit: 31 Zeichen pro Blattname
            bereinigt.to_excel(writer, sheet_name=tab_name, index=False)
        pd.DataFrame({"Protokoll": gesamtes_protokoll}).to_excel(writer, sheet_name="Protokoll", index=False)

    print(f"\nFertig. Bereinigte Datei liegt hier:\n{ausgabe_pfad}")
    print(f"({len(gesamtes_protokoll)} Eintraege im Protokoll, siehe Blatt 'Protokoll')")


if __name__ == "__main__":
    try:
        main()
    except ValueError as fehler:
        # Bekannte, verstaendliche Fehler (z.B. keine passende Konfiguration
        # gefunden) ohne Python-Traceback anzeigen - fuer die Nutzung per
        # Drag-and-Drop ohne Terminal-Erfahrung.
        print(f"\nFehler: {fehler}")
        sys.exit(1)
