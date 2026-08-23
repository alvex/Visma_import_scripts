#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
visma_sammanstall_loggar.py
===========================

Sammanstaller Visma-loggfiler (visma_inbetalningar_logg_*_*.csv) till en enda
CSV-fil for en vald manad, ett valt ar och en vald betalningstyp.

Loggfilerna skapas av visma_register_inbetalningar och har foljande kolumner
(faltavgransare ';', kodning utf-8-sig):

    Rad;Betalningsdatum;VismaDatum;Kundnamn;KundID;Fakturanr;Belopp;Status;Felmeddelande;Tidpunkt

Skriptet:
  * fragar efter mapp, manad/ar (t.ex. 2026-06) och betalningstyp (Bank/Skatteverket),
  * letar upp alla loggfiler i mappen,
  * tar med poster som har status OK och betalningsdatum inom vald manad/ar,
  * bevarar ursprungliga kolumner och format,
  * sorterar efter betalningsdatum stigande,
  * markerar dubbletter (samma fakturanummer) i en ny kolumn 'Dubblett',
  * sparar resultatet i samma mapp och visar en sammanstallning.

Krav: Python 3.7+ (endast standardbibliotek, inga externa beroenden).

Kor:
    python visma_sammanstall_loggar.py
"""

from __future__ import annotations

import csv
import glob
import os
import re
import sys
from collections import OrderedDict, defaultdict
from datetime import datetime

# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

# Filnamnsmonster for de loggfiler som ska sammanstallas.
LOG_GLOB = "visma_inbetalningar_logg_*_*"

# Faltavgransare och kodningar (utf-8-sig forst, cp1252 som reserv for aldre export).
CSV_DELIMITER = ";"
READ_ENCODINGS = ("utf-8-sig", "cp1252")
OUTPUT_ENCODING = "utf-8-sig"

# Namnet pa den nya kolumnen for dubblettmarkering.
DUPLICATE_COLUMN = "Dubblett"

# Kolumnalias sa att inlasningen tal sma variationer i rubriker.
BETALNINGSDATUM_ALIASES = ["Betalningsdatum", "Datum", "Bet.dag", "Betalningsdag"]
VISMADATUM_ALIASES = ["VismaDatum", "Visma-datum"]
STATUS_ALIASES = ["Status"]
FAKTURANR_ALIASES = ["Fakturanr", "Fakturanummer", "Fakt.nr", "Faktura"]

# Giltiga betalningstyper (indata -> etikett i filnamnet).
PAYMENT_TYPES = {
    "bank": "bank",
    "skatteverket": "skatteverket",
}

# Svenska manadsnamn (1-12) i gemener for filnamnet.
SWEDISH_MONTHS = {
    1: "januari", 2: "februari", 3: "mars", 4: "april",
    5: "maj", 6: "juni", 7: "juli", 8: "augusti",
    9: "september", 10: "oktober", 11: "november", 12: "december",
}


# ===========================================================================
# Hjalpfunktioner: inmatning
# ===========================================================================
def prompt_folder() -> str:
    """Fragar efter mappen och upprepar tills en giltig mapp anges."""
    while True:
        raw = input("Ange full sokvag till mappen med loggfilerna: ").strip()
        # Ta bort ev. omgivande citattecken (vanligt vid inklistrade Windows-sokvagar).
        raw = raw.strip('"').strip("'")
        if not raw:
            print("  FEL: Ingen sokvag angavs. Forsok igen.\n")
            continue
        folder = os.path.abspath(os.path.expanduser(raw))
        if not os.path.exists(folder):
            print(f"  FEL: Sokvagen finns inte: {folder}\n  Forsok igen.\n")
            continue
        if not os.path.isdir(folder):
            print(f"  FEL: Sokvagen ar inte en mapp: {folder}\n  Forsok igen.\n")
            continue
        return folder


def prompt_month_year() -> tuple[int, int]:
    """Fragar efter manad/ar (t.ex. 2026-06). Upprepar tills korrekt format."""
    while True:
        raw = input("Ange manad och ar (t.ex. 2026-06): ").strip()
        # Tillat 2026-06, 2026/06, 2026 06 och 2026-6.
        m = re.match(r"^\s*(\d{4})\s*[-/ ]\s*(\d{1,2})\s*$", raw)
        if not m:
            print("  FEL: Ogiltigt format. Ange som 2026-06.\n")
            continue
        year = int(m.group(1))
        month = int(m.group(2))
        if not (1 <= month <= 12):
            print("  FEL: Manaden maste vara mellan 01 och 12.\n")
            continue
        return year, month


def prompt_payment_type() -> str:
    """Fragar efter betalningstyp (Bank/Skatteverket). Upprepar tills giltig."""
    while True:
        raw = input("Ange betalningstyp (Bank/Skatteverket): ").strip().lower()
        if raw in PAYMENT_TYPES:
            return PAYMENT_TYPES[raw]
        print("  FEL: Ange antingen 'Bank' eller 'Skatteverket'.\n")


# ===========================================================================
# Hjalpfunktioner: kolumner och datum
# ===========================================================================
def find_column(fieldnames: list[str], aliases: list[str]) -> str | None:
    """
    Hittar det faktiska kolumnnamnet som matchar nagot alias, oberoende av
    stora/sma bokstaver och extra mellanslag. Returnerar None om inget hittas.
    """
    if not fieldnames:
        return None
    normalized = {}
    for fn in fieldnames:
        key = (fn or "").strip().lower()
        # Behall forsta forekomsten om samma rubrik skulle uppträda flera ganger.
        normalized.setdefault(key, fn)
    for alias in aliases:
        actual = normalized.get(alias.strip().lower())
        if actual is not None:
            return actual
    return None


def parse_date(value: str):
    """
    Tolkar ett datum ur loggens text. Returnerar ett datetime.date eller None.

    Hanterar bl.a.:
        2026-04-07
        2026-04-07 00:00:00
        26-04-07            (Visma-format, 2-siffrigt ar)
        2026/04/07
        07/04/2026          (svensk dag/manad/ar)
        20260407
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    # Ta bort ev. klockslag ("2026-04-07 00:00:00" -> "2026-04-07").
    text = text.split(" ")[0].split("T")[0]

    formats = [
        "%Y-%m-%d",   # 2026-04-07
        "%y-%m-%d",   # 26-04-07 (Visma)
        "%Y/%m/%d",   # 2026/04/07
        "%d/%m/%Y",   # 07/04/2026
        "%d/%m/%y",   # 07/04/26
        "%d-%m-%Y",   # 07-04-2026
        "%d.%m.%Y",   # 07.04.2026
        "%Y%m%d",     # 20260407
    ]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def status_is_ok(value: str) -> bool:
    """True om statustexten ar 'OK' oberoende av gemener/versaler och mellanslag."""
    return (value or "").strip().upper() == "OK"


# ===========================================================================
# Inlasning av en loggfil
# ===========================================================================
def read_log_file(path: str):
    """
    Laser en loggfil och returnerar (fieldnames, rows).
    Provar flera kodningar. Kastar undantag vid ogiltig/olasbar fil sa att
    anroparen kan hoppa over filen med en varning.
    """
    last_error = None
    for enc in READ_ENCODINGS:
        try:
            with open(path, "r", newline="", encoding=enc) as fh:
                reader = csv.DictReader(fh, delimiter=CSV_DELIMITER)
                fieldnames = reader.fieldnames
                rows = list(reader)
            return fieldnames, rows
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    # Om ingen kodning fungerade, las binart och avkoda med ersattning som sista utvag.
    with open(path, "r", newline="", encoding="cp1252", errors="replace") as fh:
        reader = csv.DictReader(fh, delimiter=CSV_DELIMITER)
        fieldnames = reader.fieldnames
        rows = list(reader)
    return fieldnames, rows


# ===========================================================================
# Huvudlogik
# ===========================================================================
def collect_rows(folder: str, year: int, month: int, output_name: str):
    """
    Gar igenom alla loggfiler i mappen och samlar poster som ar OK och inom
    vald manad/ar. Returnerar en resultatstruktur med statistik och varningar.
    """
    pattern = os.path.join(folder, LOG_GLOB)
    matches = sorted(glob.glob(pattern))
    # Uteslut ev. tidigare skapad utdatafil om den skulle rakas matcha monstret.
    matches = [p for p in matches if os.path.basename(p) != output_name]

    stats = {
        "found_files": len(matches),
        "processed_files": 0,
        "rows_read": 0,
        "rows_included": 0,
        "rows_skipped": 0,
    }
    warnings: list[str] = []
    included: list[dict] = []          # varje post + intern sorteringsnyckel
    column_order: "OrderedDict[str, None]" = OrderedDict()

    for path in matches:
        name = os.path.basename(path)

        # Tom fil?
        try:
            if os.path.getsize(path) == 0:
                warnings.append(f"Tom fil hoppades over: {name}")
                continue
        except OSError as exc:
            warnings.append(f"Kunde inte lasa storlek pa {name}: {exc}")
            continue

        # Las filen.
        try:
            fieldnames, rows = read_log_file(path)
        except Exception as exc:  # robust: en trasig fil far inte stoppa korningen
            warnings.append(f"Kunde inte lasa {name}: {exc}")
            continue

        if not fieldnames:
            warnings.append(f"Filen saknar rubrikrad och hoppades over: {name}")
            continue

        # Hitta obligatoriska kolumner i just denna fil.
        col_datum = find_column(fieldnames, BETALNINGSDATUM_ALIASES)
        col_visma = find_column(fieldnames, VISMADATUM_ALIASES)
        col_status = find_column(fieldnames, STATUS_ALIASES)
        col_faktnr = find_column(fieldnames, FAKTURANR_ALIASES)

        missing = []
        if col_datum is None and col_visma is None:
            missing.append("Betalningsdatum")
        if col_status is None:
            missing.append("Status")
        if col_faktnr is None:
            missing.append("Fakturanr")
        if missing:
            warnings.append(
                f"Filen saknar kolumn(er) {', '.join(missing)} och hoppades over: {name} "
                f"(hittade: {list(fieldnames)})"
            )
            # Filen kunde lasas men kan inte bearbetas -> rakna raderna som overhoppade.
            stats["rows_read"] += len(rows)
            stats["rows_skipped"] += len(rows)
            continue

        stats["processed_files"] += 1

        # Registrera kolumnordning (union, forsta forekomst avgor ordningen).
        for fn in fieldnames:
            column_order.setdefault(fn, None)

        for row in rows:
            stats["rows_read"] += 1

            # Endast status OK.
            if not status_is_ok(row.get(col_status, "")):
                stats["rows_skipped"] += 1
                continue

            # Datum: primart Betalningsdatum, annars VismaDatum.
            raw_date = row.get(col_datum, "") if col_datum else ""
            date_obj = parse_date(raw_date)
            if date_obj is None and col_visma:
                date_obj = parse_date(row.get(col_visma, ""))

            if date_obj is None:
                stats["rows_skipped"] += 1
                warnings.append(
                    f"Ogiltigt/saknat datum hoppades over i {name}: "
                    f"'{raw_date}' (fakturanr {row.get(col_faktnr, '')})"
                )
                continue

            # Ratt manad och ar?
            if date_obj.year != year or date_obj.month != month:
                stats["rows_skipped"] += 1
                continue

            included.append({
                "row": row,
                "sort_key": date_obj,
                "source": name,
                "faktnr_col": col_faktnr,
            })
            stats["rows_included"] += 1

    return {
        "stats": stats,
        "warnings": warnings,
        "included": included,
        "column_order": list(column_order.keys()),
    }


def mark_duplicates(included: list[dict]) -> tuple[list[str], int]:
    """
    Markerar dubbletter baserat pa fakturanummer. Satter rad['row'][DUPLICATE_COLUMN]
    till 'JA' eller 'NEJ'. Behaller alla forekomster.

    Returnerar (varningsrader, antal_dubblettposter).
    """
    # Gruppera radindex per fakturanummer.
    by_faktnr: "defaultdict[str, list[int]]" = defaultdict(list)
    for idx, item in enumerate(included):
        row = item["row"]
        faktnr = (row.get(item["faktnr_col"], "") or "").strip()
        by_faktnr[faktnr].append(idx)

    duplicate_rows = 0
    warnings: list[str] = []

    for faktnr, idxs in by_faktnr.items():
        # Tomt fakturanummer betraktas inte som dubblett av varandra.
        is_dup = len(idxs) > 1 and faktnr != ""
        for idx in idxs:
            included[idx]["row"][DUPLICATE_COLUMN] = "JA" if is_dup else "NEJ"
        if is_dup:
            duplicate_rows += len(idxs)
            sources = sorted({included[i]["source"] for i in idxs})
            warnings.append(
                f"Dubblett: fakturanr {faktnr} forekommer {len(idxs)} ganger "
                f"(kallfiler: {', '.join(sources)})"
            )

    return warnings, duplicate_rows


def write_output(folder: str, output_name: str, column_order: list[str],
                 included: list[dict]) -> str:
    """
    Skriver den sammanstallda filen. Bevarar ursprungliga kolumner/format och
    lagger till kolumnen 'Dubblett' sist. Returnerar full sokvag till filen.
    """
    out_path = os.path.join(folder, output_name)

    # Slutlig kolumnordning: ursprungliga kolumner + Dubblett sist.
    out_columns = list(column_order)
    if DUPLICATE_COLUMN not in out_columns:
        out_columns.append(DUPLICATE_COLUMN)

    with open(out_path, "w", newline="", encoding=OUTPUT_ENCODING) as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=out_columns,
            delimiter=CSV_DELIMITER,
            restval="",              # kolumner som saknas i en rad blir tomma
            extrasaction="ignore",   # okanda extra falt ignoreras tyst
        )
        writer.writeheader()
        for item in included:
            writer.writerow(item["row"])

    return out_path


# ===========================================================================
# Main
# ===========================================================================
def main() -> int:
    print("=" * 60)
    print(" Sammanstallning av Visma-loggfiler")
    print("=" * 60)

    folder = prompt_folder()
    year, month = prompt_month_year()
    payment_type = prompt_payment_type()

    month_name = SWEDISH_MONTHS[month]
    output_name = f"visma_inbetalningar_{payment_type}_{month_name}_{year}_logg.csv"

    print(
        f"\nMapp:          {folder}\n"
        f"Period:        {month_name} {year} ({year}-{month:02d})\n"
        f"Betalningstyp: {payment_type}\n"
        f"Utdatafil:     {output_name}\n"
    )

    result = collect_rows(folder, year, month, output_name)
    stats = result["stats"]
    warnings = list(result["warnings"])
    included = result["included"]

    if stats["found_files"] == 0:
        print(
            f"Inga loggfiler som matchar '{LOG_GLOB}' hittades i mappen.\n"
            "Ingen fil skapades."
        )
        return 1

    # Sortera efter betalningsdatum stigande (stabil sortering bevarar filordning).
    included.sort(key=lambda item: item["sort_key"])

    # Dubblettmarkering.
    dup_warnings, duplicate_rows = mark_duplicates(included)
    warnings.extend(dup_warnings)

    # Skriv utdata (aven om 0 poster -> ger en fil med enbart rubrikrad).
    out_path = write_output(folder, output_name, result["column_order"], included)

    # --- Varningar ---
    if warnings:
        print("VARNINGAR:")
        for w in warnings:
            print(f"  - {w}")
        print()

    # --- Sammanstallning ---
    print("=" * 60)
    print(" SAMMANSTALLNING")
    print("=" * 60)
    print(f"  Hittade loggfiler:          {stats['found_files']}")
    print(f"  Behandlade filer:           {stats['processed_files']}")
    print(f"  Lasta poster:               {stats['rows_read']}")
    print(f"  Inkluderade (status OK):    {stats['rows_included']}")
    print(f"  Overhoppade poster:         {stats['rows_skipped']}")
    print(f"  Dubblettposter (markerade): {duplicate_rows}")
    print(f"  Skapad fil:                 {out_path}")
    print("=" * 60)

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nAvbruten av anvandaren.")
        sys.exit(130)
