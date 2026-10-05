#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dashboard.py
============

Interaktiv terminalmeny som orkestrerar hela betalnings- och Visma-flödet.
Verktyget kör de befintliga skripten (via --input/--output) i rätt ordning –
det innehåller ingen egen affärslogik, utan är enbart ett bekvämt gränssnitt.

Konvention för mappar:
    Arbetsmapp (WORK)   = mappen där råfilerna (Bankgiro/Swish) ligger.
    Redigeringsmapp     = WORK\\edit, där rensade/sammanställda filer och CSV hamnar.

Kör:
    python dashboard.py
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VISMA = ROOT / "visma_import"
PY = sys.executable

# --- Skript som menyn anropar -------------------------------------------------
CLEAN_BG = ROOT / "clean_bankgiro_files.py"
SAMMAN_BG = ROOT / "sammanstall_betalningar.py"
CSV_BG = ROOT / "convert_betalningar_to_csv.py"
CLEAN_SW = ROOT / "clean_swish_files.py"
SAMMAN_SW = ROOT / "samla_swish_betalningar.py"
CSV_SW = VISMA.parent / "step_4_skapa_swish_csv.py"
KONV = VISMA / "visma_konvertera_betalningar.py"
REG = VISMA / "visma_register_inbetalningar_fixed.py"
LOGG = VISMA / "visma_sammanstall_loggar.py"


# =============================================================================
# Hjälpfunktioner
# =============================================================================
def run(args: list, stdin_text: str | None = None) -> int:
    """Kör ett underskript och skriver ut resultatet. Returnerar returkoden.

    stdin_text matar in svar automatiskt till skript som frågar interaktivt."""
    cmd = [PY, *[str(a) for a in args]]
    print("\n>>> " + " ".join(f'"{c}"' if " " in c else c for c in cmd))
    print("-" * 70)
    try:
        result = subprocess.run(
            cmd,
            cwd=str(ROOT),
            input=stdin_text,
            text=True,
        )
    except KeyboardInterrupt:
        print("\n(avbrutet – tillbaka till menyn)")
        return 130
    print("-" * 70)
    print(f"[klart, returkod {result.returncode}]")
    return result.returncode


def ask_folder(prompt: str) -> Path | None:
    """Fråga efter en mapp. Tom rad = avbryt."""
    raw = input(prompt).strip().strip('"').strip("'").strip()
    if not raw:
        return None
    path = Path(raw)
    if not path.is_dir():
        print(f"  FEL: mappen finns inte: {path}")
        return None
    return path


def latest_visma_csv(edit_dir: Path) -> Path | None:
    """Senast ändrade *_visma.csv i redigeringsmappen (indata till Visma-reg)."""
    candidates = [
        p for p in edit_dir.glob("*_visma.csv")
        if p.is_file() and not p.name.startswith("~$")
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def pause() -> None:
    input("\nTryck Enter för att återgå till menyn ...")


# =============================================================================
# Menyåtgärder
# =============================================================================
def do_clean_bg(work: Path, edit: Path) -> None:
    run([CLEAN_BG, "--input", work, "--output", edit])


def do_samman_bg(edit: Path) -> None:
    run([SAMMAN_BG, "--input", edit, "--output", edit])


def do_csv_bg(edit: Path) -> None:
    run([CSV_BG, "--input", edit, "--output", edit])


def do_clean_sw(work: Path, edit: Path) -> None:
    run([CLEAN_SW, "--input", work, "--output", edit])


def do_samman_sw(edit: Path) -> None:
    run([SAMMAN_SW, "--input", edit, "--output", edit])


def do_csv_sw(edit: Path) -> None:
    run([CSV_SW, "--input", edit, "--output", edit])


def do_konvertera(edit: Path) -> None:
    """Konvertera alla CSV i redigeringsmappen till Visma-format.

    visma_konvertera_betalningar.py frågar interaktivt efter indata-/utdatamapp;
    här matas redigeringsmappen in automatiskt för både in- och utdata."""
    run([KONV], stdin_text=f"{edit}\n{edit}\n")


def do_register(edit: Path) -> None:
    """Registrera inbetalningar i Visma från senaste *_visma.csv."""
    csv_file = latest_visma_csv(edit)
    if csv_file is None:
        print(f"\nHittade ingen *_visma.csv i {edit}.")
        print("Kör steg 7 (Konvertera till Visma-format) först.")
        return
    print(f"\nIndatafil: {csv_file.name}")
    lage = input("Körläge – (t)orrkörning eller (s)karp? [t]: ").strip().lower()
    args = [REG, csv_file]
    if lage in ("s", "skarp"):
        args.append("--live")
        print("SKARP körning: skriptet frågar om bekräftelse (skriv 'ja').")
    else:
        print("Torrkörning: inget skrivs i Visma.")
    run(args)


def do_sammanstall_loggar(edit: Path) -> None:
    print(f"\nTips: körloggarna (visma_inbetalningar_logg_*.csv) ligger i:\n  {edit}")
    run([LOGG])


# =============================================================================
# Meny
# =============================================================================
def print_menu(work: Path | None) -> None:
    work_txt = str(work) if work else "(ej vald – välj M)"
    print("\n" + "=" * 46)
    print("==== BETALNINGS-DASHBOARD ====")
    print("=" * 46)
    print(f" Arbetsmapp: {work_txt}")
    print()
    print(" BANKGIRO")
    print("  1) Rensa Bankgiro-filer")
    print("  2) Sammanställ Bankgiro")
    print("  3) Skapa CSV (Bankgiro)")
    print(" SWISH")
    print("  4) Rensa Swish-filer")
    print("  5) Sammanställ Swish")
    print("  6) Skapa CSV (Swish)")
    print(" VISMA")
    print("  7) Konvertera till Visma-format")
    print("  8) Registrera i Visma (torr/skarp)")
    print("  9) Sammanställ & stäm av loggar")
    print()
    print("  A) Kör hela Bankgiro-kedjan (1->2->3)")
    print("  B) Kör hela Swish-kedjan (4->5->6)")
    print("  M) Byt arbetsmapp")
    print("  0) Avsluta")


def require_work(work: Path | None) -> bool:
    if work is None:
        print("\nVälj en arbetsmapp först (M).")
        return False
    return True


def main() -> int:
    print("Betalnings-dashboard – orkestrerar hela flödet.")
    work = ask_folder("Ange arbetsmapp (mappen med råfilerna): ")

    while True:
        edit = (work / "edit") if work else None
        print_menu(work)
        val = input(" Välj: ").strip().lower()

        if val == "0":
            print("Avslutar.")
            return 0
        if val == "m":
            new = ask_folder("Ange arbetsmapp (mappen med råfilerna): ")
            if new is not None:
                work = new
            continue

        if val in {"1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b"}:
            if not require_work(work):
                continue
            assert work is not None and edit is not None
            try:
                if val == "1":
                    do_clean_bg(work, edit)
                elif val == "2":
                    do_samman_bg(edit)
                elif val == "3":
                    do_csv_bg(edit)
                elif val == "4":
                    do_clean_sw(work, edit)
                elif val == "5":
                    do_samman_sw(edit)
                elif val == "6":
                    do_csv_sw(edit)
                elif val == "7":
                    do_konvertera(edit)
                elif val == "8":
                    do_register(edit)
                elif val == "9":
                    do_sammanstall_loggar(edit)
                elif val == "a":
                    do_clean_bg(work, edit)
                    do_samman_bg(edit)
                    do_csv_bg(edit)
                elif val == "b":
                    do_clean_sw(work, edit)
                    do_samman_sw(edit)
                    do_csv_sw(edit)
            except KeyboardInterrupt:
                print("\n(avbrutet – tillbaka till menyn)")
            pause()
            continue

        print(f"  Ogiltigt val: {val!r}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nAvbruten av användaren.")
        raise SystemExit(130)
