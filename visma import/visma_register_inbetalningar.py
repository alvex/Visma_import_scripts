#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
visma_register_inbetalningar.py
================================

Halvautomatisk registrering av kundinbetalningar i Visma Compact 6
(Program -> Kundreskontra -> Inbetalningar) baserat pa en Excel-fil.

VIKTIGT (ekonomisk automation):
  - Scriptet oppnar INTE meny/program sjalv. Anvandaren forbereder Visma manuellt.
  - Scriptet klickar ALDRIG slutlig bokforing / Bankgiro automatiskt i denna version.
  - Standardlage ar sakert: DRY_RUN = True och CONFIRM_EACH_ROW = True.
  - Testa alltid pa en backup forst.

Se instruktioner langst ner i filen (modul-docstring INSTALL/USAGE) samt README-utskrift
via kommandoraden:  python visma_register_inbetalningar.py --help

Beroenden:
    pip install pandas openpyxl pyautogui pywinauto pyperclip
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Beroenden importeras "mjukt" sa att t.ex. Excel-lasning fungerar aven om
# GUI-biblioteken saknas (praktiskt vid utveckling pa annan dator).
# ---------------------------------------------------------------------------
try:
    import pandas as pd
except ImportError:  # pragma: no cover
    print("FEL: 'pandas' saknas. Kor: pip install pandas openpyxl")
    sys.exit(1)

try:
    import pyperclip  # saker textinmatning via urklipp
except ImportError:  # pragma: no cover
    pyperclip = None

try:
    import pyautogui  # tangentbord/mus
    # Sakerhet: flytta musen till hornet (0,0) for att avbryta (failsafe).
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.05
except ImportError:  # pragma: no cover
    pyautogui = None

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
except ImportError:  # pragma: no cover
    Application = None
    Desktop = None
    ElementNotFoundError = Exception


# ===========================================================================
# 1. KONFIGURATION  (justera har)
# ===========================================================================

# --- Fonster / konto ---
VISMA_WINDOW_TITLE_CONTAINS = "Visma Compact 6"  # del av Visma-fonstrets titel
INBETALNING_WINDOW_TITLE_CONTAINS = "Inbetalning"  # del av inbetalningsfonstrets titel
DIALOG_TITLE_CONTAINS = "Ratt belopp"              # dialogen "Ratt belopp?"
BANK_ACCOUNT = "1930"                              # forvantat konto (kontrolleras bara visuellt)

# --- Sakerhetslagen ---
DRY_RUN = True             # True = simulera, klicka aldrig OK pa riktigt
CONFIRM_EACH_ROW = True    # True = pausa och fraga fore varje OK
ALLOW_DUPLICATES = False   # False = stoppa om samma fakturanr forekommer flera ganger
AUTO_FINALIZE_PAYMENT_BATCH = False  # MASTE vara False. Klicka aldrig Bankgiro/Bokfor auto.

# --- Belopp ---
AMOUNT_TOLERANCE = Decimal("1.00")  # tillaten differens Excel vs Visma

# --- Timing (sekunder) ---
WAIT_AFTER_ENTER = 1.0     # vantetid efter att fakturanr skrivits + Enter
WAIT_AFTER_OK = 0.8        # vantetid efter OK
WAIT_SHORT = 0.25          # kort paus mellan tangenttryck/falt
DIALOG_TIMEOUT = 5.0       # hur lange vi vantar pa "Ratt belopp?"-dialogen

# --- Test ---
MAX_ROWS: Optional[int] = 3   # kor bara N rader vid test. Satt None for alla rader.

# --- Fakturanummer-extraktion ---
INVOICE_MIN_DIGITS = 4
INVOICE_MAX_DIGITS = 8

# --- Loggning ---
LOG_PREFIX = "visma_inbetalningar_logg_"

# Forvantade Excel-kolumner (exakt dessa rubriker kravs)
REQUIRED_COLUMNS = ["Datum", "Avsandare", "Betalningsreferens", "Belopp"]
# Not: "Avsandare" stavas har utan svenska tecken for att undvika encoding-strul.
# Om din Excel har "Avsandare" med a-ring, se COLUMN_ALIASES nedan.
COLUMN_ALIASES = {
    "Avsandare": ["Avsandare", "Avsändare", "Avsandare "],
    "Datum": ["Datum"],
    "Betalningsreferens": ["Betalningsreferens", "Referens", "Bet.ref", "Betalningsref"],
    "Belopp": ["Belopp"],
}


# ===========================================================================
# Statuskonstanter for loggen
# ===========================================================================
class Status:
    OK = "OK"
    HOPPAD = "HOPPAD"
    FEL = "FEL"
    STOPPAD = "STOPPAD"
    DRY_RUN = "DRY_RUN"


# ===========================================================================
# Loggpost / logg
# ===========================================================================
@dataclass
class LogRow:
    rad: int
    datum: str = ""
    visma_datum: str = ""
    avsandare: str = ""
    betalningsreferens: str = ""
    fakturanr: str = ""
    belopp: str = ""
    status: str = ""
    felmeddelande: str = ""
    tidpunkt: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def as_dict(self) -> dict[str, Any]:
        return {
            "Rad": self.rad,
            "Datum": self.datum,
            "VismaDatum": self.visma_datum,
            "Avsandare": self.avsandare,
            "Betalningsreferens": self.betalningsreferens,
            "Fakturanr": self.fakturanr,
            "Belopp": self.belopp,
            "Status": self.status,
            "Felmeddelande": self.felmeddelande,
            "Tidpunkt": self.tidpunkt,
        }


class Logger:
    """Samlar loggrader och skriver till Excel bredvid indatafilen."""

    def __init__(self, out_dir: Path):
        self.rows: list[LogRow] = []
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = out_dir / f"{LOG_PREFIX}{stamp}.xlsx"

    def add(self, row: LogRow) -> None:
        self.rows.append(row)
        # Snabb konsol-spegling
        msg = f"[Rad {row.rad}] {row.status} | faktura={row.fakturanr} belopp={row.belopp}"
        if row.felmeddelande:
            msg += f" | {row.felmeddelande}"
        print("   " + msg)

    def save(self) -> None:
        if not self.rows:
            return
        df = pd.DataFrame([r.as_dict() for r in self.rows])
        try:
            df.to_excel(self.path, index=False)
            print(f"\nLogg sparad: {self.path}")
        except Exception as exc:  # pragma: no cover
            print(f"VARNING: Kunde inte spara logg-Excel ({exc}).")
            # Fallback till CSV sa att inget forloras.
            csv_path = self.path.with_suffix(".csv")
            df.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")
            print(f"Logg sparad som CSV istallet: {csv_path}")


# ===========================================================================
# 4. EXTRAHERA FAKTURANUMMER
# ===========================================================================
def extract_invoice_number(reference: Any) -> str:
    """
    Plockar ut fakturanummer (4-8 siffror) ur en betalningsreferens.

    Exempel:
        "49577"           -> "49577"
        "Faktura 49577"   -> "49577"
        "OCR 49577"       -> "49577"
        "Bet ref 49577"   -> "49577"
        "Faktura 49571"   -> "49571"

    Regel:
        - Leta efter siffersekvenser pa 4-8 siffror.
        - Om flera hittas: valj det mest sannolika fakturanumret
          (foredra det som foljer pa ordet "faktura", annars langsta/sista).
        - Hittas inget -> returnera "" (raden hanteras som fel av anroparen).
    """
    if reference is None:
        return ""
    text = str(reference).strip()
    if not text:
        return ""

    lower = text.lower()

    # 1) Foredra nummer som star direkt efter "faktura"/"fakt"/"ocr"/"ref".
    for keyword in ("faktura", "fakt.nr", "fakt", "ocr", "ref"):
        m = re.search(rf"{keyword}[^\d]*(\d{{{INVOICE_MIN_DIGITS},{INVOICE_MAX_DIGITS}}})", lower)
        if m:
            return m.group(1)

    # 2) Annars: hitta alla kandidater med ratt langd.
    candidates = re.findall(rf"\d{{{INVOICE_MIN_DIGITS},{INVOICE_MAX_DIGITS}}}", text)
    if not candidates:
        # Sista utvag: langre sifferkorm (t.ex. OCR) -> ta forsta 4-8 siffrorna? Hoppa hellre.
        return ""

    if len(candidates) == 1:
        return candidates[0]

    # Flera kandidater: valj det langsta; vid lika langd, det sista
    # (fakturanumret star ofta sist i referenstexten).
    best = candidates[0]
    for c in candidates:
        if len(c) > len(best) or (len(c) == len(best)):
            best = c
    return best


# ===========================================================================
# 5. DATUMFORMAT
# ===========================================================================
def format_visma_date(value: Any) -> str:
    """
    Konverterar ett datum till Visma-format YY-MM-DD.

    Exempel:
        2026-04-07               -> "26-04-07"
        datetime(2026,4,7)       -> "26-04-07"
        "2026-04-07"             -> "26-04-07"
        "7/4 2026"               -> "26-04-07"  (svensk dag/manad)
        "" / None / NaT          -> ""  (anroparen hanterar som fel)

    Kastar inte exception pa tomma varden; returnerar tom strang.
    """
    if value is None:
        return ""

    # pandas NaT / NaN
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    # Redan datetime/date
    if isinstance(value, (datetime, date)):
        d = value
        return f"{d.year % 100:02d}-{d.month:02d}-{d.day:02d}"

    text = str(value).strip()
    if not text:
        return ""

    # Om pandas redan gav t.ex. "2026-04-07 00:00:00"
    text = text.split(" ")[0]

    # Prova en rad kanda format.
    fmts = [
        "%Y-%m-%d",   # 2026-04-07
        "%y-%m-%d",   # 26-04-07
        "%Y/%m/%d",
        "%d/%m/%Y",   # svensk dag/manad/ar
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%Y%m%d",
    ]
    for fmt in fmts:
        try:
            d = datetime.strptime(text, fmt)
            return f"{d.year % 100:02d}-{d.month:02d}-{d.day:02d}"
        except ValueError:
            continue

    # Sista utvag: lat pandas gissa.
    try:
        d = pd.to_datetime(text, dayfirst=True, errors="raise")
        return f"{d.year % 100:02d}-{d.month:02d}-{d.day:02d}"
    except Exception:
        return ""


# ===========================================================================
# 6. BELOPPSFORMAT
# ===========================================================================
def parse_amount(value: Any) -> Optional[Decimal]:
    """
    Normaliserar ett belopp till Decimal for jamforelse.

    Hanterar:
        2017        -> Decimal("2017")
        2017.0      -> Decimal("2017.0")
        "2017,00"   -> Decimal("2017.00")
        "2 017,00"  -> Decimal("2017.00")   (mellanslag/tusenavskiljare)
        "1 396,50"  -> Decimal("1396.50")
        "" / None   -> None
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None

    text = str(value).strip()
    if not text:
        return None

    # Ta bort valuta och mellanslag (aven hardmellanslag \xa0).
    text = text.replace("kr", "").replace("SEK", "")
    text = text.replace("\xa0", "").replace(" ", "")

    # Svensk decimal: komma -> punkt. Punkt anvands ibland som tusenavskiljare.
    if "," in text:
        # Anta komma = decimaltecken; ta bort punkter (tusenavskiljare).
        text = text.replace(".", "").replace(",", ".")
    # annars: punkt tolkas som decimaltecken (t.ex. 2017.0)

    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_visma_amount(value: Any) -> str:
    """
    Formaterar belopp for inmatning i Visma med svensk decimal (komma, 2 dec).

    Exempel:
        2017        -> "2017,00"
        2017.0      -> "2017,00"
        "2 017,00"  -> "2017,00"
        1396.5      -> "1396,50"

    Not: Vissa Visma-falt vill ha tusenavskiljare (2 017,00). Byt UTAN_TUSEN
    till False nedan om ditt Visma kraver mellanslag som tusenavskiljare.
    """
    UTAN_TUSEN = True  # <-- justera vid behov (True = "2017,00", False = "2 017,00")

    dec = parse_amount(value)
    if dec is None:
        return ""
    # Kvantisera till 2 decimaler.
    dec = dec.quantize(Decimal("0.01"))

    if UTAN_TUSEN:
        s = f"{dec:.2f}"          # "2017.00"
        return s.replace(".", ",")  # "2017,00"
    else:
        # Med mellanslag som tusenavskiljare: "2 017,00"
        int_part, dec_part = f"{dec:.2f}".split(".")
        neg = int_part.startswith("-")
        int_part = int_part.lstrip("-")
        grouped = ""
        while len(int_part) > 3:
            grouped = " " + int_part[-3:] + grouped
            int_part = int_part[:-3]
        grouped = int_part + grouped
        if neg:
            grouped = "-" + grouped
        return f"{grouped},{dec_part}"


# ===========================================================================
# 14. GUI-AUTOMATION (hjalpfunktioner)
# ===========================================================================
class VismaGui:
    """
    Kapslar in fonster-/tangent-/mus-hantering mot Visma Compact 6.

    OBS: Falt-navigering (t.ex. till "Bet.dag" och "Fakt.nr") ar det
    stalle som oftast maste justeras for just din Visma-installation.
    Se metoderna focus_fakturanr_field() och write_payment_row().
    """

    def __init__(self):
        self.app = None
        self.main_window = None

    # --- Fonsterhantering ---------------------------------------------------
    def connect(self) -> None:
        """Anslut till redan oppet Visma-fonster. Stoppar om det inte hittas."""
        if Application is None:
            print("VARNING: pywinauto saknas. Kor i begransat lage (endast pyautogui).")
            return
        try:
            self.app = Application(backend="uia").connect(
                title_re=f".*{re.escape(VISMA_WINDOW_TITLE_CONTAINS)}.*",
                timeout=5,
            )
            self.main_window = self.app.top_window()
            print(f"Ansl0t till Visma-fonster: '{self.main_window.window_text()}'")
        except Exception as exc:
            raise RuntimeError(
                f"Kunde inte hitta Visma-fonster som innehaller "
                f"'{VISMA_WINDOW_TITLE_CONTAINS}'. Ar programmet oppet? ({exc})"
            )

    def activate_visma_window(self) -> None:
        """Ta fram och fokusera Visma-fonstret."""
        if self.main_window is not None:
            try:
                self.main_window.set_focus()
                time.sleep(WAIT_SHORT)
                return
            except Exception:
                pass
        # Fallback: pyautogui kan inte byta fonster; be anvandaren klicka.
        print("VARNING: Kunde inte automatiskt fokusera Visma. Klicka i fonstret sjalv.")

    # --- Tangent/urklipp ----------------------------------------------------
    @staticmethod
    def paste_text(text: str) -> None:
        """Skriver text via urklipp (snabbt och sakert mot special-tecken)."""
        if pyperclip is not None and pyautogui is not None:
            pyperclip.copy(text)
            time.sleep(WAIT_SHORT)
            pyautogui.hotkey("ctrl", "v")
        elif pyautogui is not None:
            # Fallback: langsam tangentinmatning.
            pyautogui.typewrite(text, interval=0.02)
        else:
            print(f"[SIMULERAT] Skriv text: {text!r}")
        time.sleep(WAIT_SHORT)

    @staticmethod
    def press_tab(times: int = 1) -> None:
        for _ in range(times):
            if pyautogui is not None:
                pyautogui.press("tab")
            else:
                print("[SIMULERAT] TAB")
            time.sleep(WAIT_SHORT)

    @staticmethod
    def press_enter() -> None:
        if pyautogui is not None:
            pyautogui.press("enter")
        else:
            print("[SIMULERAT] ENTER")

    @staticmethod
    def press_esc() -> None:
        if pyautogui is not None:
            pyautogui.press("esc")
        else:
            print("[SIMULERAT] ESC")

    @staticmethod
    def select_all_and_clear() -> None:
        """Markera allt i aktivt falt och rensa (Ctrl+A + Delete)."""
        if pyautogui is not None:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(WAIT_SHORT)
            pyautogui.press("delete")
            time.sleep(WAIT_SHORT)

    # --- Dialoghantering ----------------------------------------------------
    def wait_for_dialog(self, title_contains: str, timeout: float = DIALOG_TIMEOUT):
        """
        Vantar pa en dialog vars titel innehaller title_contains.
        Returnerar dialog-wrapper (pywinauto) eller None om den inte dyker upp.
        """
        if Desktop is None:
            # Utan pywinauto kan vi inte inspektera dialogen; returnera None
            # sa att anroparen faller tillbaka pa tangentbekraftelse.
            time.sleep(min(timeout, WAIT_AFTER_ENTER))
            return None

        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                dlg = Desktop(backend="uia").window(
                    title_re=f".*{re.escape(title_contains)}.*"
                )
                if dlg.exists(timeout=0.3):
                    return dlg
            except (ElementNotFoundError, Exception):
                pass
            time.sleep(0.2)
        return None

    @staticmethod
    def dialog_text(dlg) -> str:
        """Samlar all synlig text i dialogen (for verifiering av fakturanr/belopp)."""
        if dlg is None:
            return ""
        try:
            parts = []
            for child in dlg.descendants():
                try:
                    t = child.window_text()
                    if t:
                        parts.append(t)
                except Exception:
                    continue
            return " | ".join(parts)
        except Exception:
            return ""

    def click_ok_on_dialog(self, dlg=None) -> bool:
        """
        Klickar OK. Med pywinauto-dialog anvands OK-knappen; annars Enter.
        Returnerar True om ett OK-klick utfordes (eller simulerades).
        """
        if dlg is not None:
            for name in ("OK", "&OK", "Ok"):
                try:
                    btn = dlg.child_window(title=name, control_type="Button")
                    if btn.exists(timeout=0.5):
                        btn.click_input()
                        return True
                except Exception:
                    continue
        # Fallback: Enter bekraftar oftast OK.
        self.press_enter()
        return True

    # --- Falt-navigering (JUSTERA HAR VID BEHOV) ----------------------------
    def focus_fakturanr_field(self) -> None:
        """
        Placerar fokus i faltet 'Fakt.nr.'.

        I forsta versionen litar vi pa att ANVANDAREN redan klickat i faltet
        (enligt startinstruktionerna). Om du vill automatisera detta:
          - antingen leta upp faltet via pywinauto child_window(...)
          - eller anvand en fast muskoordinat (pyautogui.click(x, y))
        Lamnas medvetet tomt/passivt for att undvika felklick.
        """
        # Exempel (avkommentera + justera koordinat om du vill):
        # if pyautogui is not None:
        #     pyautogui.click(FAKTNR_X, FAKTNR_Y)
        #     time.sleep(WAIT_SHORT)
        pass


# ===========================================================================
# Excel-lasning + kolumnvalidering
# ===========================================================================
def resolve_columns(df: "pd.DataFrame") -> dict[str, str]:
    """
    Matchar de faktiska kolumnnamnen mot REQUIRED_COLUMNS via COLUMN_ALIASES.
    Returnerar mappning {logiskt_namn: faktiskt_kolumnnamn}.
    Stoppar med tydligt fel om nagon obligatorisk kolumn saknas.
    """
    actual = {str(c).strip(): c for c in df.columns}
    mapping: dict[str, str] = {}
    missing: list[str] = []

    for logical in REQUIRED_COLUMNS:
        aliases = COLUMN_ALIASES.get(logical, [logical])
        found = None
        for alias in aliases:
            for actual_name, orig in actual.items():
                if actual_name.lower() == alias.lower():
                    found = orig
                    break
            if found:
                break
        if found is None:
            missing.append(logical)
        else:
            mapping[logical] = found

    if missing:
        raise ValueError(
            "Excel-filen saknar obligatoriska kolumner: "
            + ", ".join(missing)
            + f"\nHittade kolumner: {list(df.columns)}"
        )
    return mapping


def read_excel(path: Path) -> tuple["pd.DataFrame", dict[str, str]]:
    """Laser Excel och validerar kolumner. Stoppar vid fel."""
    if not path.exists():
        raise FileNotFoundError(f"Excel-filen hittades inte: {path}")
    try:
        df = pd.read_excel(path)
    except Exception as exc:
        raise RuntimeError(f"Kunde inte lasa Excel-filen: {exc}")

    mapping = resolve_columns(df)
    return df, mapping


# ===========================================================================
# 8. DUBBLETTKONTROLL
# ===========================================================================
def check_duplicates(rows: list[dict[str, Any]]) -> None:
    """
    Kontrollerar dubbla fakturanummer i indata.
    Stoppar (raise) om dubbletter finns och ALLOW_DUPLICATES=False.
    """
    seen: dict[str, list[int]] = {}
    for r in rows:
        fnr = r["fakturanr"]
        if not fnr:
            continue
        seen.setdefault(fnr, []).append(r["rad"])

    dups = {k: v for k, v in seen.items() if len(v) > 1}
    if dups:
        lines = [f"  Fakturanr {k}: rader {v}" for k, v in dups.items()]
        msg = "Dubbletter av fakturanummer hittades:\n" + "\n".join(lines)
        if ALLOW_DUPLICATES:
            print("VARNING: " + msg + "\n(ALLOW_DUPLICATES=True -> fortsatter anda)")
        else:
            raise ValueError(msg + "\n(ALLOW_DUPLICATES=False -> avbryter)")


# ===========================================================================
# Startchecklista
# ===========================================================================
def print_start_checklist() -> None:
    print(
        "\n"
        "==================== FORE START ====================\n"
        "1. Ta backup i Visma forst.\n"
        "2. Oppna Kundreskontra -> Inbetalningar.\n"
        f"3. Kontrollera att 'Obetalda' ar valt.\n"
        f"4. Kontrollera att Konto = {BANK_ACCOUNT}.\n"
        "5. Klicka i Fakt.nr.-faltet.\n"
        "6. Tryck Enter i terminalen for att starta.\n"
        "====================================================\n"
    )
    print(
        f"Lagen just nu:  DRY_RUN={DRY_RUN}  CONFIRM_EACH_ROW={CONFIRM_EACH_ROW}  "
        f"MAX_ROWS={MAX_ROWS}  ALLOW_DUPLICATES={ALLOW_DUPLICATES}\n"
        f"AUTO_FINALIZE_PAYMENT_BATCH={AUTO_FINALIZE_PAYMENT_BATCH} (maste vara False)\n"
    )


# ===========================================================================
# 7 + 10. REGISTRERINGSFLODE PER RAD
# ===========================================================================
def process_row(
    gui: VismaGui,
    row: dict[str, Any],
    logger: Logger,
) -> str:
    """
    Behandlar en rad. Returnerar en Status.* -konstant.
    Kan kasta RuntimeError vid allvarligt fel dar vi vill stoppa hela koningen.
    """
    log = LogRow(
        rad=row["rad"],
        datum=str(row["datum_raw"]),
        visma_datum=row["visma_datum"],
        avsandare=row["avsandare"],
        betalningsreferens=str(row["referens"]),
        fakturanr=row["fakturanr"],
        belopp=row["visma_belopp"],
    )

    # --- Validering av raddata ---
    if not row["fakturanr"]:
        log.status = Status.FEL
        log.felmeddelande = "Inget fakturanummer kunde extraheras."
        logger.add(log)
        return Status.FEL

    if not row["visma_datum"]:
        log.status = Status.FEL
        log.felmeddelande = "Ogiltigt/saknat datum."
        logger.add(log)
        return Status.FEL

    if row["belopp_dec"] is None:
        log.status = Status.FEL
        log.felmeddelande = "Ogiltigt/saknat belopp."
        logger.add(log)
        return Status.FEL

    # --- Visa vad som ska goras ---
    print(
        f"\n--- Rad {row['rad']} ---\n"
        f"  Datum:      {row['visma_datum']}\n"
        f"  Avsandare:  {row['avsandare']}\n"
        f"  Referens:   {row['referens']}\n"
        f"  Fakturanr:  {row['fakturanr']}\n"
        f"  Belopp:     {row['visma_belopp']}"
    )

    # --- GUI-inmatning ---
    # OBS: Tabbordningen nedan kan behova justeras for din Visma-layout.
    #      Ordning enligt spec: skriv Bet.dag, skriv Fakt.nr, Enter.
    gui.activate_visma_window()

    if not DRY_RUN:
        # (a) Bet.dag  -- justera navigering vid behov (har antas anvandaren
        #     redan star i Fakt.nr; vi lamnar Bet.dag orort om det redan ar satt).
        #     Om du vill skriva datum per rad: fokusera Bet.dag-faltet forst.
        #     Exempel (avkommentera + anpassa):
        #     gui.focus_field_betdag()
        #     gui.select_all_and_clear()
        #     gui.paste_text(row["visma_datum"])

        # (b) Fakt.nr  -- anvandaren star i faltet (enligt checklistan).
        gui.focus_fakturanr_field()
        gui.select_all_and_clear()
        gui.paste_text(row["fakturanr"])
        time.sleep(WAIT_SHORT)

        # (c) Enter -> Visma markerar fakturan och oppnar ev. "Ratt belopp?"
        gui.press_enter()
        time.sleep(WAIT_AFTER_ENTER)
    else:
        print("  [DRY_RUN] Skulle skriva fakturanr + Enter i Visma har.")

    # --- Vanta pa dialogen "Ratt belopp?" ---
    dlg = None
    if not DRY_RUN:
        dlg = gui.wait_for_dialog(DIALOG_TITLE_CONTAINS, timeout=DIALOG_TIMEOUT)
    dtext = gui.dialog_text(dlg) if dlg is not None else ""

    if not DRY_RUN and dlg is None:
        # Dialogen kom inte -> osaker situation, stoppa raden.
        log.status = Status.FEL
        log.felmeddelande = (
            "Dialogen 'Ratt belopp?' visades inte. Kontrollera om fakturan "
            "markerades i listan (kan redan vara betald / fel nummer)."
        )
        logger.add(log)
        return Status.FEL

    # --- Verifiera fakturanr i dialogen ---
    if dtext:
        if row["fakturanr"] not in dtext:
            log.status = Status.FEL
            log.felmeddelande = (
                f"Dialogen namner inte fakturanr {row['fakturanr']}. "
                f"Dialogtext: {dtext[:200]}"
            )
            logger.add(log)
            # Stang dialogen sakert med Avbryt.
            gui.press_esc()
            return Status.FEL

    # --- Belopps-jamforelse (Excel vs ev. Visma-belopp i dialogen) ---
    # I forsta versionen skriver vi ALLTID Excel-beloppet i 'Andra till'
    # och loggar. Vi forsoker ocksa lasa Vismas foreslagna belopp for jamf.
    visma_amount = _extract_amount_from_text(dtext)
    if visma_amount is not None:
        diff = abs(visma_amount - row["belopp_dec"])
        if diff > AMOUNT_TOLERANCE:
            print(
                f"  VARNING: Beloppsdifferens over tolerans!\n"
                f"    Excel: {row['visma_belopp']}   Visma: {visma_amount}   "
                f"diff: {diff} (tolerans {AMOUNT_TOLERANCE})"
            )
            svar = input("  Skriv 'j' for att anda skriva Excel-beloppet och OK, "
                         "annat = hoppa/stopp: ").strip().lower()
            if svar != "j":
                log.status = Status.STOPPAD
                log.felmeddelande = (
                    f"Beloppsdiff {diff} > tolerans. Excel={row['visma_belopp']} "
                    f"Visma={visma_amount}. Anvandaren stoppade."
                )
                logger.add(log)
                gui.press_esc()
                raise RuntimeError("Stoppad av anvandare pga beloppsdifferens.")

    # --- Skriv Excel-belopp i 'Andra till' ---
    if not DRY_RUN:
        # OBS: Fokus antas ligga i 'Andra till'-faltet nar dialogen oppnats.
        #      Om inte: navigera dit (Tab) forst. JUSTERA vid behov.
        gui.select_all_and_clear()
        gui.paste_text(row["visma_belopp"])
        time.sleep(WAIT_SHORT)
    else:
        print(f"  [DRY_RUN] Skulle skriva belopp {row['visma_belopp']} i 'Andra till'.")

    # --- 10. Bekraftelse fore OK ---
    if CONFIRM_EACH_ROW:
        print(
            "\n"
            f"Faktura: {row['fakturanr']}\n"
            f"Datum: {row['visma_datum']}\n"
            f"Belopp: {row['visma_belopp']}\n"
            f"Avsandare: {row['avsandare']}\n"
        )
        svar = input("Tryck Enter for att klicka OK, eller skriv s for att stoppa: ").strip().lower()
        if svar == "s":
            log.status = Status.STOPPAD
            log.felmeddelande = "Anvandaren stoppade fore OK."
            logger.add(log)
            if not DRY_RUN:
                gui.press_esc()
            raise RuntimeError("Stoppad av anvandare.")

    # --- Klicka OK (eller simulera) ---
    if DRY_RUN:
        log.status = Status.DRY_RUN
        log.felmeddelande = "DRY_RUN: inget OK klickades."
        logger.add(log)
        # I DRY_RUN, stang ev. dialog utan att bokfora.
        if dlg is not None:
            gui.press_esc()
        return Status.DRY_RUN

    gui.click_ok_on_dialog(dlg)
    time.sleep(WAIT_AFTER_OK)

    log.status = Status.OK
    logger.add(log)
    return Status.OK


def _extract_amount_from_text(text: str) -> Optional[Decimal]:
    """Forsoker plocka ut ett belopp (svensk decimal) ur dialogtext."""
    if not text:
        return None
    # Leta efter monster som 1 396,00 / 2017,00 / 2 017.00
    candidates = re.findall(r"-?\d[\d\s\xa0.]*,\d{2}", text)
    if not candidates:
        candidates = re.findall(r"-?\d[\d\s\xa0.]*\.\d{2}", text)
    if not candidates:
        return None
    # Ta det sista (oftast 'Andra till'-beloppet).
    return parse_amount(candidates[-1])


# ===========================================================================
# Bygg radlista fran DataFrame
# ===========================================================================
def build_rows(df: "pd.DataFrame", mapping: dict[str, str]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        datum_raw = r[mapping["Datum"]]
        avsandare = r[mapping["Avsandare"]]
        referens = r[mapping["Betalningsreferens"]]
        belopp_raw = r[mapping["Belopp"]]

        rows.append({
            "rad": i,
            "datum_raw": datum_raw,
            "visma_datum": format_visma_date(datum_raw),
            "avsandare": "" if pd.isna(avsandare) else str(avsandare).strip(),
            "referens": "" if pd.isna(referens) else str(referens).strip(),
            "fakturanr": extract_invoice_number(referens),
            "belopp_dec": parse_amount(belopp_raw),
            "visma_belopp": format_visma_amount(belopp_raw),
        })
    return rows


# ===========================================================================
# Huvudflode
# ===========================================================================
def run(excel_path: Path) -> None:
    print(f"Laser Excel: {excel_path}")
    df, mapping = read_excel(excel_path)
    print(f"Kolumnmappning: {mapping}")

    rows = build_rows(df, mapping)
    print(f"Antal rader i filen: {len(rows)}")

    # Dubblettkontroll INNAN nagon registrering.
    check_duplicates(rows)

    # Begransa antal rader vid test.
    if MAX_ROWS is not None:
        rows = rows[:MAX_ROWS]
        print(f"TESTLAGE: kor endast {len(rows)} rad(er) (MAX_ROWS={MAX_ROWS}).")

    logger = Logger(out_dir=excel_path.parent)

    # Anslut till Visma (om inte DRY_RUN utan pywinauto).
    gui = VismaGui()
    if not DRY_RUN:
        gui.connect()
    else:
        print("DRY_RUN: hoppar over anslutning till Visma (ingen inmatning sker).")

    # Startchecklista + bekraftelse.
    print_start_checklist()
    input("Tryck Enter for att starta (Ctrl+C for att avbryta) ... ")

    summary = {Status.OK: 0, Status.DRY_RUN: 0, Status.HOPPAD: 0,
               Status.FEL: 0, Status.STOPPAD: 0}

    try:
        for row in rows:
            try:
                status = process_row(gui, row, logger)
            except RuntimeError as stop_exc:
                # Anvandaren/systemet begarde stopp -> avbryt hela koningen.
                print(f"\nSTOPP: {stop_exc}")
                summary[Status.STOPPAD] += 1
                break
            summary[status] = summary.get(status, 0) + 1
    finally:
        logger.save()

    # Sammanfattning.
    print("\n==================== SAMMANFATTNING ====================")
    for k, v in summary.items():
        print(f"  {k:10s}: {v}")
    print("========================================================")

    # 15. Ingen automatisk slutbokforing.
    print(
        "\nAlla rader ar behandlade.\n"
        "Kontrollera nedre listan 'Till betalning >>>' i Visma.\n"
        "Klicka SJALV pa Bankgiro/Bokfor i Visma om allt stammer.\n"
        f"(AUTO_FINALIZE_PAYMENT_BATCH={AUTO_FINALIZE_PAYMENT_BATCH} - "
        "ingen automatisk bokforing utford.)\n"
    )


# ===========================================================================
# CLI
# ===========================================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Halvautomatisk registrering av inbetalningar i Visma Compact 6."
    )
    parser.add_argument(
        "excel", nargs="?", help="Sokvag till Excel-filen med inbetalningar."
    )
    parser.add_argument(
        "--self-test", action="store_true",
        help="Kor interna tester av parser-funktionerna (ingen Visma-kontakt)."
    )
    args = parser.parse_args()

    if args.self_test:
        _self_test()
        return

    if args.excel:
        excel_path = Path(args.excel).expanduser().resolve()
    else:
        raw = input("Ange sokvag till Excel-filen: ").strip().strip('"')
        excel_path = Path(raw).expanduser().resolve()

    try:
        run(excel_path)
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"\nAVBRUTET: {exc}")
        sys.exit(2)
    except KeyboardInterrupt:
        print("\nAvbrutet av anvandaren (Ctrl+C).")
        sys.exit(130)


# ===========================================================================
# Enkla interna tester (kor: python visma_register_inbetalningar.py --self-test)
# ===========================================================================
def _self_test() -> None:
    print("Kor self-test av parser-funktioner ...\n")

    # extract_invoice_number
    assert extract_invoice_number("49577") == "49577"
    assert extract_invoice_number("Faktura 49577") == "49577"
    assert extract_invoice_number("OCR 49577") == "49577"
    assert extract_invoice_number("Bet ref 49571") == "49571"
    assert extract_invoice_number("") == ""
    assert extract_invoice_number(None) == ""
    print("  extract_invoice_number: OK")

    # format_visma_date
    assert format_visma_date("2026-04-07") == "26-04-07"
    assert format_visma_date(date(2026, 4, 7)) == "26-04-07"
    assert format_visma_date(datetime(2026, 4, 7)) == "26-04-07"
    assert format_visma_date("07/04/2026") == "26-04-07"
    assert format_visma_date("") == ""
    assert format_visma_date(None) == ""
    print("  format_visma_date: OK")

    # parse_amount
    assert parse_amount(2017) == Decimal("2017")
    assert parse_amount("2017,00") == Decimal("2017.00")
    assert parse_amount("2 017,00") == Decimal("2017.00")
    assert parse_amount("1 396,50") == Decimal("1396.50")
    assert parse_amount("") is None
    print("  parse_amount: OK")

    # format_visma_amount
    assert format_visma_amount(2017) == "2017,00"
    assert format_visma_amount(2017.0) == "2017,00"
    assert format_visma_amount("2 017,00") == "2017,00"
    assert format_visma_amount(1396.5) == "1396,50"
    print("  format_visma_amount: OK")

    print("\nAlla self-tester gick igenom.")


if __name__ == "__main__":
    main()
