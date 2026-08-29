#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
visma_register_inbetalningar.py
================================

Halvautomatisk registrering av kundinbetalningar i Visma Compact 6
(Program -> Kundreskontra -> Inbetalningar) baserat pa en CSV-fil.

Indatafil:
    - Standardnamn: data.csv
    - UTF-8-kodad
    - Semikolon (;) som faltavgransare
    - Komma (,) som decimaltecken
    - Kolumnrubriker (exakt dessa):
        Betalningsdatum ; Fakturanr ; Belopp ; Kundnamn ; KundID

VIKTIGT (ekonomisk automation):
  - Scriptet oppnar INTE meny/program sjalv. Anvandaren forbereder Visma manuellt.
  - Scriptet klickar ALDRIG slutlig bokforing / Bankgiro automatiskt i denna version.
  - Standardlage ar sakert: DRY_RUN = True och CONFIRM_EACH_ROW = True.
  - Testa alltid pa en backup forst.

Se instruktioner langst ner i filen (modul-docstring INSTALL/USAGE) samt README-utskrift
via kommandoraden:  python visma_register_inbetalningar.py --help

Beroenden:
    pip install pandas pyautogui pywinauto pyperclip
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
# Beroenden importeras "mjukt" sa att t.ex. CSV-lasning fungerar aven om
# GUI-biblioteken saknas (praktiskt vid utveckling pa annan dator).
# ---------------------------------------------------------------------------
try:
    import pandas as pd
except ImportError:  # pragma: no cover
    print("FEL: 'pandas' saknas. Kor: pip install pandas")
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
# OBS: dessa anvands som REGEX mot fonstertiteln. "R[aä]tt" matchar bade
# "Rätt" (med a-ring) och "Ratt", sa svenska tecken inte stjalper matchningen.
DIALOG_TITLE_CONTAINS = "R[aä]tt belopp"           # dialogen "Rätt belopp?"
DIALOG_DIFFERENS_CONTAINS = "Differens"            # dialogen "Differens"
BANK_ACCOUNT = "1930"                              # forvantat konto (kontrolleras bara visuellt)

# --- Faltetiketter i huvudfonstret (for att hitta ratt inmatningsfalt) ---
BETDAG_LABEL_CONTAINS = "Bet.dag"     # etikett bredvid datum-faltet
FAKTNR_LABEL_CONTAINS = "Fakt.nr"     # etikett bredvid fakturanr-faltet
ANDRA_TILL_LABEL_CONTAINS = "Andra till"  # etikett i "Ratt belopp?"-dialogen ("Ändra till:")

# --- Valfria faltoverstyrningar (satts bara om auto-matchningen missar) ---
# Fasta skarmkoordinater (x, y) att klicka for att fokusera faltet. None = auto.
BETDAG_CLICK_XY: Optional[tuple[int, int]] = None
FAKTNR_CLICK_XY: Optional[tuple[int, int]] = None

# Antal Enter-tryck for att flytta fokus fran Bet.dag till Fakt.nr i Visma.
# (I Visma flyttas fokus framat med Enter; standard ar 2 tryck.)
ENTERS_AFTER_BETDAG = 2

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
WAIT_AFTER_DATE = 0.3      # paus efter att Bet.dag skrivits
DIALOG_TIMEOUT = 5.0       # hur lange vi vantar pa "Ratt belopp?"-dialogen
DIFFERENS_TIMEOUT = 3.0    # hur lange vi vantar pa ev. "Differens"-dialog
DIALOG_CLOSE_TIMEOUT = 8.0 # hur lange vi vantar pa att dialogerna stangs

# --- Test ---
MAX_ROWS: Optional[int] = 3   # kor bara N rader vid test. Satt None for alla rader.

# --- Indatafil ---
DEFAULT_CSV_FILENAME = "data.csv"  # anvands om ingen sokvag anges
CSV_SEPARATOR = ";"                # faltavgransare
CSV_DECIMAL = ","                  # decimaltecken (informativt; belopp lases som text)
CSV_ENCODING = "utf-8"             # filkodning

# --- Fakturanummer-extraktion ---
INVOICE_MIN_DIGITS = 4
INVOICE_MAX_DIGITS = 8

# --- Loggning ---
LOG_PREFIX = "visma_inbetalningar_logg_"

# Forvantade CSV-kolumner (exakt dessa rubriker kravs)
REQUIRED_COLUMNS = ["Betalningsdatum", "Fakturanr", "Belopp", "Kundnamn", "KundID"]
# Alias gor inlasningen tolerant mot sma stavningsvariationer/rubriker.
COLUMN_ALIASES = {
    "Betalningsdatum": ["Betalningsdatum", "Datum", "Bet.dag", "Betalningsdag"],
    "Fakturanr": ["Fakturanr", "Fakturanummer", "Fakt.nr", "Faktura"],
    "Belopp": ["Belopp"],
    "Kundnamn": ["Kundnamn", "Kund", "Avsandare", "Avsändare"],
    "KundID": ["KundID", "Kundid", "Kundnr", "Kund-ID"],
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
    kundnamn: str = ""
    kundid: str = ""
    fakturanr: str = ""
    belopp: str = ""
    status: str = ""
    felmeddelande: str = ""
    tidpunkt: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def as_dict(self) -> dict[str, Any]:
        return {
            "Rad": self.rad,
            "Betalningsdatum": self.datum,
            "VismaDatum": self.visma_datum,
            "Kundnamn": self.kundnamn,
            "KundID": self.kundid,
            "Fakturanr": self.fakturanr,
            "Belopp": self.belopp,
            "Status": self.status,
            "Felmeddelande": self.felmeddelande,
            "Tidpunkt": self.tidpunkt,
        }


class Logger:
    """Samlar loggrader och skriver till CSV bredvid indatafilen."""

    def __init__(self, out_dir: Path):
        self.rows: list[LogRow] = []
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = out_dir / f"{LOG_PREFIX}{stamp}.csv"

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
        # Semikolon + utf-8-sig sa att Excel oppnar CSV:n korrekt (svenska tecken).
        df.to_csv(self.path, index=False, sep=";", encoding="utf-8-sig")
        print(f"\nLogg sparad: {self.path}")


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
        self._dumped: set = set()

    def dump_dialog(self, dlg, tag: str) -> None:
        """
        Skriver dialogens kontrolltrad till fil (en gang per tag) for felsokning.
        Kallas under skarp korning nar dialogen ar oppen -> palitlig ogonblicksbild.
        """
        if dlg is None or tag in self._dumped:
            return
        try:
            out = Path.cwd() / f"visma_inspect_{tag}.txt"
            dlg.print_control_identifiers(depth=4, filename=str(out))
            print(f"  [DEBUG] Dialogens falt sparade till: {out}")
            self._dumped.add(tag)
        except Exception as exc:
            print(f"  [DEBUG] Kunde inte dumpa dialog '{tag}': {exc}")

    # --- Fonsterhantering ---------------------------------------------------
    def connect(self) -> None:
        """Anslut till redan oppet Visma-fonster. Stoppar om det inte hittas."""
        if Application is None:
            raise RuntimeError(
                "pywinauto saknas men kravs for skarp korning (--live). "
                "Installera med:  pip install pywinauto\n"
                "(eller kor utan --live for att bara simulera.)"
            )
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
        """
        Markera allt i aktivt falt och rensa.

        VIKTIGT: Ctrl+A far INTE anvandas i Visma Compact - det oppnar
        Artikelregistret (ser ut som att en ny artikel ska skapas).
        Vi markerar i stallet med End -> Shift+Home och raderar sedan.
        """
        if pyautogui is not None:
            pyautogui.press("end")            # till radslut
            time.sleep(WAIT_SHORT)
            pyautogui.hotkey("shift", "home")  # markera hela faltet bakat
            time.sleep(WAIT_SHORT)
            pyautogui.press("delete")          # rensa markeringen
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

        pat = re.compile(title_contains, re.IGNORECASE)
        deadline = time.time() + timeout
        while time.time() < deadline:
            for w in self._all_top_windows():
                try:
                    t = w.window_text() or ""
                except Exception:
                    continue
                if pat.search(t):
                    return w
            time.sleep(0.2)
        return None

    @staticmethod
    def _all_top_windows() -> list:
        """Alla toppniva-fonster (bade via Desktop och Visma-appen)."""
        wins = []
        try:
            wins.extend(Desktop(backend="uia").windows())
        except Exception:
            pass
        return wins

    def list_open_window_titles(self) -> list[str]:
        """Returnerar titlarna pa alla oppna toppniva-fonster (for felsokning)."""
        titles = []
        for w in self._all_top_windows():
            try:
                t = (w.window_text() or "").strip()
            except Exception:
                t = ""
            if t:
                titles.append(t)
        return titles

    def dump_open_windows(self, tag: str = "windows") -> None:
        """Skriver alla oppna fonstertitlar till fil (for att se dialogens titel)."""
        try:
            out = Path.cwd() / f"visma_inspect_{tag}.txt"
            titles = self.list_open_window_titles()
            out.write_text(
                "Oppna fonstertitlar:\n" + "\n".join(f"  - {t!r}" for t in titles),
                encoding="utf-8",
            )
            print(f"  [DEBUG] Oppna fonster sparade till: {out}")
        except Exception as exc:
            print(f"  [DEBUG] Kunde inte lista fonster: {exc}")

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

    def wait_for_dialog_closed(self, title_contains: str,
                               timeout: float = DIALOG_CLOSE_TIMEOUT) -> bool:
        """
        Vantar tills dialogen med titel *title_contains* INTE langre finns.
        Returnerar True om den stangdes inom timeout, annars False.
        """
        if Desktop is None:
            time.sleep(min(timeout, WAIT_AFTER_OK))
            return True
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                dlg = Desktop(backend="uia").window(
                    title_re=f".*{title_contains}.*"
                )
                if not dlg.exists(timeout=0.2):
                    return True
            except Exception:
                return True
            time.sleep(0.2)
        return False

    @staticmethod
    def _amount_edit_in_dialog(dlg):
        """
        Hittar redigeringsfaltet 'Andra till:' i "Ratt belopp?"-dialogen.
        Strategi: valj det Edit-falt vars innehall ser ut som ett belopp;
        annars forsta Edit-faltet. Returnerar wrapper eller None.
        """
        try:
            edits = dlg.descendants(control_type="Edit")
        except Exception:
            edits = []
        if not edits:
            return None
        for e in edits:
            try:
                if re.search(r"\d", e.window_text() or ""):
                    return e
            except Exception:
                continue
        return edits[0]

    def write_amount_in_dialog(self, dlg, value: str) -> None:
        """
        Skriver *value* i dialogens beloppsfalt ('Andra till:').

        Anvander SAMMA metod som fungerar for Bet.dag/Fakt.nr:
        riktigt musklick pa faltet (click_input) -> rensa (End->Shift+Home)
        -> klistra in via urklipp. (set_edit_text undviks - den kan "lyckas"
        utan att faltet faktiskt andras.)

        Hittas inget falt: faltet ar redan fokuserat nar dialogen oppnas, sa
        vi rensar och klistrar in anda.
        """
        target = value.strip()
        edit = self._amount_edit_in_dialog(dlg)

        if edit is not None:
            try:
                edit.click_input()   # riktigt klick fokuserar faltet
            except Exception:
                pass
            time.sleep(WAIT_SHORT)

        # Rensa och skriv (samma som fungerar for datum/fakturanr).
        self.select_all_and_clear()
        self.paste_text(target)

    # --- Falt-navigering i huvudfonstret ------------------------------------
    def _find_edit_by_label(self, label_contains: str):
        """
        Hittar inmatningsfaltet (Edit) som hor till etiketten *label_contains*
        i huvudfonstret, via geometrisk matchning: faltet ligger pa samma
        rad, till hoger om etiketten. Returnerar wrapper eller None.
        """
        if self.main_window is None:
            return None
        try:
            descendants = self.main_window.descendants()
        except Exception:
            return None

        # Kandidat-etiketter (text som innehaller label_contains).
        labels = []
        edits = []
        for c in descendants:
            try:
                txt = c.window_text() or ""
                ctype = c.friendly_class_name()
            except Exception:
                continue
            if label_contains.lower() in txt.lower():
                labels.append(c)
            if ctype in ("Edit", "ComboBox"):
                edits.append(c)
        if not labels or not edits:
            return None

        for lbl in labels:
            try:
                lr = lbl.rectangle()
            except Exception:
                continue
            lbl_mid_y = (lr.top + lr.bottom) / 2
            best = None
            best_dx = 10 ** 9
            for e in edits:
                try:
                    er = e.rectangle()
                except Exception:
                    continue
                # Samma rad (vertikal overlappning/narhet) och till hoger.
                same_row = er.top - 6 <= lbl_mid_y <= er.bottom + 6
                dx = er.left - lr.right
                if same_row and dx >= -5 and dx < best_dx:
                    best_dx = dx
                    best = e
            if best is not None:
                return best
        return None

    def focus_field_by_label(self, label_contains: str, click_xy=None,
                             what: str = "faltet"):
        """
        Fokuserar ett falt i huvudfonstret.

        Ordning:
          1) Fast koordinat (click_xy) om angiven.
          2) Automatisk etikett->falt-matchning.
        Stoppar (RuntimeError) om faltet inte kan hittas - da fortsatter vi
        INTE till nasta faktura.
        """
        # 1) Fast koordinat.
        if click_xy is not None and pyautogui is not None:
            pyautogui.click(click_xy[0], click_xy[1])
            time.sleep(WAIT_SHORT)
            return None

        # 2) Auto via etikett.
        edit = self._find_edit_by_label(label_contains)
        if edit is None:
            raise RuntimeError(
                f"Hittade inte {what} (etikett '{label_contains}') i Visma-fonstret. "
                f"Stoppar. Ange en fast koordinat i konfigurationen om det behovs."
            )
        try:
            edit.click_input()  # klick fokuserar och satter markor
        except Exception as exc:
            raise RuntimeError(f"Kunde inte fokusera {what}: {exc}")
        time.sleep(WAIT_SHORT)
        return edit

    def focus_betdag_field(self):
        return self.focus_field_by_label(
            BETDAG_LABEL_CONTAINS, click_xy=BETDAG_CLICK_XY, what="Bet.dag-faltet"
        )

    def focus_fakturanr_field(self):
        return self.focus_field_by_label(
            FAKTNR_LABEL_CONTAINS, click_xy=FAKTNR_CLICK_XY, what="Fakt.nr-faltet"
        )

    def clear_and_write(self, text: str) -> None:
        """Rensar aktivt falt (utan Ctrl+A) och skriver text via urklipp."""
        self.select_all_and_clear()
        self.paste_text(text)


# ===========================================================================
# CSV-lasning + kolumnvalidering
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
            "CSV-filen saknar obligatoriska kolumner: "
            + ", ".join(missing)
            + f"\nHittade kolumner: {list(df.columns)}"
        )
    return mapping


def read_csv(path: Path) -> tuple["pd.DataFrame", dict[str, str]]:
    """
    Laser CSV och validerar kolumner. Stoppar vid fel.

    Forutsattningar (se KONFIGURATION):
        - UTF-8-kodad
        - Semikolon (;) som faltavgransare
        - Komma (,) som decimaltecken

    Alla falt lases som text (dtype=str) sa att befintlig parsning av
    datum och belopp (svensk decimal, tusenavskiljare) fungerar oforandrat.
    """
    if not path.exists():
        raise FileNotFoundError(f"CSV-filen hittades inte: {path}")
    if path.stat().st_size == 0:
        raise ValueError(f"CSV-filen ar tom: {path}")

    try:
        df = pd.read_csv(
            path,
            sep=CSV_SEPARATOR,
            dtype=str,            # las allt som text; parse_amount/format_visma_date hanterar resten
            encoding=CSV_ENCODING,
            keep_default_na=True,
            skip_blank_lines=True,
        )
    except pd.errors.EmptyDataError:
        raise ValueError(f"CSV-filen saknar data (inga kolumner/rader): {path}")
    except UnicodeDecodeError as exc:
        raise RuntimeError(
            f"Kunde inte avkoda CSV-filen som {CSV_ENCODING}: {exc}"
        )
    except Exception as exc:
        raise RuntimeError(f"Kunde inte lasa CSV-filen: {exc}")

    if df.empty:
        raise ValueError(f"CSV-filen innehaller inga datarader: {path}")

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
        "5. Lat Inbetalningsfonstret vara oppet och synligt (falten Bet.dag/Fakt.nr\n"
        "   fokuseras automatiskt av scriptet - ror inte mus/tangentbord under korning).\n"
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
        kundnamn=row["kundnamn"],
        kundid=row["kundid"],
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
        f"  Kundnamn:   {row['kundnamn']}\n"
        f"  KundID:     {row['kundid']}\n"
        f"  Fakturanr:  {row['fakturanr']}\n"
        f"  Belopp:     {row['visma_belopp']}"
    )

    # --- DRY_RUN: simulera hela flodet, ror aldrig Visma ---
    if DRY_RUN:
        print(
            "  [DRY_RUN] Skulle utfora:\n"
            f"    1) Bet.dag  = {row['visma_datum']}\n"
            f"    2) {ENTERS_AFTER_BETDAG}x Enter -> Fakt.nr = {row['fakturanr']} + Enter\n"
            f"    3) vanta pa 'Ratt belopp?'\n"
            f"    4) 'Andra till:' = {row['visma_belopp']}  + Enter\n"
            "    5) ev. 'Differens' -> 'Restbelopp pa fakturan' + Enter"
        )
        log.status = Status.DRY_RUN
        log.felmeddelande = "DRY_RUN: inget skrevs i Visma."
        logger.add(log)
        return Status.DRY_RUN

    # --- Fokusera Visma-fonstret ---
    gui.activate_visma_window()

    # --- (valfri) bekraftelse innan vi ror Visma ---
    if CONFIRM_EACH_ROW:
        print(
            f"\nKlar att registrera i Visma:\n"
            f"  Bet.dag:  {row['visma_datum']}\n"
            f"  Fakt.nr:  {row['fakturanr']}\n"
            f"  Belopp:   {row['visma_belopp']}\n"
            f"  Kund:     {row['kundnamn']} ({row['kundid']})\n"
        )
        svar = input("Tryck Enter for att kora raden, eller skriv s for att stoppa: ").strip().lower()
        if svar == "s":
            log.status = Status.STOPPAD
            log.felmeddelande = "Anvandaren stoppade fore registrering."
            logger.add(log)
            raise RuntimeError("Stoppad av anvandare.")

    def _stoppa(felmeddelande: str, stang_dialog: bool = False) -> None:
        """Loggar fel och stoppar hela koningen (fortsatter INTE till nasta faktura)."""
        log.status = Status.FEL
        log.felmeddelande = felmeddelande
        logger.add(log)
        if stang_dialog:
            gui.press_esc()
        raise RuntimeError(felmeddelande)

    # --- Steg 2: skriv Betalningsdatum i Bet.dag (AA-MM-DD) ---
    try:
        gui.focus_betdag_field()
    except RuntimeError as exc:
        _stoppa(str(exc))
    gui.clear_and_write(row["visma_datum"])
    time.sleep(WAIT_AFTER_DATE)

    # --- Steg 3: flytta till Fakt.nr med Enter, skriv Fakturanr, tryck Enter ---
    # I Visma flyttas fokus framat med Enter. Fran Bet.dag kravs (som standard)
    # tva Enter for att hamna i Fakt.nr-faltet (se ENTERS_AFTER_BETDAG).
    for _ in range(ENTERS_AFTER_BETDAG):
        gui.press_enter()
        time.sleep(WAIT_SHORT)
    gui.clear_and_write(row["fakturanr"])
    time.sleep(WAIT_SHORT)
    gui.press_enter()  # bekraftar Fakt.nr -> oppnar "Ratt belopp?"
    time.sleep(WAIT_AFTER_ENTER)

    # --- Steg 4: vanta tills "Ratt belopp?" oppnas ---
    dlg = gui.wait_for_dialog(DIALOG_TITLE_CONTAINS, timeout=DIALOG_TIMEOUT)
    if dlg is None:
        # Felsokning: skriv ut alla oppna fonstertitlar sa vi ser dialogens
        # verkliga titel (dialogen syns pa skarmen men matchades inte).
        gui.dump_open_windows("windows")
        print("  [DEBUG] Oppna fonster just nu:")
        for t in gui.list_open_window_titles():
            print(f"    - {t!r}")
        _stoppa(
            "Dialogen 'Ratt belopp?' oppnades inte inom tidsgransen. Kontrollera "
            f"att fakturanr {row['fakturanr']} finns bland Obetalda. Stoppar."
        )
    dtext = gui.dialog_text(dlg)

    # Kontroll: galler dialogen ratt faktura?
    if dtext and row["fakturanr"] not in dtext:
        _stoppa(
            f"'Ratt belopp?' namner inte fakturanr {row['fakturanr']}. "
            f"Dialogtext: {dtext[:200]}. Stoppar.",
            stang_dialog=True,
        )

    # Informativ loggning av ev. beloppsdifferens (hanteras via Differens-dialogen).
    visma_amount = _extract_amount_from_text(dtext)
    if visma_amount is not None:
        diff = abs(visma_amount - row["belopp_dec"])
        if diff > AMOUNT_TOLERANCE:
            print(
                f"  INFO: Visma foreslar {visma_amount}, Excel-belopp "
                f"{row['visma_belopp']} (diff {diff}). Skriver Excel-beloppet; "
                "ev. differens hanteras som 'Restbelopp pa fakturan'."
            )

    # --- Steg 5: skriv Excel-belopp i 'Andra till:' och tryck Enter ---
    # Felsokning: spara dialogens falt till fil forsta gangen (dialogen ar oppen nu).
    gui.dump_dialog(dlg, "ratt_belopp")
    gui.write_amount_in_dialog(dlg, row["visma_belopp"])
    time.sleep(WAIT_SHORT)
    gui.press_enter()  # bekraftar 'Ratt belopp?'
    time.sleep(WAIT_AFTER_OK)

    # --- Steg 6: ev. "Differens" -> behall 'Restbelopp pa fakturan' + Enter ---
    diff_dlg = gui.wait_for_dialog(DIALOG_DIFFERENS_CONTAINS, timeout=DIFFERENS_TIMEOUT)
    if diff_dlg is not None:
        diff_text = gui.dialog_text(diff_dlg)
        if diff_text and row["fakturanr"] not in diff_text:
            _stoppa(
                f"'Differens' namner inte fakturanr {row['fakturanr']}. "
                f"Dialogtext: {diff_text[:200]}. Stoppar.",
                stang_dialog=True,
            )
        print("  Differens-dialog: behaller 'Restbelopp pa fakturan' och bekraftar.")
        gui.click_ok_on_dialog(diff_dlg)
        time.sleep(WAIT_AFTER_OK)
        # --- Steg 7: vanta tills Differens stangts ---
        if not gui.wait_for_dialog_closed(DIALOG_DIFFERENS_CONTAINS,
                                          timeout=DIALOG_CLOSE_TIMEOUT):
            _stoppa("'Differens'-dialogen stangdes inte. Stoppar.")

    # --- Steg 7: sakerstall att 'Ratt belopp?' stangts innan nasta rad ---
    if not gui.wait_for_dialog_closed(DIALOG_TITLE_CONTAINS, timeout=DIALOG_CLOSE_TIMEOUT):
        _stoppa("'Ratt belopp?'-dialogen stangdes inte. Stoppar.")

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
# Inspektionslage: lista Vismas falt/dialogkontroller (for felsokning)
# ===========================================================================
def inspect_windows() -> None:
    """
    Skriver ut kontrolltradet for huvudfonstret och for dialogerna
    'Ratt belopp?' / 'Differens'. Anvands for att hitta exakt vilket falt
    beloppet ska skrivas i. Kraver pywinauto.
    """
    if Application is None or Desktop is None:
        print("pywinauto saknas. Kor: pip install pywinauto")
        return

    # OBS: Vi listar ENDAST dialogerna - inte huvudfonstret (det innehaller hela
    # fakturalistan och blir enormt). Beloppsfaltet 'Andra till:' finns i dialogen.
    print(
        "\nINSPEKTIONSLAGE (endast dialoger).\n"
        "Huvudfonstret listas inte (det innehaller alla fakturor)."
    )

    for title in (DIALOG_TITLE_CONTAINS, DIALOG_DIFFERENS_CONTAINS):
        input(
            f"\n>>> Oppna dialogen '{title}' i Visma (registrera en faktura manuellt),\n"
            f"    Alt-Tab sedan tillbaka hit och tryck Enter (eller Enter for att hoppa over) ... "
        )
        try:
            dlg = Desktop(backend="uia").window(
                title_re=f".*{title}.*"
            )
            if dlg.exists(timeout=1.0):
                safe = re.sub(r"[^A-Za-z0-9]+", "_", title).strip("_").lower()
                out_file = Path.cwd() / f"visma_inspect_{safe}.txt"
                print(f"\n========== DIALOG: {title} ==========")
                # Skriv till fil (och skarm). Dialogen ar liten -> grund djup racker.
                dlg.print_control_identifiers(depth=4, filename=str(out_file))
                dlg.print_control_identifiers(depth=4)
                print(f"\n>>> Sparat till: {out_file}")
            else:
                print(f"(Hittade ingen dialog med titel som innehaller '{title}'.)")
        except Exception as exc:
            print(f"({title}: {exc})")

    print("\nKlar. Filerna 'visma_inspect_*.txt' skapades i mappen - sag till sa laser jag dem.")


# ===========================================================================
# Bygg radlista fran DataFrame
# ===========================================================================
def build_rows(df: "pd.DataFrame", mapping: dict[str, str]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        datum_raw = r[mapping["Betalningsdatum"]]
        fakturanr_raw = r[mapping["Fakturanr"]]
        belopp_raw = r[mapping["Belopp"]]
        kundnamn_raw = r[mapping["Kundnamn"]]
        kundid_raw = r[mapping["KundID"]]

        # Fakturanr kommer nu i egen kolumn. Anvand det direkt om det ar rena
        # siffror; annars falla tillbaka pa extraktion (t.ex. "Faktura 49577").
        fnr = "" if pd.isna(fakturanr_raw) else str(fakturanr_raw).strip()
        fakturanr = fnr if fnr.isdigit() else extract_invoice_number(fnr)

        rows.append({
            "rad": i,
            "datum_raw": datum_raw,
            "visma_datum": format_visma_date(datum_raw),
            "kundnamn": "" if pd.isna(kundnamn_raw) else str(kundnamn_raw).strip(),
            "kundid": "" if pd.isna(kundid_raw) else str(kundid_raw).strip(),
            "fakturanr": fakturanr,
            "belopp_dec": parse_amount(belopp_raw),
            "visma_belopp": format_visma_amount(belopp_raw),
        })
    return rows


# ===========================================================================
# Interaktivt urval: fran vilket fakturanr + hur manga
# ===========================================================================
def ask_invoice_selection(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Fragar anvandaren fran vilket fakturanr registreringen ska borja och
    hur manga fakturor som ska registreras. Returnerar en filtrerad radlista
    i ursprunglig ordning.

    Regler:
        - Tomt startvarde  -> borja fran forsta raden.
        - Tomt antal       -> anvand MAX_ROWS om satt, annars alla kvarvarande.
        - Antal 0 (noll)   -> alla rader fran startpunkten.
        - Ogiltig inmatning -> fraga igen (ingen krasch).
    """
    if not rows:
        return rows

    fakturanummer = [r["fakturanr"] for r in rows]

    # 1) Fran vilket fakturanr ska vi borja?
    start_idx = 0
    while True:
        svar = input(
            "Fran vilket fakturanr vill du borja? (Enter = forsta raden): "
        ).strip()
        if not svar:
            start_idx = 0
            break
        traff = next((i for i, f in enumerate(fakturanummer) if f == svar), None)
        if traff is None:
            print(f"  Hittade inget fakturanr '{svar}' i filen. Forsok igen.")
            continue
        start_idx = traff
        break

    kvar = len(rows) - start_idx

    # 2) Hur manga fakturor ska registreras?
    default_antal = MAX_ROWS if MAX_ROWS is not None else kvar
    if default_antal > kvar:
        default_antal = kvar
    while True:
        svar = input(
            f"Hur manga fakturor vill du registrera? "
            f"(Enter = {default_antal}, 0 = alla {kvar} kvarvarande): "
        ).strip()
        if not svar:
            antal = default_antal
            break
        try:
            antal = int(svar)
        except ValueError:
            print("  Ange ett heltal.")
            continue
        if antal < 0:
            print("  Antalet kan inte vara negativt.")
            continue
        break

    if antal <= 0 or antal > kvar:
        antal = kvar

    urval = rows[start_idx:start_idx + antal]
    print(
        f"\nUrval: {len(urval)} faktura(or) - fran rad {urval[0]['rad']} "
        f"(fakturanr {urval[0]['fakturanr']}) till rad {urval[-1]['rad']} "
        f"(fakturanr {urval[-1]['fakturanr']})."
    )
    return urval


# ===========================================================================
# Huvudflode
# ===========================================================================
def run(csv_path: Path) -> None:
    print(f"Laser CSV: {csv_path}")
    df, mapping = read_csv(csv_path)
    print(f"Kolumnmappning: {mapping}")

    rows = build_rows(df, mapping)
    print(f"Antal rader i filen: {len(rows)}")

    # Dubblettkontroll INNAN nagon registrering.
    check_duplicates(rows)

    # Interaktivt urval: fran vilket fakturanr och hur manga fakturor.
    rows = ask_invoice_selection(rows)
    if not rows:
        print("Inga rader att registrera. Avbryter.")
        return

    logger = Logger(out_dir=csv_path.parent)

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
        "csv", nargs="?",
        help=f"Sokvag till CSV-filen med inbetalningar (standard: {DEFAULT_CSV_FILENAME})."
    )
    parser.add_argument(
        "--self-test", action="store_true",
        help="Kor interna tester av parser-funktionerna (ingen Visma-kontakt)."
    )
    parser.add_argument(
        "--inspect", action="store_true",
        help="Listar Vismas falt/dialogkontroller for felsokning (kraver oppet Visma)."
    )
    parser.add_argument(
        "--live", action="store_true",
        help="Kor SKARPT mot Visma (DRY_RUN=False). Utan flaggan simuleras allt "
             "och inget skrivs i Visma."
    )
    parser.add_argument(
        "--rader", type=int, metavar="N",
        help="Kor endast N rader (overstyr MAX_ROWS). Ange 0 for alla rader."
    )
    args = parser.parse_args()

    if args.self_test:
        _self_test()
        return

    if args.inspect:
        inspect_windows()
        return

    # --- Overstyr sakerhetslagen fran kommandoraden ---
    global DRY_RUN, MAX_ROWS
    if args.live:
        DRY_RUN = False
    if args.rader is not None:
        MAX_ROWS = None if args.rader <= 0 else args.rader

    if not DRY_RUN:
        print(
            "\n*** SKARP KORNING (--live): scriptet kommer att skriva i Visma. ***\n"
            f"    Rader som kors: {'ALLA' if MAX_ROWS is None else MAX_ROWS}   "
            f"CONFIRM_EACH_ROW={CONFIRM_EACH_ROW}\n"
        )
        if input("    Skriv JA for att fortsatta: ").strip() != "JA":
            print("Avbrutet (ingen skarp korning).")
            return

    if args.csv:
        raw = args.csv
    else:
        raw = input(
            f"Ange sokvag till CSV-filen [{DEFAULT_CSV_FILENAME}]: "
        ).strip().strip('"')
        if not raw:
            raw = DEFAULT_CSV_FILENAME
    csv_path = Path(raw).expanduser().resolve()

    try:
        run(csv_path)
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
