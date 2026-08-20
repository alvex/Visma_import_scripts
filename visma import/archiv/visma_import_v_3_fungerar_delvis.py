import json
import time
import ctypes
import re

import pandas as pd
import pyautogui
import pyperclip

from ctypes import wintypes
from pathlib import Path
from datetime import datetime


# =============================================================================
# VISMA COMPACT - AUTOMATISK KUNDFAKTURAREGISTRERING
# =============================================================================
#
# HUVUDPRINCIPER
# --------------
# * Fälten Kund ID, Bokf.dag och Fak.belopp hittas via KALIBRERADE punkter
#   (sparade i visma_config.json) - inte via gissade fasta koordinater.
# * Efter varje klick VERIFIERAS fokus via Windows API (GetGUIThreadInfo)
#   mot den kontroll som kalibrerades. Scriptet antar aldrig att ett klick
#   hamnade rätt - matchar inte fokus, avbryts fakturan (inget skrivs, inget
#   Num+ skickas).
# * Fakturanumret läses INTE från skärmen. Användaren anger startnummer och
#   antal; hela serien kontrolleras mot Excel före körning och per rad.
#
# VIKTIGT OM Ctrl+A
# -----------------
# I Visma Compact är Ctrl+A en genväg som ÖPPNAR ARTIKELREGISTRET - inte
# "markera allt". Därför markeras fält i stället med End -> Shift+Home innan
# nytt värde skrivs. (Samma effekt, utan att trigga Artikelregistret.)
#
# VISMA-FLÖDE PER FAKTURA
# -----------------------
#   Kund ID   -> klick (kalibrerat) -> verifiera fokus -> skriv -> Tab (ladda kund)
#   Bokf.dag  -> klick (kalibrerat) -> verifiera fokus -> markera -> skriv datum
#   Fak.belopp-> klick (kalibrerat) -> verifiera fokus -> markera -> skriv belopp
#   -> exakt 5 x Tab
#   -> Num+
#
# Tab används ALDRIG för att navigera mellan Bokf.dag och Fak.belopp.
#
# =============================================================================


# =============================================================================
# LÄGEN
# =============================================================================

# Visa varje GUI-steg i terminalen.
DEBUG = True

# Testläge: fyll fälten men gör INTE 5 Tab och INTE Num+.
# Fakturan sparas alltså inte - för visuell kontroll innan skarp körning.
DRY_RUN = True


# =============================================================================
# GRUNDINSTÄLLNINGAR
# =============================================================================

pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = True

# Vänta efter att KundID skrivits så Visma hinner ladda kunden.
CUSTOMER_LOAD_DELAY = 0.8

# Vänta efter Num+ så Visma hinner spara.
SAVE_DELAY = 1.5

# Exakt 5 Tab efter Fak.belopp.
TABS_AFTER_BELOPP = 5

# Max antal fakturor per körning.
MAX_INVOICES_PER_RUN = 50

# Tolerans (pixlar) vid jämförelse av kontrollens rektangel vid fokusverifiering.
CTRL_RECT_TOLERANCE = 8

# Fältordning för kalibrering och registrering.
FIELD_ORDER = ["Kund ID", "Bokf.dag", "Fak.belopp"]

# Config-fil för kalibrerade positioner.
CONFIG_PATH = Path(__file__).parent / "visma_config.json"


# =============================================================================
# WINDOWS API
# =============================================================================

user32 = ctypes.windll.user32

VK_ADD = 0x6B
KEYEVENTF_KEYUP = 0x0002
SW_RESTORE = 9


class POINT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
    ]


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),
    ]


user32.WindowFromPoint.argtypes = [POINT]
user32.WindowFromPoint.restype = wintypes.HWND

# WM_GETTEXT/WM_SETTEXT: läser/sätter en kontrolls textinnehåll direkt.
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT = 0x000C

user32.SendMessageW.restype = ctypes.c_ssize_t
user32.SendMessageW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    ctypes.c_size_t,
    wintypes.LPWSTR,
]


# =============================================================================
# FORMATERING
# =============================================================================

def format_date(value):
    """
    Konverterar Fakturadatum från Excel till Visma-format YY-MM-DD.

    Exempel:
        2026-06-15 -> 26-06-15
        2026-08-17 -> 26-08-17

    Returnerar "" om värdet saknas eller inte kan tolkas som ett datum.
    Vi returnerar ALDRIG ett oformaterat datum - Visma kräver YY-MM-DD.
    """

    if pd.isna(value):
        return ""

    try:
        dt = pd.to_datetime(value)
    except Exception:
        return ""

    if pd.isna(dt):
        return ""

    return dt.strftime("%y-%m-%d")


def excel_date_display(value):
    """
    Visar Excel-datumet i läsbar form (utan tid), för utskrift.
    """

    if pd.isna(value):
        return "<tomt>"

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    try:
        dt = pd.to_datetime(value)
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    return str(value).strip()


def format_amount(value):
    """Konverterar belopp till svenskt decimalformat (1949 -> 1949,00)."""

    if pd.isna(value):
        return "0,00"

    try:
        amount = float(value)
        return f"{amount:.2f}".replace(".", ",")
    except Exception:
        return str(value).strip().replace(".", ",")


def normalize_customer_id(value):
    """Normaliserar KundID/personnummer."""

    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def normalize_invoice_number(value):
    """Normaliserar fakturanummer till en ren sträng (4545.0 -> '4545')."""

    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()

    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass

    return text


def parse_invoice_number(value):
    """Returnerar fakturanumret som heltal, eller None om ogiltigt."""

    text = normalize_invoice_number(value)

    if not text or not re.fullmatch(r"\d+", text):
        return None

    return int(text)


# =============================================================================
# WINDOWS-FÖNSTER
# =============================================================================

def get_window_title(hwnd):
    """Hämtar titeln på ett fönster."""

    if not hwnd:
        return ""

    length = user32.GetWindowTextLengthW(hwnd)

    if length <= 0:
        return ""

    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)

    return buffer.value.strip()


def get_foreground_window_title():
    """Returnerar titeln på aktivt fönster."""

    return get_window_title(user32.GetForegroundWindow())


def get_window_rect(hwnd):
    """Hämtar ett fönsters/en kontrolls rektangel."""

    rect = wintypes.RECT()

    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None

    return rect


def get_class_name(hwnd):
    """Hämtar klassnamnet på en kontroll."""

    if not hwnd:
        return ""

    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)

    return buffer.value


def get_focused_control():
    """
    Returnerar hwnd för den kontroll som har tangentbordsfokus i det
    aktiva fönstrets tråd, eller 0.
    """

    gui = GUITHREADINFO()
    gui.cbSize = ctypes.sizeof(gui)

    thread_id = user32.GetWindowThreadProcessId(
        user32.GetForegroundWindow(),
        None,
    )

    if not user32.GetGUIThreadInfo(thread_id, ctypes.byref(gui)):
        return 0

    return gui.hwndFocus or 0


def get_control_at_point(x, y):
    """Returnerar hwnd för kontrollen under en skärmpunkt (utan att klicka)."""

    return user32.WindowFromPoint(POINT(x, y)) or 0


def wm_get_text(hwnd):
    """
    Läser en kontrolls textinnehåll via WM_GETTEXT.

    Returnerar strängen (kan vara tom), eller None om hwnd saknas.
    """

    if not hwnd:
        return None

    length = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, None)

    if length <= 0:
        return ""

    buffer = ctypes.create_unicode_buffer(length + 1)

    user32.SendMessageW(hwnd, WM_GETTEXT, length + 1, buffer)

    return buffer.value


def wm_set_text(hwnd, text):
    """Sätter en kontrolls textinnehåll direkt via WM_SETTEXT."""

    if not hwnd:
        return False

    return bool(user32.SendMessageW(hwnd, WM_SETTEXT, 0, str(text)))


# =============================================================================
# HITTA / AKTIVERA VISMA
# =============================================================================

def find_visma_window():
    """Letar efter öppet Visma-fönster (prioriterar Kundbokning)."""

    windows = []

    CALLBACK = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HWND,
        wintypes.LPARAM,
    )

    @CALLBACK
    def enum_proc(hwnd, lparam):

        if not user32.IsWindowVisible(hwnd):
            return True

        title = get_window_title(hwnd)

        if not title:
            return True

        lower_title = title.lower()

        if "visma" not in lower_title:
            return True

        score = 1

        if "compact" in lower_title:
            score += 10
        if "kundreskontra" in lower_title:
            score += 10
        if "kundbokning" in lower_title:
            score += 20

        windows.append((score, hwnd, title))

        return True

    user32.EnumWindows(enum_proc, 0)

    if not windows:
        return None, ""

    windows.sort(key=lambda item: item[0], reverse=True)

    _, hwnd, title = windows[0]

    return hwnd, title


def activate_visma_window():
    """Aktiverar Visma-fönstret."""

    hwnd, _ = find_visma_window()

    if not hwnd:
        raise RuntimeError(
            "Kunde inte hitta Visma Compact. "
            "Öppna Kundreskontra -> Kundbokning."
        )

    user32.ShowWindow(hwnd, SW_RESTORE)
    time.sleep(0.1)
    user32.BringWindowToTop(hwnd)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.35)

    if "visma" not in get_foreground_window_title().lower():

        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.4)

    if "visma" not in get_foreground_window_title().lower():
        raise RuntimeError(
            "Kunde inte aktivera Visma. "
            f"Aktivt fönster är '{get_foreground_window_title()}'."
        )

    return hwnd


def verify_visma_is_active():
    """Stoppar importen om Visma tappat fokus."""

    title = get_foreground_window_title()

    if "visma" not in title.lower():
        raise RuntimeError(
            f"Visma har tappat fokus. Aktivt fönster är '{title}'."
        )


# =============================================================================
# KONFIGURATION (KALIBRERADE POSITIONER)
# =============================================================================

def load_config():
    """Läser visma_config.json, eller None om den saknas/är trasig."""

    if not CONFIG_PATH.exists():
        return None

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception:
        return None

    fields = config.get("fields", {})

    for name in FIELD_ORDER:
        if name not in fields:
            return None

    return config


def save_config(config):
    """Sparar konfigurationen."""

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def compute_point(field_cfg, win_rect):
    """Räknar ut absolut skärmposition för ett fält utifrån aktuellt fönster."""

    width = win_rect.right - win_rect.left

    x = win_rect.left + int(width * field_cfg["x_ratio"])
    y = win_rect.top + field_cfg["y_offset"]

    return x, y


# =============================================================================
# KALIBRERING
# =============================================================================

def calibrate():
    """
    Kalibrerar Kund ID, Bokf.dag och Fak.belopp.

    Användaren håller muspekaren mitt i fältet och trycker Enter.
    Positionen sparas relativt Visma-fönstret. Om Visma exponerar
    kontrollen (child window) sparas även dess klass + rektangel för
    fokusverifiering under körning.
    """

    print()
    print("=" * 70)
    print("KALIBRERING AV VISMA")
    print("=" * 70)

    hwnd, title = find_visma_window()

    if not hwnd:
        print("\nFEL: Hittade inget Visma-fönster.")
        print("Öppna Visma -> Kundreskontra -> Kundbokning och försök igen.")
        return

    activate_visma_window()

    win_rect = get_window_rect(hwnd)

    if not win_rect:
        print("\nFEL: Kunde inte läsa Visma-fönstrets position.")
        return

    width = win_rect.right - win_rect.left

    print(f"\nVisma hittat: {title}")
    print("Se till att Visma är MAXIMERAT och att fakturaformuläret visas.\n")

    fields = {}

    for name in FIELD_ORDER:

        try:
            input(
                f'Placera muspekaren mitt i fältet "{name}" '
                f"och tryck Enter..."
            )
        except EOFError:
            print("Avbrutet.")
            return

        x, y = pyautogui.position()

        x_ratio = (x - win_rect.left) / width
        y_offset = y - win_rect.top

        # Försök identifiera kontrollen under punkten (utan att klicka).
        ctrl_hwnd = get_control_at_point(x, y)

        ctrl_class = None
        ctrl_rel_rect = None

        if ctrl_hwnd and ctrl_hwnd != hwnd:

            ctrl_rect = get_window_rect(ctrl_hwnd)

            if ctrl_rect:
                ctrl_class = get_class_name(ctrl_hwnd)
                ctrl_rel_rect = [
                    ctrl_rect.left - win_rect.left,
                    ctrl_rect.top - win_rect.top,
                    ctrl_rect.right - win_rect.left,
                    ctrl_rect.bottom - win_rect.top,
                ]

        fields[name] = {
            "x_ratio": round(x_ratio, 4),
            "y_offset": int(y_offset),
            "ctrl_class": ctrl_class,
            "ctrl_rel_rect": ctrl_rel_rect,
        }

        verify_note = (
            "fokusverifiering PÅ"
            if ctrl_rel_rect
            else "fokusverifiering AV (positionsläge)"
        )

        print(f"  Position sparad: X={x} Y={y}  ({verify_note})")

    save_config({"fields": fields})

    print()
    print("=" * 70)
    print(f"Kalibrering sparad: {CONFIG_PATH}")
    print("=" * 70)
    print()


# =============================================================================
# FOKUS + SKRIVNING
# =============================================================================

def rects_match(rel_rect, expected, tolerance):
    """Jämför två relativa rektanglar inom tolerans."""

    if not expected or not rel_rect:
        return False

    return all(
        abs(a - b) <= tolerance
        for a, b in zip(rel_rect, expected)
    )


def focus_field(name, field_cfg, win_rect):
    """
    Klickar på fältets kalibrerade punkt och verifierar fokus.

    Returnerar (x, y, focus_hwnd) - focus_hwnd är den verifierade
    kontrollen (0 i positionsläge).

    Höjer RuntimeError om fokus inte kan verifieras - då skrivs inget
    och ingen faktura sparas.
    """

    x, y = compute_point(field_cfg, win_rect)

    expected = field_cfg.get("ctrl_rel_rect")
    main_hwnd, _ = find_visma_window()

    last_info = ""

    for attempt in range(2):

        verify_visma_is_active()

        pyautogui.click(x, y)
        time.sleep(0.20)

        verify_visma_is_active()

        # -------------------------------------------------------------
        # Positionsläge: kontrollen kunde inte identifieras vid kalibrering.
        # Vi litar då på den kalibrerade punkten (spec pkt 3).
        # -------------------------------------------------------------
        if not expected:
            if DEBUG:
                print(f"    {name} -> klick X={x} Y={y} (positionsläge)")
            return x, y, 0

        # -------------------------------------------------------------
        # Verifiera att den fokuserade kontrollen matchar den kalibrerade.
        # -------------------------------------------------------------
        focus_hwnd = get_focused_control()

        if focus_hwnd and focus_hwnd != main_hwnd:

            ctrl_rect = get_window_rect(focus_hwnd)

            if ctrl_rect:
                rel = [
                    ctrl_rect.left - win_rect.left,
                    ctrl_rect.top - win_rect.top,
                    ctrl_rect.right - win_rect.left,
                    ctrl_rect.bottom - win_rect.top,
                ]

                if rects_match(rel, expected, CTRL_RECT_TOLERANCE):
                    if DEBUG:
                        print(
                            f"    {name} -> klick X={x} Y={y} "
                            f"| fokus OK ({get_class_name(focus_hwnd)})"
                        )
                    return x, y, focus_hwnd

                last_info = f"fokus i fel kontroll rel={rel} != {expected}"
            else:
                last_info = "kunde inte läsa fokuskontrollens rektangel"
        else:
            last_info = "ingen urskiljbar kontroll fick fokus"

        if DEBUG:
            print(
                f"    {name} -> klick X={x} Y={y} "
                f"| MISSLYCKAD verifiering ({last_info}) - försök {attempt + 1}"
            )

        time.sleep(0.15)

    raise RuntimeError(
        f"Kunde inte verifiera fokus på fältet '{name}' "
        f"({last_info}). Avbryter för att inte skriva i fel fält."
    )


def read_field_via_clipboard():
    """
    Läser det aktiva fältets innehåll via urklipp.

    Markerar med Home -> Shift+End (INTE Ctrl+A) och kopierar med Ctrl+C.
    Returnerar fältets text (tom sträng om tomt).
    """

    marker = "__VISMA_EMPTY__"

    pyperclip.copy(marker)

    pyautogui.press("home", _pause=False)
    time.sleep(0.05)

    pyautogui.keyDown("shift", _pause=False)
    pyautogui.press("end", _pause=False)
    pyautogui.keyUp("shift", _pause=False)
    time.sleep(0.05)

    pyautogui.hotkey("ctrl", "c", _pause=False)
    time.sleep(0.15)

    value = str(pyperclip.paste())

    if value == marker:
        return ""

    return value


def read_field_value(focus_hwnd, main_hwnd):
    """
    Läser det aktuella fältets värde.

    Primärt via WM_GETTEXT på den fokuserade kontrollen (när den är
    identifierad). Annars via urklipp.
    """

    if focus_hwnd and focus_hwnd != main_hwnd:

        text = wm_get_text(focus_hwnd)

        # Använd WM_GETTEXT bara om det gav ett faktiskt värde. Tomt kan
        # betyda att kontrollen inte svarar på WM_GETTEXT - då är urklipp
        # säkrare än att felaktigt tolka fältet som tomt.
        if text is not None and text.strip() != "":
            return text

    return read_field_via_clipboard()


def value_matches(actual, expected):
    """
    Jämför fältvärde mot förväntat värde på siffernivå (tål olika
    avgränsare, mellanslag och ledande år).

    Exempel:
        '26-04-15' matchar '26-04-15' och '2026-04-15'
        '1 845,00' matchar '1845,00'
    """

    actual_digits = re.sub(r"\D", "", str(actual))
    expected_digits = re.sub(r"\D", "", str(expected))

    if not expected_digits:
        return False

    return (
        actual_digits == expected_digits
        or actual_digits.endswith(expected_digits)
    )


def write_and_verify(name, field_cfg, win_rect, value, type_variants=None):
    """
    Klickar på fältet, skriver värdet och LÄSER TILLBAKA fältet för att
    bekräfta att värdet verkligen står där.

    Prövar i tur och ordning:
      1. varje sträng i 'type_variants' via tangentbord. För maskade
         datumfält skickas siffrorna utan bindestreck ('260615') först,
         eftersom fältets egen mask lägger in bindestrecken.
      2. WM_SETTEXT direkt på Edit-kontrollen som sista utväg.

    Verifieras alltid mot 'value' på siffernivå. Höjer RuntimeError om
    värdet inte kan verifieras - inget skrivs vidare och ingen Num+ skickas.

    Returnerar (x, y) för den klickade punkten.
    """

    main_hwnd, _ = find_visma_window()

    if not type_variants:
        type_variants = [value]

    x = y = None
    actual = ""

    # --- Strategi 1: tangentbord, ett försök per variant ---
    for variant in type_variants:

        x, y, focus_hwnd = focus_field(name, field_cfg, win_rect)

        type_into_field(variant)
        time.sleep(0.15)

        actual = read_field_value(focus_hwnd, main_hwnd)

        if value_matches(actual, value):
            if DEBUG:
                print(
                    f"    {name} -> {value} | VERIFIERAT via tangentbord "
                    f"('{variant}' -> fältet '{actual.strip()}')"
                )
            return x, y

        if DEBUG:
            print(
                f"    {name}: tangentbord '{variant}' gav fältet "
                f"'{actual.strip()}' (förväntat '{value}') - provar nästa"
            )

        time.sleep(0.15)

    # --- Strategi 2: WM_SETTEXT direkt på kontrollen ---
    x, y, focus_hwnd = focus_field(name, field_cfg, win_rect)

    if focus_hwnd and focus_hwnd != main_hwnd:

        wm_set_text(focus_hwnd, value)
        time.sleep(0.15)

        actual = read_field_value(focus_hwnd, main_hwnd)

        if value_matches(actual, value):
            if DEBUG:
                print(
                    f"    {name} -> {value} | VERIFIERAT via WM_SETTEXT "
                    f"(fältet '{actual.strip()}')"
                )
            return x, y

    raise RuntimeError(
        f"Kunde inte verifiera att '{value}' står i fältet '{name}'. "
        f"Fältet visar '{actual.strip()}'. Avbryter - ingen Num+ skickas."
    )


def type_into_field(text):
    """
    Markerar befintligt värde och skriver nytt.

    Ctrl+A UNDVIKS (öppnar Artikelregistret i Visma Compact).
    Markering sker med End -> Shift+Home.
    """

    verify_visma_is_active()

    pyautogui.press("end", _pause=False)
    time.sleep(0.05)

    pyautogui.keyDown("shift", _pause=False)
    pyautogui.press("home", _pause=False)
    pyautogui.keyUp("shift", _pause=False)
    time.sleep(0.05)

    pyautogui.write(str(text), interval=0.01)
    time.sleep(0.10)


def press_tab(times=1):
    """Trycker Tab angivet antal gånger."""

    verify_visma_is_active()

    for i in range(times):
        pyautogui.press("tab", _pause=False)
        if DEBUG:
            print(f"    Tab {i + 1}/{times}")
        time.sleep(0.10)


def press_num_plus():
    """Skickar Num+ (Visma sparar/registrerar)."""

    verify_visma_is_active()

    user32.keybd_event(VK_ADD, 0, 0, 0)
    time.sleep(0.10)
    user32.keybd_event(VK_ADD, 0, KEYEVENTF_KEYUP, 0)


# =============================================================================
# REGISTRERA EN FAKTURA
# =============================================================================

def register_invoice(row, config, position_in_run, count, expected_number):
    """
    Registrerar en kundfaktura.

    Returnerar en dict med de klickade positionerna (för loggen).

    Vid DRY_RUN görs varken 5 Tab eller Num+.
    """

    kund_id = normalize_customer_id(row["KundID"])
    fakturadatum = format_date(row["Fakturadatum"])
    belopp = format_amount(row["Belopp"])

    if not kund_id:
        raise ValueError("KundID saknas.")
    if not fakturadatum:
        raise ValueError("Fakturadatum saknas.")

    fields = config["fields"]

    activate_visma_window()
    time.sleep(0.25)

    hwnd, _ = find_visma_window()
    win_rect = get_window_rect(hwnd)

    if not win_rect:
        raise RuntimeError("Kunde inte läsa Visma-fönstrets position.")

    positions = {}

    if DEBUG:
        print(f"[{position_in_run}/{count}] Fakturanr {expected_number}")

    # =========================================================================
    # 1. KUND ID
    # =========================================================================

    x, y, _ = focus_field("Kund ID", fields["Kund ID"], win_rect)
    positions["Kund ID"] = (x, y)

    type_into_field(kund_id)

    if DEBUG:
        print(f"    Kund ID -> {kund_id}")

    # Bekräfta så kunden laddas.
    press_tab(1)
    time.sleep(CUSTOMER_LOAD_DELAY)

    # =========================================================================
    # 2. BOKF.DAG  (direkt klick - ingen Tab-navigation)
    #
    # HÖGSTA PRIORITET: Fakturadatum MÅSTE konverteras till YY-MM-DD, skrivas
    # och läsas tillbaka och verifieras. Fortsätt INTE till Fak.belopp om
    # datumet inte finns i Bokf.dag.
    # =========================================================================

    print(f"  Excel datum: {excel_date_display(row['Fakturadatum'])}")
    print(f"  Visma datum: {fakturadatum}")

    # Skriv siffrorna utan bindestreck först (maskat datumfält), annars
    # datumet med bindestreck.
    date_digits = re.sub(r"\D", "", fakturadatum)

    positions["Bokf.dag"] = write_and_verify(
        "Bokf.dag",
        fields["Bokf.dag"],
        win_rect,
        fakturadatum,
        type_variants=[date_digits, fakturadatum],
    )

    # Exakt 1 x Tab efter Bokf.dag så Visma skapar Förf.dag automatiskt.
    press_tab(1)

    if DEBUG:
        print("    Tab x1 efter Bokf.dag (Visma skapar Förf.dag)")

    time.sleep(0.15)

    # =========================================================================
    # 3. FAK.BELOPP  (direkt klick - får ALDRIG hamna i Info)
    #
    # Även beloppet läses tillbaka och verifieras innan Num+.
    # =========================================================================

    positions["Fak.belopp"] = write_and_verify(
        "Fak.belopp",
        fields["Fak.belopp"],
        win_rect,
        belopp,
    )

    # =========================================================================
    # 3b. SLUTKONTROLL AV BOKF.DAG PRECIS FÖRE SPARANDE
    #
    # Läser om Bokf.dag-kontrollen (utan att flytta fokus) för att fånga att
    # Visma inte återställt datumet efter att vi lämnade fältet. Stämmer det
    # inte -> ingen Num+.
    # =========================================================================

    bokf_x, bokf_y = compute_point(fields["Bokf.dag"], win_rect)
    bokf_hwnd = get_control_at_point(bokf_x, bokf_y)
    bokf_now = wm_get_text(bokf_hwnd)

    if bokf_now is None or not value_matches(bokf_now, fakturadatum):
        raise RuntimeError(
            f"Bokf.dag innehåller '"
            f"{'' if bokf_now is None else bokf_now.strip()}"
            f"' precis före sparande, förväntat '{fakturadatum}'. "
            f"Avbryter - ingen Num+ skickas."
        )

    if DEBUG:
        print(f"    Slutkontroll Bokf.dag OK: '{bokf_now.strip()}'")

    # =========================================================================
    # 4. DRY_RUN -> stoppa här, spara inget
    # =========================================================================

    if DRY_RUN:

        print()
        print("DRY RUN")
        print()
        print("Kontrollera i Visma:")
        print(f"KundID: {kund_id}")
        print(f"Bokf.dag: {fakturadatum}")
        print(f"Fak.belopp: {belopp}")
        print()
        print("Ingen faktura har sparats.")

        try:
            input("Tryck Enter för att fortsätta...")
        except EOFError:
            pass

        return positions

    # =========================================================================
    # 5. EXAKT 5 TAB + NUM+
    # =========================================================================

    press_tab(TABS_AFTER_BELOPP)

    if DEBUG:
        print("    Skickar Num+")

    press_num_plus()
    time.sleep(SAVE_DELAY)

    if DEBUG:
        print("    SPARAD")

    return positions


# =============================================================================
# EXCEL
# =============================================================================

def resolve_excel_path(excel_path):
    """Hittar Excel-fil även om .xlsx inte skrivits."""

    excel_path = excel_path.strip().strip('"').strip("'").strip()

    file_path = Path(excel_path)

    if file_path.exists():
        return file_path

    if file_path.suffix.lower() not in (".xlsx", ".xls"):
        for ext in (".xlsx", ".xls"):
            candidate = file_path.with_name(file_path.name + ext)
            if candidate.exists():
                return candidate

    return None


def apply_column_aliases(df):
    """Tillåter alternativa Excel-kolumnnamn."""

    aliases = {
        "KundID": ["KundID", "Personnummer", "Kundnummer"],
        "Fakturanr": [
            "Fakturanr", "Fakturanummer", "Faktura nr", "Faktura Nr", "FakturaNr"
        ],
        "Fakturadatum": ["Fakturadatum", "Faktura datum"],
        "Belopp": ["Belopp", "Fakturabelopp"],
    }

    for target, candidates in aliases.items():

        if target in df.columns:
            continue

        for candidate in candidates:
            if candidate in df.columns:
                df = df.rename(columns={candidate: target})
                break

    return df


def check_required_columns(df):
    """Kontrollerar att obligatoriska kolumner finns."""

    required = ["KundID", "Fakturanr", "Fakturadatum", "Belopp"]

    missing = [c for c in required if c not in df.columns]

    if missing:
        print("\nExcel-filen saknar obligatoriska kolumner:")
        for column in missing:
            print(f"  - {column}")
        return False

    return True


def validate_invoice_series(df, start_number, count):
    """
    Kontrollerar att de 'count' första Excel-raderna exakt motsvarar serien
    start_number, start_number+1, ... Returnerar (ok, felmeddelande).
    """

    available = len(df)

    if count > available:
        message = (
            "IMPORT STOPPAD\n\n"
            f"För få rader i Excel. Du vill registrera {count} fakturor men\n"
            f"Excel innehåller endast {available} rader.\n\n"
            "Ingen registrering har startats i Visma."
        )
        return False, message

    selected = df.head(count)

    for position, (index, row) in enumerate(selected.iterrows()):

        expected = start_number + position
        excel_number = parse_invoice_number(row["Fakturanr"])
        excel_raw = normalize_invoice_number(row["Fakturanr"])
        excel_display = excel_raw if excel_raw else "<tomt/ogiltigt>"

        if excel_number is None or excel_number != expected:
            message = (
                "IMPORT STOPPAD\n\n"
                f"Excel-rad: {index + 2}\n"
                f"Förväntat Fakturanr: {expected}\n"
                f"Excel Fakturanr: {excel_display}\n\n"
                "Ingen registrering har startats i Visma."
            )
            return False, message

    return True, ""


# =============================================================================
# LOGG
# =============================================================================

def save_log(file_path, log_rows):
    """Sparar importlogg bredvid Excel-filen."""

    if not log_rows:
        return None

    columns = [
        "Excel-rad", "KundID", "Fakturanr", "Förväntat Fakturanr",
        "Fakturadatum", "Belopp",
        "KundID-position", "Bokf.dag-position", "Fak.belopp-position",
        "Status", "Fel",
    ]

    log_df = pd.DataFrame(log_rows, columns=columns)

    log_file = (
        file_path.parent
        / f"visma_import_logg_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    log_df.to_excel(log_file, index=False)

    return log_file


# =============================================================================
# INMATNING
# =============================================================================

def ask_int(prompt, min_value, max_value):
    """Frågar efter ett heltal i [min_value, max_value]. None = avbryt."""

    while True:
        try:
            answer = input(prompt).strip()
        except EOFError:
            return None

        if answer == "":
            return None

        if not re.fullmatch(r"\d+", answer):
            print(
                f"  Ogiltigt värde. Ange ett heltal mellan "
                f"{min_value} och {max_value}."
            )
            continue

        value = int(answer)

        if value < min_value or value > max_value:
            print(f"  Värdet måste vara mellan {min_value} och {max_value}.")
            continue

        return value


def pos_to_text(positions, name):
    """Formaterar en klickposition för loggen."""

    if name in positions:
        x, y = positions[name]
        return f"X={x} Y={y}"

    return ""


# =============================================================================
# IMPORT
# =============================================================================

def run_import():
    """Kör hela importflödet."""

    config = load_config()

    if config is None:
        print()
        print("Ingen kalibrering hittades (eller ofullständig).")
        print("Välj 'Kalibrera Visma-fält' i menyn först.")
        return

    if DRY_RUN:
        print("\n*** DRY_RUN AKTIVT: ingen faktura sparas (ingen Num+). ***")
    if DEBUG:
        print("*** DEBUG AKTIVT: varje GUI-steg visas. ***")

    # -------------------------------------------------------------------------
    # Excel
    # -------------------------------------------------------------------------

    try:
        excel_path = input("\nAnge sökväg till Excel-filen: ")
    except EOFError:
        return

    file_path = resolve_excel_path(excel_path)

    if file_path is None:
        print(f"\nFilen hittades inte: {excel_path}")
        return

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"\nKunde inte läsa Excel-filen: {e}")
        return

    df = apply_column_aliases(df)

    if not check_required_columns(df):
        return

    total_rows = len(df)

    if total_rows == 0:
        print("\nExcel-filen innehåller inga rader.")
        return

    # -------------------------------------------------------------------------
    # Startnummer + antal
    # -------------------------------------------------------------------------

    start_number = ask_int("\nAnge första Visma Fakturanr: ", 1, 99999999)
    if start_number is None:
        print("Avbrutet.")
        return

    max_allowed = min(MAX_INVOICES_PER_RUN, total_rows)

    count = ask_int(
        f"Hur många kundfakturor ska registreras? (1-{max_allowed}): ",
        1,
        max_allowed,
    )
    if count is None:
        print("Avbrutet.")
        return

    # -------------------------------------------------------------------------
    # Förkontroll av serien (innan Visma påverkas)
    # -------------------------------------------------------------------------

    ok, error_message = validate_invoice_series(df, start_number, count)

    if not ok:
        print("\n" + "=" * 80)
        print(error_message)
        print("=" * 80 + "\n")
        return

    selected = df.head(count)
    first_excel = normalize_invoice_number(selected.iloc[0]["Fakturanr"])
    last_excel = normalize_invoice_number(selected.iloc[-1]["Fakturanr"])
    last_expected = start_number + count - 1

    print("\n" + "=" * 80)
    print("FÖRKONTROLL GODKÄND")
    print("=" * 80)
    print(f"\nVisma startnummer:      {start_number}")
    print(f"Antal fakturor:         {count}")
    print(f"Första Excel Fakturanr: {first_excel}")
    print(f"Sista Excel Fakturanr:  {last_excel}")
    print(f"\nFörväntad nummerserie:\n{start_number} - {last_expected}\n")

    hwnd, visma_title = find_visma_window()

    if not hwnd:
        print("FEL: Visma Compact hittades inte.")
        print("Öppna Visma -> Kundreskontra -> Kundbokning.")
        return

    print(f"Visma hittat: {visma_title}")
    print("\nNÖDSTOPP: flytta muspekaren till skärmens övre vänstra hörn.\n")

    try:
        confirm = input(
            "Skriv JA och tryck Enter för att starta registreringen: "
        ).strip()
    except EOFError:
        return

    if confirm.lower() != "ja":
        print("Avbrutet - inget registrerades.")
        return

    print()
    for seconds in range(3, 0, -1):
        print(f"Startar om {seconds}...")
        time.sleep(1)

    # -------------------------------------------------------------------------
    # Registrering
    # -------------------------------------------------------------------------

    log_rows = []
    stopped = False
    expected_invoice_number = start_number

    for position, (index, row) in enumerate(selected.iterrows()):

        excel_rad = index + 2
        kund_id = normalize_customer_id(row["KundID"])
        excel_fakturanr = normalize_invoice_number(row["Fakturanr"])
        fakturadatum = format_date(row["Fakturadatum"])
        belopp = format_amount(row["Belopp"])
        positions = {}

        print("\n" + "-" * 80)
        print(f"Registrerar {position + 1}/{count}")
        print(f"Förväntat Fakturanr: {expected_invoice_number}")
        print(f"Excel Fakturanr:     {excel_fakturanr}")

        def log(status, fel):
            log_rows.append({
                "Excel-rad": excel_rad,
                "KundID": kund_id,
                "Fakturanr": excel_fakturanr,
                "Förväntat Fakturanr": expected_invoice_number,
                "Fakturadatum": fakturadatum,
                "Belopp": belopp,
                "KundID-position": pos_to_text(positions, "Kund ID"),
                "Bokf.dag-position": pos_to_text(positions, "Bokf.dag"),
                "Fak.belopp-position": pos_to_text(positions, "Fak.belopp"),
                "Status": status,
                "Fel": fel,
            })

        # Kontroll under körning: raden måste matcha förväntat nummer.
        excel_number = parse_invoice_number(row["Fakturanr"])

        if excel_number is None or excel_number != expected_invoice_number:
            error_text = (
                f"Oväntad avvikelse: förväntat {expected_invoice_number}, "
                f"Excel {excel_fakturanr}."
            )
            print("\n" + "=" * 80)
            print("IMPORTEN STOPPAD")
            print("=" * 80)
            print(f"\n{error_text}\nIngen Num+ skickades.")
            log("STOPPAD", error_text)
            stopped = True
            break

        try:
            positions = register_invoice(
                row, config, position + 1, count, expected_invoice_number
            )

            if DRY_RUN:
                log("DRY_RUN", "")
                print(f"  DRY_RUN klar (inget sparat)")
            else:
                log("OK", "")
                print(f"  SPARAD")
                expected_invoice_number += 1

            time.sleep(0.40)

        except pyautogui.FailSafeException:
            error_text = "PyAutoGUI failsafe aktiverades (nödstopp)."
            print(f"\n{error_text}")
            log("STOPPAD", error_text)
            stopped = True
            break

        except KeyboardInterrupt:
            error_text = "Importen avbröts manuellt."
            print(f"\n{error_text}")
            log("STOPPAD", error_text)
            stopped = True
            break

        except Exception as e:
            error_text = str(e)
            print("\n" + "=" * 80)
            print("IMPORTEN STOPPAD")
            print("=" * 80)
            print(f"\nExcel-rad: {excel_rad}")
            print(f"KundID: {kund_id}")
            print(f"Förväntat Fakturanr: {expected_invoice_number}")
            print(f"\nOrsak: {error_text}")
            print("\nIngen Num+ skickades efter felet.")
            log("STOPPAD", error_text)
            stopped = True
            break

    # -------------------------------------------------------------------------
    # Logg + summering
    # -------------------------------------------------------------------------

    log_file = None
    try:
        log_file = save_log(file_path, log_rows)
    except Exception as e:
        print(f"\nKunde inte spara loggfilen: {e}")

    ok_count = sum(1 for r in log_rows if r["Status"] in ("OK", "DRY_RUN"))
    err_count = sum(1 for r in log_rows if r["Status"] == "STOPPAD")

    print("\n" + "=" * 80)
    print("KÖRNINGEN STOPPAD" if stopped else "KÖRNINGEN KLAR")
    print("=" * 80)
    print(f"\nBehandlade fakturor: {ok_count}")
    print(f"Stoppade/fel:        {err_count}")

    if log_file:
        print(f"\nLogg sparad:\n{log_file}")

    print()


# =============================================================================
# MENY
# =============================================================================

def main():

    print()
    print("=" * 80)
    print("VISMA COMPACT - AUTOMATISK KUNDFAKTURAREGISTRERING")
    print("=" * 80)

    while True:

        has_config = load_config() is not None

        print()
        print("MENY")
        print("  1. Starta import" + ("" if has_config else "  (kräver kalibrering)"))
        print("  2. Kalibrera Visma-fält")
        print("  3. Avsluta")

        try:
            choice = input("\nVälj (1-3): ").strip()
        except EOFError:
            break

        if choice == "1":
            run_import()
        elif choice == "2":
            calibrate()
        elif choice == "3":
            print("Avslutar.")
            break
        else:
            print("Ogiltigt val.")


if __name__ == "__main__":
    main()
