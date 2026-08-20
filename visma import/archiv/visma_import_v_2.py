import pandas as pd
import pyautogui
import pyperclip
import time
import ctypes
import re

from ctypes import wintypes
from pathlib import Path
from datetime import datetime


# =============================================================================
# OCR - LÄSER VISMA "Nr."-FÄLTET SOM TEXT
# =============================================================================
#
# Nr.-fältet är skrivskyddat och går inte att kopiera via urklipp.
# I stället fotas fältet och läses med Tesseract-OCR (gratis, lokalt).
#
# Installera EN gång:
#   1. Ladda ner Tesseract:
#        https://github.com/UB-Mannheim/tesseract/wiki
#   2. Installera (standardval).
#   3. pip install pytesseract
#
# =============================================================================

# Sökväg till Tesseract-programmet. Ändra om du installerat på annan plats.
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Storlek i pixlar på området runt Nr.-siffran som fotas för OCR.
# Området centreras på den kalibrerade punkten (INVOICE_NR_X_RATIO / _Y).
INVOICE_NR_BOX_WIDTH = 110
INVOICE_NR_BOX_HEIGHT = 28

try:

    import pytesseract

    if Path(TESSERACT_PATH).exists():
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    OCR_AVAILABLE = True

except Exception:

    OCR_AVAILABLE = False


# =============================================================================
# VISMA COMPACT - AUTOMATISK FAKTURAREGISTRERING
# =============================================================================
#
# FLÖDE:
#
#   KundID
#      ↓
#   Visma laddar kunden
#      ↓
#   Kontrollera "Nr." uppe till höger mot Fakturanr i Excel
#      ↓
#   Klicka direkt på Bokf.dag
#      ↓
#   Skriv Fakturadatum
#      ↓
#   Klicka direkt på Fak.belopp
#      ↓
#   Skriv Belopp
#      ↓
#   TAB x 5
#      ↓
#   NUM+
#
# =============================================================================


# =============================================================================
# GRUNDINSTÄLLNINGAR
# =============================================================================

pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = True

# Vänta efter att KundID skrivits så Visma hinner ladda kunden.
CUSTOMER_LOAD_DELAY = 0.8

# Vänta innan fakturanumret kontrolleras.
INVOICE_NUMBER_WAIT = 0.5

# Fakturanummer-kontrollen läser Visma "Nr."-fältet via pixelposition
# och urklipp. Det fältet är ofta skrivskyddat och går inte alltid att
# kopiera, vilket gör att avläsningen misslyckas även när Visma laddat
# RÄTT faktura.
#
#   False = kan numret inte LÄSAS -> varning i loggen, importen fortsätter.
#           (En verklig FELMATCHNING stoppar ändå alltid importen.)
#   True  = kan numret inte läsas/verifieras -> importen stoppas (som förr).
#
# Behåll False om importen tidigare stannade på "Kunde inte verifiera Fakturanr".
STOP_IF_INVOICE_NUMBER_UNREADABLE = False

# Vänta efter Num+ så Visma hinner spara.
SAVE_DELAY = 1.5

# EXAKT 5 TAB efter Fak.belopp
TABS_AFTER_BELOPP = 5


# =============================================================================
# TESTLÄGE
# =============================================================================
#
# 1 = endast första fakturan
# 2 = två fakturor
# 0 = alla fakturor
#
# Behåll 1 tills allt fungerar korrekt.
# =============================================================================

TEST_LIMIT = 0


# =============================================================================
# VISMA-POSITIONER
# =============================================================================
#
# Dessa positioner är anpassade efter Visma-layouten på dina skärmbilder.
#
# X anges som andel av Visma-fönstrets bredd.
# Y anges i pixlar från Visma-fönstrets överkant.
#
# Exempel:
#   0.13 = 13 % från vänster
#
# =============================================================================

# KundID-fält
KUNDID_X_RATIO = 0.036
KUNDID_Y = 99

# Fakturanummer "Nr." uppe till höger
INVOICE_NR_X_RATIO = 0.917
INVOICE_NR_Y = 62

# Bokf.dag
BOKFDAG_X_RATIO = 0.302
BOKFDAG_Y = 123

# Fak.belopp
FAK_BELOPP_X_RATIO = 0.789
FAK_BELOPP_Y = 124


# =============================================================================
# WINDOWS API
# =============================================================================

user32 = ctypes.windll.user32

VK_ADD = 0x6B
KEYEVENTF_KEYUP = 0x0002

SW_RESTORE = 9


# =============================================================================
# FORMATERING
# =============================================================================

def format_date(value):
    """
    Konverterar datum till Visma-format YY-MM-DD.

    Exempel:
        2026-08-17 -> 26-08-17
    """

    if pd.isna(value):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%y-%m-%d")

    text = str(value).strip()

    try:
        dt = pd.to_datetime(text)
        return dt.strftime("%y-%m-%d")
    except Exception:
        return text


def format_amount(value):
    """
    Konverterar belopp till svenskt decimalformat.

    Exempel:
        1949       -> 1949,00
        1949.50    -> 1949,50
    """

    if pd.isna(value):
        return "0,00"

    try:
        amount = float(value)

        return f"{amount:.2f}".replace(
            ".",
            ","
        )

    except Exception:

        return str(value).strip().replace(
            ".",
            ","
        )


def normalize_customer_id(value):
    """
    Normaliserar KundID/personnummer.
    """

    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def normalize_invoice_number(value):
    """
    Normaliserar fakturanummer.

    Exempel:
        48991
        48991.0
        "48991"

    blir:
        "48991"
    """

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


# =============================================================================
# WINDOWS-FÖNSTER
# =============================================================================

def get_window_title(hwnd):
    """
    Hämtar titeln på ett Windows-fönster.
    """

    if not hwnd:
        return ""

    length = user32.GetWindowTextLengthW(
        hwnd
    )

    if length <= 0:
        return ""

    buffer = ctypes.create_unicode_buffer(
        length + 1
    )

    user32.GetWindowTextW(
        hwnd,
        buffer,
        length + 1
    )

    return buffer.value.strip()


def get_foreground_window_title():
    """
    Returnerar titeln på aktuellt aktivt fönster.
    """

    hwnd = user32.GetForegroundWindow()

    return get_window_title(
        hwnd
    )


def get_window_rect(hwnd):
    """
    Hämtar Visma-fönstrets position på skärmen.
    """

    rect = wintypes.RECT()

    result = user32.GetWindowRect(
        hwnd,
        ctypes.byref(rect)
    )

    if not result:
        return None

    return rect


# =============================================================================
# HITTA VISMA
# =============================================================================

def find_visma_window():
    """
    Letar efter öppet Visma-fönster.

    Prioriterar fönster med:
        Kundbokning
        Kundreskontra
        Compact
    """

    windows = []

    CALLBACK = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HWND,
        wintypes.LPARAM
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

        windows.append(
            (
                score,
                hwnd,
                title
            )
        )

        return True

    user32.EnumWindows(
        enum_proc,
        0
    )

    if not windows:
        return None, ""

    windows.sort(
        key=lambda item: item[0],
        reverse=True
    )

    _, hwnd, title = windows[0]

    return hwnd, title


# =============================================================================
# AKTIVERA VISMA
# =============================================================================

def activate_visma_window():
    """
    Aktiverar Visma-fönstret.
    """

    hwnd, title = find_visma_window()

    if not hwnd:

        raise RuntimeError(
            "Kunde inte hitta Visma Compact. "
            "Öppna Kundreskontra -> Kundbokning."
        )

    user32.ShowWindow(
        hwnd,
        SW_RESTORE
    )

    time.sleep(0.1)

    user32.BringWindowToTop(
        hwnd
    )

    user32.SetForegroundWindow(
        hwnd
    )

    time.sleep(0.35)

    active_title = get_foreground_window_title()

    if "visma" not in active_title.lower():

        # Försök en gång till

        user32.ShowWindow(
            hwnd,
            SW_RESTORE
        )

        user32.BringWindowToTop(
            hwnd
        )

        user32.SetForegroundWindow(
            hwnd
        )

        time.sleep(0.4)

        active_title = get_foreground_window_title()

    if "visma" not in active_title.lower():

        raise RuntimeError(
            "Kunde inte aktivera Visma. "
            f"Aktivt fönster är '{active_title}'."
        )

    return hwnd


def verify_visma_is_active():
    """
    Stoppar importen om Visma tappat fokus.
    """

    title = get_foreground_window_title()

    if "visma" not in title.lower():

        raise RuntimeError(
            "Visma har tappat fokus. "
            f"Aktivt fönster är '{title}'."
        )


# =============================================================================
# POSITIONER I VISMA
# =============================================================================

def get_visma_point(
    x_ratio,
    y_offset
):
    """
    Omvandlar en relativ Visma-position till
    absoluta skärmkoordinater.
    """

    hwnd, _ = find_visma_window()

    if not hwnd:

        raise RuntimeError(
            "Visma-fönstret hittades inte."
        )

    rect = get_window_rect(
        hwnd
    )

    if not rect:

        raise RuntimeError(
            "Kunde inte läsa Visma-fönstrets position."
        )

    width = (
        rect.right
        -
        rect.left
    )

    x = (
        rect.left
        +
        int(
            width * x_ratio
        )
    )

    y = (
        rect.top
        +
        y_offset
    )

    return x, y


def click_visma_position(
    x_ratio,
    y_offset
):
    """
    Klickar på en bestämd position i Visma.
    """

    verify_visma_is_active()

    x, y = get_visma_point(
        x_ratio,
        y_offset
    )

    pyautogui.click(
        x,
        y
    )

    time.sleep(0.20)


# =============================================================================
# KUNDID
# =============================================================================

def focus_kund_id():
    """
    Klickar direkt i KundID-fältet.
    """

    click_visma_position(
        KUNDID_X_RATIO,
        KUNDID_Y
    )


# =============================================================================
# BOKF.DAG
# =============================================================================

def focus_bokf_dag():
    """
    Klickar direkt i Bokf.dag-fältet.
    """

    click_visma_position(
        BOKFDAG_X_RATIO,
        BOKFDAG_Y
    )


# =============================================================================
# FAK.BELOPP
# =============================================================================

def focus_fak_belopp():
    """
    Klickar direkt i Fak.belopp-fältet.

    På detta sätt används ingen Tab-navigation för att
    försöka hitta beloppsfältet.
    """

    click_visma_position(
        FAK_BELOPP_X_RATIO,
        FAK_BELOPP_Y
    )


# =============================================================================
# SKRIVA I FÄLT
# =============================================================================

def write_text(text):
    """
    Skriver text i aktivt fält.
    """

    verify_visma_is_active()

    pyautogui.write(
        str(text),
        interval=0.01
    )


def replace_current_text(text):
    """
    Ersätter hela innehållet i aktivt fält.
    """

    verify_visma_is_active()

    # OBS: Ctrl+A används INTE här.
    # I Visma Compact är Ctrl+A en genväg som ÖPPNAR Artikelregistret,
    # inte "markera allt". Vi markerar i stället hela fältet med
    # End -> Shift+Home och skriver över markeringen.

    pyautogui.press(
        "end",
        _pause=False
    )

    time.sleep(0.05)

    pyautogui.keyDown(
        "shift",
        _pause=False
    )

    pyautogui.press(
        "home",
        _pause=False
    )

    pyautogui.keyUp(
        "shift",
        _pause=False
    )

    time.sleep(0.05)

    pyautogui.write(
        str(text),
        interval=0.01
    )

    time.sleep(0.10)


def press_tab(times=1):
    """
    Trycker Tab angivet antal gånger.
    """

    verify_visma_is_active()

    for _ in range(times):

        pyautogui.press(
            "tab",
            _pause=False
        )

        time.sleep(0.10)


# =============================================================================
# LÄS TEXT FRÅN EN BESTÄMD POSITION
# =============================================================================

def copy_text_from_point(
    x,
    y
):
    """
    Klickar på en punkt och försöker kopiera
    fältets innehåll.

    Accepterar inte gammal clipboard-text.
    """

    verify_visma_is_active()

    pyautogui.click(
        x,
        y
    )

    time.sleep(0.15)

    marker = "__VISMA_NO_VALUE__"

    pyperclip.copy(
        marker
    )

    # Försök 1:
    # Markera med End -> Shift+Home (INTE Ctrl+A, som öppnar
    # Artikelregistret i Visma) och kopiera med Ctrl+C.

    pyautogui.press(
        "end",
        _pause=False
    )

    time.sleep(0.05)

    pyautogui.keyDown(
        "shift",
        _pause=False
    )

    pyautogui.press(
        "home",
        _pause=False
    )

    pyautogui.keyUp(
        "shift",
        _pause=False
    )

    time.sleep(0.08)

    pyautogui.hotkey(
        "ctrl",
        "c",
        _pause=False
    )

    time.sleep(0.25)

    value = str(
        pyperclip.paste()
    ).strip()

    if (
        value
        and
        value != marker
    ):
        return value

    # -------------------------------------------------------------------------
    # Försök 2
    #
    # HOME -> SHIFT+END -> CTRL+C
    #
    # Vissa äldre Windows-fält reagerar bättre på detta.
    # -------------------------------------------------------------------------

    pyperclip.copy(
        marker
    )

    pyautogui.press(
        "home",
        _pause=False
    )

    time.sleep(0.05)

    pyautogui.keyDown(
        "shift",
        _pause=False
    )

    pyautogui.press(
        "end",
        _pause=False
    )

    pyautogui.keyUp(
        "shift",
        _pause=False
    )

    time.sleep(0.08)

    pyautogui.hotkey(
        "ctrl",
        "c",
        _pause=False
    )

    time.sleep(0.25)

    value = str(
        pyperclip.paste()
    ).strip()

    if value == marker:
        return ""

    return value


# =============================================================================
# OCR-AVLÄSNING AV Nr.-FÄLTET
# =============================================================================

def ocr_read_invoice_number(rect):
    """
    Fotar Nr.-fältet och läser siffrorna med Tesseract-OCR.

    Returnerar en ren siffersträng, eller "" om OCR inte är
    tillgängligt eller inget kunde läsas.
    """

    if not OCR_AVAILABLE:
        return ""

    width = rect.right - rect.left

    if width <= 0:
        return ""

    # Mittpunkt på den kalibrerade Nr.-positionen.
    center_x = rect.left + int(width * INVOICE_NR_X_RATIO)
    center_y = rect.top + INVOICE_NR_Y

    left = center_x - INVOICE_NR_BOX_WIDTH // 2
    top = center_y - INVOICE_NR_BOX_HEIGHT // 2

    try:

        image = pyautogui.screenshot(
            region=(
                left,
                top,
                INVOICE_NR_BOX_WIDTH,
                INVOICE_NR_BOX_HEIGHT,
            )
        )

    except Exception:
        return ""

    # Gråskala + förstoring ger tydligare siffror för OCR.
    try:

        image = image.convert("L").resize(
            (
                INVOICE_NR_BOX_WIDTH * 4,
                INVOICE_NR_BOX_HEIGHT * 4,
            )
        )

    except Exception:
        pass

    # psm 7 = en enda textrad. Endast siffror tillåts.
    config = (
        "--psm 7 "
        "-c tessedit_char_whitelist=0123456789"
    )

    try:

        text = pytesseract.image_to_string(
            image,
            config=config,
        )

    except Exception as e:

        print(
            f"  OCR kunde inte köras: {e}"
        )

        return ""

    digits = re.sub(
        r"\D",
        "",
        text,
    )

    return digits


# =============================================================================
# LÄS VISMA FAKTURANUMMER
# =============================================================================

def read_visma_invoice_number(
    expected_invoice_number
):
    """
    Läser Fakturanr från fältet:

        Nr.: 48991

    uppe till höger.

    Vi försöker flera punkter inom Nr.-fältet.

    Säkerhetsregel:
    Endast ett rent numeriskt värde accepteras.
    """

    verify_visma_is_active()

    expected_invoice_number = normalize_invoice_number(
        expected_invoice_number
    )

    hwnd, _ = find_visma_window()

    if not hwnd:
        return ""

    rect = get_window_rect(
        hwnd
    )

    if not rect:
        return ""

    # -------------------------------------------------------------------------
    # PRIMÄR METOD: OCR (fotar Nr.-fältet och läser siffrorna).
    # -------------------------------------------------------------------------

    ocr_value = ocr_read_invoice_number(
        rect
    )

    if ocr_value:

        ocr_value = normalize_invoice_number(
            ocr_value
        )

        print(
            f"  Visma Nr-fält (OCR): '{ocr_value}'"
        )

        return ocr_value

    # -------------------------------------------------------------------------
    # RESERVMETOD: urklipp. Används om OCR saknas eller misslyckas.
    # Denna metod returnerar bara ett värde vid exakt match mot Excel.
    # -------------------------------------------------------------------------

    width = (
        rect.right
        -
        rect.left
    )

    # -------------------------------------------------------------------------
    # Prova flera punkter inuti Fakturanr-fältet.
    #
    # Fältet ligger ungefär:
    #
    #               Nr.: [ 48991 ] Ver: B 739
    #
    # -------------------------------------------------------------------------

    x_ratios = [
        0.925,
        0.930,
        0.934,
        0.938,
        0.942
    ]

    y_offsets = [
        17,
        20,
        23,
        26,
        29
    ]

    values_found = []

    for x_ratio in x_ratios:

        x = (
            rect.left
            +
            int(
                width * x_ratio
            )
        )

        for y_offset in y_offsets:

            y = (
                rect.top
                +
                y_offset
            )

            try:

                raw_value = copy_text_from_point(
                    x,
                    y
                )

            except Exception:
                continue

            raw_value = str(
                raw_value
            ).strip()

            if not raw_value:
                continue

            # -----------------------------------------------------------------
            # Acceptera ENDAST om hela clipboardvärdet är numeriskt.
            # -----------------------------------------------------------------

            if re.fullmatch(
                r"\d+",
                raw_value
            ):

                value = normalize_invoice_number(
                    raw_value
                )

                values_found.append(
                    value
                )

                # Exakt match mot Excel
                if value == expected_invoice_number:

                    print(
                        f"  Visma Nr-fält läst: '{value}'"
                    )

                    return value

    if values_found:

        print(
            "  Numeriska värden som hittades: "
            f"{sorted(set(values_found))}"
        )

    return ""


# =============================================================================
# KONTROLLERA FAKTURANUMMER
# =============================================================================

def verify_invoice_number(
    excel_invoice_number
):
    """
    Fakturanummer i Excel måste vara exakt samma
    som Visma Nr. uppe till höger.
    """

    excel_invoice_number = normalize_invoice_number(
        excel_invoice_number
    )

    if not excel_invoice_number:

        raise ValueError(
            "Fakturanummer saknas i Excel."
        )

    if not re.fullmatch(
        r"\d+",
        excel_invoice_number
    ):

        raise ValueError(
            f"Ogiltigt Fakturanr i Excel: "
            f"'{excel_invoice_number}'."
        )

    time.sleep(
        INVOICE_NUMBER_WAIT
    )

    verify_visma_is_active()

    visma_invoice_number = read_visma_invoice_number(
        excel_invoice_number
    )

    # -------------------------------------------------------------------------
    # Bekräftad exakt match -> allt OK.
    # -------------------------------------------------------------------------

    if visma_invoice_number == excel_invoice_number:

        print(
            f"  OK - Fakturanummer "
            f"{visma_invoice_number} matchar Excel."
        )

        return visma_invoice_number

    # -------------------------------------------------------------------------
    # Visma visade ett ANNAT nummer -> verklig felmatchning.
    # Detta stoppar alltid importen, oavsett inställning.
    # -------------------------------------------------------------------------

    if visma_invoice_number:

        raise ValueError(
            "FAKTURANUMMER MATCHAR INTE! "
            f"Excel={excel_invoice_number} | "
            f"Visma={visma_invoice_number}. "
            "Fakturan har INTE sparats."
        )

    # -------------------------------------------------------------------------
    # Numret kunde inte LÄSAS ur Visma "Nr."-fältet.
    #
    # Detta beror oftast på att fältet är skrivskyddat och inte går att
    # kopiera via urklipp - inte på att fel faktura är laddad.
    # -------------------------------------------------------------------------

    if STOP_IF_INVOICE_NUMBER_UNREADABLE:

        raise ValueError(
            f"Kunde inte läsa Fakturanr "
            f"{excel_invoice_number} ur Visma-fältet 'Nr.'. "
            f"Fakturan har INTE sparats. "
            f"(Sätt STOP_IF_INVOICE_NUMBER_UNREADABLE = False "
            f"för att fortsätta ändå.)"
        )

    print(
        f"  VARNING: Kunde inte läsa Visma Nr.-fältet - "
        f"fortsätter utan bekräftelse (Excel={excel_invoice_number})."
    )

    return ""


# =============================================================================
# NUM+
# =============================================================================

def press_num_plus():
    """
    Skickar riktig Num+ från numeriska tangentbordet.

    Visma använder Num+ för att spara/registrera.
    """

    verify_visma_is_active()

    # Num+ DOWN
    user32.keybd_event(
        VK_ADD,
        0,
        0,
        0
    )

    time.sleep(0.10)

    # Num+ UP
    user32.keybd_event(
        VK_ADD,
        0,
        KEYEVENTF_KEYUP,
        0
    )


def save_record():
    """
    Sparar fakturan.
    """

    press_num_plus()

    time.sleep(
        SAVE_DELAY
    )


# =============================================================================
# REGISTRERA EN FAKTURA
# =============================================================================

def register_invoice(row):
    """
    Registrerar en kundfaktura.

    FLÖDE:

        1. Klicka KundID
        2. Skriv KundID
        3. Tab för att bekräfta kunden
        4. Kontrollera Visma Nr. mot Excel
        5. Klicka Bokf.dag
        6. Skriv Fakturadatum
        7. Klicka Fak.belopp
        8. Skriv Belopp
        9. Tab x 5
       10. Num+

    Ingen Tab-navigation används för att försöka hitta
    Bokf.dag eller Fak.belopp.
    """

    kund_id = normalize_customer_id(
        row["KundID"]
    )

    excel_fakturanr = normalize_invoice_number(
        row["Fakturanr"]
    )

    fakturadatum = format_date(
        row["Fakturadatum"]
    )

    belopp = format_amount(
        row["Belopp"]
    )

    # -------------------------------------------------------------------------
    # Validering
    # -------------------------------------------------------------------------

    if not kund_id:

        raise ValueError(
            "KundID saknas."
        )

    if not excel_fakturanr:

        raise ValueError(
            "Fakturanr saknas."
        )

    if not fakturadatum:

        raise ValueError(
            "Fakturadatum saknas."
        )

    # -------------------------------------------------------------------------
    # Aktivera Visma
    # -------------------------------------------------------------------------

    activate_visma_window()

    time.sleep(0.25)

    # =========================================================================
    # 1. KUNDID
    # =========================================================================

    focus_kund_id()

    replace_current_text(
        kund_id
    )

    print(
        f"  KundID: {kund_id}"
    )

    # Bekräfta KundID så Visma laddar kunden
    press_tab(1)

    time.sleep(
        CUSTOMER_LOAD_DELAY
    )

    # =========================================================================
    # 2. FAKTURANUMMER
    # =========================================================================

    visma_fakturanr = verify_invoice_number(
        excel_fakturanr
    )

    # =========================================================================
    # 3. BOKF.DAG
    # =========================================================================
    #
    # Direkt klick.
    # Fakturanummerkontrollen kan därför lämna fokus var som helst.
    # =========================================================================

    focus_bokf_dag()

    replace_current_text(
        fakturadatum
    )

    print(
        f"  Bokf.dag: {fakturadatum}"
    )

    time.sleep(0.15)

    # =========================================================================
    # 4. FAK.BELOPP
    # =========================================================================
    #
    # Direkt klick på rätt beloppsfält.
    #
    # Detta förhindrar att scriptet tabbar ner i artikelraderna.
    # =========================================================================

    focus_fak_belopp()

    replace_current_text(
        belopp
    )

    print(
        f"  Fak.belopp: {belopp}"
    )

    time.sleep(0.15)

    # =========================================================================
    # 5. EXAKT 5 TAB
    # =========================================================================

    press_tab(
        TABS_AFTER_BELOPP
    )

    # =========================================================================
    # 6. NUM+
    # =========================================================================

    save_record()

    return visma_fakturanr


# =============================================================================
# EXCEL-SÖKVÄG
# =============================================================================

def resolve_excel_path(excel_path):
    """
    Hittar Excel-fil även om .xlsx inte skrivits.
    """

    excel_path = (
        excel_path
        .strip()
        .strip('"')
        .strip("'")
        .strip()
    )

    file_path = Path(
        excel_path
    )

    if file_path.exists():
        return file_path

    if file_path.suffix.lower() not in (
        ".xlsx",
        ".xls"
    ):

        for ext in (
            ".xlsx",
            ".xls"
        ):

            candidate = file_path.with_name(
                file_path.name + ext
            )

            if candidate.exists():
                return candidate

    return None


# =============================================================================
# KOLUMNALIAS
# =============================================================================

def apply_column_aliases(df):
    """
    Tillåter alternativa Excel-kolumnnamn.
    """

    aliases = {

        "KundID": [
            "KundID",
            "Personnummer",
            "Kundnummer"
        ],

        "Fakturanr": [
            "Fakturanr",
            "Fakturanummer",
            "Faktura nr",
            "Faktura Nr",
            "FakturaNr"
        ],

        "Fakturadatum": [
            "Fakturadatum",
            "Faktura datum"
        ],

        "Belopp": [
            "Belopp",
            "Fakturabelopp"
        ]
    }

    for target, candidates in aliases.items():

        if target in df.columns:
            continue

        for candidate in candidates:

            if candidate in df.columns:

                df = df.rename(
                    columns={
                        candidate: target
                    }
                )

                break

    return df


# =============================================================================
# VALIDERING AV EXCEL
# =============================================================================

def validate_excel(df):
    """
    Kontrollerar Excel innan Visma påverkas.
    """

    required_columns = [
        "KundID",
        "Fakturanr",
        "Fakturadatum",
        "Belopp"
    ]

    missing = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing:

        print()
        print(
            "Excel-filen saknar obligatoriska kolumner:"
        )

        for column in missing:

            print(
                f"  - {column}"
            )

        return False

    # -------------------------------------------------------------------------
    # Normaliserade fakturanummer
    # -------------------------------------------------------------------------

    invoice_numbers = df[
        "Fakturanr"
    ].apply(
        normalize_invoice_number
    )

    # -------------------------------------------------------------------------
    # Tomma fakturanummer
    # -------------------------------------------------------------------------

    empty_mask = (
        invoice_numbers == ""
    )

    if empty_mask.any():

        print()
        print(
            "FEL: Excel innehåller tomma Fakturanr."
        )

        for index in df.index[
            empty_mask
        ]:

            print(
                f"  - Excel-rad {index + 2}"
            )

        return False

    # -------------------------------------------------------------------------
    # Ogiltiga fakturanummer
    # -------------------------------------------------------------------------

    invalid_mask = ~invoice_numbers.apply(
        lambda value: bool(
            re.fullmatch(
                r"\d+",
                value
            )
        )
    )

    if invalid_mask.any():

        print()
        print(
            "FEL: Ogiltiga Fakturanr i Excel:"
        )

        for index in df.index[
            invalid_mask
        ]:

            print(
                f"  - Rad {index + 2}: "
                f"{invoice_numbers.loc[index]}"
            )

        return False

    # -------------------------------------------------------------------------
    # Dubletter
    # -------------------------------------------------------------------------

    duplicate_mask = invoice_numbers.duplicated(
        keep=False
    )

    if duplicate_mask.any():

        duplicates = sorted(
            set(
                invoice_numbers[
                    duplicate_mask
                ].tolist()
            )
        )

        print()
        print(
            "FEL: Excel innehåller dubbla Fakturanr:"
        )

        for number in duplicates:

            print(
                f"  - {number}"
            )

        print()
        print(
            "Importen startas inte."
        )

        return False

    return True


# =============================================================================
# LOGG
# =============================================================================

def save_log(
    file_path,
    log_rows
):
    """
    Sparar importlogg bredvid Excel-filen.
    """

    if not log_rows:
        return None

    log_df = pd.DataFrame(
        log_rows
    )

    log_file = (
        file_path.parent
        /
        f"visma_import_logg_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    log_df.to_excel(
        log_file,
        index=False
    )

    return log_file


# =============================================================================
# MAIN
# =============================================================================

def main():

    print()
    print("=" * 80)
    print(
        "VISMA COMPACT - AUTOMATISK FAKTURAREGISTRERING"
    )
    print("=" * 80)

    print()
    print(
        "Flöde:"
    )

    print(
        "KundID -> kontroll Nr. -> Bokf.dag -> "
        "Fak.belopp -> TAB x5 -> NUM+"
    )

    print()
    print(
        "Excel Fakturanr måste vara exakt samma som Visma Nr."
    )

    print()

    # =========================================================================
    # Excel-fil
    # =========================================================================

    excel_path = input(
        "Ange sökväg till Excel-filen: "
    )

    file_path = resolve_excel_path(
        excel_path
    )

    if file_path is None:

        print()
        print(
            f"Filen hittades inte: {excel_path}"
        )

        return

    # =========================================================================
    # Läs Excel
    # =========================================================================

    try:

        df = pd.read_excel(
            file_path
        )

    except Exception as e:

        print()
        print(
            f"Kunde inte läsa Excel-filen: {e}"
        )

        return

    df = apply_column_aliases(
        df
    )

    if not validate_excel(
        df
    ):
        return

    total_rows = len(
        df
    )

    # =========================================================================
    # Testläge
    # =========================================================================

    if TEST_LIMIT > 0:

        df = df.head(
            TEST_LIMIT
        )

        print()
        print(
            "TESTLÄGE AKTIVERAT"
        )

        print(
            f"Kör {len(df)} av totalt "
            f"{total_rows} fakturor."
        )

        print(
            "Ändra TEST_LIMIT = 0 "
            "när första fakturan fungerar korrekt."
        )

    # =========================================================================
    # Kontrollera Visma
    # =========================================================================

    hwnd, visma_title = find_visma_window()

    if not hwnd:

        print()
        print(
            "FEL: Visma Compact hittades inte."
        )

        print(
            "Öppna Visma -> Kundreskontra -> Kundbokning."
        )

        return

    print()
    print(
        f"Visma hittat: {visma_title}"
    )

    print()
    print(
        f"Antal fakturor: {len(df)}"
    )

    print()
    print(
        "Efter Fak.belopp används EXAKT 5 Tab och därefter Num+."
    )

    print()
    print(
        "NÖDSTOPP: flytta muspekaren till skärmens övre vänstra hörn."
    )

    print()

    input(
        "Tryck Enter för att starta..."
    )

    # =========================================================================
    # Nedräkning
    # =========================================================================

    print()

    for seconds in range(
        3,
        0,
        -1
    ):

        print(
            f"Startar om {seconds}..."
        )

        time.sleep(1)

    # =========================================================================
    # IMPORT
    # =========================================================================

    log_rows = []

    stopped = False

    for index, row in df.iterrows():

        excel_fakturanr = normalize_invoice_number(
            row["Fakturanr"]
        )

        kund_id = normalize_customer_id(
            row["KundID"]
        )

        visma_fakturanr = ""

        print()
        print("-" * 80)

        print(
            f"Rad {index + 2} | "
            f"KundID: {kund_id} | "
            f"Excel Fakturanr: {excel_fakturanr}"
        )

        try:

            visma_fakturanr = register_invoice(
                row
            )

            # Tom retur = fakturan sparades men Visma Nr. kunde inte läsas.
            nummerkontroll = (
                "MATCH"
                if visma_fakturanr
                else "EJ LÄST"
            )

            log_rows.append({

                "Rad": index + 2,

                "KundID": kund_id,

                "Excel Fakturanr": excel_fakturanr,

                "Visma Fakturanr": (
                    visma_fakturanr
                    or excel_fakturanr
                ),

                "Nummerkontroll": nummerkontroll,

                "Status": "OK",

                "Fel": ""
            })

            print(
                f"  SPARAD: Faktura "
                f"{visma_fakturanr or excel_fakturanr}"
            )

            time.sleep(
                0.40
            )

        # =====================================================================
        # FAILSAFE
        # =====================================================================

        except pyautogui.FailSafeException:

            error_text = (
                "PyAutoGUI failsafe aktiverades."
            )

            print()
            print(
                error_text
            )

            log_rows.append({

                "Rad": index + 2,

                "KundID": kund_id,

                "Excel Fakturanr": excel_fakturanr,

                "Visma Fakturanr": visma_fakturanr,

                "Nummerkontroll": "AVBRUTEN",

                "Status": "STOPPAD",

                "Fel": error_text
            })

            stopped = True

            break

        # =====================================================================
        # MANUELLT STOPP
        # =====================================================================

        except KeyboardInterrupt:

            error_text = (
                "Importen avbröts manuellt."
            )

            print()
            print(
                error_text
            )

            log_rows.append({

                "Rad": index + 2,

                "KundID": kund_id,

                "Excel Fakturanr": excel_fakturanr,

                "Visma Fakturanr": visma_fakturanr,

                "Nummerkontroll": "AVBRUTEN",

                "Status": "STOPPAD",

                "Fel": error_text
            })

            stopped = True

            break

        # =====================================================================
        # ANNAT FEL
        # =====================================================================

        except Exception as e:

            error_text = str(
                e
            )

            print()
            print("=" * 80)
            print(
                "IMPORTEN STOPPAD"
            )
            print("=" * 80)

            print()
            print(
                f"Rad: {index + 2}"
            )

            print(
                f"KundID: {kund_id}"
            )

            print(
                f"Excel Fakturanr: {excel_fakturanr}"
            )

            print()
            print(
                f"Orsak: {error_text}"
            )

            print()
            print(
                "Ingen Num+ skickades efter felet."
            )

            log_rows.append({

                "Rad": index + 2,

                "KundID": kund_id,

                "Excel Fakturanr": excel_fakturanr,

                "Visma Fakturanr": visma_fakturanr,

                "Nummerkontroll": "FEL",

                "Status": "STOPPAD",

                "Fel": error_text
            })

            stopped = True

            break

    # =========================================================================
    # LOGG
    # =========================================================================

    log_file = None

    try:

        log_file = save_log(
            file_path,
            log_rows
        )

    except Exception as e:

        print()
        print(
            f"Kunde inte spara loggfilen: {e}"
        )

    # =========================================================================
    # SUMMERING
    # =========================================================================

    success_count = sum(
        1
        for item in log_rows
        if item["Status"] == "OK"
    )

    error_count = sum(
        1
        for item in log_rows
        if item["Status"] != "OK"
    )

    print()
    print("=" * 80)

    if stopped:

        print(
            "KÖRNINGEN STOPPAD"
        )

    else:

        print(
            "KÖRNINGEN KLAR"
        )

    print("=" * 80)

    print()
    print(
        f"Sparade fakturor: {success_count}"
    )

    print(
        f"Stoppade/fel: {error_count}"
    )

    if log_file:

        print()
        print(
            "Logg sparad:"
        )

        print(
            log_file
        )

    print()


# =============================================================================
# START
# =============================================================================

if __name__ == "__main__":
    main()