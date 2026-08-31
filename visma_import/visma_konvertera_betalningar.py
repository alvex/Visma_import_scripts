import re
import sys

from pathlib import Path
from datetime import datetime

import pandas as pd


# =============================================================================
# VISMA - KONVERTERA BETALNINGSFILER TILL VISMA-FORMAT
# =============================================================================
#
# FLÖDE:
#
#   Fråga efter indata-mapp
#      ↓
#   Fråga efter utdata-mapp (skapas vid behov)
#      ↓
#   Läs alla .csv i indata-mappen (inte undermappar)
#      ↓
#   Per fil: byt namn på kolumner, sortera om, spara med "_visma"
#      ↓
#   Skriv logg till skärm och textfil i utdata-mappen
#
# =============================================================================


# =============================================================================
# GRUNDINSTÄLLNINGAR
# =============================================================================

# Byt namn på kolumnerna (original -> Visma). Alla dessa måste finnas i filen.
KOLUMN_BYT = {
    "Datum": "Betalningsdatum",
    "Avsändare": "Kundnamn",
    "Betalningsreferens": "KundID",
    "Fakturanummer": "Fakturanr",
    "Belopp": "Belopp",
}

# Exakt kolumnordning i utdata. Eventuella extra kolumner läggs sist.
KOLUMN_ORDNING = ["Betalningsdatum", "Fakturanr", "Belopp", "Kundnamn", "KundID"]

# Teckenkodningar som testas vid inläsning, i tur och ordning.
KODNINGAR = ["utf-8-sig", "cp1252"]

# Utdata skrivs alltid som UTF-8 med BOM.
UTDATA_KODNING = "utf-8-sig"


# =============================================================================
# HJÄLPFUNKTIONER
# =============================================================================

def fråga_mapp(text):
    """Fråga användaren efter en mappsökväg tills något anges."""
    while True:
        svar = input(text).strip().strip('"')
        if svar:
            return Path(svar)
        print("  Ange en sökväg.")


def hitta_kodning(sökväg):
    """Returnera första kodningen som kan läsa hela filen, annars None."""
    for kodning in KODNINGAR:
        try:
            with open(sökväg, "r", encoding=kodning) as f:
                f.read()
            return kodning
        except (UnicodeDecodeError, UnicodeError):
            continue
    return None


def hitta_avgränsare(sökväg, kodning):
    """Identifiera avgränsare från rubrikraden: semikolon eller komma.

    Semikolon väljs vid lika antal (vanligast i svenska exportfiler)."""
    with open(sökväg, "r", encoding=kodning) as f:
        rubrik = f.readline()
    if rubrik.count(";") >= rubrik.count(","):
        return ";"
    return ","


def tolka_belopp(värde):
    """Tolka ett belopp med mellanslag som tusental och komma som decimal.

    Returnerar float. Kastar ValueError om värdet inte kan tolkas."""
    s = str(värde).strip()
    if not s:
        return 0.0
    # Ta bort alla sorters mellanslag (inkl. hårt mellanslag \xa0).
    s = re.sub(r"\s", "", s.replace("\xa0", " "))
    # Komma är decimaltecken.
    s = s.replace(",", ".")
    return float(s)


# =============================================================================
# BEARBETA EN FIL
# =============================================================================

def bearbeta_fil(sökväg, utdata_mapp, logg):
    """Bearbeta en enskild CSV-fil.

    Returnerar (antal_poster, summa) om filen bearbetades,
    annars None om den hoppades över. Loggar alltid utfallet."""
    kodning = hitta_kodning(sökväg)
    if kodning is None:
        logg(f"  HOPPAR ÖVER {sökväg.name}: kunde inte läsas med "
             f"någon av kodningarna {KODNINGAR}.")
        return None

    avgränsare = hitta_avgränsare(sökväg, kodning)

    # Läs allt som text (dtype=str) så beloppen behålls exakt oförändrade.
    try:
        df = pd.read_csv(
            sökväg,
            sep=avgränsare,
            encoding=kodning,
            dtype=str,
            keep_default_na=False,
        )
    except Exception as fel:
        logg(f"  HOPPAR ÖVER {sökväg.name}: kunde inte läsas ({fel}).")
        return None

    # Kontrollera att alla förväntade kolumner finns (originalnamn).
    saknade = [k for k in KOLUMN_BYT if k not in df.columns]
    if saknade:
        logg(f"  HOPPAR ÖVER {sökväg.name}: saknar kolumner: "
             f"{', '.join(saknade)}.")
        return None

    # Byt namn och sortera om. Extra kolumner läggs sist, efter KundID.
    df = df.rename(columns=KOLUMN_BYT)
    extra = [c for c in df.columns if c not in KOLUMN_ORDNING]
    df = df[KOLUMN_ORDNING + extra]

    # Summera Belopp (oförändrat i filen, tolkat endast för summering).
    summa = 0.0
    fel_rader = 0
    for rad, värde in enumerate(df["Belopp"], start=2):  # rad 1 = rubrik
        try:
            summa += tolka_belopp(värde)
        except ValueError:
            fel_rader += 1

    # Spara med "_visma" före filändelsen, samma avgränsare, UTF-8 med BOM.
    ut_namn = f"{sökväg.stem}_visma{sökväg.suffix}"
    ut_sökväg = utdata_mapp / ut_namn
    df.to_csv(
        ut_sökväg,
        sep=avgränsare,
        encoding=UTDATA_KODNING,
        index=False,
    )

    antal = len(df)
    summa_txt = f"{summa:,.2f}".replace(",", " ").replace(".", ",")
    rad = f"  {sökväg.name}: {antal} poster, summa Belopp {summa_txt}  ->  {ut_namn}"
    if fel_rader:
        rad += f"  (OBS: {fel_rader} belopp kunde inte tolkas vid summering)"
    logg(rad)

    return antal, summa


# =============================================================================
# HUVUDPROGRAM
# =============================================================================

def main():
    print("=" * 60)
    print("  VISMA - KONVERTERA BETALNINGSFILER")
    print("=" * 60)

    indata_mapp = fråga_mapp("Sökväg till mappen med indata-filer: ")
    if not indata_mapp.is_dir():
        print(f"FEL: indata-mappen finns inte: {indata_mapp}")
        sys.exit(1)

    utdata_mapp = fråga_mapp("Sökväg till mappen där resultatet ska sparas: ")
    utdata_mapp.mkdir(parents=True, exist_ok=True)

    # Loggen skrivs både till skärm och till en lista som sparas till fil.
    logg_rader = []

    def logg(text):
        print(text)
        logg_rader.append(text)

    tidsstämpel = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logg("")
    logg(f"Konvertering startad {tidsstämpel}")
    logg(f"Indata:  {indata_mapp}")
    logg(f"Utdata:  {utdata_mapp}")
    logg("-" * 60)

    # Endast .csv direkt i mappen (inte undermappar).
    csv_filer = sorted(p for p in indata_mapp.glob("*.csv") if p.is_file())

    if not csv_filer:
        logg("Inga .csv-filer hittades i indata-mappen.")

    antal_bearbetade = 0
    antal_hoppade = 0
    totalsumma = 0.0

    for fil in csv_filer:
        try:
            resultat = bearbeta_fil(fil, utdata_mapp, logg)
        except Exception as fel:
            # Avbryt inte hela körningen på grund av ett fel i en enskild fil.
            logg(f"  HOPPAR ÖVER {fil.name}: oväntat fel ({fel}).")
            resultat = None

        if resultat is None:
            antal_hoppade += 1
        else:
            _, summa = resultat
            totalsumma += summa
            antal_bearbetade += 1

    # Sammanfattning.
    logg("-" * 60)
    logg(f"Filer bearbetade: {antal_bearbetade}")
    logg(f"Filer hoppade över: {antal_hoppade}")
    total_txt = f"{totalsumma:,.2f}".replace(",", " ").replace(".", ",")
    logg(f"Totalsumma Belopp (alla filer): {total_txt}")
    logg("=" * 60)

    # Skriv loggfil till utdata-mappen.
    logg_namn = f"konverteringslogg_{datetime.now():%Y%m%d_%H%M%S}.txt"
    logg_sökväg = utdata_mapp / logg_namn
    logg_sökväg.write_text("\n".join(logg_rader) + "\n", encoding=UTDATA_KODNING)
    print(f"\nLogg sparad: {logg_sökväg}")


if __name__ == "__main__":
    main()
