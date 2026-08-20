# Betalningsregistrering med Playwright

Det här verktyget läser den senaste filen som matchar `total_betalningar_*.xlsx` i en mapp du anger och använder Playwright för att öppna betalningssidor i Hemfresh fakturasystem.

Scriptet prioriterar säkerhet: det registrerar aldrig en betalning om beloppet från Excel inte matchar texten `Att betala:` på fakturasidan.

## Filer

- `payment_register.py` - huvudscriptet.
- `requirements.txt` - Python-paket som behövs.
- `payment_log_example.csv` - exempel på loggfilens struktur.
- `auth_state.json` - skapas efter manuell inloggning och innehåller session/cookies.
- `screenshots/` - skapas automatiskt vid fel där screenshot kan tas.

## Installation

Kör i PowerShell:

```powershell
cd "C:\Ekonomi\Fak 2026\April"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
```

## Kör scriptet

```powershell
python payment_register.py
```

Scriptet frågar:

```text
Ange sökvägen till mappen där total_betalningar_*.xlsx finns:
```

Exempel:

```text
C:\Ekonomi\Fak 2026\April\edit
```

Om flera `total_betalningar_*.xlsx` finns i mappen väljer scriptet den senast ändrade filen och skriver ut vilken fil som används.

## Första inloggning

Första gången öppnas Chromium i synligt läge. Logga in manuellt i fakturasystemet. När du ser fakturasystemet går du tillbaka till terminalen och trycker Enter.

Sessionen sparas i:

```text
auth_state.json
```

Behandla filen som känslig eftersom den kan innehålla inloggningscookies.

## Dry-run

Välj:

```text
1. Dry-run/testläge
```

Dry-run gör detta:

- läser Excel-filen
- öppnar fakturasidan
- kontrollerar fakturanummer när det kan läsas på sidan
- läser `Att betala:`
- jämför beloppet från Excel med fakturasidan
- fyller i formuläret om allt matchar
- klickar inte på `Betalningsprocess`
- loggar status `DRY_RUN_OK`

## Testa en rad först

När scriptet frågar:

```text
Hur många rader vill du behandla?
Skriv 1 för första testet eller ALLA för alla rader:
```

Tryck Enter eller skriv:

```text
1
```

Detta är den säkra första testen.

## Testa alla rader utan registrering

Kör:

```powershell
python payment_register.py
```

Välj dry-run och skriv:

```text
ALLA
```

Kontrollera loggfilen innan riktig registrering.

## Riktig registrering

Kör:

```powershell
python payment_register.py
```

Välj:

```text
2. Riktig registrering
```

Bekräfta genom att skriva exakt:

```text
REGISTRERA
```

Välj sedan antal rader. Börja gärna med `1` även vid riktig registrering.

Scriptet klickar bara på `Betalningsprocess` om:

- fakturanummer kunde tolkas från `Betalningsreferens`
- fakturasidan öppnas
- formuläret verkar finnas
- `Att betala:` finns och kan tolkas
- beloppet från Excel matchar `Att betala:` numeriskt med två decimaler
- fakturan inte verkar vara redan betald

## Loggfil

Varje körning skapar en loggfil i samma mapp som Excel-filen:

```text
payment_log_YYYY-MM-DD_HHMMSS.csv
```

Kolumner:

- `Datum`
- `Avsändare`
- `Betalningsreferens`
- `Fakturanummer`
- `Belopp_fran_Excel`
- `Belopp_pa_fakturasidan`
- `Status`
- `Meddelande`
- `Tidpunkt`
- `Screenshot`

Exempel på status:

- `DRY_RUN_OK`
- `REGISTRERAD`
- `BELOPP_MATCHAR_INTE`
- `REDAN_BETALD`
- `SAKNAR_FAKTURANUMMER`
- `SAKNAR_DATUM`
- `SAKNAR_FORMULAR`
- `FEL_FAKTURA`
- `TIMEOUT`
- `FEL`
- `OKLAR_STATUS`

## Screenshots

Vid vissa fel sparas screenshot i:

```text
screenshots
```

Filnamnet innehåller fakturanummer, status och tidpunkt, till exempel:

```text
49782_BELOPP_MATCHAR_INTE_2026-05-15_103012.png
```

## Säker teststrategi

1. Kör `python payment_register.py`.
2. Välj dry-run.
3. Behandla `1` rad.
4. Kontrollera webbläsaren och loggfilen.
5. Kör igen, välj dry-run och `ALLA`.
6. Kontrollera loggfilen.
7. Kör riktig registrering först när dry-run ser korrekt ut.

Hellre att scriptet hoppar över en rad och loggar felet än att fel betalning registreras.
