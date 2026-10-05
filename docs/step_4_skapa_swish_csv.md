# step_4_skapa_swish_csv.py

## Översikt

- **Skript:** `step_4_skapa_swish_csv.py` (projektroten)
- **Syfte:** Skapar en CSV för betalningsregistrering från den samlade Swish-filen (`samla_swish_bet_lista_*.xlsx`). Producerar **samma CSV-format** som Bankgiro-steget (`convert_betalningar_to_csv.py`) så att registreringen kan konsumera båda likadant.
- **Användning:** Steg 6 i Swish-kedjan. Körs normalt via `dashboard.py` (menyval 6 eller B), eller fristående.

## Funktion och användning

### Huvudflöde

1. `--input` kan vara en mapp (senast ändrade `samla_swish_bet_lista_*.xlsx` väljs) eller en direkt `.xlsx`-fil.
2. Läser arket `Swish` (annars första arket). Obligatoriska kolumner: `Bokförd, Avsändare, Meddelande, Insättningar`.
3. Mappning till CSV: `Datum ← Bokförd`, `Avsändare ← Avsändare`, `Betalningsreferens ← Meddelande`, `Belopp ← Insättningar`, `Fakturanummer ←` 5-siffrigt tal extraherat ur `Meddelande` (regex `(?<!\d)\d{5}(?!\d)`; flera träffar → sista med varning).
4. Rader utan både avsändare och meddelande (t.ex. `SUMMA`-raden) hoppas över som `ÖVERHOPPAD`.
5. Skriver `swish_betalningar_for_registrering_<datum>[_n].csv` (semikolon, UTF-8 med BOM, CRLF) med kolumnerna `Datum;Avsändare;Betalningsreferens;Fakturanummer;Belopp` samt loggfil `<csv-namn>_log.txt`.

### Startkommandon

```powershell
python step_4_skapa_swish_csv.py --input "C:\...\processed" --output "C:\...\output"
python step_4_skapa_swish_csv.py --input "...\samla_swish_bet_lista_2026-10-01.xlsx" --date 2026-10-01
```

Utan `--output`: ligger indatafilen i en mapp som heter `processed` sparas CSV + logg i syskonmappen `output` (skapas automatiskt), annars bredvid indatafilen.

### Beroenden

- `openpyxl`.

### Felhantering

- Samma mönster som Bankgiro-konverteraren: tydliga fel för saknad fil/kolumner, returkod 1; rader med saknat fakturanummer exporteras med tomt fält och `VARNING` i loggen.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Rubrikbaserad kolumnidentifiering, beloppsbevarande utan omräkning, loggfil med rad-för-rad-status, unika utdatanamn.

**Brister/risker med stöd i koden:**
- Swish-meddelanden skrivs av kunder fritt – saknat/felaktigt 5-siffrigt tal ger tomt fakturanummer som stoppas först vid registreringens validering. Flera tal → sista väljs, vilket kan vara fel.
- Koden är nästan identisk med `convert_betalningar_to_csv.py` (duplicering, se roadmap CC-1 i det dokumentet).
- Till skillnad från Bankgiro-varianten räcker det att avsändare **eller** meddelande finns för att raden ska exporteras – en summarad med ifyllt belopp och text i någon av kolumnerna skulle följa med.

**Behöver verifieras:**
- Att `SUMMA`-raden i den samlade Swish-filen alltid saknar både avsändare och meddelande (annars exporteras den).

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| SW-1 | Dela kod med `convert_betalningar_to_csv.py` (gemensam modul) | Samma logik underhålls på två ställen | Medel | Planerad | Se CC-1 |
| SW-2 | Rapportera andel rader utan fakturanummer tydligare (t.ex. avsluta med varningsruta) | Swish-referenser är opålitliga | Medel | Planerad | – |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Nytt filflöde: läser från `processed`, sparar färdig CSV i `output`. Standardutdata utan `--output` ändrad: `processed`-indata → syskonmappen `output` | Separera mellanresultat från färdiga filer | Körd mot syntetisk testdata – CSV + logg hamnade i `output` (både med explicit `--output` och med standardlogiken) | – |
