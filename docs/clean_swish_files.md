# clean_swish_files.py

## Översikt

- **Skript:** `clean_swish_files.py` (projektroten)
- **Syfte:** Rensar Swish-transaktionsexporter (`Transaktioner * Swish*.xlsx`) till en standardiserad lista med endast kolumnerna `Bokförd, Typ, Avsändare, Meddelande, Insättningar, Summa`, med perioden överst.
- **Användning:** Steg 4 i Swish-kedjan. Körs normalt via `dashboard.py` (menyval 4 eller B), eller fristående.

## Funktion och användning

### Huvudflöde

1. Hittar filer som matchar `Transaktioner * Swish*.xlsx` i indatamappen.
2. Per fil (första arbetsbladet):
   - Letar upp periodraden (`Period: YYYY-MM-DD - YYYY-MM-DD`). Saknas den används dagens datum i filnamnet och en fel-rad skrivs till stderr (körningen fortsätter).
   - Hittar rubrikraden som innehåller flest av de önskade kolumnerna; **alla sex kolumner måste finnas**, annars avbryts filen med fel.
   - Läser raderna under rubriken: tomma rader och upprepade rubrikrader tas bort; datum formateras `YYYY-MM-DD`; belopp behålls som tal eller svensk kommasträng.
3. Skriver `clean_swish_lista_<periodslut>[_n].xlsx` (ark `Swish`): periodrad på rad 1, rubriker på rad 3, data från rad 4. Befintliga filnamn får suffix i stället för att skrivas över.

### Startkommandon

```powershell
python clean_swish_files.py --input "C:\...\arbetsmapp" --output "C:\...\arbetsmapp\processed"
```

`--output` är valfritt (standard: `<input>\processed`; skapas automatiskt).

### Indata/utdata

| Typ | Format |
|---|---|
| Indata | `Transaktioner * Swish*.xlsx` (bankens Swish-export) |
| Utdata | `clean_swish_lista_<YYYY-MM-DD>[_n].xlsx` |

### Beroenden

- `openpyxl`.

### Felhantering

- Fil som inte kan öppnas eller saknar tabell/kolumner räknas som fel; övriga filer bearbetas vidare. Returkod 1 om någon fil misslyckades.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Periodsökning, rubrikradsdetektering ("flest träffar vinner"), borttagning av tomrader/dubbelrubriker, unika utdatanamn, transaktionsräkning.

**Brister/risker med stöd i koden:**
- Alla sex kolumner är obligatoriska – om banken byter rubriknamn (t.ex. `Inbetalningar` i stället för `Insättningar`) avvisas filen.
- `clean_amount` behåller beloppstext som den är; blandade format (tal vs. text) förs vidare till nästa steg och hanteras först där.
- Saknad period gör att filnamnet baseras på körningsdagens datum, vilket kan ge missvisande namn.

**Behöver verifieras:**
- Att bankens aktuella Swish-export fortfarande har periodraden i formatet `Period: <start> - <slut>` och samma kolumnrubriker.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| CS-1 | Kolumnalias (t.ex. tolerans för rubrikvarianter) | Robusthet mot ändrad bankexport | Medel | Planerad | Exempel på varianter |
| CS-2 | Varna tydligare när periodslut saknas (filnamn blir dagens datum) | Spårbarhet | Låg | Planerad | – |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Standardutdatamapp ändrad från `<input>\edit` till `<input>\processed` | Nytt filflöde: mellanresultat i `processed`, färdiga CSV i `output` | Körd mot syntetisk testfil i temporär mapp – `clean_swish_lista_*.xlsx` hamnade i `processed` som skapades automatiskt | – |
