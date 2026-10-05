# samla_swish_betalningar.py

## Översikt

- **Skript:** `samla_swish_betalningar.py` (projektroten)
- **Syfte:** Slår ihop alla rensade Swish-filer (`clean_swish_lista_*.xlsx`) till en gemensam fil `samla_swish_bet_lista_<datum>.xlsx`. Period och filtotal visas en gång per källfil. Eget separat steg – rör inte Bankgiro-flödet.
- **Användning:** Steg 5 i Swish-kedjan. Körs normalt via `dashboard.py` (menyval 5 eller B), eller fristående.

## Funktion och användning

### Huvudflöde

1. Hittar `clean_swish_lista_*.xlsx` i indatamappen (sorterade på filnamn).
2. Per fil (första arbetsbladet): läser periodraden samt tabellen med kolumnerna `Bokförd, Avsändare, Meddelande, Insättningar` (alla fyra krävs). Endast rader med giltigt belopp i `Insättningar` tas med – totalraden (bara `Summa` ifylld) och tomrader hoppas över automatiskt.
3. Skriver `samla_swish_bet_lista_<YYYY-MM-DD>[_n].xlsx` (ark `Swish`) med kolumnerna `Bokförd, Avsändare, Meddelande, Insättningar, Period, Filtotal`. Period/Filtotal skrivs bara på gruppens första rad; filtotalen **beräknas som summan av filens insättningsrader**. Sist läggs en `SUMMA`-rad med totalsumman.

### Startkommandon

```powershell
python samla_swish_betalningar.py --input "C:\...\processed" --output "C:\...\processed"
```

`--output` är valfritt (standard = `--input`).

### Beroenden

- `openpyxl`.

### Felhantering

- Fil som inte kan öppnas eller saknar tabellen räknas som fel; övriga bearbetas. Returkod 1 om någon fil misslyckades.
- Inga transaktioner alls → ingen utdatafil, returkod 1.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Beloppstolkning av både tal och svenska kommasträngar, gruppvis period/filtotal, formaterad utdatafil med SUMMA-rad, unika utdatanamn.

**Brister/risker med stöd i koden:**
- Filtotalen beräknas av skriptet (summan av raderna) och stäms **inte** av mot källfilens egen `Summa`-kolumn – en avvikelse där upptäcks inte.
- Kolumnen `Typ` från den rensade filen följer inte med; ingen filtrering på transaktionstyp görs (alla rader med belopp i `Insättningar` antas vara inbetalningar).
- Rad utan tolkbart belopp hoppas över tyst (ingen varning per rad).

**Behöver verifieras:**
- Att `Insättningar` alltid är tom på icke-betalningsrader i den rensade filen (annars kan fel rader följa med).

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| SS-1 | Stämma av beräknad filtotal mot källfilens `Summa` | Avstämningsskydd motsvarande Bankgiro-förslaget SB-1 | Hög | Planerad | – |
| SS-2 | Logga rader som hoppas över p.g.a. otolkbart belopp | Spårbarhet | Låg | Planerad | – |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Nytt filflöde: läser från och sparar i `processed` (endast docstring-exempel ändrat i koden – styrs av `--input`/`--output`) | Mellanresultat samlas i `processed` | Körd mot rensad testfil i `processed` – sammanställningen hamnade i `processed` | – |
