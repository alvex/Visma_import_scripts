# sammanstall_betalningar.py

## Översikt

- **Skript:** `sammanstall_betalningar.py` (projektroten)
- **Syfte:** Slår ihop alla rensade Bankgiro-filer (`clean_bet_lista_*.xlsx`) till en gemensam Excel-fil `samlade_betalningar_<datum>.xlsx` med samtliga betalningsrader, löpnummer och filtotaler.
- **Användning:** Steg 2 i Bankgiro-kedjan. Körs normalt via `dashboard.py` (menyval 2 eller A), eller fristående.

## Funktion och användning

### Huvudflöde

1. Hittar `clean_bet_lista_*.xlsx` i indatamappen (sorterade på filnamn).
2. Per fil (första arbetsbladet som har betalningstabell):
   - Letar upp **datum** i sektionen "Insättningsuppgift" (annars var som helst i bladet, annars ur filnamnet `_YYYYMMDD`).
   - Letar upp **filtotal** och **löpnummer** i sektionen "Belopp insatt på konto".
   - Läser betalningstabellen med rubrikerna `Avsändare`, `Betalningsreferens`, `Belopp` tills en tom rad eller ny sektion.
3. Sorterar alla rader på datum (fallande) och skriver `samlade_betalningar_<YYYY-MM-DD>.xlsx` (ark `Betalningar`) med kolumnerna `Datum, Avsändare, Betalningsreferens, Belopp, Löpnummer, Total`. Löpnummer/Total visas bara på första raden per källfil. Sist läggs en `SUMMA`-rad med summan av filtotalerna.
4. Om filnamnet redan finns läggs suffix `_2`, `_3` … på (ingen överskrivning).

### Startkommandon

```powershell
python sammanstall_betalningar.py --input "C:\...\processed" --output "C:\...\processed"
```

`--output` är valfritt (standard = `--input`). Utan `--input` används skriptmappen.

### Indata/utdata

| Typ | Format |
|---|---|
| Indata | `clean_bet_lista_*.xlsx` (skapade av `clean_bankgiro_files.py`) |
| Utdata | `samlade_betalningar_<datum>[_n].xlsx` med formatering (fet rubrikrad, frysta rutor, autofilter, kolumnbredder, talformat) |

### Beroenden

- `openpyxl`.

### Felhantering

- Fil som inte kan öppnas loggas som fel; körningen fortsätter.
- Saknat datum/total/löpnummer/tabell ger varningar per fil.
- Inga giltiga rader alls → ingen utdatafil, returkod 1.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Robust sektions-/rubriksökning, datumtolkning i flera format inkl. filnamn, unika utdatanamn, formaterad sammanställning.

**Brister/risker med stöd i koden:**
- `SUMMA`-raden är summan av **filtotalerna** ("Belopp insatt på konto"), inte av de exporterade radbeloppen. Om en fils radbelopp inte summerar till filtotalen upptäcks det inte – ingen avstämning rad ↔ total görs.
- Endast det första arbetsbladet med betalningstabell per fil läses (`break` efter första träffen) – fler blad i samma fil ignoreras tyst.
- Filer med fel (`has_errors`) utesluts ur totalsumman men varnings-filer (t.ex. saknad total) räknas som 0 i summan utan tydlig markering i utdatafilen.
- Gruppindelningen (Löpnummer/Total på första raden per fil) bygger på att raderna från samma fil ligger i följd efter sorteringen; detta håller eftersom varje fil har ett datum och sorteringen är stabil, men bryts om en fil någon gång innehåller flera datum.

**Behöver verifieras:**
- Att antagandet "ett datum per fil" stämmer för alla Bankgiro-exporter.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| SB-1 | Avstämning per fil: summera radbelopp och jämför mot filtotalen, varna vid differens | Fångar ofullständigt inlästa tabeller | Hög | Planerad | – |
| SB-2 | Markera i utdatafilen när en källfil saknar total/löpnummer | I dag syns det bara i terminalen | Låg | Planerad | – |
| SB-3 | Verifiera "ett datum per fil"-antagandet | Gruppering av Löpnummer/Total beror på det | Medel | Behöver verifieras | Genomgång av verkliga exportfiler |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Nytt filflöde: läser från och sparar i `processed` (ingen kodändring – styrs av `--input`/`--output` från dashboarden) | Mellanresultat samlas i `processed` | Körd mot rensad testfil i `processed` – sammanställningen hamnade i `processed` | – |
