# convert_betalningar_to_csv.py

## Översikt

- **Skript:** `convert_betalningar_to_csv.py` (projektroten)
- **Syfte:** Konverterar den samlade Bankgiro-filen (`samlade_betalningar_*.xlsx`) till en semikolonseparerad CSV för betalningsregistrering, med extraherat fakturanummer per rad.
- **Användning:** Steg 3 i Bankgiro-kedjan. Körs normalt via `dashboard.py` (menyval 3 eller A), eller fristående.

## Funktion och användning

### Huvudflöde

1. `--input` kan vara en mapp (senast ändrade `samlade_betalningar_*.xlsx` väljs) eller en direkt `.xlsx`-fil. Positionsargument stöds bakåtkompatibelt.
2. Läser arket `Betalningar` (annars första arket). Kolumner identifieras via rubriknamn – obligatoriska: `Datum, Avsändare, Betalningsreferens, Belopp, Löpnummer, Total`.
3. Normaliserar värden: riktiga datum → `YYYY-MM-DD`, heltal utan `.0`, belopp skrivs ut **oförändrade** (ingen omräkning/avrundning).
4. Extraherar fakturanummer: exakt 5 siffror via regex `(?<!\d)\d{5}(?!\d)` (fångar även `F49457`). Flera träffar → sista väljs med `VARNING`; ingen träff → tomt fält med `VARNING`.
5. Rader utan avsändare, referens **och** belopp (t.ex. `SUMMA`-raden) hoppas över men loggas som `ÖVERHOPPAD`.
6. Skriver CSV `betalningar_lista_to_reg_<datum>[_n].csv` (semikolon, UTF-8 med BOM, CRLF) med kolumnerna `Datum;Avsändare;Betalningsreferens;Fakturanummer;Belopp`, samt loggfil `<csv-namn>_log.txt` med rad-för-rad-status och summering.

### Startkommandon

```powershell
python convert_betalningar_to_csv.py --input "C:\...\processed" --output "C:\...\output"
python convert_betalningar_to_csv.py --input "C:\...\samlade_betalningar_2026-10-01.xlsx" --date 2026-10-01
```

Utan `--output`: ligger indatafilen i en mapp som heter `processed` sparas CSV + logg i syskonmappen `output` (skapas automatiskt), annars bredvid indatafilen.

### Beroenden

- `openpyxl` (medvetet inte pandas – motiveras i filhuvudet: bevarar råa celltyper).

### Felhantering

- Tydliga felmeddelanden för saknad fil, fel filtyp, saknade kolumner, tom fil; returkod 1.
- Rader med saknat fakturanummer exporteras ändå (med tomt `Fakturanummer`) och flaggas i loggen.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Rubrikbaserad kolumnidentifiering, exakt beloppsbevarande, unika utdatanamn, utförlig loggfil med summa (svenskt format).

**Brister/risker med stöd i koden:**
- Rader utan fakturanummer följer med i CSV:n med tomt fält – nedströms (`visma_konvertera_betalningar.py` → `visma_register_inbetalningar_fixed.py`) stoppar registreringen på sådana rader först vid validering.
- Vid flera 5-siffriga tal väljs det **sista** – rätt val är inte garanterat, bara varnat.
- `Löpnummer` och `Total` krävs i indata men används inte i utdata – en i övrigt giltig fil utan dessa kolumner avvisas.
- Skriptet delar nästan all logik med `step_4_skapa_swish_csv.py` (duplicerad kod).

**Behöver verifieras:**
- Att fakturanummer i verksamheten alltid är exakt 5 siffror (jfr `payment_register.py` som tillåter 4–7 och `visma_register_inbetalningar_fixed.py` som tillåter 4–8).

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| CC-1 | Bryta ut gemensam logik med `step_4_skapa_swish_csv.py` till en delad modul | Två nästan identiska kodbaser att underhålla | Medel | Planerad | Regressionstest på båda flödena |
| CC-2 | Flagga/filtrera rader utan fakturanummer tydligare (t.ex. separat fil) | Undviker stopp längre fram i kedjan | Medel | Planerad | Beslut om önskat beteende |
| CC-3 | Göra `Löpnummer`/`Total` valfria vid inläsning | De används inte i utdata | Låg | Planerad | – |
| CC-4 | Enhetlig fakturanummer-definition i hela repot | Tre olika sifferlängdsregler finns i dag | Medel | Behöver verifieras | Bekräfta nummerserien |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Nytt filflöde: läser från `processed`, sparar färdig CSV i `output`. Standardutdata utan `--output` ändrad: `processed`-indata → syskonmappen `output` | Separera mellanresultat från färdiga filer | Körd mot syntetisk testdata – CSV + logg hamnade i `output` (både med explicit `--output` och med standardlogiken); innehåll kontrollerat (datum, fakturanummer, belopp korrekta) | – |
