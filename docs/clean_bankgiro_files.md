# clean_bankgiro_files.py

## Översikt

- **Skript:** `clean_bankgiro_files.py` (projektroten)
- **Syfte:** Skapar rensade kopior av Bankgirots insättningsfiler (`Bg*_Insättningsuppgifter_*.xlsx`) utan att röra originalen. Tar bort känsliga/onödiga uppgifter så att filerna kan sammanställas vidare.
- **Användning:** Steg 1 i Bankgiro-kedjan. Körs normalt via `dashboard.py` (menyval 1 eller A), eller fristående.

## Funktion och användning

### Huvudflöde

1. Hittar filer som matchar `Bg*_Insättningsuppgifter_*.xlsx` i indatamappen (temporära `~$`-filer och filer som redan ligger i utdatamappen ignoreras).
2. Per fil och arbetsblad:
   - **Sektionen "Insättningsuppgift":** alla celler töms utom kolumnen under rubriken `Datum`.
   - **Sektionen "Belopp insatt på konto":** endast kolumnerna `Belopp` och `Löpnummer…` behålls.
   - **Sektionen "Insättningar":** kolumnen `Bankgironummer/Avinummer` tas bort helt.
3. Sparar kopian som `clean_bet_lista_<originalnamn>` i utdatamappen. Originalfilen ändras aldrig.
4. Skriver per-fil-logg och sammanfattning (hittade/bearbetade/varningar/fel) till terminalen.

### Körlägen

- **Utan flaggor:** bearbetar endast den hårdkodade testfilen `Bg819-5968_Insättningsuppgifter_20260415.xlsx` i skriptmappen (testläge).
- **`--all`, `--input` eller `--output`:** batch-läge – alla matchande filer i indatamappen bearbetas.

```powershell
python clean_bankgiro_files.py --input "C:\...\arbetsmapp" --output "C:\...\arbetsmapp\processed"
```

### Indata/utdata

| Typ | Format |
|---|---|
| Indata | `Bg*_Insättningsuppgifter_*.xlsx` (Bankgirots exportlayout med sektionerna ovan) |
| Utdata | `clean_bet_lista_<originalnamn>.xlsx` i utdatamappen (standard: `<skriptmapp>\processed`; skapas automatiskt) |

### Beroenden

- `openpyxl`. Ingen konfigfil.

### Felhantering

- Fel i en fil loggas och nästa fil bearbetas (returkod 1 om någon fil fick fel).
- Saknade sektioner/rubriker ger varningar i terminalen, inte stopp.
- Varningen "Print area cannot be set" från openpyxl filtreras bort medvetet.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Rensning av tre sektioner, skydd av originalfiler, batch- och testläge, varnings-/felsammanfattning.

**Brister/risker med stöd i koden:**
- Testläget (körning utan flaggor) kräver den hårdkodade filen `Bg819-5968_Insättningsuppgifter_20260415.xlsx` i skriptmappen – saknas den avbryts körningen. Datumet antyder att filen är inaktuell.
- Befintlig utdatafil med samma namn skrivs över utan varning (`workbook.save`).
- Cellstilar/format följer med `load_workbook` men formler ersätts inte av värden (`data_only` används inte här) – rensade celler som refererades av formler kan ge `#REF!`-liknande fel i kopian.

**Behöver verifieras:**
- Att Bankgirots aktuella exportlayout fortfarande använder exakt sektionsrubrikerna `Insättningsuppgift`, `Belopp insatt på konto` och `Insättningar`.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| CB-1 | Ta bort/parametrisera den hårdkodade testfilen | Testläget är beroende av en fil från april 2026 | Låg | Planerad | Besluta om testläget behövs alls när dashboard används |
| CB-2 | Varna innan befintlig utdatafil skrivs över | Skydd mot oavsiktlig överskrivning | Låg | Planerad | – |
| CB-3 | Verifiera sektionsrubriker mot färsk Bankgiro-export | Layoutändring bryter rensningen tyst (bara varningar) | Medel | Behöver verifieras | Färsk exportfil |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Standardutdatamapp ändrad från `edit` till `processed` | Nytt filflöde: mellanresultat i `processed`, färdiga CSV i `output` | Körd mot syntetisk testfil i temporär mapp – rensad kopia hamnade i `processed` som skapades automatiskt | – |
