# visma_sammanstall_loggar.py

Sammanställer Visma-loggfiler (`visma_inbetalningar_logg_*_*.csv`) till **en** CSV-fil
för en vald månad, ett valt år och en vald betalningstyp.

- **Fil:** `visma_import/visma_sammanstall_loggar.py`
- **Beroenden:** Inga för kärnflödet – endast Pythons standardbibliotek (Python 3.7+).
  `openpyxl` krävs **endast** om jämförelsefilen (steg 14) är en Excel-fil (`.xlsx`/`.xlsm`);
  installeras vid behov med `pip install openpyxl`. CSV fungerar utan externa beroenden.
- **Syfte:** Slå ihop de enskilda körloggarna som skapas av
  `visma_register_inbetalningar` till en samlad månadsfil, med dubblettkontroll,
  valfri dubblettrensning och valfri avstämning mot en export över betalda fakturor.

---

## Snabbstart

```bash
python visma_sammanstall_loggar.py
```

Skriptet frågar interaktivt efter tre saker:

1. **Full sökväg till mappen** där loggfilerna ligger.
2. **Månad och år**, t.ex. `2026-06`.
3. **Betalningstyp**: `Bank` eller `Skatteverket`.

Resultatet sparas i **samma mapp** som loggfilerna, t.ex.:

```
visma_inbetalningar_bank_juni_2026_logg.csv
visma_inbetalningar_skatteverket_juni_2026_logg.csv
```

Därefter erbjuder skriptet två valfria steg:

1. **Dubblettrensning** → `..._clean_logg.csv`
2. **Avstämning** mot en export över betalda fakturor → `..._avvikelser.csv`

Hela flödet:

```
Sammanställ  ->  (fråga) Skapa fil utan dubbletter?  ->  clean_logg
                                                   |
                                                   +->  (fråga) Jämföra med betalda fakturor?  ->  avvikelser
```

---

## Indata: loggfilernas format

Loggfilerna skapas av `visma_register_inbetalningar` och har följande kolumner
(fältavgränsare `;`, kodning `utf-8-sig`):

```
Rad;Betalningsdatum;VismaDatum;Kundnamn;KundID;Fakturanr;Belopp;Status;Felmeddelande;Tidpunkt
```

| Kolumn | Betydelse | Exempel |
|--------|-----------|---------|
| `Betalningsdatum` | Rådatum från indata (används för filtrering/sortering) | `2026-06-07` |
| `VismaDatum` | Visma-format (2-siffrigt år), reserv för datumfiltrering | `26-06-07` |
| `Fakturanr` | Fakturanummer (grund för dubblettkontroll) | `49561` |
| `Status` | Postens status – endast `OK` tas med | `OK` |

Filnamnsmönstret som söks är `visma_inbetalningar_logg_*_*` (t.ex.
`visma_inbetalningar_logg_20260607_101010.csv`).

---

## Krav och hur de uppfylls

| # | Krav | Hur det uppfylls |
|---|------|------------------|
| 1 | Fråga efter mapp, månad/år och betalningstyp vid start | `prompt_folder()`, `prompt_month_year()`, `prompt_payment_type()` |
| 2 | Kontrollera att mappen finns; annars tydligt fel + ny chans | `prompt_folder()` loopar tills sökvägen finns och är en mapp |
| 3 | Sök filer som matchar `visma_inbetalningar_logg_*_*` | `glob.glob()` med `LOG_GLOB` i vald mapp |
| 4 | Ta bara med poster med `Status = OK` och betalningsdatum i vald månad/år | `status_is_ok()` (tål gemener/versaler + mellanslag) och `parse_date()` + jämförelse av `year`/`month` |
| 5 | En rubrikrad, bevarade kolumner/format, sorterat stigande på datum | Läses/skrivs med `csv`-modulen som text; `included.sort(key=sort_key)`; en `writeheader()` |
| 6 | Dubblettkontroll på fakturanr: behåll alla, kolumn `Dubblett` (`JA`/`NEJ`), varna | `mark_duplicates()` grupperar per fakturanr och listar källfiler i varningar |
| 7 | Spara i loggmappen med rätt filnamn | `write_output()` bygger `visma_inbetalningar_{typ}_{månad}_{år}_logg.csv` |
| 8 | Tåla tomma filer, saknade kolumner, ogiltiga datum utan att krascha | Varje fil/rad hanteras i `try`/skip med varning i stället för avbrott |
| 9 | Visa sammanställning i terminalen | Statistikblock skrivs ut i slutet av `main()` |
| 10 | Fråga om en ny fil utan dubbletter ska skapas | `prompt_yes_no(...)` efter sammanställningen |
| 11 | Vid `Ja`: ta bort dubbletter (behåll första kronologiskt), varna vid olika belopp/datum, spara `..._clean_logg.csv`, visa antal borttagna | `build_clean()` + `write_output()` |
| 12 | Vid `Nej`: avsluta kontrollerat, behåll den sammanställda filen | `return 0` utan att röra filen |
| 13 | Fråga om jämförelse mot export över betalda fakturor | `prompt_yes_no(...)` efter clean-filen |
| 14 | Be om sökväg, kontrollera fil, stöd CSV+Excel, jämför på fakturanr, normalisera nummer | `prompt_existing_file()`, `read_comparison_file()`, `normalize_invoice()` |
| 15 | Identifiera avvikelser (`SAKNAS_I_VISMA`, `SAKNAS_I_BETALDA_FAKTUROR`, `BELOPPSAVVIKELSE`, `DATUMAVVIKELSE`) | `run_comparison()` mängdjämförelse + belopp/datum-kontroll |
| 16–17 | Skapa avvikelsefil med angivna kolumner, spara som `..._avvikelser.csv` | `run_comparison()` skriver med `AVVIKELSE_COLUMNS` |
| 18 | Visa avstämningssammanställning + sökväg | Statistikblock i `run_comparison()` |

---

## Så fungerar skriptet (steg för steg)

1. **Inmatning och validering.** Sökväg, period och betalningstyp läses in.
   Sökvägen valideras direkt; fel ger ny fråga. Perioden accepterar `2026-06`,
   `2026/06`, `2026 06` och `2026-6`.

2. **Filsökning.** Alla filer som matchar `visma_inbetalningar_logg_*_*` i mappen
   listas. En eventuell tidigare skapad utdatafil med samma namn utesluts.

3. **Inläsning per fil** (`read_log_file`): provar kodningarna `utf-8-sig` och
   `cp1252`, med binär reserv (`errors="replace"`) som sista utväg. Tomma filer
   och filer utan rubrikrad hoppas över med varning.

4. **Kolumnmatchning** (`find_column`): hittar rätt kolumner oberoende av
   stora/små bokstäver och extra mellanslag, med alias
   (t.ex. `Betalningsdatum` / `Datum` / `Bet.dag`). Saknas obligatorisk kolumn
   hoppas filen över (dess rader räknas som överhoppade).

5. **Urval per rad:**
   - `Status` måste vara `OK` (trimmat, versaloberoende).
   - Betalningsdatum tolkas av `parse_date()` (primärt `Betalningsdatum`, annars
     `VismaDatum`). Ogiltigt datum → raden hoppas över med varning.
   - Datumet måste ligga i vald månad **och** år.

6. **Sortering.** Inkluderade poster sorteras stigande på det tolkade datumet.
   Sorteringen är stabil, så poster med samma datum behåller fil-/radordning.

7. **Dubblettmarkering** (`mark_duplicates`): poster grupperas på fakturanummer.
   Fakturanummer som förekommer fler än en gång markeras `JA` i kolumnen
   `Dubblett`, övriga `NEJ`. Alla förekomster behålls, och en varning visar
   fakturanumret och vilka källfiler det förekom i. Tomt fakturanummer räknas
   inte som dubblett.

8. **Skrivning** (`write_output`): skrivs som `utf-8-sig` (BOM) med `;` så att
   Excel öppnar filen korrekt med svenska tecken. Ursprungliga kolumner bevaras
   och `Dubblett` läggs sist. Även 0 träffar ger en giltig fil med enbart
   rubrikrad.

9. **Sammanställning:** statistik och den fullständiga sökvägen skrivs ut.

---

## Dubblettrensning (steg 10–12)

Efter sammanställningen frågar skriptet: **"Vill du skapa en ny fil utan dubbletter?"**

- **Ja** → `build_clean()` går igenom de (redan kronologiskt sorterade) posterna och
  behåller **första** förekomsten per fakturanummer. Övriga tas bort. Om en borttagen
  dubblett har **annat belopp** eller **annat betalningsdatum** än den behållna visas
  en varning. Resultatet sparas som `..._clean_logg.csv` (alla rader får `Dubblett = NEJ`),
  och antalet borttagna dubbletter visas.
- **Nej** → skriptet avslutas kontrollerat och den sammanställda filen behålls oförändrad.

Tomt fakturanummer kan inte avgöras som dubblett och behålls alltid.

---

## Avstämning mot betalda fakturor (steg 13–18)

När clean-filen skapats frågar skriptet: **"Vill du jämföra filen med en export över
betalda fakturor?"**

- **Nej** → avslutar utan att ta bort clean-filen.
- **Ja** → skriptet ber om sökväg till jämförelsefilen (`prompt_existing_file`),
  läser den (`read_comparison_file`) och jämför mot clean-filens poster.

**Filstöd:** CSV (auto-detekterad avgränsare `;`, `,` eller tab) och Excel `.xlsx`/`.xlsm`
(kräver `openpyxl` – installeras vid behov med `pip install openpyxl`). Gamla `.xls`
stöds inte.

**Nyckel:** fakturanummer, normaliserat med `normalize_invoice()` (tar bort mellanslag
och avslutande `.0`/`,0`, så att Excel-tal som `49561.0` matchar `49561`).

**Avvikelsetyper:**

| Typ | Betydelse |
|-----|-----------|
| `SAKNAS_I_VISMA` | Finns i betalda-exporten men saknas i Visma-loggen (ev. oregistrerad inbetalning) |
| `SAKNAS_I_BETALDA_FAKTUROR` | Finns i Visma-loggen men saknas i jämförelsefilen |
| `BELOPPSAVVIKELSE` | Finns i båda men beloppen skiljer sig (jämförs som tal, `parse_amount`) |
| `DATUMAVVIKELSE` | Finns i båda men betalningsdatumen skiljer sig (jämförs som datum, `parse_date`) |

En faktura kan ge både en belopps- och en datumavvikelse (en rad per typ).

**Avvikelsefil** (`..._avvikelser.csv`, `;`-separerad, `utf-8-sig`) med kolumnerna:

```
Avvikelsetyp;Fakturanummer;Belopp_Visma;Belopp_Betalda_fakturor;Betalningsdatum_Visma;Betalningsdatum_Betalda_fakturor;Kommentar
```

Till slut skrivs en avstämningssammanställning ut: antal fakturor per fil, matchande,
saknas i Visma, endast i Visma, belopps-/datumavvikelser och sökvägen till avvikelsefilen.

---

## Datumtolkning

`parse_date()` klarar bl.a.:

```
2026-04-07              2026/04/07
2026-04-07 00:00:00     07/04/2026   (svensk dag/månad/år)
26-04-07  (Visma)       20260407
```

Klockslag tas bort automatiskt. Går datumet inte att tolka hoppas raden över
(räknas som överhoppad + varning) i stället för att stoppa körningen.

---

## Utdata

Exempel på resultatfil (`;`-separerad, `utf-8-sig`):

```
Rad;Betalningsdatum;VismaDatum;Kundnamn;KundID;Fakturanr;Belopp;Status;Felmeddelande;Tidpunkt;Dubblett
2;2026-06-01;26-06-01;Kund G;1007;49570;750,00;OK;;2026-06-01 12:00:01;NEJ
2;2026-06-03;26-06-03;Kund B;1002;49562;1396,50;OK;;2026-06-03 10:10:11;NEJ
1;2026-06-07;26-06-07;Kund A;1001;49561;2017,00;OK;;2026-06-07 10:10:10;JA
1;2026-06-20;26-06-20;Kund F;1006;49561;2017,00;OK;;2026-06-20 12:00:00;JA
```

Terminalens sammanställning:

```
============================================================
 SAMMANSTALLNING
============================================================
  Hittade loggfiler:          4
  Behandlade filer:           2
  Lasta poster:               8
  Inkluderade (status OK):    4
  Overhoppade poster:         4
  Dubblettposter (markerade): 2
  Skapad fil:                 ...\visma_inbetalningar_bank_juni_2026_logg.csv
============================================================
```

---

## Skapade filer (översikt)

| Fil | Skapas när | Innehåll |
|-----|-----------|----------|
| `visma_inbetalningar_{typ}_{månad}_{år}_logg.csv` | Alltid | Alla OK-poster i perioden, med `Dubblett`-markering |
| `visma_inbetalningar_{typ}_{månad}_{år}_clean_logg.csv` | Vid `Ja` på steg 10 | Samma poster utan dubbletter (första per fakturanr) |
| `visma_inbetalningar_{typ}_{månad}_{år}_avvikelser.csv` | Vid `Ja` på steg 13 | Avvikelser mot betalda-exporten |

---

## Att tänka på

- **Betalningstypen styr endast filnamnet.** Det finns ingen kolumn i loggen som
  skiljer Bank från Skatteverket, så urvalet sker på datum + status. Tanken är
  att du pekar ut den mapp som innehåller rätt typ av loggar. Ska typen i stället
  filtrera på ett fält behöver skriptet justeras.
- **Inga poster tas bort.** Dubbletter behålls alltid och markeras bara – ingen
  data raderas.
- **Robust mot dålig indata.** Tomma filer, saknade kolumner och ogiltiga datum
  ger varningar men avbryter aldrig körningen.

---

## Konfiguration (i skriptets topp)

| Konstant | Standard | Beskrivning |
|----------|----------|-------------|
| `LOG_GLOB` | `visma_inbetalningar_logg_*_*` | Mönster för loggfiler som ska sammanställas |
| `CSV_DELIMITER` | `;` | Fältavgränsare för in- och utdata |
| `READ_ENCODINGS` | `utf-8-sig`, `cp1252` | Kodningar som provas vid inläsning |
| `OUTPUT_ENCODING` | `utf-8-sig` | Kodning för utdatafilen |
| `DUPLICATE_COLUMN` | `Dubblett` | Namn på den tillagda dubblettkolumnen |
| `BETALNINGSDATUM_ALIASES`, `STATUS_ALIASES`, `FAKTURANR_ALIASES`, `BELOPP_ALIASES`, `VISMADATUM_ALIASES` | – | Tillåtna rubrikvarianter i loggfilerna |
| `CMP_FAKTURANR_ALIASES`, `CMP_BELOPP_ALIASES`, `CMP_DATUM_ALIASES` | – | Bredare rubrikvarianter för jämförelsefilen (betalda fakturor) |
| `AVVIKELSE_COLUMNS` | 7 kolumner | Kolumnordning i avvikelsefilen |
| `PAYMENT_TYPES` | `bank`, `skatteverket` | Giltiga betalningstyper |
| `SWEDISH_MONTHS` | 1–12 | Svenska månadsnamn för filnamnet |

### Funktionsöversikt

| Funktion | Ansvar |
|----------|--------|
| `prompt_folder` / `prompt_month_year` / `prompt_payment_type` | Inmatning steg 1–2 |
| `prompt_yes_no` / `prompt_existing_file` | Ja/Nej-frågor och filväg (steg 10, 13, 14) |
| `collect_rows` | Läser loggfiler, filtrerar på OK + period (steg 3–4) |
| `read_log_file` / `find_column` / `parse_date` / `status_is_ok` | Inläsning och tolkning |
| `mark_duplicates` | Markerar `Dubblett` JA/NEJ (steg 6) |
| `write_output` | Skriver logg- och clean-filen (steg 5, 7, 11) |
| `build_clean` | Tar bort dubbletter, varnar vid belopps-/datumskillnad (steg 11) |
| `normalize_invoice` / `parse_amount` | Normaliserar fakturanr och belopp för jämförelse |
| `read_comparison_file` / `read_excel_rows` | Läser jämförelsefil (CSV/Excel) (steg 14) |
| `run_comparison` | Avstämning, skriver avvikelsefil, visar sammanställning (steg 15–18) |
