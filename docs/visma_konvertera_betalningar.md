# visma_konvertera_betalningar.py

## Översikt

- **Skript:** `visma_import\visma_konvertera_betalningar.py`
- **Syfte:** Konverterar betalnings-CSV:er (från Bankgiro-/Swish-stegen) till Visma-format: byter kolumnnamn, sorterar om kolumnerna och sparar som `<namn>_visma.csv`.
- **Användning:** Steg 7 i flödet. Körs normalt via `dashboard.py` (menyval 7, som matar in output-mappen som både in- och utdatamapp), eller fristående.

## Funktion och användning

### Huvudflöde

1. Frågar interaktivt efter indatamapp och utdatamapp (utdatamappen skapas vid behov).
2. Läser **alla `.csv` direkt i indatamappen** (inte undermappar). Kodning testas i ordningen `utf-8-sig`, `cp1252`; avgränsare detekteras från rubrikraden (`;` eller `,`, semikolon vid lika).
3. Per fil: kontrollerar att kolumnerna `Datum, Avsändare, Betalningsreferens, Fakturanummer, Belopp` finns – annars hoppas filen över med loggrad.
4. Byter kolumnnamn enligt: `Datum→Betalningsdatum`, `Avsändare→Kundnamn`, `Betalningsreferens→KundID`, `Fakturanummer→Fakturanr`, `Belopp→Belopp`. Ny kolumnordning: `Betalningsdatum, Fakturanr, Belopp, Kundnamn, KundID` (extra kolumner läggs sist). **Observera:** betalningsreferensen hamnar alltså i kolumnen `KundID` – ett medvetet mappningsval, inte ett riktigt kund-ID.
5. Allt läses som text (`dtype=str`) så beloppen bevaras exakt; beloppen tolkas endast för summering i loggen.
6. Sparar `<namn>_visma.csv` i utdatamappen (samma avgränsare, UTF-8 med BOM) och skriver `konverteringslogg_<tidsstämpel>.txt` med antal poster och summa per fil samt totalsumma.

### Startkommandon

```powershell
python visma_import\visma_konvertera_betalningar.py
```

Ingen flagghantering – mapparna anges via prompt (dashboarden automatiserar detta via stdin).

### Beroenden

- `pandas`.

### Felhantering

- Fel i en enskild fil (kodning, läsfel, saknade kolumner) loggas som `HOPPAR ÖVER`; körningen fortsätter.
- Belopp som inte kan tolkas vid summering räknas och noteras i loggen men stoppar inget.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Kodnings- och avgränsardetektering, exakt beloppsbevarande, logg till både skärm och fil, robust per-fil-felhantering.

**Brister/risker med stöd i koden:**
- **Alla** `.csv` i indatamappen behandlas. I output-mappen kan det ligga andra CSV:er:
  - Redan konverterade `*_visma.csv` hoppas visserligen över (kolumnerna heter redan Visma-namn), men de genererar "HOPPAR ÖVER"-brus i loggen – och körs konverteringen mot en mapp där en fil råkar ha originalkolumnerna igen skapas `*_visma_visma.csv`.
  - `payment_log_*.csv` från `payment_register.py` innehåller kolumnerna `Datum, Avsändare, Betalningsreferens, Fakturanummer, Belopp...` och skulle **konverteras som om den vore en betalfil**. Loggarna sparas normalt i Excel-mappen, men hamnar de i output-mappen följer de med.
- Befintlig `*_visma.csv` med samma namn skrivs över utan varning.
- Mappningen `Betalningsreferens → KundID` kan förvåna – registreringsskriptet använder dock inte KundID för inmatning (endast logg), så effekten är begränsad.

**Behöver verifieras:**
- Att inga främmande CSV:er (loggar, exporter) förekommer i output-mappen vid körning via dashboarden.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| VK-1 | Filtrera indata till kända prefix (`betalningar_lista_to_reg_*`, `swish_betalningar_for_registrering_*`) i stället för `*.csv` | Hindrar att loggfiler/övriga CSV:er konverteras av misstag | Hög | Planerad | Bekräfta samtliga giltiga indataprefix |
| VK-2 | Hoppa över filer som redan slutar på `_visma.csv` explicit | Mindre loggbrus, skydd mot dubbelkonvertering | Medel | Planerad | – |
| VK-3 | Stöd för `--input`/`--output`-flaggor utöver prompt | Enhetligt med övriga skript, enklare för dashboarden | Låg | Planerad | Dashboard steg 7 behöver uppdateras samtidigt |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
| 2026-10-06 | Dashboarden matar nu in `WORK\output` (tidigare `edit`) som in- och utdatamapp – ingen kodändring i detta skript | Nytt filflöde: färdiga CSV:er ligger i `output` | Dashboard-anropet granskat; skriptet självt oförändrat och har inte körts | – |
