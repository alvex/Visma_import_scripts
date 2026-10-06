# sammanstall_fakturor.py

## Översikt

- **Skript:** `sammanstall_fakturor.py` (projektroten)
- **Syfte:** Sammanställer alla betalningsposter som förberetts för import under en kalendermånad (Bankgiro + Swish) till en kontrollista i Excel, med tydlig markering av återkommande fakturanummer. Rör aldrig källfilerna och ändrar ingen importlogik.
- **Användning:** Körs efter steg 3 och 6 (när registrerings-CSV:erna finns i `WORK\output`), fristående från terminalen. Kan köras om när som helst – omkörning mot oförändrat underlag skapar inga extra rader.

## Funktion och användning

### Huvudflöde

1. `--input` pekar på output-mappen med registrerings-CSV:erna, eller på månadsmappen (då används undermappen `output`).
2. Läser `betalningar_lista_to_reg_*.csv` och `swish_betalningar_for_registrering_*.csv`. Filer med `_visma` i namnet samt `~$`-temporärfiler exkluderas; loggfiler (`.txt`) matchas aldrig.
3. Skriver/uppdaterar `output\fakturor_lista\fakturor_<mappnamn>.xlsx`. **Tolkning av `<mappnamn>`:** månadsmappen är arbetsmappen som `dashboard.py` arbetar mot (t.ex. `sept 01`); eftersom CSV:erna ligger i `WORK\output` hämtas namnet från output-mappens förälder – aldrig namnet `output`.
4. Befintlig fil läses in och kompletteras; tidigare poster bevaras alltid.

### Kolumnmappning (ark `Fakturor`)

| Kolumn | Källa |
|---|---|
| Datum | `Datum` ur CSV:n (betalningsdatum), text `ÅÅÅÅ-MM-DD` |
| Avsändare | `Avsändare` ur CSV:n, trimmad |
| Fakturanummer | `Fakturanummer` ur CSV:n, lagrat som **text** (format `@`) så att inledande nollor bevaras; alla blanksteg (även hårda) tas bort |
| Belopp | `Belopp` ur CSV:n, numeriskt med två decimaler (format `#,##0.00`); svenska format (`1 234,50`) och negativa belopp hanteras, tecknet ändras aldrig |
| Status | `Förberedd` eller `Dubblett` (se nedan) |
| Source | Källfilens exakta filnamn inkl. filändelse |

Kolumnen `Betalningsreferens` i källfilerna används inte. `Förberedd` betyder endast att posten finns i importunderlaget – inte att importen är genomförd.

### Periodurval

Perioden är **kalendermånaden i postens betalningsdatum** (kolumnen `Datum`), aldrig dagens datum. Prioritet: `--month ÅÅÅÅ-MM` > perioden lagrad i befintlig fil (arket `Metadata`) > entydig månad i de nya posterna. Spänner underlaget över flera månader utan `--month` stoppar skriptet med en lista per månad. Giltiga rader utanför perioden tas inte med men rapporteras per källfil (antalet sparas i arket `Källfiler`). En rad med saknat/ogiltigt datum blir felrad – dagens datum används aldrig som ersättning.

### Dubblettlogik

- Samma fakturanummer i flera källposter ⇒ **samtliga** berörda rader markeras `Dubblett` (röd markering), även tidigare registrerade rader och även när belopp eller datum skiljer sig.
- Upptäcks inom en källfil och mellan Bankgiro- och Swish-filer (status räknas om över hela listan vid varje körning).
- Ingen rad tas bort, slås ihop eller stoppas – markeringen är en signal för manuell kontroll (t.ex. delbetalningar).

### Omkörningsskydd

Varje källfil identifieras med SHA-256 i arket `Källfiler` (filnamn, hash, tidpunkt, period, radantal). Vid körning:

- **Oförändrad, redan inläst fil** hoppas över ⇒ inga extra rader, även om filen innehåller två identiska rader (de bevarades som två separata förekomster vid första inläsningen).
- **Ändrad fil** (samma namn, annan hash) rapporteras med tydlig varningsruta och returkod 2; den läses **inte** in igen och ingenting skrivs över i tysthet.
- **Ny fil** läses in och kompletterar listan.

### Validering och felrader

Rader med saknat fakturanummer, ogiltigt datum eller ogiltigt belopp registreras inte som giltiga poster. De rapporteras i terminalen med källfil och radnummer och sparas i arket `Felrader` (gulmarkerade) med rådata.

### Arbetsbokens ark

| Ark | Innehåll |
|---|---|
| `Fakturor` | Kontrollistan: fryst rubrikrad, autofilter, kolumnbredder, dubbletter rödmarkerade, sorterad på datum + fakturanummer |
| `Felrader` | Ogiltiga rader med källfil, radnummer och problem |
| `Källfiler` | Metadata per inläst källfil (SHA-256 m.m.) – grunden för omkörningsskyddet |
| `Metadata` | Period, månadsmapp, antal poster, totalsumma (inkl. ev. dubbletter), senast uppdaterad |

Totalsumman omfattar **samtliga** rader; eventuella dubbletter ingår och räknas aldrig bort automatiskt.

### Säker filhantering

Skrivs via temporär fil (`.xlsx.tmp`) som ersätter målfilen atomiskt (`os.replace`) först när hela skrivningen lyckats. Är målfilen öppen i Excel visas ett begripligt felmeddelande, temporärfilen städas bort och befintliga data bevaras. En befintlig målfil som inte följer förväntat format (saknade ark/kolumner, otolkbara rader) stoppar körningen i stället för att skrivas över.

### Startkommandon

```powershell
python sammanstall_fakturor.py --input "C:\ekonomi\Fak_2026\sept 01\output"
python sammanstall_fakturor.py --input "C:\ekonomi\Fak_2026\sept 01" --month 2026-10
python sammanstall_fakturor.py --input "C:\ekonomi\Fak_2026\sept 01" --dry-run
```

Returkoder: `0` = OK, `1` = fel (inget skrevs), `2` = OK men ändrade källfiler upptäcktes.

### Beroenden

- `openpyxl`.

## Nuläge och kända problem

**Implementerat och verifierat genom körning (2026-10-07, syntetisk testdata + skarp körning mot `sept 01`):**
- Första körningen skapar filen på rätt plats; omkörning mot samma underlag ändrar inte antalet rader; ny källfil kompletterar befintlig sammanställning.
- Samma fakturanummer i Bankgiro och Swish markeras på båda raderna, retroaktivt även på tidigare `Förberedd`-rader.
- Inledande nollor (`00123`), svenska belopp (`1 234,50`), negativa belopp och månadsskiften (rad utanför perioden exkluderas och rapporteras) hanteras korrekt.
- Felrader rapporteras med källfil + radnummer; ändrad källfil ger varningsruta och returkod 2 utan att data ändras.
- Låst målfil (både vid läsning och skrivning) ger begripligt fel, ingen kvarlämnad temporärfil och intakt befintlig fil.

**Begränsningar:**
- Manuell formatering/egna kolumner i målfilen överlever inte en uppdatering (hela arbetsboken skrivs om; cellvärdena i de egna kolumnerna bevaras inte). Använd filen som kontrollista, inte som arbetsyta.
- En ändrad källfil läses aldrig in igen automatiskt – avsiktlig omläsning kräver att posterna/metadataraden för filen hanteras manuellt (eller att en ny fil med nytt namn skapas).
- En källfil som spänner över två månader får sina rader utanför perioden exkluderade (rapporteras och räknas i `Källfiler`); de dyker inte upp i någon annan månadsfil automatiskt.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| SF-1 | Menyval i `dashboard.py` (t.ex. "10) Sammanställ fakturor-lista") | Slipper fristående kommando | Låg | Planerad | Kräver ändring i dashboard (medvetet utelämnad i första versionen) |
| SF-2 | Flagga för avsiktlig omläsning av ändrad källfil (t.ex. `--reimport <fil>` som tar bort filens gamla rader först) | Idag krävs manuell hantering | Låg | Planerad | Måste förbli explicit – aldrig automatiskt |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-07 | Skriptet skapat | Kontrollista över förberedda betalningsposter per månad med dubblettmarkering | `py_compile`; åtta scenarier mot syntetisk testdata (första körning, omkörning, komplettering, dubbletter BG↔Swish, inledande nollor, svenska/negativa belopp, månadsskifte, felrader, ändrad källfil, låst målfil, torrkörning); skarp körning mot `Fak_2026\sept 01` (11 poster, 19 232,00) | – |
