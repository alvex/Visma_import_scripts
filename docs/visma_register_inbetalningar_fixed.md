# visma_register_inbetalningar_fixed.py

## Översikt

- **Skript:** `visma_import\visma_register_inbetalningar_fixed.py`
- **Syfte:** Halvautomatisk registrering av **kundinbetalningar** i Visma Compact 6 (Kundreskontra → Inbetalningar) från en Excel-/CSV-fil, via GUI-automation (pywinauto + pyautogui). Skriver Bet.dag och Fakt.nr, hanterar dialogerna "Rätt belopp?" och "Differens" med exakt beloppsverifiering på öresnivå.
- **Användning:** Steg 8 i flödet. Körs via `dashboard.py` (menyval 8, torr/skarp) eller fristående. Detaljerad användarmanual finns i `visma_import\docs\Manual_Visma_inbetalningar.md` – detta dokument sammanfattar och kompletterar den, utan att duplicera innehållet.

## Funktion och användning

### Huvudflöde per rad

1. Fokuserar `Bet.dag` (via etikettmatchning i pywinauto eller sparad kalibrering) och skriver datumet i Visma-format `YY-MM-DD`.
2. Fokuserar `Fakt.nr`, skriver fakturanumret och trycker Enter → dialogen **"Rätt belopp?"** öppnas.
3. Kontrollerar att dialogen nämner rätt fakturanummer, skriver Excel-beloppet i "Ändra till:" och **läser tillbaka värdet** – endast exakt matchning på öresnivå godkänns, annars stopp utan OK.
4. Om beloppet avviker från fakturabeloppet måste dialogen **"Differens"** visas; differensen verifieras räknemässigt (Excel-belopp − Visma-belopp) och valet måste vara "Restbelopp på fakturan" innan OK klickas.
5. Väntar på att dialogerna stängts innan nästa rad. Allt loggas till `visma_inbetalningar_logg_<tidsstämpel>.csv` bredvid indatafilen.

Fältrensning sker med End → Shift+Home (Ctrl+A öppnar Artikelregistret i Vismas huvudfönster; i den modala beloppsdialogen används dock Ctrl+A säkert som reservmetod).

### Säkerhetslägen (konstanter i filen)

- `DRY_RUN = True` som standard – `--live` krävs för skarp körning, plus bekräftelsen `ja`.
- `AUTO_FINALIZE_PAYMENT_BATCH = False` – skriptet klickar **aldrig** Bankgiro/Bokför; slutbokföringen görs manuellt i Visma.
- `ALLOW_DUPLICATES = False` – dubbla fakturanummer i urvalet stoppar körningen.
- `MAX_INVOICES_PER_RUN = 50`.
- Nödstopp: muspekaren till övre vänstra hörnet (pyautogui failsafe) eller Ctrl+C.

### Kommandoradsflaggor

| Flagga | Funktion |
|---|---|
| `<fil>` | Excel (.xlsx/.xls/.xlsm) eller CSV (`;`, UTF-8/cp1252) |
| `--live` | Skarp körning (annars dry-run) |
| `--rader N` | Tak för antal rader (0 = alla) |
| `--calibrate` | Peka ut Bet.dag/Fakt.nr med musen; sparas fönsterrelativt i `visma_inbetalningar_coordinates.json` |
| `--self-test` | Interna tester av parser-/dialogfunktioner, ingen Visma-kontakt |
| `--inspect` | Listar dialogkontroller för felsökning (kräver öppet Visma) |
| `--debug-gui` | Sparar långsamma dialogdumpar (`visma_inspect_*.txt`) |

### Indata

- Obligatoriska kolumner: `Betalningsdatum, Fakturanr, Belopp` (alias accepteras, t.ex. `Datum`, `Fakturanummer`). Valfria: `Kundnamn`, `KundID` (endast logg).
- Interaktivt urval: startfakturanummer + antal (max 50). Hela urvalet valideras (datum/belopp/fakturanr) innan någon GUI-inmatning sker; belopp ≤ 0 avvisas.

### Utdata/loggning

- `visma_inbetalningar_logg_<YYYYMMDD_HHMMSS>.csv` (semikolon, UTF-8 med BOM) med kolumnerna `Rad;Betalningsdatum;VismaDatum;Kundnamn;KundID;Fakturanr;Belopp;Status;Felmeddelande;Tidpunkt`. Statusar: `OK`, `DRY_RUN`, `HOPPAD`, `FEL`, `STOPPAD`. Loggarna konsumeras av `visma_sammanstall_loggar.py`.

### Beroenden

- `pandas`, `openpyxl`; för skarp körning dessutom `pywinauto`, `pyautogui`, `pyperclip` (importeras "mjukt" – dry-run fungerar utan GUI-biblioteken). `xlrd` endast för gamla `.xls`.

### Felhantering

- Vid varje kritiskt fel (fält hittas inte, dialog uteblir, belopp kan inte verifieras, fel faktura i dialogen) **stoppas hela körningen** – skriptet fortsätter aldrig till nästa faktura efter ett fel.
- Tidsmätning per rad/steg skrivs ut (`SHOW_ROW_TIMING`).

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Flerstegsverifiering av belopp och differens, dubblettkontroll, urvalsvalidering före GUI-start, kalibreringsflöde med fönsterstorlekskontroll, dialogcache med regressionstestade buggfixar, omfattande self-test (`--self-test` körs utan Visma).

**Brister/risker med stöd i koden:**
- Säkerhetslägen (`CONFIRM_EACH_ROW`, `AMOUNT_TOLERANCE` m.fl.) ändras genom att redigera konstanter i filen – ingen konfigfil/flagga.
- `visma_config.json` i samma mapp används **inte** av detta skript (kalibreringen ligger i `visma_inbetalningar_coordinates.json`). Konfigfilen är en kvarleva – se roadmap.
- Fakturanummer-extraktion tillåter 4–8 siffror (bredare än CSV-stegens exakt 5) – inkonsekvent regelverk i repot.
- Etikettmatchningen (`Bet.dag`/`Fakt.nr`) är beroende av att Visma exponerar etiketterna; annars krävs kalibrering, vilket hanteras men måste göras om vid ändrad fönsterstorlek.

**Behöver verifieras:**
- Att dialogtitlarna `R[aä]tt belopp` och `Differens` matchar den Visma-version som används.
- Att `--self-test` fortfarande går igenom i aktuell Python-miljö.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| VR-1 | Ta bort eller koppla in `visma_config.json` | Oanvänd fil förvirrar (punkt #3 i tidigare genomgång) | Låg | Planerad | Bekräfta att inget annat läser filen |
| VR-2 | Flytta säkerhetslägen till flaggor/konfigfil | Undvika att redigera skriptet för lägesbyte | Låg | Planerad | – |
| VR-3 | Enhetlig fakturanummer-definition med CSV-stegen | 4–8 vs exakt 5 siffror | Medel | Behöver verifieras | Bekräfta nummerserien |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång; manualen i `visma_import\docs` bevarad som användarguide | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | ac33cbe m.fl. (tidigare fixar) |
