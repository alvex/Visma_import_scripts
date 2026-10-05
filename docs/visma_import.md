# visma_import.py

## Översikt

- **Skript:** `visma_import\visma_import.py`
- **Syfte:** Automatisk registrering av **kundfakturor** (inte inbetalningar) i Visma Compact via GUI-automation (pyautogui + Windows API). Läser fakturor från Excel och matar in KundID, bokföringsdag och belopp i fönstret Kundreskontra → Kundbokning, sparar med Num+.
- **Användning:** Körs manuellt och fristående (ingår inte i `dashboard.py`-menyn). Kräver att Visma Compact är öppet med Kundbokning synligt.

## Funktion och användning

### Huvudflöde per faktura

1. Klickar i KundID-fältet (position relativ Visma-fönstret), skriver KundID och trycker Tab så kunden laddas (0,8 s paus).
2. Klickar direkt i `Bokf.dag`, ersätter innehållet med fakturadatum i formatet `YY-MM-DD`.
3. Klickar direkt i `Fak.belopp`, skriver beloppet med svensk decimal (`1949,00`).
4. Trycker exakt 5 × Tab, pausar 1 s och skickar riktig Num+ (Windows API `keybd_event`) som sparar fakturan i Visma.

Fältmarkering görs med **End → Shift+Home** (aldrig Ctrl+A – det öppnar Artikelregistret i Visma Compact).

### Start och interaktivt urval

```powershell
python visma_import\visma_import.py
```

1. Frågar efter Excel-filens sökväg (`.xlsx`-ändelsen kan utelämnas).
2. Validerar filen: obligatoriska kolumner `KundID, Fakturanr, Fakturadatum, Belopp` (alias accepteras, t.ex. `Personnummer`, `Fakturanummer`, `Fakturabelopp`). Tomma, ogiltiga eller dubbla fakturanummer stoppar körningen innan Visma påverkas.
3. Frågar antal fakturor (1–50, aldrig fler än raderna) och startfakturanummer. Den valda serien måste vara **strikt löpande (+1)** – hopp, dubbletter eller fel ordning stoppar körningen med detaljerat felmeddelande.
4. Letar upp Visma-fönstret (prioriterar titlar med "Kundbokning"/"Kundreskontra"/"Compact"), visar nedräkning 3-2-1 och börjar registrera.

### Säkerhet och stopp

- **Nödstopp:** flytta muspekaren till skärmens övre vänstra hörn (pyautogui failsafe).
- Vid varje inmatning kontrolleras att Visma fortfarande har fokus; annars stoppas importen.
- Vid fel skickas ingen Num+ – fakturan sparas inte halvfärdig.
- Logg sparas som `visma_import_logg_<tidsstämpel>.xlsx` bredvid Excel-filen (kolumner: Rad, KundID, Excel/Visma Fakturanr, Nummerkontroll, Status, Fel).

### Positioner

Fältens positioner är hårdkodade: X som andel av fönsterbredden (`KUNDID_X_RATIO = 0.036` m.fl.), Y i **pixlar** från fönstrets överkant (99/123/124).

### Beroenden

- `pandas`, `openpyxl` (via pandas), `pyautogui`; Windows API via `ctypes` (endast Windows).

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Komplett valideringskedja före GUI-inmatning (kolumner, tomma/ogiltiga/dubbla fakturanummer, strikt löpande serie), fokuskontroll, failsafe, loggning, direktklick i stället för Tab-navigation.

**Brister/risker med stöd i koden:**
- **Ingen kontroll mot Vismas faktiska "Nr."-fält** – avläsningen togs bort medvetet (Ctrl+C-simulering orsakade KeyboardInterrupt). Att Vismas nummerserie matchar Excel garanteras bara indirekt via startnummer + strikt löpande serie. `Nummerkontroll` i loggen blir alltid `EJ LÄST`.
- Y-positionerna anges i fasta pixlar – fel skärmupplösning/DPI-skalning eller ändrad fönsterhöjd ger klick i fel fält. Ingen kalibreringsfunktion finns (till skillnad från inbetalningsskriptet).
- Konstanten `TEST_LIMIT` (rad 69) används inte längre någonstans – död konfiguration som kan vilseleda.
- Inget dry-run-läge finns; minsta körning är 1 skarp faktura.

**Behöver verifieras:**
- Att "exakt 5 Tab efter Fak.belopp" fortfarande stämmer med aktuell Visma-layout.
- Att positionsandelarna stämmer på den skärm som används i produktion.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| VI-1 | Dry-run-läge (simulera utan Num+) | Övriga skript i repot har dry-run; detta saknar det | Hög | Planerad | – |
| VI-2 | Kalibreringsläge som i `visma_register_inbetalningar_fixed.py` (--calibrate) | Fasta pixel-Y är skört mot upplösning/DPI | Medel | Planerad | Testas på produktionsskärmen |
| VI-3 | Ta bort oanvänd `TEST_LIMIT` | Död kod/konfiguration | Låg | Planerad | – |
| VI-4 | Läsa tillbaka Vismas "Nr." på säkert sätt (utan Ctrl+C-simulering) | Återinföra faktisk nummerkontroll | Medel | Planerad | Kräver pywinauto-läsning av fältet; Behöver verifieras mot Visma |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
