# Vad som gjordes

**Nytt: `dashboard.py` (CLI-meny, i projektroten)**

- Interaktiv meny exakt enligt din valda layout, med arbetsmapp som sätts en gång; rensade/sammanställda filer och CSV hamnar automatiskt i `<arbetsmapp>\edit`.
- Anropar de befintliga skripten via `--input`/`--output` (ingen egen affärslogik → skripten förblir enda sanningskälla).
- `A`/`B` kör hela Bankgiro-/Swish-kedjan i ett svep; steg 8 hittar automatiskt senaste `*_visma.csv` och frågar torr/skarp; steg 7 matar in mapparna åt konverteraren.
- Verifierat: kompilerar, menyn startar/avslutar rent, och ett end-to-end-test bekräftade att rätt skript anropas med rätt argument (`--input <mapp> --output <mapp>\edit`).

**Kommentarerna uppdaterade** i `clean_swish_files.py`, `samla_swish_betalningar.py`, `step_4_skapa_swish_csv.py` (pekar nu på `dashboard.py`) och `convert_betalningar_to_csv.py` (den döda `config.json`-referensen borttagen).

**Guiden uppdaterad** (`workflow-guide.html` + `.pdf`, 18 sidor):

- Ny "Genväg"-ruta i huvudflödet som förklarar `dashboard.py`.
- `dashboard.py` tillagd i referenstabellen.
- "Lösa trådar (a)" omskriven till åtgärdat; ändringsloggens #2 → Åtgärdat.

## Status på OKLART

| # | Punkt | Status |
|---|-------|--------|
| 1 | `archiv/` som arkiv | ✅ Bekräftat |
| 2 | dashboard/`config.json` | ✅ Åtgärdat (dashboard byggd, kommentarer fixade) |
| 3 | `visma_config.json` (oanvänd) | 📄 Dokumenterat – väntar på ta bort/koppla in |
| 4 | `total_betalningar_*.xlsx` (extern fil) | 📄 Dokumenterat – väntar på ev. nytt steg |
| 5 | `requirements.txt` | ✅ Åtgärdat |

**Kvar när du vill:** ta bort `visma_config.json` (#3), bygga ett steg som skapar `total_betalningar_*.xlsx` åt webbflödet (#4), eller radera de nu ersatta gamla doc-filerna. Säg till.
