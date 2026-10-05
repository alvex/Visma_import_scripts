# dashboard.py

## Översikt

- **Skript:** `dashboard.py` (projektroten)
- **Syfte:** Interaktiv terminalmeny som orkestrerar hela betalnings- och Visma-flödet. Anropar de befintliga skripten i rätt ordning via `--input`/`--output` – innehåller **ingen egen affärslogik** (skripten förblir enda sanningskälla).
- **Användning:** Startpunkten för det normala månadsflödet.

## Funktion och användning

### Mappkonvention

- **Arbetsmapp (WORK):** mappen med råfilerna (Bankgiro/Swish). Väljs vid start, kan bytas med `M`.
- **`WORK\processed`:** bearbetade mellanresultat (rensade och sammanställda Excel-filer).
- **`WORK\output`:** färdiga CSV-filer (registrerings-CSV, `*_visma.csv`, körloggar).

Båda mapparna ligger på samma nivå och skapas automatiskt av dashboarden om de saknas.

### Meny

| Val | Åtgärd | Skript som anropas |
|---|---|---|
| 1 | Rensa Bankgiro-filer | `clean_bankgiro_files.py --input WORK --output WORK\processed` |
| 2 | Sammanställ Bankgiro | `sammanstall_betalningar.py --input processed --output processed` |
| 3 | Skapa CSV (Bankgiro) | `convert_betalningar_to_csv.py --input processed --output output` |
| 4 | Rensa Swish-filer | `clean_swish_files.py --input WORK --output WORK\processed` |
| 5 | Sammanställ Swish | `samla_swish_betalningar.py --input processed --output processed` |
| 6 | Skapa CSV (Swish) | `step_4_skapa_swish_csv.py --input processed --output output` |
| 7 | Konvertera till Visma-format | `visma_import\visma_konvertera_betalningar.py` (output matas in via stdin som både in- och utdatamapp) |
| 8 | Registrera i Visma (torr/skarp) | `visma_import\visma_register_inbetalningar_fixed.py <senaste *_visma.csv>` (+ `--live` vid skarp) |
| 9 | Sammanställ & stäm av loggar | `visma_import\visma_sammanstall_loggar.py` |
| A | Hela Bankgiro-kedjan | 1 → 2 → 3 |
| B | Hela Swish-kedjan | 4 → 5 → 6 |
| M | Byt arbetsmapp | – |
| 0 | Avsluta | – |

### Startkommando

```powershell
python dashboard.py
```

### Beroenden

- Endast Python-standardbibliotek (subprocess). Underskripten har egna beroenden.

### Felhantering

- Returkoden från varje underskript skrivs ut. Ctrl+C i ett steg återgår till menyn.
- Steg 8 vägrar starta om ingen `*_visma.csv` finns i output-mappen och hänvisar till steg 7.
- Skarp registrering (steg 8) kräver dessutom bekräftelsen `ja` inne i registreringsskriptet.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Hela menystrukturen, kedjekörning A/B, automatisk stdin-matning till den interaktiva Visma-konverteraren, val av senaste `*_visma.csv` för steg 8. Enligt tidigare genomgång (`notion-uppdatering.md`) är menyn end-to-end-testad: rätt skript anropas med rätt argument.

**Brister/risker med stöd i koden:**
- Kedjorna A/B fortsätter till nästa steg även om föregående steg returnerade fel (returkoden kontrolleras inte).
- `visma_import.py` (fakturaregistrering i Kundbokning) finns inte med i menyn – det flödet körs helt separat.
- Steg 9 kör loggsammanställaren, som själv frågar efter mapp – dashboarden skickar inte in output-mappen utan visar den bara som tips.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| DB-1 | Avbryta kedja A/B vid returkod ≠ 0 | Undviker att sammanställa/konvertera efter ett misslyckat steg | Hög | Planerad | – |
| DB-2 | Mata in output-mappen till steg 9 (loggsammanställaren) automatiskt | Mindre manuell inmatning | Låg | Planerad | Kräver att `visma_sammanstall_loggar.py` tar emot argument eller stdin |
| DB-3 | Ev. lägga till `visma_import.py` som menyval | Samla hela flödet på ett ställe | Låg | Planerad | Beslut om fakturaflödet ska ingå |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång | Dokumentationsuppdrag | Koden läst; tidigare end-to-end-test dokumenterat i `notion-uppdatering.md` | 69d3b36 (dashboard tillkom i V.2.0) |
| 2026-10-06 | Nytt filflöde: `edit` ersatt av `processed` (mellanresultat) + `output` (färdiga CSV). Steg 1–2, 4–5 → processed; steg 3, 6 → output; steg 7–9 arbetar mot output. Mapparna skapas automatiskt. | Tydligare separation mellan mellanresultat och färdiga filer | Syntaxkontroll + hela kedjan steg 1–6 körd mot syntetiska testfiler i temporär mapp (utan förskapade mappar); dashboard-menyn själv ej interaktivt testad – anropsargumenten följer samma mönster som verifierats per skript | – |
