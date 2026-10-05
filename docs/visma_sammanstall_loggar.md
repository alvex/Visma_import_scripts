# visma_sammanstall_loggar.py

## Översikt

- **Skript:** `visma_import\visma_sammanstall_loggar.py`
- **Syfte:** Sammanställer körloggarna från inbetalningsregistreringen (`visma_inbetalningar_logg_*_*.csv`) till en månadsfil per betalningstyp, med dubblettmarkering, valfri dubblettrensning och valfri avstämning mot en export över betalda fakturor.
- **Användning:** Steg 9 i flödet. Körs via `dashboard.py` (menyval 9) eller fristående. Detaljerad dokumentation finns i `visma_import\docs\visma_sammanstall_loggar.md` – detta dokument sammanfattar och lägger till nuläge/roadmap utan att duplicera den.

## Funktion och användning

### Huvudflöde

1. Frågar interaktivt efter mapp, månad/år (t.ex. `2026-06`) och betalningstyp (`Bank`/`Skatteverket`).
2. Läser alla loggfiler som matchar `visma_inbetalningar_logg_*_*` i mappen (kodning `utf-8-sig`, reserv `cp1252`; trasiga/tomma filer hoppas över med varning). Kolumnalias tolereras.
3. Tar med poster med **status OK** och betalningsdatum inom vald månad (primärt `Betalningsdatum`, reserv `VismaDatum`). Sorterar stigande på datum och markerar dubbletter (samma fakturanr) i ny kolumn `Dubblett` (JA/NEJ).
4. Skriver `visma_inbetalningar_<typ>_<månad>_<år>_logg.csv` i samma mapp och visar statistik.
5. Valfritt: skapar `..._clean_logg.csv` utan dubbletter (första förekomsten behålls; varnar om dubbletter har olika belopp/datum).
6. Valfritt: jämför clean-filen mot en export över betalda fakturor (CSV eller `.xlsx`/`.xlsm`) och skriver `..._avvikelser.csv` med avvikelsetyperna `SAKNAS_I_VISMA`, `SAKNAS_I_BETALDA_FAKTUROR`, `BELOPPSAVVIKELSE`, `DATUMAVVIKELSE`.

### Startkommando

```powershell
python visma_import\visma_sammanstall_loggar.py
```

### Beroenden

- Endast standardbiblioteket; `openpyxl` krävs bara om jämförelsefilen är Excel (lazy import med tydligt felmeddelande).

### Felhantering

- En trasig loggfil stoppar aldrig körningen – den hoppas över med varning.
- Utdatafilen utesluts ur inläsningen om den råkar matcha filmönstret.
- `.xls` som jämförelsefil avvisas med instruktion att spara om.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Robust inläsning med alias/kodningsreserver, dubblettmarkering och -rensning, trevägsavstämning mot betald-export, statistik och varningslistor.

**Brister/risker med stöd i koden:**
- Endast poster med status `OK` tas med – `DRY_RUN`-poster filtreras bort, så en månadssammanställning som körs efter enbart torrkörningar blir tom (korrekt, men kan förvåna).
- Betalningstypen (`Bank`/`Skatteverket`) påverkar bara **filnamnet** – ingen filtrering av poster per typ görs; ligger båda typernas loggar i samma mapp blandas de.
- Helt interaktiv – kan inte skriptas/parametriseras (se även dashboard-roadmap DB-2).

**Behöver verifieras:**
- Hur Skatteverkets loggfiler särskiljs från bankens i praktiken (mappstruktur eller filnamn?).

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| VL-1 | Filtrera poster per betalningstyp (eller dokumentera att typerna ska hållas i separata mappar) | Typen påverkar i dag bara filnamnet | Medel | Behöver verifieras | Klargöra hur loggarna lagras per typ |
| VL-2 | Kommandoradsargument för mapp/period/typ | Möjliggör automatisering via dashboard | Låg | Planerad | Se DB-2 |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång; detaljdokumentet i `visma_import\docs` bevarat | Dokumentationsuppdrag | Koden läst i sin helhet; skriptet har inte körts | – |
