# Dokumentation – Ekonomiskript (C:\Ekonomi\Scripts)

Register över samtliga Python-skript i projektet, med länk till respektive skriptdokument. Varje dokument innehåller översikt, funktion/användning, nuläge, levande roadmap och ändringshistorik. Gemensam ändringslogg finns i [CHANGELOG.md](CHANGELOG.md).

## Flödesöversikt

```
BANKGIRO                          SWISH
1. clean_bankgiro_files           4. clean_swish_files
2. sammanstall_betalningar        5. samla_swish_betalningar
3. convert_betalningar_to_csv     6. step_4_skapa_swish_csv
                 \                   /
                  7. visma_konvertera_betalningar  (CSV -> *_visma.csv)
                  8. visma_register_inbetalningar_fixed  (inbetalningar i Visma)
                  9. visma_sammanstall_loggar  (månadslogg + avstämning)

Separata flöden:
  dashboard.py          – meny som orkestrerar steg 1–9
  payment_register.py   – webbregistrering i Hemfresh fakturasystem (Playwright)
  visma_import.py       – fakturaregistrering i Visma Kundbokning (ej inbetalningar)
```

## Skriptregister

| Skript | Dokument | Kort beskrivning |
|---|---|---|
| `dashboard.py` | [dashboard.md](dashboard.md) | Interaktiv meny som kör hela flödet i rätt ordning (ingen egen affärslogik) |
| `clean_bankgiro_files.py` | [clean_bankgiro_files.md](clean_bankgiro_files.md) | Rensar Bankgirots insättningsfiler till `clean_bet_lista_*.xlsx` (original orörda) |
| `sammanstall_betalningar.py` | [sammanstall_betalningar.md](sammanstall_betalningar.md) | Slår ihop rensade Bankgiro-filer till `samlade_betalningar_*.xlsx` |
| `convert_betalningar_to_csv.py` | [convert_betalningar_to_csv.md](convert_betalningar_to_csv.md) | Bankgiro-Excel → registrerings-CSV med fakturanummer-extraktion |
| `clean_swish_files.py` | [clean_swish_files.md](clean_swish_files.md) | Rensar Swish-exporter till `clean_swish_lista_*.xlsx` |
| `samla_swish_betalningar.py` | [samla_swish_betalningar.md](samla_swish_betalningar.md) | Slår ihop rensade Swish-filer till `samla_swish_bet_lista_*.xlsx` |
| `step_4_skapa_swish_csv.py` | [step_4_skapa_swish_csv.md](step_4_skapa_swish_csv.md) | Swish-Excel → registrerings-CSV (samma format som Bankgiro) |
| `visma_import\visma_konvertera_betalningar.py` | [visma_konvertera_betalningar.md](visma_konvertera_betalningar.md) | CSV → Visma-kolumnformat (`*_visma.csv`) |
| `visma_import\visma_register_inbetalningar_fixed.py` | [visma_register_inbetalningar_fixed.md](visma_register_inbetalningar_fixed.md) | Halvautomatisk inbetalningsregistrering i Visma Compact 6 (dry-run/`--live`) |
| `visma_import\visma_sammanstall_loggar.py` | [visma_sammanstall_loggar.md](visma_sammanstall_loggar.md) | Månadssammanställning av körloggar, dubblettrensning, avstämning |
| `visma_import\visma_import.py` | [visma_import.md](visma_import.md) | Fakturaregistrering i Visma Kundbokning via GUI-automation |
| `payment_register.py` | [payment_register.md](payment_register.md) | Betalningsregistrering i Hemfresh webbsystem via Playwright |

## Undantagna filer (med motivering)

- **`visma_import\archiv\*.py`** (10 filer, bl.a. `visma_import_v_2.py`–`v_6`, `visma_register_inbetalningar_fixed_v_2*.py`, `*_old.py`): arkiverade, ersatta versioner av de aktiva skripten. Mappnamnet och versionssuffixen visar att de bara sparats som historik – de dokumenteras inte separat och ska inte köras.
- **`__pycache__\`**: kompilerad cache, ingen källkod.

## Övrig befintlig dokumentation

- `visma_import\docs\Manual_Visma_inbetalningar.md` (+ PDF): detaljerad användarmanual för inbetalningsregistreringen – bevarad, skriptdokumentet länkar dit.
- `visma_import\docs\visma_sammanstall_loggar.md`: detaljerad referens för loggsammanställaren – bevarad.
- `docs\notion-uppdatering.md`: historisk anteckning från när `dashboard.py` byggdes (status på tidigare oklarheter).
- `docs\visma_import_tutorial.html/.pdf`, `docs\filtrera_swish_inbetalningar.pdf`, `docs\starta_workflow_dashboard.txt`: äldre guider, bevarade oförändrade.
- `requirements.txt`: samtliga Python-beroenden med kommentarer per skript.

Den tidigare `docs\README.md` innehöll användarguiden för `payment_register.py`; den är inarbetad i [payment_register.md](payment_register.md).

## Rutin för framtida uppdateringar

Projektet har inga andra agent-/utvecklarinstruktioner (ingen `CLAUDE.md`/`AGENTS.md` finns), så detta är den gällande rutinen. Uppdateringarna sker **inte** automatiskt – den som ändrar ett skript ansvarar för att:

1. **Uppdatera berörd skriptdokumentation** i `docs\` (funktion, indata/utdata, begränsningar).
2. **Uppdatera roadmapen** i samma dokument: sätt status (**Planerad / Pågår / Klar / Behöver verifieras**) och beskriv återstående arbete.
3. **Lägga en post i skriptets ändringshistorik** (datum, vad, varför, hur det verifierades, commitreferens när den finns).
4. **Lägga en post överst i [CHANGELOG.md](CHANGELOG.md)** (datum, berörda skript, kategori Tillagt/Ändrat/Åtgärdat/Borttaget, verifiering).

Markera sådant som inte kunnat testas som **Behöver verifieras** i stället för att anta att det fungerar.
