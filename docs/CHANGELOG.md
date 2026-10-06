# CHANGELOG

Gemensam ändringslogg för skripten i `C:\Ekonomi\Scripts`. Senaste ändringen överst.

Varje post anger datum, berörda skript, kategori (**Tillagt / Ändrat / Åtgärdat / Borttaget**), kort beskrivning och hur ändringen verifierades. Historik före 2026-10-06 har inte rekonstruerats i efterhand – se git-historiken (`git log`) för tidigare commits.

---

## 2026-10-07 – Nytt skript: månadsvis kontrollista över förberedda betalningsposter

- **Berörda skript:** `sammanstall_fakturor.py` (nytt); inga befintliga skript ändrade.
- **Tillagt:** Sammanställer registrerings-CSV:erna (`betalningar_lista_to_reg_*`, `swish_betalningar_for_registrering_*`) i `WORK\output` till kontrollistan `output\fakturor_lista\fakturor_<mappnamn>.xlsx` (mappnamn = månadsmappen, output-mappens förälder). Perioden är kalendermånaden i betalningsdatum (`--month ÅÅÅÅ-MM` vid behov). Återkommande fakturanummer markeras `Dubblett` på samtliga berörda rader (inom och mellan källfiler, retroaktivt); inga poster tas bort eller slås ihop. Omkörningsskydd via SHA-256 per källfil (ark `Källfiler`): oförändrade filer hoppas över, ändrade rapporteras med returkod 2 utan att läsas om. Felrader (saknat fakturanummer, ogiltigt datum/belopp) hamnar i arket `Felrader` med källfil + radnummer. Atomisk skrivning via temporär fil; låst målfil ger begripligt fel med bevarade data. Även `--dry-run`. Dokumentation i `docs\sammanstall_fakturor.md` + rad i registret.
- **Verifiering:** `py_compile`; fullständig scenariokörning mot syntetisk testdata (första körning, omkörning utan radändring, komplettering med ny fil, dubblett Bankgiro↔Swish, inledande nollor, svenska/negativa belopp, månadsskifte, felrader, ändrad källfil, låst målfil vid läsning och skrivning, torrkörning) samt skarp körning mot `Fak_2026\sept 01` (11 poster, totalsumma 19 232,00, inga dubbletter). Inga källfiler eller befintliga skript berördes.
- **Commit:** – (läggs till vid commit)

## 2026-10-06 – Nytt filflöde: processed/output ersätter edit

- **Berörda skript:** `clean_bankgiro_files.py`, `clean_swish_files.py`, `sammanstall_betalningar.py`, `samla_swish_betalningar.py`, `convert_betalningar_to_csv.py`, `step_4_skapa_swish_csv.py`, `dashboard.py`.
- **Ändrat:** Bearbetade mellanresultat sparas nu i `WORK\processed`, färdiga CSV-filer i `WORK\output` (syskonmappar på samma nivå). Rensningsskripten läser oförändrat från arbetsmappen och sparar i `processed`; sammanställningsskripten läser/sparar i `processed`; CSV-skripten läser från `processed` och sparar i `output`. Dashboarden skapar båda mapparna automatiskt och kör Visma-stegen 7–9 mot `output`. Standardvärden uppdaterade: rensningsskriptens standardutdata är `processed` (tidigare `edit`); CSV-skripten utan `--output` sparar i syskonmappen `output` när indatafilen ligger i `processed` (annars som förut bredvid indatafilen). Betalningslogik, filnamn och filformat oförändrade; inga befintliga filer flyttade eller raderade.
- **Verifiering:** Syntaxkontroll (`py_compile`) av alla sju filer. Hela kedjan steg 1–6 körd mot syntetiska Bankgiro-/Swish-testfiler i en temporär mapp **utan** förskapade `processed`/`output` – mellanresultaten hamnade i `processed`, CSV + loggar i `output`, och CSV-innehållet kontrollerades (datum, fakturanummer, belopp). Standardlogiken utan `--output` verifierades separat. Ingen produktionsdata berördes. Dashboard-menyn har inte körts interaktivt (**Behöver verifieras** vid nästa skarpa månadskörning).
- **Commit:** – (läggs till vid commit)

## 2026-10-06 – Dokumentation skapad

- **Berörda skript:** samtliga 12 aktiva (se `docs\README.md`); inga skript ändrade.
- **Tillagt:** Ett dokument per skript i `docs\` med översikt, användning, nuläge, levande roadmap och ändringshistorik. Nytt register i `docs\README.md` med rutin för framtida uppdateringar. Denna changelog.
- **Ändrat:** Tidigare `docs\README.md` (användarguide för `payment_register.py`) inarbetad i `docs\payment_register.md`; README är nu dokumentationsregistret.
- **Verifiering:** All dokumentation är skriven utifrån fullständig kodläsning av de aktiva skripten samt befintliga manualer. Inga skript har körts och ingen ekonomidata har ändrats. Punkter som inte kunnat fastställas från koden är markerade **Behöver verifieras** i respektive dokument.
- **Commit:** – (läggs till vid commit)
