# payment_register.py

## Översikt

- **Skript:** `payment_register.py` (projektroten)
- **Syfte:** Registrerar kundbetalningar automatiskt i Hemfresh webbaserade fakturasystem (`hemfresh.com/fakturorab`) utifrån den senaste Excel-filen `total_betalningar_*.xlsx`. Löser det manuella arbetet att öppna varje faktura och registrera betalningen.
- **Användning:** Körs manuellt av ekonomiansvarig när en sammanställd betalningslista finns. Fristående från dashboard-flödet (anropas inte av `dashboard.py`).

## Funktion och användning

### Huvudflöde

1. Frågar efter mappen där `total_betalningar_*.xlsx` finns och väljer den senast ändrade filen (temporära `~$`-filer ignoreras).
2. Frågar efter körläge: `1` = dry-run (standard) eller `2` = riktig registrering. Riktig registrering kräver bekräftelsen `REGISTRERA` – annars faller skriptet tillbaka till dry-run.
3. Frågar hur många rader som ska behandlas (`1` är standard, `ALLA` för samtliga).
4. Läser Excel-filen: hittar rubrikraden dynamiskt (kolumnerna `Datum`, `Avsändare`, `Betalningsreferens`, `Belopp` måste finnas), hoppar över tomma rader och `Summa`-rader, extraherar fakturanummer (4–7 siffror) ur betalningsreferensen.
5. Öppnar Chromium via Playwright (synligt läge). Första gången loggar användaren in manuellt; sessionen sparas i `auth_state.json` och återanvänds.
6. Per rad: öppnar betalningssidan för fakturan, kontrollerar fakturanummer på sidan, läser `Att betala:`, jämför beloppet exakt (två decimaler), kontrollerar att fakturan inte redan är betald och att formuläret finns. Fyller i belopp, datum, ev. betalningsmetod (`Bankbetalning`) och anteckning.
7. I dry-run stannar skriptet före knappen `Betalningsprocess` och loggar `DRY_RUN_OK`. Vid riktig registrering klickas knappen och resultatet verifieras mot sidtexten (`REGISTRERAD` eller `OKLAR_STATUS`).
8. Skriver loggfil `payment_log_YYYY-MM-DD_HHMMSS.csv` (semikolon, UTF-8 med BOM) i samma mapp som Excel-filen. Vid fel sparas skärmdump i `screenshots/`.

### Indata, utdata och filformat

| Typ | Fil/format |
|---|---|
| Indata | `total_betalningar_*.xlsx` – kolumner `Datum`, `Avsändare`, `Betalningsreferens`, `Belopp` (rubrikraden hittas automatiskt) |
| Utdata | `payment_log_*.csv` med kolumnerna `Datum;Avsändare;Betalningsreferens;Fakturanummer;Belopp_fran_Excel;Belopp_pa_fakturasidan;Status;Meddelande;Tidpunkt;Screenshot` |
| Session | `auth_state.json` (cookies – **känslig fil**, får inte spridas eller committas) |
| Fel | `screenshots/<fakturanr>_<STATUS>_<tidpunkt>.png` |

### Startkommandon

```powershell
python payment_register.py
```

Skriptet är helt interaktivt – inga kommandoradsflaggor finns.

### Beroenden och förberedelser

- `openpyxl`, `playwright` (se `requirements.txt`) samt `python -m playwright install chromium`.
- Manuell inloggning i fakturasystemet vid första körningen.
- Ingen fil `total_betalningar_*.xlsx` skapas av övriga skript i repot – den tas fram utanför detta flöde (se roadmap).

### Statusar i loggen

`DRY_RUN_OK`, `REGISTRERAD`, `BELOPP_MATCHAR_INTE`, `REDAN_BETALD`, `SAKNAR_FAKTURANUMMER`, `SAKNAR_DATUM`, `SAKNAR_FORMULAR`, `SAKNAR_BELOPP`, `FEL_FAKTURA`, `TIMEOUT`, `FEL`, `OKLAR_STATUS`.

### Felhantering och begränsningar

- Registrerar aldrig om Excel-beloppet inte matchar `Att betala:` exakt.
- Ett fel på en rad stoppar inte batchen – raden loggas och nästa behandlas.
- Slumpad paus 1–2 s mellan rader.
- Hellre att en rad hoppas över och loggas än att fel betalning registreras.

### Säker teststrategi

1. Dry-run med 1 rad → kontrollera webbläsare + logg.
2. Dry-run med `ALLA` → kontrollera loggen.
3. Riktig registrering, börja med 1 rad.

## Nuläge och kända problem

**Implementerat och verifierat i koden:**
- Komplett dry-run-läge, dubbelbekräftelse för skarp körning, radbegränsning, exakt beloppsjämförelse, logg + skärmdumpar, sessionsåteranvändning.

**Brister/risker med stöd i koden:**
- `extract_invoice_number` tillåter 4–7 siffror och förkastar i reservläget kandidater som börjar på `19`/`20` (årtalsskydd). Giltiga fakturanummer som börjar på 19/20 kan därmed ratas när referensen saknar nyckelord som "faktura". Swish-/Bankgiro-CSV-flödet använder i stället exakt 5 siffror – regelverken är inte samordnade.
- `verify_payment_result` godkänner breda textmarkörer (`"betald"`, `"betalningen"`); en felsida som innehåller ordet kan ge falsk `REGISTRERAD`.
- `BASE_URL`/`PAYMENT_URL_TEMPLATE` är hårdkodade; ändras systemets URL slutar skriptet fungera.
- `auth_state.json` ligger i projektroten och innehåller inloggningscookies.

**Behöver verifieras:**
- Att CSS-selektorerna (`amount_selectors` m.fl.) fortfarande matchar fakturasidans aktuella HTML.
- Att markörerna för "redan betald" täcker systemets alla formuleringar.

## Levande roadmap

| ID | Förbättring/åtgärd | Motivering | Prioritet | Status | Beroenden/verifieringskrav |
|---|---|---|---|---|---|
| PR-1 | Bygga ett steg som skapar `total_betalningar_*.xlsx` ur befintligt flöde | Indatafilen skapas i dag utanför repot (punkt #4 i tidigare genomgång) | Medel | Planerad | Beslut om källa (Bankgiro-/Swish-sammanställning) |
| PR-2 | Samordna fakturanummer-regler (4–7 vs exakt 5 siffror) mellan webb- och CSV-flödet | Olika regler kan ge olika fakturanummer för samma referens | Medel | Planerad | Bekräfta faktisk nummerserie |
| PR-3 | Striktare resultatverifiering efter `Betalningsprocess` (t.ex. kräva fakturanumret i bekräftelsetexten) | Minskar risk för falsk `REGISTRERAD` | Medel | Planerad | Exempel på systemets bekräftelsesida |
| PR-4 | Flytta `auth_state.json` till skyddad plats och/eller lägga till i `.gitignore` | Känslig sessionsdata | Hög | Planerad | Kontrollera att filen inte är spårad i git |

## Ändringshistorik

| Datum | Ändring | Varför | Verifiering | Commit |
|---|---|---|---|---|
| 2026-10-06 | Dokumentet skapat utifrån kodgenomgång; innehåll från tidigare `docs/README.md` (användarguiden) inarbetat | Dokumentationsuppdrag – ett dokument per skript | Koden läst i sin helhet; skriptet har inte körts | – |
