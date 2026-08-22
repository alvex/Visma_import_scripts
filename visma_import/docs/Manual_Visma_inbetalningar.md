# Manual: Registrera kundinbetalningar i Visma Compact 6

**Script:** `visma_register_inbetalningar_fixed.py`  
**För:** HemFresh  
**Version:** 1.0 - 20 augusti 2026

---

## 1. Vad gör scriptet?

Scriptet läser betalningar från en Excel- eller CSV-fil och registrerar dem i:

**Visma Compact 6 → Kundreskontra → Inbetalningar**

För varje vald faktura gör scriptet följande:

1. Skriver betalningsdatum i **Bet.dag**.
2. Skriver fakturanumret i **Fakt.nr**.
3. Öppnar dialogen **Rätt belopp?**.
4. Skriver Excel-beloppet i **Ändra till:**.
5. Kontrollerar beloppet exakt på öresnivå.
6. Om dialogen **Differens** visas kontrolleras differensen och valet **Restbelopp på fakturan**.
7. Loggar resultatet för fakturan.

> **VIKTIGT:** Scriptet klickar aldrig automatiskt på den slutliga knappen för Bankgiro/Bokför. Du måste själv kontrollera listan **Till betalning >>>** och slutföra bokföringen i Visma.

---

## 2. Säkerhetsregler

Följ alltid dessa regler:

- Ta en säkerhetskopia av Visma innan en skarp körning.
- Börja alltid med en torrkörning utan `--live`.
- Testa sedan med **en faktura** i skarpt läge.
- Använd inte mus eller tangentbord medan scriptet registrerar.
- Kontrollera loggen och Vismas lista innan slutlig bokföring.
- Kör högst 50 fakturor åt gången.
- Om scriptet stoppar: rätta felet innan du startar om.

Scriptet fortsätter inte till nästa faktura om ett kritiskt fel upptäcks.

---

## 3. Krav

### Program

- Windows
- Visma Compact 6
- Python 3
- PowerShell eller Kommandotolken

### Python-paket

Öppna PowerShell och installera paketen:

```powershell
python -m pip install pandas openpyxl pyautogui pywinauto pyperclip
```

För gamla Excel-filer i formatet `.xls` kan även detta behövas:

```powershell
python -m pip install xlrd
```

---

## 4. Förbered Excel-filen

### Obligatoriska kolumner

| Kolumn | Exempel | Förklaring |
|---|---:|---|
| `Betalningsdatum` | `2026-04-07` | Konverteras automatiskt till `26-04-07` i Visma. |
| `Fakturanr` | `49561` | Måste motsvara fakturanumret i Visma. |
| `Belopp` | `1465,00` | Det belopp som faktiskt har betalats. |

### Valfria kolumner

| Kolumn | Exempel |
|---|---|
| `Kundnamn` | `Peter Zäll` |
| `KundID` | `650505-1059` |

### Exempel

| Betalningsdatum | Fakturanr | Belopp | Kundnamn | KundID |
|---|---:|---:|---|---|
| 2026-04-07 | 49561 | 1465,00 | Peter Zäll | 650505-1059 |

Scriptet stoppar innan Visma ändras om en vald rad har:

- tomt eller ogiltigt datum,
- tomt eller ogiltigt fakturanummer,
- tomt, ogiltigt eller negativt belopp,
- dubblerat fakturanummer.

Filen kan vara `.xlsx`, `.xlsm`, `.xls` eller `.csv`.

---

## 5. Lägg filerna i samma mapp

Exempel:

```text
C:\ekonomi\scripts\
├── visma_register_inbetalningar_fixed.py
├── betalningar.xlsx
└── visma_inbetalningar_coordinates.json
```

Den sista filen skapas automatiskt när du kalibrerar scriptet.

Öppna PowerShell i mappen:

```powershell
cd C:\ekonomi\scripts
```

Om mappnamnet innehåller mellanslag använder du citattecken:

```powershell
cd "C:\ekonomi\mina scripts"
```

---

## 6. Kontrollera att scriptet fungerar

Kör de inbyggda testerna:

```powershell
python visma_register_inbetalningar_fixed.py --self-test
```

Korrekt resultat avslutas med:

```text
Alla self-tester gick igenom.
```

Detta test kontaktar inte Visma och registrerar ingenting.

---

## 7. Kalibrera fälten första gången

Kalibrering behövs eftersom Visma Compact ibland visar fälten visuellt utan att Windows kan identifiera dem tekniskt.

Kalibrera första gången, samt när:

- Visma-fönstrets storlek har ändrats,
- fönsterlayouten har ändrats,
- klicken hamnar i fel fält,
- meddelandet **Hittade inte Bet.dag-fältet** visas.

### Gör så här

1. Öppna Visma Compact 6.
2. Gå till **Kundreskontra → Inbetalningar**.
3. Kontrollera att registreringsfönstret är synligt.
4. Kör:

```powershell
python visma_register_inbetalningar_fixed.py --calibrate
```

5. Tryck `Enter` när terminalen ber om **Bet.dag**.
6. Du får fem sekunder. Flytta muspekaren till mitten av fältet **Bet.dag** och håll den stilla.
7. Gå tillbaka till terminalen med `Alt+Tab`.
8. Upprepa samma procedur för **Fakt.nr**.

Kalibreringen sparas i:

```text
visma_inbetalningar_coordinates.json
```

> **TIPS:** Behåll samma storlek på Visma-fönstret vid kalibrering och registrering.

---

## 8. Gör alltid en torrkörning först

En torrkörning läser och kontrollerar filen, men skriver ingenting i Visma.

```powershell
python visma_register_inbetalningar_fixed.py "C:\ekonomi\scripts\betalningar.xlsx"
```

Scriptet frågar:

1. **Från vilket fakturanummer vill du börja?**
2. **Hur många fakturor vill du registrera?**

Exempel:

```text
Från vilket fakturanr vill du börja? 49561
Hur många fakturor vill du registrera? 1
```

Kontrollera listan med valda fakturor. I torrkörningen visas statusen:

```text
DRY_RUN: inget skrevs i Visma.
```

---

## 9. Förbered Visma för skarp körning

Innan du kör med `--live`:

1. Ta en säkerhetskopia.
2. Öppna **Kundreskontra → Inbetalningar**.
3. Välj **Obetalda**.
4. Kontrollera att konto är **1930**.
5. Kontrollera att fälten **Bet.dag** och **Fakt.nr** syns.
6. Låt Visma-fönstret vara öppet och synligt.
7. Stäng andra dialogrutor i Visma.

---

## 10. Kör skarpt

Kommando:

```powershell
python visma_register_inbetalningar_fixed.py "C:\ekonomi\scripts\betalningar.xlsx" --live
```

### Bekräfta körningen

Scriptet visar en tydlig varning. Skriv exakt:

```text
JA
```

Välj därefter:

- startfaktura,
- antal fakturor, högst 50.

Kontrollera förhandsvisningen och tryck `Enter` för att starta.

> **Rör inte mus eller tangentbord under registreringen.** Scriptet behöver kontroll över fokus och tangenttryckningar.

---

## 11. Vad händer för varje faktura?

### Steg 1 - Bet.dag och Fakt.nr

Scriptet skriver:

- `Betalningsdatum` i **Bet.dag**, exempelvis `26-04-07`,
- `Fakturanr` i **Fakt.nr**, exempelvis `49561`.

Därefter trycker scriptet `Enter` för att öppna fakturan.

### Steg 2 - Rätt belopp?

Dialogen **Rätt belopp?** visar fakturans ursprungliga belopp.

Scriptet:

1. skriver Excel-beloppet i **Ändra till:**,
2. lämnar fältet så att Visma formaterar beloppet,
3. läser tillbaka värdet,
4. kräver exakt samma belopp på öresnivå.

Exempel:

```text
Excel: 1465,00
Godkänt i Visma: 1465,00
Inte godkänt: 1465,01
```

Om ens ett öre avviker stoppas scriptet utan att bekräfta dialogen.

### Steg 3 - Differens

Om fakturan är på `2881,00` och betalningen är `1465,00` ska Visma visa:

```text
Skillnad: -1 416,00
```

Scriptet kontrollerar att:

- differensen är exakt korrekt,
- valet är **Restbelopp på fakturan**.

Först därefter bekräftas dialogen.

### Steg 4 - Nästa faktura

Scriptet väntar tills dialogerna har stängts. Därefter behandlas nästa vald faktura.

---

## 12. Kontroll efter körningen

När scriptet är klart:

1. Läs sammanfattningen i terminalen.
2. Öppna loggfilen bredvid Excel-filen.
3. Kontrollera Vismas lista **Till betalning >>>**.
4. Jämför fakturanummer och belopp mot Excel.
5. Klicka själv på Bankgiro/Bokför först när allt stämmer.

Loggfilens namn ser ut så här:

```text
visma_inbetalningar_logg_20260820_153000.csv
```

Vanliga statusvärden:

| Status | Betydelse |
|---|---|
| `OK` | Fakturan registrerades i arbetslistan. |
| `DRY_RUN` | Fakturan simulerades, men inget skrevs i Visma. |
| `FEL` | Ett fel upptäcktes och körningen stoppades. |
| `STOPPAD` | Körningen avbröts. |

---

## 13. Felsökning

| Meddelande eller problem | Orsak | Åtgärd |
|---|---|---|
| `Hittade inte Bet.dag-fältet` | Visma exponerar inte fältet tekniskt. | Kör `python visma_register_inbetalningar_fixed.py --calibrate`. |
| Klicket hamnar i fel fält | Visma-fönstrets storlek eller layout har ändrats. | Kör kalibreringen igen med samma fönsterstorlek som ska användas. |
| Visma-fönstret hittas inte | Visma är stängt eller fönstertiteln matchar inte. | Öppna Visma Compact 6 och gå till Inbetalningar. |
| `Rätt belopp?` öppnas inte | Fakturanumret saknas bland obetalda fakturor eller fel vy är vald. | Kontrollera fakturanumret och välj **Obetalda**. |
| Beloppet kan inte verifieras | Visma visar ett annat värde, exempelvis `1465,01`. | Bekräfta inte manuellt. Kontrollera beloppet och kör endast en faktura igen. |
| Fel differens | Differensen motsvarar inte fakturabelopp minus Excel-belopp. | Kontrollera fakturan, Excel-raden och Vismas belopp innan nytt försök. |
| `DRY_RUN` visas | Scriptet körs i testläge. | Lägg till `--live` endast när testet är korrekt. |
| Python-paket saknas | Beroenden är inte installerade. | Kör installationskommandot i avsnitt 3. |
| Dubblerat fakturanummer | Samma fakturanummer förekommer flera gånger i urvalet. | Rätta Excel-filen innan körning. |

### Om körningen stoppas mitt i en faktura

1. Klicka inte automatiskt vidare i Visma.
2. Läs felmeddelandet i terminalen.
3. Kontrollera den öppna dialogen och den aktuella fakturan.
4. Kontrollera loggfilens sista rad.
5. Starta om från rätt fakturanummer först när du vet om den senaste fakturan registrerades eller inte.

---

## 14. Snabbguide för daglig användning

### Före körning

- [ ] Säkerhetskopia är skapad.
- [ ] Excel-filen är kontrollerad.
- [ ] Visma visar **Obetalda**.
- [ ] Konto **1930** är valt.
- [ ] Torrkörningen är godkänd.

### Körning

```powershell
python visma_register_inbetalningar_fixed.py "C:\ekonomi\scripts\betalningar.xlsx" --live
```

- [ ] Skriv `JA`.
- [ ] Ange startfaktura.
- [ ] Ange antal, börja med 1 vid test.
- [ ] Kontrollera förhandsvisningen.
- [ ] Rör inte mus eller tangentbord.

### Efter körning

- [ ] Kontrollera terminalens sammanfattning.
- [ ] Kontrollera CSV-loggen.
- [ ] Kontrollera **Till betalning >>>** i Visma.
- [ ] Slutför Bankgiro/Bokför manuellt.

---

## 15. Kommandon i korthet

```powershell
# Installera beroenden
python -m pip install pandas openpyxl pyautogui pywinauto pyperclip

# Självtest
python visma_register_inbetalningar_fixed.py --self-test

# Kalibrering
python visma_register_inbetalningar_fixed.py --calibrate

# Torrkörning
python visma_register_inbetalningar_fixed.py "C:\ekonomi\scripts\betalningar.xlsx"

# Skarp körning
python visma_register_inbetalningar_fixed.py "C:\ekonomi\scripts\betalningar.xlsx" --live
```

> **Rekommenderad rutin:** torrkörning → en skarp testfaktura → kontroll → resterande fakturor.
