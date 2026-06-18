# PERSONALSHEET TEST CASES

## TEST 1 — Future Hour Change

### Scenario
Employee hours changed in July.

### Expected
- May unchanged
- June unchanged
- July updated
- August+ updated

---

## TEST 2 — Exit Employee

### Scenario
Employee exit date entered in August.

### Expected
- Employee visible in August
- Employee removed from September onward

---

## TEST 3 — Future Employee Survival

### Scenario
New employee manually added in October.
Backward copy started from August.

### Expected
- October employee survives
- No accidental deletion

---

## TEST 4 — Override Survival

### Scenario
D/E values manually changed in future months.

### Expected
- Overrides survive propagation
- Previous months unchanged

---

## TEST 5 — O18:Q25 Propagation

### Scenario
Month propagation executed.

### Expected
- O18:Q25 always copied correctly

---

## TEST 6 — Column L

### Scenario
Month propagation executed.

### Expected
- Column L never propagates

---

## TEST 7 — Excel 2016 Compatibility

### Scenario
Workbook opened on Excel 2016.

### Expected
- No broken formulas
- No unsupported functions
- No VBA compile errors

---

## TEST 8 — Mac Compatibility

### Scenario
Workbook opened on MacOS Excel.

### Expected
- Macros function correctly
- No path issues
- No Windows-only dependencies

---

## TEST 9 — Repeated Hour Change Same Month (FP-030)

### Scenario
1. Change July hours to 150, run CopyData.
2. Change July hours again to 140, run CopyData again.

### Expected
- Second change wins: July=140 propagates to December.
- System does NOT stick to the first value (150).
- Windows + Mac identical.

---

## TEST 10 — Independent Later Override Survives Earlier Edit (FP-030)

### Scenario
1. Change July to 150, run CopyData.
2. Change November to 160, run CopyData.
3. Re-edit July to 140, run CopyData.

### Expected
- July–October = 140 (earlier fix propagates forward).
- November–December = 160 (independent later override survives the July edit).
- Editing an earlier month must NOT wipe a later month's explicit override.

---

## TEST 11 — Middle-Month Edit Keeps Both Neighbours (FP-030)

### Scenario
1. July=150, November=160 (each followed by CopyData).
2. Change September to 145, run CopyData.

### Expected
- July–August = 150, September–October = 145, November–December = 160.
- No override is silently deleted by editing a month between two existing overrides.

### Diagnostic
- Run `PID_ShowHourOverrideLog` before/after each CopyData to inspect stored overrides
  (read-only; does not modify data).

---

## TEST 12 — Einheitliche Fluktuation: Monat (FP-Flukt)

### Scenario
Monatsfluktuation = Austritte des Monats / Durchschnittsbestand des Monats,
Durchschnittsbestand = (Bestand 1. des Monats + Bestand Monatsende) / 2.

### Expected
- Beispiel: Bestand 1. Tag = 20, Monatsende = 18, Austritte = 2 → 2 / ((20+18)/2) = 2/19 ≈ 10,53 %.
- `Q31` (Live nach Austrittsdatum-Eingabe) und Monatswert im FLUKTUATION-Blatt verwenden denselben Nenner.
- Mitarbeiter mit Austritt genau am Stichtag zählt am Stichtag NICHT mehr zum Bestand.

---

## TEST 13 — Einheitliche Fluktuation: Quartal (FP-Flukt)

### Scenario
Quartalswert direkt aus Gesamtaustritten des Quartals / ((Bestand Quartalsanfang + Bestand Quartalsende)/2).

### Expected
- Quartal wird NICHT aus den drei Monatsprozenten gemittelt.
- Q1 = Austritte Jan–Mär / ((Bestand 01.01. + Bestand 31.03.)/2).
- Bei Bestand 0 im Zeitraum → 0 % (keine Division durch 0).

---

## TEST 14 — Einheitliche Fluktuation: YTD + Bewertung (FP-Flukt)

### Scenario
YTD von Januar bis zum aktuellen Auswertungsmonat; zusätzlich Textbewertung neben KPI „Jahresfluktuation".

### Expected
- YTD = Gesamtaustritte (Jan–aktueller Monat) / ((Bestand 01.01. + Bestand Periodenende)/2).
- YTD wird NICHT aus Monatsprozenten gemittelt.
- Bewertungstext laut Grenzwerten: <20 % „Sehr gut / stabil", <35 % „Gut / normal",
  <50 % „Erhöht / beobachten", <70 % „Hoch / analysieren", <100 % „Sehr hoch / kritisch",
  ≥100 % „Extrem hoch / akuter Handlungsbedarf".

---

## TEST 15 — Spalte „Ø Verlust-Score" (FP-Flukt)

### Scenario
Monatstabelle im FLUKTUATION-Blatt prüfen.

### Expected
- Die frühere Spalte „Durchschnitt" heißt jetzt „Ø Verlust-Score" (= monatlicher Verlust-Score / Austritte).
- Sie wird NICHT mit dem durchschnittlichen Personalbestand verwechselt.
- Zielzellen `Q31`, `UBERSICHT` Spalte Q und das FLUKTUATION-Layout bleiben unverändert.

---

## TEST 16 — Eine einzige Bewertungslogik (Rate, kein Verlust-Score-Status) (FP-Flukt)

### Scenario
FLUKTUATION-Blatt: oberer Status (B5) und Bewertungszeile (B10) prüfen.

### Expected
- **Status `B5`** = Kurz-Ampel aus YTD-Rate: „Sehr gut / Gut / Erhöht / Hoch / Kritisch / Extrem"
  (Grenzwerte 0,2 / 0,35 / 0,5 / 0,7 / 1,0), eingefärbt grün → rot.
- **Bewertung `B10`** = Langform aus `PID_GetFluctuationRating(ytdFluctuation)`
  („Sehr gut / stabil" … „Extrem hoch / akuter Handlungsbedarf").
- Status und Bewertung stammen aus DERSELBEN Rate; B5 und B10 zeigen NICHT denselben Text.
- Es gibt KEINEN „Kritisch"-Status mehr aus dem Verlust-Score.
- Der Statustext `C5` bezieht sich auf die YTD-Rate, nicht auf den Verlust-Score.

### Negative checks
- Verlust-Score, „Ø Verlust-Score" und „Kritische Austritte" bleiben als informative KPIs sichtbar,
  bestimmen aber Status/Bewertung NICHT.
- Bei fehlenden Austrittsgründen: KEIN „Daten prüfen"-Status-Override; nur ein informativer Hinweis
  (KPI „Daten offen" + Abschnitt „Sofort prüfen" bleiben sichtbar).
- „Sofort prüfen" und „Empfehlungen" bleiben rein operative Hinweise und überschreiben nichts.