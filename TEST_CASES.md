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

---

## TEST 17 — Q31-Sync auf allen Monatsblättern (FP-Flukt)

### Scenario
Mappe mit Austritten speichern, schließen und neu öffnen. Danach jedes Monatsblatt aktivieren.

### Expected
- `Q31` auf JEDEM Monatsblatt zeigt denselben Monatswert wie die FLUKTUATION-Monatstabelle/UBERSICHT
  (Quelle: `PID_ComputeFluctuationForPeriod`), auch ohne vorheriges Aktivieren des Blatts.
- Kein 0 % mehr auf einzelnen Monatsblättern, obwohl Austritte vorhanden sind.
- `O31` zeigt das Label „Fluktuation:" (persistent aus dem Format-Makro, zusätzlich vom Sync abgesichert).
- Beim Öffnen werden alle Monatsblätter deferred synchronisiert (nach `ScreenUpdating=True`),
  beim Aktivieren eines Monatsblatts wird dessen `Q31` erneut synchronisiert.

### Negative checks
- Der Blattschutz ist nach dem Sync wieder aktiv (UserInterfaceOnly, `AllowSorting:=False`).
- Q31-Schreibzugriffe lösen beim Öffnen keine Change-/Activate-Folgeevents aus (Events temporär aus).
- Keine Änderung an CopyData, Stundenlogik, Passwortlogik oder Mac-Schutzlogik.

---

## TEST 18 — FLUKTUATION: Legende, Zielwert & Erklärungen (FP-Flukt)

### Scenario
FLUKTUATION-Blatt öffnen, bis ganz nach unten unter „Kurz erklärt" scrollen sowie KPI-Block prüfen.

### Expected
- Neuer Block „Bewertung & Ziel" unter „Kurz erklärt" mit:
  - Zielwert „unter 20 % Jahresfluktuation (Sehr gut / stabil)".
  - Legende der 6 Bewertungsstufen (0–19,99 % … ab 100 %) mit den HR-Controlling-Texten.
  - Hinweis „Monatswerte" (einzelner Austritt ≈ 1–2 %/Monat, Führungsbewertung über YTD).
- KPI-Label heißt „Verlust-Score (nur Info)".
- Erklärzeile „Verlust-Score (nur Info)" enthält den Hinweis, dass der Score die
  Fluktuationsbewertung NICHT beeinflusst.
- Management-Kurzbewertung (`C5`) enthält zusätzlich „Ziel: unter 20 % …" und „Handlungsbedarf: …".

### Negative checks
- Der Legendenblock ist rein additiv (ganz unten); bestehende Sektionen/Spalten bleiben unverändert.
- „Aktueller Personalbestand" als KPI ist NICHT umgesetzt (nur dokumentiert, siehe CHANGELOG).

---

## TEST 19 — Späterer Override == früherer Quellwert (FP-028)

### Scenario
Für denselben Mitarbeiter (Feld F, Stunden):
1. April F=173, CopyData ab April.
2. Juli F=69, CopyData ab Juli.
3. November F=173 (bewusst zurück auf den Ausgangswert), CopyData.
4. Zur Sicherheit zusätzlich CopyData **ab April** erneut ausführen (verarbeitet den gesamten Log).

### Expected
- April–Juni = 173
- Juli–Oktober = 69
- November–Dezember = 173
- Der November-Override (173) bleibt erhalten, obwohl sein Wert dem April-Quellwert (173) entspricht.
- Auch nach erneutem CopyData ab April bleibt November–Dezember = 173 (Regression-Kern von FP-028).

### Diagnostic
- `PID_ShowHourOverrideLog` vor/nach jedem CopyData: der Eintrag `(11, F, 173)` darf NICHT
  durch `PID_PruneHourOverrideLogForCopy` gelöscht werden (Vergleich gegen laufenden
  Segmentwert 69, nicht gegen Quellwert 173).

### Negative checks
- TEST 9–11 (FP-030) bleiben unverändert grün:
  - Mehrfache Änderung im selben Monat: letzter Wert gewinnt.
  - Unabhängiger späterer Override überlebt eine frühere Editierung.
  - Mittelmonats-Override löscht keine Nachbar-Overrides.

---

## TEST 20 — Neutraler Austritt „Karenz" zählt nicht in die Rate (FP-Flukt FIX 2)

### Scenario
1. Mitarbeiter mit Austrittsdatum im Zeitraum, Austrittsgrund = „Karenz".
2. UBERSICHT öffnen, FLUKTUATION neu aufbauen, betroffenes Monatsblatt `Q31` prüfen.

### Expected
- Die Fluktuationsrate (Monat/Quartal/YTD) steigt durch diesen Austritt NICHT.
- Der Austritt bleibt in FLUKTUATION als „Neutrale Bewegung" sichtbar (Anzeige unverändert).
- `Q31` (Monatsblatt), UBERSICHT Spalte Q und FLUKTUATION zeigen denselben Ratenwert.

---

## TEST 21 — Neutraler Austritt „Store transfer" zählt nicht in die Rate (FP-Flukt FIX 2)

### Scenario
1. Austrittsgrund = „Store transfer" (mit Leerzeichen) bzw. „Storetransfer" (ohne).
2. FLUKTUATION/UBERSICHT/`Q31` prüfen.

### Expected
- Beide Schreibweisen werden als neutral erkannt (Helper toleriert Leerzeichen/Groß-/Kleinschreibung).
- Rate steigt nicht; Eintrag bleibt als „Neutrale Bewegung" sichtbar.

---

## TEST 22 — Neutraler Austritt „Beförderung" zählt nicht in die Rate (FP-Flukt FIX 2)

### Scenario
1. Austrittsgrund = „Beförderung" (auch „Befoerderung").
2. FLUKTUATION/UBERSICHT/`Q31` prüfen.

### Expected
- Mit und ohne Umlaut wird neutral erkannt; Rate steigt nicht.
- Eintrag bleibt als „Neutrale Bewegung" sichtbar.

---

## TEST 23 — Neutraler Austritt „Nicht eingetreten" zählt nicht in die Rate (FP-Flukt FIX 2)

### Scenario
1. Austrittsgrund = „Nicht eingetreten".
2. FLUKTUATION/UBERSICHT/`Q31` prüfen.

### Expected
- Rate steigt nicht; Eintrag bleibt als „Neutrale Bewegung" sichtbar.

### Negative checks (TEST 20–23 gemeinsam)
- Ein NICHT-neutraler Austritt (z.B. „Dienstnehmer Kündigung") im selben Zeitraum erhöht die Rate weiterhin korrekt.
- Es gibt nur EINE zentrale Wahrheit für neutrale Gründe: `PID_IsNeutralFluctuationExitReason`.
  Sowohl die Klassifizierung (`GetFluctuationCategory` → „Neutrale Bewegung") als auch BEIDE
  Zähler (`PID_CountExitsInPeriod` und der Live-Pfad `PID_CalculateFluctuation`) nutzen diesen
  Helper — keine doppelte Logik.
- **Live-Pfad:** Trägt man auf dem Monatsblatt ein Austrittsdatum für einen neutralen Grund ein
  (Spalte I gefüllt, Spalte N = z.B. „Beförderung"), zeigt `Q31` sofort 0,00 % — identisch zum
  späteren Sync und zu UBERSICHT/FLUKTUATION.
- Nenner (durchschnittlicher Personalbestand) bleibt unverändert.

---

## TEST 24 — Reine Austrittsgrund-Änderung berechnet neu (FP-Flukt FIX 2)

### Scenario
1. Auf einem Monatsblatt existiert bereits ein Austrittsdatum (Spalte I), Grund = z.B. „Dienstnehmer Kündigung" → `Q31` zeigt einen Wert > 0.
2. NUR Spalte N (Austrittsgrund) auf „Beförderung" ändern (Austrittsdatum bleibt unverändert).

### Expected
- `Q31` aktualisiert sich SOFORT auf 0,00 %, wenn dies der einzige Austritt im Monat ist
  (gleiche Neuberechnung wie bei einer Änderung von Spalte I).
- Beim Aktivieren von FLUKTUATION/UBERSICHT sind die Werte konsistent (Dirty-Flag wird bei
  N-Änderung gesetzt — wie bei D/I).
- Umgekehrt: Grund von „Beförderung" zurück auf einen nicht-neutralen Grund ändern → `Q31` steigt
  sofort wieder.

### Negative checks
- Eine N-Änderung in einer Zeile OHNE Austrittsdatum ändert die Rate nicht (kein Austritt).
- Layout unverändert; nur der unmittelbare Q31-Trigger hört jetzt zusätzlich auf Spalte N.

---

## TEST 25 — Paste als Text in B/C/M/N + Panel O18:Q28 (FP-032 / TR-02)

### Scenario
1. Monatsblatt: in B3 eine Personal-ID mit führenden Nullen aus einer anderen Zelle/Datei einfügen (z.B. `00123`).
2. In C3 einen namenähnlichen Text mit Datums-/Zahlenoptik einfügen (z.B. `01.02.Muster`).
3. Optional: M/N Freitext einfügen; Panelzelle in O18:Q28 (inkl. Zeile 26–28) aus formatierter Quelle einfügen.
4. Mehrzellen-Paste über B3:C4.
5. Normale Tastatur-Eingabe in D3 (Datum) und F3 (Stunden) — Kontrolle, dass Zahl-/Datumsformat erhalten bleibt.

### Expected
- B/C (und M/N) haben nach Paste `NumberFormat = @` (Text); Wert unverändert, führende Nullen erhalten.
- Panel O18:Q28: nur Werte übernommen (kein Quell-Zahlen-/Währungsformat „kleben“).
- D/F unverändert nutzbar (Datum bzw. Zahl), kein erzwungenes `@`.

### Negative checks
- Ein-Zellen-Tipp in B/C setzt/hält Textformat, ohne bestehende Werte zu löschen.
- CopyData / E-F-Logging / Fluktuation durch diesen Paste-Pfad nicht beeinträchtigt.