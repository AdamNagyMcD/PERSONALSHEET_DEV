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

## TEST 8 — Windows Platform

### Scenario
Workbook opened on Excel for Windows.

### Expected
- Macros function correctly
- No path issues
- Smoke check reports PASS for the Windows platform test

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

## TEST 25 — Personal-ID / Name korrigieren (TR-06)

Modul: `mod_MitarbeiterPflege.bas`.
Makro: `PersonalIdKorrigieren` — Button **„Personal-ID korrigieren" (Q7:R7)** auf jedem
Monatsblatt, Alt+F8, oder Admin-Button „Personal-ID fix".

### Scenario A — falsche ID auf allen Monaten korrigieren
1. Auf Januar einen Mitarbeiter mit falscher Personal-ID anlegen, CopyData bis Dezember.
2. Auf einem beliebigen Monatsblatt die Zeile des Mitarbeiters markieren.
3. Makro starten, neue (richtige) ID eingeben, Namen leer lassen, bestätigen.

### Expected A
- Bestätigungsdialog listet ALT/NEU und die betroffenen Monate mit Zeilenzahl.
- Nach dem Lauf steht in ALLEN 12 Monatsblättern die richtige ID.
- Anschliessendes CopyData ab Januar bringt die falsche ID NICHT zurück
  (keine zusätzliche „Geister"-Zeile).

### Scenario B — Stunden-Override überlebt die Korrektur
1. Für denselben Mitarbeiter in einem Monat (z.B. Juli) die Stunden (F) abweichend setzen.
2. ID über das Makro korrigieren.
3. CopyData ab einem früheren Monat ausführen.

### Expected B
- Der Juli-Override bleibt erhalten (das Log `PID_HOUR_OVERRIDES` wurde mit umgeschlüsselt).
- Die Abschlussmeldung nennt die Anzahl der geänderten Log-Einträge.

### Scenario C — Konflikt: Ziel-ID gehört schon jemandem
1. Zwei Mitarbeiter A und B mit unterschiedlichen IDs.
2. Für A die ID von B eintragen wollen.

### Expected C
- Meldung „Korrektur nicht möglich - es wurde nichts geändert" mit Blatt und Zeile.
- KEIN Blatt wurde verändert (Abbruch vor dem ersten Schreibvorgang).

### Scenario D — nur Namensänderung
1. Zeile markieren, ID leer lassen, neuen Namen eingeben.

### Expected D
- Name auf allen Monaten geändert, Log-Schlüssel angepasst, Stunden-Overrides intakt.

### Negative checks
- Start ohne markierte Mitarbeiterzeile (z.B. vom Admin-Panel): Makro fragt per InputBox
  nach der aktuellen ID; bei mehreren Namen zu einer ID bricht es mit Hinweis ab.
- ID mit führender Null (`00123`) bleibt nach der Korrektur als Text erhalten.
- Blattschutz ist nach dem Lauf auf allen berührten Monatsblättern wieder aktiv.
- Berechnungsmodus steht nach dem Lauf wieder auf dem Ausgangswert (nicht „Manuell").

---

## TEST 26 — Mitarbeiter aus Monaten entfernen (TR-08)

Modul: `mod_MitarbeiterPflege.bas`.
Makro: `MitarbeiterEntfernen` — Button **„Mitarbeiter entfernen" (O7:P7)** auf jedem
Monatsblatt, Alt+F8, oder Admin-Button „MA entfernen".

### Scenario A — aus allen 12 Monaten entfernen
1. Mitarbeiter in Januar anlegen, CopyData bis Dezember.
2. Zeile des Mitarbeiters markieren, Makro starten.
3. Bei der Frage nach dem Zeitraum **JA** wählen (alle 12 Monate), bestätigen.

### Expected A
- Bestätigungsdialog listet die betroffenen Monate mit Zeilenzahl und weist auf das
  Austrittsdatum (Spalte I) als Alternative hin.
- Nach dem Lauf ist die Zeile in ALLEN 12 Monaten leer (B:N ohne Inhalt).
- Zeilen wurden NICHT gelöscht: Zeilenanzahl, Zebra-Formatierung und Struktur unverändert.
- Andere Mitarbeiter sind unverändert.

### Scenario B — erst ab einem bestimmten Monat entfernen
1. Mitarbeiter existiert Januar–Dezember.
2. Makro starten, **NEIN** wählen, Monatsnummer `7` eingeben.

### Expected B
- Januar–Juni bleiben unverändert (Mitarbeiter weiterhin vorhanden).
- Juli–Dezember sind geleert.
- Die Abschlussmeldung nennt „Juli bis Dezember".

### Scenario C — Stunden-Log wird mit bereinigt
1. Für den Mitarbeiter in Oktober abweichende Stunden (F) setzen (Override entsteht).
2. Mitarbeiter ab Juli entfernen.
3. `PID_ShowHourOverrideLog` prüfen.

### Expected C
- Der Oktober-Override dieses Mitarbeiters ist weg (Monate >= Startmonat).
- Overrides ANDERER Mitarbeiter sind unverändert.
- Ein späteres erneutes Anlegen desselben Mitarbeiters bringt die alten Stunden NICHT zurück.

### Scenario D — CopyData nach dem Entfernen
1. Mitarbeiter ab Juli entfernen.
2. CopyData ab Juni ausführen.

### Expected D
- Der Mitarbeiter erscheint ab Juli NICHT wieder (er existiert ab Juni nicht mehr in der Quelle
  bzw. wird nicht als Neuzugang erkannt).
- Hinweis: CopyData ab einem Monat VOR dem Entfernen (z.B. ab Januar) verteilt den Mitarbeiter
  bewusst wieder nach vorne — das ist erwartetes CopyData-Verhalten, kein Fehler.

### Negative checks
- Abbrechen im Zeitraum-Dialog ändert nichts.
- Ungültige Monatsnummer (0, 13, Text) wird abgelehnt, ohne etwas zu ändern.
- Mitarbeiter im gewählten Zeitraum nicht vorhanden → Meldung, keine Änderung.
- Blattschutz nach dem Lauf auf allen berührten Monatsblättern wieder aktiv.
- Berechnungsmodus steht wieder auf dem Ausgangswert (nicht „Manuell").
- G/H/K/L in der geleerten Zeile zeigen keine Fehlerwerte; nach erneutem Befüllen der Zeile
  rechnen sie wieder korrekt.

---

## TEST 27 — Personal-ID Eindeutigkeit (TR-07)

Modul: `mod_PersonalIdUnique.bas`, ausgelöst aus `Workbook_SheetChange`.

### Scenario A — doppelte ID von Hand eingeben
1. Auf einem Monatsblatt hat Zeile 5 die ID `10457` (Max Mustermann).
2. In Zeile 12 dieselbe ID `10457` eintippen.

### Expected A
- Meldung „Personal-ID bereits vergeben" nennt Zeile 12, die ID und Zeile 5 mit Namen.
- **B12 ist leer** — die neue Eingabe wurde zurückgewiesen.
- **B5 bleibt unverändert** — der Bestand wird nie angefasst.

### Scenario B — Einfügen mehrerer Zeilen mit derselben ID
1. Zwei Zeilen mit identischer Personal-ID in den Bereich B3:B82 einfügen.

### Expected B
- Die erste eingefügte Zeile behält die ID, jede weitere wird geleert.
- Eine einzige Sammelmeldung (nicht ein Popup pro Zeile).
- Bei mehr als 8 Konflikten endet die Liste mit „… und N weitere".

### Scenario C — weicher Hinweis bei abweichendem Namen
1. In Januar existiert ID `10457` = „Max Mustermann".
2. In März in einer freien Zeile ID `10457` und Name „Maxi Muster" eintragen.

### Expected C
- Kein Zurückweisen (im März ist die ID eindeutig).
- Info „Personal-ID prüfen" listet `Januar: Max Mustermann` und verweist auf das Makro
  `Personal-ID korrigieren`.
- Gleicher Name in beiden Monaten → **keine** Meldung.

### Scenario D — Makros werden nicht blockiert
1. CopyData über mehrere Monate laufen lassen.
2. `PersonalIdKorrigieren` und `MitarbeiterEntfernen` ausführen.

### Expected D
- Keine Eindeutigkeits-Meldung während dieser Makros (`EnableEvents = False`).
- Kein Datenverlust, keine geleerten B-Zellen durch die Prüfung.

### Negative checks
- Leere B-Zellen lösen nichts aus (mehrere leere Zeilen sind erlaubt).
- Führende Nullen: `00123` und `123` gelten als verschiedene IDs (bekannte Grenze, TR-02).
- Änderung in anderen Spalten (C, E, F …) löst die Prüfung nicht aus.
- Nach dem Zurückweisen ist der Blattschutz unverändert aktiv.

---

## TEST 28 — Fehler melden + Aktionsprotokoll (TR-09)

Module: `mod_PIDFeedback.bas`, `mod_PIDActionLog.bas`.
Makro: `FehlerMelden` — Button **„Fehler melden" (S7:T7)** auf jedem Monatsblatt,
Alt+F8, oder Admin-Button „Fehler melden".

### Scenario A — Meldung erstellen
1. Auf einem Monatsblatt eine Zelle markieren (z.B. `B12`).
2. Button „Fehler melden" klicken.
3. Beide Fragen beantworten.

### Expected A
- Datei `Feedback\Fehlermeldung_JJJJ-MM-TT_hhmmss.txt` liegt neben der Mappe,
  der Pfad steht in der Abschlussmeldung.
- Inhalt enthält Version, Datei, Ordner, Excel-Version, Benutzer, Jahr,
  **Blatt und Auswahl `B12`** (also den Stand VOR den Dialogen), Rechenmodus.
- Strg+V in einer E-Mail oder im Notepad fügt denselben Text ein.

### Scenario B — letzte Aktionen landen im Bericht
1. `MitarbeiterEntfernen` oder `DataClear` ausführen.
2. Danach „Fehler melden".

### Expected B
- Abschnitt „Letzte Aktionen" listet die eben ausgeführte Aktion mit Zeitpunkt,
  Blatt und Detail (z.B. Zeilenzahl).
- Admin-Button „Aktionsprotokoll" zeigt dieselben Einträge.
- Das Blatt `PID_ACTION_LOG` bleibt **sehr versteckt** und die vorher aktive
  Registerkarte bleibt aktiv (kein Blattwechsel beim ersten Protokolleintrag).

### Scenario C — Abbrechen
1. „Fehler melden" starten und die erste Frage abbrechen.

### Expected C
- Keine Datei, keine Meldung, keine Änderung an der Mappe.

### Scenario D — E-Mail an Adam
1. Meldung erstellen und die Frage nach der E-Mail mit **Ja** beantworten.

### Expected D
- Outlook öffnet ein **neues, nicht abgesendetes** Mailfenster.
- Empfänger `adam.nagy@at.mcd.com`, Betreff mit Version, Bericht im Text,
  die Textdatei als Anhang.
- Ohne Outlook: das Standard-Mailprogramm öffnet sich über `mailto` (gekürzter Text).
- Bei **Nein**: nichts passiert, Datei und Zwischenablage bleiben erhalten.

### Negative checks
- Beide Fragen leer lassen → Hinweis, keine Datei.
- Nach CopyData steht ein Eintrag `CopyData | <Quellmonat> -> Dezember (N Monate)` im Protokoll,
  und CopyData selbst läuft unverändert durch (Bootstrap-Modul, nur diese eine Zeile).
- Nach 500 Protokolleinträgen werden die ältesten entfernt, die Datei wächst nicht weiter.
- Ist kein Protokoll vorhanden, steht im Bericht „(keine aufgezeichnet)" statt eines Fehlers.
- Der Rechenmodus im Bericht zeigt „Automatisch", solange kein Makro hängen geblieben ist.

---

## TEST 29 — Einfügen immer nur als Wert (TR-02)

Module: `mod_PIDPasteValues.bas`, `DieseArbeitsmappe.cls`.
Kein Makroaufruf nötig — der Schutz läuft über `Strg+V` und über `Workbook_SheetChange`.

### Scenario A — externe Quelle (Word / Browser)
1. In Word einen **farbigen, fetten** Namen schreiben und kopieren.
2. Auf einem Monatsblatt in `C12` mit **Strg+V** einfügen.

### Expected A
- Nur der Text steht in der Zelle.
- Schriftart, Farbe, Rahmen und Zahlenformat der Zelle bleiben wie im Blatt definiert.
- Das Dropdown in derselben Zeile (Spalte F) funktioniert weiterhin.

### Scenario B — andere Excel-Datei
1. In einer anderen Excel-Datei einen Block `B:C` mit eigener Formatierung kopieren.
2. Im Monatsblatt `B12` markieren und **Strg+V**.

### Expected B
- Nur Werte, keine Quellformatierung.
- Bei doppelter Personal-ID greift zusätzlich die Prüfung aus TEST 27.

### Scenario C — Menüband und Rechtsklick
1. Denselben Inhalt über **Start → Einfügen** und über **Rechtsklick → Einfügen** einfügen.

### Expected C
- Ergebnis identisch zu Scenario A (das Netz in `Workbook_SheetChange` räumt nach).
- Auch beim Einfügen in **Spalte E** bleibt nichts von der Quellformatierung stehen
  (das war der alte Fehlerfall: der Dropdown-Neuaufbau hatte die Undo-Historie geleert).

### Scenario D — geschützte Formelspalten
1. Einen 6 Spalten breiten Block kopieren.
2. In `B12` einfügen, sodass die Formelspalten G/H erreicht würden.

### Expected D
- Hinweis „… würde geschützte Zellen … überschreiben", **nichts** wird eingefügt.
- Die Formeln in G/H/K/L sind unverändert.

### Scenario E — Ausschneiden
1. Eine Zeile mit **Strg+X** ausschneiden und **Strg+V** drücken.

### Expected E
- Hinweis, dass Ausschneiden + Einfügen nicht vorgesehen ist; keine Änderung im Blatt.

### Negative checks
- **Andere Datei:** zweite Excel-Datei öffnen, dorthin wechseln, formatiert kopieren und
  einfügen → dort funktioniert **Strg+V wie gewohnt** (mit Formatierung). Zurück zum
  Personalsheet wechseln → dort wieder nur Werte.
- Nach dem Schließen der Mappe ist `Strg+V` in Excel wieder Excel-Standard.
- Normale Eingabe über die Tastatur und die Dropdowns verhalten sich unverändert.
- Löschen einer Mehrfachauswahl (Strg-Klick, dann Entf) löscht wie gewohnt.
- Einfügen einer sehr langen Liste (> 20 000 Zellen) wird mit Hinweis abgelehnt,
  es gehen keine Daten verloren.

---

## TEST 30 — Formeln überleben das Löschen (TR-10)

Module: `mod_DataClear.bas`, `mod_MitarbeiterPflege.bas`, `Modul1.bas`.
Betroffene Spalten: **G** (Monatslohn), **H** (Aktuelle Stunden), **K** (Urlaub Euro),
**L** (Letztes Gehalt).

### Scenario A — Mitarbeiter entfernen
1. Auf einem Monatsblatt eine gefüllte Mitarbeiterzeile markieren (z.B. Zeile 12).
2. Button „Mitarbeiter entfernen", alle Monate wählen, bestätigen.
3. In jedem Monatsblatt Zeile 12 anklicken und `G12`, `H12`, `K12`, `L12` in der
   Bearbeitungsleiste prüfen.

### Expected A
- B:F, I:J und M:N sind leer.
- **In G, H, K und L steht weiterhin die Formel**, die Zellen sehen leer aus
  (B/C-Guard), zeigen aber keine 0 und kein `#WERT!`.
- Wird danach in dieselbe Zeile ein neuer Mitarbeiter eingetragen (ID, Name, Eintritt,
  KV-Gruppe, Stunden), rechnen G, H, K und L sofort wieder.

### Scenario B — Zeilen löschen und Monat löschen
1. `PID_ClearOnlySelectedEmployeeRows` auf zwei markierten Zeilen ausführen.
2. `DataClear` auf einem Monatsblatt ausführen.

### Expected B
- Gleiches Ergebnis wie A; nach `DataClear` haben **alle** Zeilen 3–82 in G/H/K/L Formeln.
- Der Bestätigungsdialog nennt „Formelspalten G, H, K, L" als „bleibt erhalten".

### Scenario C — Selbstheilung alter Schäden
1. Eine Zeile suchen, in der eine ältere Version die Formeln bereits gelöscht hat
   (G/H/K/L leer, keine Formel).
2. Dort einen Mitarbeiter eintragen und ihn danach wieder entfernen.

### Expected C
- Nach dem Entfernen stehen in G, H, K und L wieder Formeln
  (`PID_RestoreFormulaColumnsForRows` setzt fehlende Formeln zeilengenau neu).

### Negative checks
- Bereits beschädigte Blätter komplett reparieren: **Admin → Full Refresh**
  (`PID_FullSystemRefresh`) setzt G/H/K/L auf allen 12 Monatsblättern neu.
- Eine vorhandene, bewusst abweichende Formel in G/H/K/L wird beim Löschen **nicht**
  überschrieben (nur fehlende Formeln werden ergänzt).
- Zeilenhöhe, Zebra-Streifen, Rahmen und Zahlenformate bleiben unverändert.
- Nach dem Entfernen bleibt das Blatt geschützt.
---

## TEST 31 — Formelspalten prüfen und reparieren (TR-10)

Module: `mod_PIDFormelCheck.bas`, `mod_PIDAdmin.bas`.
Voraussetzung: `_ADMIN` sichtbar (**Alt+F8** → `PID_ToggleAdminSheet`).

Ausgangslage in der Testdatei (mit `tools/check_formula_columns.py` ermittelt):
in **Februar bis Dezember** fehlt in **L3, L4 und L5** die Formel — 33 Zellen insgesamt.

### Scenario A — Diagnose
1. `_ADMIN` → **„Formeln prüfen"** (oder **Alt+F8** → `PID_PruefeFormelspalten`).

### Expected A
- Meldung listet „Geprüft: 12 Monatsblätter", die Spalten G, H, K, L und Zeile 3 bis 82.
- Jedes Blatt mit Lücken erscheint mit der Anzahl fehlender Zellen
  (erwartet: Februar bis Dezember mit je 3).
- Es wird nichts verändert: ein zweiter Lauf zeigt dieselben Zahlen.

### Scenario B — Reparatur
1. `_ADMIN` → **„Formeln reparieren"** (oder **Alt+F8** → `PID_FormelspaltenReparieren`).
2. Danach erneut **„Formeln prüfen"**.

### Expected B
- Erste Meldung: „Neue Formeln: 33 Zellen auf 11 Monatsblättern".
- Zweite Meldung: „Alle Formeln in den Spalten G,H,K,L sind vorhanden."
- `Februar!L4` (Zeile mit Mitarbeiter) zeigt in der Bearbeitungsleiste wieder eine Formel
  und rechnet einen Betrag.

### Scenario C — Full Refresh meldet die Reparatur
1. Datei schließen ohne zu speichern, neu öffnen (Schaden ist wieder da).
2. `_ADMIN` → **„Full Refresh"**.

### Expected C
- Abschlussmeldung enthält die drei neuen Zeilen: „Formelspalten G/H/K/L geprüft: 12",
  „Fehlende Formeln ergänzt: N Zellen", „Monatsindex A1 korrigiert: 0".
- Danach meldet „Formeln prüfen" keine Lücken mehr.

### Scenario D — Monatsindex A1
1. Blattschutz aufheben (`UnprotectEverything`), auf **Mai** die Zelle `A1` leeren.
2. `_ADMIN` → **„Formeln prüfen"**, danach **„Formeln reparieren"**.

### Expected D
- Die Prüfung meldet „Monatsindex in A1 fehlt oder passt nicht: Mai".
- Nach der Reparatur steht in `Mai!A1` wieder **5**, und „Monatsindex A1 korrigiert: 1".
- Hintergrund: ohne A1 überspringen die Wiederherstellungen das Blatt stillschweigend.

### Negative checks
- Vorhandene Formeln werden nicht überschrieben (nur fehlende ergänzt).
- Blattschutz ist nach beiden Makros wieder aktiv.
- Rechenmodus bleibt auf Automatisch.
- Nach dem Speichern meldet `python3 tools/check_formula_columns.py`
  „Fehlende Formelzellen gesamt: 0".

---

## TEST 32 — Alle Daten löschen (TR-03)

Modul: `mod_DataClear.bas`. Einstieg: **Alt+F8** → `AlleDatenLoeschen`
oder `_ADMIN` → **„Alle Daten löschen"**.

**Vorher unbedingt eine Kopie der Datei anlegen.**

### Scenario A — Abbruch
1. Makro starten, im ersten Dialog **Nein** wählen.
2. Makro erneut starten, ersten Dialog mit **Ja**, zweiten mit **Nein** beantworten.

### Expected A
- In beiden Fällen bleibt jede Zelle unverändert.
- Der zweite Dialog steht standardmäßig auf **Nein**.

### Scenario B — Löschen
1. Makro starten und beide Dialoge mit **Ja** bestätigen.

### Expected B
- Alle 12 Monatsblätter: `B:F`, `I:J`, `M:N` (Zeile 3–82), `O18:Q28`, `O45` und `Q31` leer.
- **G, H, K, L behalten ihre Formeln** in allen 80 Zeilen (Bearbeitungsleiste prüfen).
- Abschlussmeldung: „Monatsblätter geleert: 12 / 12" und „Stunden-Log geleert".
- `EINSTELLUNG!C35` (Jahr), LOHNTABELLE und UEBERSICHT sind unverändert.
- `Q12` (Vormonat) bleibt stehen — bewusst nicht gelöscht.
- Blattschutz auf allen Monatsblättern weiterhin aktiv.

### Scenario C — Weiterarbeiten auf der leeren Datei
1. In Januar einen Mitarbeiter eintragen (ID, Name, Eintritt, KV-Gruppe, Stunden).
2. `CopyData` von Januar aus starten.
3. FLUKTUATION-Tab öffnen.

### Expected C
- G, H, K und L rechnen sofort.
- CopyData verteilt den Mitarbeiter bis Dezember.
- Fluktuation und UEBERSICHT bauen ohne Fehler neu auf (0 Austritte).

### Negative checks
- Stunden-Log (`PID_AdminShowActionLog` / `PID_ShowHourOverrideLog`) enthält keine alten
  Overrides mehr; ein früher überschriebener Monatswert kehrt nicht zurück.
- Das Aktionsprotokoll enthält den Eintrag „Alle Daten loeschen".
- Formate, Zebra-Streifen, Kopfzeilen, Buttons und Dropdowns bleiben erhalten.

---

## TEST 33 — Einheitliches Layout der Monatsblätter

Module: `mod_FormatMonthSheet.bas`, `mod_FormatEinstellung.bas`.
Einstieg: **Alt+F8** → `ADMIN_30_Format_Alle_Monate`, danach `ADMIN_32_Format_EINSTELLUNG`.

### Scenario A — Zeilenhöhe
1. `ADMIN_30_Format_Alle_Monate` ausführen.
2. Auf mehreren Monatsblättern die Zeilen 3 bis 82 markieren und die Höhe ablesen
   (Rechtsklick auf den Zeilenkopf → *Zeilenhöhe*).

### Expected A
- Jede Zeile 3–82 ist **22** hoch, auf allen zwölf Blättern einschließlich Januar.
- Die Höhe bleibt gleich, auch in Zeilen mit ausgefülltem Austrittsgrund (Spalte N).

### Scenario B — Zeilenhöhe bleibt nach CopyData
1. In einer Zeile einen Austrittsgrund aus dem Dropdown in Spalte N wählen.
2. `CopyData` starten.

### Expected B
- Die betroffene Zeile ist weiterhin 22 hoch — vorher wurde sie auf 18 zusammengezogen.

### Scenario C — Ausrichtung und Schriftgröße
1. Beliebige Zellen in `A3:N82` anklicken und Ausrichtung sowie Schriftgröße prüfen.

### Expected C
- Alle Spalten A bis N sind vertikal **mittig** ausgerichtet.
- Spalten B bis N stehen auf **11 Punkt** (vorher B und N auf 12), Spalte A bleibt bei 8.
- Die horizontale Ausrichtung ist unverändert: B, C, M, N links, der Rest zentriert.

### Scenario D — EINSTELLUNG Nachtzuschläge
1. `ADMIN_32_Format_EINSTELLUNG` ausführen, dann `O6:O17` ansehen.

### Expected D
- Die Beträge erscheinen als Eurowert (z. B. `€ 2.860,00`), nicht als `2860,00`.
- `N6:N17` zeigt weiterhin die Monatsbeschriftungen.

---

## TEST 34 — Spalte L wächst nicht mehr, Spalte G bleibt als Wert stehen

Module: `Modul1.bas`, `mod_KVLohnLookup.bas`, `mod_PIDFormelCheck.bas`.
Hintergrund: Die L-Formel wurde bei jedem Öffnen erneut umhüllt (1467 → 2981 Zeichen,
Kernformel 4× enthalten) und wäre bei rund 8192 Zeichen an der Excel-Grenze gescheitert.
Die Formelprüfung meldete gleichzeitig bis zu 50 Zeilen je Monatsblatt als „Formel fehlt",
obwohl der Zahlenwert in `G` der gewollte schnelle Zustand ist.

**Voraussetzung:** VBA neu importieren (`ADMIN_01_VBA_Import` bzw. `ResetAndImportVBAFiles`),
kompilieren, speichern.

### Scenario A — Einmalige Bereinigung der L-Formel
1. Vor dem Import in `Januar` die Zelle `L3` anklicken und die Formellänge notieren
   (Bearbeitungsleiste, alternativ im Direktfenster: `?Len(Range("L3").Formula)`).
2. VBA importieren, Datei speichern, schließen und **einmal neu öffnen**.
3. `L3` erneut ansehen: `?Len(Range("L3").Formula)`.

### Expected A
- Vorher rund **2981** Zeichen, danach rund **745** Zeichen.
- Die angezeigten Beträge in `L3:L82` sind auf allen zwölf Blättern **unverändert**.
- `Q17` (`=SUM(L3:L82)`) und `Q42` (AVG Bruttolohn) zeigen dieselben Werte wie vorher.
- Zeilen ohne Mitarbeiter bleiben **leer** (kein `€ 0,00`) — das übernimmt jetzt das
  Zahlenformat mit leerem Null-Abschnitt.

### Scenario B — Zweites Öffnen schreibt nichts mehr
1. Datei erneut schließen und öffnen.
2. `?Len(Range("L3").Formula)` in mehreren Monatsblättern prüfen.

### Expected B
- Die Länge bleibt bei rund **745** Zeichen und wächst nicht weiter.
- Das Öffnen ist nicht langsamer als vorher; es werden keine L-Spalten mehr neu geschrieben.

### Scenario C — Prüfung meldet Spalte G nicht mehr falsch
1. **Alt+F8** → `PID_PruefeFormelspalten` (oder `_ADMIN` → *Formeln prüfen*).

### Expected C
- Meldung: **0 Zellen ohne Formel** auf allen zwölf Monatsblättern.
- Vorher wurden hier rund **402** Zellen gemeldet, obwohl nichts defekt war.

### Scenario D — Reparatur ersetzt keine schnellen G-Werte
1. In `Februar` eine belegte Zeile suchen, in der `G` einen Zahlenwert ohne Formel trägt
   (Bearbeitungsleiste zeigt die Zahl, keine `=`-Formel).
2. **Alt+F8** → `PID_FormelspaltenReparieren` ausführen.
3. Dieselbe Zelle erneut ansehen.

### Expected D
- Die Zelle trägt weiterhin den **Zahlenwert**, keine `PID_KVLohnLookup`-Formel.
- Der angezeigte Lohn ist unverändert.

### Scenario E — Formel kommt zurück, wo sie fehlt
1. In `März` einen Mitarbeiter über *Dolgozó törlése* (`MitarbeiterEntfernen`) aus einem
   Monat entfernen.
2. In derselben Zeile `G`, `H`, `K` und `L` prüfen.
3. Danach in dieselbe Zeile einen neuen Mitarbeiter mit KV-Gruppe und Stunden eintragen.

### Expected E
- Nach dem Löschen steht in `G` wieder die Formel (kein stehengebliebener Lohn des
  gelöschten Mitarbeiters), `H`, `K`, `L` behalten ihre Formeln.
- Der neue Mitarbeiter erhält sofort Lohn, aktuelle Stunden, Urlaubsgeld und letztes Gehalt.

### Negative checks
- `python3 tools/check_formula_columns.py` (oder auf Windows `python tools/...`) meldet
  „Fehlende Formelzellen gesamt: 0".
- Ein Blattwechsel zwischen den Monaten fühlt sich nicht langsamer an als vorher
  (die Prüfung liest E, F und G jetzt als Array statt Zelle für Zelle).
- `PID_RunSystemSmokeCheck`: TEST 15 (Monatslohn Spalte G) bleibt **PASS**.


## TEST 35 — Blattschutz ab dem Öffnen und nach einem Abbruch

Module: `mod_SchutzHinzufugen.bas`, `Modul1.bas`, `mod_FormatEinstellung.bas`.
Hintergrund: Der Schutz wurde bisher erst beim **ersten Besuch** eines Tabs gesetzt. Wer die
Mappe auf `UEBERSICHT` liegend öffnete, hatte kein einziges Monatsblatt geschützt und konnte
die Formelspalten `G`, `H`, `K`, `L` überschreiben. Zusätzlich blieben Blätter entsperrt
zurück, wenn ein Makro mitten im Lauf abbrach.

**Voraussetzung:** VBA neu importieren (`ADMIN_01_VBA_Import`), kompilieren, speichern.

### Scenario A — Alle Blätter sind sofort nach dem Öffnen geschützt
1. Die Datei auf `UEBERSICHT` stehend speichern und schließen.
2. Datei öffnen und **ohne** einen Monats-Tab anzuklicken direkt auf `Dezember` wechseln.
3. In `G3` etwas eintippen (z. B. `1`).
4. Ebenso in `H3`, `K3`, `L3` und in `A3`.
5. Rechtsklick auf eine Zeilennummer → *Zeilen löschen* versuchen.
6. Dasselbe auf `EINSTELLUNG` in einer Zelle außerhalb der Eingabeblöcke (z. B. `B6`).

### Expected A
- Excel lehnt jede dieser Eingaben mit der Schutzmeldung ab.
- Das Löschen und Einfügen von Zeilen und Spalten ist gesperrt.
- Die erlaubten Eingabefelder funktionieren weiter: `B3:F3`, `I3:J3`, `M3:N3` und das
  Freitextfeld `O18` sind beschreibbar.

### Scenario B — Eine ungeschützt gespeicherte Datei kommt geschützt beim Benutzer an
1. `ADMIN_03_Schutz_AUS` (`UnprotectEverything`) ausführen und mit *Ja* bestätigen.
2. Datei **speichern** und schließen (das ist der Fehler, der bisher nicht auffing).
3. Datei erneut öffnen und Scenario A Punkt 3 bis 5 wiederholen.

### Expected B
- Nach dem Öffnen sind alle Blätter wieder geschützt, `G3` ist nicht beschreibbar.
- Die technischen Blätter (`PID_HOUR_OVERRIDES`, `FLUKTUATION_DATEN`, `KV_DROPDOWN_HELPER`,
  `_ADMIN`) sind wieder ausgeblendet.

### Scenario C — Sortieren bleibt auf Monatsblättern gesperrt
1. `ADMIN_11_Format_EINSTELLUNG` (bzw. `PID_FormatEinstellungSheet`) ausführen.
2. Auf `Januar` wechseln, eine Zelle in `C3:C82` anklicken und *Daten → Sortieren* versuchen.

### Expected C
- Sortieren wird abgelehnt. Vorher hatte dieses Makro auf jedem angefassten Monatsblatt
  `AllowSorting:=True` hinterlassen — ein versehentliches Sortieren hätte die Zuordnung
  aller Mitarbeiterzeilen zerrissen.
- Der AutoFilter funktioniert weiterhin.

### Scenario D — Nach einem Abbruch bleibt kein Blatt offen
1. Auf `UEBERSICHT` wechseln (das löst den Durchrechnungs-Refresh aus).
2. Während eines längeren Makros (z. B. `ADMIN_10_Format_Monatsblaetter`) **Strg+Untbr**
   drücken und im Dialog *Beenden* wählen.
3. Auf `UEBERSICHT` und auf ein Monatsblatt wechseln und dort in `G3` bzw. in eine
   Formelzelle schreiben wollen.

### Expected D
- Beide Blätter sind geschützt; die Eingabe wird abgelehnt.
- Vor der Änderung blieb genau das Blatt der abgebrochenen Runde dauerhaft entsperrt.

### Scenario E — Excel-Reset löst die häufigsten Blockaden
1. Ein Makro mit **Strg+Untbr** abbrechen, während es läuft.
2. Symptome prüfen: Ausfüllkästchen fehlt, Ziehen von Zellen geht nicht, Dropdowns in `E`/`F`
   aktualisieren nicht mehr, `Q31`/Finanzwerte reagieren nicht auf Eingaben.
3. **Alt+F8** → `ADMIN_05_Excel_Reset` (bzw. `PID_ResetExcelState`).

### Expected E
- Meldung „Excel wurde zurückgesetzt".
- Ausfüllkästchen und Ziehen sind wieder verfügbar (auch in anderen offenen Dateien).
- Eingaben in `E`/`F` lösen die Dropdown-Aktualisierung wieder aus, `Q31` und die
  Finanzspalten rechnen wieder mit.
- Einfügen per Strg+V fügt weiterhin nur Werte ein (TR-02, TEST 29).

### Negative checks
- Das Öffnen dauert nicht merkbar länger: die Schutzarbeit ist nur vom ersten Tabwechsel nach
  vorne gewandert, und die Sperrliste läuft je Blatt jetzt **einmal** statt zwei- bis dreimal.
- `PID_RunSystemSmokeCheck`: TEST 17 (Blattschutz) und TEST 18 (`Q12` auf Feb/Mai/Aug/Nov)
  bleiben **PASS**.
- `CopyData`, `DataClear`, `MitarbeiterEntfernen` und `PersonalIdKorrektur` laufen unverändert
  durch — sie schützen sich selbst um ihre Schreibvorgänge herum.


## TEST 36 — Admin-Panel wird nicht bei jedem Öffnen neu gebaut

Modul: `mod_PIDAdminSheet.bas`.
Hintergrund: `PID_EnsureAdminSheet` läuft bei jedem Öffnen und hat dabei das ganze Blatt neu
beschriftet und alle 22 Buttons gelöscht und neu angelegt — obwohl `_ADMIN` sofort danach
wieder auf `xlSheetVeryHidden` geht und ein normaler Benutzer es nie zu sehen bekommt.
Jetzt entscheidet eine Signatur (Buttonanzahl + Beschriftungen + Makronamen) in `_ADMIN!BZ1`,
ob ein Neuaufbau nötig ist.

### Scenario A — Erstes Öffnen nach dem Import baut einmal, danach nicht mehr
1. VBA importieren, speichern, schließen.
2. Datei öffnen (erster Start: `BZ1` ist leer, das Panel wird einmal gebaut), speichern, schließen.
3. Datei erneut öffnen und die Zeit bis zur Bedienbarkeit mit vorher vergleichen.
4. **Alt+F8** → `PID_ToggleAdminSheet` und das Panel ansehen.

### Expected A
- Alle **22** Buttons sind vorhanden, mit korrekter Beschriftung, und jeder startet sein Makro.
- Kopfzeile stimmt: `C3` = aktuelles Jahr aus `EINSTELLUNG!C35`, `E3` = Excel-Version.
- Das Öffnen ist ab dem zweiten Start nicht langsamer, tendenziell schneller (keine 44
  Shape-Operationen und kein `Cells.Font`-Durchlauf mehr).

### Scenario B — Selbstheilung bleibt erhalten
1. Panel anzeigen, **einen** Button von Hand löschen, `_ADMIN` wieder ausblenden.
2. Datei schließen (speichern) und erneut öffnen.
3. Panel wieder anzeigen.

### Expected B
- Der fehlende Button ist wieder da (Anzahl weicht ab → vollständiger Neuaufbau).

### Scenario C — Geänderte Button-Liste im Code wird übernommen
1. Nach einem VBA-Import, der eine Beschriftung oder einen Makronamen in
   `PID_AdminGetButtonSpec` ändert, die Datei öffnen.
2. Panel anzeigen.

### Expected C
- Die neue Beschriftung ist sichtbar; die Signatur hat den Neuaufbau ausgelöst.

### Negative checks
- `_ADMIN` ist nach dem Öffnen weiterhin **ausgeblendet** (`xlSheetVeryHidden`).
- `PID_RunSystemSmokeCheck` bleibt vollständig **PASS**.
- Ein Öffnen ohne jede Änderung markiert die Mappe nicht mehr allein wegen des
  Admin-Blatts als geändert.
