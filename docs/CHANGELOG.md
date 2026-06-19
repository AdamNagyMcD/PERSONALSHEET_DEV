# PERSONALSHEET CHANGELOG

Entwicklungs- und Release-Log.

## Unreleased

### Added
- mod_PIDAdminSheet.bas: verstecktes Admin-Panel `_ADMIN` (VeryHidden) mit Buttons (Smoke, Full Refresh, VBA Import, Schutz AN/AUS, …); `PID_ToggleAdminSheet`.
- Modul1.bas (FP-029): `PID_GetUrlaubGeldFormulaR1C1`, `PID_RestoreUrlaubGeldFormulasSilent`, `PID_RestoreUrlaubGeldFormulas` — kanonische Spalte-K-Formel mit B/C-Guard, J-Guard und 0→leer.
- mod_PIDCopyright.bas (FP-027): `PID_CopyrightAlreadyCurrent` — Skip-Guard in `PID_ApplyCopyrightToAllSheets`; kein Unprotect/Write wenn Copyright schon aktuell.

### Fixed
- mod_BuildFluktuationDaten.bas + mod_FluctuationCalculation.bas (FP-Flukt FIX 2): Neutrale Austrittsgründe (Karenz, Store transfer, Nicht eingetreten, Beförderung) flossen fälschlich in den Zähler der Fluktuationsrate ein und trieben sie nach oben. Neue zentrale Wahrheit `PID_IsNeutralFluctuationExitReason` (normalisiert via `ReasonKey`, ignoriert Groß-/Kleinschreibung, Umlaute und Leerzeichen-Unterschiede — „Store transfer" == „Storetransfer"). `GetFluctuationCategory` nutzt jetzt diesen Helper für die „Neutrale Bewegung"-Klassifizierung (keine doppelte Reason-Text-Logik mehr). Neutrale Austritte werden an BEIDEN Zähler-Stellen ausgeschlossen: `PID_CountExitsInPeriod` (Spalte F in FLUKTUATION_DATEN; speist UBERSICHT Q, FLUKTUATION, Quartal, YTD und den Q31-Sync) UND `PID_CalculateFluctuation` (Live-Pfad: zählt beim Eintragen eines Austrittsdatums direkt vom Monatsblatt, Spalte N = Austrittsgrund). Damit liefert Q31 sofort 0,00 %, wenn der einzige Austritt im Monat neutral ist — identisch zu UBERSICHT/FLUKTUATION. Alle Raten laufen über `PID_ComputeFluctuationForPeriod` → `PID_CountExitsInPeriod` bzw. den konsistenten Live-Pfad.
- DieseArbeitsmappe.cls (FP-Flukt FIX 2): Eine reine Änderung des Austrittsgrundes (Spalte N) löste die unmittelbare Q31-Neuberechnung nicht aus (der Live-Trigger hörte nur auf Spalte I). Da der Grund entscheidet, ob ein Austritt in die Rate zählt, hört der unmittelbare `PID_CalculateFluctuation`-Trigger jetzt auf `I3:I82,N3:N82`. Das Dirty-Flag für die volle Analyse wurde bereits bei D/I/N gesetzt (unverändert). TEST 24. Anzeige unverändert: neutrale Austritte bleiben als „Neutrale Bewegung" sichtbar; der Nenner (Durchschnittsbestand) bleibt unverändert. Neue Tests: TEST 20–23.
- mod_FluctuationCalculation.bas + DieseArbeitsmappe.cls (FP-Flukt): `Q31` zeigte auf einzelnen Monatsblättern 0 %, obwohl UBERSICHT/FLUKTUATION korrekt waren (UserInterfaceOnly-Schutz wird beim Speichern nicht persistiert → lautloser Schreibfehler in das gesperrte `Q31`). Neue robuste Sync-Helfer `PID_SyncMonthSheetFluctuationQ31` (Schutz aufheben → O31-Label „Fluktuation:" + `Q31` aus `PID_ComputeFluctuationForPeriod` schreiben → kanonisch reprotecten) und `PID_SyncAllMonthSheetsFluctuationQ31`. `PID_SyncFluktuationToMonthSheets` nutzt jetzt diesen robusten Pfad (Werte identisch zu UBERSICHT/FLUKTUATION). Sync wird beim Öffnen deferred (nach `ScreenUpdating=True`, Events temporär aus) und beim Aktivieren je Monatsblatt ausgelöst. Keine Änderung an CopyData, Stundenlogik, Passwortlogik oder Mac-Schutzlogik.
- mod_CopyData.bas (FP-030): Mehrfache Stunden-Aenderungen blieben am ersten Wert haengen. `PID_LogEFAenderungForSheet` rief `PID_ClearHourOverrideLogAfterMonth` auf, das ALLE spaeteren Log-Eintraege desselben MA/Feldes loeschte. Da CopyData-Propagation keine Log-Eintraege erzeugt, war jeder spaetere Eintrag eine echte Benutzer-Aenderung — und wurde beim Editieren eines frueheren Monats vernichtet. Aufruf entfernt; Funktion als DEPRECATED markiert. Redundanz wird weiterhin durch Upsert (gleicher Monat), Prune (== Quellwert) und ApplyLogged-Dedup (== laufender Wert) behandelt. Win + Mac. Mac-Schutz (`gCopyDataRunning`, SelectionChange-Backup, `PID_FlushPendingEFLog`, `PID_ReconcileUnloggedFChangesForMac`) unveraendert.
- mod_CopyData.bas (FP-030): `PID_ShowHourOverrideLog` — schreibgeschuetzte Diagnose, zeigt gespeicherte Monats-Overrides (vor/nach CopyData ausfuehren).
- DieseArbeitsmappe.cls (FP-027): `PID_ConfigureDeferredWorkbookCalculationOnOpen` wird erst nach `ScreenUpdating=True` aufgerufen — Blatt erscheint sofort, „Berechnet…" laeuft danach sichtbar statt das Open zu blockieren.
- Modul1.bas / mod_CopyData.bas (FP-029): Spalte K zeigte 0,00 € in leeren MA-Zeilen; neue kanonische Formel liefert leer bei fehlendem MA oder J=leer; CopyData propagiert Formel (nicht mehr vom Quellblatt kopiert).
- mod_KVLohnLookup.bas + mod_SmokeCheck.bas: TEST 15 erkennt VBA-geschriebene G-Werte (nicht nur G3-Formel); `PID_EnsureMonatslohnFormulas` prueft alle 12 Monate.
- mod_CopyData.bas + DieseArbeitsmappe.cls (FP-028): E/F-Overrides nur aus User-Log; bei Aenderung Monat M Log > M loeschen; redundante Log-Zeilen und Reconcile nach CopyData entfernt.
- mod_CopyData.bas (FP-028): `PID_GetOrCreateHourOverrideLogSheet` hebt Blattschutz auf — verhindert lautlose Schreibfehler nach Workbook-Reopen (UserInterfaceOnly wird nicht gespeichert).
- mod_CopyData.bas (FP-028): Ein bewusster spaeterer Stunden-Override blieb nicht erhalten, wenn sein Wert zufaellig dem Quellwert eines frueheren Monats entsprach (z.B. April F=173 → Juli F=69 → November F=173; nach CopyData ab April fiel November faelschlich auf 69 zurueck). Ursache: `PID_PruneHourOverrideLogForCopy` loeschte und `PID_ApplyLoggedHourOverrides` uebersprang Log-Eintraege per Vergleich gegen den **Quellwert**. Beide vergleichen jetzt gegen den **laufenden Segmentwert** (running value, in Monatsreihenfolge fortgeschrieben, Seed = Quellwert): redundant ist nur, was dem bereits wirksamen Wert entspricht — eine echte spaetere Aenderung bleibt erhalten. Der erste Eintrag eines MA/Feldes vergleicht sich weiterhin mit dem Quellwert (identisches Verhalten zu vorher), spaetere Eintraege mit dem fortgeschriebenen Wert. FP-030-Verhalten (TEST 9–11) und Mac-Schutz (`gCopyDataRunning`, SelectionChange-Backup, `PID_FlushPendingEFLog`, `PID_ReconcileUnloggedFChangesForMac`) unveraendert. Neuer Regressionstest: TEST 19.
- mod_CopyData.bas (FP-028): `PID_PruneHourOverrideLogForCopy` verarbeitet Monate jetzt aufsteigend und loescht markierte Zeilen erst nach dem Durchlauf (Index-sicher); `PID_ApplyLoggedHourOverrides` ohne Quellwert-Skip — Deduplizierung ausschliesslich gegen laufenden Segmentwert.
- mod_PIDAdminSheet.bas: 14. Admin-Button (Stunden-Log) fehlte; `ws.Move` schlug lautlos fehl wenn `_ADMIN` schon an Position 1 war — `PID_AdminMoveSheetToFront` prueft Index vorher.
- mod_PIDAdminSheet.bas: Button-Spezifikation auf `Select Case` umgestellt (kein VBA-Array-Rueckgabe-Limit mehr).
- mod_KVStundenDropdown.bas + DieseArbeitsmappe.cls: KV-Gruppe (E) wechseln baut F-Dropdown sofort neu (scoped pro Zeile); Validierung muss zur aktuellen Gruppe passen (BG1_5→BG1), nicht nur „irgendein“ KV_DG_*.
- DieseArbeitsmappe.cls: E-Dropdown-Auswahl — F-Refresh beim Verlassen von E und beim Sprung E→F; ganze Monatsblatt-F-Spalte neu (x14-Gruppen-Validierung).
- mod_KVStundenDropdown.bas + Modul1.bas + mod_AddNewKVPeriodOnTop.bas: Mac-only sofortiger Dropdown-Refresh nach Eigene Stunden; Windows lazy-Refresh unveraendert.
- mod_KVStundenDropdown.bas: Mac — auch nach Eigene Stunden und bei SheetActivate immer `forceFullRebuild` (x14-Validierung sonst veraltete Helper-Zellrange trotz neuer Stunde).
- mod_KVStundenDropdown.bas: `forceFullRebuild` wendet F-Validierung immer neu an (nicht ueberspringen wenn alte x14-Liste noch „gueltig“ wirkt); aggressives Loeschen pro Zelle; Named Range KV_DG_* wieder auf Mac.
- Modul1.bas + mod_KVStundenDropdown.bas + mod_AddNewKVPeriodOnTop.bas: Nach Eigene Stunden / Stunde Loeschen bleibt LOHNTABELLE sichtbar (Scroll + Zeile G); Mac-Monats-Refresh springt nicht mehr auf Dezember.
- mod_KVStundenDropdown.bas + DieseArbeitsmappe.cls: F-Dropdown-Gueltigkeit exakter Named-Range-Vergleich (BG1 vs BG1_5); F-Refresh auch beim Sprung auf F wenn E-Gruppe gewechselt (Mac-Fallback).
- mod_KVStundenDropdown.bas + DieseArbeitsmappe.cls: Mac — kein Stunden-Cache; E-Wechsel per SelectionChange (Dropdown meldet kein SheetChange); F-Refresh bei E verlassen / F fokussieren erzwungen; `PID_ForceRefreshFStundenDropdownForSheet`; Dirty-Clean nach Mac-Period-Refresh.

### Changed
- mod_FluctuationCalculation.bas + mod_BuildFluktuationAnalyse.bas (FP-Flukt): O31-Label auf Monatsblättern wird im Sync jetzt IMMER auf „Fluktuation:" gesetzt (nicht nur wenn leer) und direkt nach Unprotect vor der Berechnung geschrieben (korrigiert auch falsche/alte Werte, bleibt bei Berechnungsfehler korrekt). FLUKTUATION-Spaltenbreiten als neuer Layout-Standard aus der manuell angepassten Optik übernommen: A 16→25.4, E 27→30.8, O neu 10.2 (nur FLUKTUATION-spezifische Konstanten; keine globalen Style-Helfer, Zeilenhöhen, Merges, Farben oder Schriften geändert).
- mod_BuildFluktuationAnalyse.bas + mod_FluctuationCalculation.bas + mod_FormatMonthSheet.bas (FP-Flukt): Verständlichkeit FLUKTUATION/Monatsblätter. Monatsblatt `O31` zeigt persistent das Label „Fluktuation:" (im Format-Makro gesetzt, vom Sync zusätzlich abgesichert). FLUKTUATION erhält unter „Kurz erklärt" einen additiven Block „Bewertung & Ziel": Zielwert „unter 20 % Jahresfluktuation", Legende der 6 Bewertungsstufen (gelten primär für Jahr/YTD) und einen Hinweis zu Monatswerten (einzelner Austritt ≈ 1–2 %/Monat, Führungsbewertung über YTD). KPI-Label „Verlust-Score" → „Verlust-Score (nur Info)"; Erklärtext stellt klar, dass der Score die Fluktuationsbewertung NICHT beeinflusst. Management-Kurzbewertung (`C5`) ergänzt um „Ziel: …" und „Handlungsbedarf: …" (neuer Helper `PID_GetFluctuationActionHint`). Block ist rein additiv (ganz unten), bestehende Sektionen/Spalten und Zielzellen unverändert. NICHT umgesetzt (nur dokumentiert): KPI „Aktueller Personalbestand: xx MA" — könnte später über `PID_CountEmployeesAtDate(ws, Periodenende)` als reiner Info-KPI ergänzt werden.
- mod_BuildFluktuationAnalyse.bas + mod_FluctuationCalculation.bas (FP-Flukt): EINE einzige Bewertungslogik. FLUKTUATION-Status (B5) und Bewertung (B10) kommen jetzt ausschliesslich aus der YTD-Fluktuationsrate; das alte Verlust-Score-basierte Statussystem (`GetFluktuationRiskLevel`, „Kritisch") treibt nichts mehr und ist DEPRECATED. Neuer Helper `PID_GetFluctuationStatusShort` liefert die Kurz-Ampel (Sehr gut/Gut/Erhöht/Hoch/Kritisch/Extrem) aus derselben Rate wie die Langform `PID_GetFluctuationRating`. `ApplyRiskFormatting` auf die 6 Rate-Stufen (grün→rot) umgestellt; `GetFluktuationManagerSummary` rate-basiert neu formuliert. Kein „Daten prüfen"-Status-Override mehr — unvollständige Daten erscheinen nur als informativer Hinweis (KPI „Daten offen" + „Sofort prüfen" bleiben sichtbar). Verlust-Score / „Ø Verlust-Score" / „Kritische Austritte" bleiben rein informative KPIs. „Sofort prüfen" und „Empfehlungen" bleiben operative Hinweise. Zielzellen `Q31`, `UBERSICHT` Spalte Q und FLUKTUATION-Layout unverändert.
- mod_FluctuationCalculation.bas + mod_BuildFluktuationAnalyse.bas (FP-Flukt): Fluktuation einheitlich nach HR-Controlling-/BDA-Standard. Rate = Austritte im Zeitraum / durchschnittlicher Personalbestand; Durchschnittsbestand = (Bestand Periodenanfang + Bestand Periodenende)/2. Monat, Quartal und YTD nutzen dieselbe Logik (`PID_ComputeFluctuationForPeriod`); Quartal/YTD werden direkt aus Gesamtaustritten + Durchschnittsbestand berechnet (nicht mehr aus Monatsprozenten bzw. Mittel der Monatsendbestände). Neue zentrale Helper: `PID_CountEmployeesAtDate`, `PID_CountExitsInPeriod`, `PID_GetAverageHeadcountForPeriod`, `PID_ComputeFluctuationForPeriod`, `PID_GetFluctuationRating`. Live-`Q31` (`PID_CalculateFluctuation`) auf denselben (Anfang+Ende)/2-Nenner umgestellt (Zähler bleibt sofortige Blatt-Zählung). Neu: HR-Bewertungstext neben KPI „Jahresfluktuation" (Grenzwerte 0,2/0,35/0,5/0,7/1,0). FLUKTUATION-Monatstabelle: Spalte „Durchschnitt" → „Ø Verlust-Score" (Klarstellung, kein Layout-Umbau). Obsoletes `PID_ComputeFluktuationForMonthRange` entfernt. Zielzellen `Q31`, `UBERSICHT` Spalte Q und FLUKTUATION-Layout unverändert. Excel 2016 + Win/Mac, keine dynamischen Array-Funktionen, kein XVERWEIS.
- mod_KVStundenDropdown.bas (FP-005): scoped KV-dirty-Refresh — nur betroffene KV-Codes/Zeilen statt 80× F-Rebuild; `MarkKVDropdownDirtyForKVCode`, `MarkKVDropdownDirtyFromLOHNTABELLERange`. Windows-Messung: scoped 0,15 s vs Voll 0,52 s (Baseline 0,51 s).
- mod_AddNewKVPeriodOnTop.bas: Eigene Stunden hinzufuegen/loeschen markiert nur den gewaehlten KV-Code dirty.
- DieseArbeitsmappe.cls: LOHNTABELLE D4:G-Aenderung → KV-Code aus geaenderten Zeilen.
- mod_PerformanceBaseline.bas: Baseline-Schritt 2b (scoped dirty BG1).
- mod_KVStundenDropdown.bas (FP-006): `RefreshKVStundenDropdownForRow` nutzt `KV_DG_*` pro KV-Code; `PID_RemoveLegacyKVDDNamedRanges` bei `RefreshAllMonthKVStundenDropdowns`. Verifiziert: `PID_CountKVDDNamedRanges` = 0.
- mod_KVLohnLookup.bas + Modul1.bas + DieseArbeitsmappe.cls (FP-007): SheetChange — kein doppeltes L-Recalc bei E/F; gebündeltes H/L-Calculate.
- DieseArbeitsmappe.cls + mod_KVStundenDropdown.bas (FP-008): SelectionChange — ScreenUpdating nur bei Dropdown-Repair; E-Validation-Cache.
- mod_RefreshFluktuationAll.bas + mod_BuildFluktuationDaten.bas + DieseArbeitsmappe.cls (FP-009): Save nur Daten-Refresh; Analyse beim FLUKTUATION-Tab; inkrementeller Monats-Rescan.
- mod_SchutzHinzufugen.bas + Monatsblatt-Protect-Aufrufer (FP-011/012): `PID_ProtectWorkerMonthSheet` — Fill Handle aus, Sortieren aus.
- mod_SchutzHinzufugen.bas + mod_FormatMonthSheet.bas (FP-013): `PID_ApplyMonthSheetLockPolicy` — Lock-all + Whitelist (B/C, D, E/F, I/J, M/N, O18:Q25); Q12 nur auf Nicht-Startmonaten entsperrt (Jan/Mar–Jul/Sep–Dez).
- mod_FormatMonthSheet.bas: Durchrechnungs-Startmonate (Februar/Mai/August/November) — Panel O15 als ein Merge O15:R15; Makro `PID_FixDurchrechnungStartMonthPanels`.
- mod_SmokeCheck.bas (FP-016): TEST 17 (Month Sheet Lock Policy: G gesperrt, E/B/I entsperrt) + TEST 18 (Q12 Vormonat: Februar gesperrt, Januar entsperrt).
- docs/Kurzanleitung_Personalsheet_A4.html (FP-015): Q12 Vormonat erklärt; G/H/K/L Blattschutz-Hinweis; Sortieren-Verbot; Stand Juni 2026.

### Removed
- mod_BuildFluktuationAnalyse.bas + DieseArbeitsmappe.cls: FLUKTUATION PDF-Export (Button, Makro, Export-Hilfen) entfernt; Legacy-Button wird bei Refresh gelöscht (FP-025 storniert).

### Added
- mod_PIDCopyright.bas: Copyright-Hinweis Zeile 2 rechts (sichtbar beim Tab-Öffnen, ohne PageSetup/Shape); VBA-Modul-Kopf Adam Nagy / McOpCo (FP-023).

### Fixed
- mod_PIDCopyright.bas: Copyright S2 (Monatsblatt) und L2 (LOHNTABELLE) linksbuendig; FLUKTUATION-Copyright nach A19 (alt A3:D3 wird bereinigt).
- mod_FormatMonthSheet.bas: A1 (Monatsindex) Schriftfarbe #DDEBF7 auf allen Monatsblaettern.
- mod_FormatMonthSheet.bas: FormatAllMonthSheets merge-sicher (kein Format-Paste auf verbundene Kopf-/Panel-Zellen; Fehler 1004 behoben).
- mod_FormatMonthSheet.bas + mod_CopyData.bas: Monatsblatt-Kopfzeilen A–N vertikal (Zeile 1+2) zusammengeführt; Spalte A Sorszahlen 1./2./… als Text (@), schmale Breite; A1 Monatsindex bleibt beim Format-Kopieren erhalten.
- mod_PIDCopyright.bas: UEBERSICHT copyright B24 statt B25; kein blindes UnMerge fremder FINANZIELL/DURCHRECHNUNG-Merges bei FullSystemRefresh.
- mod_BuildFluktuationDaten/Analyse.bas: Austritte gesamt = Summe Monatstabelle (Dedup, Jahresfilter, Gesamt-Zeile); PDF-Export-Button auf FLUKTUATION.
- mod_BuildFluktuationAnalyse.bas: PDF Export Mac/Win — GetSaveAsFilename ohne benannte Parameter, MacScript-Fallback.
- mod_BuildFluktuationAnalyse.bas: PDF Export — Druckbereich ohne Hilfsspalten, Mac-Pfad/HFS, PageSetup-Restore, Button nach Export + bei Tab-Activate.
- mod_BuildFluktuationAnalyse.bas: PDF Export Mac — kein GetSaveAsFilename, fest `Fluktuation.pdf` neben xlsm, Temp-Workbook-Export.
- mod_BuildFluktuationAnalyse.bas: PDF Export — FitToPagesTall=0 (mehrseitig statt 1-Seiten-Zusammenpressung), ein Exportlauf, Erfolg per Dateiprüfung, Diagramme mit PrintObject.
- mod_BuildFluktuationAnalyse.bas: PDF Export Mac — Zoom 100% statt FitToPages, Diagramme als Bild, Staging-Datei `.pid_export.pdf` + Überschreiben, Chart-Block in Druckbereich.
- mod_BuildFluktuationAnalyse.bas: PDF Export Mac 2016 — ChartObject.PrintObject statt CopyPicture/Shape.PrintObject (Compile-fix), minimales PageSetup.
- mod_BuildFluktuationAnalyse.bas: PDF Export Mac — direkt nach `Fluktuation.pdf` (kein Temp-Workbook/Staging), weniger Sandbox-Dialoge.
- mod_BuildFluktuationAnalyse.bas: PDF Export — Druckbereich via End(xlUp)/Find statt Rows.Count-Scan (Mac-Performance).
- mod_BuildFluktuationAnalyse.bas: PDF Export — Zoom aus Seitenbreite (1 Seite breit, mehrere Seiten hoch, Mac-tauglich).
- mod_BuildFluktuationAnalyse.bas: PDF Export — Querformat (A4), groessere Schrift via breiterer Seite; Mac-Dialog ohne Ordner-Hinweis.
- mod_BuildFluktuationAnalyse.bas: Monatstabelle Zeilenhoehe dynamisch (Monats-Hinweis), PDF-Druckbereich inkl. Merge-Spalten; Win+Mac.
- mod_ResetAndImportVBAFiles.bas: vor .bas-Import mod_PIDUtils1-Duplikate entfernen (mehrdeutige Public-Funktionen).
- mod_BuildFluktuationAnalyse.bas: PDF Monats-Hinweis Merge aufloesen, Spaltenbreite M/N, Zoom aus Spaltenpunkten, Temp-Export.
- mod_BuildFluktuationAnalyse.bas: FLUKTUATION Spaltenbreiten A-N als Konstanten (manuell abgestimmt), PID_ApplyFluktuationColumnWidths.
- mod_BuildFluktuationAnalyse.bas: PDF-Export-Button gelb/kontrastreich, Hoehe 16pt (passt in Zeile 2).
- mod_BuildFluktuationAnalyse.bas: PDF-Button Mac — ScreenUpdating=True fuer Position, feste Geometrie, immer neu bei Tab-Activate.
- mod_BuildFluktuationAnalyse.bas: PDF-Button in Titelzeile A1:E1, Mac Form-Control, Fallback Shape, DisplayDrawingObjects.
- mod_BuildFluktuationAnalyse.bas: PDF-Button — nur Shape (Mac-compile-safe), Titelzeile A1:E1 rechts.
- mod_BuildDurchrechnung.bas + mod_SumMergedCells.bas: UEBERSICHT FINANZIELL — Formeln F/I/H/K und Quartale wiederhergestellt; G/J aus Monatsblaettern per Sync.
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE — kein Erfolgs-Dialog nach Periode loeschen; alle KV-Datenzeilen Schriftfarbe Navy (auch alte Perioden beim Oeffnen).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE — kein Erfolgs-Dialog nach „Neue Periode“ (wie Eigene Stunden / Stunde loeschen).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE — einheitliche Zeilenhoehe (24 pt) und Spaltenbreiten B/C/E/F nach neuer KV-Periode; `ClearContents` statt `Clear` beim Einfuegen.
- mod_KVStundenDropdown.bas + Modul1.bas + DieseArbeitsmappe.cls: erster Monats-Tab nach KV-Änderung schneller — Bulk-F-Dropdown, Manual während Refresh, H/K/L einmal (FP-018).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE — kein Erfolgs-Dialog nach Eigene Stunden / Stunde löschen; Monatslohn-Hinweis ohne Spaltenbuchstaben; Umlaute in Dialogen (FP-022).
- mod_AddNewKVPeriodOnTop.bas: K-Marker (`PID_EIGEN`) bleibt erhalten — Trim loescht nur L:XFD, nicht K; Markierung nach FormatKVPeriodArea erneut (FP-022).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE Button 4 „Stunde löschen“ — nur markierte/extra Stunden in Liste, kein Fallback auf alle Standard-Stunden (FP-022).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE Button 4 „Stunde löschen“ — Erkennung per K-Marker, ältere Periode/min. Block; Legacy `PID_MarkSelectedLOHNTABELLECustomHour` (FP-022).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE Button 4 „Stunde löschen“ statt Hilfe; `DeleteCustomKVMonatsstunden` mit Bestätigung und Dropdown-Refresh (FP-022).
- mod_FormatMonthSheet.bas + Modul1.bas: Monatsblätter D/I Spaltenbreite 13 für Datumsfelder (FP-017).
- mod_SchutzHinzufugen.bas: Compile-Fix Excel 2016 Mac — Protect ohne AllowSelectingLockedCells; EnableSelection numerisch (FP-021).
- mod_SchutzHinzufugen.bas: UEBERSICHT geschuetzt, nur E30/I30 editierbar (AllowEditRanges, Lock-all) (FP-021 v4).
- DieseArbeitsmappe.cls + mod_SchutzHinzufugen.bas: Compile-Fix UEBERSICHT-Erkennung (`PID_IsUbersichtWorksheet`).
- mod_SchutzHinzufugen.bas + mod_SumMergedCells.bas + mod_BuildDurchrechnung.bas: FINANZ-Sync nutzt zentrales `PID_ReprotectWorksheet` (FP-021).
- mod_SchutzHinzufugen.bas: Q12:R12 (Vormonat +/-) entsperrt; `PID_ReprotectWorksheet` für konsistenten Schutz (FP-020).
- Modul1.bas + DieseArbeitsmappe.cls: nach Öffnen wieder Automatische Berechnung; H/K/L-Formeln auf Monatsblättern per Tab-Wechsel neu berechnet (FP-024).
- mod_CopyData.bas: CopyData startet ohne Ja/Nein-Bestätigung vor dem Kopieren (FP-019).

### Changed
- docs/Kurzanleitung_Personalsheet_A4.html: CopyData ohne Bestätigungsdialog beschrieben.
- docs/RELEASE.md: erste Release (v1.0) an FP-017–FP-022 gekoppelt; FP-023 Copyright optional.
- docs/FUTURE_PLANS.md: FP-023 Adam Nagy / McOpCo — Blatt-Fußzeile + VBA-Modul-Copyright-Kopf.
- docs/FUTURE_PLANS.md: Excel-16-Test-Feedback als FP-017–FP-022 (Spalten D/I, KV-Ladezeit, CopyData ohne Dialog, Q12/E30/I30 Unlock, LOHNTABELLE Eigene-Stunden löschen).
- docs/Kurzanleitung_Personalsheet_A4.html: vereinfachte Anleitung für Einsteiger (Umlaute ä/ö/ü/ß), drei Grundregeln, Schritt-für-Schritt CopyData/Austritt, kein Ziehen in E/F, FormatAllMonthSheets-Hinweis.
- docs/FUTURE_PLANS.md: Schutz-Paket (Amateur-Vermeidung) — Übersicht + FP-011–FP-016 (Fill Handle, Sort off, Lock-all/Whitelist, EnableSelection, Endanwender-Doku, Schutz-Smoke).
- docs/FUTURE_PLANS.md: Performance-Backlog FP-005–FP-010 (Windows-Referenzmessung, F-Dropdown/Names/SheetChange/Fluktuation).
- mod_FormatMonthSheet.bas: Referenz-Layout aus manuell formatiertem Januar (OOXML-Analyse); FormatAllMonthSheets kopiert Formate von Januar.
- mod_PIDSheetStyle.bas: PID_StyleApplyToolbarButton fuer einheitliche Toolbar-Buttons (Navy/Header/Accent/Zebra).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE an UEBERSICHT-Palette angeglichen (Titel, Header, Periodenband, Datenzeilen, Toolbar).
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE Zeile 2 nur Buttons nebeneinander (kein A2-Text); verwaiste Toolbar-Shapes (z.B. altes Hilfe) werden entfernt.
- mod_PIDAdmin.bas: Admin-/Release-Makros (FullSystemRefresh, QuickSystemCheck) und PID_ConfirmAdminAction; Bestaetigung vor ResetAndImportVBAFiles und UnprotectEverything.
- mod_PIDUtils.bas: gemeinsame Helper (PID_CollectionHasKey, PID_MonthNames, PID_GetMonthIndexFromName/SheetName); Duplikate in CopyData, DataClear, SumMergedCells, Durchrechnung, Modul1, DieseArbeitsmappe, KV-Module entfernt.
- mod_CopyData.bas / mod_DataClear.bas: CopyData/DataClear nur auf gueltigem Monatsblatt; Bestaetigung mit Blattname und Zielmonaten vor CopyData.
- mod_SmokeCheck.bas: TEST 9-16 (Jahr, A1-Index, UEBERSICHT, Durchrechnung E30/I30, LOHNTABELLE, KV_CODE_LIST, Monatslohn, VBProject); TEST 8 Mac-Pfad fix.
- Repo: OOXML probe mappak (_xlsm_*, _einstellung_probe) kivetele git trackingbol; .gitignore bovitve.
- README.md + docs/RELEASE.md hozzaadva (workflow, release checklist, admin makro figyelmeztetes).

### Fixed
- mod_KVStundenDropdown.bas: F-Stunden-Dropdown nach LOHNTABELLE-Aenderung bei dirty-Refresh Validation neu (FP-004, kein E-Re-Select noetig).
- mod_SumMergedCells.bas: FINANZIELL-Sofort-Sync nur noch bei Q17:R29 und S35; O18:Q25-Freitext loest kein Panel-Flackern mehr aus (FP-001).
- Modul1.bas + mod_FormatMonthSheet.bas: Spalte L Restore ohne Format-Kopie (PasteSpecial Formeln); Mitarbeiterblock-Rahmen danach wiederherstellen.
- mod_FormatMonthSheet.bas + mod_CopyData.bas: Zebra-Hintergrund nach CopyData/Sort auf Zielmonaten wiederherstellen.
- mod_CopyData.bas: fehlende Konstante `PID_HOUR_OVERRIDE_LOG_SHEET` wiederhergestellt (Compile-Fehler vor CopyData).
- mod_CopyData.bas: O18:Q25 Panel-Kopie merge-anker (O/Q) + O18:R25 Block; Panel vor RestoreFormulas.
- mod_FormatMonthSheet.bas: CopyData-Button bei jedem Monats-Tab neu positionieren statt loeschen/neu anlegen (Mac/Februar-Fix).
- mod_PIDSheetStyle.bas: Toolbar-Button-TextFrame per Late Binding (Mac: kein Compile-Fehler auf WordWrap).
- mod_AddNewKVPeriodOnTop.bas + DieseArbeitsmappe.cls: LOHNTABELLE Status/Pruefung (I/J) explizit berechnen bei Manual Calculation.
- mod_ResetAndImportVBAFiles.bas: `FixLegacyModul11Name` renames legacy `Modul11` to `mod_BuildDurchrechnung` (VBA max 31 chars; old name was 32).
- mod_BuildDurchrechnung.bas: module/file renamed from `mod_BuildDurchrechnungUebersicht` (name exceeded VBA 31-char limit).
- mod_BuildDurchrechnung.bas: fix error 5 — merge before CF, safe borders, ISNUMBER-based diff/ueber CF.
- mod_BuildDurchrechnung.bas: Anzeigetexte mit Umlauten via ChrW (Verfuegbar, Ueberstunden, Jaenner usw.); Blattnamen bleiben ASCII.
- mod_BuildDurchrechnung.bas: Tabellenzellen zentriert; Eingabezeile 30 bereinigt (G30-Altlast, J:Q-Merge).
- mod_BuildDurchrechnung.bas: Lohn/h (G) und EUR (I) mit Euro-Zeichen via PID_ApplyEuroNumberFormat.
- mod_ResetAndImportVBAFiles.bas: VBA-Import Erfolgsmeldung gekuerzt.
- mod_BuildDurchrechnung.bas: FINANZIELL-Block (B2:Q23) im blau/gelb Design; B2/B23 #BEZUG! -> EINSTELLUNG!C35.
- mod_FluctuationCalculation.bas: gemeinsame Fluktuationslogik; UEBERSICHT Q + Monatsblatt Q31 sync mit FLUKTUATION.
- mod_SumMergedCells.bas: SumMergedCells volatile + S36-Neuberechnung auf allen Monatsblaettern (kein manuelles Enter).
- mod_SumMergedCells.bas: PID_RecalculateFinanzSummaryChain — Monats-S35:S37, EINSTELLUNG E22:E33, UEBERSICHT G/J/H/K (FINANZIELL null-Werte behoben).
- mod_SumMergedCells.bas + DieseArbeitsmappe.cls: FINANZIELL G/J sofort nach Lohn-/Crew-Labor-Aenderung auf Monatsblaettern (SheetChange).
- mod_SumMergedCells.bas: FINANZIELL G/J wie Fluktuation Q — direkte Value-Sync via SumMergedCells statt stale Cross-Sheet-Formeln; kurzes Unprotect beim Schreiben.
- mod_SumMergedCells.bas: gFinanzSummaryDirty + RefreshFinanzSummaryIfDirty; UEBERSICHT-Activate nur bei Bedarf; Batch-Unprotect, VBA-Quartals-/Diff-Berechnung.
- mod_KVLohnLookup.bas: LOHNTABELLE-Cache in PID_KVLohnLookup UDF (kein Invalidate pro Zelle); Cache-Clear bei MarkAllKVLohnDirty.
- mod_KVLohnLookup.bas + DieseArbeitsmappe.cls: G (Monatslohn) Calculate bei E/F-Aenderung (UDF+Cache Recalc-Fix).
- DieseArbeitsmappe.cls: kein MarkAllKVDropdownsDirty mehr bei Open; PID_ResetMonthView ohne Activate/Select.
- Modul1.bas + DieseArbeitsmappe.cls: kein CalculateFull mehr beim Oeffnen; Letztes-Gehalt nur bei fehlenden Formeln; E-Validierung lazy (SelectionChange) statt 24x Rebuild bei Open.
- mod_SchutzHinzufugen.bas: kein E-Dropdown-Rebuild mehr im Protection-Setup (nur E/F entsperren).
- mod_KVStundenDropdown.bas: KV-Code-Liste nur anlegen wenn Named Range fehlt.
- mod_KVLohnLookup.bas: G-Recalc via Formel-Reset statt Calculate (UDF-Cache-Fix); RefreshKVLohnIfDirty berechnet G wirklich neu.
- mod_SchutzHinzufugen.bas: lazy Blattschutz beim Open (nur Hidden + Active), rest bei erstem Tab-Wechsel.
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE-Buttons nur anlegen wenn fehlend (nicht jedes Open neu).
- DieseArbeitsmappe.cls: LOHNTABELLE-Setup deferred auf Tab-Activate; Calculation=Manual waehrend Open.
- mod_KVStundenDropdown.bas: F-Dropdown bei E-Aenderung nur Helper-Werte aktualisieren wenn Validation schon OK.
- mod_SumMergedCells.bas + DieseArbeitsmappe.cls: FINANZIELL-Sync bei E/F deferred (MarkDirty), nicht sofort bei jeder KV-Aenderung.
- mod_KVStundenDropdown.bas: F-Dropdown lazy bei E-Aenderung (Invalidate), Rebuild erst bei F-Klick; Stunden-Lookup-Cache.
- Modul1.bas + DieseArbeitsmappe.cls: EnableCalculation nur fuer aktiven Tab beim Open (kein 12-Blatt-UDF-Recalc); kein MarkFinanzSummaryDirty beim Open.
- Modul1.bas: FormatStaleValues=False (kein Strikethrough bei Manual/Partial Calc); Manual nach Open ohne Sheet-Calculate.
- mod_KVLohnLookup.bas: G (Lohn) direkt via GetKVLohnByPeriod statt UDF-Formel-Reset bei E/F-Aenderung.
- mod_KVStundenDropdown.bas: RefreshKVStundenDropdownForSingleRow fuer schnellen F-Dropdown bei F-Klick.
- DieseArbeitsmappe.cls: Formula-Checks lazy beim ersten Monats-Tab; Hour-Override-Log nur bei F-Aenderung.
- Modul1.bas + DieseArbeitsmappe.cls: L (Letztes Gehalt) Calculate bei D/E/F/G/I/K-Aenderung (Manual-Calc-Fix).
- Modul1.bas: L via Application.Evaluate statt Calculate/Formel-Reset; L nur einmal pro G-Update.
- Modul1.bas: L-Recalc nur .Calculate (Formel bleibt); AutoFill-Reparatur bei statischen L-Werten.
- mod_SumMergedCells.bas + mod_DataClear.bas: FINANZIELL-Sync nach Monats-Loeschen nur betroffener Monat (kein 12x S37-Recalc).
- mod_KVLohnLookup.bas + DieseArbeitsmappe.cls: LOHNTABELLE Monatslohn (H) ohne F-Dropdown-Rebuild; L batch Calculate nach KV-Refresh.
- mod_KVLohnLookup.bas: MarkAllKVLohnDirty setzt auch MarkFinanzSummaryDirty (UEBERSICHT G/J nach KV-Aenderung).
- DieseArbeitsmappe.cls: SelectionChange nur D/E/F (+ O45-Bereinigung), sonst sofortiger Exit.
- mod_KVStundenDropdown.bas: F-Dropdown-Refresh nur Zeilen mit KV-Code (E nicht leer); ScreenUpdating aus waehrend Refresh.
- DieseArbeitsmappe.cls: FLUKTUATION-Activate nur bei gFluktuationDirty; kein S35-S37-Calculate bei Monats-Activate.
- mod_FluctuationCalculation.bas: UEBERSICHT nach Fluktuation-Sync wieder schuetzen.
- mod_DataClear.bas: MarkFinanzSummaryDirty nach Monatsdaten-Loeschen.
- Modul1.bas: FullSystemRefresh ohne doppeltes Protection-Setup am Ende.
- DieseArbeitsmappe.cls: `Union`-Aufruf in `EnforcePasteValuesOnly` ohne Zeilenfortsetzung (Mac VBA Syntaxfehler behoben).
- mod_ResetAndImportVBAFiles.bas: ueberarbeiteter Import (Workbook-/Tabellenmodule per Code-Update statt Import); `ReadVBAFileWithoutAttributes` liest nur ab `Option Explicit` (kein VERSION/Attribute-Header mehr im Modul).
- Modul1.bas: `SyncDieseArbeitsmappeFromExport` macro alias for manual workbook-module sync.
- Personalsheet.xlsm: reverted broken direct XML patch on `UBERSICHT`; Durchrechnung block must be created via `BuildDurchrechnungUebersicht` macro (avoids corrupted `sharedStrings` counts on Mac Excel).
- mod_FormatEinstellung.bas: Compile-Fix — kein `ws.DisplayGridlines` (nur Window-Eigenschaft).

### Changed
- mod_PIDSheetStyle.bas: gemeinsame Styles wieder am Original-UEBERSICHT FINANZIELL (Navy-Titel 13pt, hellblaue fette Header 10pt, Gelb nur Akzent/Summen/Eingaben).
### Changed
- mod_FormatEinstellung.bas: alle nicht editierbaren Bereiche gelb (Monatsspalten B/H/K/N, E6:E17, D22:E33); Wertspalten weiss.
- mod_BuildFluktuationAnalyse.bas: FLUKTUATION gelb = nur Anzeige (Monatsspalte + berechnete Werte), wie EINSTELLUNG.
- mod_BuildDurchrechnung.bas: UEBERSICHT Jänner-Plan E30/I30 weiss und editierbar; gelbe Labels daneben.
- mod_FormatEinstellung.bas: EINSTELLUNG-Blatt im UEBERSICHT-Design (Navy-Titelbaender, Header, Zebra, Rahmen); `FormatEinstellung`-Makro; C35 (Arbeitsjahr) gelb hervorgehoben.
- mod_BuildFluktuationAnalyse.bas: FLUKTUATION-Blatt im gleichen Design wie UEBERSICHT FINANZIELL (Navy-Titel, hellblaue Header, Zebra-Zeilen, Akzent-Gelb, Tabellenrahmen); Risiko-Farben bleiben erhalten.
- mod_BuildDurchrechnungUebersicht.bas: Durchrechnung block clearer for managers; new column `AVG Lohn/h` from Schlussmonat `Q42`; Ueberstunden EUR = Std x Lohn/h x 1,5 (removed manual C30 Stundenlohn); only Jaenner plan inputs (E30/G30) stay yellow; full block styling (title/header colors, CF on Differenz/Status/Ueberstunden, `#,##0.00` formats); `FormatDurchrechnungUebersicht` macro; Mac-safe formatting unmerges block first, formats, then re-merges display rows (fixes error 1004).
- mod_BuildDurchrechnungUebersicht.bas: Ueberstunden EUR no longer requires ISNUMBER on C30 (text numbers like `1` or `12,5` now calculate).
- mod_KVLohnLookup.bas: column `G` uses `PID_KVLohnLookup` UDF (same logic as VBA lookup; fixes `BG3_15` etc.).
- mod_KVStundenDropdown.bas: column `E` dropdown uses named range `PID_KV_CODE_LIST` (fixes `#REF!` and German list separator); applied on every sheet protect/open.
- mod_KVStundenDropdown.bas + DieseArbeitsmappe.cls: safe read of broken `#REF!` validation (no debugger break); direct helper-sheet list reference; silent restore on every open.
- Modul1.bas / mod_KVStundenDropdown.bas / mod_KVLohnLookup.bas: removed duplicate `Restore*` macro aliases (fixes VBA compile error and missing Alt+F8 entries).
- mod_KVStundenDropdown.bas: moved `PID_KV_CODE_*` constants to module top (fixes Mac VBA “Variable not defined” compile error).
- mod_KVLohnLookup.bas: `PID_RestoreMonatslohnFormulasSilent` made Public (fixes Modul1 compile error in `FullSystemRefresh`).
- Modul1.bas + DieseArbeitsmappe.cls + mod_CopyData.bas: auto-restore column `L` Letztes Gehalt formulas on open; CopyData uses canonical L formula (not broken source copy).
- Personalsheet.xlsm: column `L` `#REF!` year refs replaced with `EINSTELLUNG!$C$35` on all month sheets (re-applied after accidental revert).
- mod_ResetAndImportVBAFiles.bas: removed auto-repair during import (caused compile/state issues); clear post-import steps via `FullSystemRefresh`.
- Modul1.bas: `FullSystemRefresh` runs `CalculateFull` after formula restore (fixes G `#NAME?` when VBA UDF was inactive).
- mod_KVStundenDropdown.bas + mod_SchutzHinzufugen.bas: unlock column `F` on protected sheets; broken F validation auto-repair on cell select; `RestoreKVStundenDropdownValidation` macro.
- Personalsheet.xlsm + DieseArbeitsmappe.cls: column `L` `#REF!` year refs fixed again; `PID_EnsureLetztesGehaltFormulas` restored on workbook open.
- Modul1.bas: column `L` restore replaces legacy `#REF!` in existing A1 formulas (AutoFill) instead of silent R1C1 overwrite; `CalculateFull` after restore.
- mod_KVLohnLookup.bas + DieseArbeitsmappe.cls: removed SheetChange VBA writes to `G` (they destroyed formulas and made lohn refresh appear dead).
- Modul1.bas: `RefreshDurchrechnungUebersicht` macro alias.

### Added
- mod_BuildDurchrechnungUebersicht.bas: Durchrechnungszeitraum block on `UBERSICHT` (rows 28+) with period sums, Schlussmonat differenz, Ueberstunden EUR estimate, and manual Jaenner plan inputs (C30/E30/G30).
- Modul1.bas: `BuildDurchrechnungUebersicht` macro alias to rebuild the UEBERSICHT Durchrechnung block safely without touching rows 2-27.
- docs/Kurzanleitung_Personalsheet_A4.html: printable A4 user guide (German) for restaurant managers.
- Modul1.bas: `RestoreAktuelleStundenFormulas` / `PID_RestoreAktuelleStundenFormulas` restores column `H` pro-rata hour formulas on all month sheets using `EINSTELLUNG!C35` as workbook year.
- Modul1.bas: `RestoreLetztesGehaltFormulas` / `PID_RestoreLetztesGehaltFormulas` restores column `L` on all month sheets (fixes `#REF!` year refs → `EINSTELLUNG!C35`; enables `AVG Bruttolohn` / Q42 again).
- Modul1.bas: `RestoreAustrittsdatumValidation` fixes month-sheet `AB1:AB2` date bounds and column `I` data validation after workbook year moved to `EINSTELLUNG!C35`.
- mod_AddNewKVPeriodOnTop.bas: `AddCustomKVMonatsstunden` with dialog flow and green sheet button to insert custom Monatsstunden into a selected KV-Code block, sorted ascending by hours.
- Initial project structure
- SPEC.md
- TEST_CASES.md
- VBA export folder
- New smoke-check macro module `mod_SmokeCheck.bas` with `PID_RunSystemSmokeCheck` to log TEST_CASES 1-8 as PASS/FAIL/REVIEW on `SYSTEM_CHECK`.
- New smoke-check helper macros: `PID_FilterSmokeReviewOnly` and `PID_ClearSmokeFilter` for quick REVIEW-focused triage.
- Added `.gitignore` entry for `.DS_Store` to prevent accidental macOS metadata commits.

### Changed
- mod_BuildFluktuationAnalyse.bas, mod_RefreshFluktuationAll.bas, mod_KVStundenDropdown.bas, Modul1.bas: removed unused Fluktuation helpers; wired dirty/clean helpers without behavior change.
- mod_BuildFluktuationAnalyse.bas: FLUKTUATION dashboard redesigned for restaurant managers — status summary, KPI row, actionable alert table with month/row locations, numbered recommendations, and monthly/category charts.
- mod_BuildFluktuationAnalyse.bas: chart source columns stay visible for Mac Excel; charts use explicit series data. Alert/recommendation rows auto-fit height.
- mod_BuildFluktuationAnalyse.bas: title row height and column A width fixed; added horizontal bar chart for Austrittsgruende (Einvernehmlich, Dienstnehmer, etc.).
- mod_BuildFluktuationAnalyse.bas: merged-row height estimated for Empfehlungen/alerts; Wo nachschauen uses Monat statt Blatt without row numbers.
- mod_BuildFluktuationAnalyse.bas: all FLUKTUATION content cells vertically center aligned after layout.
- mod_BuildFluktuationAnalyse.bas: Monatsuebersicht lists only months with at least one exit again (no empty months).
- mod_BuildFluktuationDaten.bas: month detection uses each sheet's `A1` month number for reliable exit assignment.
- DieseArbeitsmappe.cls: opening `FLUKTUATION` always rebuilds analysis (not only when dirty flag is set).
- mod_RefreshFluktuationAll.bas: added `RefreshFluktuationNow` manual refresh macro.
- Modul1.bas, DieseArbeitsmappe.cls, mod_BuildFluktuationAnalyse.bas: analyse sheet name updated to `FLUKTUATION` (via `PID_FLUKTUATION_SHEET`).
- LOHNTABELLE migration: legacy salary sheet removed in workbook; `LOHNTABELLE_TEST` renamed to `LOHNTABELLE` in VBA (KV table, buttons, events, protection). Legacy `_TEST` public macro names kept as thin aliases.
- Modul1.bas: centralized workbook year and EINSTELLUNG config constants; `PID_GetWorkbookYear` now reads `EINSTELLUNG!C35` instead of legacy `LOHNTABELLE!G3`.
- mod_BuildFluktuationDaten.bas: Fluktuation reason weights (`EINSTELLUNG!B38:C49`) and time factors (`EINSTELLUNG!C53:C59`) are read from EINSTELLUNG instead of hardcoded values or legacy LOHNTABELLE cells.
- mod_BuildFluktuationAnalyse.bas: fixed compile error on Fluktuation sheet open (`yearValue` replaced with `currentYear` after EINSTELLUNG year migration).
- mod_FluctuationCalculation.bas, mod_KVLohnLookup.bas, mod_KVStundenDropdown.bas, mod_CopyData.bas: year lookup switched to `PID_GetWorkbookYear` / EINSTELLUNG.
- mod_AddNewKVPeriodOnTop.bas: team-friendly KV button labels (1/2/3 + Hilfe), step-by-step dialogs, single-period delete with double confirm, plain-language errors, and updated A2 guidance text.
- mod_AddNewKVPeriodOnTop.bas: `EnsureAddNewKVPeriodButton` now creates both KV sheet action buttons via `PID_EnsureLOHNTABELLE_TESTButtons`.
- mod_DataClear.bas: error handling tightened by replacing broad `On Error Resume Next` blocks with focused sheet protect/unprotect helper procedures and preserved original error details in failure messages.
- Password handling centralized: introduced `PID_WORKBOOK_PASSWORD` in Modul1 and replaced module-local password constants plus hardcoded `"company"` literals across VBA modules.
- mod_DataClear.bas: delete confirmation dialog now explicitly lists `Q31` (Fluktuation) as part of the cleared data.
- mod_SmokeCheck.bas: added `Manual Steps` output column so REVIEW cases include concrete manual verification steps in `SYSTEM_CHECK`.
- mod_DataClear.bas: added `Selection` type guard before row-clear logic to avoid non-range selection runtime issues.
- mod_AddNewKVPeriodOnTop.bas: improved KV period insertion flow (year-only input + configurable contract count), restored explicit period title rows, added robust trailing-area cleanup, and added `RestoreLOHNTABELLE_TESTBase2025_2026` for deterministic rollback baseline.
- mod_AddNewKVPeriodOnTop.bas: `RestoreLOHNTABELLE_TESTBase2025_2026` now filters invalid/partial rows before rebuilding the base period to prevent malformed top rows.
- mod_AddNewKVPeriodOnTop.bas: restored merged period title row (`A:J`) during rebuild/insert, enabled wrapped warning text in `A2`, removed dark first-row tint in KV visual grouping, and improved base-row recovery to preserve missing contract lines (e.g. 173 hours).
- mod_AddNewKVPeriodOnTop.bas: base restore now force-adds missing `BG1_Basis` row `173,00 / 2.021,00` before `151,38`, and `G:H` cells are unlocked so Monatsstunden/Monatslohn remain editable on protected sheet.
- mod_AddNewKVPeriodOnTop.bas: visual grouping now applies a stronger bottom border on the last KV table row (`A:J`) to match the table frame.
- mod_AddNewKVPeriodOnTop.bas: strengthened full outer table frame during visual grouping (`A:J`), with explicit thick bottom border for reliable visible closure.
- mod_AddNewKVPeriodOnTop.bas: fixed lock handling for `Monatsstunden`/`Monatslohn` by unlocking only real data rows (skip merged title rows), and normalized bottom frame weight/color to match left/right borders.
- mod_AddNewKVPeriodOnTop.bas: `FormatKVPeriodArea` now enforces uniform row font/border styling across `A4:J(lastRow)` so trailing rows match the rest of the table.
- mod_AddNewKVPeriodOnTop.bas: visual grouping now adds medium separator lines between KV code subgroups (`Basis`, `_5`, `_10`, `_15`) inside BG1/BG2/BG3 blocks.
- mod_AddNewKVPeriodOnTop.bas: status/check formulas in `I:J` are now reapplied for all valid data rows during formatting, so `Status` and `Pruefung` update dynamically after edits.
- mod_AddNewKVPeriodOnTop.bas: added `FixLOHNTABELLE_TEST_StatusFormulas`, auto-repair on open/sheet activate, removed static `OK` writes during rebuild/restore, and hardened formula application on protected sheets.

### Fixed
- mod_AddNewKVPeriodOnTop.bas: Mac Excel compile fix for KV sheet buttons (`Shape.Placement` via late binding; project compiles again on macOS).
- mod_AddNewKVPeriodOnTop.bas: restored missing `PID_GetKVTeamAfterChangeHint` and `PID_GetLOHNTABELLE_TESTTeamHelpText` helpers that caused compile errors.
- mod_AddNewKVPeriodOnTop.bas: LOHNTABELLE_TEST header layout reserves I2:J2 for action buttons; warning text stays in A2:H2 without overlap.
- mod_AddNewKVPeriodOnTop.bas: custom Monatsstunden insert now follows block sort order (descending by default), fills A-F metadata, and formats rows from the detected first data row instead of hardcoded row 4.
- mod_ResetAndImportVBAFiles.bas: clearer failure handling when VBProject access is blocked on Windows and when workbook path or vba folder is missing.
- mod_SmokeCheck.bas: TEST 8 no longer fails on its own detector string; API scan now matches only real `Declare` statements at line start and skips `mod_SmokeCheck.bas`.
- LOHNTABELLE_TEST: `Status`/`Pruefung` no longer stay as static `OK` text; formulas are restored automatically and update when `G`/`H` values are cleared.
- mod_AddNewKVPeriodOnTop.bas: widened `Status`/`Pruefung` columns (`I:J`) so long messages like `Monatsstunden fehlen` display fully.
- mod_AddNewKVPeriodOnTop.bas: fixed VBA compile error in `FormatKVPeriodArea` (`firstRow` was undefined).
- mod_AddNewKVPeriodOnTop.bas: reworked `AddNewKVPeriodOnTop` insertion flow (stable row insert, correct period bounds/title parsing, deterministic full-sheet formatting, no stale format copy after row shift).
- mod_AddNewKVPeriodOnTop.bas: new KV periods no longer copy `Monatslohn` from template; column `H` stays empty for manual entry.
- mod_AddNewKVPeriodOnTop.bas: restore/rebuild now removes trailing empty rows after shorter periods (15-row test -> 13-row base), filters rows without hours/wage, and clears leftover formats reliably.
- mod_AddNewKVPeriodOnTop.bas: removed extra confirmation dialog before inserting a new KV period; trims/deletes rows below table end to shrink scroll area (`CleanupLOHNTABELLE_TESTTrailingArea` / auto after format).
- mod_KVLohnLookup.bas: fixed period fallback when new KV period has empty `Monatslohn` (no longer aborts lookup on non-OK rows); month sheets refresh `G` on activate via cached single-sheet lookup (CopyData unchanged).
- mod_KVLohnLookup.bas: faster lohn refresh (single unprotect/protect per sheet, batched euro format, cached workbook year); month/LOHNTABELLE_TEST sheet activate no longer runs heavy work every time (dirty-flag + one refresh per month after KV table changes).
- mod_KVStundenDropdown.bas: fixed hour dropdown list after LOHNTABELLE_TEST edits (period name matching + refresh current month sheet when KV table is dirty).
- VBA encoding: all 16 .bas/.cls files confirmed BOM-free and pure ASCII-compatible for Windows-1252 import; fixed garbled bytes (0x8A, 0x9F) in mod_RefreshFluktuationAll.bas MsgBox strings ("vollstaendig", "zurueckgesetzt").

### Notes
- Project moved to GitHub version control.