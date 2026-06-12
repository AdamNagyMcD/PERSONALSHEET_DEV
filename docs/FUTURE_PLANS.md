# FUTURE PLANS — PERSONALSHEET

Technischer Backlog für geplante, aber **noch nicht umgesetzte** Verbesserungen.  
Aktuelles Verhalten bleibt unverändert, bis ein Eintrag explizit umgesetzt und in `CHANGELOG.md` dokumentiert wird.

---

## FP-001 — FINANZIELL-Sync bei freien Texteingaben im Panel (O18:Q25)

**Status:** Behoben (2026-05-25) — `O18:Q25` aus Immediate-Finanz-Watch entfernt  
**Priorität:** Niedrig / UX-Optimierung  
**Betroffene Bereiche:** Monatsblätter (JANUAR–DEZEMBER), FINANZIELL-Kette, UEBERSICHT, EINSTELLUNG

### Beobachtetes Verhalten (historisch)

Wird in **O20** (oder generell in **O18:Q25**) Text eingegeben, wirkte es so, als würde sich das **gesamte Blatt** aktualisieren (kurzes Flackern / sichtbare Neuberechnung im rechten Panel).

### Ursache

`PID_MonthChangeNeedsImmediateFinanzSync` behandelte **O18:Q25** wie FINANZIELL-relevant. OOXML/Formelanalyse: Crew-Labor fuer UEBERSICHT kommt nur aus **Q17:R29** und **S35** — das Personal-Panel **O18:Q25** hat keine Formelbezuege dorthin (nur CopyData-Propagation laut SPEC).

### Fix (Ist-Zustand)

- `mod_SumMergedCells.bas`: Immediate-Watch nur noch `Q17:R29` und `S35`.
- Aenderungen in **O18:Q25** loesen **keinen** `PID_SyncFinanzSummaryForMonth` mehr aus.
- Mitarbeiterdaten (E–L) → weiterhin deferred Sync (`MarkFinanzSummaryDirtyForMonth`); UEBERSICHT-Tab → `RefreshFinanzSummaryIfDirty`.

### Akzeptanzkriterien

- [x] Freitext in O20 (O18:Q25) loest keine sofortige FINANZIELL-Kette aus.
- [x] Q17:R29 / S35 Aenderungen syncen weiterhin sofort (Watch-Range unveraendert fuer diese Bereiche).
- [x] CopyData kopiert O18:Q25 unveraendert (separater Codepfad).
- [ ] Manuell: O20 tippen ohne Flackern; S35/Crew-Labor-Aenderung → UEBERSICHT korrekt.

### Betroffene Dateien (Referenz)

- `vba/mod_SumMergedCells.bas` — `PID_MonthChangeNeedsImmediateFinanzSync`
- `vba/DieseArbeitsmappe.cls` — `Workbook_SheetChange`
- `SPEC.md` — Copy Areas (`O18:Q25`)

---

## FP-002 — CopyData: O18:Q25 propagiert nicht in Folgemonate

**Status:** Behoben (2026-05-25, Snapshot O/Q) — `PID_ReadMonthPanelSnapshot` / `PID_WriteMonthPanelSnapshot`  
**Priorität:** Hoch — SPEC-konformes Verhalten fehlt  
**Betroffene Bereiche:** `CopyData`, Monatsblätter O18:Q25, Panel / Crew-Labor-Info

### Beobachtetes Verhalten (historisch)

Beim Ausführen von **CopyData** wurde der Bereich **O18:Q25** laut SPEC nicht zuverlässig in die **folgenden Monate** kopiert (Merge-Zellen; `FormulaR1C1`-Array wirkte nicht). Zusätzlich blieb auf Zielmonaten eine Markierung um O21:Q24.

### Fix (Ist-Zustand)

- `mod_CopyData.bas`: `PID_ReadMonthPanelSnapshot` / `PID_WriteMonthPanelSnapshot` — O/Q-Ankerzeilen (Merge), Quellblatt kurz entsperrt.
- `PID_ResetFollowingMonthSelections` setzt nach CopyData auf allen Zielmonaten die Auswahl auf A1 (ScreenUpdating aus).

### Akzeptanzkriterien

- [x] Nach CopyData vom Monat M sind O18:Q25 in M+1 … Dezember identisch mit Quellmonat M.
- [x] Keine sichtbare Markierung O21:Q24 auf Zielmonaten nach CopyData.
- [ ] Smoke / manueller CopyData-Test grün (Mac).

### Betroffene Dateien (Referenz)

- `vba/mod_CopyData.bas` — `PID_CopyMonthPanelBlock`, `PID_ResetFollowingMonthSelections`
- `SPEC.md` — Copy Areas

---

## FP-003 — Spalte L: bei Ergebnis 0 Zelle leer lassen (kein €0,00)

**Status:** Behoben (2026-05-25 v2) — B/C-Guard + 0→leer; Restore per `RC[-10]`-Marker  
**Priorität:** Mittel — UX / Konsistenz mit Spalte G  
**Betroffene Bereiche:** Monatsblätter Spalte L (Letztes Gehalt / Laborcost), Formel-Restore

### Fix (Ist-Zustand)

- `Modul1.bas`: L-Formel liefert bei Ergebnis 0 `""` statt sichtbarem €0,00.
- `PID_MonthSheetNeedsLetztesGehaltFormulaUpdate`: alte L-Formeln werden beim Open/Restore aktualisiert.

### Akzeptanzkriterien

- [x] Leere / irrelevante Mitarbeiterzeilen: L ohne €0,00, Zelle optisch leer.
- [ ] Zeilen mit echtem Laborcost-Wert > 0: weiterhin korrekt formatiert (€) — manuell pruefen.
- [x] CopyData ueberschreibt L in Zielmonaten weiterhin nicht (SPEC: L informational only).
- [ ] Mac + Windows Excel 2016+ kompatibel — manuell pruefen.

### Betroffene Dateien (Referenz)

- `vba/Modul1.bas` — `PID_GetLetztesGehaltFormulaR1C1`, Restore-Pfade

---

## FP-004 — LOHNTABELLE „Eigene Stunden“: F-Dropdown auf Monatsblatt erst nach erneuter E-Auswahl

**Status:** Behoben (2026-05-25) — dirty-Refresh baut F-Validation neu auf  
**Priorität:** Mittel — Workaround existierte (E erneut waehlen)  
**Betroffene Bereiche:** `LOHNTABELLE`, Monatsblätter Spalte E/F, KV-Stunden-Dropdown

### Beobachtetes Verhalten (historisch)

Nach **„Eigene Stunden“** auf **LOHNTABELLE** und Wechsel zum Monatsblatt erschien die neue Stundenzahl nicht in **F**, bis **E** erneut gewaehlt wurde.

### Ursache

`RefreshKVStundenDropdownForRow` aktualisierte bei bestehender Validation nur den Helper/Named Range, uebersprang aber `Validation.Add` (`PID_RowHasValidFStundenDropdown` → Exit). Excel (v. a. Mac) cached die alte Dropdown-Liste.

### Fix (Ist-Zustand)

- `mod_KVStundenDropdown.bas`: Bei `gKVDropdownsDirty` Named Range mit `replaceIfDifferent` und F-Validation immer neu anwenden.
- Lazy-Refresh pro Monatstab unveraendert (`RefreshKVDropdownsIfDirtyForSheet`); kein 12-Blatt-Refresh beim Open.

### Akzeptanzkriterien

- [x] Nach Eigene Stunden + Monats-Tab: F-Liste enthaelt neue Stunde ohne E-Re-Select (Code-Pfad).
- [ ] Workbook-Open bleibt schnell — manuell pruefen.
- [ ] F-Overrides / Zukunftsplanung — manuell pruefen.
- [ ] Mac + Windows — manuell pruefen.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas` — `RefreshKVStundenDropdownForRow`

---

## Performance-Backlog (geplant: Windows, leistungsstarker Rechner)

Analyse 2026-05-25. Bestehende Optimierungen (lazy Open, dirty Flags, deferred FINANZIELL) **nicht** zurueckbauen.  
Umsetzung und Messung bewusst **spaeter auf Windows** — Mac-Regressionen weiterhin mitdenken, aber Primaer-Testplattform Windows + grosses Workbook.

---

## FP-005 — Performance: F-Dropdown dirty-Refresh nur betroffene Zeilen

**Status:** Erledigt (2026-06-12) — Windows: scoped 0,15 s vs Voll 0,52 s  
**Priorität:** Hoch (beste ROI nach FP-004)  
**Plattform-Ziel:** Windows (Referenz-Messung), Mac smoke  
**Betroffene Bereiche:** `mod_KVStundenDropdown.bas`, `DieseArbeitsmappe.cls`

### Ist-Zustand

Nach `MarkAllKVDropdownsDirty` baut `RefreshKVStundenDropdownForSheet` beim ersten Monats-Tab **alle 80 F-Zeilen** neu (Validation + `KV_DD_*` Named Range). FP-004 korrekt, aber spuerbar langsam.

### Geplante Verbesserung

- Nur Zeilen mit gesetztem **E (KV-Code)** und passendem LOHNTABELLE-Zeitraum refreshen, **oder**
- Nur Zeilen, deren KV-Code von der letzten LOHNTABELLE-Aenderung betroffen ist (Zielgerichtet nach `AddCustomKVMonatsstunden`), **oder**
- Lazy: F-Liste erst bei Fokus/Klick auf F (Approach D aus alter FP-004-Analyse).

### Akzeptanzkriterien

- [x] Nach Eigene Stunden + Monats-Tab: neue Stunde in F ohne E-Re-Select (Regression FP-004) — Windows OK (2026-06-12).
- [x] Erster Monats-Tab nach scoped dirty spuerbar schneller: **0,15 s** (2b) vs **0,51 s** Baseline / **0,52 s** Voll-Refresh.
- [x] Workbook-Open weiterhin ohne 12× F-Rebuild (lazy dirty unveraendert).
- [ ] Mac smoke (Excel 2016+).

### Umsetzung (2026-06-12)

- `MarkKVDropdownDirtyForKVCode` / `MarkKVDropdownDirtyFromLOHNTABELLERange` — scoped dirty statt immer `MarkAllKVDropdownsDirty`.
- `RefreshKVStundenDropdownForSheetBulk` — nur betroffene KV-Codes/Zeilen; Helper-Spalten bleiben fuer unberuehrte Codes erhalten.
- `AddCustomKVMonatsstunden` / `DeleteCustomKVMonatsstunden` — scoped dirty fuer gewaehlten KV-Code.
- `DieseArbeitsmappe` LOHNTABELLE D4:G — KV-Code aus geaenderten Zeilen.
- FP-010 Schritt **2b** misst scoped dirty (BG1, Februar).

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas` — `RefreshKVStundenDropdownForSheet`, `RefreshKVStundenDropdownForRow`
- `vba/mod_AddNewKVPeriodOnTop.bas` — optional KV-Code/Periode an dirty-Refresh uebergeben

---

## FP-006 — Performance: weniger Workbook Named Ranges (`KV_DD_*`)

**Status:** Erledigt (2026-06-12) — `PID_CountKVDDNamedRanges` = 0 nach FullSystemRefresh (Windows)  
**Priorität:** Mittel — grosser Refactor, langfristig Open/Save  
**Plattform-Ziel:** Windows (viele Names = spuerbar bei Open/Save)  
**Betroffene Bereiche:** `mod_KVStundenDropdown.bas`, Workbook-Names

### Ist-Zustand

Pro Monatszeile ein Name `KV_DD_<Sheet>_<Row>` (ca. 80 × 12 ≈ 960 Names). Belastet Excel Open, Save, Validation.

### Geplante Verbesserung

- Weniger Names: z. B. eine Helper-Spalte pro KV-Code/Monat statt 80 Zeilen-Namen, **oder**
- Validation `Formula1` mit direkter Bereichsadresse ohne Workbook-Name (Mac-Merge-Fall testen).

### Akzeptanzkriterien

- [x] F-Dropdown funktioniert (E/F-Test + FP-004/FP-005 weiter OK).
- [x] `KV_DD_*` Names entfernt — `PID_CountKVDDNamedRanges` = **0** (2026-06-12).
- [ ] Open/Save-Zeit verbessert (vor/nach messen auf Windows) — optional MANU-Baseline.
- [ ] Mac smoke.
- [ ] Keine Regression CopyData — bei Gelegenheit pruefen.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas` — `GetDropdownNameForMonthRow`, `PID_EnsureWorkbookNameRefersTo`

---

## FP-007 — Performance: SheetChange — weniger doppelte Recalc pro Zeile

**Status:** Erledigt (2026-06-12) — Windows E/F/D/I-Test OK  
**Priorität:** Mittel — Alltags-UX beim Tippen  
**Plattform-Ziel:** Windows + Mac  
**Betroffene Bereiche:** `DieseArbeitsmappe.cls`, `mod_KVLohnLookup.bas`, `Modul1.bas`

### Ist-Zustand

Eine Aenderung in E/F kann nacheinander ausloesen: Monatslohn-VBA, Aktuelle Stunden, Letztes Gehalt, F-Invalidate, Fluktuation dirty, FINANZ deferred — mehrfaches Unprotect/Calculate pro Zeile.

### Geplante Verbesserung

- Paste/Bulk: ein Handler fuer den geaenderten Bereich, pro Zeile einmal recalculieren.
- `PID_PreloadKVLohnCaches` nur wenn LOHNTABELLE-Cache wirklich invalid (nicht bei jedem Einzelcell-Event doppelt).

### Akzeptanzkriterien

- [x] E/F-Aenderung: G und abhaengige Spalten korrekt (2026-06-12).
- [ ] Grosser Paste in B:F blockweise ohne Timeout — bei Gelegenheit.
- [ ] Keine stale UEBERSICHT/FINANZIELL-Werte — bei Gelegenheit.

### Umsetzung (2026-06-12)

- `PID_RecalculateMonatslohnForChangedRows`: Preload einmal; L-Recalc gesammelt am Ende (nicht pro Zelle doppelt).
- `PID_RecalculateLetztesGehaltForChangedRows`: E/F aus Watch-Range (Monatslohn uebernimmt L).
- `PID_RecalculateAktuelleStundenForChangedRows`: H-`Calculate` gesammelt per `Union`.
- `DieseArbeitsmappe`: redundantes `PID_PreloadKVLohnCaches` entfernt.

### Betroffene Dateien (Referenz)

- `vba/DieseArbeitsmappe.cls` — `Workbook_SheetChange`
- `vba/mod_KVLohnLookup.bas` — `PID_RecalculateMonatslohnForChangedRows`
- `vba/Modul1.bas` — `PID_RecalculateLetztesGehaltForChangedRows`, Aktuelle Stunden

---

## FP-008 — Performance: SheetSelectionChange entlasten

**Status:** Erledigt (2026-06-12) — Windows D/E/F-Navigation OK  
**Priorität:** Niedrig  
**Plattform-Ziel:** Windows  
**Betroffene Bereiche:** `DieseArbeitsmappe.cls`

### Ist-Zustand

Jede Auswahl in D3:F82 setzt `ScreenUpdating = False` und prueft E/F-Dropdown-Validierung; D-Zelle zeigt Beschaeftigungsdauer (O45).

### Geplante Verbesserung

- F/E-Rebuild nur wenn Validation wirklich fehlt/invalid (bereits teilweise — pruefen ob redundant).
- O45/D-Hinweis ohne globales ScreenUpdating-Off bei harmlosen Spruengen.

### Akzeptanzkriterien

- [x] Navigation E/F weiterhin fluessig (2026-06-12).
- [x] Kaputte Dropdowns werden weiterhin lazy repariert.

### Umsetzung (2026-06-12)

- `SheetSelectionChange`: `ScreenUpdating = False` nur bei kaputtem E/F-Dropdown (Repair-Pfad).
- D/O45-Hinweis ohne globales ScreenUpdating.
- `PID_MonthSheetHasValidKVCodeDropdown`: Sheet-Cache nach erstem gueltigen Check.

---

## FP-009 — Performance: Fluktuation inkrementell (optional)

**Status:** Erledigt (2026-06-12) — Compile OK, Save/Tab-Pfad getrennt  
**Priorität:** Niedrig (grosser Aufwand)  
**Plattform-Ziel:** Windows (grosses Workbook)  
**Betroffene Bereiche:** `mod_RefreshFluktuationAll.bas`, `mod_BuildFluktuationDaten.bas`, `mod_BuildFluktuationAnalyse.bas`

### Ist-Zustand

`RefreshFluktuationAll` baut Daten + Analyse komplett neu (dirty, vor Save). Bewusst deferred — kann trotzdem mehrere Sekunden dauern.

### Geplante Verbesserung

- Inkrementelle Aktualisierung nur geaenderte Mitarbeiter/Monate, **oder**
- Schwere Analyse nur beim Oeffnen des FLUKTUATION-Tabs (strenger als heute).

### Akzeptanzkriterien

- [ ] FLUKTUATION-Inhalt fachlich identisch nach D/I/N-Aenderung.
- [ ] Save mit dirty spuerbar schneller oder gleichwertig akzeptabel dokumentiert.

### Umsetzung (2026-06-12)

- **BeforeSave:** nur `BuildFluktuationDaten` (`RefreshFluktuationDataIfDirty`) — Analyse deferred.
- **FLUKTUATION-Tab:** `RefreshFluktuationIfDirty` — Daten (falls noetig) + Analyse.
- **Inkrementell:** `MarkFluktuationDirtyForMonthSheet` bei D/I/N; `BuildFluktuationDaten` rescannt nur dirty Monate.
- FP-010 Schritte **6** (Save-Daten) und **6b** (Tab voll).

---

## FP-010 — Performance: Mess-Protokoll Windows

**Status:** Erledigt AUTO (2026-06-12); FP-005–009 Re-Messung offen; MANU optional  
**Priorität:** Hoch — vor Implementierung FP-005–009  
**Plattform-Ziel:** Windows, leistungsstarker PC, Produktionsnahes Workbook

### Werkzeuge

- **Doku:** `docs/PERFORMANCE_BASELINE.md` (Tabellen fuer Baseline + Re-Messung)
- **Makro:** `PID_RunPerformanceBaseline` / `RunPerformanceBaseline` — AUTO-Schritte 2–7, Log-Blatt `PID_PERFORMANCE_LOG` (very hidden)

### Checkliste Messung

| # | Schritt | Methode |
|---|---------|---------|
| 1 | Cold Open bis erster Monats-Tab nutzbar | MANU (Stoppuhr) |
| 2 | LOHNTABELLE → Eigene Stunden → erster Monats-Tab | MANU + AUTO (KV-Refresh Proxy) |
| 3 | E/F-Aenderung → G stabil | MANU + AUTO (G-Recalc 1 Zeile) |
| 4 | CopyData Januar → Dezember | MANU (Testkopie!) |
| 5 | UEBERSICHT nach FINANZIELL-Aenderung | AUTO (`PID_SyncFinanzSummaryToUbersicht`) |
| 6 | Save mit `gFluktuationDirty = True` | MANU + AUTO (Fluktuation-Refresh) |
| 7 | `FullSystemRefresh` (Admin-Referenz) | AUTO |

### Ergebnis

- [x] Erste Windows-Baseline AUTO in `docs/PERFORMANCE_BASELINE.md` (2026-06-12: KV 0,51 s, G 0,02 s, FINANZ 0,11 s, Fluktuation 1,24 s, FullRefresh 8,18 s).
- [ ] MANU-Schritte (Cold Open, CopyData, Save) noch dokumentieren.
- [ ] Nach jedem FP-005–009-Fix: gleiche Schritte wiederholen.

---

## Schutz-Paket — Amateur-Vermeidung (Übersicht)

Ziel: Monatsblätter und kritische Bereiche **so weit wie möglich** gegen versehentliche Excel-Gesten absichern, ohne SPEC/TEST-konforme manuelle Eingaben zu blockieren (B/C Neuer MA, H/I Exit, D/E Override, O18:Q25 Panel, E/F KV).

| ID | Thema | Aufwand | Priorität |
|----|--------|---------|-----------|
| FP-011 | Fill Handle / Drag-Format | Klein (A) / Mittel (B) | Hoch |
| FP-012 | Sortieren auf Monatsblättern verbieten | Klein | Hoch |
| FP-013 | Lock-all + Whitelist-Unlock (Monatsblätter) | Mittel | Hoch |
| FP-014 | Nur entsperrte Zellen auswählbar (`EnableSelection`) | Klein–mittel | Mittel |
| FP-015 | Endanwender-Hinweise + Recovery-Doku | Klein (Doku) | Mittel |
| FP-016 | Schutz-Smoke / Regression nach Paket | Klein | Nach Umsetzung |

**Umsetzungsreihenfolge (empfohlen):** FP-011 A → FP-012 → FP-013 → FP-011 B (falls nötig) → FP-014 (Mac+Win testen) → FP-015 → FP-016.

**Bereits vorhanden (nicht neu):** Lapvédelem + Jelszó, Paste nur Werte (`EnforcePasteValuesOnly`), rejtett helper lapok, `FormatAllMonthSheets` als Layout-Reparatur.

---

## FP-011 — Véletlen cellahúzás (Fill Handle) tönkreteszi a formázást

**Status:** Erledigt (2026-06-12) — Windows teszt OK  
**Priorität:** Hoch — UX / védelem nem technikai szakértőknek  
**Aufwand:** **Klein bis mittel** (siehe unten)  
**Betroffene Bereiche:** Monatsblätter B3:N82, `mod_SchutzHinzufugen.bas`, `mod_FormatMonthSheet.bas`, optional `DieseArbeitsmappe.cls`

### Beobachtetes Verhalten

Nutzer hält **E** oder **F** (entsperrt) gedrückt und **zieht** auf **gesperrte** Zellen (**G, H, K, L**, B–D mit Zebra/Input-Look). Excel blockiert oft Werte, kopiert aber **Format** (Hintergrund, Rahmen) → Zebra/Guide-Stil „kaputt“, ähnlich wie früher AutoFill auf Spalte L (VBA), hier **User-Geste**.

### Ursache

- Schutz entsperrt nur **E:F**; `Locked` verhindert nicht zuverlässig **Format-Drag** (v. a. Mac).
- Kein `EnableFillHandle`-Disable, kein Auto-Restore nach Format-Schaden.

### Geplante Verbesserung (Aufwand)

| Stufe | Ansatz | Aufwand | Risiko |
|-------|--------|---------|--------|
| **A (empfohlen)** | Beim Schutz/Tab: `EnableFillHandle = False` auf Monatsblättern (und ggf. `CellDragAndDrop = False`) | **Klein** (~1 Modul, wenige Zeilen) | Gering — Nutzer nutzen F/E per Dropdown, selten Fill |
| **B** | `Workbook_SheetChange`: Format-Änderung auf Guide/Locked-Bereich → `PID_ApplyMonthEmployeeZebraRows` + Rahmen | **Mittel** | Mittel — Events, Performance bei Paste |
| **C** | Nur Doku + `FormatAllMonthSheets` als Reparatur-Makro | **Minimal** | Löst Unfall nicht, nur Recovery |

**Nicht** nötig: großer Schutz-Refactor, Named Ranges, 12-Blatt-Rebuild.

### Akzeptanzkriterien

- [x] Fill-Handle-Zug von F/E nach G/L/B–D zerstört Zebra/Guide-Format **nicht** (Application-Guard auf Monats-Tabs).
- [x] E/F-Dropdown und normales Tippen unverändert.
- [ ] Mac smoke.

### Umsetzung Stufe A (2026-06-12)

- `PID_ProtectWorkerMonthSheet` in `mod_SchutzHinzufugen.bas`: `AllowSorting:=False`.
- Fill Handle: `Application.EnableFillHandle=False` auf Monats-Tabs via `PID_ApplyMonthSheetFillHandleGuard` (Excel 2016 — kein Worksheet-Property).
- Alle Monatsblatt-`Protect`-Stellen nutzen zentral diese Funktion.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas` — `PID_ApplySheetProtectionForMacros`
- `vba/mod_FormatMonthSheet.bas` — `PID_ApplyMonthEmployeeZebraRows`, `FormatAllMonthSheets`
- `vba/DieseArbeitsmappe.cls` — optional `SheetChange`-Restore (Stufe B)

---

## FP-012 — Sortieren auf Monatsblättern deaktivieren

**Status:** Erledigt (2026-06-12) — Windows teszt OK  
**Priorität:** Hoch — verhindert schwere Daten-/Zeilenvermischung  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, alle Re-`Protect`-Stellen (CopyData, Modul1, FormatMonthSheet, …)

### Beobachtetes Verhalten

Nutzer löst versehentlich **Sort** aus (Menü oder Kontext). Zeilen 3–82 werden neu geordnet → Mitarbeiterzeilen, Formeln, KV-Zuordnung und Zebra passen nicht mehr zusammen.

### Ursache

`PID_ApplySheetProtectionForMacros` setzt für Monatsblätter `AllowSorting:=True` (analog andere Blätter).

### Geplante Verbesserung

- Monatsblätter (Januar–Dezember): `AllowSorting:=False` beim `Protect`.
- `AllowFiltering` nur beibehalten, wenn fachlich nötig; sonst ebenfalls `False` prüfen (optional, separates Mini-Ticket).
- Alle Stellen, die Monatsblätter erneut schützen, konsistent anpassen (grep `AllowSorting`).

### Akzeptanzkriterien

- [x] Auf Monatsblatt ist Sortieren im geschützten Zustand nicht möglich (Windows).
- [x] E/F-Dropdown, Tippen, CopyData, FormatAllMonthSheets unverändert funktionsfähig.

### Umsetzung (2026-06-12)

- `PID_ProtectWorkerMonthSheet` — `AllowSorting:=False`; `mod_KVLohnLookup`, `mod_KVStundenDropdown`, `mod_DataClear`, `Modul1` umgestellt.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas`
- `vba/mod_CopyData.bas`, `vba/Modul1.bas`, `vba/mod_FormatMonthSheet.bas` (Re-Protect-Parameter)

---

## FP-013 — Monatsblätter: Lock-all + Whitelist-Unlock

**Status:** Offen — Teil des **Schutz-Pakets** (Kern: „fast alles gesperrt“)  
**Priorität:** Hoch  
**Aufwand:** **Mittel**  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, `mod_FormatMonthSheet.bas`, `DieseArbeitsmappe.cls` (Paste-Bereiche), SPEC/TEST_CASES

### Beobachtetes Verhalten

Viele Zellen sehen durch **Input-Styling** editierbar aus (B:F, I:J, M:N), sind aber fachlich Formel-/Guide-Bereiche (G, K, L, …). Wenn `Locked` im Template/Format-Kopie `False` bleibt, können Nutzer **in geschützte Formeln tippen** oder Werte überschreiben.

### Ursache

Schutz setzt nur explizit **E:F** auf `Locked = False`, ohne vorher `Cells.Locked = True` für das ganze Blatt (Muster **EINSTELLUNG** in `mod_FormatEinstellung.bas` fehlt auf Monatsblättern).

### Geplante Verbesserung

1. Vor `Protect`: `ws.Cells.Locked = True` (ganzes Monatsblatt).
2. **Whitelist** entsperren (SPEC + TEST_CASES):
   - `E3:F82` — KV-Code / Stunden (täglich)
   - `B3:C82` — Neuer Mitarbeiter / Schlüssel (TEST 3)
   - `D3:D82` — Stunden-Override (TEST 4; E bleibt separat)
   - `I3:J82` — Exit-Daten (TEST 2)
   - `M3:N82` — falls weiter manuell genutzt (CopyData-Bereich)
   - `O18:Q25` — Panel-Freitext (SPEC)
3. **Gesperrt bleiben:** G, H (wenn nicht in Whitelist), K, L, Q17:R29, S35, Summen, Guide-Zebra außerhalb Whitelist.
4. `PID_MSRestoreMonthSheetDropdowns` / Format-Kopie: `xlPasteFormats` darf **Locked** auf E/F nicht wieder kaputt machen (bereits Kommentar in Code — nach Lock-all erneut validieren).
5. Nach Schutz: `PID_ApplySheetProtectionForMacros` erneut aufrufen oder zentral eine `PID_ApplyMonthSheetLockPolicy(ws)`.

### Akzeptanzkriterien

- [ ] Nutzer kann nur Whitelist-Bereiche bearbeiten; G/K/L/Q-Formelbereiche nicht.
- [ ] TEST 2 (Exit I/J), TEST 3 (B/C), TEST 4 (D/E-Override), Panel O18:Q25 weiterhin möglich.
- [ ] CopyData, KV-Dropdown, L-Restore, FINANZIELL-Sync ohne Regression.
- [ ] `FormatAllMonthSheets` setzt Lock-Policy nicht zurück (oder ruft Policy am Ende auf).
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas` — zentrale Policy
- `vba/mod_FormatMonthSheet.bas` — nach Format-Kopie
- `vba/mod_FormatEinstellung.bas` — Referenzmuster Lock-all/Whitelist
- `SPEC.md`, `TEST_CASES.md`

---

## FP-014 — Auswahl nur auf entsperrte Zellen (`EnableSelection`)

**Status:** Offen — Teil des **Schutz-Pakets**  
**Priorität:** Mittel  
**Aufwand:** **Klein–mittel** (Mac-Verhalten testen)  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, Monatsblätter

### Beobachtetes Verhalten

Nutzer markieren große Bereiche inkl. **gesperrter** Formelzellen → Löschen, Format-Paste oder Drag mit höherer Fehlerwahrscheinlichkeit.

### Geplante Verbesserung

- Nach `Protect` auf Monatsblättern: `EnableSelection = xlUnlockedCell` (nur entsperrte Zellen auswählbar).
- Optional: `xlNoRestrictions` auf Admin-Makro-Pfad kurzzeitig (wie bestehendes Unprotect für Makros).

### Akzeptanzkriterien

- [ ] Normale Arbeit in E/F und Whitelist (FP-013) unverändert.
- [ ] Makros (CopyData, Format, L-Restore) funktionieren mit `UserInterfaceOnly:=True`.
- [ ] Mac Excel 2016: keine blockierenden UX-Regressionen (Panel, Dropdowns).
- [ ] Windows gleicher Smoke wie Mac.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas`

---

## FP-015 — Endanwender-Hinweise und Recovery (Deutsch, kurz)

**Status:** Offen — Teil des **Schutz-Pakets**  
**Priorität:** Mittel  
**Aufwand:** **Klein** (Doku + optional feste Hinweiszelle)  
**Betroffene Bereiche:** Monatsblätter (z. B. feste Zelle), `README.md` oder `docs/`, Toolbar-Texte

### Ziel

Restaurant-Manager ohne Excel-Wissen: **was tun / was nicht** + **was tun bei kaputtem Layout**.

### Geplante Inhalte (Deutsch, kurz)

- Nur **E und F** täglich ändern; **nicht ziehen** (Fill Handle).
- Kein Sortieren, kein großflächiges Löschen.
- Bei kaputtem Zebra/Rahmen: Button **`FormatAllMonthSheets`** (oder Admin-Hinweis).
- Wichtige Aktionen nur über **Toolbar-Buttons** (CopyData, Aktualisierung, …).

**Ist (2026-05):** Ausgearbeitet in `docs/Kurzanleitung_Personalsheet_A4.html` (A4-Druck, einfache Sprache, Umlaute).

### Akzeptanzkriterien

- [x] Hinweis für Endnutzer in Kurzanleitung HTML (druckbar A4).
- [ ] Optional: gleicher Kurztext als Hinweiszelle auf Monatsblättern.
- [ ] Keine technischen Begriffe (kein „Named Range“, „SheetChange“).
- [ ] Recovery-Pfad zu `FormatAllMonthSheets` dokumentiert.

### Betroffene Dateien (Referenz)

- `docs/FUTURE_PLANS.md` (dieser Eintrag)
- Optional: `vba/mod_FormatMonthSheet.bas` (Hinweiszelle), `README.md`

---

## FP-016 — Schutz-Paket: Smoke / Regression

**Status:** Offen — nach FP-011–FP-015  
**Priorität:** Pflicht vor Release mit Schutz-Paket  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** `mod_SmokeCheck.bas`, `docs/RELEASE.md`, manuelle Mac/Win-Checkliste

### Geplante Prüfungen

- [ ] Monatsblatt: E/F editierbar, G/L nicht editierbar (nach FP-013).
- [ ] Sort auf Monatsblatt blockiert (FP-012).
- [ ] Fill-Handle-Zug von E/F zerstört Layout nicht (FP-011 A) oder Restore (FP-011 B).
- [ ] Paste in E/F nur Werte (bestehend + Whitelist).
- [ ] TEST 1–6 aus `TEST_CASES.md` grün.
- [ ] `FormatAllMonthSheets` stellt Zebra nach absichtlichem Format-Schaden wieder her.

### Betroffene Dateien (Referenz)

- `vba/mod_SmokeCheck.bas` (optional neue Tests)
- `docs/RELEASE.md`

---

## FP-017 — Monatsblätter: Spalte D und I breiter / datenabhängig

**Status:** Behoben (2026-05-26) — `PID_ApplyMonthSheetDateColumnWidths`, Breite 13 (wie G/J)  
**Priorität:** Mittel — UX, Daten sichtbar  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** Monatsblätter (Januar–Dezember), `mod_FormatMonthSheet.bas` oder `FormatAllMonthSheets`

### Beobachtetes Verhalten

Spalten **D** (Eintritt) und **I** (Austrittsdatum) sind zu schmal; Datumswerte werden abgeschnitten oder nicht vollständig angezeigt.

### Geplante Verbesserung

- Feste Mindestbreite für D und I (z. B. aus Referenzblatt Januar) **oder**
- `ColumnWidth` / `AutoFit` nur für D und I nach Format-Lauf (Mac/Win 2016+ testen).
- In `FormatAllMonthSheets` / Monats-Layout mitziehen.

### Akzeptanzkriterien

- [x] Feste Mindestbreite 13 für D und I (Code).
- [ ] Typische Datumsformate (dd.mm.yyyy) in D und I auf Monatsblatt voll sichtbar (Win Excel 2016, manuell bestätigt).
- [ ] Kein Layout-Bruch im Mitarbeiterblock B:N.
- [ ] Mac + Windows.

### Betroffene Dateien (Referenz)

- `vba/mod_FormatMonthSheet.bas`

---

## FP-018 — LOHNTABELLE: neue KV-Periode → erster Monats-Tab sehr langsam

**Status:** Behoben (2026-05-26) — F-Dropdown pro KV-Code (Bulk), Manual während Refresh, H/K/L danach einmal  
**Priorität:** Mittel — Performance (Alltag)  
**Aufwand:** **Mittel–groß** (an FP-005–FP-010 anknüpfen)  
**Betroffene Bereiche:** `mod_KVStundenDropdown.bas`, `DieseArbeitsmappe.cls` (lazy refresh), LOHNTABELLE

### Beobachtetes Verhalten

Nach **neuer KV-Periode** auf LOHNTABELLE dauert das **erste Öffnen** eines Monatsblatts sehr lange — vermutlich vollständiger KV-/Stunden-Refresh über alle Zeilen/Named Ranges.

### Geplante Verbesserung (Richtung)

- Dirty-Refresh weiter eingrenzen (nur aktives Blatt, nur betroffene KV-Codes/Zeilen).
- Named-Range-Strategie vereinfachen (siehe FP-006).
- Optional: Hintergrund-Refresh oder Fortschritts-Hinweis statt „eingefroren“ wirkender Pause.
- Windows-Messung laut FP-010 vor großem Umbau.

### Akzeptanzkriterien

- [x] Bulk-Refresh: ein Named Range pro KV-Gruppe statt 80× pro Zeile; `PID_BeginHeavyMaintenance` während SheetActivate.
- [ ] Nach neuer KV-Periode: erster Monats-Tab unter akzeptabler Zeit — manuell prüfen.
- [ ] F-Dropdown listet neue Stunden korrekt (Regression FP-004).
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas`
- `vba/DieseArbeitsmappe.cls`
- `docs/FUTURE_PLANS.md` FP-005–FP-010

---

## FP-019 — CopyData: Bestätigungsdialog entfernen

**Status:** Behoben (2026-05-26) — `PID_ConfirmCopyDataAction` entfernt  
**Priorität:** Niedrig — UX  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** `mod_CopyData.bas`

### Beobachtetes Verhalten

`PID_ConfirmCopyDataAction` zeigt vor CopyData ein Ja/Nein-Fenster. Nutzer, die den Button/Makro wählen, wollen **direkt** kopieren — Dialog wirkt überflüssig.

### Geplante Verbesserung

- `PID_ConfirmCopyDataAction`-Aufruf entfernen oder optional (Admin-Flag); CopyData startet sofort nach gültigem Monatsblatt.
- Kurze **Erfolgs-Meldung** am Ende beibehalten (oder optional stumm + nur Statuszeile).

### Akzeptanzkriterien

- [x] CopyData ohne Vorab-Dialog; Kopie läuft wie bisher (Overrides, Panel, SPEC).
- [x] Ungültiges Blatt / Fehler weiterhin mit Meldung (`PID_ValidateWorkerMonthSheet`).
- [ ] TEST 1–5 / CopyData manuell grün (nach Import in xlsm).

### Betroffene Dateien (Referenz)

- `vba/mod_CopyData.bas` — `PID_CopyDataToFollowingMonths`

---

## FP-020 — Monatsblatt Q12 (Vormonat +/-) entsperren

**Status:** Behoben (2026-05-26) — `Q12:R12` in `PID_UnlockSheetEditRanges`  
**Priorität:** Mittel — fachlich nötige Eingabe  
**Aufwand:** **Klein** (Teil von FP-013 Whitelist)  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, Panel rechts, `mod_FormatMonthSheet.bas`

### Beobachtetes Verhalten

**Q12** ist geschützt/gesperrt; Nutzer trägt dort **Vormonat +/-** ein. Layout: `Q12:R12` Input-Stil (`PID_MSApplyStyleToRangeMergedOnce`).

### Geplante Verbesserung

- In Schutz-Whitelist: `Q12` (bzw. `Q12:R12` merge) `Locked = False`.
- Mit FP-013 abgleichen: Panel-Eingaben vs. Formelbereiche Q17:R29 trennen.

### Akzeptanzkriterien

- [x] Q12:R12 auf geschütztem Monatsblatt editierbar (Code-Pfad).
- [ ] Q17:R29 / FINANZIELL-Formeln weiterhin nicht versehentlich überschreibbar (Ziel FP-013).
- [ ] CopyData / Panel-Snapshot unverändert — manuell prüfen.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas` — `PID_UnlockSheetEditRanges`, `PID_ReprotectWorksheet`

---

## FP-021 — UEBERSICHT: E30 und I30 entsperren (Durchrechnung Plan)

**Status:** Behoben (2026-05-26) — E30/I30 in `PID_UnlockSheetEditRanges`, UEBERSICHT-Schutz-Zweig  
**Priorität:** Mittel — fachlich nötige Eingabe  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, `mod_BuildDurchrechnung.bas`

### Beobachtetes Verhalten

Auf **UBERSICHT** sind **E30** (Jänner Verfügbar Plan) und **I30** (Jänner Muster Plan) trotz Doku/Smoke **gesperrt** — Nutzer muss dort manuell planen.

`PID_UnlockDurchrechnungInputs` setzt Unlock, wird aber vermutlich durch globales `Protect` ohne erneutes Unlock überschrieben.

### Geplante Verbesserung

- Beim Schutz von UEBERSICHT: nach `PID_UnlockDurchrechnungInputs` oder in `PID_ApplySheetProtectionForMacros` E30/I30 explizit `Locked = False`.
- Smoke TEST 12 bleibt grün.

### Akzeptanzkriterien

- [x] E30 und I30 auf geschütztem UEBERSICHT editierbar (Code-Pfad).
- [ ] Durchrechnungs-Formeln in anderen Zellen unverändert — manuell prüfen.
- [ ] TEST 12 PASS nach Import.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas` — `PID_UnlockSheetEditRanges`, `PID_ReprotectWorksheet`

---

## FP-022 — LOHNTABELLE: Hilfe-Button → „Eigene Stunden löschen“

**Status:** Behoben (2026-05-26) — Button `4) Stunde löschen`, Makro `DeleteCustomKVMonatsstunden`  
**Priorität:** Mittel — Alltag (ausgeschiedene MA, Stunde nicht mehr nötig)  
**Aufwand:** **Mittel**  
**Betroffene Bereiche:** `mod_AddNewKVPeriodOnTop.bas`, LOHNTABELLE, KV-Dropdown-Refresh

### Beobachtetes Verhalten

Button **Hilfe** auf LOHNTABELLE wird nicht benötigt. Stattdessen: **eigene Stunden** (Eigene Stunden / KV-Zeile) **löschen** können, wenn z. B. Mitarbeiter ausgeschieden und die Stundenzahl sonst niemandem mehr zugeordnet ist.

### Geplante Verbesserung

- Toolbar: Hilfe-Shape entfernen oder ersetzen durch z. B. **„Eigene Stunde löschen“** (deutscher Kurztext).
- Dialog: KV-Code / Stunde wählen → Zeile in LOHNTABELLE entfernen oder deaktivieren (Fachregel klären: nur leere Nutzung vs. harte Löschung).
- `MarkKVDropdownsDirty` / Monats-F-Listen aktualisieren.
- Sicherheitsabfrage vor Löschung (eine kurze Ja/Nein reicht hier).

### Akzeptanzkriterien

- [x] Hilfe-Button durch `4) Stunde löschen` ersetzt (`DeleteCustomKVMonatsstunden`).
- [x] Nur zusaetzliche Zeilen (mehr als Schema-Zeilen pro KV-Code) loeschbar; Ja/Nein-Bestaetigung.
- [ ] Gelöschte „eigene Stunde“ erscheint nicht mehr in F-Dropdown auf Monatsblättern (nach Refresh) — manuell prüfen.
- [ ] Bestehende KV-Perioden und normale KV-Zeilen unverändert — manuell prüfen.
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_AddNewKVPeriodOnTop.bas`
- `vba/mod_KVStundenDropdown.bas`

---

## FP-023 — Copyright: Blätter + VBA-Module (Adam Nagy / McOpCo)

**Status:** Behoben (2026-05-26) — `mod_PIDCopyright.bas`, Open/Refresh/FormatAll  
**Priorität:** Niedrig — Branding / Urheberrecht, kein Fachfeature  
**Aufwand:** **Klein–mittel** (Blätter) + **Klein** (VBA-Header, viele Dateien)  
**Betroffene Bereiche:** Alle sichtbaren Worksheets; alle `vba/*.bas` und `vba/*.cls`; zentral `mod_PIDUtils.bas` oder `mod_PIDCopyright.bas`

### Urheber (fest)

- **Name:** Adam Nagy  
- **Organisation:** McOpCo  

### Teil A — Hinweis auf jedem Excel-Blatt

Dezent, **nicht aufdringlich**, passend zum PERSONALSHEET-Design (kleine Calibri, gedämpftes Grau/helles Navy, Fußzeile rechts unten o. ä.).

**Vorschlag Anzeige-Text (einheitlich):**

`© Adam Nagy · McOpCo · Personalsheet [Jahr]`

- `[Jahr]` = Arbeitsjahr aus `EINSTELLUNG!C35` oder Release-Jahr (Konstante `PID_COPYRIGHT_YEAR`).
- Zentrale Konstante z. B. in `mod_PIDCopyright.bas`:
  - `PID_COPYRIGHT_AUTHOR = "Adam Nagy"`
  - `PID_COPYRIGHT_ORG = "McOpCo"`

**Technik:** `PID_ApplyCopyrightToAllSheets` — Open / `FormatAllMonthSheets` / `FullSystemRefresh`; sichtbare Blätter ja, `FLUKTUATION_DATEN` + `KV_DROPDOWN_HELPER` optional ohne.

### Teil B — VBA-Modul-Kopf (ja, üblich und sinnvoll)

**Nicht** in jede einzelne `Sub`/`Function` (zu laut, wartungsintensiv), sondern **ein Standard-Kommentarblock am Dateianfang** jeder exportierten Modul-Datei (direkt unter `Attribute VB_Name`, vor `Option Explicit`). Das ist in der Branche üblich; der VBA-Editor zeigt ihn beim Öffnen des Moduls; Export/Import behält die Zeilen.

**Vorschlag Block (Englisch — üblich bei Copyright; Kommentare sonst Deutsch):**

```vba
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================
```

Optional zusätzlich **eine Zeile** nur in zentralen Entry-Makros (`CopyData`, `FullSystemRefresh`, …) — nur wenn gewünscht; Standard = Modul-Kopf reicht.

**Umsetzung:** alle `vba/*.bas`, `vba/*.cls` einheitlich; bei `ResetAndImportVBAFiles` nicht überschreiben ohne Template.

### Fix (Ist-Zustand)

- `mod_PIDCopyright.bas`: `PID_ApplyCopyrightToAllSheets` — Monatsblatt `S2:V2` (neben CopyData-Button), LOHNTABELLE `L2:N2`, EINSTELLUNG `S2:U2`, UEBERSICHT `B25`, FLUKTUATION `A3`; kein PageSetup/Shape; Jahr aus `EINSTELLUNG!C35`.
- Aufruf: `Workbook_Open`, `PID_FullSystemRefresh`, `FormatAllMonthSheets`.
- Ausnahme: `FLUKTUATION_DATEN`, `KV_DROPDOWN_HELPER` (very hidden).
- Alle `vba/*.bas` und `vba/*.cls`: einheitlicher Copyright-Kommentarblock am Modulanfang.

### Akzeptanzkriterien

- [x] Alle sichtbaren Blätter: Hinweis mit **Adam Nagy** und **McOpCo** (Code-Pfad).
- [ ] Keine Störung von Eingabe, Formeln, Druck (A4 Spot-Check) — manuell prüfen.
- [x] Jedes VBA-Modul im Repo trägt den Copyright-Kopf (Adam Nagy / McOpCo).
- [ ] Mac + Windows Excel 2016+; nach Format-Lauf bleibt Blatt-Hinweis erhalten — manuell prüfen.

### Betroffene Dateien (Referenz)

- Neu oder `vba/mod_PIDSheetStyle.bas`
- `vba/mod_FormatMonthSheet.bas`, `vba/DieseArbeitsmappe.cls` (Open)
- `docs/RELEASE.md` (vor v1.0 optional mit umsetzen)

---

## FP-024 — Berechnung: Automatisch statt dauerhaft Manuell (H/K/L-Formeln)

**Status:** Behoben (2026-05-26)  
**Priorität:** Hoch — Endanwender / falsche Anzeige in H, K, L  
**Betroffene Bereiche:** `Modul1.bas`, `DieseArbeitsmappe.cls`

### Beobachtetes Verhalten

Nach jedem Öffnen stand Excel auf **Manuelle Berechnung**; Spalten **H, K, L** (Formeln) blieben oft leer oder veraltet, bis der Nutzer manuell auf Automatisch schaltete.

### Ursache

`Workbook_Open` und `PID_ConfigureDeferredWorkbookCalculationOnOpen` setzten `xlCalculationManual` ohne Rückstellung auf Automatisch.

### Fix (Ist-Zustand)

- Open-Ende: `PID_EnableCalculationForAllSheets` + `xlCalculationAutomatic`.
- `PID_RecalculateMonthFormulaColumns` bei Open (aktives Monatsblatt) und `SheetActivate` (H/K/L `.Calculate`).
- `Workbook_BeforeClose`: Automatisch für die Excel-Sitzung.
- Kurz **Manual** nur während der Open-Initialisierung (Performance), danach Automatisch.

### Akzeptanzkriterien

- [x] Nach Öffnen: Excel-Status „Automatische Berechnung“ (ohne manuelles Umstellen).
- [ ] H/K/L auf Monatsblatt zeigen Werte nach Tab-Wechsel (manuell Excel 2016).
- [ ] Open-Zeit akzeptabel (ggf. FP-018 separat).

---

## FP-025 — FLUKTUATION: PDF-Export (storniert)

**Status:** Storniert (2026-06-12) — Feature und zugehöriger VBA-Code entfernt  
**Grund:** PDF-Export auf FLUKTUATION wird nicht benötigt; Button und Makro `ExportFluktuationSheetToPDF` gelöscht. Legacy-Shape `btn_FluktuationPdfExport` wird bei Refresh/Tab-Activate entfernt.

---

## Test-Notiz (Excel 2016, 2026-05)

Manueller Durchlauf: **Smoke grün/gelb**, **manuelle Tests ohne Fehler**. Offene Punkte oben als FP-017–FP-023 erfasst (kein Release-Blocker für dokumentierte Fixes FP-001–FP-004).

---

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
