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

**Status:** Offen  
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

- [ ] Nach Eigene Stunden + Monats-Tab: neue Stunde in F ohne E-Re-Select (Regression FP-004).
- [ ] Erster Monats-Tab nach dirty spuerbar schneller als Voll-80-Zeilen-Refresh (Zeit messen vor/nach).
- [ ] Workbook-Open weiterhin ohne 12× F-Rebuild.
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas` — `RefreshKVStundenDropdownForSheet`, `RefreshKVStundenDropdownForRow`
- `vba/mod_AddNewKVPeriodOnTop.bas` — optional KV-Code/Periode an dirty-Refresh uebergeben

---

## FP-006 — Performance: weniger Workbook Named Ranges (`KV_DD_*`)

**Status:** Offen  
**Priorität:** Mittel — grosser Refactor, langfristig Open/Save  
**Plattform-Ziel:** Windows (viele Names = spuerbar bei Open/Save)  
**Betroffene Bereiche:** `mod_KVStundenDropdown.bas`, Workbook-Names

### Ist-Zustand

Pro Monatszeile ein Name `KV_DD_<Sheet>_<Row>` (ca. 80 × 12 ≈ 960 Names). Belastet Excel Open, Save, Validation.

### Geplante Verbesserung

- Weniger Names: z. B. eine Helper-Spalte pro KV-Code/Monat statt 80 Zeilen-Namen, **oder**
- Validation `Formula1` mit direkter Bereichsadresse ohne Workbook-Name (Mac-Merge-Fall testen).

### Akzeptanzkriterien

- [ ] F-Dropdown funktioniert auf allen Monatsblaettern (inkl. Merge in Panel-Naehe irrelevant fuer F).
- [ ] Anzahl `KV_DD_*` Names deutlich reduziert oder ersetzt.
- [ ] Open/Save-Zeit verbessert (vor/nach messen auf Windows).
- [ ] Keine Regression CopyData / E-F-Overrides.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas` — `GetDropdownNameForMonthRow`, `PID_EnsureWorkbookNameRefersTo`

---

## FP-007 — Performance: SheetChange — weniger doppelte Recalc pro Zeile

**Status:** Offen  
**Priorität:** Mittel — Alltags-UX beim Tippen  
**Plattform-Ziel:** Windows + Mac  
**Betroffene Bereiche:** `DieseArbeitsmappe.cls`, `mod_KVLohnLookup.bas`, `Modul1.bas`

### Ist-Zustand

Eine Aenderung in E/F kann nacheinander ausloesen: Monatslohn-VBA, Aktuelle Stunden, Letztes Gehalt, F-Invalidate, Fluktuation dirty, FINANZ deferred — mehrfaches Unprotect/Calculate pro Zeile.

### Geplante Verbesserung

- Paste/Bulk: ein Handler fuer den geaenderten Bereich, pro Zeile einmal recalculieren.
- `PID_PreloadKVLohnCaches` nur wenn LOHNTABELLE-Cache wirklich invalid (nicht bei jedem Einzelcell-Event doppelt).

### Akzeptanzkriterien

- [ ] E/F-Aenderung: G und abhaengige Spalten korrekt.
- [ ] Grosser Paste in B:F blockweise ohne Timeout.
- [ ] Keine stale UEBERSICHT/FINANZIELL-Werte.

### Betroffene Dateien (Referenz)

- `vba/DieseArbeitsmappe.cls` — `Workbook_SheetChange`
- `vba/mod_KVLohnLookup.bas` — `PID_RecalculateMonatslohnForChangedRows`
- `vba/Modul1.bas` — `PID_RecalculateLetztesGehaltForChangedRows`, Aktuelle Stunden

---

## FP-008 — Performance: SheetSelectionChange entlasten

**Status:** Offen  
**Priorität:** Niedrig  
**Plattform-Ziel:** Windows  
**Betroffene Bereiche:** `DieseArbeitsmappe.cls`

### Ist-Zustand

Jede Auswahl in D3:F82 setzt `ScreenUpdating = False` und prueft E/F-Dropdown-Validierung; D-Zelle zeigt Beschaeftigungsdauer (O45).

### Geplante Verbesserung

- F/E-Rebuild nur wenn Validation wirklich fehlt/invalid (bereits teilweise — pruefen ob redundant).
- O45/D-Hinweis ohne globales ScreenUpdating-Off bei harmlosen Spruengen.

### Akzeptanzkriterien

- [ ] Navigation E/F weiterhin fluessig.
- [ ] Kaputte Dropdowns werden weiterhin lazy repariert.

---

## FP-009 — Performance: Fluktuation inkrementell (optional)

**Status:** Offen — nur wenn FLUKTUATION-Tab/Save zu langsam  
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

---

## FP-010 — Performance: Mess-Protokoll Windows

**Status:** Offen (Voraussetzung fuer FP-005–009 Priorisierung)  
**Priorität:** Hoch — vor Implementierung  
**Plattform-Ziel:** Windows, leistungsstarker PC, Produktionsnahes Workbook

### Checkliste Messung (Stopuhr / StatusBar)

1. Cold Open (Manual Calc) bis erster Monats-Tab nutzbar.
2. LOHNTABELLE → Eigene Stunden → erster Monats-Tab (F-Dropdown oeffnen).
3. Eine Zelle E/F aendern (Reaktionszeit bis G stabil).
4. CopyData Januar → Dezember (Dauer).
5. UEBERSICHT-Tab nach FINANZIELL-Aenderung.
6. Save mit `gFluktuationDirty = True`.
7. Optional: `FullSystemRefresh` als Admin-Referenz (nicht Alltags-KPI).

### Ergebnis

- [ ] Baseline-Zeiten in diesem Eintrag oder `docs/RELEASE.md` Notiz festhalten.
- [ ] Nach jedem FP-00x-Fix: gleiche Schritte wiederholen.

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

**Status:** Offen — Teil des **Schutz-Pakets** (s. Übersicht oben)  
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

- [ ] Fill-Handle-Zug von F/E nach G/L/B–D zerstört Zebra/Guide-Format **nicht** (oder wird sofort unsichtbar durch A).
- [ ] E/F-Dropdown und normales Tippen unverändert.
- [ ] Nach Unfall (falls B): `FormatAllMonthSheets` oder ein Zeilen-Restore stellt Layout wieder her.
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas` — `PID_ApplySheetProtectionForMacros`
- `vba/mod_FormatMonthSheet.bas` — `PID_ApplyMonthEmployeeZebraRows`, `FormatAllMonthSheets`
- `vba/DieseArbeitsmappe.cls` — optional `SheetChange`-Restore (Stufe B)

---

## FP-012 — Sortieren auf Monatsblättern deaktivieren

**Status:** Offen — Teil des **Schutz-Pakets**  
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

- [ ] Auf Monatsblatt ist Sortieren im geschützten Zustand nicht möglich (Win + Mac Excel 2016+).
- [ ] E/F-Dropdown, Tippen, CopyData, FormatAllMonthSheets unverändert funktionsfähig.
- [ ] Kein Regression bei UEBERSICHT / LOHNTABELLE (dort ggf. weiter Sort erlaubt, falls gewünscht).

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

**Status:** Offen — aus Excel-16-Manualtest (2026-05)  
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

- [ ] Typische Datumsformate (dd.mm.yyyy) in D und I auf Monatsblatt voll sichtbar (Win Excel 2016, manuell bestätigt).
- [ ] Kein Layout-Bruch im Mitarbeiterblock B:N.
- [ ] Mac + Windows.

### Betroffene Dateien (Referenz)

- `vba/mod_FormatMonthSheet.bas`

---

## FP-018 — LOHNTABELLE: neue KV-Periode → erster Monats-Tab sehr langsam

**Status:** Offen — aus Excel-16-Manualtest (2026-05)  
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

- [ ] Nach neuer KV-Periode: erster Monats-Tab unter akzeptabler Zeit (Ziel nach Baseline-Messung).
- [ ] F-Dropdown listet neue Stunden korrekt (Regression FP-004).
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_KVStundenDropdown.bas`
- `vba/DieseArbeitsmappe.cls`
- `docs/FUTURE_PLANS.md` FP-005–FP-010

---

## FP-019 — CopyData: Bestätigungsdialog entfernen

**Status:** Offen — Wunsch aus Manualtest (2026-05)  
**Priorität:** Niedrig — UX  
**Aufwand:** **Klein**  
**Betroffene Bereiche:** `mod_CopyData.bas`

### Beobachtetes Verhalten

`PID_ConfirmCopyDataAction` zeigt vor CopyData ein Ja/Nein-Fenster. Nutzer, die den Button/Makro wählen, wollen **direkt** kopieren — Dialog wirkt überflüssig.

### Geplante Verbesserung

- `PID_ConfirmCopyDataAction`-Aufruf entfernen oder optional (Admin-Flag); CopyData startet sofort nach gültigem Monatsblatt.
- Kurze **Erfolgs-Meldung** am Ende beibehalten (oder optional stumm + nur Statuszeile).

### Akzeptanzkriterien

- [ ] CopyData ohne Vorab-Dialog; Kopie läuft wie bisher (Overrides, Panel, SPEC).
- [ ] Ungültiges Blatt / Fehler weiterhin mit Meldung.
- [ ] TEST 1–5 / CopyData manuell grün.

### Betroffene Dateien (Referenz)

- `vba/mod_CopyData.bas` — `PID_CopyDataToFollowingMonths`, `PID_ConfirmCopyDataAction`

---

## FP-020 — Monatsblatt Q12 (Vormonat +/-) entsperren

**Status:** Offen — aus Excel-16-Manualtest (2026-05)  
**Priorität:** Mittel — fachlich nötige Eingabe  
**Aufwand:** **Klein** (Teil von FP-013 Whitelist)  
**Betroffene Bereiche:** `mod_SchutzHinzufugen.bas`, Panel rechts, `mod_FormatMonthSheet.bas`

### Beobachtetes Verhalten

**Q12** ist geschützt/gesperrt; Nutzer trägt dort **Vormonat +/-** ein. Layout: `Q12:R12` Input-Stil (`PID_MSApplyStyleToRangeMergedOnce`).

### Geplante Verbesserung

- In Schutz-Whitelist: `Q12` (bzw. `Q12:R12` merge) `Locked = False`.
- Mit FP-013 abgleichen: Panel-Eingaben vs. Formelbereiche Q17:R29 trennen.

### Akzeptanzkriterien

- [ ] Q12 auf geschütztem Monatsblatt editierbar (Win Excel 2016).
- [ ] Q17:R29 / FINANZIELL-Formeln weiterhin nicht versehentlich überschreibbar (Ziel FP-013).
- [ ] CopyData / Panel-Snapshot unverändert.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas`
- `vba/mod_FormatMonthSheet.bas`

---

## FP-021 — UEBERSICHT: E30 und I30 entsperren (Durchrechnung Plan)

**Status:** Offen — aus Excel-16-Manualtest (2026-05); Smoke TEST 12 erwartet bereits Unlock  
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

- [ ] E30 und I30 auf geschütztem UEBERSICHT editierbar.
- [ ] Durchrechnungs-Formeln in anderen Zellen unverändert.
- [ ] TEST 12 PASS nach Fix.

### Betroffene Dateien (Referenz)

- `vba/mod_SchutzHinzufugen.bas`
- `vba/mod_BuildDurchrechnung.bas` — `PID_UnlockDurchrechnungInputs`

---

## FP-022 — LOHNTABELLE: Hilfe-Button → „Eigene Stunden löschen“

**Status:** Offen — aus Excel-16-Manualtest (2026-05)  
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

- [ ] Hilfe-Button weg oder durch Lösch-Workflow ersetzt.
- [ ] Gelöschte „eigene Stunde“ erscheint nicht mehr in F-Dropdown auf Monatsblättern (nach Refresh).
- [ ] Bestehende KV-Perioden und normale KV-Zeilen unverändert.
- [ ] Mac + Windows Excel 2016+.

### Betroffene Dateien (Referenz)

- `vba/mod_AddNewKVPeriodOnTop.bas`
- `vba/mod_KVStundenDropdown.bas`

---

## Test-Notiz (Excel 2016, 2026-05)

Manueller Durchlauf: **Smoke grün/gelb**, **manuelle Tests ohne Fehler**. Offene Punkte oben als FP-017–FP-022 erfasst (kein Release-Blocker für dokumentierte Fixes FP-001–FP-004).

---

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
