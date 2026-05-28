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

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
