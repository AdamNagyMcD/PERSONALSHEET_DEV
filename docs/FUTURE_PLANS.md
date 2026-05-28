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

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
