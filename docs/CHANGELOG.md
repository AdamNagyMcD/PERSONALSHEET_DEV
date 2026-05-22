# PERSONALSHEET CHANGELOG

## Unreleased

### Fixed
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
- mod_KVStundenDropdown.bas: F-Dropdown-Refresh nur Zeilen mit KV-Code (E nicht leer); ScreenUpdating aus waehrend Refresh.
- DieseArbeitsmappe.cls: FLUKTUATION-Activate nur bei gFluktuationDirty; kein S35-S37-Calculate bei Monats-Activate.
- mod_FluctuationCalculation.bas: UEBERSICHT nach Fluktuation-Sync wieder schuetzen.
- mod_DataClear.bas: MarkFinanzSummaryDirty nach Monatsdaten-Loeschen.
- Modul1.bas: FullSystemRefresh ohne doppeltes Protection-Setup am Ende.
- DieseArbeitsmappe.cls: `Union`-Aufruf in `EnforcePasteValuesOnly` ohne Zeilenfortsetzung (Mac VBA Syntaxfehler behoben).
- mod_ResetAndImportVBAFiles.bas: ueberarbeiteter Import (Workbook-/Tabellenmodule per Code-Update statt Import); `ReadVBAFileWithoutAttributes` liest nur ab `Option Explicit` (kein VERSION/Attribute-Header mehr im Modul).
- Modul1.bas: `SyncDieseArbeitsmappeFromExport` macro alias for manual workbook-module sync.
- Personalsheet.xlsm: reverted broken direct XML patch on `UBERSICHT`; Durchrechnung block must be created via `BuildDurchrechnungUebersicht` macro (avoids corrupted `sharedStrings` counts on Mac Excel).

### Changed
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