# PERSONALSHEET CHANGELOG

## Unreleased

### Added
- Modul1.bas: `RestoreAktuelleStundenFormulas` / `PID_RestoreAktuelleStundenFormulas` restores column `H` pro-rata hour formulas on all month sheets using `EINSTELLUNG!C35` as workbook year.
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
- mod_BuildFluktuationAnalyse.bas: FLUKTUATION dashboard redesigned for restaurant managers — status summary, KPI row, actionable alert table with month/row locations, numbered recommendations, and monthly/category charts.
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