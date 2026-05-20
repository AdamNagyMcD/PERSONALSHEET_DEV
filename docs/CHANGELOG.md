# PERSONALSHEET CHANGELOG

## Unreleased

### Added
- Initial project structure
- SPEC.md
- TEST_CASES.md
- VBA export folder
- New smoke-check macro module `mod_SmokeCheck.bas` with `PID_RunSystemSmokeCheck` to log TEST_CASES 1-8 as PASS/FAIL/REVIEW on `SYSTEM_CHECK`.
- New smoke-check helper macros: `PID_FilterSmokeReviewOnly` and `PID_ClearSmokeFilter` for quick REVIEW-focused triage.
- Added `.gitignore` entry for `.DS_Store` to prevent accidental macOS metadata commits.

### Changed
- mod_DataClear.bas: error handling tightened by replacing broad `On Error Resume Next` blocks with focused sheet protect/unprotect helper procedures and preserved original error details in failure messages.
- Password handling centralized: introduced `PID_WORKBOOK_PASSWORD` in Modul1 and replaced module-local password constants plus hardcoded `"company"` literals across VBA modules.
- mod_DataClear.bas: delete confirmation dialog now explicitly lists `Q31` (Fluktuation) as part of the cleared data.
- mod_SmokeCheck.bas: added `Manual Steps` output column so REVIEW cases include concrete manual verification steps in `SYSTEM_CHECK`.
- mod_DataClear.bas: added `Selection` type guard before row-clear logic to avoid non-range selection runtime issues.
- mod_AddNewKVPeriodOnTop.bas: improved KV period insertion flow (year-only input + configurable contract count), restored explicit period title rows, added robust trailing-area cleanup, and added `RestoreLOHNTABELLE_TESTBase2025_2026` for deterministic rollback baseline.
- mod_AddNewKVPeriodOnTop.bas: `RestoreLOHNTABELLE_TESTBase2025_2026` now filters invalid/partial rows before rebuilding the base period to prevent malformed top rows.

### Fixed
- VBA encoding: all 16 .bas/.cls files confirmed BOM-free and pure ASCII-compatible for Windows-1252 import; fixed garbled bytes (0x8A, 0x9F) in mod_RefreshFluktuationAll.bas MsgBox strings ("vollstaendig", "zurueckgesetzt").

### Notes
- Project moved to GitHub version control.