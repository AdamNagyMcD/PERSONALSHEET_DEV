# AGENTS.md

Project-wide guidance for AI agents working in this repository. See also
[`.cursor/rules.md`](.cursor/rules.md) (authoritative coding rules), [`README.md`](README.md),
[`SPEC.md`](SPEC.md) and [`TEST_CASES.md`](TEST_CASES.md).

## Cursor Cloud specific instructions

### What this project is (important for understanding "run/build/test")

- This is a **Windows-only Excel/VBA application**. The runnable artifact is the macro-enabled
  workbook `Personalsheet.xlsm`; the git source of truth for development is the exported VBA in
  `vba/*.bas` / `*.cls` (UTF-8). There is **no package manager, build system, or automated test
  runner** in the usual sense.
- The real develop/lint/test/run loop **requires Microsoft Excel for Windows** and cannot execute
  on the Linux Cloud VM:
  - "Lint/compile" = VBA Editor → *Debug → VBAProject kompilieren* (Windows/Excel only).
  - "Test" = run `PID_RunSystemSmokeCheck` in Excel + manual cases in [`TEST_CASES.md`](TEST_CASES.md)
    (Windows/Excel only).
  - "Build/run" = open `Personalsheet.xlsm` in Excel and use the macros/buttons (Windows/Excel only).
  - The dev import step is `tools/import_vba_and_repair.ps1` (PowerShell + Excel COM, Windows only)
    or the in-Excel macro `ResetAndImportVBAFiles`. See [`README.md`](README.md) → "Fejlesztoi workflow".
- Do **not** try to add macOS/Linux runtime support to the product — the codebase is intentionally
  Windows-only (see `.cursor/rules.md`).

### What CAN be done on the Linux Cloud VM

The update script installs lightweight Python tooling (`oletools`, `openpyxl`). Use it to validate
changes without Excel:

- **Validate the embedded VBA project parses** (closest available proxy for a VBA compile check):
  `python3 -m oletools.olevba Personalsheet.xlsm` — the workbook currently contains 53 VBA
  components (~22.7k lines) and parses cleanly.
- **Inspect workbook data/formulas** with `openpyxl` (load with `data_only=False` to see formulas;
  note openpyxl does **not** evaluate formulas and warns that the Data Validation extension is
  dropped on load — cosmetic, don't "fix" it).
- **Actually open & recalculate the workbook** with **LibreOffice Calc** (headless) — this renders
  the real sheets and recalculates *native* Excel formulas. Caveats:
  - LibreOffice is **not** installed by the update script (heavy system dependency). Install on
    demand: `sudo apt-get install -y libreoffice-calc poppler-utils`.
  - The `Lohn` column (`G`) uses the **VBA UDF** `PID_KVLohnLookup`, which does **not** evaluate in
    LibreOffice (shows blank/`#NAME?`). Native formulas (e.g. `Aktuelle Stunden` `H`,
    `Gesamt Crew Stunden` `Q8=SUM(H3:H82)`) **do** recalculate.
  - When driving LibreOffice via the UNO bridge (`python3-uno`), formulas passed to `setFormula`
    use **semicolons** as argument separators (e.g. `=DATE(2026;1;1)`, not commas) — commas produce
    `#NAME?`.
  - Month sheets `Januar`..`Dezember` layout: `B`=Personal ID, `C`=Name, `D`=Eintrittsdatum,
    `E`=KV Gruppe, `F`=Stunden, `G`=Lohn (VBA UDF), `H`=Aktuelle Stunden. The active year lives in
    `EINSTELLUNG!C35`; `H` prorates `F` by entry/exit date within that year.

### Editing rules reminder

- `vba/mod_ResetAndImportVBAFiles.bas` and `vba/mod_CopyData.bas` are **critical bootstrap modules** —
  never modify/regenerate/re-import them automatically (see `.cursor/rules.md`).
- All VBA comments and user-facing strings are **German**; keep Excel 2016 formula compatibility
  (no XLOOKUP/LET/FILTER/dynamic arrays).
