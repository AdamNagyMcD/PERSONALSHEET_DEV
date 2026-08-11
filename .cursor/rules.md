# PERSONALSHEET Cursor Rules

## Core Working Principle

* Use the smallest necessary context for each task.
* Do not scan or analyze the whole project unless the user explicitly asks for a full audit, refactor, or architecture review.
* First inspect only the directly relevant file, module, procedure, worksheet, formula, or documented error.
* Expand context only when the first inspection is not enough to make a safe change.
* Prefer precise, local fixes over broad project-wide changes.
* Never rewrite working logic without explicit approval.

---

## General Rules

* Never change business logic unless explicitly requested.
* Always preserve Excel 2016 compatibility.
* Target platform is Windows only (Excel for Windows). Do not preserve, add, or restore macOS compatibility.
* Never use XLOOKUP / XVERWEIS.
* Never use LET, FILTER, UNIQUE or dynamic array formulas.
* Never delete future employee data.
* Never overwrite future planning accidentally.
* Prefer stability over aggressive optimization.
* Prefer minimally invasive changes.
* Never modify unrelated modules during a task.
* Do not refactor unrelated logic automatically.

---

## Context Control Rules

* Do not read large files completely if only one procedure, formula, range, or section is relevant.
* Do not repeatedly reread files already inspected in the same task.
* Do not open documentation, changelog, tests, and implementation files all at once unless the task requires it.
* Use targeted search for exact names, procedures, sheet names, ranges, error messages, or module names.
* For small tasks, avoid creating a long plan. Give a short diagnosis and apply the smallest safe fix.
* If a task only affects documentation, do not inspect VBA modules unless needed.
* If a task only affects VBA, do not inspect documentation unless needed.
* If a task only affects formulas, do not inspect unrelated modules unless needed.

---

## Documentation Rules

* Update `docs/CHANGELOG.md` only when a real code, formula, workbook behavior, or user-facing workflow change was made.
* Do not update the changelog for pure explanations, analysis, planning, or unchanged code.
* Keep changelog entries short and factual.
* Never invent undocumented changes.
* After completed implementation, bugfix, or refactor, provide a suggested git commit message.
* Prefix the suggestion with:

  * `Empfohlener Commit-Name:`
* Summarize the functional impact in 1-3 short bullet points.
* Do not force a commit message when no file was changed.

---

## Git / Export Rules

* Never intentionally modify exported `.bas` / `.cls` file encoding.
* Preserve UTF-8 compatibility.
* Avoid BOM-related export problems.
* Never auto-delete exported VBA modules.
* Preserve existing Git workflow compatibility.

---

## VBA Rules

* Avoid breaking existing module dependencies.
* Prefer readable VBA over overly complex optimizations.
* Preserve existing naming conventions unless explicitly requested.
* Write all VBA comments in German.
* Write all user-facing texts and messages in German.
* Do not rename existing procedures, variables, modules, worksheets, or named ranges unless explicitly requested.
* Never remove existing error handling unless explicitly requested.
* Before changing VBA, identify the exact procedure and the affected worksheet/range if applicable.
* Do not regenerate full modules when a small patch is enough.

---

## CRITICAL BOOTSTRAP MODULES

The following modules are critical bootstrap / infrastructure modules:

* `mod_ResetAndImportVBAFiles`
* `mod_CopyData`

For these modules the following rules apply:

* Never modify automatically.
* Never regenerate automatically.
* Never overwrite automatically.
* Never re-import automatically.
* Never refactor automatically.
* Never optimize automatically.
* Never rename automatically.

Before any change to these modules:

1. Explain why the change is necessary.
2. Name the exact affected lines of code.
3. Explain the risks.
4. Wait for explicit approval.

Without explicit user approval, these modules must never be changed.

### Script protection (compile / import / repair / test helpers)

* Every compile, import, repair, or test helper (PowerShell or VBA) — including temporary
  scripts — MUST use the same skip list and never delete or re-import these modules:

  * `mod_ResetAndImportVBAFiles`
  * `mod_CopyData`

* The skip list must be applied in BOTH the delete step and the import step.
* `tools/import_vba_and_repair.ps1` is the reference implementation of this skip list.

---

## Performance Rules

* Avoid unnecessary `Select` / `Activate` usage.
* Prefer array-based operations for large datasets.
* Disable `ScreenUpdating`, `EnableEvents`, and `Calculation` during heavy operations.
* Always restore Excel settings after errors.
* Avoid looping cell-by-cell when not necessary.
* Avoid unnecessary worksheet switching.
* Optimize only the affected area unless the user explicitly requests broader optimization.

---

## Stability Rules

* Keep fixes as isolated as possible.
* Preserve existing workflows unless explicitly requested.
* Prioritize predictable behavior over clever code.
* If uncertain about side effects, explain the risk briefly before changing critical logic.
* Never replace stable logic only for cleaner code.

---

## Worksheet / Structure Rules

* Never change existing worksheet names unless explicitly requested.
* Preserve all hidden sheets.
* Preserve all named ranges.
* Preserve existing table structures.
* Never reorder worksheets automatically.
* Never remove formulas or formatting outside the requested scope.
* Never change protected, locked, or user-guided areas unless the task requires it.

---

## Formula Rules

* Always use formulas compatible with Excel 2016.
* Preserve German Excel formula compatibility.
* Avoid modern Excel-only functions.
* Prefer stable formulas over shorter formulas.
* Do not replace working formulas with newer alternatives.

---

## Debug Rules

* When fixing bugs:

  * first explain the likely root cause briefly,
  * then explain the planned local fix,
  * only then generate or modify code.
* Do not perform a full project audit unless the bug cannot be isolated locally.
* If the cause is uncertain, make the smallest safe diagnostic step first.

---

## Testing Rules

* Before finishing a code or formula change, validate only the affected logic and directly related edge cases.
* For month-to-month copy changes, validate:

  * future employee handling,
  * termination date handling,
  * D/E hour overrides,
  * formula preservation,
  * Excel 2016 compatibility,
  * Windows Excel behavior (macOS is out of scope).
* Do not run or inspect unrelated test cases unless the affected logic requires it.
* Respect `TEST_CASES.md` when the task touches documented tested behavior.

---

## PERSONALSHEET-Specific Rules

* Never delete employees automatically only because rows appear empty.
* Preserve future D/E hour overrides.
* Preserve employee exit dates correctly.
* Preserve LOHNTABELLE / KV schema table structures and formatting.
* Preserve management hour logic.
* Preserve productivity calculation logic.
* Preserve existing month sheet structures.
* Preserve hidden helper sheet functionality.
* Preserve CopyData month-to-month logic unless explicitly requested.

---

## User Experience / Safety Rules

* The workbook is used by restaurant managers with very limited Excel knowledge.
* Assume users may accidentally click, overwrite, drag, paste wrong values, delete cells, delete rows, break filters, or change formatting.
* All workflows must be safe and foolproof.
* Protect critical areas whenever possible.
* Prevent accidental editing wherever possible.
* Avoid requiring technical Excel knowledge from users.
* Prefer locked cells over relying on user discipline.
* Prefer guided workflows and buttons over manual steps.
* Avoid hidden technical complexity for end users.
* All user-facing texts must remain:

  * in German,
  * simple,
  * clear,
  * short,
  * easy to understand for non-technical users.
* Avoid technical terminology in user messages whenever possible.
* Error messages must explain:

  * what happened,
  * why,
  * what the user should do next.
* Never expose internal helper logic or technical implementation details to end users.
* Design all features defensively against accidental misuse.

---

## Terminal Usage Rules

* The project root is `C:\DEV\PERSONALSHEET_DEV`.
* Do not use `cd "C:\DEV\PERSONALSHEET_DEV"` in terminal commands when the workspace is already opened from this folder.
* Do not chain many commands with `&&`.
* Prefer one command at a time.
* Read-only diagnostic commands are allowed:

  * `git status`
  * `git remote -v`
  * `git branch -a`
  * `git log -5 --oneline`
  * `git diff --stat`
  * `git diff --name-status`
* Do not run destructive commands without explicit confirmation:

  * `git reset`
  * `git clean`
  * `git restore`
  * `rm`
  * `del`
  * `rmdir`
* Do not run `git push`, `git commit`, or `git add .` automatically without user confirmation.

---

## Response Rules

* For small tasks, keep the response short.
* Do not produce long explanations unless the user asks for them.
* After file changes, summarize:

  * what changed,
  * which area was touched,
  * what was intentionally not touched.
* Do not claim that tests were run unless they were actually run.

NEVER MODIFY mod_ResetAndImportVBAFiles.
This is a critical bootstrap module.