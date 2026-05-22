# PERSONALSHEET Cursor Rules

## General Rules

- Never change business logic unless explicitly requested.
- Always preserve Excel 2016 compatibility.
- Always preserve MacOS compatibility.
- Never use XLOOKUP.
- Never use LET, FILTER, UNIQUE or dynamic array formulas.
- Never delete future employee data.
- Never overwrite future planning accidentally.
- Prefer stability over aggressive optimization.
- Prefer minimal invasive changes.
- Never rewrite working logic without explicit approval.
- Never modify unrelated modules during a task.

---

## Documentation Rules

- After every code change, update docs/CHANGELOG.md.
- Keep changelog entries short and factual.
- Never invent undocumented changes.
- After every completed implementation, bugfix or refactor:
  - always provide a suggested git commit message.
- Prefix the suggestion with:
  - "Empfohlener Commit-Name:"
- Never finish a response without suggesting a commit message.
- Summarize the functional impact in 1-3 short bullet points.

---

## Git / Export Rules

- Never modify exported .bas/.cls file encoding intentionally.
- Preserve UTF-8 compatibility.
- Avoid BOM-related export problems.
- Never auto-delete exported VBA modules.
- Preserve existing Git workflow compatibility.

---

## VBA Rules

- Avoid breaking existing module dependencies.
- Prefer readable VBA over overly complex optimizations.
- Preserve existing naming conventions unless explicitly requested.
- Write all VBA comments in German.
- Write all user-facing texts and messages in German.
- Do not rename existing procedures, variables or modules unless explicitly requested.
- Never remove existing error handling unless explicitly requested.

---

## Performance Rules

- Avoid unnecessary Select / Activate usage.
- Prefer array-based operations for large datasets.
- Disable ScreenUpdating, EnableEvents and Calculation during heavy operations.
- Always restore Excel settings after errors.
- Avoid looping cell-by-cell when not necessary.
- Avoid unnecessary worksheet switching.

---

## Stability Rules

- Keep fixes as isolated as possible.
- Never refactor unrelated logic automatically.
- Preserve existing workflows unless explicitly requested.
- Prioritize predictable behavior over clever code.

---

## Worksheet / Structure Rules

- Never change existing worksheet names unless explicitly requested.
- Preserve all hidden sheets.
- Preserve all named ranges.
- Preserve existing table structures.
- Never reorder worksheets automatically.
- Never remove formulas or formatting outside the requested scope.

---

## Formula Rules

- Always use formulas compatible with Excel 2016.
- Preserve German Excel formula compatibility.
- Avoid modern Excel-only functions.
- Prefer stable formulas over shorter formulas.

---

## Debug Rules

- When fixing bugs:
  - first explain the root cause briefly,
  - then explain the planned fix,
  - only then generate code.
- If uncertain about side effects, warn before changing logic.

---

## Testing Rules

- Before finishing:
  - mentally validate month-to-month copy logic,
  - validate future employee handling,
  - validate termination date handling,
  - validate formula compatibility,
  - validate MacOS compatibility,
  - validate Excel 2016 compatibility.
- Respect TEST_CASES.md requirements.

---

## PERSONALSHEET-Specific Rules

- Never delete employees automatically only because rows appear empty.
- Preserve future D/E hour overrides.
- Preserve employee exit dates correctly.
- Preserve LOHNTABELLE (KV schema table) structures and formatting.
- Preserve management hour logic.
- Preserve productivity calculation logic.
- Preserve existing month sheet structures.
- Preserve hidden helper sheet functionality.

---

## Safety Rules

- If uncertain, ask before modifying critical logic.
- Never perform destructive refactors automatically.
- Never replace stable logic only for cleaner code.
- Always prefer safe incremental fixes.

## User Experience / Safety Rules

- The workbook is used by restaurant managers with very limited Excel knowledge.
- Assume users may accidentally click, overwrite, drag or delete cells unintentionally.
- All workflows must be as safe and foolproof as possible.
- Protect critical areas whenever possible.
- Prevent accidental editing wherever possible.
- Avoid requiring technical Excel knowledge from users.
- Prefer locked cells over relying on user discipline.
- Prefer guided workflows and buttons over manual steps.
- Avoid hidden technical complexity for end users.
- All user-facing texts must remain:
  - in German,
  - simple,
  - clear,
  - short,
  - easy to understand for non-technical users.
- Avoid technical terminology in user messages whenever possible.
- Error messages must explain:
  - what happened,
  - why,
  - what the user should do next.
- Never expose internal helper logic or technical implementation details to end users.
- Design all features defensively against accidental misuse.

- Always assume users can accidentally:
  - delete rows,
  - overwrite formulas,
  - drag cells,
  - paste wrong values,
  - break filters,
  - change formatting,
  - click wrong buttons.

  ## Terminal Usage Rules

- The project root is `C:\DEV\PERSONALSHEET_DEV`.
- Do not use `cd "C:\DEV\PERSONALSHEET_DEV"` in terminal commands when the workspace is already opened from this folder.
- Do not chain many commands with `&&`.
- Prefer one command at a time.
- Read-only diagnostic commands are allowed:
  - `git status`
  - `git remote -v`
  - `git branch -a`
  - `git log -5 --oneline`
  - `git diff --stat`
  - `git diff --name-status`
- Do not run destructive commands without explicit confirmation:
  - `git reset`
  - `git clean`
  - `git restore`
  - `rm`
  - `del`
  - `rmdir`
- Do not run `git push`, `git commit`, or `git add .` automatically without user confirmation.