# PERSONALSHEET Cursor Rules

## General Rules

- Never change business logic unless explicitly requested.
- Always preserve Excel 2016 compatibility.
- Always preserve MacOS compatibility.
- Never use XLOOKUP.
- Never delete future employee data.
- Never overwrite future planning accidentally.
- Prefer stability over aggressive optimization.

---

## Documentation Rules

- After every code change, update docs/CHANGELOG.md.
- Keep changelog entries short and factual.
- Never invent undocumented changes.

---

## VBA Rules

- Avoid breaking existing module dependencies.
- Prefer readable VBA over overly complex optimizations.
- Preserve existing naming conventions unless explicitly requested.
- Write all VBA comments in German.
- Write all user-facing texts and messages in German.
- Do not rename existing procedures, variables or modules unless explicitly requested.

---

## Safety Rules

- If uncertain, ask before modifying critical logic.
- Never perform destructive refactors automatically.
- Respect TEST_CASES.md requirements.