# PERSONALSHEET TEST CASES

## TEST 1 — Future Hour Change

### Scenario
Employee hours changed in July.

### Expected
- May unchanged
- June unchanged
- July updated
- August+ updated

---

## TEST 2 — Exit Employee

### Scenario
Employee exit date entered in August.

### Expected
- Employee visible in August
- Employee removed from September onward

---

## TEST 3 — Future Employee Survival

### Scenario
New employee manually added in October.
Backward copy started from August.

### Expected
- October employee survives
- No accidental deletion

---

## TEST 4 — Override Survival

### Scenario
D/E values manually changed in future months.

### Expected
- Overrides survive propagation
- Previous months unchanged

---

## TEST 5 — O18:Q25 Propagation

### Scenario
Month propagation executed.

### Expected
- O18:Q25 always copied correctly

---

## TEST 6 — Column L

### Scenario
Month propagation executed.

### Expected
- Column L never propagates

---

## TEST 7 — Excel 2016 Compatibility

### Scenario
Workbook opened on Excel 2016.

### Expected
- No broken formulas
- No unsupported functions
- No VBA compile errors

---

## TEST 8 — Mac Compatibility

### Scenario
Workbook opened on MacOS Excel.

### Expected
- Macros function correctly
- No path issues
- No Windows-only dependencies

---

## TEST 9 — Repeated Hour Change Same Month (FP-030)

### Scenario
1. Change July hours to 150, run CopyData.
2. Change July hours again to 140, run CopyData again.

### Expected
- Second change wins: July=140 propagates to December.
- System does NOT stick to the first value (150).
- Windows + Mac identical.

---

## TEST 10 — Independent Later Override Survives Earlier Edit (FP-030)

### Scenario
1. Change July to 150, run CopyData.
2. Change November to 160, run CopyData.
3. Re-edit July to 140, run CopyData.

### Expected
- July–October = 140 (earlier fix propagates forward).
- November–December = 160 (independent later override survives the July edit).
- Editing an earlier month must NOT wipe a later month's explicit override.

---

## TEST 11 — Middle-Month Edit Keeps Both Neighbours (FP-030)

### Scenario
1. July=150, November=160 (each followed by CopyData).
2. Change September to 145, run CopyData.

### Expected
- July–August = 150, September–October = 145, November–December = 160.
- No override is silently deleted by editing a month between two existing overrides.

### Diagnostic
- Run `PID_ShowHourOverrideLog` before/after each CopyData to inspect stored overrides
  (read-only; does not modify data).