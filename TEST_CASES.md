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