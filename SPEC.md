# PERSONALSHEET SPECIFICATION

## Core Requirements

- Excel 2016 compatible
- MacOS compatible
- Windows compatible
- No XLOOKUP usage
- Stable performance on large datasets
- No business logic changes without approval

---

## Employee Logic

### Employee Key
- Unique employee identifier:
  - Column B + Column C

### Exit Logic
- Employee remains visible in exit month
- Employee removed from all following months

### Future Employees
- Employees created in future months must survive backward copy operations

### Hour Changes
- Future hour changes must not overwrite previous months
- New hour settings apply only from intended month onward

---

## Monthly Propagation Rules

### Copy Areas
- B3:E82 copied
- H3:I82 handled separately
- O18:Q25 must always propagate

### Overrides
- D/E overrides must survive copy operations
- H/I exit data must survive copy operations
- Column L is informational only and must not propagate

---

## Compatibility Rules

- Must work on Excel 2016
- Must work on latest Excel
- Must work on Mac and Windows

---

## Stability Rules

- Never delete employees accidentally
- Never overwrite future planning accidentally
- Always preserve manually entered future data