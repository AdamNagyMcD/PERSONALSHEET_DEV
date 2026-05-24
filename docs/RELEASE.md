# PERSONALSHEET Release Checklist

Hasznalat minden etterembe kikerulo verzio elott.

## Elokeszites

- [ ] Git working tree tiszta (`git status`)
- [ ] `docs/CHANGELOG.md` Unreleased szekció naprakesz
- [ ] VBA valtozasok importalva a `Personalsheet.xlsm`-be

## Automatikus ellenorzes (Excelben)

1. **Alt+F8** → `ResetAndImportVBAFiles` (ha csak `vba/` valtozott)
2. VBA Editor → **Debug → Compile VBAProject** (hiba = STOP)
3. Workbook mentese
4. Excel ujrainditasa (ajanlott import utan)
5. **Alt+F8** → `FullSystemRefresh`
6. **Alt+F8** → `PID_RunSystemSmokeCheck`
   - TEST 7 = **PASS** kotelezo
   - TEST 9-16 = strukturális **PASS** ahol automatikus (Jahr, A1, UEBERSICHT, KV, Monatslohn)
   - TEST 1–6, 8, 16 = REVIEW OK, de manualis teszt kell

## Manualis kritikus tesztek

Lsd. [TEST_CASES.md](../TEST_CASES.md):

- [ ] **TEST 1** — Future Hour Change (CopyData)
- [ ] **TEST 2** — Exit Employee (Austrittsdatum)
- [ ] **TEST 3** — Future Employee Survival (backward copy)

## Vizuális spot-check

- [ ] UEBERSICHT: FINANZIELL blokk, Durchrechnung E30/I30 feher + szerkesztheto
- [ ] EINSTELLUNG: sarga = nem editierbar, weiss = Eingabe
- [ ] Egy honaplap: E/F dropdown mukodik, Monatslohn szamol
- [ ] FLUKTUATION: dashboard betolt (ha volt valtozas)

## Release

- [ ] `Personalsheet.xlsm` mentese
- [ ] Git commit + tag (pl. `v2026.05.22`)
- [ ] Backup: elozo verzio megmarad (OneDrive verzioelozmeny / masolat)
- [ ] Ettermeknek: csak a tagelt xlsm telepitese

## Rollback

1. Elozo tagelt `Personalsheet.xlsm` visszaallitasa
2. Ne futtass `ResetAndImportVBAFiles` a regi fajlon uj `vba/` kodddal ellentetben
3. Ha adat problema: etterem sajat mentesebol restore

## Admin makrok — SOHA ettermi usernek

- `ResetAndImportVBAFiles`
- `RebuildLOHNTABELLE` / `RestoreLOHNTABELLEBase2025_2026`
- `UnprotectEverything`
