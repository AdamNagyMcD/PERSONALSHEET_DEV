# PERSONALSHEET Release Checklist

Hasznalat minden etterembe kikerulo verzio elott.

## Erste Release (v1.0) — wann?

**Geplant:** Erstes ettermi Release **nach** Umsetzung der Excel-16-Test-Punkte **FP-017–FP-022** (siehe `docs/FUTURE_PLANS.md`).

| Status | Inhalt |
|--------|--------|
| Erledigt (Basis) | FP-001–FP-004, Smoke + manuelle Tests Excel 2016 grün/gelb (2026-05) |
| Erledigt (v1.0-Code) | FP-017–FP-022 inkl. FP-018 KV-Ladezeit (Bulk) — manuell nach Import prüfen |
| Erledigt (v1.0-Code) | FP-023 Copyright Adam Nagy / McOpCo — nach Import prüfen |
| Optional / danach | Schutz-Paket FP-011–FP-016, Performance FP-005–FP-010 (Windows-Messung) |

**Magyarul:** Az első éttermi kiadás akkor jön, ha a fenti FP-017–FP-022 kész — utána ez a checklista + tagelt `Personalsheet.xlsm`.

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

## Release (v1.0)

- [ ] FP-017–FP-022 in `FUTURE_PLANS.md` auf **Behoben** gesetzt
- [ ] `docs/CHANGELOG.md`: Unreleased → neuer Versionsabschnitt (Datum)
- [ ] `Personalsheet.xlsm` mentese
- [ ] Git commit + tag (pl. `v1.0.0` oder `v2026.06.xx`)
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
