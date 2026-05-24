# PERSONALSHEET

Excel-alapu szemelyzeti es labor-planning rendszer ettermek szamara.

- **Workbook:** `Personalsheet.xlsm`
- **VBA forraskod:** `vba/` (git source of truth fejleszteshez)
- **Specifikacio:** [SPEC.md](SPEC.md)
- **Tesztesetek:** [TEST_CASES.md](TEST_CASES.md)
- **Valtozasok:** [docs/CHANGELOG.md](docs/CHANGELOG.md)
- **Felhasznaloi rovid utmutato:** [docs/Kurzanleitung_Personalsheet_A4.html](docs/Kurzanleitung_Personalsheet_A4.html)
- **Release checklist:** [docs/RELEASE.md](docs/RELEASE.md)

## Kompatibilitas

- Excel 2016+
- Windows es macOS
- Nincs XLOOKUP, LET, FILTER, dinamikus tombok

## Projektstruktura

```
Personalsheet.xlsm     <- futtathato workbook (1 etterem / ev)
vba/                   <- exportalt VBA modulok
docs/                  <- dokumentacio
SPEC.md                <- uzleti szabalyok
TEST_CASES.md          <- manualis tesztesetek
```

A `_xlsm_*` es `_einstellung_probe` mappak **lokalis fejlesztoi probe** celokra — nincsenek git-ben.

## Fejlesztoi workflow

1. VBA modositas a `vba/` mappaban (UTF-8)
2. Excelben: **Alt+F8** → `ResetAndImportVBAFiles`
3. Compile ellenorzes (VBA Editor)
4. Mentés → Excel ujrainditas (ajanlott)
5. **Alt+F8** → `FullSystemRefresh`
6. **Alt+F8** → `PID_RunSystemSmokeCheck`
7. Manualis tesztek: TEST 1–3 ([TEST_CASES.md](TEST_CASES.md))
8. `docs/CHANGELOG.md` frissitese
9. Release: [docs/RELEASE.md](docs/RELEASE.md)

## Ettermi felhasznalok (🟢 makrok)

| Makro | Mit csinal |
|-------|------------|
| `CopyData` | Honap adatainak masolasa a kovetkezo honapokba |
| `DataClear` | Aktualis honap adatainak torlese (megerosites kell) |
| `RefreshFluktuationNow` | Fluktuacio ujraszamolas |

Reszletek: Kurzanleitung HTML.

## Admin / fejlesztoi makrok (🔴 — ne ettermi usernek)

| Makro | Veszely |
|-------|---------|
| `ResetAndImportVBAFiles` | Minden VBA modul ujraimport |
| `RebuildLOHNTABELLE` | KV tabla ujraepites |
| `BuildLohntabelleTest` | LOHNTABELLE teljes torles |
| `UnprotectEverything` | Minden lap feloldasa |

## Git

- Branch: `main` (release-ready)
- Commit uzenetek: rovid, nemet prefix (pl. `EINSTELLUNG: ...`)
- A `vba/` diff olvashato; a `Personalsheet.xlsm` binary
