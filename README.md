# PERSONALSHEET

Excel-alapu szemelyzeti es labor-planning rendszer ettermek szamara.

- **Workbook:** `Personalsheet.xlsm`
- **VBA forraskod:** `vba/` (git source of truth fejleszteshez)
- **Specifikacio:** [SPEC.md](SPEC.md)
- **Tesztesetek:** [TEST_CASES.md](TEST_CASES.md)
- **Valtozasok:** [docs/CHANGELOG.md](docs/CHANGELOG.md)
- **Felhasznaloi rovid utmutato:** [docs/Kurzanleitung_Personalsheet_A4.html](docs/Kurzanleitung_Personalsheet_A4.html)
- **Release checklist:** [docs/RELEASE.md](docs/RELEASE.md)
- **Geplante Verbesserungen (Backlog):** [docs/FUTURE_PLANS.md](docs/FUTURE_PLANS.md)

## Kompatibilitas

- Excel 2016+
- Csak Windows (macOS nem tamogatott)
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
| `PersonalIdKorrigieren` | Personal ID / nev javitasa mind a 12 honapon + Stunden-Log |
| `MitarbeiterEntfernen` | Dolgozo torlese minden honaprol vagy egy valasztott honaptol |
| `FehlerMelden` | Hibabejelentes: kontext + 2 kerdes → txt fajl a Feedback mappaba + vagolap |

Harom gomb minden honaplapon: „Mitarbeiter entfernen" (O7:P7),
„Personal-ID korrigieren" (Q7:R7) es „Fehler melden" (S7:T7).

Egy dolgozo = egy Personal ID: ugyanaz az ID egy honaplapon belul nem irhato be ketszer
(a masodik bevitel torlodik), es figyelmeztetes jon, ha az ID mas honapban mas nevhez tartozik.

Bemasolas mindig csak ertekkent: a Ctrl+V (es a menuszalag / jobbklikk beillesztes) sosem
hozza at a forras formazasat. Zarolt formulaoszlopot nem lehet vele felulirni, es a
Ctrl+X + beillesztes le van tiltva. A Ctrl+V atallitas csak addig el, amig ez a fajl az
aktiv — mas Excel fajlban valtozatlan marad.

Reszletek: Kurzanleitung HTML.

## Admin / fejlesztoi makrok (🔴 — ne ettermi usernek)

Modul: `mod_PIDAdmin.bas` — `PID_ShowAdminMacroInfo` listazza.

| Makro | Veszely |
|-------|---------|
| `ResetAndImportVBAFiles` | Minden VBA modul ujraimport (Bestaetigung) |
| `FullSystemRefresh` | Teljes workbook refresh |
| `RebuildLOHNTABELLE` | KV tabla ujraepites |
| `UnprotectEverything` | Minden lap feloldasa (Bestaetigung) |

## Git

- Branch: `main` (release-ready)
- Commit uzenetek: rovid, nemet prefix (pl. `EINSTELLUNG: ...`)
- A `vba/` diff olvashato; a `Personalsheet.xlsm` binary
