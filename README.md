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
| `AlleDatenLoeschen` | Minden honap minden adatanak torlese (dupla megerosites, keplet marad) |

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

Modul: `mod_ADMIN.bas`. Az **Alt+F8** listaban az `ADMIN_` nevek szamozottan,
egymas mellett, a lista elejen jelennek meg. `ADMIN_00_Hilfe` mutatja az attekintest.

| Csoport | Makrok | Mit csinal |
|---------|--------|------------|
| **01–05 Setup / VBA** | `ADMIN_01_VBA_Import`, `ADMIN_02_VBA_Export`, `ADMIN_03_VBA_Reparatur_Nach_Import`, `ADMIN_04_Admin_Panel`, `ADMIN_05_Makro_Uebersicht` | VBA import a `vba\` mappabol, export a `vba_export\` mappaba, import utani keplet-helyreallitas, `_ADMIN` panel |
| **10–18 Teszt / diagnosztika** | `ADMIN_10_Test_Smoke_Check`, `ADMIN_11_Test_Schnellcheck`, `ADMIN_12_Test_Performance`, `ADMIN_13_Test_Formelspalten`, `ADMIN_14_Test_Stunden_Log`, `ADMIN_15_Test_Aktionsprotokoll`, `ADMIN_16_Test_Ergebnisblatt`, `ADMIN_17/18_Tech_Blaetter_*` | Smoke teszt, gyorsellenorzes, teljesitmenymeres, keplet-ellenorzes, naplok, technikai lapok |
| **20–28 Javitas** | `ADMIN_20_Reparatur_Full_Refresh`, `ADMIN_21_Reparatur_Formelspalten`, `ADMIN_22/23/24_Reparatur_*`, `ADMIN_25_Namen_Aufraeumen`, `ADMIN_26_Schutz_AN`, `ADMIN_27_Schutz_AUS`, `ADMIN_28_UEBERSICHT_Schutz` | Full refresh, G/H/K/L kepletek, dropdownok, `#REF!` nevek torlese, lapvedelem |
| **30–38 Formazas / LOHNTABELLE** | `ADMIN_30..38` | Honaplapok es UEBERSICHT formazasa, KV tabla ujraepitese es javitasa |
| **40–41 Adattorles** | `ADMIN_40_Daten_Stunden_Log_Leeren`, `ADMIN_41_Daten_Alles_Loeschen` | Visszafordithatatlan — dupla megerosites |

Elnevezesi szabaly:

- `ADMIN_NN_...` = fejlesztoi / karbantartasi makro (csak neked)
- rovid nemet nev (`CopyData`, `DataClear`, …) = ettermi felhasznaloi makro
- `PID_...` = belso technika, kezzel nem inditando

A regi nevek valtozatlanul mukodnek: az `ADMIN_` bejegyzesek csak atiranyitasok,
a gombok, esemenyek es a dokumentacio erintetlenek.

## Fejlesztoi ellenorzo szkriptek (Windows nelkul is futnak)

| Szkript | Mit ellenoriz |
|---------|---------------|
| `python3 tools/vba_lint.py` | A `vba/` forrasok statikus ellenorzese — hianyzo/privat eljaras, duplikalt nev, hianyzo cimke, blokk-hiba, `Option Explicit`, nem Excel 2016-kompatibilis fuggveny |
| `python3 tools/check_vba_sync.py` | A `Personalsheet.xlsm`-be agyazott VBA egyezik-e a `vba/` mappaval |
| `python3 tools/check_formula_columns.py` | Hol hianyzik keplet G/H/K/L-ben a munkafuzetben (a `PID_PruefeFormelspalten` Linux-parja) |

Mindharom csak olvas. Ajanlott sorrend VBA modositas utan: `vba_lint.py` → Excelben import
+ Kompilieren → mentes utan `check_vba_sync.py`.

## Git

- Branch: `main` (release-ready)
- Commit uzenetek: rovid, nemet prefix (pl. `EINSTELLUNG: ...`)
- A `vba/` diff olvashato; a `Personalsheet.xlsm` binary
