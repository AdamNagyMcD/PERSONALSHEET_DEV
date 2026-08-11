# Test-Release v0.9.0-test — TODO (1. nagy tesztfázis visszajelzés)

**Forrás:** Restaurant-Manager teszt, első nagy tesztfázis  
**Dátum:** 2026-06-19 (utolsó frissítés: 2026-07-09)  
**Verzió:** Personalsheet Test-Release v0.9.0-test  
**Státusz:** Gyűjtés folyamatban — implementáció a következő napokban indul  
**Teljes release:** 2027 (addig több kolléga tesztel)

### Teszt összkép (2026-07-09)

- Általános visszajelzés: **jó, stabilan működik** a sheet a gyakorlati használatban.
- Ismert problémák: TR-01 … TR-04 (lásd alább) + TR-05 áttekintés.
- Új észrevételek: folyamatosan ide kerülnek; commit/push csak kérésre.

---

## ⏭️ KÖVETKEZŐ FELADAT (itt folytatjuk) — 2026-08-11 este

**TR-10 nincs kész. Két nyitott pont, ebben a sorrendben:**

### 1. `PID_RestoreFormulaColumnsForRows` „nincs meg" (compile hiba)

A makró futtatásakor „Sub or Function not defined". A függvény a `vba/Modul1.bas`
végén van (a `PID_RestoreUrlaubGeldFormulas` után), tehát a fájl megvan, de a
munkafüzetben lévő modul nem az aktuális.

Ellenőrizendő sorrendben:

- [ ] `ResetAndImportVBAFiles` lefutott-e a mentés óta, és melyik mappából importál
      (a repo `vba\` mappája vagy egy másik példány?).
- [ ] Az import kiírja a „skipped" darabszámot — `Modul1.bas` benne van-e a kihagyottakban.
- [ ] VBA-szerkesztőben `Modul1` végén ott van-e `PID_RestoreFormulaColumnsForRows`,
      `PID_GetEmployeeInputCellsForRows`, `PID_EnsureCellFormula`.
- [ ] Debug → Compile VBAProject: mi az első hibás sor.
- [ ] Ha az import nem hozza be: kézi import / beillesztés, aztán újra fordítás.

### 2. Full Refresh nem hozza vissza a hiányzó képleteket

`PID_FullSystemRefresh` hívja mind a négyet (`PID_RestoreMonatslohnFormulasSilent`,
`PID_RestoreAktuelleStundenFormulasSilent`, `PID_RestoreUrlaubGeldFormulasSilent`,
`PID_RestoreLetztesGehaltFormulasSilent`), a sérült sorok mégis üresek maradnak.

Vizsgálandó hipotézisek:

- [ ] **Lapvédelem:** `PID_TryUnprotectMonthSheetForMacro` / `ws.Unprotect` sikertelen →
      a `FormulaR1C1` írás csendben elszáll (a Silent függvények mindent elnyelnek).
      Teszt: `UnprotectEverything` után Full Refresh — ha így működik, ez az ok.
- [ ] **Korai kilépés:** `PID_RestoreAktuelleStundenFormulasOnSheet` kilép, ha
      `A1` nem szám; `PID_RestoreMonatslohnFormulasOnSheet` szintén szűr — meg kell nézni,
      hogy tényleg a teljes `3:82` tartományra ír-e, vagy csak feltételesen.
- [ ] **Hibaelnyelés:** minden Silent függvény `On Error GoTo SafeExit` — egy hiba az első
      lapon csendben leállíthatja a maradék 11-et. Ideiglenes debug-számláló kell.
- [ ] Egyetlen cellán manuálisan tesztelni: entsperren →
      `Cells(12,"K").FormulaR1C1 = PID_GetUrlaubGeldFormulaR1C1()` → hibát dob-e.

**Amíg ez nincs kész:** a `MitarbeiterEntfernen` / `DataClear` javítás (nem törli többé a
G/H/K/L oszlopot) önmagában már véd az ÚJ sérülés ellen — de csak akkor, ha a modulok
importja rendben lefutott.

---

## Prioritás összefoglaló

| ID | Prioritás | Típus | Rövid név | Státusz |
|----|-----------|-------|-----------|---------|
| TR-01 | 🔴 Magas | Bug | Personal ID beragad CopyData után | ⬜ Nyitott |
| TR-06 | 🔴 Magas | Feature | Personal ID / név javító makró (TR-01 gyakorlati megoldása) | 🟩 Kód kész — teszt nyitott |
| TR-08 | 🔴 Magas | Feature | Dolgozó eltávolítása a hónapokból (mind / adott hónaptól) | 🟩 Kód kész — teszt nyitott |
| TR-07 | 🔴 Magas | Feature | Personal ID egyediség-ellenőrzés (megelőzés) | 🟩 Kód kész — teszt nyitott |
| TR-09 | 🟡 Közepes | Feature | Hibabejelentő gomb + akciónapló (tesztfázis támogatás) | 🟩 Kód kész — teszt nyitott |
| TR-05 | 🟡 Közepes | Review | Alap gyorsítás & egyszerűsítés áttekintés | ⬜ Nyitott |
| TR-10 | 🔴 Magas | Bug | Dolgozó törlésekor a képletek is törlődtek (G/H/K/L) | 🟥 **Nyitott — itt folytatjuk** |
| TR-02 | 🟡 Közepes | Bug | Bemásolás mindig csak érték (formázás nélkül) | 🟩 Kód kész — teszt nyitott |
| TR-03 | 🟡 Közepes | Feature (1. fázis) | Minden adat törlése gomb | ⬜ Nyitott |
| TR-04 | 🟢 Alacsony (2. fázis) | Feature | Új év indítása | ⬜ Nyitott |

---

## TR-01 — Personal ID beragad javítás után is (CopyData)

**Státusz:** ⬜ Nyitott  
**Prioritás:** 🔴 Magas (tesztblokkoló)  
**Modul(ok):** `mod_CopyData.bas` (jóváhagyás szükséges — bootstrap modul)

### Tünet (teszt visszajelzés)

- Rossz Personal ID (B oszlop) bemásolódik a későbbi hónapokba CopyData-val.
- A forráshónapban javítás után újra CopyData — a későbbi hónapok **nem** frissülnek.
- A rossz érték „bent marad”, mintha le lenne fagyasztva.

### Valószínű ok (elemzés)

- Dolgozó-kulcs: `B | C` (Personal ID + Név) — `PID_BuildEmployeeKey`.
- Ha a későbbi hónapban rossz B van, más kulcs keletkezik, mint a javított forrásban.
- `PID_CollectFutureOverrides` ezt **új dolgozóként** (`futureNewStarts`) rögzíti, és lefagyasztja a B/C/D/E/F értékeket.
- Újramásoláskor a „szellem” sor visszakerül — a forrás javítása nem „nyeri fel” automatikusan.
- Az E/F stundák külön override-logban vannak (`PID_HOUR_OVERRIDES`); a B oszlop **nem** ugyanabban a védett logikában van, de a fenti mechanizmus hasonlóan blokkolja a javítást.

### Elméleti megoldási irányok

- [ ] **A)** ID-javítás felismerése: ha C (név) megegyezik, B eltér → korrekció, nem új MA; régi kulcs törlése `futureNewStarts`-ból.
- [ ] **B)** CopyData opció: „későbbi hónapok B/C/D mindig forrásból” (kivéve tényleges új belépő).
- [ ] **C)** Reprodukció rögzítése tesztben: dupla sor jelenik meg, vagy ugyanabban a sorban marad a rossz B?

### Érintett eljárások (vizsgálandó)

- `PID_CollectFutureOverrides`
- `PID_BuildTargetMonthData` / `PID_AddFutureNewEmployees`
- `PID_ApplyOverridesUntilMonth` (jelenleg csak E, F, I, J, M, N — nem B/C/D)

### Elfogadási kritériumok

- [ ] Rossz B → CopyData → forrás javítás → CopyData újra → **minden** későbbi hónap helyes B-t kap.
- [ ] Szándékos későbbi B-változtatás (ha üzletileg kell) továbbra is működik, vagy dokumentáltan nem támogatott.
- [ ] Mac + Windows, meglévő E/F override viselkedés (FP-028/FP-030) érintetlen.
- [ ] Új Smoke/Regression teszt (pl. TEST 25).

### Ideiglenes workaround (teszt alatt)

1. Admin: Stunden-Log leeren (`PID_AdminResetHourOverrideLog`).
2. Érintett későbbi hónapok: `DataClear` (egyenként).
3. Javított forráshónapból CopyData újra.
4. Vagy: minden érintett hónapon kézi B javítás (lásd chat 2026-07-09).

**Figyelem a kézi javításnál:** a `PID_HOUR_OVERRIDES` log a dolgozót `ID|NÉV` kulccsal tárolja.
Kézi B-átírás után a log bejegyzései árván maradnak (az óra-override-ok elveszhetnek).
Ezt automatikusan a TR-06 makró kezeli.

---

## TR-06 — Personal ID / név javító makró

**Státusz:** 🟩 Implementálva (2026-08-11) — manuális teszt (TEST 25) nyitott  
**Prioritás:** 🔴 Magas — TR-01 gyakorlati megoldása  
**Modul(ok):** **új** `mod_MitarbeiterPflege.bas` — `mod_CopyData.bas` **érintetlen maradt**

### Megvalósult (2026-08-11)

| Fájl | Mi történt |
|------|------------|
| `vba/mod_MitarbeiterPflege.bas` | **Új modul** — `PersonalIdKorrigieren` (Alt+F8), `PID_AdminKorrigierePersonalId` (gomb) |
| `vba/mod_PIDAdminSheet.bas` | „Personal-ID fix” gomb (`Case 13`) |
| `vba/mod_PIDUserText.bas` | `PID_UTxtGeaendert`, `PID_UTxtUnveraendert`, `PID_UTxtMoeglich` |
| `vba/mod_PIDAdmin.bas` | Admin makró-lista kiegészítve |
| `README.md` | Manager makró tábla kiegészítve |
| `TEST_CASES.md` | **TEST 25** (A–D forgatókönyv + negatív ellenőrzések) |

**Következő lépés:** Excel-ben `ResetAndImportVBAFiles` → Kompilieren → Speichern → TEST 25 lefuttatása.

### Cél

Egy elgépelt vagy megváltozott dolgozó-azonosító (Personal ID és/vagy név) javítása
**egy lépésben, konzisztensen** mind a 12 hónaplapon és a rejtett óra-override logban.

### Miért kell

- A dolgozó kulcsa `B|C` (ID + név) — `PID_BuildEmployeeKey` (`mod_CopyData.bas`).
- Rossz ID esetén a CopyData a későbbi hónapok sorát **új belépőnek** hiszi és megőrzi (TR-01).
- Kézi javítás hónaponként működik, de:
  - hosszadalmas (12 lap),
  - könnyű kihagyni egy hónapot,
  - **az óra-override log kulcsa nem frissül** → csendes adatvesztés.

### Felhasználói folyamat (javasolt)

1. Felhasználó egy hónaplapon **kijelöli a javítandó dolgozó sorát** (B vagy C cella elég).
2. Makró indítása (Alt+F8 vagy `_ADMIN` gomb).
3. Makró kiolvassa és megmutatja a jelenlegi ID-t és nevet.
4. InputBox: **új Personal ID** (üresen hagyva marad a régi).
5. InputBox: **új név** (üresen hagyva marad a régi).
6. Megerősítő dialógus: mi változik, **mely hónapokban hány sor**.
7. Végrehajtás + záró üzenet a módosított hónapok listájával.

### Mit módosít

| Cél | Művelet |
|-----|---------|
| Mind a 12 hónaplap | B és/vagy C oszlop átírása minden sorban, ahol a régi `B\|C` kulcs egyezik |
| `PID_HOUR_OVERRIDES` | C oszlop (`EmployeeKey`) átírása régi kulcsról az újra |
| Fluktuáció | `MarkFluktuationDirty` — a `FLUKTUATION_DATEN` a hónaplapokból épül újra |
| Finanz összesítő | Nem érintett (ID nem szerepel benne) |

### Validáció (végrehajtás előtt, minden hónapon)

- [ ] A régi kulcs **létezik** legalább egy hónapon — különben hibaüzenet.
- [ ] Az új ID **nem foglalt** másik dolgozó által egyik hónapban sem (más név ugyanazzal az ID-vel) → megtagadás.
- [ ] Az új `B|C` kulcs nem ütközik meglévő sorral ugyanazon a hónaplapon (dupla sor elkerülése).
- [ ] Üres ID + üres név nem engedélyezett.
- [ ] Összehasonlítás normalizálva: `Trim$`, `UCase$` — a `PID_BuildEmployeeKey` mintája szerint.
- [ ] Ha ütközés van, **semmit nem ír át** (mindent vagy semmit).

### Technikai megvalósítás

- Lapvédelem: minden hónaplapon `Unprotect` → írás → `PID_ReprotectWorksheet` (`mod_SchutzHinzufugen.bas`).
- Az override log lap `xlSheetVeryHidden` és védett lehet — ideiglenes feloldás, majd visszaállítás.
- `Application.EnableEvents = False` / `ScreenUpdating = False` a futás idejére, `CleanFail` ágban visszaállítás
  (a `mod_DataClear.bas` mintája szerint).
- Hibakezelés: `On Error GoTo CleanFail`, eredeti `Err.Number`/`Description` megőrzése az üzenetben.

### Újrahasznosítható helperek (ellenőrizve a kódban)

| Helper | Modul | Láthatóság |
|--------|-------|-----------|
| `PID_ValidateWorkerMonthSheet` | `mod_CopyData.bas` | Public |
| `PID_IsWorkerMonthSheet` | `mod_CopyData.bas` | Public |
| `PID_MonthNames`, `PID_GetMonthIndexFromSheetName` | `mod_PIDUtils.bas` | Public |
| `PID_ReprotectWorksheet` | `mod_SchutzHinzufugen.bas` | Public |
| `MarkFluktuationDirty` | `mod_RefreshFluktuationAll.bas` | Public |
| `PID_ConfirmAdminAction` | `mod_PIDAdmin.bas` | Public |
| `PID_UTxt*` szövegépítők | `mod_PIDUserText.bas` | Public |

**Megjegyzés:** `PID_BuildEmployeeKey` a `mod_CopyData.bas`-ban **Private**.
A javító modul saját, azonos logikájú lokális helpert kap (a bootstrap modult nem módosítjuk).

### Elnevezés (javaslat)

- Modul: `mod_MitarbeiterPflege.bas` (TR-08-cal közös, a helperek osztottak)
- Belépési pont (Alt+F8): `PersonalIdKorrigieren` — a `DataClear` / `CopyData` mintája szerint
- Fő eljárás: `PID_KorrigierePersonalIdUndName`
- Admin wrapper: `PID_AdminKorrigierePersonalId`

### Gomb (opcionális, 2. lépés)

`mod_PIDAdminSheet.bas` → `PID_AdminGetButtonSpec`:
jelenleg `Case 0..12` + `Case Else` (Admin verbergen); új `Case 13` felvétele és
`PID_ADMIN_BTN_COUNT` növelése.

### Peremesetek

- [ ] Ugyanaz a dolgozó **több sorban** egy hónapon belül → mindegyiket javítsa, de jelezze.
- [ ] Kilépett dolgozó (nem szerepel minden hónapban) → csak ahol létezik.
- [ ] Csak név változik, ID marad (házasság, ékezet) → ugyanaz a folyamat.
- [ ] Vezető nullás ID (`00123`) → szövegként kezelendő (TR-02 kapcsolat).
- [ ] Makró **nem** fut, ha nem hónaplap az aktív lap.

### Elfogadási kritériumok

- [ ] Rossz ID minden hónapon javítva egy futással.
- [ ] `PID_HOUR_OVERRIDES` kulcsai frissültek — az óra-override-ok a javítás után is érvényesek.
- [ ] Utána CopyData a forráshónapból **nem** hozza vissza a rossz ID-t (nincs „szellem” sor).
- [ ] Ütközés esetén semmi nem íródik át, érthető hibaüzenet (német).
- [ ] Lapvédelem és számítási mód a futás után visszaáll.
- [ ] Új Smoke/regressziós teszt: **TEST 25**.

### Tesztesetek (TEST 25)

1. Januárban rossz ID → CopyData decemberig → javító makró → minden hónap helyes.
2. Javítás után CopyData januárból → nem jelenik meg a régi ID.
3. Dolgozó óra-override-dal (pl. Juli F=150) → ID javítás → az override megmarad.
4. Új ID már foglalt egy másik dolgozónál → megtagadás, nincs írás.
5. Csak név módosítása → kulcs frissül, override megmarad.

### Eldöntött kérdések

- [x] Azonosítás: **sor kijelölése**, InputBox csak fallback (admin gomb).
- [x] Elérhetőség: **manager (Alt+F8) és `_ADMIN` gomb** is.
- [x] Hatókör: **mind a 12 hónap** (a „csak innentől” a TR-08-ban van meg).

---

## TR-08 — Dolgozó eltávolítása a hónapokból

**Státusz:** 🟩 Implementálva (2026-08-11) — manuális teszt (TEST 26) nyitott  
**Prioritás:** 🔴 Magas — a TR-06 párja  
**Modul:** `mod_MitarbeiterPflege.bas` (TR-06-tal közös) — `mod_CopyData.bas` **érintetlen**

### Cél

Egy tévedésből felvitt vagy duplán szereplő dolgozó eltávolítása **egy lépésben**:
vagy mind a 12 hónapról, vagy egy választott hónaptól decemberig.

### Felhasználói folyamat

1. Dolgozó sorának kijelölése egy hónaplapon (B vagy C cella elég) — vagy InputBox az ID-vel.
2. Zeitraum-kérdés: **JA** = mind a 12 hónap, **NEIN** = hónapszám 1–12 (alapérték: az aktív hónap), **ABBRECHEN** = kilépés.
3. Megerősítő dialógus: mely hónapokban hány sor + figyelmeztetés, hogy valódi kilépéshez az **I oszlop (Austrittsdatum)** való.
4. Végrehajtás + záró üzenet (törölt sorok és log-bejegyzések száma).

### Mit módosít

| Cél | Művelet |
|-----|---------|
| Hónaplapok (startMonth…december) | `B:N` tartalom törlése minden egyező soron — **sor nem törlődik**, formátum/zebra marad |
| `PID_HOUR_OVERRIDES` | A dolgozó bejegyzései **törlődnek**, ahol a hónap >= startMonth (visszafelé, indexbiztosan) |
| Formulák | Érintett laponként `PID_EnsureMonatslohnFormulasOnSheet` (G visszaáll) |
| Dirty flagek | `MarkFinanzSummaryDirtyForMonth` laponként, majd `MarkFluktuationDirty` + `MarkAllKVDropdownsDirty` |

### Miért törlünk log-bejegyzést (a TR-06-tal ellentétben)

Ha az override-ok bennmaradnának, egy későbbi azonos `ID|NÉV` felvitelekor **visszatérnének**
a régi óraszámok — csendes adathiba.

### Peremesetek

- [ ] A dolgozó több sorban szerepel egy hónapon → mindegyik törlődik.
- [ ] A kijelölt hónap előtti hónapok érintetlenek maradnak.
- [ ] CopyData a törlés előtti hónapból **szándékosan** visszahozza a dolgozót — ez CopyData-viselkedés, nem hiba (TEST 26/D).
- [ ] Érvénytelen hónapszám (0, 13, szöveg) → elutasítás, nincs írás.

### Elfogadási kritériumok

- [ ] Egy futással eltűnik a dolgozó a kért hónapokból, a struktúra sértetlen.
- [ ] Az óra-override-jai eltűnnek a kért időszakra, másoké érintetlen.
- [ ] Lapvédelem és számítási mód visszaáll.
- [ ] Új teszt: **TEST 26**.

---

## TR-07 — Personal ID egyediség-ellenőrzés (megelőzés)

**Státusz:** 🟩 Implementálva (2026-08-11) — manuális teszt (TEST 27) nyitott  
**Prioritás:** 🔴 Magas — a TR-01 valódi megelőzése  
**Modul(ok):** **új** `mod_PersonalIdUnique.bas` + `DieseArbeitsmappe.cls` (`Workbook_SheetChange`)

### Megvalósult (2026-08-11)

| Fájl | Mi történt |
|------|------------|
| `vba/mod_PersonalIdUnique.bas` | **Új modul** — `PID_CheckPersonalIdUniqueness` (kemény elutasítás + puha kereszt-hónap figyelmeztetés) |
| `vba/DieseArbeitsmappe.cls` | `Workbook_SheetChange`: hívás a beillesztés-védelem UTÁN, `B3:B82` metszetnél |
| `vba/mod_FormatMonthSheet.bas` | Két gomb minden hónaplapon: „Mitarbeiter entfernen” (O7:P7), „Personal-ID korrigieren” (Q7:R7) |
| `TEST_CASES.md` | **TEST 27** (A–D forgatókönyv + negatív ellenőrzések) |

**Két szint:**

1. **Kemény:** egy hónaplapon belül duplikált ID → a **most beírt** cella törlődik, a meglévő sor érintetlen.
   Több soros beillesztésnél az első marad, a többi elutasításra kerül, egyetlen összesített üzenettel.
2. **Puha:** ha a beírt ID más hónapban **más névhez** tartozik → figyelmeztetés a javító makróra.
   Csak egyetlen kézi beírásnál fut (beillesztésnél nem), és csak ha ID és név is ki van töltve.

### Igény (teszt visszajelzés, 2026-08-11)

> „Egy dolgozó = 1 ID. Be se lehessen írni ugyanazt az ID-t, ami már be van írva egy dolgozóhoz.”

### Szabály

- Hatókör: **egy hónaplapon belül** a `B3:B82` tartomány.
- Több hónapon ugyanaz az ID ugyanannál a dolgozónál **helyes** — nem duplikátum.
- Üres cella megengedett.
- Összehasonlítás normalizálva (`Trim$`, `UCase$`, szövegként).

### Viselkedés ütközéskor (eldöntve)

- [x] **Elutasítás**: a most beírt B cella tartalma törlődik, a meglévő sor változatlan.
- [x] Az üzenet tartalmazza a sort, az ID-t, a már foglalt sort és a hozzá tartozó nevet.

### Fontos korlátok

- CopyData futásakor az események ki vannak kapcsolva → a másolást nem blokkolja (helyes).
- **Nem véd az elgépelés ellen**, ami egyedi értéket ad (pl. `10457` → `10475`) —
  ezért TR-06 önmagában is szükséges.
- Több cellás beillesztésnél minden érintett cellát ellenőrizni kell.
- Előfeltétel: TR-02 (szöveg formátum), különben `00123` és `123` külön értéknek látszik.

### Elfogadási kritériumok

- [ ] Duplikált ID kézi beírásnál nem marad a lapon észrevétlenül.
- [ ] CopyData és a javító makró (TR-06) nem akad el rajta.
- [ ] Beillesztés (több cella) is ellenőrzött.

---

## TR-09 — Hibabejelentő gomb + akciónapló

**Státusz:** 🟩 Implementálva (2026-08-11) — manuális teszt (TEST 28) nyitott  
**Prioritás:** 🟡 Közepes — a 2027-ig tartó tesztfázis kiszolgálása  
**Modul(ok):** **új** `mod_PIDFeedback.bas`, **új** `mod_PIDActionLog.bas`

### Cél

A tesztelő kollégák „nem működik" típusú visszajelzése helyett egy gombnyomással
összeálljon minden, ami a hiba visszajátszásához kell.

### Megvalósult (2026-08-11)

| Fájl | Mi történt |
|------|------------|
| `vba/mod_PIDFeedback.bas` | `FehlerMelden` — kontext + 2 kérdés → txt fájl a `Feedback` mappába + vágólap |
| `vba/mod_PIDActionLog.bas` | `PID_TrackAction` → rejtett `PID_ACTION_LOG` lap (max 500 sor), `PID_AdminShowActionLog` |
| `vba/mod_FormatMonthSheet.bas` | 3. gomb a hónaplapokon: „Fehler melden” (S7:T7) |
| `vba/mod_PIDAdminSheet.bas` | „Fehler melden” + „Aktionsprotokoll” gomb, `PID_ADMIN_BTN_COUNT` 16 → 18 |
| `vba/mod_DataClear.bas`, `mod_MitarbeiterPflege.bas`, `mod_PersonalIdUnique.bas` | `PID_TrackAction` hívások |
| `TEST_CASES.md` | **TEST 28** |

### A jelentés tartalma

Verzió, fájl, mappa, Excel verzió/build, felhasználó, év, aktív lap, kijelölés,
számítási mód, a felhasználó két válasza, és az utolsó 15 naplóbejegyzés.

### Eldöntött kérdések

- [x] **E-mail:** `PID_FB_MAIL_TO = adam.nagy@at.mcd.com`, elsődlegesen **Outlook**
      (késői kötés, a txt csatolmányként, `Display` és nem `Send`), tartalék a `mailto`.
- [x] **CopyData naplózása:** jóváhagyva — egyetlen `PID_TrackAction` sor a
      `PID_CopyDataToFollowingMonths` sikeres ágán, a `PID_HideUnwantedTechnicalSheets`
      után, a `CleanExit` előtt. Másolási logika érintetlen.
- [ ] Verzió (`PID_FB_VERSION`) kézzel frissítendő release-nél.

---

## TR-05 — Alap gyorsítás & egyszerűsítés áttekintés

**Státusz:** ⬜ Nyitott  
**Prioritás:** 🟡 Közepes — TR-01 után, TR-02 előtt  
**Típus:** Review + mérés; csak kis, biztonságos finomítások  
**Modul(ok):** első körben nincs kötelező kód — `mod_PerformanceBaseline.bas`, `docs/PERFORMANCE_BASELINE.md`, `docs/FUTURE_PLANS.md`

### Cél

- Összhangban a teszt visszajelzéssel: a sheet **jó**, de érdemes tudni, hol lehet még finomítani (sebesség, UX, kevesebb kattintás).
- **Nem** nagy refactor; **nem** CopyData logika (az TR-01).

### 1. fázis — mérés és gyűjtés

- [ ] **FP-010 MANU** lépések teszt gépen (stoppóra):
  - Cold Open → első hónap használható
  - CopyData (forrás hónap → december)
  - Save után várakozás
  - LOHNTABELLE → első hónap F-dropdown
- [ ] Admin: `PID_RunPerformanceBaseline` — log: `PID_PERFORMANCE_LOG`
- [ ] Tesztelői kérdés: *„Mi érződik lassúnak?”* (megnyitás, CopyData, KV, Fluktuation tab, Full Refresh)
- [ ] Eredmények rögzítése: `docs/PERFORMANCE_BASELINE.md` táblák

### 2. fázis — finomítási jelöltek (csak ha mérés indokolja)

| Terület | Lehetséges lépés | Kapcsolat |
|---------|------------------|-----------|
| CopyData formátum | B/C `@` másoláskor (opcionális) | TR-02 |
| Open | FP-027 manuális Win/Mac lezárása | FUTURE_PLANS |
| UX | TR-03 nullázás gomb — kevesebb workaround | TR-03 |
| Full Refresh | Mikor kell admin refresh vs. automatikus dirty | Admin |
| Mac F-dropdown | FP-026 — post-release, alacsony prio | FUTURE_PLANS |

### Amit most **nem** csinálunk

- CopyData override/propagáció átírása (TR-01)
- `mod_ResetAndImportVBAFiles` érintése
- Cél nélküli „minden gyorsabb” refaktor

### Elfogadási kritériumok

- [ ] MANU baseline kitöltve legalább 1 Win gépen (opcionálisan Mac)
- [ ] Rövid lista: top 3 „érzett lassúság” vagy UX súrlódás
- [ ] Döntés: mely finomítások mennek implementációba (külön FP/TR)

### Meglévő alap (már kész — ne törjük el)

- FP-005 scoped KV refresh (~0,15 s)
- FP-007/008 SheetChange / SelectionChange
- FP-009 Fluktuation deferred
- CopyData: `PID_APPLY_FORMATS_DURING_COPY = False`, `PID_CALCULATE_FLUCTUATION_DURING_COPY = False`

---

## TR-10 — Dolgozó törlésekor a képletek is törlődtek

**Státusz:** 🟥 Nyitott — a kód megvan, de a munkafüzetben nem fut le
(lásd „KÖVETKEZŐ FELADAT" a dokumentum elején)  
**Prioritás:** 🔴 Magas (adatvesztés-szerű: a sor véglegesen „halott" marad)  
**Modul(ok):** `mod_DataClear.bas`, `mod_MitarbeiterPflege.bas`, `Modul1.bas`

### Tünet (teszt visszajelzés)

- Ahonnan dolgozót töröltünk, abból a sorból a képletek is eltűntek.
- A sor később sem éled fel: új dolgozó beírásakor G/H/K/L üres marad.

### Megtalált ok

Mind a három törlési útvonal (`PID_ClearMonthInputAreas`, `PID_ClearOnlySelectedEmployeeRows`,
`PID_MPClearEmployeeRows`) egyben ürítette a `B:N` tartományt — abban viszont ott van a négy
képletoszlop is: **G** (Monatslohn), **H** (Aktuelle Stunden), **K** (Urlaub Euro),
**L** (Letztes Gehalt).

Utána csak G-t próbálta visszaállítani, és az sem működött: a
`PID_EnsureMonatslohnFormulasOnSheet` a `PID_MonthSheetHasMonatslohnFormula` alapján kilép,
ha a lapon **bármelyik** sorban megvan még a képlet — egyetlen kiürített sornál ez mindig igaz.
H, K és L visszaállítása pedig sehol nem történt meg.

### Megoldás

- A képletoszlopokat **nem töröljük többé**. Új központi helper:
  `PID_GetEmployeeInputCellsForRows(ws, elsőSor, utolsóSor)` → `B:F`, `I:J`, `M:N`.
  Mindhárom törlési útvonal ezt használja.
- A képletek B/C-guardja miatt a sor magától üresnek látszik, ha nincs benne dolgozó.
- Javítóháló: `PID_RestoreFormulaColumnsForRows` — ha az érintett sorban G/H/K/L-ből
  hiányzik a képlet, soronként pontosan visszateszi (`HasFormula` ellenőrzés, meglévő
  képlethez nem nyúl). Így a régi verzió által kiürített sorok a következő törléskor
  maguktól meggyógyulnak.
- A megerősítő dialógusok a tényleges tartományokat írják, és a képletoszlopokat
  „megmarad" tételként sorolják fel.

### Már meglévő sérülés javítása — ⚠️ NEM MŰKÖDIK (2026-08-11 teszt)

- **Admin → Full Refresh** (`PID_FullSystemRefresh`) elvileg mind a 12 hónaplapon újraírja
  G/H/K/L képleteit, a gyakorlatban viszont a sérült sorok üresen maradtak.
  Okok és teendők: lásd „KÖVETKEZŐ FELADAT" a dokumentum elején.

### Elfogadási kritériumok

- [ ] `MitarbeiterEntfernen` után G/H/K/L-ben ott a képlet (szerkesztőlécen ellenőrizve).
- [ ] Ugyanabba a sorba új dolgozót írva a lohn/stunden/urlaub/letztes Gehalt azonnal számol.
- [ ] `DataClear` után mind a 80 sorban van képlet.
- [ ] Formátumok, zebra, keretek, lapvédelem változatlan.

---

## TR-02 — Bemásolás mindig csak érték (formázás nélkül)

**Státusz:** 🟩 Kód kész — teszt nyitott  
**Prioritás:** 🟡 Közepes  
**Modul(ok):** `mod_PIDPasteValues.bas` (új), `DieseArbeitsmappe.cls`, `Modul1.bas`

### Tünet (teszt visszajelzés)

- Bemásoláskor néha bejön a forrás formázása (szín, betűtípus, szám/dátum formátum).
- „Nem mindig" — vagyis ugyanaz a művelet hol tiszta, hol formázott lett.

### Megtalált ok

1. A régi védelem `Application.Undo`-ra épült, de a `Workbook_SheetChange`-ben **előtte** futott
   VBA-írás (F-dropdown újraépítés az E oszlopnál). Minden VBA-írás **kiüríti az Excel undo-vermét**,
   így az `Application.Undo` már nem csinált semmit → a formázás bent maradt.
2. A védelem csak a hónaplapok néhány tartományára futott (`B3:C82`, `D`, `E:F`, `I:J`, `M:N`, `O18:Q25`),
   minden más cellába és minden más lapra formázottan lehetett beilleszteni.
3. Reaktív volt: előbb megtörtént a formázott beillesztés, utána próbáltuk visszavonni.

### Megoldás — két szint

**1. Megelőzés (`mod_PIDPasteValues.bas`):**

- `Application.OnKey` elfogja: `Ctrl+V`, `Ctrl+Shift+V`, `Shift+Insert` → `PID_PasteValuesOnly`.
- Excelen belüli másolás → `PasteSpecial xlPasteValues` (csak érték).
- Külső forrás (Word, böngésző, másik Excel-példány) → a vágólap szövegét **saját kezűleg**
  írja be cellánként (nyelvi verziótól független, formázás nem jöhet át).
- `Ctrl+X` + beillesztés le van tiltva üzenettel: a mozgatás formulákat és hivatkozásokat törhet.
- **Zárolt cella védelem:** a lapok `UserInterfaceOnly:=True` védelemmel futnak, ezért a VBA
  beleírhatna a zárolt formulaoszlopokba is. Beillesztés előtt ellenőrzi a célterületet
  (kijelölés + a vágólap blokkja), és zárolt cella esetén üzenettel elutasít.
- Az `OnKey` az egész Excel-alkalmazásra hat → `Workbook_Activate`-nél beáll,
  `Workbook_Deactivate` / `Workbook_BeforeClose`-nál visszaáll az alapértelmezett Ctrl+V.

**2. Háló (`EnforcePasteValuesOnly` a `DieseArbeitsmappe.cls`-ben):**

- A menüszalag / jobbklikk / húzás útján érkező beillesztést takarítja utólag.
- A hívás a `SheetChange` **legelejére** került (minden VBA-írás elé) → az undo-verem ép.
- Hatóköre: **az egész munkafüzet minden cellája**, nem csak a hónaplap input tartományai.
- Kihagyja magát, ha épp a saját Ctrl+V fut (`PID_IsManagedPasteRunning`).
- Biztonsági korlátok: több területű kijelölésnél nem nyúl bele, és 20 000 cella fölött kilép
  az `Application.Undo` **előtt** (nehogy adat vesszen).

### Szándékosan kimaradt

- **Nincs** kényszerített `NumberFormat = "@"` a B/C oszlopon. A cél a formázás átvitelének
  megakadályozása volt; az ID text-formátumúvá alakítása külön kérdés (érintené a
  `PID_HOUR_OVERRIDES` kulcsokat és a `00123` / `123` összehasonlítást) — lásd TR-07 jegyzet.
- **Nincs** `Application.CellDragAndDrop = False`. A húzás/kitöltő fogantyú átviszi a formázást,
  de a letiltása a normál kitöltést is elvenné. A háló utólag ezt is kitakarítja.

### Elfogadási kritériumok

- [ ] Színes/félkövér szöveg beillesztése Wordből → tiszta érték, a cella eredeti formázása marad.
- [ ] Másik Excel-fájlból (más formátum) beillesztés → csak érték.
- [ ] Menüszalag „Beillesztés" gomb és jobbklikk → szintén csak érték.
- [ ] E oszlopba beillesztés után is tiszta (ez volt a régi hibás eset).
- [ ] Formulaoszlop (G/H/K/L) nem írható felül széles beillesztéssel.
- [ ] Ctrl+V másik megnyitott Excel-fájlban változatlanul működik (fájlváltás után).
- [ ] Normál kézi beírás és dropdown nem törik el.

---

## TR-03 — „Minden adat törlése” gomb (1. fázis)

**Státusz:** ⬜ Nyitott  
**Prioritás:** 🟡 Közepes  
**Modul(ok):** `mod_DataClear.bas` (kiterjesztés), esetleg `mod_PIDAdminSheet.bas` (gomb)

### Igény (teszt visszajelzés)

- Egy gomb, ami **minden adatot** kitöröl — nulláról induljon a sheet kitöltése.

### Meglévő alap

- `PID_ClearCurrentMonthData` — **egy** hónapot töröl (B:N, O18:Q25, Q31, O45).
- `PID_ResetHourOverrideLog` — stunden override log ürítése (adminban már elérhető).

### Elméleti terv — `PID_ClearAllWorkbookData` (munkanév)

- [ ] Erős megerősítő dialógus (dupla confirm).
- [ ] Mind a 12 hónaplap:
  - B3:N82 tartalom
  - O18:Q28 panel (frissített tartomány)
  - Q31, O45
- [ ] Formulák megmaradnak (G, H, K, L).
- [ ] `PID_HOUR_OVERRIDES` log ürítése.
- [ ] Dirty flag-ek: Fluktuation, KV, Finanz összesítő.
- [ ] **NEM** törlődik: EINSTELLUNG, LOHNTABELLE, UEBERSICHT sablon, védelem, makrók.
- [ ] Gomb: `_ADMIN` lapon vagy dedikált user-facing hely (döntés szükséges).

### Elfogadási kritériumok

- [ ] Egy kattintás (+ megerősítés) után minden hónap üres input állapotban.
- [ ] Struktúra, formulák, védelem, KV-tábla érintetlen.
- [ ] CopyData és Fluktuation működik az üres sheet-en.

---

## TR-04 — „Új év indítása” (2. fázis)

**Státusz:** ⬜ Nyitott  
**Prioritás:** 🟢 Alacsony — TR-03 után, külön feature  
**Modul(ok):** új makró (pl. `PID_StartNewYear`), **ne** keverjük TR-03-mal

### Igény (teszt visszajelzés)

- Megkérdezi a következő évet.
- Új fájl/munkafüzet: minden adat nullázva.
- Év átvitele: `EINSTELLUNG!C35`.
- December dolgozói → következő év **január** (név, ID, KV — **órák/dátumok nélkül**).
- Eredeti évi fájl archívumként megmarad.

### Elméleti lépések

- [ ] Dialógus: új év → `EINSTELLUNG!C35`.
- [ ] `SaveAs` új fájlnévvel (pl. `Personalsheet_2027.xlsm`).
- [ ] TR-03 logika (teljes nullázás) az új fájlban.
- [ ] December → január átemelés: B, C, D (opcionálisan M/N); **nem** E, F, I, J.
- [ ] Kilépési dátummal rendelkező dolgozók szűrése decemberből.
- [ ] Override log ürítése.
- [ ] Fluktuation / UEBERSICHT frissítés.

### Kockázatok / döntések

- [ ] CopyData bootstrap érintetlen marad — külön, szűkebb átemelő makró.
- [ ] UX: SaveAs vs. in-place (javasolt: SaveAs + archívum).
- [ ] Kurzanleitung / manager-dokumentáció frissítése.

---

## Kapcsolódó, még nem bejelentett tételek

Ezeket a teszt során érdemes figyelni; külön ticket, ha előjönnek:

| ID | Megjegyzés |
|----|------------|
| TR-XX | CopyData panel tartomány spec: O18:Q25 vs. lock O18:Q28 — szinkron ellenőrzés |
| TR-XX | `PID_APPLY_FORMATS_DURING_COPY = False` — másoláskor formátumok nem frissülnek (TR-02 kapcsolat) |

---

## Implementációs sorrend (javaslat)

1. **TR-06 + TR-08** — javító és eltávolító makró (azonnali gyakorlati megoldás, nem érinti a CopyData-t) ✅ kód kész
2. **TR-07** — ID egyediség-ellenőrzés (megelőzés) ✅ kód kész
3. **TR-02** — beillesztés csak értékként ✅ kód kész
4. **TR-01** — CopyData logika fix (gyökeres megoldás, bootstrap modul, jóváhagyás kell)
5. **TR-05** — perf/UX áttekintés + FP-010 mérés (párhuzamosan gyűjtéssel is mehet)
6. **TR-03** — nullázás gomb (gyorsítja a tesztelést és az újraindítást)
7. **TR-04** — új év (önálló release feature, 2027 release előtt)

**Változás indoka (2026-08-11):** TR-06 + TR-07 kombinációja a gyakorlatban megoldja a
TR-01 fájdalmát anélkül, hogy a `mod_CopyData.bas`-hoz hozzá kellene nyúlni.
TR-01 így kényelmi kérdéssé válik, nem blokkolóvá.

---

## Jóváhagyás / korlátok

- `mod_CopyData.bas` — **explicit jóváhagyás** szükséges (TR-01).
- `mod_ResetAndImportVBAFiles.bas` — **soha** módosítani jóváhagyás nélkül.
- Minden TR után: `docs/CHANGELOG.md` + Smoke teszt + commit csak kérésre.

---

## Verziózás (tervezett FP-kódok)

| TR ID | Tervezett FP/TR kód | CHANGELOG szekció |
|-------|---------------------|-------------------|
| TR-01 | FP-031 | Fixed |
| TR-06 | FP-036 | Added |
| TR-07 | FP-037 | Added |
| TR-08 | FP-038 | Added |
| TR-09 | FP-039 | Added |
| TR-05 | FP-035 | Changed / docs |
| TR-02 | FP-032 | Fixed |
| TR-10 | FP-040 | Fixed |
| TR-03 | FP-033 | Added |
| TR-04 | FP-034 | Added |
