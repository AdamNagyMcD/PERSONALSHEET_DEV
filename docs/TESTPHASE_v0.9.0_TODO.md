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

## Prioritás összefoglaló

| ID | Prioritás | Típus | Rövid név | Státusz |
|----|-----------|-------|-----------|---------|
| TR-01 | 🔴 Magas | Bug | Personal ID beragad CopyData után | ⬜ Nyitott |
| TR-05 | 🟡 Közepes | Review | Alap gyorsítás & egyszerűsítés áttekintés | ⬜ Nyitott |
| TR-02 | 🟡 Közepes | Bug | Bemásolás nem mindig sima TEXT | ⬜ Nyitott |
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

## TR-02 — Bemásolás nem mindig sima TEXT formátum

**Státusz:** ⬜ Nyitott  
**Prioritás:** 🟡 Közepes  
**Modul(ok):** `DieseArbeitsmappe.cls`, esetleg `mod_FormatMonthSheet.bas`

### Tünet (teszt visszajelzés)

- Bemásoláskor nem mindig lesz a cella sima szöveg (TEXT).
- Néha az eredeti formátum marad (pl. szám, dátum, pénznem).

### Valószínű ok (elemzés)

- `EnforcePasteValuesOnly` csak **értéket** ír vissza, **NumberFormat**-ot nem állít `@`-ra.
- `IsProbablyPaste` nem minden paste-módot érzék el (egy cella, Mac, bizonyos Undo szövegek, drag-fill).
- B/C oszlopokon nincs kötelező `@` formátum a hónaplapokon.
- Paste allowlist még **O18:Q25** — az O18:Q28 bővítés (FP lock policy) itt lehet nincs szinkronban.

### Teendők

- [ ] Paste handler: B/C (esetleg M/N) mindig `NumberFormat = "@"` beillesztés után.
- [ ] `IsProbablyPaste` erősítése (több nyelv, Mac).
- [ ] Allowlist frissítése: `O18:Q25` → `O18:Q28`.
- [ ] Opcionális: engedélyezett tartományban **minden** SheetChange után formátum-ellenőrzés, nem csak paste-nél.
- [ ] Teszt: Personal ID `00123`, dátum-szerű szöveg, több cellás paste, Mac + Win.

### Elfogadási kritériumok

- [ ] B/C bemásolás után mindig szöveg formátum, érték változatlan.
- [ ] O18:Q28 panel paste szintén values-only.
- [ ] Normál kézi beírás nem tör el.

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

1. **TR-01** — reprodukció + fix (legkritikusabb tesztblokkoló)
2. **TR-05** — perf/UX áttekintés + FP-010 mérés (párhuzamosan gyűjtéssel is mehet)
3. **TR-02** — paste + formátum (gyakori user friction)
4. **TR-03** — nullázás gomb (gyorsítja a tesztelést és az újraindítást)
5. **TR-04** — új év (önálló release feature, 2027 release előtt)

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
| TR-05 | FP-035 | Changed / docs |
| TR-02 | FP-032 | Fixed |
| TR-03 | FP-033 | Added |
| TR-04 | FP-034 | Added |
