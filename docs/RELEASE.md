# PERSONALSHEET — Release-Checkliste

Vor **jeder Version**, die in Restaurants ausgerollt wird.

Verknüpfungen: [`FUTURE_PLANS.md`](FUTURE_PLANS.md) · [`CHANGELOG.md`](CHANGELOG.md) · [`TEST_CASES.md`](../TEST_CASES.md) · [`PERFORMANCE_BASELINE.md`](PERFORMANCE_BASELINE.md)

---

## Erste Release (v1.0) — wann?

**Geplant:** Erstes Restaurant-Release **nach** Umsetzung der Excel-2016-Test-Punkte **FP-017–FP-022** (siehe [`FUTURE_PLANS.md`](FUTURE_PLANS.md)).

| Status | Inhalt |
|--------|--------|
| Erledigt (Basis) | FP-001–FP-004, Smoke + manuelle Tests Excel 2016 grün/gelb (2026-05) |
| Erledigt (v1.0-Code) | FP-017–FP-022 inkl. FP-018 KV-Ladezeit (Bulk) — manuell nach Import prüfen |
| Erledigt (v1.0-Code) | FP-023 Copyright Adam Nagy / McOpCo — nach Import prüfen |
| Erledigt (Zusatz) | Schutz-Paket FP-011–FP-016, Performance FP-005–FP-010 |
| Offen vor Rollout | **FP-028** (Stunden-CopyData — Teilfix, Rest zuhause); FP-027, FP-029, FP-014 (UX/optional); FP-026 post-release Mac |

**Ablauf:** Wenn FP-017–FP-022 im Workbook verifiziert sind → diese Checkliste durchgehen → getaggte `Personalsheet.xlsm` an die Restaurants.

---

## Vorbereitung

- [ ] Git working tree sauber (`git status`)
- [ ] `docs/CHANGELOG.md` — Unreleased-Abschnitt aktuell
- [ ] VBA-Änderungen in `Personalsheet.xlsm` importiert

---

## Automatische Prüfung (in Excel)

1. **Alt+F8** → `ResetAndImportVBAFiles` (wenn nur `vba/` geändert wurde)
2. VBA-Editor → **Debug → VBAProject kompilieren** (Fehler = STOP)
3. Workbook speichern
4. Excel neu starten (nach Import empfohlen)
5. **Alt+F8** → `FullSystemRefresh`
6. **Alt+F8** → `PID_RunSystemSmokeCheck`
   - TEST 7 = **PASS** (Pflicht)
   - TEST 9–16 = strukturell **PASS**, wo automatisch (Jahr, A1, UEBERSICHT, KV, Monatslohn)
   - TEST 1–6, 8 = **REVIEW** OK, aber manueller Test nötig

---

## Manuelle kritische Tests

Siehe [TEST_CASES.md](../TEST_CASES.md):

- [ ] **TEST 1** — Stundenänderung in der Zukunft (CopyData)
- [ ] **TEST 2** — Austritt (Austrittsdatum)
- [ ] **TEST 3** — Neuer MA über Monatsgrenze (CopyData rückwärts)

---

## Schutz-Paket (FP-011–FP-013)

- [ ] **TEST 17** — `PID_RunSystemSmokeCheck` → TEST 17 PASS (G3 gesperrt, E3/B3/I3 frei)
- [ ] **TEST 18** — TEST 18 PASS (Q12: Februar gesperrt, Januar frei)
- [ ] Manuell: Januarblatt — E/F-Dropdown funktioniert, G nicht editierbar (Meldung erscheint)
- [ ] Manuell: Sortieren auf Monatsblatt → blockiert (Excel-Meldung)
- [ ] Manuell: Fill-Handle von E/F ziehen → funktioniert nicht / kein Layout-Schaden (FP-011)

---

## Visueller Spot-Check

- [ ] **UBERSICHT:** FINANZIELL-Block, Durchrechnung E30/I30 weiß und editierbar
- [ ] **EINSTELLUNG:** gelb = nicht editierbar, weiß = Eingabe
- [ ] **Ein Monatsblatt:** E/F-Dropdown OK, Monatslohn wird berechnet, Q31 Fluktuation nach Austrittsdatum plausibel
- [ ] **FLUKTUATION:** Dashboard lädt (nach relevanter Änderung Tab öffnen)

---

## Release (v1.0)

- [ ] Offene FP in [`FUTURE_PLANS.md`](FUTURE_PLANS.md) geprüft (Release-Blocker vs. „danach“)
- [ ] `docs/CHANGELOG.md`: Unreleased → neuer Versionsabschnitt (Datum)
- [ ] `Personalsheet.xlsm` speichern
- [ ] Git commit + Tag (z. B. `v1.0.0` oder `v2026.06.xx`)
- [ ] Backup: vorherige Version bleibt erhalten (OneDrive-Versionen / Kopie)
- [ ] Restaurants: **nur** die getaggte `.xlsm` installieren

---

## Rollback

1. Vorherige getaggte `Personalsheet.xlsm` wiederherstellen
2. **Nicht** `ResetAndImportVBAFiles` auf alter Datei mit neuem `vba/`-Stand mischen
3. Bei Datenproblem: Restore aus Restaurant-eigener Sicherung

---

## Admin-Makros — **niemals** an Restaurant-User

- `ResetAndImportVBAFiles`
- `RebuildLOHNTABELLE` / `RestoreLOHNTABELLEBase2025_2026`
- `UnprotectEverything`

Siehe auch Hinweis in `mod_PIDAdmin.bas` und [`README.md`](../README.md).
