# FUTURE PLANS — PERSONALSHEET

Technischer Backlog und Umsetzungsstand.  
**Zuerst die Übersicht lesen** — Details zu erledigten Punkten stehen unten im Archiv.

Verknüpfungen: [`CHANGELOG.md`](CHANGELOG.md) · [`PERFORMANCE_BASELINE.md`](PERFORMANCE_BASELINE.md) · [`RELEASE.md`](RELEASE.md)

---

## Status-Legende

| Symbol | Bedeutung |
|--------|-----------|
| 🔴 **Offen** | Noch **nicht** umgesetzt — Code- oder Doku-Arbeit nötig |
| 🟡 **Teilweise** | Code **fertig**, aber manuelle Tests / Restarbeit offen |
| 🟢 **Erledigt** | Umgesetzt und für Release akzeptiert (evtl. einzelne optionale Tests offen) |
| ⚫ **Storniert** | Bewusst nicht umgesetzt |
| 🔵 **Mac-only** | Post-release, nur Mac Excel (Adam) |

**Priorität:** Hoch → Mittel → Niedrig (innerhalb „Offen“).

---

## Übersicht — alle FP-Einträge

### 🔴 Offen (noch umsetzen)

| ID | Thema | Prio | Plattform |
|----|--------|------|-----------|
| [FP-028](#fp-028--copydata--monatsstunden-stunden-override-log) | Stunden: nach CopyData zurück auf Original (Teilfix) | **Hoch** | Win + Mac |
| [FP-027](#fp-027--workbook-open-flackern--recalc) | Open: Flackern / Recalc mehrere Sekunden | Mittel | Win (2016), ggf. 365 |
| [FP-029](#fp-029--spalte-k-urlaub-in--0-in-leerer-zeile) | Spalte K (Urlaub €): 0 statt leer | Mittel | Win + Mac |
| [FP-014](#fp-014--enableselection-nur-entsperrte-zellen) | `EnableSelection` auf Monatsblättern | Mittel | Win + Mac |
| [FP-026](#fp-026--mac-only-f-dropdown-performance) | Mac F-Dropdown Performance | Niedrig | **Nur Mac** (post-release) |

### 🟡 Erledigt — manuelle Tests / Rest offen

| ID | Thema | Was fehlt noch |
|----|--------|----------------|
| FP-010 | Performance-Messprotokoll | MANU: Cold Open, CopyData, Save (Stoppuhr) |
| FP-016 | Schutz-Smoke | Sort/Fill-Handle bei jedem Release manuell |
| FP-018 | KV-Periode → erster Monats-Tab | Akzeptable Zeit nach neuer Periode messen |
| FP-001–004, 006, 011, 013, 017, 019–024 | diverse Fixes | Einzelne Mac/Win-Spot-Checks in Akzeptanzkriterien |

→ Details: [Archiv — Teilweise](#archiv-teilweise-manuelle-restpunkte)

### 🟢 Erledigt (Archiv)

FP-001 · FP-002 · FP-003 · FP-004 · FP-005 · FP-006 · FP-007 · FP-008 · FP-009 · FP-011 · FP-012 · FP-013 · FP-015 · FP-016 (auto) · FP-017 · FP-018 (Code) · FP-019 · FP-020 · FP-021 · FP-022 · FP-023 · FP-024

→ Kurzfassung: [Archiv — Erledigt](#archiv-erledigt)

### ⚫ Storniert

| ID | Thema |
|----|--------|
| FP-025 | FLUKTUATION PDF-Export (entfernt 2026-06-12) |

---

## Empfohlene Reihenfolge (nächste Schritte)

1. **FP-028** — Stunden-CopyData (Teilfix offen, **zuhause** abschließen)
2. **FP-029** — UX K-Spalte (schnell, analog FP-003/L)
3. **FP-027** — Open-Performance (nach Messung Win2016 + Win365)
4. **FP-014** — EnableSelection Monatsblätter (Mac testen)
5. **FP-026** — erst nach v1.0-Release, Mac-only

---

# Offen — Details

## FP-028 — CopyData / Monatsstunden: Stunden-Override-Log

**Status:** 🟡 Teilfix — Win2016-Test 2026-06 **nicht grün** · Rest **zuhause**  
**Priorität:** Hoch (Release-Blocker)  
**Plattform:** Windows + Mac · `mod_CopyData.bas`

### Problem (Original)

1. MA ab Juli mit 173 h — OK.  
2. Ab Oktober weniger + Aktualisierung ab Mai — OK.  
3. Oktober zurückdrehen / neu ändern → springt auf **erste** Änderung zurück.

### Feedback nach Teilfix (2026-06, Win2016)

- Im **Änderungsmonat** wirkt die Korrektur zunächst — Wert bzw. Log-Eintrag bleibt dort sichtbar.  
- Nach **CopyData** („Aktualisierung des restlichen Jahres“) springt **Monatsstunden (F)** in den Folgemonaten wieder auf den **ursprünglichen** KV-Wert zurück.  
- Verstecktes Blatt **`PID_HOUR_OVERRIDES`** kann den neuen Wert trotzdem führen — Blatt und Log weiter **desynchron**.

### Bereits umgesetzt (Teilfix, uncommitted)

| # | Änderung |
|---|----------|
| 1 | `PID_ApplyLoggedHourOverrides` aus CopyData-Cascade entfernt |
| 2 | `CollectFutureOverrides`: E/F aus Folgemonats-Blatt (`PID_CollectFutureEFOverrideIfChanged`) |
| 3 | Nach CopyData: `PID_ReconcileHourOverrideLogFromMonthSheets` |

→ Reicht für Win2016-Szenario **noch nicht** — weiter debuggen zuhause.

### Nächste Schritte (zuhause)

- [ ] Repro Schritt für Schritt: welcher Monat als Quelle, welche Folgemonate falsch  
- [ ] Prüfen: `PID_BuildTargetMonthData` / `currentValues`-Rollforward überschreibt F?  
- [ ] Prüfen: SheetChange-Log vs. Reconcile-Reihenfolge  
- [ ] Ggf. Log bei F-Änderung in Monat M → Einträge für Monate **> M** invalidieren  
- [ ] Regression: November-Sonderwert, zwei Okt-Änderungen hintereinander

### Akzeptanzkriterien

- [ ] Juli 173 → Okt weniger → zurück 173 → CopyData ab Mai **und** ab Okt: überall 173
- [ ] Zwei nacheinander verschiedene Okt-Werte — **letzter** bleibt
- [ ] November-Sonderwert bei CopyData ab Mai bleibt (Regression)
- [ ] Smoke + CopyData-SPEC grün

### Workaround (bis Fix)

Aktualisierung immer vom **Monat der letzten Änderung** aus starten (nicht von früherem Monat).

---

## FP-027 — Workbook-Open: Flackern / Recalc

**Status:** 🔴 Offen — Feedback Excel 2016 (2026-06)  
**Priorität:** Mittel  
**Plattform:** Primär Win2016; Mechanismus betrifft alle Versionen

### Problem

Beim Öffnen: Excel flackert / „Berechnet…“ mehrere Sekunden (Win2016). Funktion danach OK.

### Ursache (vermutet)

Open-Ende: `xlCalculationAutomatic` → Full-Workbook-Recalc; plus Copyright, Schutz, H/K/L auf aktivem Tab.

### Geplante Ansätze

| # | Ansatz |
|---|--------|
| A | Messung FP-010 Schritt 1 auf Win2016 **und** Win365 |
| B | Open-scoped Recalc (nur aktives Blatt H/K/L) |
| C | Lazy Tab-Recalc konsistent |
| D | Deferred Copyright/Schutz |

### Akzeptanzkriterien

- [ ] Win2016: Open spürbar kürzer (FP-010 messen)
- [ ] Win365 + Mac: keine Regression
- [ ] H/K/L nach Open korrekt (FP-024)

---

## FP-029 — Spalte K (Urlaub in €): 0 in leerer Zeile

**Status:** 🔴 Offen — Feedback (2026-06)  
**Priorität:** Mittel  
**Plattform:** Win + Mac · Spalte **K**, `Modul1.bas` / CopyData

### Problem

Leere MA-Zeilen (kein Name/ID) zeigen **0,00 €** in K statt leerer Zelle.

### Ansatz

Analog **FP-003** (Spalte L): B/C-Leer-Guard, J leer → K leer, 0 → `""`, kanonische Formel in `Modul1.bas`.

### Akzeptanzkriterien

- [ ] Leere Zeile: K ohne 0
- [ ] MA ohne J: K leer
- [ ] MA + J + Lohn: K korrekt (€)
- [ ] CopyData propagiert Formel

---

## FP-014 — EnableSelection (nur entsperrte Zellen)

**Status:** 🔴 Offen (Code auf Monatsblättern **nicht** umgesetzt; Kurzanleitung FP-015 erledigt)  
**Priorität:** Mittel  
**Plattform:** Win + Mac · `mod_SchutzHinzufugen.bas`

### Problem

Nutzer markieren große Bereiche inkl. gesperrter Formelzellen → höheres Fehlerrisiko.

### Geplant

`EnableSelection = xlUnlockedCell` auf **Monatsblättern** nach Protect (UBERSICHT hat bereits einen Pfad).

### Akzeptanzkriterien

- [ ] E/F und Whitelist normal nutzbar
- [ ] CopyData, Format, L-Restore mit `UserInterfaceOnly:=True`
- [ ] Mac 2016: keine UX-Regression (Dropdowns, Panel)
- [ ] Windows gleicher Smoke

---

## FP-026 — Mac-only: F-Dropdown Performance

**Status:** 🔵 Geplant — **post-release v1.0**  
**Priorität:** Niedrig  
**Plattform:** Nur Mac · Windows-Pfad **nicht** anfassen

### Ist-Zustand (Mac, 2026-06)

E/F-Aktionen ~0,5 s (voller F-Rebuild). Zuverlässig nach KV-Fixes, aber langsam.

### Geplant

Sor-/KV-szintű refresh; Stunden-Cache mit strikter Invalidierung; weniger Doppel-Refresh.

### Akzeptanzkriterien

- [ ] Mac smoke: BG1 ↔ BG1_5, Eigene Stunden, LOHNTABELLE-View
- [ ] ~0,1 s pro E/F-Aktion (Ziel)
- [ ] Windows Baseline 2b unverändert (~0,15 s)

---

# Archiv

## Archiv — Erledigt

Kompakte Liste. Vollständige Historie in [`CHANGELOG.md`](CHANGELOG.md).

| ID | Status | Kurz | Datum |
|----|--------|------|-------|
| FP-001 | 🟢 | O18:Q25 löst kein FINANZIELL-Flackern mehr aus | 2026-05-25 |
| FP-002 | 🟢 | CopyData kopiert Panel O18:Q25 in Folgemonate | 2026-05-25 |
| FP-003 | 🟢 | Spalte L: 0 → leer (B/C-Guard) | 2026-05-25 |
| FP-004 | 🟢 | F-Dropdown nach Eigene Stunden ohne E-Re-Select | 2026-05-25 |
| FP-005 | 🟢 | Scoped KV-dirty F-Refresh (0,15 s vs 0,52 s) | 2026-06-12 |
| FP-006 | 🟢 | `KV_DD_*` Names entfernt (`PID_CountKVDDNamedRanges` = 0) | 2026-06-12 |
| FP-007 | 🟢 | SheetChange: gebündeltes H/L/G-Recalc | 2026-06-12 |
| FP-008 | 🟢 | SelectionChange entlastet | 2026-06-12 |
| FP-009 | 🟢 | Fluktuation: Save nur Daten; Analyse beim Tab | 2026-06-12 |
| FP-010 | 🟢 AUTO | Baseline-Makro + Doku; MANU-Stoppuhr offen | 2026-06-12 |
| FP-011 | 🟢 | Fill Handle / Drag auf Monatsblättern aus | 2026-06-12 |
| FP-012 | 🟢 | Sortieren auf Monatsblättern aus | 2026-06-12 |
| FP-013 | 🟢 | Lock-all + Whitelist (B/C, D, E/F, I/J, M/N, O18:Q25, Q12) | 2026-06-12 |
| FP-015 | 🟢 | Kurzanleitung A4 (Deutsch, Du-Form, Juni 2026) | 2026-06-12 |
| FP-016 | 🟢 | Smoke TEST 17/18 + RELEASE-Checkliste Schutz | 2026-06-12 |
| FP-017 | 🟢 | Spalten D/I Breite 13 | 2026-05-26 |
| FP-018 | 🟢 | KV-Periode: Bulk F-Refresh, HeavyMaintenance | 2026-05-26 |
| FP-019 | 🟢 | CopyData ohne Bestätigungsdialog | 2026-05-26 |
| FP-020 | 🟢 | Q12 Vormonat +/- editierbar (Nicht-Startmonate) | 2026-05-26 |
| FP-021 | 🟢 | UEBERSICHT E30/I30 editierbar | 2026-05-26 |
| FP-022 | 🟢 | LOHNTABELLE „4) Stunde löschen“ | 2026-05-26 |
| FP-023 | 🟢 | Copyright Blätter + VBA-Module (Adam Nagy / McOpCo) | 2026-05-26 |
| FP-024 | 🟢 | Automatische Berechnung nach Open; H/K/L per Tab | 2026-05-26 |

### FP-028 — CopyData Stunden (Teilfix, offen)

**Stand 2026-06:** Erster Fix in `mod_CopyData.bas` (E/F aus Blatt, Log-Reconcile) — Win2016-Test **fehlgeschlagen** (F revert). Siehe [Offen — FP-028](#fp-028--copydata--monatsstunden-stunden-override-log).

### Schutz-Paket (FP-011–016) — erledigt

| ID | Thema | Ergebnis |
|----|--------|----------|
| FP-011 | Fill Handle | `EnableFillHandle=False` auf Monats-Tabs |
| FP-012 | Sortieren | `AllowSorting:=False` |
| FP-013 | Lock-all + Whitelist | `PID_ApplyMonthSheetLockPolicy` |
| FP-015 | Endanwender-Doku | `docs/Kurzanleitung_Personalsheet_A4.html` |
| FP-016 | Smoke | TEST 17 (Lock), TEST 18 (Q12 Feb/Jan) |

**Bereits vor Release:** Paste nur Werte, versteckte Helper-Blätter, `FormatAllMonthSheets` als Layout-Reparatur.

### Performance-Paket (FP-005–010) — erledigt

Messung: [`PERFORMANCE_BASELINE.md`](PERFORMANCE_BASELINE.md) · Makro `PID_RunPerformanceBaseline`.

| Schritt | Inhalt | AUTO (2026-06-12) |
|---------|--------|-------------------|
| 2b | Scoped dirty BG1, Februar | **0,15 s** |
| 2 | Voll KV-Refresh 1 Blatt | 0,52 s |
| 3 | G-Recalc 1 Zeile | 0,02 s |
| 5 | FINANZ-Sync | 0,10 s |
| 6 | Fluktuation Save-Daten | 1,27 s |
| 7 | FullSystemRefresh | 7,98 s |

---

## Archiv — Teilweise (manuelle Restpunkte)

Diese Punkte sind **im Code umgesetzt**; offene Kästchen in den alten Akzeptanzkriterien = **manuelle Regression**, kein Blocker unless noted.

| ID | Offen (manuell) |
|----|------------------|
| FP-001 | O20 tippen ohne Flackern; S35 → UEBERSICHT |
| FP-002 | CopyData Smoke Mac |
| FP-003 | L > 0 €-Format; Mac/Win |
| FP-004 | Open schnell; F-Overrides; Mac/Win |
| FP-005 | Mac smoke |
| FP-006 | Open/Save-Zeit optional; Mac; CopyData |
| FP-007 | Grosser Paste-Block; stale FINANZ |
| FP-009 | Fachlicher Identitäts-Check nach D/I/N |
| FP-010 | MANU Cold Open, CopyData, Save |
| FP-011 | Mac Fill-Handle smoke |
| FP-013 | Mac + Win Excel 2016+ Whitelist |
| FP-016 | Sort/Fill bei Release manuell |
| FP-017 | Datum sichtbar Win2016; Mac |
| FP-018 | Erster Monats-Tab Zeit nach neuer Periode |
| FP-019 | TEST 1–5 CopyData |
| FP-020–021 | Panel/FINANZ manuell |
| FP-022 | F-Dropdown nach Löschen; Mac/Win |
| FP-023 | Druck A4 Spot-Check; Copyright nach Format |

---

## Archiv — Storniert

### FP-025 — FLUKTUATION PDF-Export

**Status:** ⚫ Storniert (2026-06-12)  
PDF-Button und Makro entfernt; Legacy-Shape wird bei Tab-Activate gelöscht.

---

## Archiv — Detailliert (Referenz)

<details>
<summary>FP-001 bis FP-004 (Früh-Fixes, eingeklappt)</summary>

### FP-001 — FINANZIELL-Sync O18:Q25

**Behoben 2026-05-25.** Freitext O18:Q25 löst keinen sofortigen FINANZIELL-Sync mehr (`mod_SumMergedCells.bas`). Q17:R29 / S35 weiterhin sofort.

### FP-002 — CopyData O18:Q25

**Behoben 2026-05-25.** Panel-Snapshot O/Q merge-sicher in Folgemonate (`PID_ReadMonthPanelSnapshot`).

### FP-003 — Spalte L leer statt €0

**Behoben 2026-05-25.** `PID_GetLetztesGehaltFormulaR1C1`: B/C-Guard, 0 → `""`.

### FP-004 — F-Dropdown nach Eigene Stunden

**Behoben 2026-05-25.** Dirty-Refresh baut F-Validation neu (`mod_KVStundenDropdown.bas`).

</details>

<details>
<summary>FP-005 bis FP-010 (Performance, eingeklappt)</summary>

### FP-005 — Scoped F-Dropdown Refresh

**Erledigt 2026-06-12.** `MarkKVDropdownDirtyForKVCode`, nur betroffene Zeilen/KV-Codes.

### FP-006 — Weniger Named Ranges

**Erledigt 2026-06-12.** `KV_DG_*` pro KV-Code; Legacy `KV_DD_*` entfernt.

### FP-007 — SheetChange Recalc

**Erledigt 2026-06-12.** Kein doppeltes L-Recalc; gebündeltes H/L.

### FP-008 — SelectionChange

**Erledigt 2026-06-12.** ScreenUpdating nur bei Dropdown-Repair.

### FP-009 — Fluktuation inkrementell

**Erledigt 2026-06-12.** Save: nur Daten; Tab: Analyse; Monats-Dirty-Rescan.

### FP-010 — Mess-Protokoll

**Erledigt AUTO 2026-06-12.** `docs/PERFORMANCE_BASELINE.md`, `PID_RunPerformanceBaseline`. MANU optional.

</details>

<details>
<summary>FP-011 bis FP-016 (Schutz-Paket, eingeklappt)</summary>

Siehe Tabelle [Schutz-Paket](#schutz-paket-fp-011016--erledigt). Volltext war in alter FUTURE_PLANS-Version; Kernthemen: Fill Handle aus, Sort aus, Lock-all/Whitelist, Kurzanleitung, TEST 17/18.

</details>

<details>
<summary>FP-017 bis FP-024 (v1.0 Excel-16-Feedback, eingeklappt)</summary>

### FP-017 — D/I Spaltenbreite

Breite 13 via `PID_ApplyMonthSheetDateColumnWidths`.

### FP-018 — KV-Periode Ladezeit

Bulk F-Dropdown, `PID_BeginHeavyMaintenance` / `EndHeavyMaintenance`.

### FP-019 — CopyData ohne Dialog

`PID_ConfirmCopyDataAction` entfernt.

### FP-020 — Q12 Vormonat

Q12:R12 in Whitelist; Feb/Mai/Aug/Nov Formel (Startmonate).

### FP-021 — UEBERSICHT E30/I30

Durchrechnungs-Planfelder editierbar.

### FP-022 — Stunde löschen

Button `4) Stunde löschen`, `DeleteCustomKVMonatsstunden`.

### FP-023 — Copyright

`mod_PIDCopyright.bas`, Blatt Zeile 2 / VBA-Modulkopf.

### FP-024 — Automatische Berechnung

Open → Automatic; H/K/L `.Calculate` bei Tab.

</details>

---

## Test-Notiz (Excel 2016, 2026-05)

Smoke grün/gelb; manuelle Tests ohne harte Fehler. Offene Punkte → **FP-028** (Blocker), FP-027/029.

**Nicht im FP-Backlog (bereits umgesetzt, 2026-06):** Q31 Fluktuation % auf Monatsblatt **sofort** bei Austrittsdatum-Änderung (I) — `PID_CalculateFluctuation` in `DieseArbeitsmappe.cls`.

---

## Weitere Einträge

Neue Punkte als **FP-03N** unten anfügen, dann in [Übersicht — Offen](#-offen-noch-umsetzen) eintragen.

**Vorlage:**

```markdown
## FP-0NN — Titel

**Status:** 🔴 Offen  
**Priorität:** …  
**Plattform:** …

### Problem
…

### Akzeptanzkriterien
- [ ] …
```
