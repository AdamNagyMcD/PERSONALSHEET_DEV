# FUTURE PLANS — PERSONALSHEET

Technischer Backlog für geplante, aber **noch nicht umgesetzte** Verbesserungen.  
Aktuelles Verhalten bleibt unverändert, bis ein Eintrag explizit umgesetzt und in `CHANGELOG.md` dokumentiert wird.

---

## FP-001 — FINANZIELL-Sync bei freien Texteingaben im Panel (O18:Q25)

**Status:** Zurückgestellt (bewusst so belassen)  
**Priorität:** Niedrig / UX-Optimierung  
**Betroffene Bereiche:** Monatsblätter (JANUAR–DEZEMBER), FINANZIELL-Kette, UEBERSICHT, EINSTELLUNG

### Beobachtetes Verhalten

Wird in **O20** (oder generell in **O18:Q25**) Text eingegeben, wirkt es so, als würde sich das **gesamte Blatt** aktualisieren (kurzes Flackern / sichtbare Neuberechnung im rechten Panel).

### Ursache (Ist-Zustand)

1. `Workbook_SheetChange` in `DieseArbeitsmappe.cls` ruft bei Monatsblättern bei relevanten Änderungen die FINANZIELL-Sync-Kette auf.
2. `PID_MonthChangeNeedsImmediateFinanzSync` in `mod_SumMergedCells.bas` behandelt **jede** Änderung in folgenden Bereichen als sofort synchronisationspflichtig:
   - `O18:Q25` (Personal-/Info-Block, CopyData-Propagation)
   - `Q17:R29` (FINANZIELL-Zusammenfassung rechts)
   - `S35` (Crew Labor %)
3. Folge bei Treffer: `PID_SyncFinanzSummaryForMonth` → u.a.
   - `PID_RefreshMonthFinanzSummaryCells` (Unprotect, `Calculate` auf `Q17:R29`, `S35`–`S37`, Reprotect)
   - Schreiben nach **UBERSICHT** (G/J) und **EINSTELLUNG** (E-Spalte)
   - Quartals-/Diff-Neuberechnung auf UEBERSICHT
4. Zusätzlich: `Application.ScreenUpdating = False` im SheetChange-Handler reduziert Flicker, kann aber Unprotect/Reprotect + große `Calculate`-Blöcke trotzdem als „Vollrefresh“ wahrnehmbar machen.

**Wichtig:** Es wird nicht das ganze Monatsblatt neu formatiert — es läuft die **FINANZIELL-Synchronisation**, die absichtlich eng an `O18:Q25` gekoppelt ist (siehe auch `SPEC.md` → Copy Areas).

### Warum vorerst unverändert

- Das Verhalten ist **fachlich konsistent**: Änderungen im Panel-Bereich können Crew-Labor / FINANZIELL beeinflussen; sofortiger Sync verhindert veraltete UEBERSICHT-Werte.
- Eine Verengung der Watch-Range ohne Analyse könnte **stale FINANZIELL-Daten** erzeugen.
- Kein funktionaler Bug — nur spürbare UX/Reaktionszeit bei „reinem“ Freitext in O20.

### Geplante Verbesserung (später)

**Ziel:** Freitext-/Notiz-Zellen (z. B. O20) dürfen **keine** schwere FINANZIELL-Kette auslösen; numerisch/relevante Zellen weiterhin sofort syncen.

**Mögliche Ansätze (noch nicht entschieden):**

| Ansatz | Idee | Pro | Contra |
|--------|------|-----|--------|
| A — Feinere Watch-Range | `O18:Q25` auf tatsächlich FINANZ-relevante Zellen einschränken (z. B. nur Q/R/S-Zellen mit Formeln/Werten, nicht gesamtes O–Q-Notizfeld) | Weniger Sync bei Notizen | Mapping nötig; Risiko bei vergessener Zelle |
| B — Zelltyp-Erkennung | In `PID_MonthChangeNeedsImmediateFinanzSync`: nur syncen wenn geänderte Zelle Formel hat, numerischen Wert trägt, oder in expliziter „Crew-Labor“-Whitelist liegt | Präzise Steuerung | Komplexer; Freitext in „falscher“ Zelle könnte Sync umgehen |
| C — Deferred Sync für Panel-Text | O-only-Änderungen → `MarkFinanzSummaryDirtyForMonth` statt sofortigem `PID_SyncFinanzSummaryForMonth` (Sync bei Tab-Wechsel / UEBERSICHT-Aktivierung) | Kein Flackern beim Tippen | UEBERSICHT kurz verzögert; muss mit bestehendem `gFinanzSummaryDirty`-Mechanismus abgestimmt werden |
| D — ScreenUpdating-Scope | Schwere Sync-Kette beibehalten, aber Unprotect/Calculate nur auf Minimalbereich; kein globales `ScreenUpdating` im ganzen SheetChange | Geringerer visueller Effekt | Löst nicht die Rechen-Last |

**Empfohlene Reihenfolge bei Umsetzung:**

1. OOXML / Blatt-Analyse: welche Zellen in `O18:Q25` wirklich FINANZIELL-relevant sind vs. reine Notizfelder.
2. Watch-Range + ggf. Whitelist in `mod_SumMergedCells.bas` anpassen.
3. Regression: CopyData-Propagation `O18:Q25`, TEST FINANZIELL-Sync (CHANGELOG-Einträge zu SumMergedCells / SheetChange).
4. Manuell: Text in O20 → kein sichtbarer Voll-Panel-Refresh; Änderung an S35 / Crew-Labor-Werten → UEBERSICHT weiterhin sofort korrekt.

### Akzeptanzkriterien (wenn umgesetzt)

- [ ] Freitext in dokumentierten Notiz-Zellen (mindestens O20) löst **keine** sofortige `PID_SyncFinanzSummaryForMonth`-Kette aus.
- [ ] Änderungen an Crew-Labor / FINANZIELL-relevanten Panel-Zellen syncen **weiterhin sofort** nach UEBERSICHT und EINSTELLUNG.
- [ ] `CopyData` kopiert `O18:Q25` unverändert korrekt.
- [ ] `PID_RunSystemSmokeCheck` und manuelle FINANZIELL-Checks grün.

### Betroffene Dateien (Referenz)

- `vba/DieseArbeitsmappe.cls` — `Workbook_SheetChange`
- `vba/mod_SumMergedCells.bas` — `PID_MonthChangeNeedsImmediateFinanzSync`, `PID_SyncFinanzSummaryForMonth`, `PID_RefreshMonthFinanzSummaryCells`
- `SPEC.md` — Copy Areas (`O18:Q25`)

---

## FP-002 — CopyData: O18:Q25 propagiert nicht in Folgemonate

**Status:** Behoben (2026-05-25, Snapshot O/Q) — `PID_ReadMonthPanelSnapshot` / `PID_WriteMonthPanelSnapshot`  
**Priorität:** Hoch — SPEC-konformes Verhalten fehlt  
**Betroffene Bereiche:** `CopyData`, Monatsblätter O18:Q25, Panel / Crew-Labor-Info

### Beobachtetes Verhalten (historisch)

Beim Ausführen von **CopyData** wurde der Bereich **O18:Q25** laut SPEC nicht zuverlässig in die **folgenden Monate** kopiert (Merge-Zellen; `FormulaR1C1`-Array wirkte nicht). Zusätzlich blieb auf Zielmonaten eine Markierung um O21:Q24.

### Fix (Ist-Zustand)

- `mod_CopyData.bas`: `PID_ReadMonthPanelSnapshot` / `PID_WriteMonthPanelSnapshot` — O/Q-Ankerzeilen (Merge), Quellblatt kurz entsperrt.
- `PID_ResetFollowingMonthSelections` setzt nach CopyData auf allen Zielmonaten die Auswahl auf A1 (ScreenUpdating aus).

### Akzeptanzkriterien

- [x] Nach CopyData vom Monat M sind O18:Q25 in M+1 … Dezember identisch mit Quellmonat M.
- [x] Keine sichtbare Markierung O21:Q24 auf Zielmonaten nach CopyData.
- [ ] Smoke / manueller CopyData-Test grün (Mac).

### Betroffene Dateien (Referenz)

- `vba/mod_CopyData.bas` — `PID_CopyMonthPanelBlock`, `PID_ResetFollowingMonthSelections`
- `SPEC.md` — Copy Areas

---

## FP-003 — Spalte L: bei Ergebnis 0 Zelle leer lassen (kein €0,00)

**Status:** Behoben (2026-05-25 v2) — B/C-Guard + 0→leer; Restore per `RC[-10]`-Marker  
**Priorität:** Mittel — UX / Konsistenz mit Spalte G  
**Betroffene Bereiche:** Monatsblätter Spalte L (Letztes Gehalt / Laborcost), Formel-Restore

### Fix (Ist-Zustand)

- `Modul1.bas`: L-Formel liefert bei Ergebnis 0 `""` statt sichtbarem €0,00.
- `PID_MonthSheetNeedsLetztesGehaltFormulaUpdate`: alte L-Formeln werden beim Open/Restore aktualisiert.

### Akzeptanzkriterien

- [x] Leere / irrelevante Mitarbeiterzeilen: L ohne €0,00, Zelle optisch leer.
- [ ] Zeilen mit echtem Laborcost-Wert > 0: weiterhin korrekt formatiert (€) — manuell pruefen.
- [x] CopyData ueberschreibt L in Zielmonaten weiterhin nicht (SPEC: L informational only).
- [ ] Mac + Windows Excel 2016+ kompatibel — manuell pruefen.

### Betroffene Dateien (Referenz)

- `vba/Modul1.bas` — `PID_GetLetztesGehaltFormulaR1C1`, Restore-Pfade

---

## FP-004 — LOHNTABELLE „Eigene Stunden“: F-Dropdown auf Monatsblatt erst nach erneuter E-Auswahl

**Status:** Offen (bekanntes UX-Problem, bewusst zurückgestellt)  
**Priorität:** Mittel — Workaround existiert; Fix erst nach Performance-Abwägung  
**Betroffene Bereiche:** `LOHNTABELLE`, Monatsblätter Spalte E/F, KV-Stunden-Dropdown

### Beobachtetes Verhalten

Nach **„2) Eigene Stunden“** auf **LOHNTABELLE** und Wechsel zu einem Monatsblatt (z. B. **Januar**) erscheint die **neue Stundenzahl nicht** im **F-Spalten-Dropdown**, solange in **Spalte E** der KV-Code **nicht erneut ausgewählt** wird.

**Workaround (manuell):** In betroffenen Zeilen E erneut wählen → danach zeigt F die neue Stunde in der Liste.

**Folge:** Fühlt sich wie ein **großer manueller Refresh** an (viele Zeilen / viele Mitarbeiter).

### Vermutete Ursache (Ist-Zustand)

- F-Dropdown wird **lazy** bzw. **dirty-flag-gesteuert** aktualisiert (`mod_KVStundenDropdown.bas`, `MarkAllKVDropdownsDirty`, `RefreshKVDropdownsIfDirtyForSheet` in `DieseArbeitsmappe.cls`).
- Nach LOHNTABELLE-Änderung werden Dropdowns markiert, aber die **Validation-Liste in F** scheint an **bestehende E-Werte** gebunden zu bleiben, bis E **neu gesetzt** wird (Re-Trigger der Zeilen-Logik).
- Performance-Optimierungen (kein F-Rebuild bei jedem Tab-Wechsel, kein E-Rebuild bei Open) können dieses Verhalten begünstigen.

### Geplante Verbesserung (später)

**Ziel:** Neue LOHNTABELLE-Stunden **automatisch** in F verfügbar machen — **ohne** E erneut wählen zu müssen — bei **Beibehaltung** der aktuellen Öffnungs-/Tab-Wechsel-Geschwindigkeit.

**Mögliche Ansätze (noch nicht entschieden):**

| Ansatz | Idee | Pro | Contra |
|--------|------|-----|--------|
| A — Zielgerichteter F-Refresh nach KV-Insert | Nach `AddCustomKVMonatsstunden` / `FormatKVPeriodArea` nur F-Validierungen für Zeilen mit passendem KV-Code (E) neu aufbauen | Präzise | Muss Schutz/Performance pro Zeile abwägen |
| B — Tab-Activate: Dirty-Refresh vollständig | Beim ersten Monats-Tab nach `MarkAllKVDropdownsDirty` alle F-Listen wirklich neu binden (nicht nur Helper) | Einfacher für User | Risiko langsamer Tab-Wechsel |
| C — E „weicher“ Re-Trigger | Intern E-Wert kurz leeren/setzen oder Validation invalidieren ohne User-Aktion | Kein manuelles Klicken | Hacky; Events/Overrides beachten |
| D — F bei Fokus/Klick | Beim Öffnen des F-Dropdowns immer aktuelle Stundenliste aus LOHNTABELLE (bereits teilweise lazy) | Schnell im Alltag | Erst beim F-Klick sichtbar |

**Empfohlene Reihenfolge bei Umsetzung:**

1. Reproduktion: Eigene Stunden einfügen → Januar → Zeile mit gleichem KV-Code wie neue Stunde → F-Liste prüfen.
2. Code-Pfad: `AddCustomKVMonatsstunden` → `MarkAllKVDropdownsDirty` → `RefreshKVDropdownsIfDirtyForSheet` / `RefreshKVStundenDropdownForSingleRow` nachvollziehen.
3. Fix mit Messung: Tab-Wechsel-Zeit vor/nach Änderung (Mac + Windows).
4. Regression: SMOKE, bestehende lazy F-Logik, kein 12× Voll-Rebuild bei Open.

### Akzeptanzkriterien (wenn umgesetzt)

- [ ] Nach neuer **Eigene Stunden**-Zeile in LOHNTABELLE erscheint die Stunde in **F** auf Monatsblättern **ohne** erneute E-Auswahl (für passenden KV-Code).
- [ ] Workbook-Open und erster Monats-Tab bleiben **spürbar schnell** (kein spürbarer Voll-Refresh aller Zeilen).
- [ ] Bestehende F-Overrides / Zukunftsplanung unverändert korrekt.
- [ ] Mac + Excel 2016+ kompatibel.

### Betroffene Dateien (Referenz)

- `vba/mod_AddNewKVPeriodOnTop.bas` — `AddCustomKVMonatsstunden`, `MarkAllKVDropdownsDirty`
- `vba/mod_KVStundenDropdown.bas` — F-Validation, Helper, Row-Refresh
- `vba/DieseArbeitsmappe.cls` — `RefreshKVDropdownsIfDirtyForSheet`, SheetActivate

---

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
