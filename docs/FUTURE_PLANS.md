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

## Weitere Einträge

Neue Backlog-Punkte unten anfügen mit ID `FP-00N`, Status, Ursache, geplantem Ansatz und Akzeptanzkriterien.
