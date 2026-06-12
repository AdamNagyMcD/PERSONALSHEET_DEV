# PERFORMANCE BASELINE — FP-010

Windows-Referenzmessung fuer FP-005–FP-009. Gleiche Schritte nach jedem Performance-Fix wiederholen.

## Umgebung (vor jeder Messung ausfuellen)

| Feld | Wert |
|------|------|
| Datum | 2026-06-12 22:37:46 (Nach FP-005) |
| Tester | Adam |
| PC / CPU | Windows 64-bit (NT 10.0) |
| RAM | — |
| Excel-Version | 16.0 |
| OS | Windows (64-bit) NT 10.00 |
| Workbook | `Personalsheet.xlsm` |
| Git-Commit / Tag | — |
| Notizen | Calculation: Automatisch; Log: `PID_PERFORMANCE_LOG` |

**Vorbereitung**

1. Excel komplett schliessen.
2. `ResetAndImportVBAFiles` + Compile + Speichern + Excel neu starten (nach VBA-Aenderung).
3. Keine anderen schweren Programme im Hintergrund.
4. Workbook lokal (nicht ueber langsames Netzwerk) oeffnen.

---

## Messmethode

| Symbol | Bedeutung |
|--------|-----------|
| **AUTO** | Makro `PID_RunPerformanceBaseline` (Alt+F8) — nicht-destruktiv |
| **MANU** | Stoppuhr / gefuehlte Wartezeit — siehe Schritte unten |
| **SEMI** | Makro misst Teil; Rest manuell (z. B. CopyData) |

Ergebnisse in die Tabellen unten eintragen. Ziel: **Sekunden (s)**, eine Dezimalstelle reicht.

---

## Schritte (Reihenfolge FP-010)

### 1 — Cold Open bis erster Monats-Tab nutzbar (MANU)

1. Excel starten (kalt).
2. `Personalsheet.xlsm` oeffnen — Stoppuhr ab **Doppelklick** bis Workbook sichtbar + berechnet.
3. Ersten Monats-Tab (z. B. **Januar**) anklicken — Stoppuhr stoppen, wenn E/F bedienbar und G nicht `#NAME?`/leer (bei Datenzeile).

| Lauf | Open bis WB sichtbar (s) | + Monats-Tab nutzbar (s) | Gesamt (s) |
|------|--------------------------|---------------------------|------------|
| Baseline | — | — | *offen (MANU)* |
| Nach FP-00x | | | |

---

### 2 — LOHNTABELLE → Eigene Stunden → erster Monats-Tab (SEMI)

**Manuell (realistischer Alltag):**

1. Tab **LOHNTABELLE** — ggf. neue Stunde / KV-Aenderung simulieren.
2. Tab **Februar** (erster Monats-Tab nach Aenderung) — Stoppuhr bis F-Dropdown oeffnet und Liste sichtbar.

**AUTO (Makro):**

- **2** — `MarkAllKVDropdownsDirty` → Refresh (Februar) — Worst-Case / Voll-Refresh.
- **2b** — `MarkKVDropdownDirtyForKVCode("BG1")` → Refresh (Februar) — FP-005 scoped dirty (typisch nach Eigene Stunden).

| Lauf | MANU: LOHNTABELLE → Monats-Tab F (s) | AUTO 2: KV-Refresh Voll (s) | AUTO 2b: KV scoped BG1 (s) |
|------|--------------------------------------|-----------------------------|----------------------------|
| Baseline | *offen (MANU)* | **0,51** | — |
| Nach FP-005 (2026-06-12) | *offen (MANU)* | **0,52** | **0,15** |

**FP-005:** scoped dirty (2b) **~3,4× schneller** als Voll-Refresh (0,15 s vs 0,51 s Baseline / 0,52 s Nachher).

---

### 3 — E/F-Aenderung → G stabil (AUTO + optional MANU)

**AUTO:** Makro recalculiert Monatslohn (G) fuer eine belegte Zeile auf Januar.

**MANuell:** E in Zeile 3 aendern → bis G-Wert stabil (Stoppuhr).

| Lauf | AUTO: G-Recalc 1 Zeile (s) | MANU: E-Aenderung → G (s) |
|------|----------------------------|---------------------------|
| Baseline | **0,02** | *offen (MANU)* |
| Nach FP-00x | | |

---

### 4 — CopyData Januar → Dezember (MANU / SEMI)

**Achtung:** Aendert Folgemonate — nur auf **Testkopie** oder bewusst vor Messung.

1. Tab **Januar** aktiv.
2. Alt+F8 → `CopyData` — Stoppuhr ab Start bis Erfolgs-Dialog.

| Lauf | Dauer (s) |
|------|-----------|
| Baseline | *offen (MANU, Testkopie)* |
| Nach FP-00x | |

---

### 5 — UEBERSICHT nach FINANZIELL-Aenderung (AUTO)

**AUTO:** `PID_SyncFinanzSummaryToUbersicht` (alle Monate).

**MANU optional:** Monatsblatt Q17:R29 aendern → UEBERSICHT pruefen — Stoppuhr bis Zahlen stimmen.

| Lauf | AUTO: FINANZ-Sync (s) | MANU (s) |
|------|----------------------|---------|
| Baseline | **0,11** | *offen (MANU)* |
| Nach FP-005 | **0,10** | |

---

### 6 — Save mit Fluktuation dirty (AUTO + optional MANU)

**AUTO:**

- **6** — `RefreshFluktuationDataIfDirty` (Save-Pfad, nur FLUKTUATION_DATEN).
- **6b** — `RefreshFluktuationIfDirty` (Tab-Pfad, Daten + Analyse).

**MANU:** I oder N auf Monatsblatt aendern → Speichern (Strg+S) — Stoppuhr bis Save fertig.

| Lauf | AUTO 6: Save Daten only (s) | AUTO 6b: Tab Daten+Analyse (s) | MANU: Save gesamt (s) |
|------|-----------------------------|--------------------------------|------------------------|
| Baseline | **1,27** (alt: voll) | — | *offen (MANU)* |
| Nach FP-009 | | | |

---

### 7 — FullSystemRefresh (AUTO, Admin-Referenz)

**AUTO:** gleiche Schritte wie `PID_FullSystemRefresh`, ohne Abschluss-Dialog.

Nicht Alltags-KPI — nur Vergleich nach grossen Aenderungen.

| Lauf | Dauer (s) |
|------|-----------|
| Baseline | **8,18** |
| Nach FP-005 | **7,98** |

**AUTO-Zusammenfassung (Baseline 2026-06-12):** Schritte 2–7 gesamt ~10,1 s (KV 0,51 + G 0,02 + FINANZ 0,11 + Fluktuation 1,24 + FullRefresh 8,18).

**AUTO-Zusammenfassung (Nach FP-005, 2026-06-12 22:37):** 2b 0,15 + 2 0,52 + G 0,02 + FINANZ 0,10 + Fluktuation 1,27 + FullRefresh 7,98 s. Alltags-Pfad (scoped): **0,15 s** statt 0,51 s.

---

## Makro

| Makro | Zweck |
|-------|--------|
| `PID_RunPerformanceBaseline` | AUTO-Schritte 2–7 (Teilmessung), Ergebnis als MsgBox + Blatt `PID_PERFORMANCE_LOG` |
| `RunPerformanceBaseline` | Alias |

Nach dem Lauf: Werte aus MsgBox oder Log-Blatt in diese Datei kopieren.

---

## Akzeptanz FP-010

- [x] Erste Baseline auf Windows dokumentiert (AUTO 2026-06-12).
- [x] `docs/PERFORMANCE_BASELINE.md` — AUTO-Spalten ausgefuellt.
- [x] Baseline Schritt 2 (KV-Refresh 1 Blatt): **0,51 s** — Referenz fuer FP-005.
- [x] Nach FP-005 Schritt 2b (scoped BG1): **0,15 s** — Ziel erreicht.
- [ ] MANU-Schritte 1, 2 (real), 4 (CopyData), 6 (Save) noch offen.

---

## Verknuepfung

- Backlog: `docs/FUTURE_PLANS.md` → FP-010, FP-005–FP-009
- Release: `docs/RELEASE.md` (optional Kurzverweis)
