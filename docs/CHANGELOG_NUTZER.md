# Personalsheet – Änderungsübersicht

**Für Restaurant Manager** · verständlich, ohne Technik-Jargon  
**Neueste Einträge zuerst** · Stand: Juni 2026

> Entwickler-Details: [`CHANGELOG.md`](CHANGELOG.md)

---

## Noch nicht überall ausgerollt (Juni 2026, Test)

Diese Punkte sind in der Entwicklung / auf dem Mac getestet — kommen mit dem nächsten Update in die Ettermen:

- **KV-Gruppe wechseln (Spalte E):** Die Stunden-Liste (Spalte F) passt sich zuverlässiger an — auch nach „Eigene Stunden“ auf der LOHNTABELLE.
- **Mac:** Gleiche Verbesserungen für Dropdown und LOHNTABELLE-Ansicht nach KV-Änderungen.

---

## 13. Juni 2026

### Behoben
- **Stunden-Dropdown (F):** Wenn du die **KV-Gruppe** änderst, zeigt die Stunden-Liste wieder die **richtigen** Werte (z. B. Wechsel BG1 ↔ BG1_5).
- **Eigene Stunden:** Nach dem Anlegen eigener Vertragsstunden auf der LOHNTABELLE erscheinen diese in den Monatsblättern zuverlässiger in der Liste.

---

## 12. Juni 2026

### Neu
- **Druck-Anleitung** für Restaurant Manager (HTML, A4) — Schritt-für-Schritt auf Deutsch, Stand Juni 2026.

### Verbessert
- **Schnelleres Arbeiten** beim Wechsel zwischen Monatsblättern und nach KV-Änderungen (weniger unnötige Neuberechnung).
- **FLUKTUATION:** Beim Speichern weniger Wartezeit; volle Auswertung erst, wenn du das Blatt **FLUKTUATION** öffnest.
- **Monatsblätter geschützt:** Lohn-Spalten und Formeln lassen sich nicht versehentlich überschreiben; du kannst weiter normal Personal ID, Name, KV-Gruppe, Stunden, Austritt usw. pflegen.
- **Sortieren** auf Monatsblättern ist deaktiviert (verhindert kaputtes Layout).
- **Zellen nach unten ziehen** bei KV-Gruppe/Stunden ist deaktiviert (Fill Handle).
- **Vormonat +/- Stunden:** In Monaten, die **kein** Durchrechnungs-Start sind (nicht Feb/Mai/Aug/Nov), kannst du die Korrektur wieder eintragen.

### Behoben
- **Durchrechnungs-Startmonate** (Februar, Mai, August, November): Anzeige im Monats-Panel konsistent.

---

## 9. Juni 2026

### Neu
- **Copyright-Hinweis** auf den Blättern (Adam Nagy / McOpCo) — klein in Zeile 2, stört die Arbeit nicht.

---

## 4.–7. Juni 2026

### Neu
- **LOHNTABELLE – Button „Stunde löschen“:** Falsch angelegte **eigene Stunden** kannst du wieder entfernen (nur selbst hinzugefügte Zeilen).
- **Copyright** auf allen sichtbaren Blättern.

### Verbessert
- **„Aktualisierung des restlichen Jahres“** startet **ohne extra Bestätigungsdialog** — ein Klick, fertig.
- **LOHNTABELLE:** Weniger störende Erfolgs-Meldungen nach „Neue Periode“, „Eigene Stunden“, „Stunde löschen“.
- **LOHNTABELLE:** Einheitlicheres Layout nach neuer KV-Periode; Navy-Schrift für alle Perioden.
- **Nach KV-Änderung:** Erstes Öffnen eines Monatsblatts spürbar **schneller**.
- **Datums-Spalten** (Eintritt / Austritt) etwas **breiter** — Datum besser lesbar.
- **Nach dem Öffnen** der Datei werden Lohn- und Stunden-Spalten wieder **automatisch berechnet** (nicht mehr leer stehen lassen).
- **UBERSICHT – Finanzübersicht** und **Durchrechnung** wieder vollständig und synchron mit den Monatsblättern.
- **Monatsblatt-Kopfzeile** und Zeilennummern (Spalte A) übersichtlicher formatiert.
- **FLUKTUATION:** Austritts-Zahlen und Ampel zuverlässiger.

### Behoben
- **UBERSICHT:** Planfelder **E30 / I30** (Durchrechnung) wieder editierbar, Rest geschützt.
- **Vormonat +/- (Q12):** In den meisten Monaten wieder editierbar; in **Feb/Mai/Aug/Nov** fest (Durchrechnungs-Start — dort gibt es keinen Vormonat).
- **CopyData / Panel:** Zusatz-Infos unter „Reine Laborcost“ (Crew Labor) werden in Folgemonate mitkopiert.
- **Zebra-Farben** auf Monatsblättern nach Aktualisierung bleiben erhalten.

### Entfernt
- **FLUKTUATION PDF-Export** (Button) — wurde kurz getestet, ist **nicht** Teil der endgültigen Version.

---

## 27.–28. Mai 2026

### Verbessert
- **Druck-Anleitung** vereinfacht, mit Umlauten (ä, ö, ü, ß).

### Behoben
- **Stunden-Liste (F)** nach Änderung auf der LOHNTABELLE: neue Stunden erscheinen **ohne** KV-Gruppe nochmal anklicken zu müssen.
- **Freitext unter „Reine Laborcost“:** Tippen dort löst **kein** Flackern mehr auf der ganzen rechten Seite aus.
- **Letztes Gehalt (Spalte L):** Wenn kein Wert — Zelle bleibt **leer** statt „0,00 €“.

---

## 24.–25. Mai 2026

### Neu
- **Druck-Anleitung** (erste Version) für Restaurant Manager.
- **Durchrechnung auf UEBERSICHT:** Jahresüberblick Stunden-Differenz, Überstunden-Risiko, Planwerte.

### Verbessert
- **LOHNTABELLE:** Buttons „Neue Periode“, „Eigene Stunden“ usw. klarer beschriftet und nebeneinander.
- **LOHNTABELLE:** Status/Prüfung-Spalten aktualisieren sich wieder bei Änderungen.
- **EINSTELLUNG, UEBERSICHT, FLUKTUATION:** Einheitlicheres, ruhigeres Layout (Navy/Gelb — wie ihr es von UEBERSICHT kennt).
- **FLUKTUATION:** Übersichtlicher für Manager — Ampel, „Sofort prüfen“, Empfehlungen.
- **Arbeitsjahr** steht zentral auf **EINSTELLUNG** (nicht mehr verstreut auf LOHNTABELLE).
- **Monatslohn** wird zuverlässiger aus der LOHNTABELLE geholt (auch Sonder-Codes wie BG3_15).
- **KV-Gruppe (E):** Dropdown-Liste funktioniert stabiler (kein #REF! mehr).
- **Geschwindigkeit:** Datei öffnet schneller; weniger Rechenarbeit im Hintergrund.

### Behoben
- **Letztes Gehalt (L)** und **Aktuelle Stunden (H):** Formeln nach Jahr-Wechsel / CopyData wieder korrekt.
- **Monatslohn (G):** Wird wieder angezeigt, wenn du KV-Gruppe oder Stunden änderst.
- **Austrittsdatum:** Datums-Prüfung passt zum Arbeitsjahr auf EINSTELLUNG.
- **CopyData-Button** auf allen Monatsblättern zuverlässig positioniert (auch Mac / Februar).

---

## Bekannt · geplant (noch nicht behoben)

| Thema | Kurz |
|--------|------|
| **Beim Öffnen kurz „Berechnet…“** | Besonders Excel 2016 — wird optimiert (FP-027). |
| **Stunden ändern, dann zurückdrehen** | Alte Stundenänderung kann „hängen bleiben“ — technisches Protokoll; Fix geplant (FP-028). **Workaround:** Aktualisierung immer vom **Monat der Änderung** aus starten. |
| **Urlaub in € (Spalte K) zeigt 0** | In leeren Zeilen steht 0 statt nichts — Fix geplant (FP-029), analog Spalte L. |

Fragen: **Adam Nagy** · adam.nagy@at.mcd.com

---

## Was du normalerweise **nicht** siehst

Interne Entwickler-Themen (VBA-Import, Smoke-Tests, Git, Performance-Messungen) stehen nur in [`CHANGELOG.md`](CHANGELOG.md) — für den Restaurant-Alltag nicht relevant.
