# PERSONALSHEET

Excel-basiertes Personal-, Lohn- und Planungswerkzeug für Restaurants.

> **Aktueller Stand:** `v0.9.1-test` – zweite Testversion, noch nicht für den uneingeschränkten Produktivbetrieb freigegeben.

## Überblick

PERSONALSHEET unterstützt Restaurantleitungen und autorisierte Führungskräfte bei der strukturierten Verwaltung personalbezogener Daten. Die Arbeitsmappe verbindet eine vertraute Excel-Oberfläche mit automatisierten VBA-Abläufen, Eingabeprüfungen, Formelkontrollen, administrativen Werkzeugen und einer ausführlichen deutschsprachigen Wissensbasis.

Das Projekt wurde für den Einsatz auf Windows-Arbeitsplätzen mit Microsoft Excel 2016 oder neuer entwickelt. Ein besonderer Schwerpunkt liegt auf stabiler Ausführung auf älteren Laptops und kleineren Bürocomputern sowie auf dem Schutz vor unbeabsichtigten Änderungen durch unerfahrene Benutzer.

## Hauptfunktionen

- Verwaltung und Pflege personalbezogener Stammdaten
- Monatsblätter von Januar bis Dezember
- automatische Berechnung und Pflege relevanter Formelspalten
- Verarbeitung von Monatslohn, aktuellen Stunden, Urlaubswerten und letztem Gehalt
- KV- und Lohnzuordnung
- Fluktuationsberechnung und Auswertungen
- kontrollierte Übernahme und Aktualisierung von Daten
- administrative Wartungs-, Reparatur- und Prüffunktionen
- Schutz wichtiger Formeln, Arbeitsblätter und Arbeitsmappenstrukturen
- Plausibilitätsprüfung bei Benutzereingaben
- deutschsprachige Offline-Wissensbasis mit Suche, Inhaltsverzeichnis und FAQ
- druckbare Schnellhilfe für Benutzer mit wenig Excel-Erfahrung

## Zielgruppe

Die Arbeitsmappe ist für unterschiedliche Benutzergruppen ausgelegt:

- **Normale Benutzer** bearbeiten ausschließlich die freigegebenen Eingabefelder.
- **Restaurantleitungen und Führungskräfte** verwenden die operativen Personal- und Planungsfunktionen.
- **Administratoren** verwalten geschützte Bereiche, führen Reparaturen durch und aktualisieren den eingebetteten VBA-Code.
- **Entwickler** bearbeiten die exportierten VBA-Quellen, führen Prüfungen aus und synchronisieren anschließend die Arbeitsmappe.

## Systemvoraussetzungen

- Windows
- Microsoft Excel 2016 oder neuer
- aktivierte VBA-Makros
- für bestimmte administrative Wartungsfunktionen: aktivierter vertrauenswürdiger Zugriff auf das VBA-Projektobjektmodell
- Python 3 nur für Entwicklungs- und Release-Prüfungen

### Kompatibilitätsregeln

Das Projekt verwendet bewusst keine Funktionen, die in Excel 2016 nicht verfügbar sind. Dazu gehören insbesondere:

- `XLOOKUP`
- `LET`
- `FILTER`
- `UNIQUE`
- dynamische Arrayformeln neuerer Excel-Versionen

macOS wird derzeit nicht unterstützt.

## Schnellstart für Benutzer

1. `Personalsheet.xlsm` herunterladen.
2. Die Datei auf einem lokalen, vertrauenswürdigen Speicherort ablegen.
3. Die Arbeitsmappe mit Microsoft Excel öffnen.
4. Falls angezeigt, die geschützte Ansicht verlassen.
5. Über **Inhalt aktivieren** die Makros freigeben.
6. Nur die vorgesehenen und freigegebenen Eingabefelder bearbeiten.
7. Die Hinweise und Meldungen der Arbeitsmappe beachten.
8. Nach der Arbeit speichern und Excel ordnungsgemäß schließen.

Für eine ausführliche Einführung steht die interaktive Wissensbasis unter [`docs/Personalsheet_Wissensbasis.html`](docs/Personalsheet_Wissensbasis.html) zur Verfügung.

## Wichtige Sicherheitshinweise

- Vor dem ersten Einsatz und vor größeren Änderungen immer eine Sicherungskopie erstellen.
- Formeln, geschützte Zellen, Arbeitsblätter und Steuerelemente nicht manuell verändern.
- Dateien nicht gleichzeitig in mehreren Excel-Instanzen öffnen.
- Während eines VBA-Imports oder einer Reparatur Excel nicht schließen.
- Administrative Funktionen nur verwenden, wenn ihre Auswirkungen bekannt sind.
- Passwörter, personenbezogene Daten und produktive Arbeitsmappen nicht öffentlich im Repository speichern.

Die eingebauten Schutzmechanismen verhindern typische Bedienfehler, stellen jedoch keinen vollständigen Schutz gegen gezielte technische Manipulation dar.

## Dokumentation

| Datei | Inhalt |
|---|---|
| [`docs/Personalsheet_Wissensbasis.html`](docs/Personalsheet_Wissensbasis.html) | Interaktive deutschsprachige Offline-Wissensbasis mit Suche und FAQ |
| [`docs/Personalsheet_Schnellhilfe.html`](docs/Personalsheet_Schnellhilfe.html) | Kurze, druckbare Hilfe für den täglichen Einsatz |
| [`docs/Kurzanleitung_Personalsheet_A4.html`](docs/Kurzanleitung_Personalsheet_A4.html) | Kompakte A4-Kurzanleitung für Benutzer |
| [`docs/Personalsheet_Wissensbasis_Wartung.md`](docs/Personalsheet_Wissensbasis_Wartung.md) | Hinweise zur Pflege und Aktualisierung der Dokumentation |
| [`docs/Personalsheet_Wissensbasis_Abdeckung.md`](docs/Personalsheet_Wissensbasis_Abdeckung.md) | Dokumentationsabdeckung der Funktionen und Arbeitsblätter |
| [`docs/Personalsheet_Wissensbasis_Offene_Fragen.md`](docs/Personalsheet_Wissensbasis_Offene_Fragen.md) | Noch offene fachliche Punkte |
| [`docs/CHANGELOG.md`](docs/CHANGELOG.md) | Technische und funktionale Änderungen |
| [`docs/AUDIT_2026-08.md`](docs/AUDIT_2026-08.md) | Performance- und Sicherheitsprüfung |
| [`docs/RELEASE.md`](docs/RELEASE.md) | Release-Ablauf und Checkliste |
| [`docs/FUTURE_PLANS.md`](docs/FUTURE_PLANS.md) | Geplante Verbesserungen und Backlog |
| [`TEST_CASES.md`](TEST_CASES.md) | Manuelle und technische Testszenarien |
| [`SPEC.md`](SPEC.md) | Fachliche Regeln und Projektspezifikation |

## Projektstruktur

```text
Personalsheet.xlsm                 Ausführbare Excel-Arbeitsmappe
vba/                               Exportierte VBA-Quellen
tools/                             Statische Prüf- und Analysewerkzeuge
docs/                              Benutzer-, Audit- und Wartungsdokumentation
SPEC.md                            Fachliche Spezifikation
TEST_CASES.md                      Testfälle und Prüfschritte
README.md                          Projektübersicht
```

Lokale Sicherungen, Arbeitskopien und temporäre Auditdateien gehören nicht in das Repository und werden über `.gitignore` ausgeschlossen.

## Entwicklungsprinzip

Die Dateien im Ordner `vba/` sind die maßgebliche, versionierte Quelle für die VBA-Entwicklung. `Personalsheet.xlsm` ist das ausführbare Ergebnis und muss vor einem Commit oder Release mit diesen Quellen synchronisiert werden.

Änderungen sollen:

- die bestehende fachliche Logik erhalten;
- vollständig mit Excel 2016 kompatibel bleiben;
- auch auf schwächeren Bürocomputern zuverlässig funktionieren;
- keine unnötigen Neuberechnungen oder zellweisen VBA-Schleifen einführen;
- bei Fehlern Excel-Einstellungen wie `ScreenUpdating`, `EnableEvents`, `Calculation` und `DisplayAlerts` zuverlässig wiederherstellen;
- mit deutschen Kommentaren und verständlichen Bezeichnungen dokumentiert werden.

## Empfohlener Entwicklungsablauf

1. Vor Arbeitsbeginn den aktuellen Stand abrufen:

   ```powershell
   git pull --ff-only
   ```

2. VBA ausschließlich in den exportierten Dateien unter `vba/` bearbeiten.

3. Die Arbeitsmappe öffnen und das Makro `ResetAndImportVBAFiles` ausführen.

4. Falls erforderlich, `SyncDieseArbeitsmappeFromExport` ausführen.

5. Arbeitsmappe speichern und Excel vollständig schließen.

6. VBA-Quellen und eingebetteten Code vergleichen:

   ```powershell
   python tools/check_vba_sync.py
   ```

7. Statische VBA-Prüfung ausführen:

   ```powershell
   python tools/vba_lint.py
   ```

8. Formelspalten prüfen:

   ```powershell
   python tools/check_formula_columns.py --verbose
   ```

9. Relevante manuelle Tests aus `TEST_CASES.md` durchführen.

10. Nur geprüfte und zusammengehörige Änderungen committen und pushen.

## Automatische Prüfwerkzeuge

### VBA-Lint

```powershell
python tools/vba_lint.py
```

Die Prüfung erkennt unter anderem:

- nicht deklarierte Prozeduren und Funktionen;
- unzulässige Aufrufe privater Prozeduren;
- doppelte öffentliche Namen;
- fehlende Sprungmarken;
- unausgeglichene VBA-Blöcke;
- fehlendes `Option Explicit`;
- Modulnamen, die nicht zum Dateinamen passen;
- nicht mit Excel 2016 kompatible Formeln.

Erwartetes Ergebnis vor einem Release:

```text
Errors: 0   Warnings: 0
```

### Synchronisationsprüfung

```powershell
python tools/check_vba_sync.py
```

Sie vergleicht den eingebetteten VBA-Code der Arbeitsmappe mit den versionierten Dateien in `vba/`.

Erwartetes Ergebnis:

```text
Abweichungen: 0
```

### Prüfung der Formelspalten

```powershell
python tools/check_formula_columns.py --verbose
```

Kontrolliert die Formelversorgung der relevanten Spalten G, H, K und L auf allen Monatsblättern.

## Release-Checkliste

Vor jeder Test- oder Produktivversion müssen mindestens folgende Punkte erfüllt sein:

- Arbeitsmappe lässt sich ohne Reparaturmeldung öffnen.
- Makros sind importiert und kompilierbar.
- `vba_lint.py` meldet keine Fehler und keine Warnungen.
- `check_vba_sync.py` meldet keine Abweichungen.
- `check_formula_columns.py` meldet keine fehlenden Formeln.
- alle kritischen Funktionen wurden gemäß `TEST_CASES.md` geprüft.
- `README.md`, Wissensbasis und Changelog entsprechen dem aktuellen Stand.
- keine personenbezogenen Daten, Passwörter oder lokalen Backups sind enthalten.
- die für Tester bestimmte `Personalsheet.xlsm` ist als Release-Datei beigefügt.
- Testversionen werden auf GitHub als **Pre-release** gekennzeichnet.

## Aktueller Teststand

Für `v0.9.1-test` wurden zuletzt folgende technische Ergebnisse dokumentiert:

- 35 VBA-Module
- 911 geprüfte Prozeduren
- 24.468 geprüfte VBA-Zeilen
- 0 VBA-Fehler
- 0 VBA-Warnungen
- 0 Abweichungen zwischen Arbeitsmappe und exportierten VBA-Quellen
- 0 fehlende Formeln in den geprüften Formelspalten
- alle Monatsblätter von Januar bis Dezember geprüft

Die aktuelle Testversion ist unter [GitHub Releases](https://github.com/AdamNagyMcD/PERSONALSHEET_DEV/releases) verfügbar.

## Fehler melden

Bei einem Fehler bitte möglichst folgende Informationen dokumentieren:

- verwendete PERSONALSHEET-Version;
- Excel- und Windows-Version;
- betroffenes Arbeitsblatt und betroffene Funktion;
- genaue Schritte bis zum Fehler;
- Wortlaut oder Screenshot der Fehlermeldung;
- erwartetes und tatsächliches Ergebnis;
- Information, ob der Fehler erneut reproduzierbar ist.

Keine echten Mitarbeiterdaten oder Passwörter in öffentlich zugängliche Fehlermeldungen aufnehmen.

## Status

Das Projekt befindet sich weiterhin in der Testphase. `v0.9.1-test` ist für kontrollierte Tests vorgesehen und noch keine uneingeschränkt freigegebene Produktivversion.

## Urheberrecht

Copyright © Adam Nagy / McOpCo. Alle Rechte vorbehalten.

Sofern nicht ausdrücklich schriftlich erlaubt, sind unbefugtes Kopieren, Verändern, Veröffentlichen oder Weitergeben des Projekts und seiner Bestandteile nicht gestattet.
