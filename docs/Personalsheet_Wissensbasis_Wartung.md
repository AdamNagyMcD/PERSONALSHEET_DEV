# Personalsheet-Wissensbasis warten

Technische Kurzanleitung für die spätere Pflege der Endanwender-Dokumentation.

## Dateien

| Datei | Zweck |
|-------|-------|
| `docs/Personalsheet_Wissensbasis.html` | Vollständige, interaktive Offline-Hilfe |
| `docs/Personalsheet_Schnellhilfe.html` | Druckbare Ein-Seiten-Hilfe |
| `docs/Personalsheet_Wissensbasis_Wartung.md` | Diese technische Pflegeanleitung |
| `docs/Personalsheet_Wissensbasis_Abdeckung.md` | Dokumentations- und Prüfmatrix |
| `docs/Personalsheet_Wissensbasis_Offene_Fragen.md` | Noch offene oder bewusst nicht erklärte Geschäftslogik |

Es gibt keine externen Bilder, Fonts, CDN- oder JavaScript-Abhängigkeiten. CSS, JavaScript und
vereinfachte visuelle Darstellungen sind direkt in der HTML-Datei enthalten.

## Inhalt ändern

1. Die passende HTML-Datei in einem Texteditor öffnen.
2. Sichtbaren Text direkt im jeweiligen `<section>` oder `<details>` ändern.
3. Exakte Excel-Namen unverändert übernehmen:
   - `LOHNTABELLE`
   - `EINSTELLUNG`
   - `Januar` … `Dezember`
   - `Marz` (ohne Umlaut)
   - `UBERSICHT` (ohne Umlaut)
   - `FLUKTUATION`
4. Nur sichtbare Buttons für Restaurant Manager beschreiben.
5. Keine Makronamen, Admin-Passwörter oder Schutz-Umgehungen ergänzen.
6. Nach jeder Änderung die Prüfungen unten durchführen.

## Suche pflegen

Jeder durchsuchbare Inhaltsblock besitzt die Klasse:

```html
class="search-item"
```

Zusätzliche Synonyme stehen in:

```html
data-keywords="austritt kündigung verlassen"
```

Regeln:

- einfache Wörter verwenden;
- deutsche und häufige umgangssprachliche Wörter ergänzen;
- Blatt- und Buttonnamen exakt aufnehmen;
- keine personenbezogenen Daten in Keywords eintragen.

Die Suche:

- ignoriert Groß-/Kleinschreibung;
- behandelt Umlaute benutzerfreundlich;
- behandelt `ß` wie `ss`;
- hebt Treffer im Text hervor;
- öffnet passende FAQ- und Anleitungsblöcke automatisch.

## Neuen Abschnitt ergänzen

```html
<section class="section" id="eindeutige-id">
  <h2>Neue Überschrift</h2>
  <p class="lead">Kurze Erklärung.</p>

  <details class="card search-item" data-keywords="wichtige synonyme">
    <summary>Frage oder Aufgabe</summary>
    <div class="details-body">
      <ol class="steps">
        <li>Ein Schritt.</li>
        <li>Ein nächster Schritt.</li>
      </ol>
    </div>
  </details>
</section>
```

Das Inhaltsverzeichnis wird beim Öffnen automatisch aus `h2` und `h3` erzeugt.

## Farben und Hinweise

Vorhandene Klassen verwenden:

| Klasse | Bedeutung |
|--------|-----------|
| `.notice.info` | allgemeine Information |
| `.notice.good` | korrekt / erfolgreich |
| `.notice.warn` | wichtige Vorsicht |
| `.notice.danger` | nicht tun / stoppen |
| `.card` | normaler Inhaltsblock |
| `.steps` | nummerierte Einzelschritte |

Nicht nur Farbe verwenden. Immer zusätzlich Text oder ein Symbol angeben.

## Visuelle Darstellungen

Die Wissensbasis nutzt vereinfachte, exakt beschriftete HTML-Darstellungen:

- Tab-Reihenfolge;
- Spalten B–N;
- Monatsblatt-Buttons;
- Farblegende.

Die Hauptwissenbasis ist bewusst eine einzige, selbstständige HTML-Datei. Wenn später echte
Screenshots ergänzt werden, müssen sie als kleine, bereinigte lokale Assets zusammen mit der
HTML-Datei ausgeliefert werden. Ab diesem Moment ist das Ergebnis ein vollständiger
Offline-Dokumentationsordner, nicht mehr nur eine Einzeldatei.

Für Screenshots gelten diese Regeln:

1. Nie das Produktions-Workbook direkt fotografieren.
2. Eine Kopie mit vollständig gelöschten Personaldaten benutzen.
3. Nur fiktive IDs, Namen, Datums- und Finanzwerte eintragen.
4. Sichtbar „Beispiel“ ergänzen.
5. Dateiname, Windows-Benutzername und Pfad ausblenden.
6. Screenshot in `docs/Personalsheet_Wissensbasis_assets/` speichern.
7. Den vollständigen Ordner immer zusammen ausliefern.
8. Keine Base64-Bilder mit mehreren Megabyte in die Hauptdatei einbetten.

## Geschäftslogik aktualisieren

Vor einer fachlichen Änderung immer abgleichen:

1. tatsächliche Arbeitsmappe `Personalsheet.xlsm`;
2. exportierte VBA-Dateien in `vba/`;
3. `SPEC.md`;
4. `TEST_CASES.md`;
5. `README.md`;
6. `docs/CHANGELOG.md`;
7. Besitzerentscheidung in `Personalsheet_Wissensbasis_Offene_Fragen.md`.

Wenn Code und Dokumentation widersprechen: nicht raten. Zuerst den Besitzer fragen.

## Qualitätsprüfung nach jeder Änderung

### Automatisch

```bash
python3 tools/vba_lint.py
python3 tools/check_vba_sync.py
```

Für reine Dokumentationsänderungen ist `check_vba_sync.py` nur ein Zustandsbericht.

HTML-Prüfung:

- alle lokalen Links existieren;
- alle `id`-Werte sind eindeutig;
- jedes `<details>` hat ein `<summary>`;
- keine `http://`, `https://`, CDN- oder externe Script-/Style-Abhängigkeit;
- keine Admin-Passwörter oder Produktionsnamen.

### Manuell in Edge und Chrome

1. HTML per Doppelklick ohne Internet öffnen.
2. Suche mit diesen Begriffen prüfen:
   - `Austritt`
   - `kuendigung`
   - `März`
   - `Marz`
   - `geschuetzt`
   - `Jubiläumsgeld`
   - `keine rückmeldung`
3. „Keine Treffer“ prüfen.
4. Inhaltsverzeichnis und alle internen Links prüfen.
5. FAQ öffnen und schließen.
6. Fenster schmal ziehen; Inhalt muss ohne horizontales Seitenscrollen lesbar bleiben.
7. Druckvorschau A4:
   - Hauptwissenbasis;
   - `Personalsheet_Schnellhilfe.html`.
8. Schnellhilfe muss auf eine Seite passen oder höchstens kontrolliert auf zwei Seiten umbrechen.

## Version und Freigabe

Vor einer Restaurant-Verteilung:

1. sichtbaren Stand im Footer aktualisieren;
2. Abdeckungsmatrix aktualisieren;
3. offene Fragen prüfen;
4. die aktuelle, getestete `.xlsm`-Version verwenden;
5. keine Entwicklerdateien oder `_ADMIN`-Informationen an Restaurant Manager verteilen;
6. Wissenbasis und Schnellhilfe zusammen mit der Arbeitsmappe bereitstellen.

