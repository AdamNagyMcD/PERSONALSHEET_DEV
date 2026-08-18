# Abdeckung der Personalsheet-Wissensbasis

Stand: August 2026

**Status:** Inhalt und technische Umsetzung sind vollständig. Chrome, Suche, Navigation,
Mobilansicht und Druck wurden geprüft. Die endgültige Freigabe für Restaurants bleibt offen,
bis die unten noch nicht abgehakten Prüfungen in Edge, in der echten Windows-Excel-Arbeitsmappe
und mit einem neuen Restaurant Manager durchgeführt wurden.

## Zielgruppe und Grenze

- Zielgruppe: Restaurant Manager.
- Diese Personen erhalten und verwenden die Arbeitsmappe selbst.
- Dokumentiert werden ausschließlich sichtbare Arbeitsblätter, Eingaben und Buttons.
- Nicht dokumentiert werden Alt+F8, Makronamen, Entwicklerwerkzeuge, Schutzpasswörter,
  versteckte Blätter oder technische Reparaturwege.

## Arbeitsblätter

| Blatt | In Wissensbasis | Zweck | Eingaben erklärt | Automatik erklärt | Häufige Fehler |
|-------|-----------------|-------|-------------------|-------------------|----------------|
| `LOHNTABELLE` | ✅ | ✅ | ✅ | ✅ | ✅ |
| `EINSTELLUNG` | ✅ | ✅ | ✅ | ✅ | ✅ |
| `Januar` … `Dezember` | ✅ | ✅ | ✅, Spalten B–N | ✅, G/H/K/L + Panel | ✅ |
| `UBERSICHT` | ✅ | ✅ | ✅, E30/I30 | ✅ | ✅ |
| `FLUKTUATION` | ✅ | ✅ | keine Eingabe | ✅ | ✅ |
| `_ADMIN` | absichtlich nicht | — | — | — | Verweis auf Admin-Hilfe |
| technische VeryHidden-Blätter | absichtlich nicht | — | — | — | — |

## Sichtbare Buttons

### Monatsblätter

| Button | Zweck | Voraussetzung | Schritte | Ergebnis | Fehlerweg |
|--------|-------|---------------|----------|----------|-----------|
| Aktualisierung des restlichen Jahres | ✅ | ✅ | ✅ | ✅ | ✅ |
| Mitarbeiter entfernen | ✅ | ✅ | ✅ | ✅ | ✅ |
| Personal-ID korrigieren | ✅ | ✅ | ✅ | ✅ | ✅ |
| Fehler melden | ✅ | ✅ | ✅ | ✅ | ✅ |

### LOHNTABELLE

| Button | Dokumentiert |
|--------|--------------|
| 1) Neue Periode | ✅ |
| 2) Eigene Stunden | ✅ |
| 3) Alte Periode löschen | ✅, mit Grenze „normalerweise ungefähr vor zwei Jahren“ |
| 4) Stunde löschen | ✅ |

## Wichtige Prozesse

| Prozess | Dokumentiert | Eigentümerentscheidung berücksichtigt |
|---------|--------------|----------------------------------------|
| Datei öffnen / Makros aktivieren | ✅ | — |
| Neues Jahr aus Teams-Datei starten | ✅ | ✅ |
| LOHNTABELLE pflegen | ✅ | ✅ |
| EINSTELLUNG pflegen | ✅ | ✅ |
| Neuer Mitarbeiter | ✅ | ✅, BG1 Basis + Payroll-Ausnahme |
| Austritt | ✅ | ✅, Datum/Grund Pflicht, Urlaub darf leer sein |
| Stunden ändern | ✅ | ✅ |
| Änderung in Folgemonate übernehmen | ✅ | ✅ |
| Vormonat +/- Stunden | ✅ | ✅ |
| Extra-Kosten | ✅ | ✅, 18–25 kopiert / 26–28 nur Monat |
| Personal-ID/Name korrigieren | ✅ | — |
| Falsch angelegte Person entfernen | ✅ | ✅ |
| Monatsabschluss | ✅ | — |
| Nächstes Jahr E30/I30 | ✅ | ✅ |
| Fehler melden / Support | ✅ | ✅ |

## Excel-Grundlagen

| Thema | Dokumentiert |
|-------|--------------|
| Arbeitsblatt / Tab | ✅ |
| Zelle | ✅ |
| Eingeben / Korrigieren | ✅ |
| Enter / Tab / Pfeiltasten | ✅ |
| Dropdown | ✅ |
| Datum | ✅ |
| Speichern / Schließen | ✅ |
| Makro als sichtbare Button-Automatik | ✅ |
| Schutz / gesperrte Zelle | ✅ |

## Sicherheit und Fehlervermeidung

| Thema | Dokumentiert |
|-------|--------------|
| Weiß = Eingabe, Gelb = automatisch | ✅ |
| Nicht sortieren | ✅ |
| Keine Zeilen/Spalten löschen | ✅ |
| Kein Fill Handle / Herunterziehen | ✅ |
| Kein Strg+X | ✅ |
| Einfügen nur als Wert | ✅, eigene einfache Erklärung zu Strg+V |
| Doppelte Personal-ID | ✅ |
| ID/Name-Konflikt | ✅ |
| Austritt vor Eintritt | ✅ |
| Payroll-Abweichung | ✅ |
| Datenschutz bei Fehlerbericht | ✅ |
| Keine Admin-Passwörter | ✅, Qualitätsprüfung |

## Fehlerbehebung und FAQ

Abgedeckt:

- Button/Makro reagiert nicht;
- Geschützte Ansicht;
- Zelle nicht editierbar;
- falsche Eingabe;
- erwartetes Ergebnis fehlt;
- langsames Excel / Keine Rückmeldung;
- falscher Klick;
- Speichern nicht möglich;
- Fehlermeldung;
- Payroll-Abweichung;
- 12 FAQ-Fragen;
- klare Admin-Eskalation.

## Interaktive und technische Anforderungen

| Anforderung | Umsetzung |
|-------------|-----------|
| Offline | ✅ Keine externe Abhängigkeit |
| Einfache lokale HTML-Datei | ✅ |
| Suche | ✅ |
| Groß-/Kleinschreibung ignorieren | ✅ |
| Umlaute/`ß` benutzerfreundlich | ✅ |
| Synonyme | ✅ `data-keywords` |
| Treffer hervorheben | ✅ |
| Kein Treffer | ✅ Hilfetext |
| Dynamisches Inhaltsverzeichnis | ✅ |
| Klickbare Kapitel | ✅ |
| Auf-/zuklappbare Bereiche | ✅ `<details>` |
| Zurück nach oben | ✅ |
| Hinweise nach Typ | ✅ Info/Gut/Warnung/Stopp |
| Responsive | ✅ Desktop/Tablet/Mobil |
| Druckfreundlich | ✅ |
| Schnellhilfe separat | ✅ |
| CSS/JavaScript lokal | ✅ inline |
| Keine echten Personendaten | ✅ Nur vereinfachte Darstellungen |

## Prüfstatus vor endgültiger Restaurant-Freigabe

- [ ] Edge offline
- [x] Chrome offline per lokaler `file://`-URL
- [x] Suche mit Umlaut, ohne Umlaut und `ue`-Schreibweise
- [x] Inhaltsverzeichnis Desktop
- [x] FAQ öffnen/schließen
- [x] Kein-Treffer-Hilfe
- [x] Zurück-nach-oben
- [x] Druckvorschau Hauptwissenbasis (A4, kein offensichtliches Abschneiden)
- [x] Druckvorschau Schnellhilfe: A4, eine Seite, kein sichtbares Abschneiden
- [x] Mobilansicht bei 390 px: lesbar, kein Seitenüberlauf, mobiles Inhaltsverzeichnis funktioniert
- [x] aktuelle sichtbare Excel-Buttons und Beschriftungen statisch mit Workbook/VBA abgeglichen
- [ ] Windows + Excel: Bearbeitung/Makros aktivieren und alle 4 Monatsblatt-Buttons ausführen
- [ ] Windows + Excel: alle 4 LOHNTABELLE-Buttons einschließlich Abbruchpfade ausführen
- [ ] Windows + Excel: Dropdowns, Schutzfarben und Ergebnisfelder mit der Anleitung vergleichen
- [ ] Windows + Excel: „Fehler melden“ bis zum gespeicherten Bericht testen
- [ ] mindestens ein vollständig neuer Restaurant Manager liest „Neuer Mitarbeiter“ und
      „Austritt“ ohne mündliche Hilfe
- [x] keine Produktions-Screenshots oder echten Personaldaten in den HTML-Dateien

