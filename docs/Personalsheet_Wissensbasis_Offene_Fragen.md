# Offene und bewusst begrenzte Geschäftslogik

Stand: August 2026

Für die aktuelle Restaurant-Manager-Wissensbasis gibt es **keine blockierende offene Frage**.
Die folgenden Punkte sind Produkt- oder Entwicklungsentscheidungen. Die Endanwender-Hilfe
formuliert sie deshalb bewusst vorsichtig oder lässt technische Details weg.

## 1. Personal-ID-Korrektur und CopyData

### Sicher dokumentiert

- Eine falsche Personal-ID oder ein falscher Name wird über den sichtbaren Button
  **„Personal-ID korrigieren“** geändert.
- Die Korrektur gilt für alle Monate und gespeicherte Stundenänderungen.

### Nicht als Benutzerregel dokumentiert

- Ob zwei gleiche Namen mit verschiedenen IDs immer dieselbe Person sind.
- Ob CopyData eine manuelle ID-Korrektur automatisch erkennen soll.

Grund: Zwei verschiedene Menschen können gleich heißen. Der Code führt solche Einträge nicht
automatisch zusammen.

## 2. Neues Jahr

### Aktueller, bestätigter Prozess

- In Teams unter **Wichtige Unterlagen** liegt eine leere Datei für das nächste Jahr.
- Restaurant Manager laden diese Datei herunter.
- Das neue Jahr wird darin manuell vorbereitet.
- Die Januar-Planwerte des neuen Jahres werden für `UBERSICHT!E30` und `UBERSICHT!I30`
  der laufenden Datei verwendet.

### Noch nicht implementiert

- automatischer „Neues Jahr“-Button;
- automatisches Kopieren von Dezember nach Januar;
- automatische Dateibenennung oder Archivierung.

Die Wissensbasis behauptet deshalb keine automatische Jahresübernahme.

## 3. Extra-Kosten

Bestätigte Regel:

- Zeilen 18–25: Extra-Kosten, die in spätere Monate kopiert werden;
- Zeilen 26–28: Extra-Kosten nur für den aktuellen Monat;
- Beispiele: Jubiläumsgeld, Austrittskosten.

Offen für eine spätere Produktverbesserung:

- feste Beschriftungen oder Kategorien für diese Zeilen;
- Pflichtfeld für Beschreibung oder Kostenart.

Die Wissensbasis beschreibt die heutige freie Eingabe.

## 4. Alte KV-Periode löschen

Bestätigte Regel:

- Restaurant Manager dürfen den sichtbaren Button verwenden.
- Gelöscht wird eine alte, nicht mehr benötigte KV-Periode.
- Typischerweise betrifft dies ungefähr das Schema von vor zwei Jahren.

Bewusste Grenze:

- Bei Unsicherheit wird nicht gelöscht.
- Zuerst Payroll oder Admin fragen.

## 5. `AlleDatenLoeschen`, `DataClear`, manuelle Fluktuationsaktualisierung

Diese Funktionen sind technisch vorhanden, haben aber keinen normalen sichtbaren
Restaurant-Manager-Button.

Eigentümerentscheidung:

- Die Wissensbasis erklärt nur sichtbare Buttons.
- Alt+F8 und Entwicklermakros werden nicht dokumentiert.
- Destruktive Admin-Funktionen gehören nicht in die Anfänger-Hilfe.

## 6. Automatische Spalte G

Spalte G kann technisch als Formel oder als von VBA geschriebener Wert gespeichert sein.
Für Restaurant Manager ist beides dasselbe:

- gelbes, automatisches Feld;
- nicht manuell ändern;
- Quelle ist die LOHNTABELLE.

Die technische Unterscheidung wird in der Endanwender-Hilfe nicht erklärt.

## 7. Screenshots

Das aktuelle Produktions-Workbook enthält echte Personal-, Budget- und Fluktuationsdaten.
Es darf nicht direkt für öffentliche Screenshots verwendet werden.

Aktuelle Lösung:

- vereinfachte, exakt beschriftete HTML-Darstellungen;
- keine Personennamen, IDs oder echten Finanzwerte.

Spätere echte Screenshots benötigen eine vollständig anonymisierte Demo-Datei.

## 8. Manuelle Produktprüfung

Linux/Cloud kann Excel-Makros und VBA-UDFs nicht ausführen. Vor der endgültigen Verteilung
müssen deshalb in Windows + Microsoft Excel geprüft werden:

- alle sichtbaren Buttons;
- CopyData-Verhalten;
- Dropdowns;
- Lohnberechnung;
- Blattfarben und Schutz;
- Fehlermeldungs-Workflow;
- die exakten Beschriftungen nach dem letzten Formatlauf.

