# Analyse und Fixes: MIS, Rohdaten Food Qualitätsfälle, Import löschen

## MIS Import

Gefundene Ursache:

- `convertExcelFilesToSheets()` hat aus der temporär konvertierten Google-Datei `getValues()` gelesen. Dadurch wird nicht der sichtbare Excel-/Sheets-Wert übernommen, sondern der interne Rohwert. Wenn Excel z. B. `190.164` anzeigt, der Rohwert aber `190163.788` ist, wurde daraus beim Schreiben `190.163.788`.
- `runMisImportSmarter()` hat zusätzlich in der XLOOKUP-Formel Werte aktiv in Zahlen umgewandelt: `A4*1` und `IMPORTRANGE(...)*1`. Das ist bei Artikelnummern, WG-Codes und Punktwerten gefährlich.

Fix:

- `IMPORT_FILES_AND_CONVERT.gs` liest jetzt primär `getDisplayValues()` und schreibt alles als Plain Text.
- `IMPORT_MIS_DATEN.gs` importiert A:D ebenfalls über `getDisplayValues()`, setzt A:G auf Plain Text und vergleicht in der XLOOKUP-Formel mit `TO_TEXT(...)` statt Zahlenumwandlung.

Wichtig:

- Bereits falsch generierte Google-Sheets-Dateien müssen nach dem Einbauen des Fixes neu aus den Excel-Dateien konvertiert/überschrieben werden. Der Fix verhindert die falsche Umwandlung künftig, kann aber eine bereits falsch gespeicherte Quelldatei nicht zuverlässig zurückraten.

## Rohdaten Food Qualitätsfälle

Vergleich alter Stand / aktueller Stand:

- Im alten Code gab es einen V4-Block mit dynamischer Suche der `WG`-Spalte in den Kopfzeilen 1-3 und spezieller WG-Bereinigung.
- Der aktive aktuelle Block hatte diese V4-Logik nicht mehr aktiv und las starr `A:H`.
- Außerdem wurde im älteren aktiven Code in die Quelldatei zurückgeschrieben (`sourceRange.setValues(...)`). Das kann Quellberichte verändern und ist bei wiederholten Imports riskant.

Fix:

- `IMPORT_DATA_Rohdaten Food Qualitätsfälle.gs` aktiviert die V4-WG-Logik wieder.
- Der Import liest sichtbare Werte, bereinigt die WG-Spalte gezielt und schreibt nur ins Zielblatt `Rohdaten Food Qualitätsfälle`, nicht mehr zurück in die Quelldatei.

## Import löschen

Neue Button-Funktion:

```text
importLoeschenAlles
```

Was sie macht:

- Leert die Import-/Generierungsbereiche in allen relevanten Report-Blättern.
- Erhält Kopfzeilen, Template-Zeilen, Formeln, Formatierungen, Rahmen und Deckblätter.
- Löscht keine Quelldateien und keine Excel-/Google-Drive-Dateien.

Einschätzung:

- Für einen normalen sauberen Import ist sie nicht zwingend nötig, weil viele Imports ihre Zielbereiche selbst leeren.
- Sinnvoll ist sie bei abgebrochenen Läufen, Testläufen, halbfertigen Ketten oder wenn die neue KW weniger Zeilen hat und Reste sichtbar bleiben.

## Einbau

Diese Dateien als einzelne Apps-Script-Dateien übernehmen/ersetzen:

- `IMPORT_FILES_AND_CONVERT.gs`
- `IMPORT_MIS_DATEN.gs`
- `IMPORT_DATA_Rohdaten Food Qualitätsfälle.gs`
- `RESET_IMPORTS.gs`

Button-Zuordnung bleibt bzw. wird:

- Excel-Dateien konvertieren: `convertExcelFilesToSheets`
- MIS Import: `runMisImportSmarter`
- Rohdaten Food Qualitätsfälle: `importRawFoodQualityCases`
- Alles löschen / Neustart: `importLoeschenAlles`
