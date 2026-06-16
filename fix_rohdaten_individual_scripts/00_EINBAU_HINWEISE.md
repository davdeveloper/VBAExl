# Einbau: Rohdaten-Imports als einzelne Apps-Script-Dateien

Diese Dateien sind dafür gedacht, die jeweiligen vorhandenen Einzel-Skripte zu ersetzen.

Wichtig:

- Nicht zusätzlich einfügen, wenn die alte Funktion noch aktiv existiert.
- Je Datei den alten Inhalt ersetzen oder die alte gleichnamige Funktion vorher entfernen/auskommentieren.
- Die Button-Funktionsnamen bleiben gleich.

## Dateien und Button-Funktionen

| Datei | Enthält Button-Funktion |
| --- | --- |
| `IMPORT_DATA_Rohdaten_Food_Sortimentswünsche.gs` | `importFoodAssortmentWishes` |
| `IMPORT_DATA_Rohdaten_Anderungswunsche_Sort.gs` | `importChangeWishesSorting` |
| `IMPORT_DATA_Rohdaten Allg. Sortimentsw..gs` | `importRawGeneralAssortment` |
| `IMPORT_DATA_Sonstiges_Food_Themen.gs` | `importRawOtherFoodTopics` |
| `IMPORT_DATA_Rohdaten Food Qualitätsfälle.gs` | `importRawFoodQualityCases` |
| `IMPORT_DATA_OG_Sortimentswunche.gs` | `importOGAssortmentWishes` |
| `IMPORT_DATA_OG_Sonstiges.gs` | `importOGOther` |
| `IMPORT_DATA_ROHDATEN_OG_QUALITATSFALLE.gs` | `importRawOGQualityCases` |
| `IMPORT_DATA_Rohdaten_OG_Qualif_o_LFN.gs` | `importRawOGQualityCasesNoLFN` |

## Was geändert wurde

- Jede Datei ist eigenständig und enthält eigene Helper mit individuellem Prefix.
- Es gibt keine Abhängigkeit von einem zentralen Helper-Script.
- Beim Import wird `Einstellungen` automatisch konsistent gehalten:
  - Spalte E = Google-Sheet-Datei-ID
  - Spalte F = Google-Sheet-URL
- Die Rohdaten werden aus den generierten Sheets mit `getDisplayValues()` gelesen.
- Es wird nicht mehr in die generierten Quelldateien zurückgeschrieben.
- Umlaut-/Unicode-Unterschiede wie `Qualitätsfälle` vs. `Qualitätsfälle` werden robuster erkannt.

## Falls du doch lieber einen großen Block willst

Der zentrale Gesamtblock liegt weiterhin hier:

`outputs/fix_all_rohdaten_imports/ROHDATEN_IMPORTS_FIXED.gs`

Dann darfst du aber die Einzeldateien nicht zusätzlich mit denselben Button-Funktionen aktiv lassen.
