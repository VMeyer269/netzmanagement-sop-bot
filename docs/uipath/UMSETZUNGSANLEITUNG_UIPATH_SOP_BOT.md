# Umsetzungsanleitung: Konkreter UiPath-Workflow fuer den SOP-Monatsvergleich

Diese Anleitung ersetzt die allgemeine Beschreibung durch einen konkreten Bauplan fuer den UiPath-Workflow. Ziel ist ein lauffaehiger Bot, der:

1. die Datei des Vormonats einliest
2. die Datei des aktuellen Monats einliest
3. den `Invoke Code`-Block ausfuehrt
4. das Ergebnis in automatische und manuelle Faelle trennt
5. die Ausgabedateien schreibt

Referenz fuer den Code-Block:
- [InvokeCode_ExcelAbgleich.cs](/Users/personliches/Documents/process-automation-engine/docs/uipath/InvokeCode_ExcelAbgleich.cs)

## Zielstruktur im Designer

Der empfohlene Aufbau in UiPath ist:

1. `Sequence` mit Name `SOP Monatsvergleich`
2. darin ein `Try Catch`
3. im `Try`:
   - mehrere `Assign`
   - `Path Exists` fuer beide Quelldateien
   - `If` zur Dateipruefung
   - `Use Excel File` fuer Vormonat
   - `Read Range` nach `dtAlt`
   - `Use Excel File` fuer aktuellen Monat
   - `Read Range` nach `dtNeu`
   - `Invoke Code`
   - `If` fuer Ergebnispruefung
   - `Filter Data Table` nach `dtManual`
   - `Filter Data Table` nach `dtAuto`
   - `Use Excel File` fuer Ergebnisdatei
   - mehrere `Write Range`
   - mehrere `Log Message`
4. im `Catch`:
   - `Log Message`
   - optional `Throw`

## Variablen

Lege diese Variablen auf Workflow-Ebene an:

- `strFileAlt` vom Typ `String`
- `strFileNeu` vom Typ `String`
- `strOutputFile` vom Typ `String`
- `blnAltExists` vom Typ `Boolean`
- `blnNeuExists` vom Typ `Boolean`
- `dtAlt` vom Typ `System.Data.DataTable`
- `dtNeu` vom Typ `System.Data.DataTable`
- `dtChanges` vom Typ `System.Data.DataTable`
- `dtManual` vom Typ `System.Data.DataTable`
- `dtAuto` vom Typ `System.Data.DataTable`
- `intChangesCount` vom Typ `Int32`
- `intManualCount` vom Typ `Int32`
- `intAutoCount` vom Typ `Int32`
- `strErrorMessage` vom Typ `String`

## Schritt 1: Startwerte setzen

Füge am Anfang der `Try`-Sektion drei `Assign`-Activities ein.

### Assign 1

- Name: `Setze Pfad Vormonat`
- To: `strFileAlt`
- Value:

```vb
"C:\Users\swhrpant\Desktop\BKS_SOP_01_2026.xlsx"
```

Wenn du weiter mit CSV arbeitest:

```vb
"C:\Users\swhrpant\Desktop\BKS_SOP_01_2026.csv"
```

### Assign 2

- Name: `Setze Pfad aktueller Monat`
- To: `strFileNeu`
- Value:

```vb
"C:\Users\swhrpant\Desktop\BK_SOP_02_2026.xlsx"
```

oder bei CSV:

```vb
"C:\Users\swhrpant\Desktop\BK_SOP_02_2026.csv"
```

### Assign 3

- Name: `Setze Ausgabedatei`
- To: `strOutputFile`
- Value:

```vb
"C:\Users\swhrpant\Desktop\SOP_Changes_2026_02.xlsx"
```

## Schritt 2: Quelldateien pruefen

Füge zwei `Path Exists`-Activities ein.

### Path Exists 1

- Input Path: `strFileAlt`
- Exists: `blnAltExists`

### Path Exists 2

- Input Path: `strFileNeu`
- Exists: `blnNeuExists`

Danach ein `If`.

### If: Dateien vorhanden

- Condition:

```vb
Not blnAltExists OrElse Not blnNeuExists
```

### Then

Füge eine `Throw`-Activity ein.

- Exception:

```vb
New BusinessRuleException("Mindestens eine Eingabedatei wurde nicht gefunden. Vormonat: " + strFileAlt + " | Aktueller Monat: " + strFileNeu)
```

### Else

- leer lassen

## Schritt 3: Vormonat einlesen

Füge eine `Use Excel File`-Activity ein.

### Use Excel File: Vormonat

- Excel-Datei: `strFileAlt`
- Referenzieren als: `ExcelAlt`
- `Aenderungen speichern`: deaktiviert
- `Erstellen, falls nicht vorhanden`: deaktiviert

Innerhalb davon eine `Read Range`-Activity.

### Read Range: Vormonat

- Bereich:

```vb
ExcelAlt.Sheet("Sheet1").UsedRange
```

Wenn das Blatt anders heißt, dort den echten Blattnamen eintragen.

- Hat Header: aktiviert
- Nur sichtbare Zeilen: deaktiviert
- Speichern unter: `dtAlt`

## Schritt 4: Aktuellen Monat einlesen

Füge direkt darunter eine zweite `Use Excel File`-Activity ein.

### Use Excel File: aktueller Monat

- Excel-Datei: `strFileNeu`
- Referenzieren als: `ExcelNeu`
- `Aenderungen speichern`: deaktiviert
- `Erstellen, falls nicht vorhanden`: deaktiviert

Innerhalb davon wieder eine `Read Range`-Activity.

### Read Range: aktueller Monat

- Bereich:

```vb
ExcelNeu.Sheet("Sheet1").UsedRange
```

- Hat Header: aktiviert
- Nur sichtbare Zeilen: deaktiviert
- Speichern unter: `dtNeu`

## Schritt 5: Eingelesene Daten pruefen

Füge ein `If` direkt unter die beiden Einlesebloecke.

### If: Tabellen vorhanden

- Condition:

```vb
dtAlt Is Nothing OrElse dtNeu Is Nothing OrElse dtAlt.Rows.Count = 0 OrElse dtNeu.Rows.Count = 0
```

### Then

Füge eine `Throw`-Activity ein.

- Exception:

```vb
New BusinessRuleException("Eine oder beide Eingabetabellen sind leer.")
```

### Else

- leer lassen

## Schritt 6: Invoke Code einhaengen

Füge jetzt den `Invoke Code`-Block direkt unter die Validierung.

### Invoke Code

- Language: `CSharp`

### Namespaces

Trage diese Namespaces ein:

- `System`
- `System.Data`
- `System.Linq`
- `System.Collections.Generic`
- `System.Globalization`

### Arguments

Lege diese Argumente an:

- `dtAlt`
  - Direction: `In`
  - Type: `System.Data.DataTable`
  - Value: `dtAlt`
- `dtNeu`
  - Direction: `In`
  - Type: `System.Data.DataTable`
  - Value: `dtNeu`
- `dtChanges`
  - Direction: `Out`
  - Type: `System.Data.DataTable`
  - Value: `dtChanges`

### Code

Den Inhalt aus dieser Datei komplett einfuegen:

- [InvokeCode_ExcelAbgleich.cs](/Users/personliches/Documents/process-automation-engine/docs/uipath/InvokeCode_ExcelAbgleich.cs)

## Schritt 7: Ergebnis pruefen

Füge direkt unter dem `Invoke Code` ein `If` ein.

### If: Aenderungen vorhanden

- Condition:

```vb
dtChanges Is Nothing OrElse dtChanges.Rows.Count = 0
```

### Then

Füge eine `Log Message`-Activity ein.

- Level: `Info`
- Message:

```vb
"Keine Aenderungen zwischen Vormonat und aktuellem Monat gefunden."
```

Danach kannst du den Workflow an dieser Stelle beenden.

### Else

Hier geht die Nachverarbeitung weiter.

## Schritt 8: Zaehler setzen

Innerhalb des `Else` nach Schritt 7 drei `Assign`-Activities einfuegen.

### Assign

- To: `intChangesCount`
- Value:

```vb
dtChanges.Rows.Count
```

### Log Message

- Level: `Info`
- Message:

```vb
"Anzahl Aenderungen gesamt: " + intChangesCount.ToString
```

## Schritt 9: Manuelle Faelle filtern

Füge eine `Filter Data Table`-Activity ein.

### Filter Data Table: manuelle Faelle

- Input: `dtChanges`
- Output: `dtManual`

Filterbedingung:

- Spalte: `ManuellePruefung`
- Operator: `=`
- Wert: `JA`

Spalten:

- alle Spalten beibehalten

Danach:

### Assign

- To: `intManualCount`
- Value:

```vb
If(dtManual Is Nothing, 0, dtManual.Rows.Count)
```

### Log Message

- Level: `Info`
- Message:

```vb
"Anzahl manueller Faelle: " + intManualCount.ToString
```

## Schritt 10: Automatische Faelle filtern

Füge direkt darunter eine zweite `Filter Data Table`-Activity ein.

### Filter Data Table: automatische Faelle

- Input: `dtChanges`
- Output: `dtAuto`

Filterbedingung:

- Spalte: `ManuellePruefung`
- Operator: `=`
- Wert: `NEIN`

Spalten:

- alle Spalten beibehalten

Danach:

### Assign

- To: `intAutoCount`
- Value:

```vb
If(dtAuto Is Nothing, 0, dtAuto.Rows.Count)
```

### Log Message

- Level: `Info`
- Message:

```vb
"Anzahl automatischer Faelle: " + intAutoCount.ToString
```

## Schritt 11: Ergebnisdatei schreiben

Füge unterhalb der beiden Filter eine `Use Excel File`-Activity ein.

### Use Excel File: Ausgabe

- Excel-Datei: `strOutputFile`
- Referenzieren als: `ExcelOut`
- `Aenderungen speichern`: aktiviert
- `Erstellen, falls nicht vorhanden`: aktiviert

Innerhalb dieser Activity drei `Write Range`-Activities einfuegen.

### Write Range 1: Gesamtergebnis

- Bereich:

```vb
ExcelOut.Sheet("Changes").Range("A1")
```

- Daten: `dtChanges`
- Kopfzeilen einbeziehen: aktiviert

### Write Range 2: manuelle Faelle

- Bereich:

```vb
ExcelOut.Sheet("ManualReview").Range("A1")
```

- Daten: `dtManual`
- Kopfzeilen einbeziehen: aktiviert

### Write Range 3: automatische Faelle

- Bereich:

```vb
ExcelOut.Sheet("AutoProcessing").Range("A1")
```

- Daten: `dtAuto`
- Kopfzeilen einbeziehen: aktiviert

Wenn `Write Range` Probleme macht, weil das Blatt noch nicht existiert, setze vorher jeweils `Write Range Workbook` ein oder lege die Sheets vorab an.

## Schritt 12: Abschluss-Logging

Unter dem Ausgabeblock eine `Log Message`.

- Level: `Info`
- Message:

```vb
"SOP-Monatsvergleich abgeschlossen. Ergebnisdatei: " + strOutputFile
```

## Schritt 13: Catch-Block

Im `Catch` des `Try Catch`:

### Assign

- To: `strErrorMessage`
- Value:

```vb
exception.Message
```

### Log Message

- Level: `Error`
- Message:

```vb
"Fehler im SOP-Monatsvergleich: " + strErrorMessage
```

Optional danach:

### Rethrow

Wenn der Prozess zentral behandelt werden soll, fuege `Rethrow` hinzu.

## Kompakte Reihenfolge der Activities

Die genaue Reihenfolge im Workflow sollte so aussehen:

1. `Try Catch`
2. `Assign` `strFileAlt`
3. `Assign` `strFileNeu`
4. `Assign` `strOutputFile`
5. `Path Exists` fuer `strFileAlt`
6. `Path Exists` fuer `strFileNeu`
7. `If` auf Dateiexistenz
8. `Use Excel File` Vormonat
9. `Read Range` nach `dtAlt`
10. `Use Excel File` aktueller Monat
11. `Read Range` nach `dtNeu`
12. `If` auf leere Tabellen
13. `Invoke Code`
14. `If` auf leeres `dtChanges`
15. `Assign` `intChangesCount`
16. `Log Message`
17. `Filter Data Table` nach `dtManual`
18. `Assign` `intManualCount`
19. `Log Message`
20. `Filter Data Table` nach `dtAuto`
21. `Assign` `intAutoCount`
22. `Log Message`
23. `Use Excel File` Ausgabe
24. `Write Range` `dtChanges`
25. `Write Range` `dtManual`
26. `Write Range` `dtAuto`
27. `Log Message` Abschluss
28. `Catch`
29. `Log Message` Fehler
30. optional `Rethrow`

## Wichtige Einstellungen, damit es stabil laeuft

- In `Read Range` immer `Hat Header` aktivieren
- bei Excel moeglichst `UsedRange` lesen statt nur `A1`
- `dtChanges` im `Invoke Code` unbedingt als `Out` setzen
- Spaltennamen in den Quelldateien muessen exakt stimmen:
  - `Anlagen-Nummer`
  - `Zählpunktbezeichnung`
  - `Prognosesmenge`
  - `Prognoses-ab`
  - `Prognoses-bis`
- falls Sonderzeichen in CSV-Dateien kaputt aussehen, nicht `Use Excel File`, sondern `Read CSV` mit passender Codierung verwenden

## Empfehlung fuer deinen aktuellen Stand

Da du `dtAlt` und `dtNeu` schon eingelesen hast und der `Invoke Code` bereits darunter sitzt, solltest du als Naechstes genau diese Bloecke ergaenzen:

1. `If` auf `dtChanges`
2. `Filter Data Table` fuer `dtManual`
3. `Filter Data Table` fuer `dtAuto`
4. `Use Excel File` fuer die Ergebnisdatei
5. drei `Write Range`-Activities
6. `Try Catch` und `Log Message`, falls noch nicht vorhanden
