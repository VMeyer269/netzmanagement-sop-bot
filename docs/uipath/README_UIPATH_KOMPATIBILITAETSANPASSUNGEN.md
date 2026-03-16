# UiPath Kompatibilitaetsanpassungen

Diese Datei dokumentiert die Anpassungen, die fuer die aktuelle UiPath-Umgebung noetig sind, damit der SOP-Monatsvergleich ohne Syntax- und Exportprobleme laeuft.

## Hintergrund

Die aktuellen Fehler entstehen an zwei Stellen:

1. Die Activity-Ausdruecke in UiPath werden im Projekt als `C#` ausgewertet, einige Bedingungen und Exceptions wurden aber in `VB`-Syntax eingetragen.
2. Die lokale UiPath-Version stellt `Write Range` in der vorhandenen Konfiguration nur fuer Cloud- bzw. OneDrive-Szenarien bereit. Fuer den lokalen Betrieb ist deshalb ein Export ueber `Write CSV` der saubere Weg.

## Grundregel

In allen `If`-, `Assign`- und `Throw`-Feldern muss C#-Syntax verwendet werden.

Nicht verwenden:

- `Not`
- `OrElse`
- `Is Nothing`
- `New ...`

Stattdessen verwenden:

- `!`
- `||`
- `== null`
- `new ...`

## Korrekturen im Workflow

## 1. If-Funktion 1

Zweck:
- prueft, ob mindestens eine Eingabedatei fehlt

Falsche VB-Syntax:

```vb
Not blnAltExists OrElse Not blnNeuExists
```

Korrekte C#-Syntax:

```csharp
!blnAltExists || !blnNeuExists
```

## 2. Throw-Funktion 1

Zweck:
- wirft einen Fehler, wenn eine Eingabedatei fehlt

Falsche VB-Syntax:

```vb
New BusinessRuleException("Mindestens eine Eingabedatei wurde nicht gefunden. Vormonat: " + strFileAlt + " | Aktueller Monat: " + strFileNeu)
```

Korrekte C#-Syntax:

```csharp
new UiPath.Core.BusinessRuleException("Mindestens eine Eingabedatei wurde nicht gefunden. Vormonat: " + strFileAlt + " | Aktueller Monat: " + strFileNeu)
```

## 3. If-Funktion 2

Zweck:
- prueft, ob `dtAlt` oder `dtNeu` leer oder nicht gesetzt sind

Falsche VB-Syntax:

```vb
dtAlt Is Nothing OrElse dtNeu Is Nothing OrElse dtAlt.Rows.Count = 0 OrElse dtNeu.Rows.Count = 0
```

Korrekte C#-Syntax:

```csharp
dtAlt == null || dtNeu == null || dtAlt.Rows.Count == 0 || dtNeu.Rows.Count == 0
```

## 4. Throw-Funktion 2

Zweck:
- wirft einen Fehler, wenn keine Daten eingelesen wurden

Korrekte C#-Syntax:

```csharp
new UiPath.Core.BusinessRuleException("Eine oder beide Eingabetabellen sind leer.")
```

## 5. If-Funktion 3

Zweck:
- prueft, ob der `Invoke Code` ueberhaupt Aenderungen geliefert hat

Falsche VB-Syntax:

```vb
dtChanges Is Nothing OrElse dtChanges.Rows.Count = 0
```

Korrekte C#-Syntax:

```csharp
dtChanges == null || dtChanges.Rows.Count == 0
```

## Tabellenfilter anpassen

Die Activity `Filter Data Table` kann in der aktuellen Umgebung Probleme machen, insbesondere bei der Spaltenkonfiguration. Deshalb wird empfohlen, die Filterung ueber `Assign` plus `Select(...).CopyToDataTable()` umzusetzen.

## Warum `Filter Data Table` hier fehlschlaegt

Der Fehler mit ungueltigen Spalten tritt in diesem Workflow sehr wahrscheinlich deshalb auf, weil die Spalte `ManuellePruefung` erst im `Invoke Code` erzeugt wird. UiPath versucht die Spalten in `Filter Data Table` aber oft schon zur Designzeit zu validieren.

Dadurch entstehen typische Probleme:

- `dtChanges` ist zur Designzeit noch leer
- die Zusatzspalten aus dem `Invoke Code` sind noch nicht bekannt
- UiPath markiert die Filterspalte deshalb als ungueltig

Fuer diesen Bot ist es deshalb sinnvoll, `Filter Data Table` vollstaendig zu ersetzen.

## Variablen

Diese Variablen muessen vorhanden sein:

- `dtChanges` (`System.Data.DataTable`)
- `dtManual` (`System.Data.DataTable`)
- `dtAuto` (`System.Data.DataTable`)

## dtManual per Assign setzen

Füge eine `Assign`-Activity ein.

- To:

```csharp
dtManual
```

- Value:

```csharp
dtChanges.Select("ManuellePruefung = 'JA'").Length > 0
    ? dtChanges.Select("ManuellePruefung = 'JA'").CopyToDataTable()
    : dtChanges.Clone()
```

## dtAuto per Assign setzen

Füge eine zweite `Assign`-Activity ein.

- To:

```csharp
dtAuto
```

- Value:

```csharp
dtChanges.Select("ManuellePruefung = 'NEIN'").Length > 0
    ? dtChanges.Select("ManuellePruefung = 'NEIN'").CopyToDataTable()
    : dtChanges.Clone()
```

## Konkrete Umsetzung in UiPath

Die beiden `Filter Data Table`-Activities sollten aus dem Workflow entfernt und durch zwei `Assign`-Activities ersetzt werden.

### Assign statt Filter 1

- To:

```csharp
dtManual
```

- Value:

```csharp
dtChanges.Select("ManuellePruefung = 'JA'").Length > 0
    ? dtChanges.Select("ManuellePruefung = 'JA'").CopyToDataTable()
    : dtChanges.Clone()
```

### Assign statt Filter 2

- To:

```csharp
dtAuto
```

- Value:

```csharp
dtChanges.Select("ManuellePruefung = 'NEIN'").Length > 0
    ? dtChanges.Select("ManuellePruefung = 'NEIN'").CopyToDataTable()
    : dtChanges.Clone()
```

## Vorteil dieser Loesung

Diese Variante ist fuer euren Workflow robuster als `Filter Data Table`, weil:

- keine Designzeit-Validierung auf unbekannte Spalten erfolgt
- die Filterung direkt auf der bereits erzeugten `dtChanges`-Tabelle arbeitet
- manuelle und automatische Faelle ohne zusaetzliche Activity-Probleme getrennt werden

## Zaehlerwerte setzen

### intManualCount

- To:

```csharp
intManualCount
```

- Value:

```csharp
dtManual == null ? 0 : dtManual.Rows.Count
```

### intAutoCount

- To:

```csharp
intAutoCount
```

- Value:

```csharp
dtAuto == null ? 0 : dtAuto.Rows.Count
```

## Alternative zu Write Range

Da `Write Range` in der aktuellen UiPath-Version nur ueber Cloud-/OneDrive-Szenarien verfuegbar ist, wird fuer den lokalen Betrieb `Write CSV` verwendet.

## Neue String-Variablen

Lege diese Variablen an:

- `strOutputChangesCsv` (`String`)
- `strOutputManualCsv` (`String`)
- `strOutputAutoCsv` (`String`)

## Beispielwerte

```csharp
strOutputChangesCsv = @"C:\Users\swhrpant\Desktop\SOP_Changes_2026_02.csv";
strOutputManualCsv = @"C:\Users\swhrpant\Desktop\SOP_ManualReview_2026_02.csv";
strOutputAutoCsv = @"C:\Users\swhrpant\Desktop\SOP_AutoProcessing_2026_02.csv";
```

## Empfohlene Export-Activities

Nach dem `Invoke Code` und nach der Trennung in `dtManual` und `dtAuto`:

1. `Write CSV`
   - DataTable: `dtChanges`
   - FilePath: `strOutputChangesCsv`
   - IncludeColumnNames: `True`
2. `Write CSV`
   - DataTable: `dtManual`
   - FilePath: `strOutputManualCsv`
   - IncludeColumnNames: `True`
3. `Write CSV`
   - DataTable: `dtAuto`
   - FilePath: `strOutputAutoCsv`
   - IncludeColumnNames: `True`

## Aktualisierte Activity-Reihenfolge

Die Reihenfolge im Hauptworkflow sollte damit so aussehen:

1. `Assign` fuer `strFileAlt`
2. `Assign` fuer `strFileNeu`
3. `Assign` fuer `strOutputChangesCsv`
4. `Assign` fuer `strOutputManualCsv`
5. `Assign` fuer `strOutputAutoCsv`
6. `File Exists` fuer `strFileAlt`
7. `File Exists` fuer `strFileNeu`
8. `If` mit:

```csharp
!blnAltExists || !blnNeuExists
```

9. `Throw` mit:

```csharp
new UiPath.Core.BusinessRuleException("Mindestens eine Eingabedatei wurde nicht gefunden. Vormonat: " + strFileAlt + " | Aktueller Monat: " + strFileNeu)
```

10. `Use Excel File` oder `Read CSV` fuer `dtAlt`
11. `Use Excel File` oder `Read CSV` fuer `dtNeu`
12. `If` mit:

```csharp
dtAlt == null || dtNeu == null || dtAlt.Rows.Count == 0 || dtNeu.Rows.Count == 0
```

13. `Throw` mit:

```csharp
new UiPath.Core.BusinessRuleException("Eine oder beide Eingabetabellen sind leer.")
```

14. `Invoke Code`
15. `If` mit:

```csharp
dtChanges == null || dtChanges.Rows.Count == 0
```

16. `Log Message` fuer "Keine Aenderungen gefunden"
17. `Assign` fuer `dtManual`
18. `Assign` fuer `intManualCount`
19. `Assign` fuer `dtAuto`
20. `Assign` fuer `intAutoCount`
21. `Write CSV` fuer `dtChanges`
22. `Write CSV` fuer `dtManual`
23. `Write CSV` fuer `dtAuto`

## Hinweis fuer den aktuellen Workflow-Stand

Wenn der Workflow aktuell noch `Filter Data Table` oder `Write Range` enthaelt, sollten diese Activities ersetzt werden:

- `Filter Data Table` ersetzen durch `Assign` auf `dtManual` und `dtAuto`
- `Write Range` ersetzen durch `Write CSV`

## Empfehlung

Fuer die aktuelle Projektphase ist der stabilste Weg:

- lokale Dateien lesen
- C#-kompatible Bedingungen und Exceptions verwenden
- Ergebnis lokal als CSV schreiben

Damit bleibt der Bot voll lauffaehig, ohne Cloud-Anbindung oder OneDrive-Berechtigungen.
