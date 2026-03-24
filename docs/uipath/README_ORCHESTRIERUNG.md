# Orchestrierung des SOP-Monatsvergleichs

Diese Anleitung beschreibt, wie der bestehende UiPath-Bot so orchestriert wird, dass er taeglich einen Eingangsordner prueft, relevante Monatsdateien automatisch erkennt und den Prozess nur dann startet, wenn ein verarbeitbarer Dateisatz vorhanden ist.

## Zielbild

Der Bot soll:

1. einmal taeglich gestartet werden
2. einen definierten Eingangsordner pruefen
3. passende SOP-Dateien automatisch erkennen
4. aus den Dateien den aktuellen Monat und den Vormonat bestimmen
5. den Vergleich nur dann starten, wenn beide Monate vorhanden sind
6. erfolgreich verarbeitete Dateien nach `Processed` verschieben
7. fehlerhafte Dateien nach `Error` verschieben
8. bei fehlenden Dateien sauber mit Info-Log abbrechen

## Empfohlene Ordnerstruktur

Lege einen Hauptordner an, zum Beispiel:

```text
C:\SOP_Automation
```

Darin diese Unterordner:

```text
C:\SOP_Automation\Input
C:\SOP_Automation\Processed
C:\SOP_Automation\Error
C:\SOP_Automation\Output
```

## Empfohlene Dateinamenskonvention

Damit der Bot Monat und Jahr automatisch erkennt, sollten die Dateinamen ein einheitliches Muster haben.

Empfehlung:

```text
BKS_SOP_01_2026.xlsx
BKS_SOP_02_2026.xlsx
BKS_SOP_12_2026.xlsx
```

Das gleiche Prinzip funktioniert auch mit CSV:

```text
BKS_SOP_01_2026.csv
BKS_SOP_02_2026.csv
```

Wichtig:

- Monatsangabe immer zweistellig
- Jahr immer vierstellig
- gleiches Praefix fuer alle verarbeitbaren Dateien

## Warum Dateien nicht geloescht werden sollten

Verwendete Dateien sollten nicht direkt geloescht werden.

Besser:

- nach erfolgreicher Verarbeitung nach `Processed` verschieben
- bei Fehlern nach `Error` verschieben

Vorteile:

- Nachvollziehbarkeit
- Wiederanlauf moeglich
- manuelle Pruefung im Fehlerfall
- klare Trennung zwischen neu und bereits verarbeitet

## Empfohlene Workflow-Struktur

Der Workflow sollte in drei logische Bereiche aufgeteilt werden:

1. Orchestrierung
2. Dateiauswahl
3. Verarbeitung plus Nachlauf

## Neue Variablen

Lege diese Variablen auf Workflow-Ebene an:

- `strBaseFolder` (`String`)
- `strInputFolder` (`String`)
- `strProcessedFolder` (`String`)
- `strErrorFolder` (`String`)
- `strOutputFolder` (`String`)
- `arrInputFiles` (`String[]`)
- `arrCandidateFiles` (`String[]`)
- `strFileAlt` (`String`)
- `strFileNeu` (`String`)
- `strCurrentFileName` (`String`)
- `strPreviousFileName` (`String`)
- `strCurrentYearMonth` (`String`)
- `strPreviousYearMonth` (`String`)
- `blnCanRun` (`Boolean`)
- `strLogMessage` (`String`)
- `dtAlt` (`System.Data.DataTable`)
- `dtNeu` (`System.Data.DataTable`)
- `dtChanges` (`System.Data.DataTable`)

Optional:

- `strProcessedAltPath` (`String`)
- `strProcessedNeuPath` (`String`)
- `strErrorAltPath` (`String`)
- `strErrorNeuPath` (`String`)

## Schritt 1: Basisordner setzen

Füge am Anfang mehrere `Assign`-Activities ein.

### Assign

- To:

```csharp
strBaseFolder
```

- Value:

```csharp
@"C:\SOP_Automation"
```

### Assign

- To:

```csharp
strInputFolder
```

- Value:

```csharp
System.IO.Path.Combine(strBaseFolder, "Input")
```

### Assign

- To:

```csharp
strProcessedFolder
```

- Value:

```csharp
System.IO.Path.Combine(strBaseFolder, "Processed")
```

### Assign

- To:

```csharp
strErrorFolder
```

- Value:

```csharp
System.IO.Path.Combine(strBaseFolder, "Error")
```

### Assign

- To:

```csharp
strOutputFolder
```

- Value:

```csharp
System.IO.Path.Combine(strBaseFolder, "Output")
```

## Schritt 2: Ordner sicherstellen

Füge fuer jeden Zielordner eine `Create Folder`-Activity ein oder nutze `Invoke Code`/`Assign` mit `System.IO.Directory.CreateDirectory(...)`.

Beispiel per `Invoke Code` oder `Assign`:

```csharp
System.IO.Directory.CreateDirectory(strInputFolder)
```

```csharp
System.IO.Directory.CreateDirectory(strProcessedFolder)
```

```csharp
System.IO.Directory.CreateDirectory(strErrorFolder)
```

```csharp
System.IO.Directory.CreateDirectory(strOutputFolder)
```

## Schritt 3: Dateien aus Input lesen

Füge eine `Assign`-Activity ein.

- To:

```csharp
arrInputFiles
```

- Value:

```csharp
System.IO.Directory.GetFiles(strInputFolder)
```

## Schritt 4: Auf leeren Input pruefen

Füge ein `If` ein.

- Condition:

```csharp
arrInputFiles == null || arrInputFiles.Length == 0
```

### Then

`Log Message`:

```csharp
"Keine Dateien im Input-Ordner gefunden. Workflow wird beendet."
```

Danach Workflow beenden.

## Schritt 5: Relevante SOP-Dateien filtern

Der Bot soll nicht jede Datei im Input verarbeiten, sondern nur passende SOP-Dateien.

Füge eine `Assign`-Activity ein.

- To:

```csharp
arrCandidateFiles
```

- Value:

```csharp
arrInputFiles
    .Where(f =>
        System.Text.RegularExpressions.Regex.IsMatch(
            System.IO.Path.GetFileName(f),
            @"^BKS_SOP_\d{2}_\d{4}\.(xlsx|csv)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase))
    .ToArray()
```

## Schritt 6: Auf zu wenige Kandidaten pruefen

Füge ein `If` ein.

- Condition:

```csharp
arrCandidateFiles == null || arrCandidateFiles.Length < 2
```

### Then

`Log Message`:

```csharp
"Nicht genug passende SOP-Dateien im Input-Ordner gefunden. Mindestens zwei Monatsdateien werden benoetigt."
```

Danach Workflow beenden.

## Schritt 7: Monat und Jahr aus den Dateinamen bestimmen

Die Dateien muessen jetzt nach Jahr und Monat sortiert werden.

Die sauberste Loesung ist hier ein kleiner `Invoke Code`-Block nur fuer die Dateiauswahl.

### Neuer Invoke Code: Dateiauswahl

Language:

- `CSharp`

Namespaces:

- `System`
- `System.Linq`
- `System.IO`
- `System.Text.RegularExpressions`
- `System.Collections.Generic`

Arguments:

- `arrCandidateFiles` (`In`, `String[]`)
- `strFileAlt` (`Out`, `String`)
- `strFileNeu` (`Out`, `String`)
- `strCurrentYearMonth` (`Out`, `String`)
- `strPreviousYearMonth` (`Out`, `String`)
- `blnCanRun` (`Out`, `Boolean`)

Code:

```csharp
var regex = new Regex(@"^BKS_SOP_(\d{2})_(\d{4})\.(xlsx|csv)$", RegexOptions.IgnoreCase);

var files = arrCandidateFiles
    .Select(path =>
    {
        var name = Path.GetFileName(path);
        var match = regex.Match(name);
        if (!match.Success) return null;

        int month = int.Parse(match.Groups[1].Value);
        int year = int.Parse(match.Groups[2].Value);

        return new
        {
            Path = path,
            Month = month,
            Year = year,
            SortKey = year * 100 + month
        };
    })
    .Where(x => x != null)
    .OrderByDescending(x => x.SortKey)
    .ToList();

blnCanRun = false;
strFileAlt = string.Empty;
strFileNeu = string.Empty;
strCurrentYearMonth = string.Empty;
strPreviousYearMonth = string.Empty;

if (files.Count >= 2)
{
    var current = files[0];
    var previous = files[1];

    int expectedPrevious = current.Month == 1
        ? ((current.Year - 1) * 100) + 12
        : (current.Year * 100) + (current.Month - 1);

    if (previous.SortKey == expectedPrevious)
    {
        strFileNeu = current.Path;
        strFileAlt = previous.Path;
        strCurrentYearMonth = current.Year.ToString() + "_" + current.Month.ToString("00");
        strPreviousYearMonth = previous.Year.ToString() + "_" + previous.Month.ToString("00");
        blnCanRun = true;
    }
}
```

## Schritt 8: Prüfen, ob ein gültiges Monatspaar vorliegt

Füge ein `If` ein.

- Condition:

```csharp
!blnCanRun
```

### Then

`Log Message`:

```csharp
"Es wurde kein gueltiges Paar aus aktuellem Monat und Vormonat gefunden. Workflow wird beendet."
```

Danach Workflow beenden.

### Else

Hier startet der eigentliche Vergleich.

## Schritt 9: Vorhandene Vergleichslogik nutzen

Ab hier wird der bestehende Workflow verwendet:

1. `Use Excel File` oder `Read CSV` fuer `strFileAlt`
2. `Use Excel File` oder `Read CSV` fuer `strFileNeu`
3. `Invoke Code` fuer den SOP-Vergleich
4. Ergebnisgruppen bilden
5. Ergebnisse exportieren

Wichtig:

- `strFileAlt` kommt jetzt aus der Dateiauswahl
- `strFileNeu` kommt jetzt aus der Dateiauswahl

## Schritt 10: Output-Dateinamen dynamisch erzeugen

Die Output-Dateien sollten aus `strCurrentYearMonth` erzeugt werden.

### Beispiel Assign

- To:

```csharp
strOutputChangesCsv
```

- Value:

```csharp
System.IO.Path.Combine(strOutputFolder, "SOP_Changes_" + strCurrentYearMonth + ".csv")
```

Weitere Beispiele:

```csharp
System.IO.Path.Combine(strOutputFolder, "SOP_NeueAnlagen_" + strCurrentYearMonth + ".csv")
```

```csharp
System.IO.Path.Combine(strOutputFolder, "SOP_Zusammengefuehrt_" + strCurrentYearMonth + ".csv")
```

```csharp
System.IO.Path.Combine(strOutputFolder, "SOP_Wegfaelle_" + strCurrentYearMonth + ".csv")
```

```csharp
System.IO.Path.Combine(strOutputFolder, "SOP_ManualReview_" + strCurrentYearMonth + ".csv")
```

## Schritt 11: Dateien nach erfolgreicher Verarbeitung verschieben

Nach erfolgreichem Lauf sollten die beiden verwendeten Dateien verschoben werden.

### Assign

- To:

```csharp
strProcessedAltPath
```

- Value:

```csharp
System.IO.Path.Combine(strProcessedFolder, System.IO.Path.GetFileName(strFileAlt))
```

### Assign

- To:

```csharp
strProcessedNeuPath
```

- Value:

```csharp
System.IO.Path.Combine(strProcessedFolder, System.IO.Path.GetFileName(strFileNeu))
```

Danach je eine `Move File`-Activity:

- Source: `strFileAlt`
- Destination: `strProcessedAltPath`

- Source: `strFileNeu`
- Destination: `strProcessedNeuPath`

## Schritt 12: Fehlerfälle nach Error verschieben

Im `Catch`-Block kannst du dieselbe Logik fuer den Fehlerfall verwenden.

### Assign

- To:

```csharp
strErrorAltPath
```

- Value:

```csharp
string.IsNullOrWhiteSpace(strFileAlt) ? string.Empty : System.IO.Path.Combine(strErrorFolder, System.IO.Path.GetFileName(strFileAlt))
```

### Assign

- To:

```csharp
strErrorNeuPath
```

- Value:

```csharp
string.IsNullOrWhiteSpace(strFileNeu) ? string.Empty : System.IO.Path.Combine(strErrorFolder, System.IO.Path.GetFileName(strFileNeu))
```

Danach jeweils nur verschieben, wenn der Dateipfad nicht leer ist.

## Empfohlene Logs

Mindestens diese Logs solltest du setzen:

- `Input-Ordner wird geprueft: ...`
- `Gefundene Dateien gesamt: ...`
- `Passende SOP-Dateien: ...`
- `Verwendeter aktueller Monat: ...`
- `Verwendeter Vormonat: ...`
- `Kein gueltiger Dateisatz vorhanden, Workflow beendet`
- `Dateien nach Processed verschoben`
- `Dateien nach Error verschoben`

## Empfohlene Reihenfolge im Workflow

1. `Assign` Basisordner
2. `Assign` Unterordner
3. `Create Folder` fuer Input/Processed/Error/Output
4. `Assign` `arrInputFiles`
5. `If` auf leeren Input
6. `Assign` `arrCandidateFiles`
7. `If` auf weniger als zwei Kandidaten
8. `Invoke Code` zur Auswahl von aktuellem Monat und Vormonat
9. `If !blnCanRun`
10. eigentliche Prozesslogik mit `strFileAlt` und `strFileNeu`
11. Ergebnisexport
12. `Move File` nach `Processed`
13. im Fehlerfall `Move File` nach `Error`

## Empfehlung fuer den produktiven Betrieb

Fuer einen hohen Automatisierungsgrad ist der beste Weg:

- taeglicher Trigger
- klar definierter `Input`-Ordner
- automatische Monatserkennung aus Dateinamen
- keine Loeschung
- stattdessen Verschieben nach `Processed`

Damit bleibt der Prozess nachvollziehbar, wiederanlaufbar und auch ueber Monats- und Jahreswechsel stabil.
