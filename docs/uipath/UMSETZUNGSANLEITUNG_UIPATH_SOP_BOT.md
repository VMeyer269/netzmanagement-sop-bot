# Umsetzungsanleitung: UiPath-Anbindung fuer den SOP-Monatsvergleich

Diese Anleitung beschreibt, wie der C#-`Invoke Code`-Block aus [InvokeCode_ExcelAbgleich.cs](/Users/personliches/Documents/process-automation-engine/docs/uipath/InvokeCode_ExcelAbgleich.cs) in einen arbeitsfaehigen UiPath-Workflow eingebunden wird.

## Zielbild

Der Workflow soll:

1. die CSV- oder Excel-Datei des Vormonats einlesen
2. die CSV- oder Excel-Datei des aktuellen Monats einlesen
3. beide Tabellen an den `Invoke Code`-Block uebergeben
4. die Ergebnistabelle `dtChanges` zur Weiterverarbeitung bereitstellen
5. manuell zu pruefende Faelle gesondert markieren oder exportieren

## Empfohlener Workflow-Aufbau

Empfohlene Reihenfolge in einer `Sequence`:

1. `Assign` fuer Pfadvariablen
2. `Path Exists` oder `File Exists` fuer beide Quelldateien
3. `If` fuer Dateipruefung mit sauberer Fehlermeldung
4. `Read CSV` oder `Use Excel File` plus `Read Range` fuer Vormonat
5. `Read CSV` oder `Use Excel File` plus `Read Range` fuer aktuellen Monat
6. `Invoke Code` fuer den Vergleich
7. `Filter Data Table` oder `For Each Row in Data Table` fuer Nachbearbeitung
8. `Write Range Workbook` oder `Write CSV` fuer Ergebnisexport
9. optional `Log Message` und `Send Mail` fuer Reporting

## Variablen

Diese Variablen solltest du mindestens auf Workflow-Ebene definieren:

- `strFileAlt` (`String`)
  - Vollstaendiger Dateipfad zum Vormonat
- `strFileNeu` (`String`)
  - Vollstaendiger Dateipfad zum aktuellen Monat
- `strOutputPath` (`String`)
  - Zielpfad fuer die Ausgabe
- `dtAlt` (`System.Data.DataTable`)
  - eingelesene Tabelle des Vormonats
- `dtNeu` (`System.Data.DataTable`)
  - eingelesene Tabelle des aktuellen Monats
- `dtChanges` (`System.Data.DataTable`)
  - Ergebnis des Invoke-Code-Blocks
- `dtManual` (`System.Data.DataTable`)
  - optional gefilterte Tabelle nur mit `ManuellePruefung = "JA"`
- `dtAuto` (`System.Data.DataTable`)
  - optional gefilterte Tabelle nur mit `ManuellePruefung = "NEIN"`
- `blnAltExists` (`Boolean`)
  - Kennzeichen, ob die Vormonatsdatei vorhanden ist
- `blnNeuExists` (`Boolean`)
  - Kennzeichen, ob die aktuelle Datei vorhanden ist
- `strLogMessage` (`String`)
  - fuer Status- und Fehlerausgaben

## Argumente fuer den Invoke-Code-Block

Im `Invoke Code`-Block:

Language:
- `CSharp`

Namespaces:
- `System`
- `System.Data`
- `System.Linq`
- `System.Collections.Generic`
- `System.Globalization`

Arguments:
- `dtAlt` (`In`, `System.Data.DataTable`)
- `dtNeu` (`In`, `System.Data.DataTable`)
- `dtChanges` (`Out`, `System.Data.DataTable`)

Wichtig:
- `dtChanges` muss als `Out` gesetzt sein.
- `dtAlt` und `dtNeu` muessen vor dem Aufruf bereits befuellt sein.
- Der Inhalt aus [InvokeCode_ExcelAbgleich.cs](/Users/personliches/Documents/process-automation-engine/docs/uipath/InvokeCode_ExcelAbgleich.cs) wird 1:1 in die Aktivitaet eingefuegt.

## Dateieinlesung

## Variante A: CSV-Dateien

Fuer deine aktuellen Beispieldateien ist `Read CSV` der direkte Weg.

Empfohlene Aktivitaeten:

1. `Assign`
   - `strFileAlt = "C:\\...\\BKS_SOP_01_2026.csv"`
   - `strFileNeu = "C:\\...\\BK_SOP_02_2026.csv"`
2. `Path Exists`
   - prueft `strFileAlt`
   - Ausgabe nach `blnAltExists`
3. `Path Exists`
   - prueft `strFileNeu`
   - Ausgabe nach `blnNeuExists`
4. `If Not blnAltExists OrElse Not blnNeuExists`
   - dann `Throw` oder `Log Message`
5. `Read CSV` fuer `strFileAlt`
   - Output: `dtAlt`
   - Delimiter: `;`
   - Encoding: moeglichst `Windows-1252` oder `System.Text.Encoding.Default`, falls Sonderzeichen sonst falsch gelesen werden
   - AddHeaders: `True`
6. `Read CSV` fuer `strFileNeu`
   - Output: `dtNeu`
   - gleiche Einstellungen wie oben

## Variante B: Excel-Dateien

Wenn die Dateien spaeter als `.xlsx` kommen:

1. `Use Excel File`
2. darin `Read Range`
3. Option `AddHeaders` aktivieren
4. jeweils in `dtAlt` und `dtNeu` einlesen

Wichtig bei Excel:
- Nur das relevante Tabellenblatt lesen
- Keine Leerzeilen oberhalb der Header zulassen
- Sicherstellen, dass die Spaltennamen exakt den erwarteten Bezeichnungen entsprechen

## Erwartete Pflichtspalten

Der `Invoke Code` erwartet mindestens diese Spalten:

- `Anlagen-Nummer`
- `Zählpunktbezeichnung`
- `Prognosesmenge`
- `Prognoses-ab`
- `Prognoses-bis`

Weitere Spalten werden uebernommen, sind aber fuer die Kernlogik optional.

## Invoke-Code einbinden

In UiPath:

1. `Invoke Code` einfuegen
2. Language auf `CSharp` setzen
3. die genannten Namespaces eintragen
4. die Argumente `dtAlt`, `dtNeu`, `dtChanges` anlegen
5. den Code aus [InvokeCode_ExcelAbgleich.cs](/Users/personliches/Documents/process-automation-engine/docs/uipath/InvokeCode_ExcelAbgleich.cs) einfuegen

Nach dem Lauf enthaelt `dtChanges`:

- Originalspalten aus den Quelldateien
- `Status`
- `ManuellePruefung`
- `Hinweis`
- `AnzahlKonsolidierterZeilen`

## Empfohlene Nachbearbeitung in UiPath

### 1. Manuelle Faelle abtrennen

Mit `Filter Data Table`:

- Bedingung `ManuellePruefung = "JA"`
- Ergebnis nach `dtManual`

Danach ein zweiter Filter:

- Bedingung `ManuellePruefung = "NEIN"`
- Ergebnis nach `dtAuto`

So kannst du automatische und manuelle Faelle getrennt weiterverarbeiten.

### 2. Ergebnisse exportieren

Empfohlene Ausgabe:

- `dtChanges` als Gesamtergebnis
- `dtManual` als Datei fuer manuelle Bearbeitung
- `dtAuto` als Datei fuer automatische Folgeprozesse

Sinnvolle Aktivitaeten:

- `Write Range Workbook`
- `Write CSV`
- optional `Create Folder`, falls Zielverzeichnis dynamisch angelegt werden soll

## Empfohlene Zusatzvariablen fuer sauberen Betrieb

Wenn du den Bot robuster machen willst, sind diese Variablen sinnvoll:

- `intChangeCount` (`Int32`)
  - `dtChanges.Rows.Count`
- `intManualCount` (`Int32`)
  - `dtManual.Rows.Count`
- `intAutoCount` (`Int32`)
  - `dtAuto.Rows.Count`
- `strRunMonth` (`String`)
  - z. B. `"2026-02"`
- `strPreviousMonth` (`String`)
  - z. B. `"2026-01"`

Diese Werte kannst du fuer Logs, Dateinamen und E-Mail-Betreffs verwenden.

## Empfohlene Logs

Mindestens diese `Log Message`-Punkte sind sinnvoll:

- Start des Workflows
- verwendeter Dateipfad fuer Vormonat
- verwendeter Dateipfad fuer aktuellen Monat
- Anzahl eingelesener Zeilen in `dtAlt`
- Anzahl eingelesener Zeilen in `dtNeu`
- Anzahl Zeilen in `dtChanges`
- Anzahl manueller Faelle
- Anzahl automatischer Faelle
- Export erfolgreich abgeschlossen

## Typische Fehlerquellen

### 1. Falscher CSV-Delimiter

Deine Beispieldateien sind mit `;` getrennt. Wenn UiPath mit `,` liest, verschiebt sich die komplette Tabelle in eine einzige Spalte.

### 2. Falsche Zeichenkodierung

Wenn `Zählpunktbezeichnung` oder andere Umlaute kaputt aussehen, stimmt meistens die CSV-Encoding-Einstellung nicht.

### 3. Header stimmen nicht exakt

Schon kleine Abweichungen bei den Pflichtspalten fuehren im `Invoke Code` zu einem Fehler.

### 4. `dtChanges` nicht als Out-Argument gesetzt

Dann liefert der Block zwar intern ein Ergebnis, aber nichts kommt im Workflow an.

### 5. Leere oder unvollstaendige Eingabedateien

Vor dem `Invoke Code` immer Dateiexistenz und nach dem Einlesen die Zeilenanzahl pruefen.

## Empfohlene Schutzmechanismen

Fuer einen reibungslosen Ablauf solltest du diese Punkte einbauen:

1. `Try Catch` um den kompletten Import- und Vergleichsblock
2. `If dtAlt Is Nothing OrElse dtAlt.Rows.Count = 0`
3. `If dtNeu Is Nothing OrElse dtNeu.Rows.Count = 0`
4. klare Fehlermeldung bei fehlenden Pflichtspalten
5. Export der manuellen Faelle in eine eigene Datei
6. nachvollziehbare Dateinamen mit Monat und Laufdatum

## Beispiel fuer sinnvolle Ausgabedateien

- `SOP_Changes_2026_02.xlsx`
- `SOP_ManualReview_2026_02.xlsx`
- `SOP_AutoProcessing_2026_02.xlsx`

## Minimaler produktiver Ablauf

Wenn du zuerst nur eine lauffaehige Basis willst, reicht dieser Flow:

1. `Assign` der beiden Dateipfade
2. `Read CSV` fuer Vormonat nach `dtAlt`
3. `Read CSV` fuer aktuellen Monat nach `dtNeu`
4. `Invoke Code` mit `dtAlt`, `dtNeu`, `dtChanges`
5. `Write Range Workbook` fuer `dtChanges`

## Empfehlung fuer den naechsten Ausbauschritt

Nach der ersten lauffaehigen Version solltest du als Naechstes:

1. `dtManual` separat exportieren
2. Fehlerbehandlung mit `Try Catch` einbauen
3. Dateinamen dynamisch aus Monat und Jahr ableiten
4. Logging fuer Zeilenanzahlen und Sonderfaelle ergaenzen
