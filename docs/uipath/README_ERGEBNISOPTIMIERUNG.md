# Ergebnisoptimierung

Diese Datei beschreibt die naechsten sinnvollen Anpassungen an der Bot-Ausgabe, damit die fachlichen Ergebnisse klarer getrennt und besser nutzbar werden.

## Aktueller Stand

Der Bot liefert bereits fachlich brauchbare Ergebnisse:

- neue Faelle werden erkannt
- Wegfaelle werden erkannt
- zusammengefuehrte Zaehlpunkte werden erkannt
- manuelle Prueffaelle koennen markiert werden

Die aktuelle Ausgabe ist aber noch zu stark gemischt, weil alle Ergebnisse zusammen in `dtChanges` landen.

## Ziel der Optimierung

Die Ergebnisse sollen in fachlich getrennte Gruppen aufgeteilt werden:

1. `Neue Anlagen im Februar`
2. `Zusammengefuehrte Zaehlpunkte`
3. `Wegfaelle`
4. `Manuelle Pruefung`

Damit entspricht die Bot-Ausgabe besser der manuellen Fachsicht.

## Empfohlene neue DataTable-Variablen

Zusaetzlich zu `dtChanges` sollten diese Variablen angelegt werden:

- `dtNeuRegulaer` (`System.Data.DataTable`)
- `dtKonsolidiert` (`System.Data.DataTable`)
- `dtWegfall` (`System.Data.DataTable`)
- `dtManual` (`System.Data.DataTable`)

## Fachliche Trennung

## 1. Neue Anlagen im Februar

Diese Gruppe soll nur die regulaeren Neuzugaenge enthalten:

- `Status = NEU`
- `ManuellePruefung = NEIN`
- kein Hinweis auf Zusammenfuehrung

### UiPath Assign

- To:

```csharp
dtNeuRegulaer
```

- Value:

```csharp
dtChanges.Select("Status = 'NEU' AND ManuellePruefung = 'NEIN' AND (Hinweis IS NULL OR Hinweis = '')").Length > 0
    ? dtChanges.Select("Status = 'NEU' AND ManuellePruefung = 'NEIN' AND (Hinweis IS NULL OR Hinweis = '')").CopyToDataTable()
    : dtChanges.Clone()
```

## 2. Zusammengefuehrte Zaehlpunkte

Diese Gruppe soll alle Faelle enthalten, bei denen mehrere Zeilen zu einem Ergebnis zusammengefuehrt wurden.

Praktisch erkennbar an:

- `Status = NEU`
- `Hinweis <> ''`

### UiPath Assign

- To:

```csharp
dtKonsolidiert
```

- Value:

```csharp
dtChanges.Select("Status = 'NEU' AND Hinweis <> ''").Length > 0
    ? dtChanges.Select("Status = 'NEU' AND Hinweis <> ''").CopyToDataTable()
    : dtChanges.Clone()
```

Optional robuster kann spaeter auch `AnzahlKonsolidierterZeilen > 1` als weiteres Kriterium verwendet werden.

## 3. Wegfaelle

Diese Gruppe soll alle Kunden oder Anlagen enthalten, die im aktuellen Monat nicht mehr vorhanden sind.

Kriterium:

- `Status = WEG`

### UiPath Assign

- To:

```csharp
dtWegfall
```

- Value:

```csharp
dtChanges.Select("Status = 'WEG'").Length > 0
    ? dtChanges.Select("Status = 'WEG'").CopyToDataTable()
    : dtChanges.Clone()
```

## 4. Manuelle Pruefung

Diese Gruppe enthaelt alle Faelle, die nicht automatisch entschieden werden sollen.

Kriterium:

- `ManuellePruefung = JA`

### UiPath Assign

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

## Empfohlene Reihenfolge im Workflow

Nach dem `Invoke Code` und nach der Pruefung auf leeres `dtChanges`:

1. `Assign` fuer `dtNeuRegulaer`
2. `Assign` fuer `dtKonsolidiert`
3. `Assign` fuer `dtWegfall`
4. `Assign` fuer `dtManual`

Danach optional weitere `Assign`-Bloecke fuer die Mengen:

- `intNeuRegulaerCount`
- `intKonsolidiertCount`
- `intWegfallCount`
- `intManualCount`

## Empfohlene Zaehler

### Neue Anlagen

```csharp
dtNeuRegulaer == null ? 0 : dtNeuRegulaer.Rows.Count
```

### Zusammengefuehrte Zaehlpunkte

```csharp
dtKonsolidiert == null ? 0 : dtKonsolidiert.Rows.Count
```

### Wegfaelle

```csharp
dtWegfall == null ? 0 : dtWegfall.Rows.Count
```

### Manuelle Pruefung

```csharp
dtManual == null ? 0 : dtManual.Rows.Count
```

## Empfohlene Exportstruktur

Die Ausgabe sollte nicht mehr nur als eine einzige Gesamttabelle weitergegeben werden, sondern fachlich getrennt exportiert werden.

Empfohlene Exportdateien:

- `SOP_NeueAnlagen_2026_02.csv`
- `SOP_Zusammengefuehrt_2026_02.csv`
- `SOP_Wegfaelle_2026_02.csv`
- `SOP_ManualReview_2026_02.csv`

Optional weiterhin zusaetzlich:

- `SOP_Changes_2026_02.csv` als Gesamtsicht

## Fachlicher Nutzen

Durch diese Aufteilung wird die Ausgabe deutlich besser lesbar:

- neue Faelle sind sofort sichtbar
- zusammengefuehrte Faelle sind separat nachvollziehbar
- Wegfaelle bleiben weiterhin erhalten
- manuelle Faelle gehen nicht in der Gesamtmenge unter

## Bewertung des aktuellen Bot-Verhaltens

Der aktuelle Bot ist fachlich bereits auf einem guten Stand, weil er:

- regulaere Neuzugaenge erkennt
- Wegfaelle erkennt
- Konsolidierungen erkennt

Die naechste Optimierung betrifft daher vor allem nicht mehr die Kernlogik, sondern die fachliche Struktur der Ausgabe.

## Empfehlung

Als naechster Ausbauschritt sollte die Ergebnisverarbeitung so angepasst werden, dass nicht mehr nur `dtChanges`, sondern die vier fachlichen Ergebnisgruppen erzeugt und exportiert werden:

1. `dtNeuRegulaer`
2. `dtKonsolidiert`
3. `dtWegfall`
4. `dtManual`
