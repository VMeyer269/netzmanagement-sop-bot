# Prozessautomatisierung Netzmanagement

Dieses Repository bildet die technische Basis fuer UiPath-Invoke-Code-Bausteine in C# zur Filterung und Aufbereitung von SOP-Daten im Netzmanagement.

Der erste fachliche Anwendungsfall ist der Vergleich von Excel-Listen des Vormonats und des aktuellen Monats, um Aenderungen strukturiert fuer den Bot weiterzuverarbeiten.

## Zielbild

- C# als gemeinsame Sprache fuer Invoke-Code-Bausteine
- klare Trennung zwischen fachlicher Logik, UiPath-Integration und Tests
- versionierbare Snippets und Hilfsklassen fuer Monatsvergleiche
- MIT-Lizenz fuer das gesamte Repository

## Repository-Struktur

```text
.
|-- process-automation-engine.sln
|-- Directory.Build.props
|-- src/
|   `-- Netzmanagement.SopBot.InvokeCode/
|       |-- InvokeCode/
|       |   `-- SopMonatsvergleichDefinition.cs
|       |-- Models/
|       |   `-- MonthlyComparisonRequest.cs
|       `-- Netzmanagement.SopBot.InvokeCode.csproj
|-- tests/
|   `-- Netzmanagement.SopBot.InvokeCode.Tests/
|       |-- MonthlyComparisonRequestTests.cs
|       `-- Netzmanagement.SopBot.InvokeCode.Tests.csproj
`-- docs/
    `-- uipath/
        |-- InvokeCode_ExcelAbgleich.cs
        `-- README_UiPath_ExcelAbgleich.md
```

## Geplanter Scope

1. Einlesen und Vereinheitlichen der Excel-Daten aus Vor- und Aktuell-Monat
2. Vergleich der Datenbestaende auf fachlich relevanten Schluesseln
3. Aufbereitung der Differenzen fuer nachgelagerte UiPath-Schritte
4. Kapselung wiederverwendbarer Regeln in testbaren C#-Bausteinen

## Lokale Voraussetzungen

- .NET SDK 8.0 oder neuer
- UiPath Studio mit `Invoke Code`-Aktivitaet fuer die spaetere Integration

## Naechste Schritte

1. Domaenenmodell fuer SOP-Datensaetze finalisieren
2. Vergleichsregeln fuer Monatsdifferenzen implementieren
3. Invoke-Code-Snippets aus den Regeln ableiten und in UiPath einbinden
