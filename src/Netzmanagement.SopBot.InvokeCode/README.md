# Modul: Netzmanagement.SopBot.InvokeCode

Dieses Projekt nimmt die wiederverwendbare C#-Logik fuer UiPath-`Invoke Code`-Bausteine auf.

## Startpunkt

- `Models/MonthlyComparisonRequest.cs` definiert den Eingabevertrag fuer Monatsvergleiche.
- `InvokeCode/SopMonatsvergleichDefinition.cs` kapselt die Standardkonfiguration fuer den SOP-Abgleich.

Die eigentliche Vergleichslogik kann anschliessend als testbare Services oder direkt als UiPath-Snippet ergaenzt werden.
