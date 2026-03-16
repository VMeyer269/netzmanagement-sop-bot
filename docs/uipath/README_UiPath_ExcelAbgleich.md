# UiPath SOP-Monatsvergleich (Invoke Code, C#)

Dieses Snippet vergleicht die SOP-Daten des Vormonats mit dem aktuellen Monat auf Ebene der `Anlagen-Nummer` und bereitet die Aenderungen fuer den weiteren UiPath-Flow auf.

## Abgedeckte Logik

1. Vergleich Vor- vs. Aktuell-Monat auf Basis der `Anlagen-Nummer`
2. Ermittlung der zusaetzlichen bzw. weggefallenen Eintraege
3. Nachpruefung der Aenderungstabelle auf Wiederholungen
4. Konsolidierung von Wiederholungen mit gleicher `Zaehlpunktbezeichnung`
5. Summierung der `Prognosesmenge` bei konsolidierten Wiederholungen
6. Kennzeichnung von Faellen mit unterschiedlicher `Zaehlpunktbezeichnung` fuer manuelle Bearbeitung

## UiPath Setup

Activity:
- `Invoke Code`

Language:
- `CSharp`

Namespaces:
- `System`
- `System.Data`
- `System.Linq`
- `System.Collections.Generic`
- `System.Globalization`

Arguments:
- `dtAlt` (`System.Data.DataTable`, `In`)
- `dtNeu` (`System.Data.DataTable`, `In`)
- `dtChanges` (`System.Data.DataTable`, `Out`)

## Erwartete Input-Spalten

- `Anlagen-Nummer`
- `Zählpunktbezeichnung`
- `Prognosesmenge`
- `Prognoses-ab`
- `Prognoses-bis`

Weitere Spalten werden in die Ausgabe uebernommen, sofern sie in den Input-Tabellen vorhanden sind.

## Output

Die Ausgabe `dtChanges` enthaelt die originalen Datenspalten plus diese Zusatzspalten:

- `Status`
  - `NEU` fuer zusaetzliche Eintraege im aktuellen Monat
  - `WEG` fuer Eintraege, die im aktuellen Monat fehlen
- `ErgebnisKategorie`
  - `NEUE_ANLAGE` fuer regulaere Neuzugaenge
  - `ZUSAMMENGEFUEHRT` fuer zusammengefuehrte Dubletten
  - `WEGFALL` fuer im aktuellen Monat fehlende Eintraege
  - `MANUELLE_PRUEFUNG` fuer Faelle mit manueller Nachbearbeitung
- `ManuellePruefung`
  - `JA`, wenn fuer dieselbe `Anlagen-Nummer` unterschiedliche `Zaehlpunktbezeichnungen` in der Aenderungstabelle vorkommen
  - `NEIN` in allen anderen Faellen
- `Hinweis`
  - Erlaeuterung zur Konsolidierung oder zur manuellen Pruefung
- `AnzahlKonsolidierterZeilen`
  - Anzahl der urspruenglichen Aenderungszeilen, die in einer Ausgabezeile aufgegangen sind

## Fachliche Wirkung

- Wiederholungen mit gleicher `Zaehlpunktbezeichnung` werden zu einer Zeile je Status zusammengefuehrt.
- Dabei wird die `Prognosesmenge` aufsummiert.
- `Prognoses-ab` wird auf das frueheste Datum gesetzt.
- `Prognoses-bis` wird auf das spaeteste Datum gesetzt.
- Wiederholungen mit unterschiedlicher `Zaehlpunktbezeichnung` bleiben getrennt und werden mit `ManuellePruefung = JA` markiert.
- Die fachliche Einordnung steht zusaetzlich direkt in `ErgebnisKategorie` und kann im Workflow fuer die Ergebnisaufteilung verwendet werden.
- Konsolidierungen werden auch dann erzeugt, wenn eine Anlage bereits im Vormonat vorhanden war, der aktuelle Monat aber mehrere Teilzeilen mit derselben `Zaehlpunktbezeichnung` enthaelt.

## Hinweis zu den Beispieldateien

Mit den Beispieldateien Januar 2026 und Februar 2026 entsteht mindestens ein Konsolidierungsfall, bei dem zwei `NEU`-Zeilen fuer dieselbe `Anlagen-Nummer` und dieselbe `Zaehlpunktbezeichnung` zu einer Zeile zusammengefuehrt werden.
