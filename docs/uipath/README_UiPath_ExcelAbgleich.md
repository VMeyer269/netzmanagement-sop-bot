# UiPath Excel-Abgleich (Invoke Code, C#)

Vergleicht zwei Monatsstaende (`dtAlt`, `dtNeu`) und erzeugt eine Aenderungstabelle `dtChanges` fuer:

- neue/weggefallene Bilanzkreise (`NEU`, `WEG`)
- neue/weggefallene Typen innerhalb bestehender Bilanzkreise (`RLM_NEU`, `RLM_WEG`, `SLP_NEU`, `SLP_WEG`)
- Regimewechsel zwischen `TAGES` und `STUNDEN` (`WECHSEL_TAGES_STUNDEN`)

## Datei

- Code: `InvokeCode_ExcelAbgleich.cs`

## UiPath Setup

1. Activity: `Invoke Code`
2. Language: `CSharp`
3. Namespaces:
   - `System`
   - `System.Data`
   - `System.Linq`
   - `System.Collections.Generic`
4. Arguments:
   - `dtAlt` (`System.Data.DataTable`, `In`)
   - `dtNeu` (`System.Data.DataTable`, `In`)
   - `dtChanges` (`System.Data.DataTable`, `Out`)

## Erwartete Input-Spalten

- `BILANZKREIS`
- `FALLGRUPPE`

Hinweis: Spaltennamen, die mit `Unnamed` beginnen, werden automatisch entfernt.

## Output-Spalten (mindestens)

- alle Spalten aus `dtAlt`/`dtNeu` (je nach Clone-Basis)
- `Status`
- `Typ`
- `Regime_alt`
- `Regime_neu`

## Statuswerte

- `NEU`: Bilanzkreis existiert nur in `dtNeu`
- `WEG`: Bilanzkreis existiert nur in `dtAlt`
- `RLM_NEU` / `RLM_WEG`: Typ RLM ist innerhalb eines bestehenden Bilanzkreises neu/weg
- `SLP_NEU` / `SLP_WEG`: Typ SLP ist innerhalb eines bestehenden Bilanzkreises neu/weg
- `WECHSEL_TAGES_STUNDEN`: Typ vorhanden in beiden Monaten, Regimewechsel Tages <-> Stunden erkannt

## Wichtige technische Entscheidungen

- Robustes Zeilenkopieren: kein direktes `ItemArray`-Setzen, damit keine Fehler bei zusaetzlichen Spalten entstehen.
- Case-insensitive Typ-Erkennung (`RLM`/`SLP`).
- Deduplizierung am Ende auf Basis `BILANZKREIS + Typ + Status + Regime_alt + Regime_neu`.
- Sortierung nach `BILANZKREIS`, `Status`, `Typ`.

## Bekannte Grenzen

- Regime-Erkennung basiert auf Textmustern in `FALLGRUPPE`.
- Bei uneindeutigen Texten wird `UNKNOWN` oder `MIXED` gesetzt.
- Wechsel werden nur fuer eindeutige Wechsel `TAGES` <-> `STUNDEN` gemeldet.

## Beispiel fuer Repo-README (Kurztext)

```md
## Excel-Abgleich in UiPath (Invoke Code, C#)

Dieses Snippet vergleicht zwei Monatsdaten (`dtAlt`, `dtNeu`) und erzeugt eine Aenderungstabelle mit:
- NEU/WEG auf Bilanzkreis-Ebene
- RLM/SLP NEU/WEG innerhalb bestehender Bilanzkreise
- Regimewechsel Tages <-> Stunden

Technisch umgesetzt als `Invoke Code` in C# mit DataTable-Indexierung und robustem Spaltenhandling.
```

## Lizenz / Nutzung

Vor dem Veröffentlichen bitte ggf. interne Begriffe, Kundendaten und Dateinamen anonymisieren.
