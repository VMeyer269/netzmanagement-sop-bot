// UiPath Invoke Code (Language: C#)
// In-Arguments:
//   dtAlt     (System.Data.DataTable)  -> Vormonat
//   dtNeu     (System.Data.DataTable)  -> aktueller Monat
// Out-Arguments:
//   dtChanges (System.Data.DataTable)  -> bereinigte Aenderungen
//
// Required namespaces in UiPath:
// System
// System.Data
// System.Linq
// System.Collections.Generic
// System.Globalization

var deCulture = new System.Globalization.CultureInfo("de-DE");

Func<object, string> AsText = value =>
{
    return (value ?? string.Empty).ToString().Trim();
};

Func<object, string> NormalizeKey = value =>
{
    return AsText(value).ToUpperInvariant();
};

Func<DataTable, string, string> GetColumnName = (table, expectedName) =>
{
    var match = table.Columns
        .Cast<DataColumn>()
        .FirstOrDefault(c => string.Equals(
            c.ColumnName?.Trim(),
            expectedName,
            StringComparison.OrdinalIgnoreCase));

    if (match == null)
    {
        throw new Exception("Pflichtspalte nicht gefunden: " + expectedName);
    }

    return match.ColumnName;
};

Action<DataTable> CleanupTable = table =>
{
    if (table == null) return;

    var unnamedColumns = table.Columns
        .Cast<DataColumn>()
        .Where(c =>
            string.IsNullOrWhiteSpace(c.ColumnName) ||
            c.ColumnName.StartsWith("Unnamed", StringComparison.OrdinalIgnoreCase))
        .ToList();

    foreach (var column in unnamedColumns)
    {
        table.Columns.Remove(column);
    }

    var rowsToDelete = table.AsEnumerable()
        .Where(row =>
        {
            bool hasValues = row.ItemArray.Any(cell => !string.IsNullOrWhiteSpace(AsText(cell)));
            if (!hasValues) return true;

            var firstValue = row.ItemArray.Length > 0 ? AsText(row.ItemArray[0]) : string.Empty;
            return firstValue.StartsWith("EXPORT aus ", StringComparison.OrdinalIgnoreCase);
        })
        .ToList();

    foreach (var row in rowsToDelete)
    {
        table.Rows.Remove(row);
    }
};

Func<object, decimal> ParseDecimal = value =>
{
    var text = AsText(value);
    if (string.IsNullOrWhiteSpace(text)) return 0m;

    decimal parsed;
    if (decimal.TryParse(text, System.Globalization.NumberStyles.Any, deCulture, out parsed)) return parsed;
    if (decimal.TryParse(text.Replace(".", string.Empty).Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return parsed;
    if (decimal.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return parsed;

    return 0m;
};

Func<decimal, string> FormatDecimal = value =>
{
    return value.ToString("0.####", deCulture);
};

Func<object, DateTime?> ParseDate = value =>
{
    var text = AsText(value);
    if (string.IsNullOrWhiteSpace(text)) return null;

    DateTime parsed;
    if (DateTime.TryParse(text, deCulture, System.Globalization.DateTimeStyles.None, out parsed)) return parsed;
    if (DateTime.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out parsed)) return parsed;

    return null;
};

Func<DateTime?, string> FormatDate = value =>
{
    return value.HasValue ? value.Value.ToString("dd.MM.yyyy") : string.Empty;
};

Func<DataTable> CreateRawTable = () =>
{
    var raw = new DataTable("RawChanges");

    foreach (DataColumn column in dtNeu.Columns)
    {
        if (!raw.Columns.Contains(column.ColumnName))
        {
            raw.Columns.Add(column.ColumnName, typeof(string));
        }
    }

    foreach (DataColumn column in dtAlt.Columns)
    {
        if (!raw.Columns.Contains(column.ColumnName))
        {
            raw.Columns.Add(column.ColumnName, typeof(string));
        }
    }

    raw.Columns.Add("Status", typeof(string));
    raw.Columns.Add("ManuellePruefung", typeof(string));
    raw.Columns.Add("Hinweis", typeof(string));
    raw.Columns.Add("AnzahlKonsolidierterZeilen", typeof(string));
    raw.Columns.Add("__AnlagenKey", typeof(string));
    raw.Columns.Add("__ZaehlpunktKey", typeof(string));

    return raw;
};

Func<DataTable, DataTable> CreateFinalTable = raw =>
{
    var finalTable = raw.Clone();
    finalTable.Columns.Remove("__AnlagenKey");
    finalTable.Columns.Remove("__ZaehlpunktKey");
    finalTable.TableName = "Changes";
    return finalTable;
};

Func<DataTable, DataRow, string, string, string, DataRow> BuildRawRow = (targetTable, sourceRow, status, anlagenCol, zaehlpunktCol) =>
{
    var row = targetTable.NewRow();

    foreach (DataColumn column in sourceRow.Table.Columns)
    {
        if (targetTable.Columns.Contains(column.ColumnName))
        {
            row[column.ColumnName] = AsText(sourceRow[column]);
        }
    }

    row["Status"] = status;
    row["ManuellePruefung"] = "NEIN";
    row["Hinweis"] = string.Empty;
    row["AnzahlKonsolidierterZeilen"] = "1";
    row["__AnlagenKey"] = NormalizeKey(sourceRow[anlagenCol]);
    row["__ZaehlpunktKey"] = NormalizeKey(sourceRow[zaehlpunktCol]);

    if (targetTable.Columns.Contains("Prognosesmenge"))
    {
        row["Prognosesmenge"] = FormatDecimal(ParseDecimal(sourceRow["Prognosesmenge"]));
    }

    return row;
};

Func<IEnumerable<DataRow>, string, string, IEnumerable<DataRow>> SortRows = (rows, dateFromCol, dateToCol) =>
{
    return rows
        .OrderBy(r => ParseDate(r[dateFromCol]) ?? DateTime.MinValue)
        .ThenBy(r => ParseDate(r[dateToCol]) ?? DateTime.MinValue)
        .ThenBy(r => AsText(r["MaLo-ID"]))
        .ThenBy(r => AsText(r["Vertrags-Nummer"]));
};

CleanupTable(dtAlt);
CleanupTable(dtNeu);

string anlagenColAlt = GetColumnName(dtAlt, "Anlagen-Nummer");
string anlagenColNeu = GetColumnName(dtNeu, "Anlagen-Nummer");
string zaehlpunktColAlt = GetColumnName(dtAlt, "Zählpunktbezeichnung");
string zaehlpunktColNeu = GetColumnName(dtNeu, "Zählpunktbezeichnung");
string prognoseColAlt = GetColumnName(dtAlt, "Prognosesmenge");
string prognoseColNeu = GetColumnName(dtNeu, "Prognosesmenge");
string vonColAlt = GetColumnName(dtAlt, "Prognoses-ab");
string vonColNeu = GetColumnName(dtNeu, "Prognoses-ab");
string bisColAlt = GetColumnName(dtAlt, "Prognoses-bis");
string bisColNeu = GetColumnName(dtNeu, "Prognoses-bis");

var rawChanges = CreateRawTable();

var altByAnlage = dtAlt.AsEnumerable()
    .GroupBy(row => NormalizeKey(row[anlagenColAlt]))
    .ToDictionary(group => group.Key, group => group.ToList());

var neuByAnlage = dtNeu.AsEnumerable()
    .GroupBy(row => NormalizeKey(row[anlagenColNeu]))
    .ToDictionary(group => group.Key, group => group.ToList());

var allAnlagen = altByAnlage.Keys
    .Union(neuByAnlage.Keys)
    .Where(key => !string.IsNullOrWhiteSpace(key))
    .OrderBy(key => key)
    .ToList();

foreach (var anlagenKey in allAnlagen)
{
    var altRows = altByAnlage.ContainsKey(anlagenKey)
        ? altByAnlage[anlagenKey]
        : new List<DataRow>();

    var neuRows = neuByAnlage.ContainsKey(anlagenKey)
        ? neuByAnlage[anlagenKey]
        : new List<DataRow>();

    var altByZaehlpunkt = altRows
        .GroupBy(row => NormalizeKey(row[zaehlpunktColAlt]))
        .ToDictionary(group => group.Key, group => group.ToList());

    var neuByZaehlpunkt = neuRows
        .GroupBy(row => NormalizeKey(row[zaehlpunktColNeu]))
        .ToDictionary(group => group.Key, group => group.ToList());

    var allZaehlpunkte = altByZaehlpunkt.Keys
        .Union(neuByZaehlpunkt.Keys)
        .OrderBy(key => key)
        .ToList();

    foreach (var zaehlpunktKey in allZaehlpunkte)
    {
        var altRowsForZp = altByZaehlpunkt.ContainsKey(zaehlpunktKey)
            ? SortRows(altByZaehlpunkt[zaehlpunktKey], vonColAlt, bisColAlt).ToList()
            : new List<DataRow>();

        var neuRowsForZp = neuByZaehlpunkt.ContainsKey(zaehlpunktKey)
            ? SortRows(neuByZaehlpunkt[zaehlpunktKey], vonColNeu, bisColNeu).ToList()
            : new List<DataRow>();

        int altCount = altRowsForZp.Count;
        int neuCount = neuRowsForZp.Count;
        int sharedCount = System.Math.Min(altCount, neuCount);

        foreach (var extraAlt in altRowsForZp.Skip(sharedCount))
        {
            rawChanges.Rows.Add(BuildRawRow(rawChanges, extraAlt, "WEG", anlagenColAlt, zaehlpunktColAlt));
        }

        foreach (var extraNeu in neuRowsForZp.Skip(sharedCount))
        {
            rawChanges.Rows.Add(BuildRawRow(rawChanges, extraNeu, "NEU", anlagenColNeu, zaehlpunktColNeu));
        }
    }
}

var processedRaw = rawChanges.Clone();

foreach (var anlagenGroup in rawChanges.AsEnumerable().GroupBy(row => AsText(row["__AnlagenKey"])))
{
    var groupRows = anlagenGroup.ToList();
    var distinctZaehlpunkte = groupRows
        .Select(row => AsText(row["__ZaehlpunktKey"]))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    if (groupRows.Count == 1)
    {
        processedRaw.ImportRow(groupRows[0]);
        continue;
    }

    if (distinctZaehlpunkte.Count > 1)
    {
        foreach (var row in groupRows)
        {
            var manualRow = processedRaw.NewRow();
            manualRow.ItemArray = row.ItemArray.Clone() as object[];
            manualRow["ManuellePruefung"] = "JA";
            manualRow["Hinweis"] = "Mehrere Aenderungen zur selben Anlagen-Nummer mit unterschiedlicher Zaehlpunktbezeichnung. Bitte manuell pruefen.";
            processedRaw.Rows.Add(manualRow);
        }

        continue;
    }

    foreach (var statusGroup in groupRows.GroupBy(row => AsText(row["Status"])))
    {
        var sameStatusRows = statusGroup.ToList();

        if (sameStatusRows.Count == 1)
        {
            processedRaw.ImportRow(sameStatusRows[0]);
            continue;
        }

        var consolidatedRow = processedRaw.NewRow();
        consolidatedRow.ItemArray = sameStatusRows[0].ItemArray.Clone() as object[];

        decimal prognoseSum = sameStatusRows.Sum(row =>
        {
            string prognoseCol = row.Table.Columns.Contains("Prognosesmenge")
                ? "Prognosesmenge"
                : prognoseColNeu;
            return ParseDecimal(row[prognoseCol]);
        });

        var fromDates = sameStatusRows
            .Select(row => ParseDate(row["Prognoses-ab"]))
            .Where(date => date.HasValue)
            .Select(date => date.Value)
            .ToList();

        var toDates = sameStatusRows
            .Select(row => ParseDate(row["Prognoses-bis"]))
            .Where(date => date.HasValue)
            .Select(date => date.Value)
            .ToList();

        consolidatedRow["Prognosesmenge"] = FormatDecimal(prognoseSum);
        consolidatedRow["AnzahlKonsolidierterZeilen"] = sameStatusRows.Count.ToString();
        consolidatedRow["Hinweis"] = "Mehrere Aenderungen mit identischer Zaehlpunktbezeichnung wurden zusammengefuehrt.";

        if (fromDates.Count > 0)
        {
            consolidatedRow["Prognoses-ab"] = FormatDate(fromDates.Min());
        }

        if (toDates.Count > 0)
        {
            consolidatedRow["Prognoses-bis"] = FormatDate(toDates.Max());
        }

        processedRaw.Rows.Add(consolidatedRow);
    }
}

dtChanges = CreateFinalTable(processedRaw);

var sortedChanges = processedRaw.AsEnumerable()
    .OrderByDescending(row => AsText(row["ManuellePruefung"]) == "JA")
    .ThenBy(row => AsText(row["Anlagen-Nummer"]))
    .ThenBy(row => AsText(row["Status"]))
    .ThenBy(row => AsText(row["Zählpunktbezeichnung"]))
    .ThenBy(row => ParseDate(row["Prognoses-ab"]) ?? DateTime.MinValue)
    .ToList();

foreach (var row in sortedChanges)
{
    var outputRow = dtChanges.NewRow();

    foreach (DataColumn column in dtChanges.Columns)
    {
        outputRow[column.ColumnName] = row[column.ColumnName];
    }

    dtChanges.Rows.Add(outputRow);
}
