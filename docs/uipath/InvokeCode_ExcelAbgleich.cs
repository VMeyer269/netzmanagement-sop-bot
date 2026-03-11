// UiPath Invoke Code (Language: C#)
// In-Arguments: dtAlt (DataTable), dtNeu (DataTable)
// Out-Argument: dtChanges (DataTable)
//
// Required namespaces in UiPath:
// System
// System.Data
// System.Linq
// System.Collections.Generic

string keyCol = "BILANZKREIS";
string fallCol = "FALLGRUPPE";

// -------- Helpers --------

Action<DataTable> DropUnnamedColumns = (dt) =>
{
    if (dt == null) return;
    var remove = dt.Columns.Cast<DataColumn>()
        .Where(c => !string.IsNullOrWhiteSpace(c.ColumnName) && c.ColumnName.StartsWith("Unnamed", StringComparison.OrdinalIgnoreCase))
        .ToList();
    foreach (var c in remove) dt.Columns.Remove(c);
};

Func<object, string> GetTyp = (fall) =>
{
    var t = (fall ?? string.Empty).ToString().ToUpperInvariant();
    if (t.Contains("RLM")) return "RLM";
    if (t.Contains("SLP")) return "SLP";
    return "OTHER";
};

Func<object, string> GetRegime = (fall) =>
{
    var t = (fall ?? string.Empty).ToString().ToUpperInvariant();

    bool isTages = t.Contains("TAGES") || t.Contains("TAGE") || t.Contains("TAG ");
    bool isStunden = t.Contains("STUND") || t.Contains("STD");

    if (isTages && !isStunden) return "TAGES";
    if (isStunden && !isTages) return "STUNDEN";
    if (isTages && isStunden) return "MIXED";
    return "UNKNOWN";
};

Func<DataTable, DataTable> CloneWithBaseCols = (source) =>
{
    var dt = (source != null) ? source.Clone() : new DataTable();

    if (!dt.Columns.Contains("Status")) dt.Columns.Add("Status", typeof(string));
    if (!dt.Columns.Contains("Typ")) dt.Columns.Add("Typ", typeof(string));
    if (!dt.Columns.Contains("Regime_alt")) dt.Columns.Add("Regime_alt", typeof(string));
    if (!dt.Columns.Contains("Regime_neu")) dt.Columns.Add("Regime_neu", typeof(string));

    return dt;
};

Action<DataRow, DataRow> CopyCommonColumns = (src, dst) =>
{
    if (src == null || dst == null) return;
    foreach (DataColumn c in src.Table.Columns)
    {
        if (dst.Table.Columns.Contains(c.ColumnName))
        {
            dst[c.ColumnName] = src[c.ColumnName];
        }
    }
};

// result[bk][typ] = set(Regime)
Func<DataTable, Dictionary<string, Dictionary<string, HashSet<string>>>> BuildBkTypRegimeIndex = (dt) =>
{
    var result = new Dictionary<string, Dictionary<string, HashSet<string>>>(StringComparer.OrdinalIgnoreCase);
    if (dt == null || dt.Rows.Count == 0) return result;
    if (!dt.Columns.Contains(keyCol) || !dt.Columns.Contains(fallCol)) return result;

    foreach (DataRow r in dt.Rows)
    {
        var bk = (r[keyCol] ?? string.Empty).ToString();
        var typ = GetTyp(r[fallCol]);
        var reg = GetRegime(r[fallCol]);

        if (!result.ContainsKey(bk))
            result[bk] = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        if (!result[bk].ContainsKey(typ))
            result[bk][typ] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        result[bk][typ].Add(reg);
    }
    return result;
};

Func<HashSet<string>, string> ReduceRegime = (set) =>
{
    if (set == null || set.Count == 0) return null;
    if (set.Count == 1) return set.First();
    return "MIXED";
};

Func<DataTable, HashSet<string>, string, string, DataTable> FilterByBkAndTypWithStatus = (source, bkSet, typWanted, status) =>
{
    var outDt = CloneWithBaseCols(source);
    if (source == null || source.Rows.Count == 0) return outDt;
    if (bkSet == null || bkSet.Count == 0) return outDt;
    if (!source.Columns.Contains(keyCol) || !source.Columns.Contains(fallCol)) return outDt;

    foreach (DataRow r in source.AsEnumerable()
        .Where(x => bkSet.Contains((x[keyCol] ?? string.Empty).ToString()) && GetTyp(x[fallCol]) == typWanted))
    {
        var n = outDt.NewRow();
        CopyCommonColumns(r, n);
        n["Status"] = status;
        n["Typ"] = typWanted;
        outDt.Rows.Add(n);
    }
    return outDt;
};

// -------- Prepare --------

DropUnnamedColumns(dtAlt);
DropUnnamedColumns(dtNeu);

var idxAlt = BuildBkTypRegimeIndex(dtAlt);
var idxNeu = BuildBkTypRegimeIndex(dtNeu);

// Output init
dtChanges = CloneWithBaseCols(dtNeu ?? dtAlt);

// -------- 1) BK neu / weg --------

var bkNeu = new HashSet<string>(idxNeu.Keys.Except(idxAlt.Keys), StringComparer.OrdinalIgnoreCase);
var bkWeg = new HashSet<string>(idxAlt.Keys.Except(idxNeu.Keys), StringComparer.OrdinalIgnoreCase);

if (dtNeu != null && dtNeu.Columns.Contains(keyCol))
{
    foreach (DataRow r in dtNeu.AsEnumerable().Where(r => bkNeu.Contains((r[keyCol] ?? string.Empty).ToString())))
    {
        var n = dtChanges.NewRow();
        CopyCommonColumns(r, n);
        n["Status"] = "NEU";
        if (dtNeu.Columns.Contains(fallCol)) n["Typ"] = GetTyp(r[fallCol]);
        dtChanges.Rows.Add(n);
    }
}

if (dtAlt != null && dtAlt.Columns.Contains(keyCol))
{
    foreach (DataRow r in dtAlt.AsEnumerable().Where(r => bkWeg.Contains((r[keyCol] ?? string.Empty).ToString())))
    {
        var n = dtChanges.NewRow();
        CopyCommonColumns(r, n);
        n["Status"] = "WEG";
        if (dtAlt.Columns.Contains(fallCol)) n["Typ"] = GetTyp(r[fallCol]);
        dtChanges.Rows.Add(n);
    }
}

// -------- 2) RLM/SLP neu/weg innerhalb bestehender BKs --------

var commonBk = new HashSet<string>(idxAlt.Keys.Intersect(idxNeu.Keys), StringComparer.OrdinalIgnoreCase);

var rlmNeuBk = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var rlmWegBk = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var slpNeuBk = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var slpWegBk = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

foreach (var bk in commonBk)
{
    bool altRlm = idxAlt.ContainsKey(bk) && idxAlt[bk].ContainsKey("RLM");
    bool neuRlm = idxNeu.ContainsKey(bk) && idxNeu[bk].ContainsKey("RLM");
    if (!altRlm && neuRlm) rlmNeuBk.Add(bk);
    if (altRlm && !neuRlm) rlmWegBk.Add(bk);

    bool altSlp = idxAlt.ContainsKey(bk) && idxAlt[bk].ContainsKey("SLP");
    bool neuSlp = idxNeu.ContainsKey(bk) && idxNeu[bk].ContainsKey("SLP");
    if (!altSlp && neuSlp) slpNeuBk.Add(bk);
    if (altSlp && !neuSlp) slpWegBk.Add(bk);
}

var dtRlmNeu = FilterByBkAndTypWithStatus(dtNeu, rlmNeuBk, "RLM", "RLM_NEU");
var dtRlmWeg = FilterByBkAndTypWithStatus(dtAlt, rlmWegBk, "RLM", "RLM_WEG");
var dtSlpNeu = FilterByBkAndTypWithStatus(dtNeu, slpNeuBk, "SLP", "SLP_NEU");
var dtSlpWeg = FilterByBkAndTypWithStatus(dtAlt, slpWegBk, "SLP", "SLP_WEG");

foreach (DataRow r in dtRlmNeu.Rows) dtChanges.ImportRow(r);
foreach (DataRow r in dtRlmWeg.Rows) dtChanges.ImportRow(r);
foreach (DataRow r in dtSlpNeu.Rows) dtChanges.ImportRow(r);
foreach (DataRow r in dtSlpWeg.Rows) dtChanges.ImportRow(r);

// -------- 3) Wechsel Tages <-> Stundenregime (pro BK + Typ) --------

var dtRegimeSwitch = new DataTable();
dtRegimeSwitch.Columns.Add(keyCol, typeof(string));
dtRegimeSwitch.Columns.Add("Typ", typeof(string));
dtRegimeSwitch.Columns.Add("Regime_alt", typeof(string));
dtRegimeSwitch.Columns.Add("Regime_neu", typeof(string));
dtRegimeSwitch.Columns.Add("Status", typeof(string));

foreach (var bk in commonBk)
{
    foreach (var typ in new[] { "RLM", "SLP" })
    {
        bool hasAlt = idxAlt.ContainsKey(bk) && idxAlt[bk].ContainsKey(typ);
        bool hasNeu = idxNeu.ContainsKey(bk) && idxNeu[bk].ContainsKey(typ);
        if (!hasAlt || !hasNeu) continue;

        var altReg = ReduceRegime(idxAlt[bk][typ]);
        var neuReg = ReduceRegime(idxNeu[bk][typ]);

        bool altOk = (altReg == "TAGES" || altReg == "STUNDEN");
        bool neuOk = (neuReg == "TAGES" || neuReg == "STUNDEN");

        if (altOk && neuOk && altReg != neuReg)
        {
            var row = dtRegimeSwitch.NewRow();
            row[keyCol] = bk;
            row["Typ"] = typ;
            row["Regime_alt"] = altReg;
            row["Regime_neu"] = neuReg;
            row["Status"] = "WECHSEL_TAGES_STUNDEN";
            dtRegimeSwitch.Rows.Add(row);
        }
    }
}

foreach (DataRow r in dtRegimeSwitch.Rows)
{
    var n = dtChanges.NewRow();
    n[keyCol] = r[keyCol];
    n["Typ"] = r["Typ"];
    n["Regime_alt"] = r["Regime_alt"];
    n["Regime_neu"] = r["Regime_neu"];
    n["Status"] = r["Status"];
    dtChanges.Rows.Add(n);
}

// -------- Optional: Duplikate entfernen (gleicher BK + Typ + Status + Regime) --------

var deduped = dtChanges.AsEnumerable()
    .GroupBy(r => string.Join("|",
        (r.Table.Columns.Contains(keyCol) ? (r[keyCol] ?? string.Empty).ToString() : string.Empty),
        (r.Table.Columns.Contains("Typ") ? (r["Typ"] ?? string.Empty).ToString() : string.Empty),
        (r.Table.Columns.Contains("Status") ? (r["Status"] ?? string.Empty).ToString() : string.Empty),
        (r.Table.Columns.Contains("Regime_alt") ? (r["Regime_alt"] ?? string.Empty).ToString() : string.Empty),
        (r.Table.Columns.Contains("Regime_neu") ? (r["Regime_neu"] ?? string.Empty).ToString() : string.Empty)))
    .Select(g => g.First())
    .ToList();

var dedupTable = dtChanges.Clone();
foreach (var row in deduped) dedupTable.ImportRow(row);
dtChanges = dedupTable;

// -------- Sortieren --------

if (dtChanges.Rows.Count > 0 && dtChanges.Columns.Contains(keyCol))
{
    var sorted = dtChanges.AsEnumerable()
        .OrderBy(r => (r[keyCol] ?? string.Empty).ToString())
        .ThenBy(r => (r["Status"] ?? string.Empty).ToString())
        .ThenBy(r => (r["Typ"] ?? string.Empty).ToString())
        .ToList();

    var tmp = dtChanges.Clone();
    foreach (var r in sorted) tmp.ImportRow(r);
    dtChanges = tmp;
}
