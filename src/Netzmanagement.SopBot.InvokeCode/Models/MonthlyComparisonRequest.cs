namespace Netzmanagement.SopBot.InvokeCode.Models;

public sealed record MonthlyComparisonRequest(
    string PreviousMonthPath,
    string CurrentMonthPath,
    string WorksheetName,
    IReadOnlyList<string> KeyColumns)
{
    public bool HasValidPaths() =>
        !string.IsNullOrWhiteSpace(PreviousMonthPath) &&
        !string.IsNullOrWhiteSpace(CurrentMonthPath);
}
