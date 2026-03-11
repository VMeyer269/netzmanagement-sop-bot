using Netzmanagement.SopBot.InvokeCode.Models;

namespace Netzmanagement.SopBot.InvokeCode.InvokeCode;

public static class SopMonatsvergleichDefinition
{
    public static readonly IReadOnlyList<string> DefaultKeyColumns =
    [
        "BILANZKREIS",
        "FALLGRUPPE"
    ];

    public static MonthlyComparisonRequest CreateDefaultRequest(
        string previousMonthPath,
        string currentMonthPath,
        string worksheetName = "SOP")
    {
        return new MonthlyComparisonRequest(
            previousMonthPath,
            currentMonthPath,
            worksheetName,
            DefaultKeyColumns);
    }
}
