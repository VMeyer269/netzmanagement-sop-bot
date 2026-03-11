using Netzmanagement.SopBot.InvokeCode.InvokeCode;

namespace Netzmanagement.SopBot.InvokeCode.Tests;

public sealed class MonthlyComparisonRequestTests
{
    [Fact]
    public void CreateDefaultRequest_SetsExpectedDefaults()
    {
        var request = SopMonatsvergleichDefinition.CreateDefaultRequest(
            "input/sop_2026_02.xlsx",
            "input/sop_2026_03.xlsx");

        Assert.Equal("SOP", request.WorksheetName);
        Assert.Equal(["BILANZKREIS", "FALLGRUPPE"], request.KeyColumns);
        Assert.True(request.HasValidPaths());
    }
}
