using Xunit;

namespace SnakeTail.Tests;

public class AppLogTests
{
    [Fact]
    public void FormatLine_uses_skill_layout()
    {
        string line = AppLog.FormatLine("INFO", @"C:\proj\Program.cs", 42, "hello");
        Assert.Contains("[INFO]", line);
        Assert.Contains("[Program.cs:42]", line);
        Assert.Contains("[hello]", line);
    }
}
