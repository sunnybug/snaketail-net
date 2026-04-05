using Xunit;

namespace SnakeTail.Tests;

public class LogFileStreamWatcherMatchTests
{
    [Fact]
    public void IsWatchEventMatch_exact_file_should_match_ignore_case()
    {
        // 精确文件模式：大小写差异不应影响匹配。
        bool match = LogFileStream.IsWatchEventMatch(
            configuredPathAbsolute: @"C:\logs\app.log",
            openedFilePath: @"C:\logs\app.log",
            eventPath: @"c:\LOGS\APP.LOG",
            fileCheckPattern: false);

        Assert.True(match);
    }

    [Fact]
    public void IsWatchEventMatch_exact_file_should_not_match_other_file()
    {
        // 精确文件模式：不同文件必须不匹配，避免误触发重载。
        bool match = LogFileStream.IsWatchEventMatch(
            configuredPathAbsolute: @"C:\logs\app.log",
            openedFilePath: @"C:\logs\app.log",
            eventPath: @"C:\logs\other.log",
            fileCheckPattern: false);

        Assert.False(match);
    }

    [Fact]
    public void IsWatchEventMatch_pattern_mode_should_accept_any_event_path()
    {
        // 通配模式：路径过滤由 watcher 的 filter 完成，这里统一放行。
        bool match = LogFileStream.IsWatchEventMatch(
            configuredPathAbsolute: @"C:\logs\*.log",
            openedFilePath: @"C:\logs\current.log",
            eventPath: @"C:\logs\next.log",
            fileCheckPattern: true);

        Assert.True(match);
    }
}
