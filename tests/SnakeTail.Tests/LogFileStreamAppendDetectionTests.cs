using System;
using System.IO;
using System.Text;
using Xunit;

namespace SnakeTail.Tests;

public class LogFileStreamAppendDetectionTests
{
    [Fact]
    public void ReadLine_should_pick_up_appended_line_after_reaching_eof()
    {
        // 到达 EOF 后追加新行，下一次读取应能命中新内容。
        string tempDir = CreateTempDir();
        string filePath = Path.Combine(tempDir, "append.log");
        try
        {
            File.WriteAllText(filePath, "line-1" + Environment.NewLine, new UTF8Encoding(false));
            using var stream = new LogFileStream(string.Empty, filePath, new UTF8Encoding(false), fileCheckFrequency: 10, fileCheckPattern: false);

            Assert.Equal("line-1", stream.ReadLine(1));
            Assert.Null(stream.ReadLine(2));

            File.AppendAllText(filePath, "line-2" + Environment.NewLine, new UTF8Encoding(false));
            string appended = WaitReadLine(stream, 2, timeoutMs: 2000);

            Assert.Equal("line-2", appended);
        }
        finally
        {
            TryDeleteDir(tempDir);
        }
    }

    static string WaitReadLine(LogFileStream stream, int lineNumber, int timeoutMs)
    {
        DateTime deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        while (DateTime.UtcNow <= deadline)
        {
            string line = stream.ReadLine(lineNumber);
            if (line != null)
                return line;
            System.Threading.Thread.Sleep(40);
        }
        return null;
    }

    static string CreateTempDir()
    {
        string path = Path.Combine(Path.GetTempPath(), "snaketail-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    static void TryDeleteDir(string path)
    {
        try
        {
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // 测试收尾清理失败不影响断言结果。
        }
    }
}

