using System;
using SnakeTail.Validation;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace SnakeTail.Tests
{
    public class MonitorRuleManagerTests
    {
        [Fact]
        public async Task Manager_aggregates_rules_and_tracks_matches()
        {
            var dir = Path.Combine(Path.GetTempPath(), "SnakeTail_Manager_Test_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var rules = new[] {
                    new MonitorRuleConfig { Name = "A", DirectoryPath = dir, FilePattern = "*a*.log", Enabled = true },
                    new MonitorRuleConfig { Name = "B", DirectoryPath = dir, FilePattern = "*b*.log", Enabled = true }
                };
                var manager = new MonitorRuleManager(rules);
                string observed = null;
                var tcs = new TaskCompletionSource<string>();
                manager.FileTailed += (p) => { observed = p; tcs.TrySetResult(p); };

                // Create a file that matches A (which is enabled)
                var pathA = Path.Combine(dir, "fileA.log");
                File.WriteAllText(pathA, "x");
                var completed = await Task.WhenAny(tcs.Task, Task.Delay(5000));
                Assert.True(completed == tcs.Task, "Did not observe tailed file from rule A");
                Assert.Equal(pathA, tcs.Task.Result);
                manager.Dispose();
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }
    }
}
