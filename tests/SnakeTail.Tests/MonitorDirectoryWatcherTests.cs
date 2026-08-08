using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace SnakeTail.Tests
{
    public class MonitorDirectoryWatcherTests
    {
        [Fact]
        public async Task Baseline_discovery_and_tail_on_create()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "SnakeTail_WildcardWatcher_Test_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var watcher = new MonitorDirectoryWatcher(tempDir, "*err*.log");
                string observed = null;
                var tcs = new TaskCompletionSource<string>();
                watcher.FileMatched += (p) => { observed = p; tcs.TrySetResult(p); };
                watcher.StartBaseline();

                // Create a matching file after baseline
                var path = Path.Combine(tempDir, "myerrfile.log");
                File.WriteAllText(path, "hello");
                // Wait for event
                var completed = await Task.WhenAny(tcs.Task, Task.Delay(5000));
                Assert.True(completed == tcs.Task, "Did not observe matching file in time");
                Assert.Equal(path, tcs.Task.Result);
                watcher.Stop();
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }
    }
}
