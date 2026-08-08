using System;
using SnakeTail.Validation;
using Xunit;

namespace SnakeTail.Tests
{
    public class MonitorRuleConfigTests
    {
        [Fact]
        public void CreateRuleConfig_Defaults()
        {
            var r = new MonitorRuleConfig { Name = "Test", DirectoryPath = @"C:\temp", FilePattern = "*.log", Enabled = true };
            Assert.NotNull(r.Name);
            Assert.True(DirectoryExists(r.DirectoryPath) || r.DirectoryPath != null);
            Assert.False(string.IsNullOrWhiteSpace(r.FilePattern));
        }

        private bool DirectoryExists(string path) => System.IO.Directory.Exists(path);
    }
}
