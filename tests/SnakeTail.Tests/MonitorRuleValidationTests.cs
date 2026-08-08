using System;
using System.IO;
using SnakeTail.Validation;
using Xunit;

namespace SnakeTail.Tests
{
    public class MonitorRuleValidationTests
    {
        [Fact]
        public void NormalizeDirectoryPath_ReturnsAbsolutePath()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "snaketail_monitor_norm_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            string result = MonitorRuleValidation.NormalizeDirectoryPath(tempDir);
            Assert.Equal(Path.GetFullPath(tempDir).TrimEnd(Path.DirectorySeparatorChar), result.TrimEnd(Path.DirectorySeparatorChar));
        }

        [Fact]
        public void ValidateDirectory_Existing_ReturnsTrue()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "snaketail_monitor_dir_exists_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            var (isValid, error) = MonitorRuleValidation.ValidateDirectory(tempDir);
            Assert.True(isValid);
            Assert.True(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void ValidateDirectory_NonExisting_ReturnsFalse()
        {
            string nonExist = Path.Combine(Path.GetTempPath(), "nonexistent_dir_" + Guid.NewGuid().ToString("N"));
            var (isValid, error) = MonitorRuleValidation.ValidateDirectory(nonExist);
            Assert.False(isValid);
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void ValidateFilePattern_Valid_ReturnsTrue()
        {
            var (isValid, error) = MonitorRuleValidation.ValidateFilePattern("*.txt");
            Assert.True(isValid);
            Assert.True(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void ValidateFilePattern_Invalid_ReturnsFalse()
        {
            var (isValid, error) = MonitorRuleValidation.ValidateFilePattern("invalid|name.txt");
            Assert.False(isValid);
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void BuildPreviewPath_ReturnsExpected()
        {
            string dir = Path.Combine(Path.GetTempPath(), "snaketail_monitor_preview_dir_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            string pattern = "*.log";
            string preview = MonitorRuleValidation.BuildPreviewPath(dir, pattern);
            string expected = Path.Combine(Path.GetFullPath(dir), pattern);
            Assert.Equal(expected, preview);
        }

        [Fact]
        public void BuildDisplayName_ReturnsNameWhenEnabled()
        {
            var rule = new MonitorRuleConfig { Name = "MyRule", Enabled = true };
            string display = MonitorRuleValidation.BuildDisplayName(rule);
            Assert.Equal("MyRule", display);
        }

        [Fact]
        public void BuildDisplayName_ReturnsDisabledSuffixWhenDisabled()
        {
            var rule = new MonitorRuleConfig { Name = "MyRule", Enabled = false };
            string display = MonitorRuleValidation.BuildDisplayName(rule);
            Assert.Equal("MyRule (Disabled)", display);
        }
    }
}
