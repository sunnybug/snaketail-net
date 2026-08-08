using System;
using System.IO;

namespace SnakeTail.Validation
{
    public static class MonitorRuleValidation
    {
        // Normalize to an absolute path for a directory string
        public static string NormalizeDirectoryPath(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir))
                throw new ArgumentException("Directory path cannot be null or whitespace.", nameof(dir));
            return Path.GetFullPath(dir);
        }

        // Validate that a directory exists on disk
        public static (bool isValid, string error) ValidateDirectory(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir))
                return (false, "Directory path is empty.");
            var full = NormalizeDirectoryPath(dir);
            if (Directory.Exists(full))
                return (true, string.Empty);
            return (false, $"Directory does not exist: {full}");
        }

        // Basic validation for wildcard patterns used to select files
        public static (bool isValid, string error) ValidateFilePattern(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern))
                return (false, "Pattern is empty.");
            // Disallow characters that are invalid in file names but commonly appear in invalid input
            foreach (char c in pattern)
            {
                if (c == '<' || c == '>' || c == '|' || c == '"' || c == ':')
                    return (false, $"Pattern contains invalid character: {c}");
            }
            return (true, string.Empty);
        }

        // Build a short, user-friendly preview path from directory + file pattern
        public static string BuildPreviewPath(string directoryPath, string filePattern)
        {
            var dir = NormalizeDirectoryPath(directoryPath);
            // Ensure no duplicate trailing slash
            if (dir.EndsWith(Path.DirectorySeparatorChar.ToString()))
                dir = dir.Substring(0, dir.Length - 1);
            return dir + Path.DirectorySeparatorChar + (filePattern ?? "");
        }

        // Display name that can be shown in the UI, with Disabled suffix when appropriate
        public static string BuildDisplayName(MonitorRuleConfig rule)
        {
            var name = rule?.Name ?? "MonitorRule";
            if (rule != null && !rule.Enabled)
                name += " (Disabled)";
            return name;
        }
    }
}
