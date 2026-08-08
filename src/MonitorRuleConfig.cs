using System;

namespace SnakeTail.Validation
{
    // Stable data shape for monitor rule configuration
    public class MonitorRuleConfig
    {
        public string Name { get; set; }
        public string DirectoryPath { get; set; }
        public string FilePattern { get; set; }
        public bool Enabled { get; set; }
    }
}
