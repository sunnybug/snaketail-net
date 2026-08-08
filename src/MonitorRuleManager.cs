using System;
using System.Collections.Generic;
using System.Linq;
using SnakeTail.Validation;

namespace SnakeTail
{
    // Coordinates multiple MonitorDirectoryWatcher instances based on MonitorRuleConfig
    public class MonitorRuleManager : IDisposable
    {
        private readonly List<MonitorRuleConfig> _rules;
        private readonly Dictionary<string, MonitorDirectoryWatcher> _watchers = new Dictionary<string, MonitorDirectoryWatcher>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _tailDedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _disposed;

        public event Action<string> FileTailed;

        public MonitorRuleManager(IEnumerable<MonitorRuleConfig> rules)
        {
            _rules = new List<MonitorRuleConfig>(rules ?? Enumerable.Empty<MonitorRuleConfig>());
            // Start any initially enabled rules
            foreach (var r in _rules.Where(x => x.Enabled))
            {
                StartWatcherForRule(r);
            }
        }

        private void StartWatcherForRule(MonitorRuleConfig rule)
        {
            if (rule == null) return;
            var watcher = new MonitorDirectoryWatcher(rule.DirectoryPath, rule.FilePattern);
            watcher.FileMatched += OnFileMatched;
            watcher.StartBaseline();
            _watchers[rule.Name ?? Guid.NewGuid().ToString()] = watcher;
        }

        private void StopWatcherForRule(MonitorRuleConfig rule)
        {
            if (rule == null) return;
            if (rule.Name != null && _watchers.TryGetValue(rule.Name, out var w))
            {
                w.Dispose();
                _watchers.Remove(rule.Name);
            }
        }

        private void OnFileMatched(string path)
        {
            if (_tailDedup.Add(path))
            {
                FileTailed?.Invoke(path);
            }
        }

        // Public API for external control (minimal): enable/disable rules by name
        public void EnableRule(string ruleName)
        {
            var rule = _rules.FirstOrDefault(r => string.Equals(r.Name, ruleName, StringComparison.OrdinalIgnoreCase));
            if (rule == null) return;
            if (rule.Enabled) return;
            rule.Enabled = true;
            StartWatcherForRule(rule);
        }

        public void DisableRule(string ruleName)
        {
            var rule = _rules.FirstOrDefault(r => string.Equals(r.Name, ruleName, StringComparison.OrdinalIgnoreCase));
            if (rule == null) return;
            if (!rule.Enabled) return;
            rule.Enabled = false;
            StopWatcherForRule(rule);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var w in _watchers.Values)
            {
                w.Dispose();
            }
            _watchers.Clear();
        }
    }
}
