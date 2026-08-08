using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SnakeTail
{
    // Watches a single directory (non-recursive) and tails files matching a wildcard pattern
    public class MonitorDirectoryWatcher : IDisposable
    {
        private readonly string _directoryPath;
        private readonly string _filePattern;
        private FileSystemWatcher _watcher;
        private readonly HashSet<string> _tailedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _sync = new object();
        private bool _disposed;

        public event Action<string> FileMatched; // Fired when a file matches baseline or on creation

        public MonitorDirectoryWatcher(string directoryPath, string filePattern)
        {
            _directoryPath = directoryPath ?? throw new ArgumentNullException(nameof(directoryPath));
            _filePattern = filePattern ?? throw new ArgumentNullException(nameof(filePattern));
        }

        public void StartBaseline()
        {
            if (!Directory.Exists(_directoryPath)) return;
            // Baseline discovery: find existing files that match the pattern
            var all = Directory.GetFiles(_directoryPath, "*", SearchOption.TopDirectoryOnly);
            foreach (var fullPath in all)
            {
                var fileName = Path.GetFileName(fullPath);
                if (IsMatch(fileName))
                {
                    lock (_sync)
                    {
                        if (_tailedFiles.Add(fullPath))
                        {
                            FileMatched?.Invoke(fullPath);
                        }
                    }
                }
            }

            // Set up watcher for new files (non-recursive)
            _watcher = new FileSystemWatcher(_directoryPath)
            {
                Filter = "*",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };
            _watcher.Created += OnCreated;
            _watcher.EnableRaisingEvents = true;
            _watcher.IncludeSubdirectories = false;
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Debounce: ensure we have a physical file and it matches
            var fileName = Path.GetFileName(e.Name);
            if (IsMatch(fileName))
            {
                var fullPath = Path.Combine(_directoryPath, fileName);
                lock (_sync)
                {
                    if (_tailedFiles.Add(fullPath))
                    {
                        FileMatched?.Invoke(fullPath);
                    }
                }
            }
        }

        private bool IsMatch(string fileName)
        {
            // Convert wildcard to regex: * -> .* , ? -> .
            var pattern = "^" + Regex.Escape(_filePattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
            return Regex.IsMatch(fileName, pattern, RegexOptions.IgnoreCase);
        }

        public void Stop()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Stop();
        }
    }
}
