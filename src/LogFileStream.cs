#region License statement
/* SnakeTail is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
#endregion

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace SnakeTail
{
    class LogFileStream : IDisposable
    {
        string _filePath = "";
        string _filePathAbsolute = "";
        Encoding _fileEncoding = Encoding.Default;
        FileStream _fileStream = null;
        StreamReader _fileReader = null;
        ThreadPoolQueue _threadPool = null;
        DateTime _lastFileCheck = DateTime.UtcNow;
        int _lastLineNumber = 0;
        string _lastFileCheckError = "";
        long _lastFileCheckLength = 0;
        TimeSpan _fileCheckFrequency = TimeSpan.FromSeconds(10);
        bool _fileCheckPattern = false;
        FileSystemWatcher _fileWatcher = null;
        int _fileChangedSignal = 0;
        DateTime _lastEventDrivenCheck = DateTime.MinValue;
        static readonly TimeSpan EventDrivenCheckDebounce = TimeSpan.FromMilliseconds(200);

        public event EventHandler FileReloadedEvent;
        public event EventHandler FileChangedEvent;

        public LogFileStream(string configPath, string filePath, Encoding fileEncoding, int fileCheckFrequency, bool fileCheckPattern)
        {
            _fileEncoding = fileEncoding;
            _filePath = filePath;
            _filePathAbsolute = Path.Combine(configPath, _filePath);
            if (fileCheckFrequency > 0)
                _fileCheckFrequency = TimeSpan.FromSeconds(fileCheckFrequency);
            _fileCheckPattern = fileCheckPattern;
            if (_fileCheckPattern)
                _threadPool = new ThreadPoolQueue(0);
            LoadFile(_filePathAbsolute, _fileEncoding, _fileCheckPattern);
            SetupFileWatcher();
        }

        ~LogFileStream()
        {
            Dispose();
        }

        public void Reset()
        {
            FileReloadedEvent = null;
         }

        public void CheckLogFile(bool forceReload)
        {
            _lastFileCheck = DateTime.UtcNow;

            try
            {
                // Refreshes the directory of the file, to ensure that we see the latest changes
                // If the directory is on a network share, then this can be a long blocking operation
                if (_threadPool != null)
                    _threadPool.CheckResult();

                DirectoryInfo dirInfo = null;
                try
                {
                    dirInfo = new DirectoryInfo(Path.GetDirectoryName(_filePathAbsolute));
                }
                catch (System.ArgumentException ex)
                {
                    // Any problems with the path should also be detected with the synchronous LoadFile-check
                    System.Diagnostics.Debug.WriteLine("Failed to refresh directory path: " + ex.Message);
                }
                catch (System.Security.SecurityException ex)
                {
                    // Any problems with the path should also be detected with the synchronous LoadFile-check
                    System.Diagnostics.Debug.WriteLine("Failed to access directory path: " + ex.Message);
                }
                catch (System.IO.IOException ex)
                {
                    // Any problems with the path should also be detected with the synchronous LoadFile-check
                    System.Diagnostics.Debug.WriteLine("Failed to read directory path: " + ex.Message);
                }
                if (dirInfo != null && _threadPool != null)
                    _threadPool.ExecuteRequest(RefreshDirectoryInfo, dirInfo);
            }
            catch (ApplicationException ex)
            {
                // Any problems with the path should also be detected with the synchronous LoadFile-check
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }

            if (_fileStream == null || forceReload)
            {
                LoadFile(_filePathAbsolute, _fileEncoding, _fileCheckPattern);
                if (_fileStream != null || forceReload)
                {
                    if (FileReloadedEvent != null)
                        FileReloadedEvent(this, null);
                }
            }
            else
            {
                string configPath = Path.GetDirectoryName(_filePathAbsolute);
                bool fileChanged = false;
                long fileCheckLength = 0;

                using (LogFileStream testLogFile = new LogFileStream(configPath, _filePathAbsolute, _fileEncoding, _fileCheckFrequency.Seconds, _fileCheckPattern))
                {
                    fileCheckLength = testLogFile.Length;
                    long currentFileLength = Length;
                    string name = testLogFile._fileStream != null ? testLogFile._fileStream.Name : null;

                    if (fileCheckLength < currentFileLength)
                        fileChanged = true;
                    else if (Position > fileCheckLength)
                        fileChanged = true;
                    else if (_fileStream.Name != name)
                        fileChanged = true;
                    else if (_lastFileCheckLength <= fileCheckLength && _lastFileCheckLength > currentFileLength)
                        fileChanged = true;
                }

                if (fileChanged)
                {
                    // The file have been renamed / deleted (reload new file)
                    LoadFile(_filePathAbsolute, _fileEncoding, _fileCheckPattern);
                    if (FileReloadedEvent != null)
                        FileReloadedEvent(this, null);
                }
                _lastFileCheckLength = fileCheckLength;
            }
        }

        static void RefreshDirectoryInfo(object state)
        {
            try
            {
                DirectoryInfo directoryInfo = state as DirectoryInfo;
                if (directoryInfo != null)
                    directoryInfo.Refresh();
            }
            catch (System.IO.IOException ex)
            {
                throw new ApplicationException("Failed to refresh directory path: " + ex.Message, ex);
            }
        }

        public long Length
        {
            get { return _fileStream != null ? _fileStream.Length : 0; }
        }

        public long Position
        {
            get { return _fileStream != null ? _fileStream.Position : 0; }
        }

        public string Name
        {
            get { return _fileStream != null ? _fileStream.Name : null; }
        }

        public Encoding FileEncoding
        {
            get { return _fileEncoding; }
        }

        public string FilePath
        {
            get { return _filePath; }
        }

        public string FilePathAbsolute
        {
            get { return _filePathAbsolute; }
        }

        public int FileCheckInterval
        {
            get { return (int)_fileCheckFrequency.TotalSeconds; }
        }

        public bool FileCheckPattern
        {
            get { return _fileCheckPattern; }
        }

        public bool ValidLineCount(int lineCount)
        {
            if (_fileStream != null && _lastLineNumber == lineCount)
                return true;
            else
                return false;
        }

        static string FindFileUsingPattern(string filePathAbsolute)
        {
            // Consider using FileSystemWatcher
            string filename = Path.GetFileName(filePathAbsolute);
            string directory = Path.GetDirectoryName(filePathAbsolute);
            DirectoryInfo dir = new DirectoryInfo(directory);
            FileInfo[] files = dir.GetFiles(filename);
            FileInfo lastestFile = null;
            foreach (FileInfo file in files)
            {
                if (lastestFile == null || lastestFile.LastWriteTime < file.LastWriteTime)
                    lastestFile = file;
            }
            if (lastestFile != null)
                return lastestFile.FullName;
            else
                return null;
        }

        public void Dispose()
        {
            Reset();

            if (_threadPool != null)
            {
                _threadPool.Dispose();
                _threadPool = null;
            }

            DisposeFileWatcher();

            CloseFile(false);
        }

        void CloseFile(bool publishEvent)
        {
            _lastFileCheckError = string.Empty;
            _lastFileCheck = DateTime.UtcNow;
            _lastFileCheckLength = 0;
            _lastLineNumber = 0;

            bool closedFile = false;
            if (_fileReader != null)
            {
                _fileReader.Dispose();
                _fileReader = null;
                closedFile = true;
            }
            if (_fileStream != null)
            {
                _fileStream.Dispose();
                _fileStream = null;
                closedFile = true;
            }

            if (publishEvent && closedFile)
            {
                if (FileReloadedEvent != null)
                    FileReloadedEvent(this, null);
            }
        }

        bool LoadFile(string filepath, Encoding fileEncoding, bool fileCheckPattern)
        {
            CloseFile(false);

            if (String.IsNullOrEmpty(filepath))
            {
                _lastFileCheckError = "No file path";
                return false;
            }
            else
            if (fileCheckPattern)
            {
                try
                {
                    filepath = FindFileUsingPattern(filepath);
                    if (filepath == null)
                    {
                        _lastFileCheckError = "No files matching pattern";
                        return false;
                    }
                }
                catch (ArgumentException ex)
                {
                    _lastFileCheckError = "Invalid file matching pattern path - " + ex.Message;
                    return false;
                }
                catch (DirectoryNotFoundException)
                {
                    _lastFileCheckError = "Directory not found";
                    return false;
                }
                catch (System.Security.SecurityException)
                {
                    _lastFileCheckError = "No permission to list folder contents";
                    return false;
                }
                catch (UnauthorizedAccessException)
                {
                    _lastFileCheckError = "Access to the directory is denied";
                    return false;
                }
                catch (IOException ex)
                {
                    _lastFileCheckError = ex.Message;
                    return false;
                }
            }

            try
            {
                _fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 65536, FileOptions.SequentialScan);
            }
            catch (ArgumentException ex)
            {
                _lastFileCheckError = "Invalid argument for opening file - " + ex.Message;
                return false;
            }
            catch (NotSupportedException ex)
            {
                _lastFileCheckError = "Not supported option used for opening file - " + ex.Message;
                return false;
            }
            catch (DirectoryNotFoundException)
            {
                _lastFileCheckError = "Directory not found";
                return false;
            }
            catch (System.Security.SecurityException)
            {
                _lastFileCheckError = "No permission to open the file";
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                _lastFileCheckError = "Access to the file is denied";
                return false;
            }
            catch (FileNotFoundException)
            {
                _lastFileCheckError = "File not found";
                return false;
            }
            catch (IOException ex)
            {
                _lastFileCheckError = ex.Message;
                return false;
            }

            _fileReader = new StreamReader(_fileStream, fileEncoding, true, 65536);

            try
            {
                if (!_fileReader.EndOfStream)
                    _lastFileCheckError = "";
            }
            catch (System.Security.SecurityException)
            {
                CloseFile(true);
                _lastFileCheckError = "No permission to read the file";
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                CloseFile(true);
                _lastFileCheckError = "Read access to the file is denied";
                return false;
            }
            catch (OperationCanceledException)
            {
                CloseFile(true);
                _lastFileCheckError = "Read file operation was aborted. File is currently not available.";
                return false;
            }
            catch (System.IO.IOException ex)
            {
                CloseFile(true);
                _lastFileCheckError = ex.Message;
                return false;
            }

            _lastLineNumber = 0;
            Interlocked.Exchange(ref _fileChangedSignal, 0);
            return true;
        }

        // 初始化文件系统监听：事件触发为主，轮询兜底。
        void SetupFileWatcher()
        {
            DisposeFileWatcher();

            try
            {
                string watchDirectory = Path.GetDirectoryName(_filePathAbsolute);
                string watchFilter = Path.GetFileName(_filePathAbsolute);
                if (string.IsNullOrEmpty(watchDirectory) || string.IsNullOrEmpty(watchFilter))
                    return;
                if (!Directory.Exists(watchDirectory))
                    return;

                _fileWatcher = new FileSystemWatcher(watchDirectory, watchFilter);
                _fileWatcher.IncludeSubdirectories = false;
                _fileWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime;
                _fileWatcher.Changed += FileWatcher_Changed;
                _fileWatcher.Created += FileWatcher_Changed;
                _fileWatcher.Deleted += FileWatcher_Changed;
                _fileWatcher.Renamed += FileWatcher_Renamed;
                _fileWatcher.Error += FileWatcher_Error;
                _fileWatcher.EnableRaisingEvents = true;
            }
            catch (ArgumentException ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to initialize file watcher due to invalid path or filter: " + ex.Message);
                DisposeFileWatcher();
            }
            catch (System.Security.SecurityException ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to initialize file watcher due to insufficient permission: " + ex.Message);
                DisposeFileWatcher();
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to initialize file watcher due to IO error: " + ex.Message);
                DisposeFileWatcher();
            }
        }

        // 释放监听资源，避免句柄泄漏。
        void DisposeFileWatcher()
        {
            if (_fileWatcher == null)
                return;

            _fileWatcher.EnableRaisingEvents = false;
            _fileWatcher.Changed -= FileWatcher_Changed;
            _fileWatcher.Created -= FileWatcher_Changed;
            _fileWatcher.Deleted -= FileWatcher_Changed;
            _fileWatcher.Renamed -= FileWatcher_Renamed;
            _fileWatcher.Error -= FileWatcher_Error;
            _fileWatcher.Dispose();
            _fileWatcher = null;
        }

        // 文件事件：仅打脏标记，避免事件线程直接做重 IO。
        void FileWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (IsWatchEventMatch(_filePathAbsolute, _fileStream != null ? _fileStream.Name : null, e != null ? e.FullPath : null, _fileCheckPattern))
            {
                Interlocked.Exchange(ref _fileChangedSignal, 1);
                if (FileChangedEvent != null)
                    FileChangedEvent(this, EventArgs.Empty);
            }
        }

        // 重命名事件：新旧路径任一命中都视为变化。
        void FileWatcher_Renamed(object sender, RenamedEventArgs e)
        {
            string oldPath = e != null ? e.OldFullPath : null;
            string newPath = e != null ? e.FullPath : null;
            if (IsWatchEventMatch(_filePathAbsolute, _fileStream != null ? _fileStream.Name : null, oldPath, _fileCheckPattern) ||
                IsWatchEventMatch(_filePathAbsolute, _fileStream != null ? _fileStream.Name : null, newPath, _fileCheckPattern))
            {
                Interlocked.Exchange(ref _fileChangedSignal, 1);
                if (FileChangedEvent != null)
                    FileChangedEvent(this, EventArgs.Empty);
            }
        }

        // 监听错误：标记脏并让兜底轮询尽快接管。
        void FileWatcher_Error(object sender, ErrorEventArgs e)
        {
            Interlocked.Exchange(ref _fileChangedSignal, 1);
            if (FileChangedEvent != null)
                FileChangedEvent(this, EventArgs.Empty);
        }

        // 判断事件路径是否与当前跟踪目标相关。
        internal static bool IsWatchEventMatch(string configuredPathAbsolute, string openedFilePath, string eventPath, bool fileCheckPattern)
        {
            if (string.IsNullOrEmpty(eventPath))
                return false;

            if (fileCheckPattern)
                return true;

            string targetPath = !string.IsNullOrEmpty(openedFilePath) ? openedFilePath : configuredPathAbsolute;
            if (string.IsNullOrEmpty(targetPath))
                return false;

            try
            {
                string fullTarget = Path.GetFullPath(targetPath);
                string fullEvent = Path.GetFullPath(eventPath);
                return string.Equals(fullTarget, fullEvent, StringComparison.OrdinalIgnoreCase);
            }
            catch (ArgumentException)
            {
                return string.Equals(targetPath, eventPath, StringComparison.OrdinalIgnoreCase);
            }
            catch (NotSupportedException)
            {
                return string.Equals(targetPath, eventPath, StringComparison.OrdinalIgnoreCase);
            }
            catch (PathTooLongException)
            {
                return string.Equals(targetPath, eventPath, StringComparison.OrdinalIgnoreCase);
            }
        }

        public string ReadLine(int lineNumber)
        {
            if (_fileReader == null || _fileStream == null)
            {
                // Check if file is available (once a second)
                if (_lastFileCheck != DateTime.UtcNow)
                    CheckLogFile(true);

                if (lineNumber == 1)
                    return "Cannot open file: " + _filePathAbsolute + (String.IsNullOrEmpty(_lastFileCheckError) ? "" : " (" + _lastFileCheckError + ")");
                else
                    return null;
            }

            try
            {
                if (lineNumber <= _lastLineNumber)
                {
                    _fileStream.Seek(0, SeekOrigin.Begin);
                    _fileReader.DiscardBufferedData();
                    _lastLineNumber = 0;
                }
                else
                {
                    lineNumber -= _lastLineNumber;
                }

                if (_fileReader.EndOfStream)
                {
                    // 事件触发优先：有变化信号时防抖检查，减少空闲轮询。
                    if (Interlocked.Exchange(ref _fileChangedSignal, 0) == 1)
                    {
                        if (DateTime.UtcNow.Subtract(_lastEventDrivenCheck) >= EventDrivenCheckDebounce)
                        {
                            CheckLogFile(false);
                            _lastEventDrivenCheck = DateTime.UtcNow;
                        }
                        else
                        {
                            Interlocked.Exchange(ref _fileChangedSignal, 1);
                        }
                    }
                    // 兜底轮询：覆盖丢事件、网络盘延迟或监听错误场景。
                    else if (DateTime.UtcNow.Subtract(_lastFileCheck) >= _fileCheckFrequency)
                    {
                        CheckLogFile(false);
                    }

                    // EOF 兜底：即使监听事件丢失，也尝试按文件长度重同步读取器。
                    if (!TryResyncReaderAtEof())
                        return null;
                }

                string line = null;
                for (int i = 0; i < lineNumber; ++i)
                {
                    line = _fileReader.ReadLine();
                    if (line == null)
                        return null;

                    _lastLineNumber++;
                }

                _lastFileCheck = DateTime.UtcNow;
                return line;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                CloseFile(true);
                if (lineNumber == 1)
                    return "Cannot read file: " + _filePathAbsolute + " (" + ex.Message + ")";
                return null;
            }
            catch (System.IO.IOException ex)
            {
                CloseFile(true);
                if (lineNumber == 1)
                    return "Cannot read file: " + _filePathAbsolute + " (" + ex.Message + ")";
                return null;
            }
        }

        // EOF 时按底层流长度重同步读取器，避免遗漏追加内容。
        bool TryResyncReaderAtEof()
        {
            if (_fileReader == null || _fileStream == null)
                return false;

            long streamLength = _fileStream.Length;
            long streamPosition = _fileStream.Position;
            if (streamLength <= streamPosition)
                return false;

            _fileStream.Seek(streamPosition, SeekOrigin.Begin);
            _fileReader.DiscardBufferedData();
            return !_fileReader.EndOfStream;
        }

        public int SkipLines(long lineCount)
        {
            // Quickly fast forward to near the file bottom
            try
            {
                long fileLength = Length - lineCount * 80 * (FileEncoding.IsSingleByte ? 1 : 2);
                long filePosiion = Position;
                for(int i = 0; i < lineCount && filePosiion < fileLength && !_fileReader.EndOfStream; ++i)
                {
                    string line = ReadLine(_lastLineNumber + 1);
                    if (line == null)
                        return _lastLineNumber;

                    filePosiion += line.Length * (FileEncoding.IsSingleByte ? 1 : 2);
                }

                return _lastLineNumber;
            }
            catch (System.Security.SecurityException)
            {
                CloseFile(true);
                _lastFileCheckError = "No permission to read the file";
                return -1;
            }
            catch (UnauthorizedAccessException)
            {
                CloseFile(true);
                _lastFileCheckError = "Read access to the file is denied";
                return -1;
            }
            catch (OperationCanceledException)
            {
                CloseFile(true);
                _lastFileCheckError = "Read file operation was aborted. File is currently not available.";
                return -1;
            }
            catch (System.IO.IOException ex)
            {
                CloseFile(true);
                _lastFileCheckError = ex.Message;
                return -1;
            }
        }
    }
}
