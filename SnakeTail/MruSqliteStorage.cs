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
using System.Data;
using System.IO;
using System.Reflection;

namespace SnakeTail
{
    /// <summary>
    /// SQLite 存储最近打开的文件列表（使用反射动态加载 SQLite）
    /// </summary>
    public class MruSqliteStorage : IDisposable
    {
        private string _dbPath;
        private object _connection;
        private readonly object _lockObject = new object();
        private Assembly _sqliteAssembly;
        private Type _sqliteConnectionType;
        private Type _sqliteCommandType;
        private Type _sqliteDataReaderType;
        private bool _isAvailable = false;

        public MruSqliteStorage(string dbPath)
        {
            if (string.IsNullOrEmpty(dbPath))
            {
                // 默认使用可执行文件目录
                string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                _dbPath = Path.Combine(exeDir, "recentfiles.db");
            }
            else
            {
                _dbPath = dbPath;
            }

            if (LoadSqliteAssembly())
            {
                InitializeDatabase();
            }
        }

        private bool LoadSqliteAssembly()
        {
            try
            {
                // 尝试从多个位置加载 SQLite
                string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string[] possiblePaths = new string[]
                {
                    Path.Combine(exeDir, "System.Data.SQLite.dll"),
                    Path.Combine(Path.Combine(exeDir, "libs"), "System.Data.SQLite.dll"),
                    "System.Data.SQLite.dll"
                };

                foreach (string dllPath in possiblePaths)
                {
                    try
                    {
                        if (File.Exists(dllPath))
                        {
                            _sqliteAssembly = Assembly.LoadFrom(dllPath);
                            _sqliteConnectionType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteConnection");
                            _sqliteCommandType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteCommand");
                            _sqliteDataReaderType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteDataReader");

                            if (_sqliteConnectionType != null && _sqliteCommandType != null && _sqliteDataReaderType != null)
                            {
                                _isAvailable = true;
                                return true;
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                // 尝试从 GAC 加载
                try
                {
                    _sqliteAssembly = Assembly.Load("System.Data.SQLite, Version=1.0.0.0, Culture=neutral, PublicKeyToken=db937bc2d44ff139");
                    _sqliteConnectionType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteConnection");
                    _sqliteCommandType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteCommand");
                    _sqliteDataReaderType = _sqliteAssembly.GetType("System.Data.SQLite.SQLiteDataReader");

                    if (_sqliteConnectionType != null && _sqliteCommandType != null && _sqliteDataReaderType != null)
                    {
                        _isAvailable = true;
                        return true;
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            _isAvailable = false;
            return false;
        }

        public bool IsAvailable
        {
            get { return _isAvailable; }
        }

        private void InitializeDatabase()
        {
            if (!_isAvailable)
                return;

            lock (_lockObject)
            {
                try
                {
                    bool isNew = !File.Exists(_dbPath);

                    // 确保目录存在
                    string directory = Path.GetDirectoryName(_dbPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    string connectionString = string.Format("Data Source={0};Version=3;", _dbPath);
                    _connection = Activator.CreateInstance(_sqliteConnectionType, new object[] { connectionString });

                    MethodInfo openMethod = _sqliteConnectionType.GetMethod("Open");
                    openMethod.Invoke(_connection, null);

                    if (isNew)
                    {
                        CreateTable();
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        string.Format("无法初始化 SQLite 数据库: {0}\n数据库路径: {1}", ex.Message, _dbPath), ex);
                }
            }
        }

        private void CreateTable()
        {
            string sql = @"
                CREATE TABLE IF NOT EXISTS RecentFiles (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FilePath TEXT NOT NULL UNIQUE,
                    LastAccessed DATETIME NOT NULL,
                    AccessCount INTEGER DEFAULT 1
                );
                CREATE INDEX IF NOT EXISTS idx_LastAccessed ON RecentFiles(LastAccessed DESC);
            ";

            ExecuteNonQuery(sql);
        }

        private void ExecuteNonQuery(string sql, params object[] parameters)
        {
            if (!_isAvailable || _connection == null || _sqliteCommandType == null)
                return;

            object command = null;
            try
            {
                command = Activator.CreateInstance(_sqliteCommandType, new object[] { sql, _connection });
                if (command == null)
                    return;

                if (parameters != null && parameters.Length > 0)
                {
                    MethodInfo addParamMethod = _sqliteCommandType.GetMethod("Parameters");
                    if (addParamMethod != null)
                    {
                        object paramCollection = addParamMethod.Invoke(command, null);
                        if (paramCollection != null)
                        {
                            MethodInfo addWithValueMethod = paramCollection.GetType().GetMethod("AddWithValue");
                            if (addWithValueMethod != null)
                            {
                                for (int i = 0; i < parameters.Length; i += 2)
                                {
                                    addWithValueMethod.Invoke(paramCollection, new object[] { parameters[i], parameters[i + 1] });
                                }
                            }
                        }
                    }
                }

                MethodInfo executeMethod = _sqliteCommandType.GetMethod("ExecuteNonQuery");
                if (executeMethod != null)
                {
                    executeMethod.Invoke(command, null);
                }
            }
            catch
            {
                // 静默失败
            }
            finally
            {
                if (command != null)
                {
                    try
                    {
                        MethodInfo disposeMethod = _sqliteCommandType.GetMethod("Dispose");
                        if (disposeMethod != null)
                        {
                            disposeMethod.Invoke(command, null);
                        }
                    }
                    catch
                    {
                        // 忽略释放错误
                    }
                }
            }
        }

        private object ExecuteReader(string sql, params object[] parameters)
        {
            if (!_isAvailable || _connection == null || _sqliteCommandType == null)
                return null;

            try
            {
                object command = Activator.CreateInstance(_sqliteCommandType, new object[] { sql, _connection });
                if (command == null)
                    return null;

                if (parameters != null && parameters.Length > 0)
                {
                    MethodInfo addParamMethod = _sqliteCommandType.GetMethod("Parameters");
                    if (addParamMethod != null)
                    {
                        object paramCollection = addParamMethod.Invoke(command, null);
                        if (paramCollection != null)
                        {
                            MethodInfo addWithValueMethod = paramCollection.GetType().GetMethod("AddWithValue");
                            if (addWithValueMethod != null)
                            {
                                for (int i = 0; i < parameters.Length; i += 2)
                                {
                                    addWithValueMethod.Invoke(paramCollection, new object[] { parameters[i], parameters[i + 1] });
                                }
                            }
                        }
                    }
                }

                MethodInfo executeReaderMethod = _sqliteCommandType.GetMethod("ExecuteReader");
                if (executeReaderMethod != null)
                {
                    return executeReaderMethod.Invoke(command, null);
                }
            }
            catch
            {
                // 静默失败
            }

            return null;
        }

        /// <summary>
        /// 添加或更新文件记录
        /// </summary>
        public void AddFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !_isAvailable || _connection == null)
                return;

            try
            {
                lock (_lockObject)
                {
                    string fullPath = Path.GetFullPath(filePath);
                    string sql = @"
                        INSERT OR REPLACE INTO RecentFiles (FilePath, LastAccessed, AccessCount)
                        VALUES (@FilePath, @LastAccessed,
                            COALESCE((SELECT AccessCount FROM RecentFiles WHERE FilePath = @FilePath), 0) + 1)
                    ";

                    ExecuteNonQuery(sql, "@FilePath", fullPath, "@LastAccessed", DateTime.Now);
                }
            }
            catch
            {
                // 静默失败，不影响主程序运行
            }
        }

        /// <summary>
        /// 获取最近打开的文件列表，按最后访问时间排序
        /// </summary>
        public List<string> GetRecentFiles(int maxCount)
        {
            List<string> files = new List<string>();

            if (!_isAvailable || _connection == null)
                return files;

            try
            {
                lock (_lockObject)
                {
                    string sql = @"
                        SELECT FilePath
                        FROM RecentFiles
                        ORDER BY LastAccessed DESC
                        LIMIT @MaxCount
                    ";

                    object reader = ExecuteReader(sql, "@MaxCount", maxCount);
                    if (reader != null)
                    {
                        try
                        {
                            MethodInfo readMethod = _sqliteDataReaderType.GetMethod("Read");
                            MethodInfo getStringMethod = _sqliteDataReaderType.GetMethod("GetString", new Type[] { typeof(int) });

                            while ((bool)readMethod.Invoke(reader, null))
                            {
                                string filePath = (string)getStringMethod.Invoke(reader, new object[] { 0 });
                                if (File.Exists(filePath))
                                {
                                    files.Add(filePath);
                                }
                            }
                        }
                        finally
                        {
                            MethodInfo closeMethod = _sqliteDataReaderType.GetMethod("Close");
                            if (closeMethod != null)
                            {
                                closeMethod.Invoke(reader, null);
                            }
                        }
                    }
                }
            }
            catch
            {
                // 返回空列表，不影响主程序运行
            }

            return files;
        }

        /// <summary>
        /// 删除文件记录
        /// </summary>
        public void RemoveFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !_isAvailable || _connection == null)
                return;

            try
            {
                lock (_lockObject)
                {
                    string fullPath = Path.GetFullPath(filePath);
                    string sql = "DELETE FROM RecentFiles WHERE FilePath = @FilePath";

                    ExecuteNonQuery(sql, "@FilePath", fullPath);
                }
            }
            catch
            {
                // 静默失败
            }
        }

        /// <summary>
        /// 重命名文件记录
        /// </summary>
        public void RenameFile(string oldFilePath, string newFilePath)
        {
            if (string.IsNullOrEmpty(oldFilePath) || string.IsNullOrEmpty(newFilePath) || !_isAvailable || _connection == null)
                return;

            try
            {
                lock (_lockObject)
                {
                    string oldFullPath = Path.GetFullPath(oldFilePath);
                    string newFullPath = Path.GetFullPath(newFilePath);

                    string sql = @"
                        UPDATE RecentFiles
                        SET FilePath = @NewFilePath, LastAccessed = @LastAccessed
                        WHERE FilePath = @OldFilePath
                    ";

                    ExecuteNonQuery(sql, "@OldFilePath", oldFullPath, "@NewFilePath", newFullPath, "@LastAccessed", DateTime.Now);
                }
            }
            catch
            {
                // 静默失败
            }
        }

        /// <summary>
        /// 清空所有记录
        /// </summary>
        public void ClearAll()
        {
            if (!_isAvailable || _connection == null)
                return;

            try
            {
                lock (_lockObject)
                {
                    string sql = "DELETE FROM RecentFiles";

                    ExecuteNonQuery(sql);
                }
            }
            catch
            {
                // 静默失败
            }
        }

        /// <summary>
        /// 清理不存在的文件记录
        /// </summary>
        public void CleanupNonExistentFiles()
        {
            if (!_isAvailable || _connection == null || _sqliteDataReaderType == null)
                return;

            try
            {
                lock (_lockObject)
                {
                    List<string> filesToRemove = new List<string>();

                    object reader = ExecuteReader("SELECT FilePath FROM RecentFiles");
                    if (reader != null && _sqliteDataReaderType != null)
                    {
                        try
                        {
                            MethodInfo readMethod = _sqliteDataReaderType.GetMethod("Read");
                            MethodInfo getStringMethod = _sqliteDataReaderType.GetMethod("GetString", new Type[] { typeof(int) });

                            if (readMethod != null && getStringMethod != null)
                            {
                                while ((bool)readMethod.Invoke(reader, null))
                                {
                                    string filePath = (string)getStringMethod.Invoke(reader, new object[] { 0 });
                                    if (!File.Exists(filePath))
                                    {
                                        filesToRemove.Add(filePath);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            MethodInfo closeMethod = _sqliteDataReaderType.GetMethod("Close");
                            if (closeMethod != null)
                            {
                                try
                                {
                                    closeMethod.Invoke(reader, null);
                                }
                                catch
                                {
                                    // 忽略关闭错误
                                }
                            }
                        }
                    }

                    foreach (string filePath in filesToRemove)
                    {
                        RemoveFile(filePath);
                    }
                }
            }
            catch
            {
                // 静默失败
            }
        }

        public void Dispose()
        {
            if (_connection == null)
                return;

            try
            {
                if (_sqliteConnectionType != null)
                {
                    lock (_lockObject)
                    {
                        try
                        {
                            MethodInfo closeMethod = _sqliteConnectionType.GetMethod("Close");
                            if (closeMethod != null && _connection != null)
                            {
                                closeMethod.Invoke(_connection, null);
                            }
                        }
                        catch
                        {
                            // 忽略关闭错误
                        }

                        try
                        {
                            MethodInfo disposeMethod = _sqliteConnectionType.GetMethod("Dispose");
                            if (disposeMethod != null && _connection != null)
                            {
                                disposeMethod.Invoke(_connection, null);
                            }
                        }
                        catch
                        {
                            // 忽略释放错误
                        }
                    }
                }
            }
            catch
            {
                // 静默失败
            }
            finally
            {
                _connection = null;
            }
        }
    }
}
