using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SnakeTail
{
    /// <summary>
    /// 已加载的显示插件信息。
    /// </summary>
    internal sealed class LoadedDisplayPlugin
    {
        public LoadedDisplayPlugin(string directoryName, string pluginDirectoryPath, string displayName, ILogDisplayPlugin instance, bool isAvailable, string unavailableReason)
        {
            DirectoryName = directoryName;
            PluginDirectoryPath = pluginDirectoryPath;
            DisplayName = displayName;
            Instance = instance;
            IsAvailable = isAvailable;
            UnavailableReason = unavailableReason ?? string.Empty;
        }

        public string DirectoryName { get; }
        public string PluginDirectoryPath { get; }
        public string DisplayName { get; }
        public ILogDisplayPlugin Instance { get; }
        public bool IsAvailable { get; }
        public string UnavailableReason { get; }
    }

    /// <summary>
    /// 显示插件管理器：负责扫描、装载、实例化与初始化。
    /// </summary>
    internal sealed class DisplayPluginManager
    {
        private readonly string _configPath;
        private readonly List<LoadedDisplayPlugin> _availablePlugins = new List<LoadedDisplayPlugin>();
        private readonly Dictionary<string, LoadedDisplayPlugin> _pluginsByName = new Dictionary<string, LoadedDisplayPlugin>(StringComparer.CurrentCultureIgnoreCase);

        public DisplayPluginManager(string configPath)
        {
            _configPath = configPath ?? string.Empty;
        }

        public IReadOnlyList<LoadedDisplayPlugin> AvailablePlugins => _availablePlugins;

        /// <summary>
        /// 重新扫描插件目录，并构建可用插件列表。
        /// </summary>
        public void Reload(string currentLogFilePath)
        {
            _availablePlugins.Clear();
            _pluginsByName.Clear();

            string pluginRootPath = GetPluginRootPath();
            if (!Directory.Exists(pluginRootPath))
                return;

            foreach (string pluginDirectory in Directory.GetDirectories(pluginRootPath))
            {
                string pluginDirectoryName = Path.GetFileName(pluginDirectory);
                if (string.IsNullOrEmpty(pluginDirectoryName))
                    continue;

                if (TryLoadPluginDirectory(pluginDirectory, pluginDirectoryName, currentLogFilePath, out LoadedDisplayPlugin loadedPlugin))
                {
                    if (_pluginsByName.ContainsKey(loadedPlugin.DisplayName))
                    {
                        AppLog.AppendDaily(AppLog.LevelWarn, string.Format("显示插件重名已忽略: Name={0}, Directory={1}", loadedPlugin.DisplayName, loadedPlugin.PluginDirectoryPath));
                        continue;
                    }

                    _availablePlugins.Add(loadedPlugin);
                    _pluginsByName[loadedPlugin.DisplayName] = loadedPlugin;
                }
            }
        }

        /// <summary>
        /// 根据显示名查找插件实例。
        /// </summary>
        public LoadedDisplayPlugin FindByDisplayName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName))
                return null;

            _pluginsByName.TryGetValue(displayName, out LoadedDisplayPlugin plugin);
            return plugin;
        }

        /// <summary>
        /// 解析插件根目录（config/plugins）。
        /// </summary>
        internal string GetPluginRootPath()
        {
            string basePath = !string.IsNullOrEmpty(_configPath) ? _configPath : Directory.GetCurrentDirectory();
            string configRootPath = Path.Combine(basePath, "config");
            return Path.Combine(configRootPath, "plugins");
        }

        /// <summary>
        /// 扫描单个插件目录，按规则加载首个合法插件。
        /// </summary>
        internal bool TryLoadPluginDirectory(string pluginDirectoryPath, string pluginDirectoryName, string currentLogFilePath, out LoadedDisplayPlugin loadedPlugin)
        {
            loadedPlugin = null;

            string[] dllFiles;
            try
            {
                dllFiles = Directory.GetFiles(pluginDirectoryPath, "*.dll", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                AppLog.AppendDaily(AppLog.LevelErr, string.Format("扫描插件目录失败: Directory={0}, Error={1}: {2}", pluginDirectoryPath, ex.GetType().FullName, ex.Message));
                return false;
            }

            if (dllFiles.Length == 0)
            {
                AppLog.AppendDaily(AppLog.LevelWarn, string.Format("插件目录未发现 DLL: Directory={0}", pluginDirectoryPath));
                return false;
            }

            foreach (string dllPath in dllFiles)
            {
                if (TryLoadPluginFromAssembly(dllPath, pluginDirectoryPath, pluginDirectoryName, currentLogFilePath, out loadedPlugin))
                    return true;
            }

            AppLog.AppendDaily(AppLog.LevelWarn, string.Format("插件目录未发现 ILogDisplayPlugin 实现: Directory={0}", pluginDirectoryPath));
            return false;
        }

        /// <summary>
        /// 从 DLL 加载并初始化插件。
        /// </summary>
        private bool TryLoadPluginFromAssembly(string assemblyPath, string pluginDirectoryPath, string pluginDirectoryName, string currentLogFilePath, out LoadedDisplayPlugin loadedPlugin)
        {
            loadedPlugin = null;

            Assembly assembly;
            try
            {
                assembly = Assembly.LoadFrom(assemblyPath);
            }
            catch (Exception ex)
            {
                AppLog.AppendDaily(AppLog.LevelErr, string.Format("加载插件 DLL 失败: Dll={0}, Error={1}: {2}", assemblyPath, ex.GetType().FullName, ex.Message));
                return false;
            }

            if (!TryCreatePluginInstance(assembly, out ILogDisplayPlugin plugin, out string createError))
            {
                if (!string.IsNullOrEmpty(createError))
                {
                    AppLog.AppendDaily(AppLog.LevelErr, string.Format("创建插件实例失败: Dll={0}, Error={1}", assemblyPath, createError));
                }
                return false;
            }

            string configRootPath = Path.Combine(!string.IsNullOrEmpty(_configPath) ? _configPath : Directory.GetCurrentDirectory(), "config");
            string hostVersion = Assembly.GetExecutingAssembly().GetName().Version != null
                ? Assembly.GetExecutingAssembly().GetName().Version.ToString()
                : string.Empty;
            PluginContext context = new PluginContext(pluginDirectoryPath, configRootPath, currentLogFilePath, hostVersion);
            try
            {
                plugin.Initialize(context);
            }
            catch (FileNotFoundException ex)
            {
                string pluginName = GetPluginDisplayName(plugin, pluginDirectoryName);
                string unavailableReason = BuildMissingConfigReason(ex);
                AppLog.AppendDaily(AppLog.LevelWarn, string.Format("插件缺少配置文件，暂不可启用: Dll={0}, Plugin={1}, Reason={2}", assemblyPath, pluginName, unavailableReason));
                loadedPlugin = new LoadedDisplayPlugin(pluginDirectoryName, pluginDirectoryPath, pluginName, null, false, unavailableReason);
                return true;
            }
            catch (DirectoryNotFoundException ex)
            {
                string pluginName = GetPluginDisplayName(plugin, pluginDirectoryName);
                string unavailableReason = string.Format("缺少配置目录：{0}", ex.Message);
                AppLog.AppendDaily(AppLog.LevelWarn, string.Format("插件缺少配置目录，暂不可启用: Dll={0}, Plugin={1}, Reason={2}", assemblyPath, pluginName, unavailableReason));
                loadedPlugin = new LoadedDisplayPlugin(pluginDirectoryName, pluginDirectoryPath, pluginName, null, false, unavailableReason);
                return true;
            }
            catch (Exception ex)
            {
                AppLog.AppendDaily(AppLog.LevelErr, string.Format("插件初始化失败: Dll={0}, Plugin={1}, Error={2}: {3}", assemblyPath, plugin.GetType().FullName, ex.GetType().FullName, ex.Message));
                return false;
            }

            string pluginDisplayName = GetPluginDisplayName(plugin, pluginDirectoryName);
            loadedPlugin = new LoadedDisplayPlugin(pluginDirectoryName, pluginDirectoryPath, pluginDisplayName, plugin, true, string.Empty);
            return true;
        }

        // 统一解析插件显示名，避免重复空值分支。
        private static string GetPluginDisplayName(ILogDisplayPlugin plugin, string pluginDirectoryName)
        {
            string pluginName = plugin != null ? plugin.Name : string.Empty;
            if (string.IsNullOrWhiteSpace(pluginName))
                pluginName = pluginDirectoryName;
            return pluginName;
        }

        // 精确拼出缺失配置文件原因，供 UI 提示与日志复用。
        private static string BuildMissingConfigReason(FileNotFoundException ex)
        {
            string missingPath = ex.FileName;
            if (!string.IsNullOrWhiteSpace(missingPath))
                return string.Format("{0}：{1}", ex.Message, missingPath);

            return ex.Message;
        }

        /// <summary>
        /// 在程序集内查找首个可实例化插件类型。
        /// </summary>
        internal static bool TryCreatePluginInstance(Assembly assembly, out ILogDisplayPlugin plugin, out string error)
        {
            plugin = null;
            error = null;

            Type[] types;
            try
            {
                types = assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                IEnumerable<string> loaderMessages = ex.LoaderExceptions != null
                    ? ex.LoaderExceptions.Where(x => x != null).Select(x => x.GetType().FullName + ": " + x.Message)
                    : Enumerable.Empty<string>();
                error = "反射类型加载失败: " + string.Join(" | ", loaderMessages);
                return false;
            }
            catch (Exception ex)
            {
                error = ex.GetType().FullName + ": " + ex.Message;
                return false;
            }

            Type pluginType = types.FirstOrDefault(t =>
                t != null &&
                typeof(ILogDisplayPlugin).IsAssignableFrom(t) &&
                t.IsClass &&
                !t.IsAbstract &&
                t.GetConstructor(Type.EmptyTypes) != null);
            if (pluginType == null)
                return false;

            try
            {
                plugin = (ILogDisplayPlugin)Activator.CreateInstance(pluginType);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.GetType().FullName + ": " + ex.Message;
                return false;
            }
        }
    }
}
