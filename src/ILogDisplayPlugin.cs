using System;

namespace SnakeTail
{
    /// <summary>
    /// 日志显示插件接口：只处理显示链路中的文本，不修改原始日志文件。
    /// </summary>
    public interface ILogDisplayPlugin
    {
        /// <summary>
        /// 插件显示名称；为空时宿主回退到插件目录名。
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 初始化插件上下文。
        /// </summary>
        void Initialize(PluginContext context);

        /// <summary>
        /// 判断当前行是否需要进入处理。
        /// </summary>
        bool CanProcess(string line);

        /// <summary>
        /// 尝试处理当前行；Handled=true 表示接管并终止后续插件。
        /// </summary>
        PluginProcessResult TryProcess(string line);
    }

    /// <summary>
    /// 插件上下文：提供只读目录与宿主信息。
    /// </summary>
    public sealed class PluginContext
    {
        public PluginContext(string pluginDirectoryAbsolutePath, string configRootDirectoryAbsolutePath, string currentLogFilePath, string hostVersion)
        {
            PluginDirectoryAbsolutePath = pluginDirectoryAbsolutePath ?? string.Empty;
            ConfigRootDirectoryAbsolutePath = configRootDirectoryAbsolutePath ?? string.Empty;
            CurrentLogFilePath = currentLogFilePath ?? string.Empty;
            HostVersion = hostVersion ?? string.Empty;
        }

        public string PluginDirectoryAbsolutePath { get; }
        public string ConfigRootDirectoryAbsolutePath { get; }
        public string CurrentLogFilePath { get; }
        public string HostVersion { get; }
    }

    /// <summary>
    /// 插件处理结果。
    /// </summary>
    public struct PluginProcessResult
    {
        public bool Handled { get; set; }
        public string Output { get; set; }
        public string ErrorMessage { get; set; }
    }
}
