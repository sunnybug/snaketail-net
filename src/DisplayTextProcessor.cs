using System;
using System.Collections.Generic;
using System.Linq;

namespace SnakeTail
{
    /// <summary>
    /// 显示文本处理器：按启用顺序串行执行插件并缓存结果。
    /// </summary>
    internal sealed class DisplayTextProcessor
    {
        private readonly Dictionary<string, string> _lineCache = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly List<LoadedDisplayPlugin> _enabledPlugins = new List<LoadedDisplayPlugin>();

        public IReadOnlyList<LoadedDisplayPlugin> EnabledPlugins => _enabledPlugins;

        /// <summary>
        /// 更新启用插件顺序。
        /// </summary>
        public void SetEnabledPlugins(IEnumerable<LoadedDisplayPlugin> plugins)
        {
            _enabledPlugins.Clear();
            if (plugins != null)
                _enabledPlugins.AddRange(plugins.Where(p => p != null));
            ClearCache();
        }

        /// <summary>
        /// 清理处理缓存。
        /// </summary>
        public void ClearCache()
        {
            _lineCache.Clear();
        }

        /// <summary>
        /// 获取处理后文本；无插件或不命中时返回原文。
        /// </summary>
        public string GetProcessedLineText(int lineKey, string rawLineText, Func<int, string> readLineByLineKey = null)
        {
            if (rawLineText == null)
                rawLineText = string.Empty;

            if (_enabledPlugins.Count == 0)
                return rawLineText;

            string signature = GetEnabledSignature();
            string cacheKey = string.Format("{0}|{1}", lineKey, signature);
            if (_lineCache.TryGetValue(cacheKey, out string cachedText))
                return cachedText;

            string currentText = rawLineText;
            foreach (LoadedDisplayPlugin loadedPlugin in _enabledPlugins)
            {
                ILogDisplayPlugin plugin = loadedPlugin.Instance;
                if (plugin == null)
                    continue;

                // 块插件优先尝试收集整块文本，确保按插件顺序决定输入粒度。
                string textForPlugin = currentText;
                if (plugin is ILogDisplayBlockPlugin blockPlugin && readLineByLineKey != null)
                {
                    try
                    {
                        if (blockPlugin.TryCollectBlock(lineKey, rawLineText, readLineByLineKey, out string blockText)
                            && !string.IsNullOrEmpty(blockText))
                        {
                            textForPlugin = blockText;
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLog.AppendDaily(AppLog.LevelErr, string.Format("插件 TryCollectBlock 失败: Plugin={0}, LineKey={1}, Error={2}: {3}", loadedPlugin.DisplayName, lineKey, ex.GetType().FullName, ex.Message));
                    }
                }

                bool canProcess;
                try
                {
                    canProcess = plugin.CanProcess(textForPlugin);
                }
                catch (Exception ex)
                {
                    AppLog.AppendDaily(AppLog.LevelErr, string.Format("插件 CanProcess 失败: Plugin={0}, LineKey={1}, Error={2}: {3}", loadedPlugin.DisplayName, lineKey, ex.GetType().FullName, ex.Message));
                    continue;
                }

                if (!canProcess)
                    continue;

                PluginProcessResult result;
                try
                {
                    result = plugin.TryProcess(textForPlugin);
                }
                catch (Exception ex)
                {
                    AppLog.AppendDaily(AppLog.LevelErr, string.Format("插件 TryProcess 失败: Plugin={0}, LineKey={1}, Error={2}: {3}", loadedPlugin.DisplayName, lineKey, ex.GetType().FullName, ex.Message));
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
                {
                    AppLog.AppendDaily(AppLog.LevelErr, string.Format("插件返回错误: Plugin={0}, LineKey={1}, Error={2}", loadedPlugin.DisplayName, lineKey, result.ErrorMessage));
                }

                if (result.Handled)
                {
                    if (result.Output != null)
                        currentText = result.Output;
                    break;
                }

                if (!string.IsNullOrEmpty(result.Output))
                    currentText = result.Output;
            }

            _lineCache[cacheKey] = currentText;
            return currentText;
        }

        /// <summary>
        /// 返回当前启用插件顺序签名。
        /// </summary>
        public string GetEnabledSignature()
        {
            if (_enabledPlugins.Count == 0)
                return string.Empty;
            return string.Join(">", _enabledPlugins.Select(x => x.DisplayName));
        }
    }
}
