using System.Text.Json;
using System.Text.RegularExpressions;
using SnakeTail;

namespace LongMaidDisplayPlugin
{
    /// <summary>
    /// 龙女仆示例插件：把 skills: 数字 扩展为附加技能名。
    /// </summary>
    public sealed class DragonMaidDisplayPlugin : ILogDisplayPlugin
    {
        // 统一匹配 skills/passive_skill 两种键名。
        private static readonly Regex SkillRegex = new Regex(@"(?<key>skills|passive_skill):\s*(?<id>\d+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private readonly Dictionary<int, string> _skillMap = new Dictionary<int, string>();
        private string _jsonPath = string.Empty;

        public string Name => "龙女仆";

        /// <summary>
        /// 初始化并读取同目录 s_skill.json。
        /// </summary>
        public void Initialize(PluginContext context)
        {
            _skillMap.Clear();
            _jsonPath = Path.Combine(context.PluginDirectoryAbsolutePath, "s_skill.json");
            if (!File.Exists(_jsonPath))
                throw new FileNotFoundException("未找到 s_skill.json", _jsonPath);

            string jsonText = File.ReadAllText(_jsonPath);
            using JsonDocument document = JsonDocument.Parse(jsonText);
            if (!document.RootElement.TryGetProperty("s_skill", out JsonElement skillsElement) || skillsElement.ValueKind != JsonValueKind.Array)
                throw new InvalidOperationException("JSON 缺少 s_skill 数组");

            foreach (JsonElement row in skillsElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() < 2)
                    continue;

                JsonElement idElement = row[0];
                JsonElement nameElement = row[1];
                if (idElement.ValueKind != JsonValueKind.Number || nameElement.ValueKind != JsonValueKind.String)
                    continue;

                if (!idElement.TryGetInt32(out int skillId))
                    continue;

                string? skillName = nameElement.GetString();
                if (string.IsNullOrWhiteSpace(skillName))
                    continue;

                _skillMap[skillId] = skillName;
            }
        }

        /// <summary>
        /// 只处理包含 skills/passive_skill 数字 的行。
        /// </summary>
        public bool CanProcess(string line)
        {
            if (string.IsNullOrEmpty(line))
                return false;
            return SkillRegex.IsMatch(line);
        }

        /// <summary>
        /// 命中且可映射时输出 key: ID 名称；否则放行后续插件。
        /// </summary>
        public PluginProcessResult TryProcess(string line)
        {
            if (string.IsNullOrEmpty(line))
                return new PluginProcessResult { Handled = false, Output = line };

            Match match = SkillRegex.Match(line);
            if (!match.Success)
                return new PluginProcessResult { Handled = false, Output = line };

            string key = match.Groups["key"].Value;
            string idText = match.Groups["id"].Value;
            if (!int.TryParse(idText, out int skillId))
            {
                return new PluginProcessResult
                {
                    Handled = false,
                    Output = line,
                    ErrorMessage = key + " 数字解析失败: " + idText
                };
            }

            if (!_skillMap.TryGetValue(skillId, out string? skillName))
                return new PluginProcessResult { Handled = false, Output = line };

            // 保持原键名，仅追加技能名。
            string replacement = string.Format("{0}: {1} {2}", key, skillId, skillName);
            string output = SkillRegex.Replace(line, replacement, 1);
            return new PluginProcessResult { Handled = true, Output = output };
        }
    }
}
