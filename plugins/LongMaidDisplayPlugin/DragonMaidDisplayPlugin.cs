using System.Text.Json;
using System.Text.RegularExpressions;
using SnakeTail;

namespace LongMaidDisplayPlugin
{
    /// <summary>
    /// 龙女仆示例插件：把 skills: 数字 扩展为附加技能名。
    /// </summary>
    public sealed class DragonMaidDisplayPlugin : ILogDisplayPlugin, ILogDisplayBlockPlugin
    {
        // 统一匹配 skills/passive_skill/aura_skills 三种键名。
        private static readonly Regex SkillRegex = new Regex(@"(?<key>skills|passive_skill|aura_skills):\s*(?<id>\d+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        // 匹配 skill: [1,2,3] 这种技能列表。
        private static readonly Regex SkillListRegex = new Regex(@"(?<key>skill):\s*\[(?<ids>\s*\d+(?:\s*,\s*\d+)*)\s*\]", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        // 匹配战力属性块起始行。
        private static readonly Regex BattleEffectsBlockStartRegex = new Regex(@"^\s*attr_data=effects\s*\{\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        // 匹配战力属性块中允许出现的行。
        private static readonly Regex BattleEffectsBlockLineRegex = new Regex(@"^\s*(attr_data=effects|effects)\s*\{\s*$|^\s*key:\s*\d+\s*$|^\s*value:\s*\d+\s*$|^\s*\}\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        // 匹配下一条日志时间头，用于块边界切分。
        private static readonly Regex TimestampHeadRegex = new Regex(@"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}\t", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        // 匹配战力属性块中的 key 行。
        private static readonly Regex BattleEffectsKeyRegex = new Regex(@"(?m)^(?<indent>\s*key:\s*)(?<id>\d+)\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private readonly Dictionary<int, string> _skillMap = new Dictionary<int, string>();
        private readonly Dictionary<int, string> _battlePowerMap = new Dictionary<int, string>();
        private string _jsonPath = string.Empty;
        private string _battlePowerJsonPath = string.Empty;

        public string Name => "龙女仆";

        /// <summary>
        /// 初始化并读取同目录 s_skill.json。
        /// </summary>
        public void Initialize(PluginContext context)
        {
            _skillMap.Clear();
            _battlePowerMap.Clear();
            _jsonPath = Path.Combine(context.PluginDirectoryAbsolutePath, "s_skill.json");
            _battlePowerJsonPath = Path.Combine(context.PluginDirectoryAbsolutePath, "s_battle_power.json");
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

            // s_battle_power.json 为可选配置，存在时启用 effects 块 key 映射。
            if (File.Exists(_battlePowerJsonPath))
            {
                string battleJsonText = File.ReadAllText(_battlePowerJsonPath);
                using JsonDocument battleDocument = JsonDocument.Parse(battleJsonText);
                if (!battleDocument.RootElement.TryGetProperty("s_battle_power", out JsonElement battlePowerElement) || battlePowerElement.ValueKind != JsonValueKind.Array)
                    throw new InvalidOperationException("s_battle_power.json 缺少 s_battle_power 数组");

                foreach (JsonElement row in battlePowerElement.EnumerateArray())
                {
                    if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() < 6)
                        continue;

                    JsonElement idElement = row[0];
                    JsonElement nameElement = row[5];
                    if (idElement.ValueKind != JsonValueKind.Number || nameElement.ValueKind != JsonValueKind.String)
                        continue;

                    if (!idElement.TryGetInt32(out int battlePowerId))
                        continue;

                    string? battlePowerName = nameElement.GetString();
                    if (string.IsNullOrWhiteSpace(battlePowerName))
                        continue;

                    _battlePowerMap[battlePowerId] = battlePowerName;
                }
            }
        }

        /// <summary>
        /// 命中 attr_data=effects 起始行时，向后收集完整 effects 属性块。
        /// </summary>
        public bool TryCollectBlock(int lineKey, string currentLine, Func<int, string> readLineByLineKey, out string blockText)
        {
            blockText = string.Empty;
            if (string.IsNullOrEmpty(currentLine) || !BattleEffectsBlockStartRegex.IsMatch(currentLine))
                return false;

            const int maxBlockLines = 600;
            List<string> lines = new List<string>(64) { currentLine };
            for (int offset = 1; offset < maxBlockLines; offset++)
            {
                string nextLine = readLineByLineKey(lineKey + offset);
                if (string.IsNullOrEmpty(nextLine))
                    break;

                if (TimestampHeadRegex.IsMatch(nextLine))
                    break;

                if (!BattleEffectsBlockLineRegex.IsMatch(nextLine))
                    break;

                lines.Add(nextLine);
            }

            if (lines.Count <= 1)
                return false;

            blockText = string.Join(Environment.NewLine, lines);
            return true;
        }

        /// <summary>
        /// 只处理包含 skills/passive_skill/aura_skills 数字 的行。
        /// </summary>
        public bool CanProcess(string line)
        {
            if (string.IsNullOrEmpty(line))
                return false;

            // 兼容单行技能映射与多行 effects 块映射两类输入。
            return SkillRegex.IsMatch(line)
                || SkillListRegex.IsMatch(line)
                || (line.Contains(Environment.NewLine) && BattleEffectsKeyRegex.IsMatch(line));
        }

        /// <summary>
        /// 命中且可映射时输出 key: ID 名称；否则放行后续插件。
        /// </summary>
        public PluginProcessResult TryProcess(string line)
        {
            if (string.IsNullOrEmpty(line))
                return new PluginProcessResult { Handled = false, Output = line };

            // 多行 effects 块：把 key: 数字 映射为 key: 数字 名称。
            if (line.Contains(Environment.NewLine) && BattleEffectsKeyRegex.IsMatch(line))
            {
                string blockOutput = BattleEffectsKeyRegex.Replace(line, match =>
                {
                    string idText = match.Groups["id"].Value;
                    if (!int.TryParse(idText, out int keyId))
                        return match.Value;

                    if (!_battlePowerMap.TryGetValue(keyId, out string? keyName))
                        return match.Value;

                    string indent = match.Groups["indent"].Value;
                    return string.Format("{0}{1} {2}", indent, keyId, keyName);
                });

                bool changed = !string.Equals(blockOutput, line, StringComparison.Ordinal);
                return new PluginProcessResult
                {
                    Handled = changed,
                    Output = blockOutput,
                    ErrorMessage = !changed && _battlePowerMap.Count == 0 ? "未加载 s_battle_power 映射，effects key 名称无法扩展" : string.Empty
                };
            }

            // skill: [数字,数字]：逐个 ID 映射并在原位置追加技能名。
            Match listMatch = SkillListRegex.Match(line);
            if (listMatch.Success)
            {
                string idsText = listMatch.Groups["ids"].Value;
                string[] idTokens = idsText.Split(',');
                bool changed = false;
                string[] mappedTokens = new string[idTokens.Length];
                for (int i = 0; i < idTokens.Length; i++)
                {
                    string token = idTokens[i].Trim();
                    if (!int.TryParse(token, out int listSkillId))
                    {
                        mappedTokens[i] = token;
                        continue;
                    }

                    if (_skillMap.TryGetValue(listSkillId, out string? listSkillName))
                    {
                        mappedTokens[i] = string.Format("{0} {1}", listSkillId, listSkillName);
                        changed = true;
                        continue;
                    }

                    mappedTokens[i] = token;
                }

                if (!changed)
                    return new PluginProcessResult { Handled = false, Output = line };

                string listKey = listMatch.Groups["key"].Value;
                string listReplacement = string.Format("{0}: [{1}]", listKey, string.Join(",", mappedTokens));
                string listOutput = SkillListRegex.Replace(line, listReplacement, 1);
                return new PluginProcessResult { Handled = true, Output = listOutput };
            }

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
