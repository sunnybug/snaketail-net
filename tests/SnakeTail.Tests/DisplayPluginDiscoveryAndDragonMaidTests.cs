using System.Reflection;
using LongMaidDisplayPlugin;
using Xunit;

namespace SnakeTail.Tests;

public class DisplayPluginDiscoveryAndDragonMaidTests
{
    [Fact]
    public void TryCreatePluginInstance_should_find_valid_plugin_type()
    {
        // 验证可从包含实现类型的程序集创建插件实例
        bool ok = DisplayPluginManager.TryCreatePluginInstance(Assembly.GetExecutingAssembly(), out ILogDisplayPlugin? plugin, out string? error);

        Assert.True(ok);
        Assert.NotNull(plugin);
        Assert.True(string.IsNullOrEmpty(error));
    }

    [Fact]
    public void TryLoadPluginDirectory_should_ignore_invalid_dll_and_accept_valid_plugin()
    {
        // 验证目录扫描可忽略非法 DLL，并识别合法插件实现
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-scan-" + Guid.NewGuid().ToString("N"));
        string pluginRoot = Path.Combine(tempRoot, "config", "plugins");
        string invalidDir = Path.Combine(pluginRoot, "invalid");
        string validDir = Path.Combine(pluginRoot, "valid");
        Directory.CreateDirectory(invalidDir);
        Directory.CreateDirectory(validDir);
        try
        {
            File.WriteAllText(Path.Combine(invalidDir, "broken.dll"), "not a dll");
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            File.Copy(assemblyPath, Path.Combine(validDir, Path.GetFileName(assemblyPath)), overwrite: true);

            var manager = new DisplayPluginManager(tempRoot);
            bool invalidLoaded = manager.TryLoadPluginDirectory(invalidDir, "invalid", string.Empty, out LoadedDisplayPlugin? _);
            bool validLoaded = manager.TryLoadPluginDirectory(validDir, "valid", string.Empty, out LoadedDisplayPlugin? loadedPlugin);

            Assert.False(invalidLoaded);
            Assert.True(validLoaded);
            Assert.NotNull(loadedPlugin);
            Assert.Equal("TestAssemblyPlugin", loadedPlugin.DisplayName);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void TryLoadPluginDirectory_should_keep_plugin_visible_when_required_config_missing()
    {
        // 验证插件缺少必需配置时仍可被发现，并返回精确原因供 UI 提示。
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-missing-config-" + Guid.NewGuid().ToString("N"));
        string pluginRoot = Path.Combine(tempRoot, "config", "plugins");
        string pluginDir = Path.Combine(pluginRoot, "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string sourceAssemblyPath = typeof(DragonMaidDisplayPlugin).Assembly.Location;
            File.Copy(sourceAssemblyPath, Path.Combine(pluginDir, Path.GetFileName(sourceAssemblyPath)), overwrite: true);

            var manager = new DisplayPluginManager(tempRoot);
            bool loaded = manager.TryLoadPluginDirectory(pluginDir, "龙女仆", string.Empty, out LoadedDisplayPlugin? loadedPlugin);

            Assert.True(loaded);
            Assert.NotNull(loadedPlugin);
            Assert.Equal("龙女仆", loadedPlugin.DisplayName);
            Assert.False(loadedPlugin.IsAvailable);
            Assert.Contains("未找到 s_skill.json", loadedPlugin.UnavailableReason);
            Assert.Contains(Path.Combine(pluginDir, "s_skill.json"), loadedPlugin.UnavailableReason);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_known_id()
    {
        // 验证龙女仆插件对 skills 命中技能 ID 的替换效果
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[10009001,"圣剑斩"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult knownResult = plugin.TryProcess("abc skills: 10009001 def");
            PluginProcessResult unknownResult = plugin.TryProcess("abc skills: 99999999 def");

            Assert.True(knownResult.Handled);
            Assert.Equal("abc skills: 10009001 圣剑斩 def", knownResult.Output);
            Assert.False(unknownResult.Handled);
            Assert.Equal("abc skills: 99999999 def", unknownResult.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_known_passive_skill_id()
    {
        // 验证龙女仆插件对 passive_skill 命中技能 ID 的替换效果
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[99501001,"测试被动1"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult knownResult = plugin.TryProcess("abc passive_skill: 99501001 def");

            Assert.True(knownResult.Handled);
            Assert.Equal("abc passive_skill: 99501001 测试被动1 def", knownResult.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_known_aura_skills_id()
    {
        // 验证龙女仆插件对 aura_skills 命中技能 ID 的替换效果
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[20244002,"光环测试"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult knownResult = plugin.TryProcess("abc aura_skills: 20244002 def");

            Assert.True(knownResult.Handled);
            Assert.Equal("abc aura_skills: 20244002 光环测试 def", knownResult.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_keep_passive_skill_zero_unchanged()
    {
        // 验证 passive_skill: 0 无映射时保持原样
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[99501001,"测试被动1"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult zeroResult = plugin.TryProcess("abc passive_skill: 0 def");

            Assert.False(zeroResult.Handled);
            Assert.Equal("abc passive_skill: 0 def", zeroResult.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_skill_list()
    {
        // 验证 skill: [id,...] 列表可逐个映射并追加技能名。
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[2001,"斩击"],[4001,"暴击"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult result = plugin.TryProcess("abc skill: [2001,3002,4001,10001041] def");

            Assert.True(result.Handled);
            Assert.Equal("abc skill: [2001 斩击,3002,4001 暴击,10001041] def", result.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_quoted_skills_list()
    {
        // 验证 "skills": [id,...] 单行数组可逐个映射并追加技能名。
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[10013001,"龙息"],[10007001,"横扫"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            PluginProcessResult result = plugin.TryProcess("""abc "skills": [10013001,10007001,10001001] def""");

            Assert.True(result.Handled);
            Assert.Equal("""abc "skills": [10013001 龙息,10007001 横扫,10001001] def""", result.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_append_skill_name_for_multiline_skills_block()
    {
        // 验证 "skills": [ ... ] 多行数组可逐行映射并追加技能名。
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            string jsonPath = Path.Combine(pluginDir, "s_skill.json");
            File.WriteAllText(jsonPath, """{"s_skill":[[10013001,"龙息"],[10007001,"横扫"]] }""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            var lines = new Dictionary<int, string>
            {
                [300] = "\"skills\": [",
                [301] = "  10013001,",
                [302] = "  10007001,",
                [303] = "  10001001",
                [304] = "],",
                [305] = "\"other\": 1"
            };

            bool collected = ((ILogDisplayBlockPlugin)plugin).TryCollectBlock(
                300,
                lines[300],
                lineKey => lines.TryGetValue(lineKey, out string? text) ? text : string.Empty,
                out string blockText);

            Assert.True(collected);
            PluginProcessResult result = plugin.TryProcess(blockText);
            Assert.True(result.Handled);
            string[] outputLines = result.Output.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            Assert.Equal("\"skills\": [", outputLines[0]);
            Assert.Equal("  10013001 龙息,", outputLines[1]);
            Assert.Equal("  10007001 横扫,", outputLines[2]);
            Assert.Equal("  10001001", outputLines[3]);
            Assert.Equal("],", outputLines[4]);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    [Fact]
    public void DragonMaid_should_expand_battle_effect_key_name_for_multiline_block()
    {
        // 验证 effects 多行块可按 key 映射扩展名称。
        string tempRoot = Path.Combine(Path.GetTempPath(), "snaketail-plugin-test-" + Guid.NewGuid().ToString("N"));
        string pluginDir = Path.Combine(tempRoot, "config", "plugins", "龙女仆");
        Directory.CreateDirectory(pluginDir);
        try
        {
            File.WriteAllText(Path.Combine(pluginDir, "s_skill.json"), """{"s_skill":[[10009001,"圣剑斩"]] }""");
            File.WriteAllText(Path.Combine(pluginDir, "s_battle_power.json"), """{"s_battle_power":[[1,1,1,0,0,"声明","",1],[2,1,1,0,0,"攻击","",1]]}""");

            var plugin = new DragonMaidDisplayPlugin();
            var context = new PluginContext(pluginDir, Path.Combine(tempRoot, "config"), Path.Combine(tempRoot, "fake.log"), "1.0.0");
            plugin.Initialize(context);

            var lines = new Dictionary<int, string>
            {
                [200] = "attr_data=effects {",
                [201] = "  key: 1",
                [202] = "  value: 7153731",
                [203] = "}",
                [204] = "effects {",
                [205] = "  key: 2",
                [206] = "  value: 46304",
                [207] = "}",
                [208] = "2026-04-05 15:50:15.208\tyyy"
            };

            bool collected = ((ILogDisplayBlockPlugin)plugin).TryCollectBlock(
                200,
                lines[200],
                lineKey => lines.TryGetValue(lineKey, out string? text) ? text : string.Empty,
                out string blockText);

            Assert.True(collected);
            PluginProcessResult result = plugin.TryProcess(blockText);
            Assert.True(result.Handled);
            Assert.Contains("key: 1 声明", result.Output);
            Assert.Contains("key: 2 攻击", result.Output);
        }
        finally
        {
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    public sealed class TestAssemblyPlugin : ILogDisplayPlugin
    {
        public string Name => "TestAssemblyPlugin";

        public void Initialize(PluginContext context)
        {
            // 测试插件无需初始化逻辑
        }

        public bool CanProcess(string line)
        {
            return false;
        }

        public PluginProcessResult TryProcess(string line)
        {
            return new PluginProcessResult { Handled = false, Output = line };
        }
    }
}
