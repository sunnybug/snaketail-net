using Xunit;

namespace SnakeTail.Tests;

public class DisplayTextProcessorTests
{
    [Fact]
    public void Processor_should_stop_when_previous_plugin_handled()
    {
        // 验证前一个插件 Handled=true 时后续插件不执行
        var firstPlugin = new CountingPlugin("P1", canProcess: true, handled: true, output: "first");
        var secondPlugin = new CountingPlugin("P2", canProcess: true, handled: true, output: "second");
        var processor = new DisplayTextProcessor();
        processor.SetEnabledPlugins(
        [
            new LoadedDisplayPlugin("A", "A", "P1", firstPlugin),
            new LoadedDisplayPlugin("B", "B", "P2", secondPlugin)
        ]);

        string output = processor.GetProcessedLineText(10, "raw");

        Assert.Equal("first", output);
        Assert.Equal(1, firstPlugin.TryProcessCallCount);
        Assert.Equal(0, secondPlugin.TryProcessCallCount);
    }

    [Fact]
    public void Processor_should_hit_cache_with_same_signature()
    {
        // 验证同一行与同一插件签名下命中缓存
        var plugin = new CountingPlugin("P1", canProcess: true, handled: true, output: "cached");
        var processor = new DisplayTextProcessor();
        processor.SetEnabledPlugins([new LoadedDisplayPlugin("A", "A", "P1", plugin)]);

        string output1 = processor.GetProcessedLineText(20, "raw");
        string output2 = processor.GetProcessedLineText(20, "raw");

        Assert.Equal("cached", output1);
        Assert.Equal("cached", output2);
        Assert.Equal(1, plugin.TryProcessCallCount);

        // 修改插件顺序签名后，缓存应失效并重新计算
        processor.SetEnabledPlugins([]);
        processor.SetEnabledPlugins([new LoadedDisplayPlugin("A", "A", "P1", plugin)]);
        _ = processor.GetProcessedLineText(20, "raw");
        Assert.Equal(2, plugin.TryProcessCallCount);
    }

    private sealed class CountingPlugin : ILogDisplayPlugin
    {
        private readonly bool _canProcess;
        private readonly bool _handled;
        private readonly string _output;

        public CountingPlugin(string name, bool canProcess, bool handled, string output)
        {
            Name = name;
            _canProcess = canProcess;
            _handled = handled;
            _output = output;
        }

        public string Name { get; }
        public int TryProcessCallCount { get; private set; }

        public void Initialize(PluginContext context)
        {
            // 测试插件无需初始化逻辑
        }

        public bool CanProcess(string line)
        {
            return _canProcess;
        }

        public PluginProcessResult TryProcess(string line)
        {
            TryProcessCallCount++;
            return new PluginProcessResult { Handled = _handled, Output = _output };
        }
    }
}
