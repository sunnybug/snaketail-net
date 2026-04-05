using System.Xml.Serialization;
using Xunit;

namespace SnakeTail.Tests;

public class TailFileConfigDefaultsTests
{
    [Fact]
    public void TailFileConfig_ctor_should_set_safe_polling_defaults()
    {
        // 新建配置应使用安全默认轮询，避免 UI 定时器 100ms 空转。
        var config = new TailFileConfig();

        Assert.Equal(TailFileConfig.DefaultFileCacheSize, config.FileCacheSize);
        Assert.Equal(TailFileConfig.DefaultFileCheckIntervalSeconds, config.FileCheckInterval);
        Assert.Equal(TailFileConfig.DefaultFileChangeCheckIntervalMs, config.FileChangeCheckInterval);
    }

    [Fact]
    public void TailFileConfig_xml_without_interval_fields_should_keep_defaults()
    {
        // 旧配置缺失轮询字段时，应保持构造默认值。
        const string xml = "<TailFileConfig />";
        var serializer = new XmlSerializer(typeof(TailFileConfig));
        using var reader = new StringReader(xml);
        var restored = (TailFileConfig)serializer.Deserialize(reader)!;

        Assert.Equal(TailFileConfig.DefaultFileCheckIntervalSeconds, restored.FileCheckInterval);
        Assert.Equal(TailFileConfig.DefaultFileChangeCheckIntervalMs, restored.FileChangeCheckInterval);
    }
}
