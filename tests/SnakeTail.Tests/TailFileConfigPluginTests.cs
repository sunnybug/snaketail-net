using System.Xml.Serialization;
using Xunit;

namespace SnakeTail.Tests;

public class TailFileConfigPluginTests
{
    [Fact]
    public void EnabledDisplayPlugins_xml_roundtrip_should_keep_order()
    {
        // 验证插件启用顺序可被 XML 正确保存与恢复
        var config = new TailFileConfig
        {
            EnabledDisplayPlugins = new List<string> { "龙女仆", "示例插件B" }
        };

        var serializer = new XmlSerializer(typeof(TailFileConfig));
        using var writer = new StringWriter();
        serializer.Serialize(writer, config);
        string xml = writer.ToString();

        using var reader = new StringReader(xml);
        var restored = (TailFileConfig)serializer.Deserialize(reader)!;

        Assert.NotNull(restored.EnabledDisplayPlugins);
        Assert.Equal(2, restored.EnabledDisplayPlugins.Count);
        Assert.Equal("龙女仆", restored.EnabledDisplayPlugins[0]);
        Assert.Equal("示例插件B", restored.EnabledDisplayPlugins[1]);
    }
}
