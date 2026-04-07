using Xunit;

namespace SnakeTail.Tests;

public class TailFormSearchTests
{
    [Fact]
    public void Text_search_should_only_use_raw_text()
    {
        // 普通文本搜索只匹配原始日志内容。
        bool useProcessedText = TailForm.ShouldSearchProcessedText(lineHighlights: false);

        Assert.False(useProcessedText);
    }

    [Fact]
    public void Highlight_search_should_still_use_processed_text()
    {
        // 高亮搜索保持显示链路语义。
        bool useProcessedText = TailForm.ShouldSearchProcessedText(lineHighlights: true);

        Assert.True(useProcessedText);
    }
}
