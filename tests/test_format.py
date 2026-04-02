"""StyleFormat 模块测试."""

from __future__ import annotations

from sword import WordDocument, StyleFormat


class TestStyleFormat:
    """StyleFormat 测试类."""

    def test_create_style_format(self) -> None:
        """测试创建样式格式对象."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        assert fmt is not None

    def test_set_heading_font_name(self) -> None:
        """测试设置标题字体名称."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, name="宋体")

    def test_set_heading_font_size(self) -> None:
        """测试设置标题字体大小."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, size=16)

    def test_set_heading_font_bold(self) -> None:
        """测试设置标题加粗."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, bold=True)

    def test_set_heading_font_color(self) -> None:
        """测试设置标题颜色."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, color=(255, 0, 0))

    def test_set_heading_font_all(self) -> None:
        """测试设置标题所有字体属性."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, name="宋体", size=16, bold=True, color=(0, 0, 255))

    def test_set_heading_font_separate_chinese_english(self) -> None:
        """测试设置标题中英文字体分开."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, east_asia="宋体", ascii="Times New Roman", size=16)

    def test_set_heading_font_level2(self) -> None:
        """测试设置 Heading 2 字体格式."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(2, name="黑体", size=14, bold=True)

    def test_set_heading_font_level3(self) -> None:
        """测试设置 Heading 3 字体格式."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(3, name="楷体", size=12)

    def test_set_heading_font_only_chinese(self) -> None:
        """测试只设置中文字体."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, east_asia="宋体")

    def test_set_heading_font_only_english(self) -> None:
        """测试只设置英文字体."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, ascii="Times New Roman")

    def test_set_heading_font_with_h_ansi(self) -> None:
        """测试设置 hAnsi 字体."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_font(1, east_asia="宋体", ascii="Times New Roman", h_ansi="Arial")

    def test_set_heading_paragraph_alignment(self) -> None:
        """测试设置标题段落对齐方式."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_paragraph(1, alignment="center")

    def test_set_heading_paragraph_spacing(self) -> None:
        """测试设置标题段落间距."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_paragraph(1, space_before=12, space_after=6)

    def test_set_heading_paragraph_all(self) -> None:
        """测试设置标题所有段落属性."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_heading_paragraph(1, alignment="left", space_before=12, space_after=6)

    def test_set_normal_font(self) -> None:
        """测试设置 Normal 样式字体."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_normal_font(name="宋体", size=11)

    def test_set_normal_font_separate_chinese_english(self) -> None:
        """测试设置 Normal 中英文字体分开."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.set_normal_font(east_asia="宋体", ascii="Times New Roman", size=11)

    def test_enable_outline_level(self) -> None:
        """测试启用大纲级别."""
        doc = WordDocument()
        fmt = StyleFormat(doc._inner)
        fmt.enable_outline_level("Heading 1", 0)
        fmt.enable_outline_level("Heading 2", 1)
