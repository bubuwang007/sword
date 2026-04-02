"""StyleFormat 模块测试."""

from __future__ import annotations

import tempfile
import os

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


class TestTableFormatting:
    """表格格式测试类（通过样式设置）."""

    def test_create_style(self) -> None:
        """测试创建样式."""
        doc = WordDocument()
        style = doc.format.create_style("自定义段落样式")
        assert style is not None
        assert style.name == "自定义段落样式"

    def test_create_style_character(self) -> None:
        """测试创建字符样式."""
        doc = WordDocument()
        style = doc.format.create_style("自定义字符样式", style_type="character")
        assert style is not None
        assert style.name == "自定义字符样式"

    def test_create_style_list(self) -> None:
        """测试创建列表样式."""
        doc = WordDocument()
        style = doc.format.create_style("自定义列表样式", style_type="list")
        assert style is not None
        assert style.name == "自定义列表样式"

    def test_create_style_based_on(self) -> None:
        """测试创建基于现有样式的样式."""
        doc = WordDocument()
        style = doc.format.create_style("自定义样式", based_on="Normal")
        assert style is not None
        assert style.base_style.name == "Normal"

    def test_create_table_style(self) -> None:
        """测试创建表格样式."""
        doc = WordDocument()
        style = doc.format.create_table_style("自定义表格样式")
        assert style is not None
        assert style.name == "自定义表格样式"

    def test_create_table_style_based_on(self) -> None:
        """测试创建基于现有样式的表格样式."""
        doc = WordDocument()
        style = doc.format.create_table_style("自定义样式", based_on="Table Grid")
        assert style is not None
        assert style.base_style.name == "Table Grid"

    def test_create_table_style_exists(self) -> None:
        """测试创建已存在的表格样式返回现有样式."""
        doc = WordDocument()
        style1 = doc.format.create_table_style("测试样式")
        style2 = doc.format.create_table_style("测试样式")
        assert style1.name == style2.name

    def test_set_table_style(self) -> None:
        """测试设置表格样式."""
        doc = WordDocument()
        table = doc._inner.add_table(rows=3, cols=3)
        doc.format.set_table_style(table, "Table Grid")
        assert table.style.name == "Table Grid"

    def test_set_table_font(self) -> None:
        """测试设置表格样式字体."""
        doc = WordDocument()
        doc.format.set_table_font("Table Grid", name="宋体", size=12, bold=True)

    def test_set_table_shading(self) -> None:
        """测试设置表格样式底纹."""
        doc = WordDocument()
        doc.format.set_table_shading("Table Grid", fill="FFFF00")

    def test_set_table_borders(self) -> None:
        """测试设置表格样式边框."""
        doc = WordDocument()
        doc.format.set_table_borders(
            "Table Grid",
            top="single",
            bottom="single",
            left="single",
            right="single",
            inside_h="single",
            inside_v="single",
            border_size=8,
            border_color="000000",
        )

    def test_set_table_borders_partial(self) -> None:
        """测试部分设置表格边框."""
        doc = WordDocument()
        doc.format.set_table_borders("Table Grid", top="single")

    def test_set_table_margins(self) -> None:
        """测试设置表格外边距."""
        doc = WordDocument()
        doc.format.set_table_margins("Table Grid", top=100, bottom=100, left=200, right=200)

    def test_set_table_cell_margins(self) -> None:
        """测试设置表格单元格内边距."""
        doc = WordDocument()
        doc.format.set_table_cell_margins("Table Grid", top=100, bottom=100, left=200, right=200)

    def test_set_table_alignment(self) -> None:
        """测试设置表格对齐方式."""
        doc = WordDocument()
        doc.format.set_table_alignment("Table Grid", alignment="center")

    def test_set_table_paragraph(self) -> None:
        """测试设置表格段落格式."""
        doc = WordDocument()
        doc.format.set_table_paragraph(
            "Table Grid",
            alignment="center",
            space_before=0,
            space_after=6,
        )

    def test_table_style_all_format(self) -> None:
        """测试设置表格样式所有格式."""
        doc = WordDocument()
        doc.format.set_table_font("Table Grid", name="宋体", size=12, bold=True)
        doc.format.set_table_shading("Table Grid", fill="E6E6E6")
        doc.format.set_table_borders("Table Grid", top="single", bottom="single")
        doc.format.set_table_margins("Table Grid", left=100, right=100)
        doc.format.set_table_cell_margins("Table Grid", left=50, right=50)
        doc.format.set_table_alignment("Table Grid", alignment="center")
        doc.format.set_table_paragraph("Table Grid", alignment="center")

    def test_save_table_document(
        self, tmp_path: tempfile.TemporaryDirectory
    ) -> None:
        """测试保存包含表格的文档."""
        doc = WordDocument()
        table = doc._inner.add_table(rows=3, cols=3)
        doc.format.set_table_style(table, "Table Grid")
        table.cell(0, 0).text = "标题"

        file_path = os.path.join(tmp_path, "table_test.docx")
        doc.save(file_path)
        assert os.path.exists(file_path)
