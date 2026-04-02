"""WordTable 模块测试."""

from __future__ import annotations

import tempfile
import os

from sword import WordDocument, WordTable


class TestWordTable:
    """WordTable 测试类."""

    def test_create_table(self) -> None:
        """测试创建表格."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3, style="Table Grid")
        assert len(table.table.rows) == 3
        assert len(table.table.columns) == 3

    def test_set_style(self) -> None:
        """测试设置表格样式."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3)
        table.set_style("Table Grid")
        assert table.table.style.name == "Table Grid"

    def test_cell(self) -> None:
        """测试获取单元格."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3)
        cell = table.cell(0, 0)
        cell.text = "测试内容"
        assert cell.text == "测试内容"

    def test_context_manager(self) -> None:
        """测试上下文管理器."""
        doc = WordDocument()
        with WordTable(doc._inner, rows=3, cols=3) as table:
            table.cell(0, 0).text = "测试内容"
            assert table.cell(0, 0).text == "测试内容"

    def test_save_table_document(
        self, tmp_path: tempfile.TemporaryDirectory
    ) -> None:
        """测试保存包含表格的文档."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3, style="Table Grid")
        table.cell(0, 0).text = "标题"

        file_path = os.path.join(tmp_path, "table_test.docx")
        doc.save(file_path)
        assert os.path.exists(file_path)


class TestWordTableFromSection:
    """从章节创建表格测试类."""

    def test_add_table_from_section(self) -> None:
        """测试从章节添加表格."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)
        section = doc.add_section("第一章")
        table = section.add_table(rows=3, cols=3, style="Table Grid")
        assert len(table.table.rows) == 3
        assert len(table.table.columns) == 3

    def test_add_table_and_set_text(self) -> None:
        """测试添加表格并设置文本."""
        doc = WordDocument()
        section = doc.add_section("第一章")
        table = section.add_table(rows=3, cols=3)
        table.set_style("Table Grid")
        table.cell(0, 0).text = "姓名"
        table.cell(0, 1).text = "年龄"
        table.cell(0, 2).text = "城市"
        table.cell(1, 0).text = "张三"
        table.cell(1, 1).text = "25"
        table.cell(1, 2).text = "北京"
