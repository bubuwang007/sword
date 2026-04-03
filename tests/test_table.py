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

    def test_rows_property(self) -> None:
        """测试 rows 属性."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=4, cols=3)
        assert table.rows == 4

    def test_cols_property(self) -> None:
        """测试 cols 属性."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=5)
        assert table.cols == 5

    def test_iter_cells(self) -> None:
        """测试 iter_cells 方法."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=2, cols=2)
        table.cell(0, 0).text = "A1"
        table.cell(0, 1).text = "A2"
        table.cell(1, 0).text = "B1"
        table.cell(1, 1).text = "B2"

        cells = list(table.iter_cells())
        assert len(cells) == 4
        texts = [cell.text for cell in cells]
        assert texts == ["A1", "A2", "B1", "B2"]

    def test_iter_rows(self) -> None:
        """测试 iter_rows 方法."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=2, cols=2)
        table.cell(0, 0).text = "A1"
        table.cell(0, 1).text = "A2"
        table.cell(1, 0).text = "B1"
        table.cell(1, 1).text = "B2"

        rows = list(table.iter_rows())
        assert len(rows) == 2
        assert [cell.text for cell in rows[0]] == ["A1", "A2"]
        assert [cell.text for cell in rows[1]] == ["B1", "B2"]

    def test_iter_cols(self) -> None:
        """测试 iter_cols 方法."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=2, cols=2)
        table.cell(0, 0).text = "A1"
        table.cell(0, 1).text = "A2"
        table.cell(1, 0).text = "B1"
        table.cell(1, 1).text = "B2"

        cols = list(table.iter_cols())
        assert len(cols) == 2
        assert [cell.text for cell in cols[0]] == ["A1", "B1"]
        assert [cell.text for cell in cols[1]] == ["A2", "B2"]

    def test_get_row(self) -> None:
        """测试 get_row 方法."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3)
        table.cell(1, 0).text = "R1C0"
        table.cell(1, 1).text = "R1C1"
        table.cell(1, 2).text = "R1C2"

        row = table.get_row(1)
        assert len(row) == 3
        assert [cell.text for cell in row] == ["R1C0", "R1C1", "R1C2"]

    def test_get_column(self) -> None:
        """测试 get_column 方法."""
        doc = WordDocument()
        table = WordTable(doc._inner, rows=3, cols=3)
        table.cell(0, 1).text = "R0C1"
        table.cell(1, 1).text = "R1C1"
        table.cell(2, 1).text = "R2C1"

        col = table.get_column(1)
        assert len(col) == 3
        assert [cell.text for cell in col] == ["R0C1", "R1C1", "R2C1"]

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
