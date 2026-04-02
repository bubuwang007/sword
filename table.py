"""Word 文档表格封装模块."""

from __future__ import annotations

from typing import Any

from docx.table import Table
from docx.document import Document as DocxDocument


class WordTable:
    """封装 Word 文档中的表格，提供便捷的表格操作."""

    def __init__(
        self,
        document: DocxDocument,
        rows: int,
        cols: int,
        style: str | None = None,
    ) -> None:
        """
        初始化表格.

        Args:
            document: 底层的 python-docx Document 对象.
            rows: 表格行数.
            cols: 表格列数.
            style: 表格样式（如 "Table Grid"）。
        """
        self._doc = document
        self._table = document.add_table(rows=rows, cols=cols)
        if style is not None:
            self._table.style = style

    @property
    def table(self) -> Table:
        """获取底层的 python-docx Table 对象."""
        return self._table

    def cell(self, row: int, col: int):
        """
        获取单元格.

        Args:
            row: 行索引（从 0 开始）。
            col: 列索引（从 0 开始）。

        Returns:
            单元格对象。
        """
        return self._table.cell(row, col)

    def set_style(self, style: str | None) -> None:
        """
        设置表格样式.

        Args:
            style: 样式名称（None 清除样式）。
        """
        if style is None:
            self._table.style = None
        else:
            self._table.style = style

    def __enter__(self) -> WordTable:
        return self

    def __exit__(self, *args: Any) -> None:
        pass
