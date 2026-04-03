"""Word 文档单元格封装模块."""

from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class WordCell:
    """封装 Word 文档中的单元格，提供便捷的单元格操作."""

    def __init__(self, cell) -> None:
        """
        初始化单元格.

        Args:
            cell: 底层的 python-docx Cell 对象.
        """
        self._cell = cell

    @property
    def inner(self):
        """获取底层的 python-docx Cell 对象."""
        return self._cell

    @property
    def text(self) -> str:
        """获取单元格文本."""
        return self._cell.text

    @text.setter
    def text(self, value: str) -> None:
        """设置单元格文本."""
        self._cell.text = value

    def _ensure_tcPr(self) -> OxmlElement:
        """确保 tcPr 元素存在."""
        tc = self._cell._element
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        return tcPr

    def _ensure_elem(self, parent: OxmlElement, tag: str) -> OxmlElement:
        """确保子元素存在，不存在则创建."""
        elem = parent.find(qn(f"w:{tag}"))
        if elem is None:
            elem = OxmlElement(f"w:{tag}")
            parent.append(elem)
        return elem

    def set_shading(self, fill: str) -> None:
        """
        设置单元格底纹/背景色.

        Args:
            fill: 填充颜色（十六进制字符串，如 "FF0000"）。
        """
        tcPr = self._ensure_tcPr()
        shd = self._ensure_elem(tcPr, "shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)

    def set_borders(
        self,
        top: str | None = None,
        bottom: str | None = None,
        left: str | None = None,
        right: str | None = None,
        border_size: int = 4,
        border_color: str = "000000",
    ) -> None:
        """
        设置单元格边框.

        Args:
            top: 上边框样式（"single", "double", "none"）。
            bottom: 下边框样式.
            left: 左边框样式.
            right: 右边框样式.
            border_size: 边框大小（磅）。
            border_color: 边框颜色（十六进制字符串）。
        """
        tcPr = self._ensure_tcPr()
        tcBorders = self._ensure_elem(tcPr, "tcBorders")

        border_attrs = {
            "top": top,
            "bottom": bottom,
            "left": left,
            "right": right,
        }

        for border_name, border_style in border_attrs.items():
            elem = self._ensure_elem(tcBorders, border_name)
            if border_style is not None:
                elem.set(qn("w:val"), border_style)
                elem.set(qn("w:sz"), str(border_size * 8))
                elem.set(qn("w:space"), "0")
                elem.set(qn("w:color"), border_color)

    def set_margins(
        self,
        top: int | None = None,
        bottom: int | None = None,
        left: int | None = None,
        right: int | None = None,
    ) -> None:
        """
        设置单元格页边距.

        Args:
            top: 上边距（磅）。
            bottom: 下边距（磅）。
            left: 左边距（磅）。
            right: 右边距（磅）。
        """
        tcPr = self._ensure_tcPr()
        tcMar = self._ensure_elem(tcPr, "tcMar")

        if top is not None:
            elem = self._ensure_elem(tcMar, "top")
            elem.set(qn("w:w"), str(int(top * 20)))
            elem.set(qn("w:type"), "dxa")
        if bottom is not None:
            elem = self._ensure_elem(tcMar, "bottom")
            elem.set(qn("w:w"), str(int(bottom * 20)))
            elem.set(qn("w:type"), "dxa")
        if left is not None:
            elem = self._ensure_elem(tcMar, "left")
            elem.set(qn("w:w"), str(int(left * 20)))
            elem.set(qn("w:type"), "dxa")
        if right is not None:
            elem = self._ensure_elem(tcMar, "right")
            elem.set(qn("w:w"), str(int(right * 20)))
            elem.set(qn("w:type"), "dxa")

    def set_vertical_alignment(self, alignment: str = "center") -> None:
        """
        设置单元格垂直对齐方式.

        Args:
            alignment: 对齐方式（"top", "center", "bottom"）。
        """
        vAlign_map = {
            "top": "top",
            "center": "center",
            "bottom": "bottom",
        }
        if alignment not in vAlign_map:
            return

        tcPr = self._ensure_tcPr()
        vAlign = self._ensure_elem(tcPr, "vAlign")
        vAlign.set(qn("w:val"), vAlign_map[alignment])

    def set_horizontal_alignment(self, alignment: str = "left") -> None:
        """
        设置单元格水平对齐方式.

        Args:
            alignment: 对齐方式（"left", "center", "right", "justify"）。
        """
        jc_map = {
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "both",
        }
        if alignment not in jc_map:
            return

        tcPr = self._ensure_tcPr()
        pPr = self._ensure_elem(tcPr, "pPr")
        jc = self._ensure_elem(pPr, "jc")
        jc.set(qn("w:val"), jc_map[alignment])

    def set_width(self, width: int, unit: str = "auto") -> None:
        """
        设置单元格宽度.

        Args:
            width: 宽度值.
            unit: 单位（"auto", "dxa"（twips）, "pct"（百分之一百分比））。
        """
        tcPr = self._ensure_tcPr()
        tcW = self._ensure_elem(tcPr, "tcW")

        if unit == "auto":
            tcW.set(qn("w:w"), "0")
            tcW.set(qn("w:type"), "auto")
        elif unit == "dxa":
            tcW.set(qn("w:w"), str(width))
            tcW.set(qn("w:type"), "dxa")
        elif unit == "pct":
            tcW.set(qn("w:w"), str(width))
            tcW.set(qn("w:type"), "pct")
        else:
            tcW.set(qn("w:w"), str(width))
            tcW.set(qn("w:type"), "dxa")

    def __enter__(self) -> WordCell:
        return self

    def __exit__(self, *args: Any) -> None:
        pass
