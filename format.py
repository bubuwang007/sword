"""Word 文档格式样式封装模块."""

from __future__ import annotations

from docx.document import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class StyleFormat:
    """Word 文档样式格式封装类."""

    def __init__(self, document: DocxDocument) -> None:
        """
        初始化样式格式.

        Args:
            document: 底层的 python-docx Document 对象.
        """
        self._doc = document

    def _set_font_rpr(self, font, ascii: str | None = None, east_asia: str | None = None, h_ansi: str | None = None) -> None:
        """
        内部方法：设置字体（支持中英文分开）.

        Args:
            font: 字体对象.
            ascii: 西文字体名称.
            east_asia: 东亚字体名称（中文）.
            h_ansi: 高 ANSI 字体名称.
        """
        rPr = font._element
        if rPr is None:
            return

        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)

        if ascii is not None:
            rFonts.set(qn("w:ascii"), ascii)
        if east_asia is not None:
            rFonts.set(qn("w:eastAsia"), east_asia)
        if h_ansi is not None:
            rFonts.set(qn("w:hAnsi"), h_ansi)

    def set_heading_font(
        self,
        level: int,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        color: tuple[int, int, int] | None = None,
        east_asia: str | None = None,
        ascii: str | None = None,
        h_ansi: str | None = None,
    ) -> None:
        """
        设置标题样式字体格式.

        Args:
            level: 标题级别（1-9）.
            name: 字体名称（同时设置中英文）.
            size: 字体大小（磅）.
            bold: 是否加粗.
            color: RGB 颜色元组（如 (255, 0, 0)）.
            east_asia: 中文字体名称（如 "宋体"）.
            ascii: 西文字体名称（如 "Times New Roman"）.
            h_ansi: 高 ANSI 字体名称.
        """
        style_name = f"Heading {level}"
        if style_name not in [s.name for s in self._doc.styles]:
            return

        style = self._doc.styles[style_name]
        font = style.font

        if name is not None:
            font.name = name
        if east_asia is not None or ascii is not None or h_ansi is not None:
            self._set_font_rpr(font, ascii=ascii, east_asia=east_asia, h_ansi=h_ansi)
        if size is not None:
            font.size = Pt(size)
        if bold is not None:
            font.bold = bold
        if color is not None:
            font.color.rgb = RGBColor(*color)

    def set_heading_paragraph(self, level: int, alignment: str | None = None, space_before: int | None = None, space_after: int | None = None) -> None:
        """
        设置标题样式段落格式.

        Args:
            level: 标题级别（1-9）.
            alignment: 对齐方式（"left", "center", "right", "justify"）.
            space_before: 段前间距（磅）.
            space_after: 段后间距（磅）.
        """
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        style_name = f"Heading {level}"
        if style_name not in [s.name for s in self._doc.styles]:
            return

        style = self._doc.styles[style_name]
        para_format = style.paragraph_format

        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }

        if alignment is not None and alignment in alignment_map:
            para_format.alignment = alignment_map[alignment]
        if space_before is not None:
            para_format.space_before = Pt(space_before)
        if space_after is not None:
            para_format.space_after = Pt(space_after)

    def set_normal_font(
        self,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        east_asia: str | None = None,
        ascii: str | None = None,
        h_ansi: str | None = None,
    ) -> None:
        """
        设置 Normal 样式字体格式.

        Args:
            name: 字体名称（同时设置中英文）.
            size: 字体大小（磅）.
            bold: 是否加粗.
            east_asia: 中文字体名称（如 "宋体"）.
            ascii: 西文字体名称（如 "Times New Roman"）.
            h_ansi: 高 ANSI 字体名称.
        """
        if "Normal" not in [s.name for s in self._doc.styles]:
            return

        style = self._doc.styles["Normal"]
        font = style.font

        if name is not None:
            font.name = name
        if east_asia is not None or ascii is not None or h_ansi is not None:
            self._set_font_rpr(font, ascii=ascii, east_asia=east_asia, h_ansi=h_ansi)
        if size is not None:
            font.size = Pt(size)
        if bold is not None:
            font.bold = bold

    def enable_outline_level(self, style_name: str, level: int) -> None:
        """
        启用大纲级别（用于目录识别）.

        Args:
            style_name: 样式名称（如 "Heading 1"）.
            level: 大纲级别（0-9）.
        """
        if style_name not in [s.name for s in self._doc.styles]:
            return

        style = self._doc.styles[style_name]
        pPr = style._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            style._element.append(pPr)

        outlineLvl = pPr.find(qn("w:outlineLvl"))
        if outlineLvl is None:
            outlineLvl = OxmlElement("w:outlineLvl")
            pPr.append(outlineLvl)
        outlineLvl.set(qn("w:val"), str(level))
