"""Word 文档格式样式封装模块."""

from __future__ import annotations

from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# 主题字体属性，需要清除以避免覆盖显式设置
_THEME_FONT_ATTRS = (
    qn("w:asciiTheme"),
    qn("w:eastAsiaTheme"),
    qn("w:hAnsiTheme"),
    qn("w:cstheme"),
    qn("w:hint"),
)

# 对齐方式映射
_ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 行距模式映射
_LINE_SPACING_MAP = {
    "at_least": WD_LINE_SPACING.AT_LEAST,
    "exactly": WD_LINE_SPACING.EXACTLY,
    "multiple": WD_LINE_SPACING.MULTIPLE,
}

# 表格对齐方式映射
_JC_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "both",
}


class StyleFormat:
    """Word 文档样式格式封装类."""

    def __init__(self, document: DocxDocument) -> None:
        """
        初始化样式格式.

        Args:
            document: 底层的 python-docx Document 对象.
        """
        self._doc = document

    def _get_style(self, style_name: str, style_type: WD_STYLE_TYPE | None = None):
        """
        获取样式对象.

        Args:
            style_name: 样式名称.
            style_type: 样式类型筛选（可选）。

        Returns:
            样式对象，不存在则返回 None.
        """
        for s in self._doc.styles:
            if s.name == style_name:
                if style_type is None or s.type == style_type:
                    return s
        return None

    def _ensure_elem(self, parent: OxmlElement, tag: str) -> OxmlElement:
        """
        确保子元素存在，不存在则创建.

        Args:
            parent: 父元素.
            tag: 子元素标签名（不含 w: 前缀）。

        Returns:
            子元素.
        """
        elem = parent.find(qn(f"w:{tag}"))
        if elem is None:
            elem = OxmlElement(f"w:{tag}")
            parent.append(elem)
        return elem

    def _set_font_rpr(
        self,
        font,
        ascii: str | None = None,
        east_asia: str | None = None,
        h_ansi: str | None = None,
    ) -> None:
        """
        设置字体属性（支持中英文分开）.

        Args:
            font: 字体对象.
            ascii: 西文字体名称.
            east_asia: 东亚字体名称（中文）.
            h_ansi: 高 ANSI 字体名称.
        """
        style_elem = font._element
        rPr = self._ensure_elem(style_elem, "rPr")
        rFonts = self._ensure_elem(rPr, "rFonts")

        # 清除主题属性
        for attr in _THEME_FONT_ATTRS:
            rFonts.attrib.pop(attr, None)

        if ascii is not None:
            rFonts.set(qn("w:ascii"), ascii)
        if east_asia is not None:
            rFonts.set(qn("w:eastAsia"), east_asia)
        if h_ansi is not None:
            rFonts.set(qn("w:hAnsi"), h_ansi)

    def _apply_font(
        self,
        font,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        color: tuple[int, int, int] | None = None,
        east_asia: str | None = None,
        ascii: str | None = None,
        h_ansi: str | None = None,
    ) -> None:
        """
        应用字体属性到字体对象.

        Args:
            font: 字体对象.
            name: 字体名称（同时设置中英文）.
            size: 字体大小（磅）.
            bold: 是否加粗.
            color: RGB 颜色元组.
            east_asia: 中文字体名称.
            ascii: 西文字体名称.
            h_ansi: 高 ANSI 字体名称.
        """
        if name is not None:
            self._set_font_rpr(font, ascii=name, east_asia=name, h_ansi=name)
        elif east_asia is not None or ascii is not None or h_ansi is not None:
            self._set_font_rpr(font, ascii=ascii, east_asia=east_asia, h_ansi=h_ansi)

        if size is not None:
            font.size = Pt(size)
        if bold is not None:
            font.bold = bold
        if color is not None:
            font.color.rgb = RGBColor(*color)

    def _apply_paragraph(
        self,
        style,
        alignment: str | None = None,
        space_before: int | None = None,
        space_after: int | None = None,
        line_spacing: int | float | None = None,
        line_spacing_rule: str | None = None,
        left_indent: int | None = None,
        right_indent: int | None = None,
        first_line_indent: int | None = None,
        hanging_indent: int | None = None,
    ) -> None:
        """
        应用段落格式到样式.

        Args:
            style: 样式对象.
            alignment: 对齐方式.
            space_before: 段前间距（磅）.
            space_after: 段后间距（磅）.
            line_spacing: 行距值.
            line_spacing_rule: 行距模式（"at_least"/"exactly"/"multiple"）。
            left_indent: 左边缩进（字符数）。
            right_indent: 右边缩进（字符数）。
            first_line_indent: 首行缩进（字符数）。
            hanging_indent: 悬挂缩进（字符数）。
        """
        para_format = style.paragraph_format

        if alignment is not None and alignment in _ALIGNMENT_MAP:
            para_format.alignment = _ALIGNMENT_MAP[alignment]
        if space_before is not None:
            para_format.space_before = Pt(space_before)
        if space_after is not None:
            para_format.space_after = Pt(space_after)

        if line_spacing is not None:
            rule = _LINE_SPACING_MAP.get(line_spacing_rule, WD_LINE_SPACING.AT_LEAST)
            para_format.line_spacing_rule = rule
            para_format.line_spacing = (
                Pt(line_spacing) if rule != WD_LINE_SPACING.MULTIPLE else line_spacing
            )

        # 缩进使用字符计数属性
        if any([left_indent, right_indent, first_line_indent, hanging_indent]):
            pPr = self._ensure_elem(style._element, "pPr")
            ind = self._ensure_elem(pPr, "ind")

            if left_indent is not None:
                ind.set(qn("w:leftChars"), str(int(left_indent * 100)))
            if right_indent is not None:
                ind.set(qn("w:rightChars"), str(int(right_indent * 100)))
            if first_line_indent is not None:
                ind.set(qn("w:firstLineChars"), str(int(first_line_indent * 100)))
            if hanging_indent is not None:
                ind.set(qn("w:hangingChars"), str(int(hanging_indent * 100)))

    # ==================== 段落样式方法 ====================

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
        style = self._get_style(f"Heading {level}")
        if style is not None:
            self._apply_font(
                style.font,
                name=name,
                size=size,
                bold=bold,
                color=color,
                east_asia=east_asia,
                ascii=ascii,
                h_ansi=h_ansi,
            )

    def set_heading_paragraph(
        self,
        level: int,
        alignment: str | None = None,
        space_before: int | None = None,
        space_after: int | None = None,
        line_spacing: int | float | None = None,
        line_spacing_rule: str | None = None,
        left_indent: int | None = None,
        right_indent: int | None = None,
        first_line_indent: int | None = None,
        hanging_indent: int | None = None,
    ) -> None:
        """
        设置标题样式段落格式.

        Args:
            level: 标题级别（1-9）.
            alignment: 对齐方式（"left", "center", "right", "justify"）.
            space_before: 段前间距（磅）.
            space_after: 段后间距（磅）.
            line_spacing: 行距值.
            line_spacing_rule: 行距模式（"at_least"/"exactly"/"multiple"）。
            left_indent: 左边缩进（字符数）.
            right_indent: 右边缩进（字符数）.
            first_line_indent: 首行缩进（字符数）。
            hanging_indent: 悬挂缩进（字符数）。
        """
        style = self._get_style(f"Heading {level}")
        if style is not None:
            self._apply_paragraph(
                style,
                alignment=alignment,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                line_spacing_rule=line_spacing_rule,
                left_indent=left_indent,
                right_indent=right_indent,
                first_line_indent=first_line_indent,
                hanging_indent=hanging_indent,
            )

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
        style = self._get_style("Normal")
        if style is not None:
            self._apply_font(
                style.font,
                name=name,
                size=size,
                bold=bold,
                east_asia=east_asia,
                ascii=ascii,
                h_ansi=h_ansi,
            )

    def set_normal_paragraph(
        self,
        alignment: str | None = None,
        space_before: int | None = None,
        space_after: int | None = None,
        line_spacing: int | float | None = None,
        line_spacing_rule: str | None = None,
        left_indent: int | None = None,
        right_indent: int | None = None,
        first_line_indent: int | None = None,
        hanging_indent: int | None = None,
    ) -> None:
        """
        设置 Normal 样式段落格式.

        Args:
            alignment: 对齐方式（"left", "center", "right", "justify"）.
            space_before: 段前间距（磅）.
            space_after: 段后间距（磅）.
            line_spacing: 行距值.
            line_spacing_rule: 行距模式（"at_least"/"exactly"/"multiple"）。
            left_indent: 左边缩进（字符数）.
            right_indent: 右边缩进（字符数）.
            first_line_indent: 首行缩进（字符数）。
            hanging_indent: 悬挂缩进（字符数）。
        """
        style = self._get_style("Normal")
        if style is not None:
            self._apply_paragraph(
                style,
                alignment=alignment,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                line_spacing_rule=line_spacing_rule,
                left_indent=left_indent,
                right_indent=right_indent,
                first_line_indent=first_line_indent,
                hanging_indent=hanging_indent,
            )

    def enable_outline_level(self, style_name: str, level: int) -> None:
        """
        启用大纲级别（用于目录识别）.

        Args:
            style_name: 样式名称（如 "Heading 1"）.
            level: 大纲级别（0-9）.
        """
        style = self._get_style(style_name)
        if style is None:
            return

        pPr = self._ensure_elem(style._element, "pPr")
        outlineLvl = self._ensure_elem(pPr, "outlineLvl")
        outlineLvl.set(qn("w:val"), str(level))

    # ==================== 样式创建方法 ====================

    def create_style(
        self,
        style_name: str,
        style_type: str = "paragraph",
        based_on: str | None = None,
    ):
        """
        创建样式.

        Args:
            style_name: 样式名称.
            style_type: 样式类型（"paragraph", "character", "table", "list"）。
            based_on: 基础样式名称（可选）。

        Returns:
            创建的样式对象，已存在则返回现有样式。
        """
        type_map = {
            "paragraph": WD_STYLE_TYPE.PARAGRAPH,
            "character": WD_STYLE_TYPE.CHARACTER,
            "table": WD_STYLE_TYPE.TABLE,
            "list": WD_STYLE_TYPE.LIST,
        }
        wd_type = type_map.get(style_type, WD_STYLE_TYPE.PARAGRAPH)

        existing = self._get_style(style_name, wd_type)
        if existing is not None:
            return existing

        # 查找基础样式
        base_style_obj = None
        if based_on is not None:
            base_style_obj = self._get_style(based_on)

        try:
            style = self._doc.styles.add_style(style_name, wd_type)
        except ValueError:
            return self._get_style(style_name, wd_type)

        if base_style_obj is not None:
            style.base_style = base_style_obj

        return style

    def create_table_style(self, style_name: str, based_on: str | None = None):
        """
        创建表格样式.

        Args:
            style_name: 样式名称.
            based_on: 基础样式名称（可选，如 "Table Grid"）。

        Returns:
            创建的样式对象，已存在则返回现有样式。
        """
        return self.create_style(style_name, style_type="table", based_on=based_on)

    # ==================== 表格样式方法 ====================

    def set_table_font(
        self,
        style_name: str,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        color: tuple[int, int, int] | None = None,
        east_asia: str | None = None,
        ascii: str | None = None,
    ) -> None:
        """
        设置表格样式字体格式.

        Args:
            style_name: 样式名称（如 "Table Grid"）。
            name: 字体名称（同时设置中英文）。
            size: 字体大小（磅）.
            bold: 是否加粗.
            color: RGB 颜色元组.
            east_asia: 中文字体名称.
            ascii: 西文字体名称.
        """
        style = self._get_style(style_name, WD_STYLE_TYPE.TABLE)
        if style is not None:
            self._apply_font(
                style.font,
                name=name,
                size=size,
                bold=bold,
                color=color,
                east_asia=east_asia,
                ascii=ascii,
            )

    def set_table_shading(self, style_name: str, fill: str) -> None:
        """
        设置表格样式底纹/背景色.

        Args:
            style_name: 样式名称.
            fill: 填充颜色（十六进制字符串，如 "FF0000"）。
        """
        style = self._get_style(style_name, WD_STYLE_TYPE.TABLE)
        if style is None:
            return

        tblPr = self._ensure_elem(style._element, "tblPr")
        shd = self._ensure_elem(tblPr, "shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)

    def set_table_borders(
        self,
        style_name: str,
        top: str | None = None,
        bottom: str | None = None,
        left: str | None = None,
        right: str | None = None,
        inside_h: str | None = None,
        inside_v: str | None = None,
        border_size: int = 4,
        border_color: str = "000000",
    ) -> None:
        """
        设置表格样式边框.

        Args:
            style_name: 样式名称.
            top: 上边框样式（"single", "double", "none"）。
            bottom: 下边框样式.
            left: 左边框样式.
            right: 右边框样式.
            inside_h: 内部水平边框样式.
            inside_v: 内部垂直边框样式.
            border_size: 边框大小（磅）.
            border_color: 边框颜色（十六进制字符串）.
        """
        style = self._get_style(style_name, WD_STYLE_TYPE.TABLE)
        if style is None:
            return

        tblPr = self._ensure_elem(style._element, "tblPr")
        tblBorders = self._ensure_elem(tblPr, "tblBorders")

        border_attrs = {
            "top": top,
            "bottom": bottom,
            "left": left,
            "right": right,
            "insideH": inside_h,
            "insideV": inside_v,
        }

        # 检查是否有任何边框样式被指定
        has_style_specified = any(v is not None for v in border_attrs.values())

        for border_name, border_style in border_attrs.items():
            elem = self._ensure_elem(tblBorders, border_name)
            if border_style is not None:
                # 指定了样式，更新所有属性
                elem.set(qn("w:val"), border_style)
                elem.set(qn("w:sz"), str(border_size*8))
                elem.set(qn("w:space"), "0")
                elem.set(qn("w:color"), border_color)
            elif has_style_specified:
                # 已指定其他边框样式但此边框未指定，保持原样
                pass
            else:
                # 没有指定任何边框样式，更新颜色和大小（作用于已有边框）
                current_val = elem.get(qn("w:val"))
                if current_val is not None:
                    elem.set(qn("w:sz"), str(border_size*8))
                    elem.set(qn("w:color"), border_color)

    def set_table_alignment(self, style_name: str, alignment: str = "center") -> None:
        """
        设置表格样式对齐方式.

        Args:
            style_name: 样式名称.
            alignment: 对齐方式（"left", "center", "right", "justify"）。
        """
        style = self._get_style(style_name, WD_STYLE_TYPE.TABLE)
        if style is None:
            return

        tblPr = self._ensure_elem(style._element, "tblPr")
        jc = self._ensure_elem(tblPr, "jc")
        if alignment in _JC_MAP:
            jc.set(qn("w:val"), _JC_MAP[alignment])

    def set_table_paragraph(
        self,
        style_name: str,
        alignment: str | None = None,
        space_before: int | None = None,
        space_after: int | None = None,
        line_spacing: int | float | None = None,
        line_spacing_rule: str | None = None,
    ) -> None:
        """
        设置表格样式段落格式.

        Args:
            style_name: 样式名称.
            alignment: 对齐方式.
            space_before: 段前间距（磅）.
            space_after: 段后间距（磅）.
            line_spacing: 行距值.
            line_spacing_rule: 行距模式（"at_least"/"exactly"/"multiple"）。
        """
        style = self._get_style(style_name, WD_STYLE_TYPE.TABLE)
        if style is not None:
            self._apply_paragraph(
                style,
                alignment=alignment,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                line_spacing_rule=line_spacing_rule,
            )
