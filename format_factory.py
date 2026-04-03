"""Word 文档批量样式设置模块."""

from __future__ import annotations

from sword.format import StyleFormat

__all__ = ["计算书格式"]


def 计算书格式(format: StyleFormat) -> None:

    format.set_normal_font(
        size=12,
        bold=False,
        italic=False,
        east_asia="宋体",
        ascii="Times New Roman",
        color=(0, 0, 0),
    )
    format.set_normal_paragraph(
        alignment="left",
        first_line_indent=2,
        space_before=0,
        space_after=0,
        line_spacing=1.2,
        line_spacing_rule="multiple",
    )

    for i in range(1, 10):
        format.set_heading_font(
            level=i,
            size=12,
            bold=True,
            italic=False,
            east_asia="宋体",
            ascii="Times New Roman",
            color=(0, 0, 0),
        )

        if i == 1:
            format.set_heading_paragraph(
                level=1,
                alignment="center",
                space_before=5,
                space_after=5,
                line_spacing=1.2,
                first_line_indent=0,
                line_spacing_rule="multiple",
            )
        else:
            format.set_heading_paragraph(
                level=i,
                alignment="left",
                left_indent=0,
                first_line_indent=0,
                space_before=5,
                space_after=5,
                line_spacing=1.2,
                line_spacing_rule="multiple",
            )

    format.create_paragraph_style(
        "Table Heading",
    )

    format.set_style_font(
        "Table Heading",
        size=12,
        bold=True,
        italic=False,
        east_asia="宋体",
        ascii="Times New Roman",
        color=(0, 0, 0),
    )

    format.set_table_font(
        "Table Grid",
        size=10,
        bold=False,
        italic=False,
        east_asia="宋体",
        ascii="Times New Roman",
        color=(0, 0, 0),
    )

    format.set_table_paragraph(
        "Table Grid",
        alignment="center",
        space_before=0,
        space_after=0,
        line_spacing=1,
        line_spacing_rule="single",
        first_line_indent=0,
    )
