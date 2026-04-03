"""WordSection 模块测试."""

from __future__ import annotations

import tempfile
import os
from sword import WordDocument, WordSection


class TestWordSection:
    """WordSection 测试类."""

    def test_create_section_with_title(self) -> None:
        """测试创建带标题的章节."""
        doc = WordDocument()
        section = WordSection(doc.inner, title="第一章")
        assert section.title == "第一章"

    def test_create_section_without_title(self) -> None:
        """测试创建不带标题的章节."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        assert section.title is None

    def test_add_heading(self) -> None:
        """测试添加标题."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        section.add_numbered_heading("第一章")
        section.add_numbered_heading("小节")

    def test_add_paragraph(self) -> None:
        """测试添加段落."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        section.add_paragraph("这是第一段内容。")
        section.add_paragraph("这是第二段内容。")

    def test_add_page_break(self) -> None:
        """测试添加分页符."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        section.add_paragraph("章节内容")
        section.add_page_break()
        section.add_paragraph("新页面内容")

    def test_add_table(self) -> None:
        """测试添加表格."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        table = section.add_table(rows=3, cols=3, style="Table Grid")
        assert len(table.table.rows) == 3
        assert len(table.table.columns) == 3

    def test_section_content_access(self) -> None:
        """测试章节内容访问."""
        doc = WordDocument()
        section = WordSection(doc.inner)
        section.add_paragraph("Test content")
        assert len(doc.inner.paragraphs) > 0

    def test_add_numbered_heading(self) -> None:
        """测试添加带自动编号的标题."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)

        section1 = doc.add_section("第一章", title_level=1)
        section1.add_numbered_heading("第一节")
        section1.add_numbered_heading("第二小节")

        section2 = doc.add_section("第二章", title_level=1)
        section2.add_numbered_heading("第一节")

        subsection = section1.add_section("子节")
        subsection.add_numbered_heading("子子节")

    def test_context_manager(self) -> None:
        """测试上下文管理器."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)
        with WordSection(doc.inner) as section:
            section.add_paragraph("Test content")
            assert section.title is None


class TestWordDocumentSections:
    """WordDocument 章节管理测试类."""

    def test_add_section_with_title(self) -> None:
        """测试通过 WordDocument 添加带标题的章节."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)

        ch1 = doc.add_section("第一章")
        ch1.add_paragraph("第一章的内容")

        ch2 = doc.add_section("第二章")
        ch2.add_paragraph("第二章的内容")

        assert ch1.title == "第一章"
        assert ch2.title == "第二章"

    def test_add_section_without_title(self) -> None:
        """测试通过 WordDocument 添加不带标题的章节."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)

        section = doc.add_section()
        section.add_paragraph("无标题章节的内容")
        assert section.title is None

    def test_add_section_with_page_break(self) -> None:
        """测试添加章节时分页功能."""
        doc = WordDocument()
        doc.set_page_break_between_sections(True)

        ch1 = doc.add_section("第一章")
        ch1.add_paragraph("第一章内容")

        ch2 = doc.add_section("第二章")
        ch2.add_paragraph("第二章内容")

    def test_save_document_with_sections(
        self, tmp_path: tempfile.TemporaryDirectory
    ) -> None:
        """测试保存包含章节的文档."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)

        ch1 = doc.add_section("第一章")
        ch1.add_paragraph("第一章的内容")

        ch2 = doc.add_section("第二章")
        ch2.add_paragraph("第二章的内容")

        file_path = os.path.join(tmp_path, "sections.docx")
        doc.save(file_path)
        assert os.path.exists(file_path)
