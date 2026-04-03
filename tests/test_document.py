"""Word 文档模块测试."""

from __future__ import annotations

import tempfile
import os
from sword import WordDocument


class TestWordDocument:
    """WordDocument 测试类."""

    def test_create_new_document(self) -> None:
        """测试创建新文档."""
        doc = WordDocument()
        assert doc.inner is not None

    def test_set_page_break_between_sections(self) -> None:
        """测试设置章节间分页."""
        doc = WordDocument()
        assert doc._page_break_between_sections is True

        doc.set_page_break_between_sections(False)
        assert doc._page_break_between_sections is False

        doc.set_page_break_between_sections(True)
        assert doc._page_break_between_sections is True

    def test_add_section(self) -> None:
        """测试添加章节."""
        doc = WordDocument()
        doc.set_page_break_between_sections(False)
        section = doc.add_section()
        section.add_paragraph("Chapter 1")

    def test_set_table_of_contents(self) -> None:
        """测试设置目录."""
        doc = WordDocument()
        doc.set_table_of_contents()

    def test_save_document(self, tmp_path: tempfile.TemporaryDirectory) -> None:
        """测试保存文档."""
        doc = WordDocument()
        file_path = os.path.join(tmp_path, "output.docx")
        doc.save(file_path)
        assert os.path.exists(file_path)

    def test_context_manager(self, tmp_path: tempfile.TemporaryDirectory) -> None:
        """测试上下文管理器."""
        file_path = os.path.join(tmp_path, "output.docx")
        with WordDocument() as doc:
            doc.set_page_break_between_sections(False)
            section = doc.add_section()
            section.add_paragraph("Test content")
