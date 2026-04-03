"""WordParagraph 模块测试."""

from __future__ import annotations

import tempfile
import os

from sword import WordDocument


class TestWordParagraph:
    """WordParagraph 测试类."""

    def test_paragraph_text(self) -> None:
        """测试获取和设置段落文本."""
        doc = WordDocument()
        para = doc.add_paragraph("测试内容")
        assert para.text == "测试内容"

    def test_paragraph_text_setter(self) -> None:
        """测试设置段落文本."""
        doc = WordDocument()
        para = doc.add_paragraph("原始内容")
        para.text = "新内容"
        assert para.text == "新内容"

    def test_paragraph_inner(self) -> None:
        """测试获取底层 Paragraph 对象."""
        doc = WordDocument()
        para = doc.add_paragraph()
        assert para.inner is not None

    def test_paragraph_context_manager(self) -> None:
        """测试上下文管理器."""
        doc = WordDocument()
        para = doc.add_paragraph()
        with para as p:
            p.text = "上下文内容"
        assert para.text == "上下文内容"

    def test_add_run(self) -> None:
        """测试添加文本Run."""
        doc = WordDocument()
        para = doc.add_paragraph()
        run = para.add_run("Run文本")
        assert run is not None
        assert run.text == "Run文本"

    def test_add_run_empty(self) -> None:
        """测试添加空Run."""
        doc = WordDocument()
        para = doc.add_paragraph()
        run = para.add_run()
        assert run is not None
        assert run.text == ""

    def test_iter_runs(self) -> None:
        """测试遍历段落中所有Run."""
        doc = WordDocument()
        para = doc.add_paragraph()
        para.add_run("第一段")
        para.add_run("第二段")
        para.add_run("第三段")

        runs = list(para.iter_runs())
        assert len(runs) == 3
        assert runs[0].text == "第一段"
        assert runs[1].text == "第二段"
        assert runs[2].text == "第三段"

    def test_iter_runs_empty(self) -> None:
        """测试遍历空段落."""
        doc = WordDocument()
        para = doc.add_paragraph()
        runs = list(para.iter_runs())
        assert len(runs) == 0

    def test_get_run_valid_index(self) -> None:
        """测试获取指定索引的Run."""
        doc = WordDocument()
        para = doc.add_paragraph()
        para.add_run("第一")
        para.add_run("第二")
        para.add_run("第三")

        run = para.get_run(1)
        assert run is not None
        assert run.text == "第二"

    def test_get_run_invalid_index(self) -> None:
        """测试获取无效索引的Run返回None."""
        doc = WordDocument()
        para = doc.add_paragraph()
        para.add_run("第一")

        run = para.get_run(5)
        assert run is None

    def test_get_run_negative_index(self) -> None:
        """测试负数索引（支持类似列表的负索引行为）."""
        doc = WordDocument()
        para = doc.add_paragraph()
        para.add_run("第一")

        run = para.get_run(-1)
        assert run is not None
        assert run.text == "第一"

    def test_save_document_with_paragraph(
        self, tmp_path: tempfile.TemporaryDirectory
    ) -> None:
        """测试保存包含段落的文档."""
        doc = WordDocument()
        para = doc.add_paragraph("测试段落")
        para.add_run(" - 带Run")

        file_path = os.path.join(tmp_path, "paragraph_test.docx")
        doc.save(file_path)
        assert os.path.exists(file_path)
