"""Word 文档段落封装模块."""

from __future__ import annotations

from typing import Any, Iterator

from sword.run import WordRun


class WordParagraph:
    """封装 Word 文档中的段落，提供便捷的段落操作."""

    def __init__(self, paragraph) -> None:
        """
        初始化段落.

        Args:
            paragraph: 底层的 python-docx Paragraph 对象.
        """
        self._para = paragraph

    @property
    def inner(self):
        """获取底层的 python-docx Paragraph 对象."""
        return self._para

    @property
    def text(self) -> str:
        """获取段落文本."""
        return self._para.text

    @text.setter
    def text(self, value: str) -> None:
        """设置段落文本."""
        self._para.text = value

    def add_run(self, text: str = "") -> WordRun:
        """
        添加文本Run.

        Args:
            text: Run文本.

        Returns:
            WordRun: 创建的Run封装对象.
        """
        run = self._para.add_run(text)
        word_run = WordRun(run)

        return word_run

    def iter_runs(self) -> Iterator[WordRun]:
        """
        遍历段落中所有Run.

        Yields:
            WordRun 对象。
        """
        for run in self._para.runs:
            yield WordRun(run)

    def get_run(self, index: int) -> WordRun | None:
        """
        获取指定索引的Run.

        Args:
            index: Run索引（从 0 开始）。

        Returns:
            WordRun 对象，若索引无效则返回 None。
        """
        try:
            return WordRun(self._para.runs[index])
        except IndexError:
            return None

    def __enter__(self) -> WordParagraph:
        return self

    def __exit__(self, *args: Any) -> None:
        pass
