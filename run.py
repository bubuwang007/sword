"""Word 文档文本Run封装模块."""

from __future__ import annotations

from typing import Any


class WordRun:
    """封装 Word 文档中的文本Run，提供文本格式化功能."""

    def __init__(self, run) -> None:
        """
        初始化文本Run.

        Args:
            run: 底层的 python-docx Run 对象.
        """
        self._run = run

    @property
    def inner(self):
        """获取底层的 python-docx Run 对象."""
        return self._run

    @property
    def text(self) -> str:
        """获取Run文本."""
        return self._run.text

    @text.setter
    def text(self, value: str) -> None:
        """设置Run文本."""
        self._run.text = value

    def __enter__(self) -> WordRun:
        return self

    def __exit__(self, *args: Any) -> None:
        pass
