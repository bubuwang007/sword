"""SWord - 简洁易用的 Word 文档封装库."""

from sword.document import WordDocument
from sword.section import WordSection
from sword.table import WordTable
from sword.cell import WordCell
from sword.format import StyleFormat

__version__ = "0.1.0"
__all__ = ["WordDocument", "WordSection", "WordTable", "WordCell", "StyleFormat"]
