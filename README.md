# SWord

A simple Python wrapper library for Microsoft Word documents.

## Installation

```bash
pip install python-docx
```

## Quick Start

```python
from sword import WordDocument
from sword.enums import NumberingFormat

# Create a new document
with WordDocument() as doc:
    doc.set_numbering_format(NumberingFormat.DECIMAL)
    doc.add_heading("第一章", level=1)
    doc.add_paragraph("Some content here")
    doc.add_heading("1.1 子章节", level=2)
    doc.save("output.docx")
```

## License

MIT
