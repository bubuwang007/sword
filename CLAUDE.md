# SWord

## 项目概述

提供一个简洁、易用的 Python 封装库，用于读写 Microsoft Word 文档（`.docx` 格式）。

## 技术栈

- **Python**: 3.11+
- **核心依赖**: `python-docx`
- **类型检查**: `mypy`
- **代码格式化**: `black`
- **单元测试**: `pytest`

## 项目结构

```
sword/
├── __init__.py       # 包导出
├── document.py       # WordDocument 主类
├── section.py        # WordSection 章节类
├── table.py          # WordTable 表格类
├── cell.py           # WordCell 单元格类
├── paragraph.py      # WordParagraph 段落类
├── run.py            # WordRun 文本Run类
├── format.py         # StyleFormat 样式类
└── tests/
    ├── __init__.py
    ├── test_document.py
    ├── test_section.py
    ├── test_table.py
    ├── test_cell.py
    ├── test_format.py
    └── test_paragraph.py
```

## 核心 API

### WordDocument

主文档类，负责文档创建、保存和章节管理。

```python
from sword import WordDocument

doc = WordDocument()
doc.set_page_break_between_sections(False)  # 章节间不分页
doc.set_start_number(level=1, number=1)       # 设置起始编号

section = doc.add_section("第一章", title_level=1)
section.add_numbered_heading("第一节")
doc.save("output.docx")
```

**主要方法：**
- `add_section(title, title_level)` - 添加章节
- `add_paragraph(text, style)` - 添加段落
- `set_page_break_between_sections(enabled)` - 设置章节分页
- `set_start_number(level, number)` - 设置起始编号
- `set_table_of_contents()` - 插入目录
- `save(path, open_after_save)` - 保存文档

### WordSection

章节类，支持层级编号和嵌套章节。

```python
section = doc.add_section("第一章")
section.add_numbered_heading("第一节")  # 输出: "1.1 第一节"
section.add_paragraph("内容")
section.add_page_break()

# 嵌套章节
subsection = section.add_section("子节")
subsection.add_numbered_heading("子子节")  # 输出: "1.1.1 子子节"
```

**主要方法：**
- `add_numbered_heading(text)` - 添加带自动编号的标题
- `add_paragraph(text, style)` - 添加段落
- `add_page_break()` - 添加分页符
- `add_section(title)` - 添加子章节（自动延续编号）

### WordTable

表格类，支持创建表格和单元格操作。

```python
table = section.add_table(rows=3, cols=3, style="Table Grid")
cell = table.cell(0, 0)
cell.text = "内容"
cell.set_shading("FFFF00")
```

**主要方法：**
- `cell(row, col)` - 获取指定单元格
- `iter_cells()` / `iter_rows()` / `iter_cols()` - 遍历
- `get_row(index)` / `get_column(index)` - 获取行/列

### WordCell

单元格封装类，提供单元格格式化功能。

```python
cell = table.cell(0, 0)
cell.text = "标题"
cell.set_shading("FFFF00")
cell.set_borders(top="single", border_size=8)
cell.set_vertical_alignment("center")
cell.set_width(500, unit="dxa")
```

**主要方法：**
- `text` - 获取/设置单元格文本
- `set_shading(color)` - 设置底纹
- `set_borders(...)` - 设置边框
- `set_vertical_alignment(align)` - 设置垂直对齐
- `set_width(width, unit)` - 设置宽度
- `add_paragraph(text)` - 添加段落
- `iter_paragraphs()` - 遍历段落

### WordParagraph

段落封装类，支持 Run 管理。

```python
para = doc.add_paragraph("段落文本")
run = para.add_run("加粗文本")
for run in para.iter_runs():
    print(run.text)
```

**主要方法：**
- `text` - 获取/设置段落文本
- `add_run(text)` - 添加文本Run
- `iter_runs()` - 遍历所有Run
- `get_run(index)` - 获取指定索引的Run

### WordRun

文本Run封装类，提供文本级别格式化。

```python
run = para.add_run("文本")
run.text = "新文本"
```

**主要方法：**
- `text` - 获取/设置Run文本
- `inner` - 获取底层 python-docx Run 对象

## 编码规范

- 遵循 PEP 8
- 使用类型注解（type hints）
- 所有公开 API 需要文档字符串
- 函数单一职责原则

## 计划和任务

制定的计划和任务写入`.claude/docs`下的`plan`和`task`文件夹中，分类根据项目文件夹结构进行组织。

## 测试

- 编写单元测试覆盖核心功能
- 使用 `pytest` 进行测试运行
- 所有测试文件命名以 `test_` 开头
- 测试文件放在项目根目录下的 `tests` 文件夹中，根据项目文件夹结构进行组织
