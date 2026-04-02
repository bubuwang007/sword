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
├── format.py         # StyleFormat 样式类
└── tests/
    ├── __init__.py
    ├── test_document.py
    └── test_section.py
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
