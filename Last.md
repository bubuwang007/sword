# Word 文档库 - 实现计划

## 项目概述

**项目名称:** ansys-word（新实现）
**类型:** Python Word 文档生成库
**核心功能:** 程序化创建 Word 文档，支持文本、表格、图片、公式和层级章节，用于技术文档（中文工程/学术风格）
**目标用户:** 生成技术报告的中文工程/学术用户

---

## 当前库分析

### 现有架构

当前库存在**两套平行的实现**，可能是不同时间开发的：

1. **旧版架构**（`__init__.py`）：
   - `Calculate_Document` - 文档构建器，含硬编码编号逻辑
   - `Document_Table` / `Document_Pic_Table` - 表格构建器
   - `Document_Format` - 静态格式化方法
   - `Calculate_Basic_Info` - Excel 数据读取

2. **新版架构**（`A_Doucment.py` 等）：
   - `A_Document` → `A_Section` → `A_Paragraph` → `A_Run`
   - `A_Table` → `A_Cell`
   - `A_Pic_Table`、`Pics_Table`
   - `Format.py` - 字体/间距常量

### 关键问题

| 问题 | 描述 | 影响 |
|------|------|------|
| **并行的 API** | 两套完全不同的类层次结构 | 用户困惑，维护负担 |
| **类级可变默认值** | `A_Section.font = FONT_NORMAL` 在实例间共享 | Bug：一个实例的修改会影响所有实例 |
| **eval() 处理变量** | `A_Paragraph.append_parse_text` 中的 `eval("{}".format(i[1]))` | 安全漏洞 |
| **全局状态** | `set_var()` 修改 `globals()` | 不可预测的行为 |
| **命名空间污染** | `from ..utils.reg import *` 将正则导入模块 | 名称冲突 |
| **文件名拼写错误** | `A_Doucment.py`（少一个 m） | 导入混淆 |
| **注释掉的代码** | `A_Table.append_component`、`Format.py` 中的死代码 | 代码膨胀 |
| **命名不一致** | `append_parse_text` vs `parse_append` | API 混乱 |
| **set_font bug** | `A_Paragraph.set_font` 注释："在有公式的段落中，设置字体会导致公式出现显示bug" | 已知问题未修复 |

---

## 新库设计

### 架构

```
word/
├── document.py          # Document 类 - 入口
├── section.py          # Section 类 - 层级章节
├── paragraph.py        # Paragraph、Run 类
├── table.py            # Table、Cell 类
├── picture.py          # Picture、PictureTable 类
├── format.py           # Font、Spacing、Color 数据类
├── formula.py          # LaTeX 转 OMML
├── parser.py           # 文本解析（变量、下标、公式）
└── __init__.py         # 公开 API 导出
```

### 核心原则

1. **单一类层次结构** - 一种方式做事，不是两种
2. **不可变默认值** - 使用不可变模式的 dataclasses
3. **不使用 eval()** - 使用 `ast.literal_eval()` 或显式变量查找
4. **无全局状态** - 显式传递上下文，不操作 `globals()`
5. **清晰的命名空间** - 显式导入，无通配符导入
6. **组合优于继承** - 小而专注的类

---

## 功能需求

### 1. 文档结构

- [ ] 创建新文档或打开已有文档
- [ ] 添加层级章节（1、1.1、1.1.1 等）
- [ ] 自动章节编号
- [ ] 章节间分页（可选）
- [ ] 保存文档到文件
- [ ] 保存后打开文档（可选）

### 2. 文本和段落

- [ ] 添加纯文本段落（默认格式）
- [ ] 添加自定义格式的段落
- [ ] 段落内富文本（混合字体/大小）
- [ ] 文本解析，支持特殊标记：
  - 变量：`<var>variable_name</var>` → 替换为值
  - 公式：`<f>LaTeX 表达式</f>` → Word OMML 公式
  - 下标：`<sub>text</sub>` → 下标
  - 上标：`<super>text</super>` → 上标
- [ ] 行距控制
- [ ] 缩进（首行、左、右、悬挂）
- [ ] 对齐（左、中、右、两端）

### 3. 表格

- [ ] 创建带标题和自动编号的表格（表 1.1-1）
- [ ] 设置表格样式（网格、无边框等）
- [ ] 向单元格写入文本
- [ ] 向单元格写入解析文本（支持变量、公式）
- [ ] 合并单元格（水平、垂直）
- [ ] 在单元格中添加图片
- [ ] 基于字典结构构建表格组件
- [ ] 表头垂直文字方向

### 4. 图片

- [ ] 插入单张图片并带标题
- [ ] 插入图片表格（多图网格，带子标题）
- [ ] 自动图片编号（图 1.1-1、(a)、(b) 等）
- [ ] 尺寸控制（宽度和/或高度，保持宽高比）
- [ ] 对齐（默认居中）

### 5. 公式

- [ ] 块公式（居中，单独一行）
- [ ] 行内公式（在文本中）
- [ ] 支持 LaTeX 语法
- [ ] 常用工程公式

### 6. 字体和格式

- [ ] 中文字体（默认：宋体）
- [ ] 英文字体（默认：Times New Roman）
- [ ] 字号（默认：12pt / 小四）
- [ ] 加粗、斜体、下划线
- [ ] 文字颜色（RGB）
- [ ] 段落间距（段前、段后、行距）

---

## API 设计

### 文档创建

```python
from ansys.word import Document, Section

# 创建文档
doc = Document()

# 添加章节（带标题）
sec1 = doc.add_section("引言")

# 添加段落
sec1.add_paragraph("纯文本段落")
sec1.add_paragraph("<var>author</var> 编写了这份报告")
sec1.add_paragraph("公式：<f>\frac{a}{b}</f>")

# 添加表格
table = sec1.add_table("结果", rows=3, cols=3)
table.write(0, 0, "表头")
table.write(0, 1, "<var>value</var>")

# 添加图片
pic = sec1.add_picture("image.png", "实验装置")

# 保存
doc.save("report.docx")
```

### 章节层级

```python
sec1 = doc.add_section("第1章")
sec1_1 = sec1.add_section("1.1 节")
sec1_1_1 = sec1_1.add_section("1.1.1 节")
```

### 表格构建

```python
table = sec.add_table("组件列表", rows=5, cols=4)

# 简单单元格写入
table.write(0, 0, "组件")

# 一行数据
table.write_row(1, 0, ["零件A", "零件B", "零件C"])

# 一列数据
table.write_col(0, 1, ["第1行", "第2行", "第3行"])

# 合并单元格
table.merge((0, 0), (0, 3))  # 表头行

# 组件字典（层级数据）
component_data = {
    "电机": ["转子", "定子"],
    "泵": None,  # 占满整行
}
table.from_components(component_data)
```

### 文本解析

```python
# 写入时替换变量
context = {"author": "张三", "date": "2024-01-01"}
sec.parse_vars(context)

# 带特殊语法的高亮文本
para = sec.add_paragraph("速度 = <var>v</var> m/s")
para += " 角度 <f>\alpha = 90^\circ</f>"
para += " 结果 (<super>ref</super>)"
```

---

## 数据结构

### 字体规格

```python
@dataclass(frozen=True)
class Font:
    east_asia: str = "宋体"                    # 中文字体
    ascii: str = "Times New Roman"             # 西文字体
    size: int = 12                             # 字号（磅）
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: tuple[int, int, int] = (0, 0, 0)  # RGB
```

### 段落格式

```python
@dataclass(frozen=True)
class ParagraphFormat:
    alignment: str = "left"       # left, center, right, justify
    line_spacing: float = 1.5
    space_before: float = 0
    space_after: float = 0
    first_line_indent: float = 2  # 字符数
    left_indent: float = 0
    right_indent: float = 0
```

### 表格格式

```python
@dataclass(frozen=True)
class TableFormat:
    style: str = "Table Grid"
    alignment: str = "center"
    cell_vertical_align: str = "center"  # top, center, bottom
    title_font: Font = Font()
    content_font: Font = Font(size=11)
    content_format: ParagraphFormat = ParagraphFormat(
        alignment="center",
        line_spacing=1,
        space_before=1,
        space_after=1
    )
```

---

## 文件结构

```
src/ansys/word/
├── __init__.py              # 公开 API：Document, Section, Table 等
├── document.py               # Document 类
├── section.py                # Section 类
├── paragraph.py              # Paragraph、Run、InlineRun 类
├── table.py                  # Table、Cell 类
├── picture.py                # Picture、PictureTable 类
├── format.py                 # Font、ParagraphFormat、TableFormat 数据类
├── formula.py                # LaTeX 转 OMML
├── parser.py                 # 变量/公式文本解析
├── exceptions.py             # 自定义异常
└── constants.py              # 默认格式、大纲级别
```

---

## 实现注意事项

### 不使用 eval() - 变量替换

```python
# 旧版（不安全）：
_temp = eval("{}".format(i[1]))

# 新版（安全）：
def resolve_var(name: str, context: dict) -> str:
    if name in context:
        return str(context[name])
    raise KeyError(f"变量 '{name}' 未找到")
```

### 无全局状态

```python
# 旧版：
def set_var(globals_dict):
    global_vars = globals()
    for key, val in globals_dict.items():
        if key not in global_vars:
            global_vars[key] = val

# 新版：显式传递上下文
class Document:
    def __init__(self):
        self.variables = {}  # 实例状态

    def set_vars(self, vars_dict):
        self.variables.update(vars_dict)
```

### 使用 frozen dataclass 实现不可变默认值

```python
from dataclasses import dataclass

@dataclass(frozen=True)
class Font:
    east_asia: str = "宋体"
    ascii: str = "Times New Roman"
    size: int = 12
    bold: bool = False
    # ...

# 默认实例
DEFAULT_FONT = Font()
```

### 显式导入（无通配符）

```python
# 旧版：
from .latex2omml import *
from ..utils.reg import *

# 新版：
from .latex2omml import latex2omml
from ..utils.reg import re_var, re_formula, re_var_content, re_formula_content
```

---

## 测试计划

1. 每个类的单元测试
2. 文档创建的集成测试
3. 表格构建测试（简单、合并、组件）
4. 公式渲染测试
5. 图片插入测试
6. 编码测试（中文文本）

---

## 从旧库迁移

### 旧类 → 新类映射

| 旧类 | 新类 | 备注 |
|------|------|------|
| `Calculate_Document` | `Document` | 统一类 |
| `Document_Table` | `Table` | 重命名 |
| `Document_Pic_Table` | `PictureTable` | 重命名 |
| `A_Section` | `Section` | 简化 |
| `A_Paragraph` | `Paragraph` | 简化 |
| `A_Run` | `InlineRun` | 富文本用 |
| `A_Table` | `Table` | 已存在 |
| `A_Cell` | `Cell` | 重命名 |
| `Document_Format` | （静态方法） | 变为 `ParagraphFormat` |
| `Format.*_NORMAL` | `DEFAULT_FONT` 等 | Frozen dataclasses |

### 用法对比

```python
# 旧版
from ansys.word import A_Document, A_Section, A_Table
doc = A_Document()
sec = doc.add_section("标题", "1")
table = sec.add_table("表", (3, 3))
table.head(["A", "B", "C"])
table.append((1, 0), ["1", "2", "3"])
doc.save("out.docx")

# 新版
from ansys.word import Document, Section, Table
doc = Document()
sec = doc.add_section("标题")
table = sec.add_table("表", rows=3, cols=3)
table.write_header(["A", "B", "C"])
table.write_row(1, 0, ["1", "2", "3"])
doc.save("out.docx")
```

---

## 验收标准

1. [ ] 单一、一致的文档创建 API
2. [ ] 不使用 `eval()`，不操作全局状态
3. [ ] 使用 frozen dataclasses 实现不可变默认值
4. [ ] 保留并改进所有旧功能：
   - 自动编号的层级章节
   - 支持合并和组件构建的表格
   - 带子标题的图片表格
   - LaTeX 公式支持
   - 文本解析（变量、下标、上标）
5. [ ] 清晰的模块结构，显式导入
6. [ ] 所有公开 API 都有类型提示
7. [ ] 核心功能的单元测试
