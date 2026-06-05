---
name: md2docx
description: 用自然语言将 Markdown 转换为 Word 文档。用户描述格式需求，你自动生成 YAML 配置并完成转换。
---

# 自然语言 Markdown → Word 转换

用户用自然语言描述想要的文档格式，你负责完成以下全流程：

1. 确认 Markdown 文件路径和输出路径
2. 将自然语言格式需求转为 YAML 配置文件
3. 执行 `md2docx` 命令完成转换
4. 提醒用户更新目录域

## 完整工作流程

### 第0步：检测 md2docx 是否已安装

首先运行 `md2docx -v` 检测 md2docx 是否已安装：

```bash
md2docx -v
```

- **如果正常输出版本号**（如 `md2docx 0.1.0`）→ 已安装，直接进入第1步
- **如果报错**（command not found 等）→ 未安装，按以下流程安装：

```bash
git clone https://github.com/RockyrLiu/md2docx.git /tmp/md2docx
cd /tmp/md2docx && pip install --editable .
```

安装完成后再次运行 `md2docx -v` 确认安装成功，再继续。

### 第1步：确认基本信息

从用户输入中提取，缺失则主动询问：

- **Markdown 文件**：要转换的 `.md` 文件路径
- **输出文件**（可选）：默认与输入同名的 `.docx`
- **封面信息**：标题、作者、班级、学号、指导教师、日期（留空=今天）
- **是否需要封面/目录**：默认都需要，用户说不要则关闭
- **格式要求**：字体、字号、行距、缩进、对齐等

### 第2步：生成 YAML 配置

从下面的默认模板出发，只修改用户明确提到的字段。**未提及的字段保持模板中的默认值不变**。

配置模板（可直接写入 `config.yaml`）：

```yaml
cover:
  enabled: true
  title: "报告标题"
  author: ""
  class_info: ""
  student_id: ""
  teacher: ""
  date: ""

toc:
  enabled: true
  level: 3
  font_name: "宋体"
  font_name_ascii: "Times New Roman"
  font_size: 12
  bold: true
  italic: false
  color: "000000"

styles:
  body:
    font_name: "宋体"
    font_name_ascii: "Times New Roman"
    font_size: 12
    line_spacing: 1.0
    bold: false
    italic: false
    alignment: "justify"
    first_line_indent: 2
    space_before: 0
    space_after: 0
    color: "000000"

  headings:
    h1:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 16
      bold: true
      italic: false
      alignment: "center"
      space_before: 12
      space_after: 12
      color: "000000"
    h2:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 14
      bold: true
      italic: false
      alignment: "center"
      space_before: 6
      space_after: 6
      color: "000000"
    h3:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 12
      bold: true
      italic: false
      alignment: "left"
      space_before: 6
      space_after: 6
      color: "000000"
    h4:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 12
      bold: true
      italic: false
      alignment: "left"
      space_before: 3
      space_after: 3
      color: "000000"
    h5:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 12
      bold: true
      italic: false
      alignment: "left"
      space_before: 3
      space_after: 3
      color: "000000"
    h6:
      font_name: "黑体"
      font_name_ascii: "Times New Roman"
      font_size: 12
      bold: true
      italic: false
      alignment: "left"
      space_before: 3
      space_after: 3
      color: "000000"

  table:
    font_name: "宋体"
    font_name_ascii: "Times New Roman"
    font_size: 10
    bold: false
    header_bold: true
    alignment: "center"
    line_spacing: 1.0

  list:
    font_name: "宋体"
    font_name_ascii: "Times New Roman"
    font_size: 12
    line_spacing: 1.0
    indent: 0                          # 列表整体左缩进字符数，0 表示不缩进
    first_line_indent: 2               # 列表项首行缩进字符数

  code:
    font_name: "Consolas"
    font_name_ascii: "Consolas"
    font_size: 10
    line_spacing: 1.0
    border_color: "999999"
    border_width: 0.75
    show_border: true

  blockquote:
    font_name: "楷体"
    font_name_ascii: "Times New Roman"
    font_size: 12
    italic: true
    line_spacing: 1.0
    left_indent: 0.5
    border_left_color: "999999"

page:
  size: "A4"
  top_margin: 2.54
  bottom_margin: 2.54
  left_margin: 3.18
  right_margin: 3.18
```

### 第3步：执行转换

```bash
md2docx <输入文件.md> -c config.yaml -o <输出文件.docx>
```

### 第4步：提醒用户

> 在 Word 中打开文档后，**右键目录 → 更新域**（或选中目录按 F9）。

---

## 自然语言 → YAML 映射规则

以下规则用于将用户的自然语言描述转换为 YAML 字段值。**凡是下文没有列出的配置项，使用默认模板中的值。**

### 中文字号 → pt

| 用户描述 | `font_size` 值 | 默认用途 |
|---------|---------------|---------|
| 二号 | 22 | — |
| 三号 | 16 | 题目 (H1) |
| 四号 | 14 | 一级标题 (H2) |
| 小四 | 12 | 正文 / 二级及其他标题 (H3–H6) |
| 五号 | 10 | 表格 / 脚注 |

### 字体

| 用户描述 | YAML |
|---------|------|
| 宋体 | `font_name: "宋体"` |
| 黑体 | `font_name: "黑体"` |
| 楷体 | `font_name: "楷体"` |
| 仿宋 | `font_name: "仿宋"` |
| 等宽/Consolas | `font_name: "Consolas"` |
| 英文 Times New Roman | `font_name_ascii: "Times New Roman"` |

### 段落格式

| 用户描述 | YAML |
|---------|------|
| 单倍行距 / 1倍行距 | `line_spacing: 1.0` |
| 1.5倍行距 | `line_spacing: 1.5` |
| 双倍行距 / 2倍行距 | `line_spacing: 2.0` |
| 首行缩进2字符 | `first_line_indent: 2` |
| 不缩进 / 顶格 | `first_line_indent: 0` |
| 居中 | `alignment: "center"` |
| 左对齐 | `alignment: "left"` |
| 右对齐 | `alignment: "right"` |
| 两端对齐 | `alignment: "justify"` |
| 加粗 | `bold: true` |
| 不加粗 | `bold: false` |

### 开关与页面

| 用户描述 | YAML |
|---------|------|
| 不要封面 | `cover.enabled: false` |
| 不要目录 | `toc.enabled: false` |
| 目录到二级标题 | `toc.level: 2` |
| 目录到三级标题 | `toc.level: 3` |
| A4纸 | `page.size: "A4"` |

### 封面信息

| 用户描述 | YAML |
|---------|------|
| 标题改成"XXX" / 封面标题"XXX" | `cover.title: "XXX"` |
| 作者"XXX" / 姓名"XXX" | `cover.author: "XXX"` |
| 班级"XXX" | `cover.class_info: "XXX"` |
| 学号"XXX" | `cover.student_id: "XXX"` |
| 指导教师"XXX" | `cover.teacher: "XXX"` |

### 常见组合（用户说了这些关键词直接套用）

| 用户说 | 意味着 |
|--------|--------|
| 「学术风格」「学术论文」 | 正文宋体12pt, 单倍行距, 首行缩进2, 两端对齐; 标题黑体加粗（h1三号居中，h2四号居中，h3+小四居左）; 三线表 |
| 「课程报告」「实验报告」 | 含封面, 目录到三级, 学术风格正文 |
| 「读书笔记」 | 无封面, 无目录, 正文楷体12pt, 1.5行距 |

---

## 示例对话

**用户**：`/md2docx 帮我把 report.md 转成 Word，正文宋体小四单倍行距首行缩进，标题黑体三号居中，封面标题"数字电路实验报告"作者张三，目录到二级`

**你的处理**：

1. 输入文件 = `report.md`，输出 = `report.docx`
2. 修改默认模板中的以下字段：
   - `cover.title: "数字电路实验报告"`
   - `cover.author: "张三"`
   - `toc.level: 2`
   - `styles.body` 保持默认（宋体12pt, 单倍行距, 首行缩进2, 两端对齐——用户说的就是默认值）
   - `styles.headings.h1` 保持默认（黑体16pt, 居中, 加粗——用户说的就是默认值）
3. 写入 `config.yaml`
4. 执行 `md2docx report.md -c config.yaml -o report.docx`
5. 提醒更新目录域

---

## 注意事项

- 用户只说「宋体小四」→ 修改正文 `font_name` 和 `font_size`，其他正文字段不动
- 用户说「标题」未指明级别 → 默认指 h1
- 用户说「表格五号字」→ 修改 `styles.table.font_size`
- 用户未提到的元素（如代码块、引用块）→ 保持默认模板值，不要擅自修改
- 如果用户需求不明确（如只说「好看一点」），主动询问字体、字号、行距偏好
