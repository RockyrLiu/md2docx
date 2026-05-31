# md2docx — Markdown to DOCX Converter

将 Markdown 文件转换为符合学术规范、样式可定制的 `.docx` 文档。

## 功能特性

-   **封面** — 可配置的报告名称、姓名、班级学号、任课老师、日期
-   **目录** — 自动插入 Word TOC 域，打开后在 Word 中右键更新即可
-   **多级标题** — H1–H6 映射为 Word 内置标题样式，各级字体/字号/加粗/对齐均可独立配置
-   **正文排版** — 中英文双字体（宋体 + Times New Roman）、1.5 倍行距、首行缩进、两端对齐
-   **学术三线表** — GFM 表格自动渲染为三线表（顶线 1.5 pt、表头下线 0.75 pt、底线 1.5 pt、无竖线）
-   **图片** — `![alt](path)` 自动居中嵌入，宽高自适应页面
-   **列表** — 有序/无序列表，支持嵌套层级
-   **LaTeX 公式** — 行内 `$...$` 和块级 `$$...$$` 转换为 Word 原生 OMML 公式
-   **代码块** — 等宽字体 + 灰色背景
-   **引用块** — 左侧竖线 + 楷体斜体
-   **YAML 驱动** — 所有样式通过单个配置文件控制，无需修改代码

## 快速开始

### 环境要求

-   Python ≥ 3.10
-   [uv](https://docs.astral.sh/uv/)（现代 Python 包管理器）

### 安装

```bash
# 克隆项目
git clone <repo-url> && cd md2docx

# 安装依赖
uv sync
```

### 使用

```bash
# 使用默认配置（config.yaml），输出与输入同名 .docx
uv run md2docx example/sample.md

# 自定义配置文件
uv run md2docx input.md -c my_config.yaml

# 指定输出路径
uv run md2docx input.md -o output.docx

# 查看帮助
uv run md2docx -h
```

## 配置文件

默认配置文件为项目根目录下的 [config.yaml](config.yaml)。完整结构如下：

```yaml
# === 封面 ===
cover:
  enabled: true                   # 是否生成封面
  title: "报告标题"               # 报告名称
  subtitle: ""                    # 副标题（可选）
  author: "姓名"                  # 作者
  class_info: "班级学号"          # 班级/学号
  teacher: "任课老师"             # 指导教师
  department: ""                  # 院系（可选）
  date: ""                        # 日期，留空使用当天

# === 目录 ===
toc:
  enabled: true
  level: 3                        # 包含的标题级别数

# === 样式 ===
styles:
  body:                           # 正文
    font_name: "宋体"
    font_name_ascii: "Times New Roman"
    font_size: 12                 # 小四
    line_spacing: 1.5
    first_line_indent: 2          # 首行缩进字符数
    alignment: "justify"

  headings:                       # 标题（可按 h1–h6 逐级配置）
    h1:
      font_name: "黑体"
      font_size: 22               # 二号
      bold: true
      alignment: "center"
    h2:
      font_name: "黑体"
      font_size: 16               # 三号
      bold: true
    # h3–h6 同理

  table:                          # 表格
    font_name: "宋体"
    font_size: 10                 # 五号
    header_bold: true
    alignment: "center"

  list:                           # 列表
    font_name: "宋体"
    font_size: 12
    line_spacing: 1.5

  code:                           # 代码块
    font_name: "Consolas"
    font_size: 10
    background_color: "F2F2F2"

  blockquote:                     # 引用块
    font_name: "楷体"
    font_size: 12
    italic: true

# === 页面设置 ===
page:
  size: "A4"
  margin_top: 2.54                # cm
  margin_bottom: 2.54
  margin_left: 3.18
  margin_right: 3.18
```

## 中文字号对照

| 字号 | pt | 典型用途 |
|------|-----|---------|
| 二号 | 22 | 一级标题 (H1) |
| 三号 | 16 | 二级标题 (H2) |
| 四号 | 14 | 三级标题 (H3) |
| 小四 | 12 | 正文 / 四级标题 |
| 五号 | 10.5 | 表格 / 脚注 |

## 支持的 Markdown 语法

| 元素 | 语法 | 说明 |
|------|------|------|
| 标题 | `# H1` – `###### H6` | 支持 ATX 风格 |
| 粗体 | `**bold**` | |
| 斜体 | `*italic*` | |
| 行内代码 | `` `code` `` | Consolas 字体 |
| 链接 | `[text](url)` | |
| 图片 | `![alt](path)` | 相对路径相对于 .md 文件 |
| 无序列表 | `- item` | 支持嵌套 |
| 有序列表 | `1. item` | 支持嵌套 |
| 表格 | GFM 表格 | 三线表输出 |
| 代码块 | `` ```lang `` | 灰色背景 |
| 引用 | `> quote` | 左侧竖线 + 楷体 |
| 分隔线 | `---` | |
| 行内公式 | `$E=mc^2$` | OMML 原生公式 |
| 块级公式 | `$$\frac{a}{b}$$` | 居中 OMML 公式 |

## 技术栈

-   [python-docx](https://python-docx.readthedocs.io/) — DOCX 文档生成
-   [mistune](https://mistune.lepture.com/) v3 — Markdown 解析
-   [latex2mathml](https://pypi.org/project/latex2mathml/) — LaTeX → MathML 转换
-   [lxml](https://lxml.de/) — OMML XML 构建与注入

## 项目结构

```
md2docx/
├── pyproject.toml
├── config.yaml                 # 默认配置
├── md2docx/                    # 核心包
│   ├── cli.py                  # CLI 入口
│   ├── config.py               # YAML 加载 & 数据类
│   ├── parser.py               # Markdown → AST
│   ├── styles.py               # Word 样式管理
│   ├── math_handler.py         # LaTeX → OMML
│   └── builder.py              # 文档构建器
└── example/
    ├── sample.md               # 示例输入
    └── sample.docx             # 示例输出
```

## License

MIT
