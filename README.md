<p align="center">
  <h1 align="center">📄 md2docx</h1>
  <p align="center"><strong>Markdown → DOCX</strong> · 学术风格 · 样式可定制</p>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-≥3.10-blue?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/uv-ready-e6b422?logo=pypi&logoColor=white" alt="uv">
</p>

---

**md2docx** 是一款命令行工具，将 Markdown 文件一键转换为符合学术规范、样式高度可定制的 Word（`.docx`）文档。适合撰写实验报告、课程论文、读书笔记等场景。

## ✨ 功能特性

| 特性 | 说明 |
|------|------|
| 📔 **封面** | 可配置的报告标题、姓名、班级学号、指导教师、日期 |
| 📑 **目录** | 自动插入 Word TOC 域，在 Word 中右键即可更新 |
| 📐 **多级标题** | H1–H6 映射为 Word 内置标题样式，各级字体/字号/加粗/对齐独立可配 |
| 📝 **正文排版** | 中英文双字体（宋体 + Times New Roman）、1.5 倍行距、首行缩进、两端对齐 |
| 📊 **学术三线表** | GFM 表格 → 三线表（顶线 1.5pt、表头下线 0.75pt、底线 1.5pt） |
| 🖼️ **图片** | `![alt](path)` 自动居中嵌入，宽高自适应页面 |
| 📋 **列表** | 有序/无序列表，支持嵌套层级 |
| 🧮 **LaTeX 公式** | 行内 `$...$` 和块级 `$$...$$` 转换为 Word 原生 OMML 公式 |
| 💻 **代码块** | 等宽字体 + 灰色背景 |
| 💬 **引用块** | 左侧竖线 + 楷体斜体 |
| ⚙️ **YAML 驱动** | 所有样式通过单个配置文件控制，无需修改代码 |

## 📦 安装

### 环境要求

- **Python** ≥ 3.10
- **[uv](https://docs.astral.sh/uv/)** — 现代 Python 包管理器

### 方式一：全局安装（推荐日常使用）

安装为全局命令行工具后，可在任意目录直接调用 `md2docx`：

```bash
# 克隆项目
git clone https://github.com/example/md2docx.git
cd md2docx

# 全局安装（发布模式）
uv tool install .

# 或安装为可编辑模式（源码改动即时生效，推荐开发）
uv tool install --editable .
```

安装完成后，在任意目录即可使用：

```bash
md2docx input.md                  # 使用默认 config.yaml
md2docx input.md -c my_conf.yaml  # 自定义配置
md2docx input.md -o output.docx   # 指定输出路径
md2docx -h                        # 查看帮助
```

**管理全局安装的工具：**

```bash
uv tool list              # 查看已安装的工具
uv tool uninstall md2docx # 卸载
uv tool upgrade md2docx   # 升级（从 PyPI/git 安装时可用）
```

### 方式二：本地开发环境

```bash
# 克隆项目
git clone https://github.com/example/md2docx.git
cd md2docx

# 安装依赖
uv sync
```

通过 `uv run` 运行：

```bash
uv run md2docx example/sample.md
```

## 🚀 快速开始

```bash
# 1. 克隆并安装
git clone https://github.com/example/md2docx.git && cd md2docx
uv tool install --editable .

# 2. 编辑配置文件（填入你的信息）
#    用任意编辑器修改 config.yaml 中的封面信息和样式

# 3. 转换
md2docx example/sample.md

# 4. 查看输出
#    打开 example/sample.docx，在 Word 中右键目录 → 更新域
```

## ⚙️ 配置文件

默认从当前目录的 `config.yaml` 加载，可通过 `-c` 指定自定义路径。缺失的键会自动回退到内置默认值。

```yaml
# === 封面 ===
cover:
  enabled: true                   # 是否生成封面
  title: "报告标题"
  subtitle: ""                    # 副标题（可选）
  author: "姓名"
  class_info: "班级学号"
  teacher: "任课老师"
  department: ""                  # 院系（可选）
  date: ""                        # 留空使用当天日期

# === 目录 ===
toc:
  enabled: true
  level: 3                        # 包含 1-3 级标题

# === 样式（部分示例） ===
styles:
  body:                           # 正文
    font_name: "宋体"
    font_name_ascii: "Times New Roman"
    font_size: 12                 # 小四
    line_spacing: 1.5
    first_line_indent: 2          # 首行缩进（字符数）
    alignment: "justify"

  headings:                       # 标题（h1–h6 可逐级配置）
    h1:
      font_name: "黑体"
      font_size: 22               # 二号
      bold: true
      alignment: "center"
    h2:
      font_name: "黑体"
      font_size: 16               # 三号
      bold: true

  table:                          # 表格
    font_name: "宋体"
    font_size: 10                 # 五号
    header_bold: true
    alignment: "center"

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

> 完整配置文件见 [config.yaml](config.yaml)。

## 📐 中文字号对照

| 字号 | pt | 典型用途 |
|------|-----|---------|
| 二号 | 22 | 一级标题 (H1) |
| 三号 | 16 | 二级标题 (H2) |
| 四号 | 14 | 三级标题 (H3) |
| 小四 | 12 | 正文 / 四级标题 |
| 五号 | 10 | 表格 / 脚注 |

## 📝 支持的 Markdown 语法

| 元素 | 语法 | 渲染效果 |
|------|------|---------|
| 标题 | `# H1` ∼ `###### H6` | Word 内置标题样式 |
| 粗体 | `**text**` | 加粗 |
| 斜体 | `*text*` | 斜体 |
| 行内代码 | `` `code` `` | Consolas 等宽 |
| 链接 | `[text](url)` | 超链接 |
| 图片 | `![alt](path)` | 居中嵌入（相对路径相对于 .md 文件） |
| 无序列表 | `- item` | 缩进 + 项目符号 |
| 有序列表 | `1. item` | 缩进 + 数字编号 |
| 表格 | GFM 表格 | 学术三线表 |
| 代码块 | ```` ```lang ```` | 灰色背景 + 等宽字体 |
| 引用 | `> quote` | 左侧竖线 + 楷体 |
| 分隔线 | `---` | 水平线 |
| 行内公式 | `$E=mc^2$` | OMML 原生公式 |
| 块级公式 | `$$\frac{a}{b}$$` | 居中 OMML 公式 |

## 🛠️ 技术栈

| 库 | 用途 |
|----|------|
| [python-docx](https://python-docx.readthedocs.io/) | DOCX 文档生成与样式管理 |
| [mistune](https://mistune.lepture.com/) v3 | Markdown 解析为 AST |
| [latex2mathml](https://pypi.org/project/late2mathml/) | LaTeX → MathML 转换 |
| [lxml](https://lxml.de/) | OMML XML 构建与注入 |
| [Pillow](https://python-pillow.org/) | 图片尺寸读取 |
| [PyYAML](https://pyyaml.org/) | YAML 配置文件解析 |

## 📁 项目结构

```
md2docx/
├── pyproject.toml               # 项目元数据 & 依赖声明
├── config.yaml                  # 默认配置文件
├── README.md
├── md2docx/                     # 核心包
│   ├── __init__.py
│   ├── cli.py                   # CLI 命令行入口
│   ├── config.py                # YAML 加载 & 配置数据类
│   ├── parser.py                # Markdown → AST
│   ├── styles.py                # Word 样式管理器
│   ├── math_handler.py          # LaTeX → OMML
│   └── builder.py               # 文档构建器（核心编排逻辑）
└── example/
    ├── sample.md                # 示例输入
    ├── sample.docx              # 示例输出
    └── image.png                # 示例图片
```

## 📄 License

[MIT](LICENSE)
