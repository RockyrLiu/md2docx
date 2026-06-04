# AGENTS.md — md2docx

> 供 Claude Code、Cursor、GitHub Copilot、Codex 等 AI 助手阅读。

## 项目简介

**md2docx** 是一个纯本地命令行工具，将 Markdown 转换为学术风格 Word (`.docx`) 文档。

- **语言**：Python ≥3.10
- **入口点**：`md2docx` CLI
- **配置**：YAML 文件（可通过内置 skill 用自然语言生成）
- **网络**：零依赖，纯本地运行

## 安装

检测并安装（按优先级尝试 uv → pip）：

```bash
# 检测 Python 版本（需要 ≥3.10）
python3 --version

# 检测是否已安装
md2docx --help 2>/dev/null && echo "已安装" || echo "需要安装"

# 安装
git clone https://github.com/RockyrLiu/md2docx.git && cd md2docx
uv tool install --editable . 2>/dev/null || \
  pip install --editable . 2>/dev/null
```

安装后，可以询问用户是否需要一并安装 **md2docx skill**（自然语言转换 skill），见下方。

## Skill：自然语言转换

项目内置了 `.agents/skills/md2docx.md`，让用户可以用自然语言描述格式需求（如「标题黑体二号居中」「正文小四宋体1.5倍行距」），skill 会自动生成 YAML 配置并完成转换。

**安装方式**：将 `.agents/skills/md2docx.md` 复制到用户项目的 `.agents/skills/` 目录下即可。用户即可通过 `/md2docx` 使用。

安装后，用户可以这样使用：

```
/md2docx 帮我把 report.md 转成 Word，学术风格，封面标题改成"数字电路实验报告"，作者张三
```

## 基本命令

```bash
md2docx input.md                    # 转换（使用默认 config.yaml）
md2docx input.md -c my_conf.yaml    # 自定义配置
md2docx input.md -o output.docx     # 指定输出路径
md2docx -ic                         # 生成默认配置模板
md2docx -h                          # 查看帮助
```

## 常见问题

| 症状 | 解决 |
|------|------|
| 目录无内容 | 在 Word 中右键目录 → 更新域（或 F9） |
| 公式显示异常 | `pip install --upgrade latex2mathml` |
| 图片不显示 | 图片路径相对于 .md 文件所在目录 |
