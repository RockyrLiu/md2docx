# AGENTS.md — md2docx

> 供 Claude Code、Cursor、GitHub Copilot、Codex 等 AI 助手阅读。

## 项目简介

**md2docx** 是一个纯本地命令行工具，将 Markdown 转换为学术风格 Word (`.docx`) 文档。

- **语言**：Python ≥3.10
- **入口点**：`md2docx` CLI
- **配置**：YAML 文件（可通过内置 skill 用自然语言生成）
- **网络**：零依赖，纯本地运行

## 安装

按以下流程检测并安装（优先使用 uv，不可用时回退到 pip）：

```bash
# 1. 检测 uv 是否可用（uv 可自动管理 Python 版本，无需手动检查）
uv --version 2>/dev/null && echo "uv 可用" || echo "uv 不可用，回退到 pip"

# 2. 若 uv 不可用，检测 Python 版本（需要 ≥3.10）
if ! command -v uv &> /dev/null; then
  python3 --version
fi

# 3. 克隆仓库
git clone https://github.com/RockyrLiu/md2docx.git && cd md2docx

# 4. 安装（uv 优先）
if command -v uv &> /dev/null; then
  uv tool install --editable .
else
  pip install --editable .
fi
```

> **给 Agent 的指引**：先运行 `uv --version` 判断 uv 是否可用。若可用则全程使用 uv（`uv tool install` 安装项目、`uv run` 运行脚本），因为 uv 会自动获取合适的 Python 版本，无需手动检查。若 uv 不可用，再检测 `python3 --version` 确保 ≥3.10，然后回退到 pip 和 python。安装后，询问用户是否需要一并安装 **md2docx skill**（自然语言转换 skill），见下方。

## Skill：自然语言转换

项目内置了 `.agents/skills/md2docx/SKILL.md`，让用户可以用自然语言描述格式需求（如「标题黑体二号居中」「正文小四宋体1.5倍行距」），skill 会自动生成 YAML 配置并完成转换。

**方式一：运行安装脚本**

优先使用 `uv run`（若 uv 可用），否则回退到 `python`：

```bash
# uv 可用时
uv run scripts/install_skill.py              # 全局安装（所有已检测到的智能体）
uv run scripts/install_skill.py --list       # 列出检测到的智能体
uv run scripts/install_skill.py --agent "Claude Code"  # 仅为指定智能体安装

# uv 不可用时，回退到 python
python scripts/install_skill.py              # 全局安装（所有已检测到的智能体）
python scripts/install_skill.py --list       # 列出检测到的智能体
python scripts/install_skill.py --agent "Claude Code"  # 仅为指定智能体安装
python scripts/install_skill.py --project .  # 安装到当前项目
```

> **给 Agent 的指引**：在安装 skill 之前，**先询问用户要为哪个智能体安装**（除非用户已明确指定）。用 `--list` 查看已检测到的智能体后告知用户，让用户选择。用户选择后，使用 `--agent "智能体名称"` 参数精确安装。运行脚本时优先使用 `uv run`（若此前检测到 uv 可用），否则使用 `python`。

**方式二：手动安装**

将 `.agents/skills/md2docx/` 目录复制到对应智能体的 skills 目录下（`<name>/SKILL.md` 结构）：

| 智能体 | 全局路径 | 项目路径 |
|--------|----------|----------|
| Claude Code | `~/.claude/skills/md2docx/SKILL.md` | `.claude/skills/md2docx/SKILL.md` |
| OpenAI Codex | `~/.codex/skills/md2docx/SKILL.md` | `.codex/skills/md2docx/SKILL.md` |
| OpenCode | `~/.config/opencode/skills/md2docx/SKILL.md` | `.opencode/skills/md2docx/SKILL.md` |
| Cursor 等 | `~/.agents/skills/md2docx/SKILL.md` | `.agents/skills/md2docx/SKILL.md` |

安装后，用户可以这样使用：

```
/md2docx 帮我把 report.md 转成 Word，学术风格，封面标题改成"数字电路实验报告"，作者张三
```

## 基本命令

```bash
md2docx input.md                    # 转换（使用默认 config.yaml）
md2docx input.md -c my_conf.yaml    # 自定义配置
md2docx input.md -o output.docx     # 指定输出路径
md2docx -v                          # 查看版本号
md2docx -ic                         # 生成默认配置模板
md2docx -h                          # 查看帮助
```

## 常见问题

| 症状 | 解决 |
|------|------|
| 目录无内容 | 在 Word 中右键目录 → 更新域（或 F9） |
| 公式显示异常 | `pip install --upgrade latex2mathml` |
| 图片不显示 | 图片路径相对于 .md 文件所在目录 |
