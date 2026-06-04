#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# ///

"""md2docx Skill Auto-Installer.

Detects AI agents (Claude Code, Codex, OpenCode, Cursor etc.) on the user's system
and installs the md2docx skill to the appropriate locations for each agent.

Usage:
    python install_skill.py              # Install to all detected agents globally
    python install_skill.py --dry-run    # Preview what would be installed
    python install_skill.py --project /path/to/project  # Install to a specific project
    python install_skill.py --list       # List all detected agents
    python install_skill.py --agent "Claude Code"  # Install only for a specific agent
    python install_skill.py --agent "Claude Code" "OpenAI Codex"  # Multiple agents
"""

import argparse
import shutil
from pathlib import Path

# ── Agent Definitions ──────────────────────────────────────────────────────

AGENT_DEFINITIONS = {
    "Claude Code": {
        "home_dirs": [".claude"],
        "skill_subdir": "skills",
        "description": "Claude Code (Anthropic)",
    },
    "OpenAI Codex": {
        "home_dirs": [".codex"],
        "skill_subdir": "skills",
        "description": "OpenAI Codex CLI",
    },
    "OpenCode": {
        "home_dirs": [".config/opencode"],
        "skill_subdir": "skills",
        "description": "OpenCode (open source agent)",
    },
    "Cursor / AGENTS.md": {
        "home_dirs": [".agents"],
        "skill_subdir": "skills",
        "description": "Cursor & agents following AGENTS.md convention",
    },
    "Cline": {
        "home_dirs": [".cline"],
        "skill_subdir": "skills",
        "description": "Cline (VS Code extension)",
    },
    "Amazon Q": {
        "home_dirs": [".amazonq"],
        "skill_subdir": "skills",
        "description": "Amazon Q Developer",
    },
    "GitHub Copilot": {
        "home_dirs": [".github"],
        "skill_subdir": "skills",
        "description": "GitHub Copilot (project-level skills)",
    },
}

# Additional project-level directories to check (for agent IDE integrations)
PROJECT_AGENT_DIRS = [
    ".agents",
    ".claude",
    ".opencode",
    ".cursor",
    ".codex",
    ".cline",
    ".amazonq",
    ".github",
]

# OpenCode project-level skills directory
OPENCODE_PROJECT_DIR = ".opencode"


def get_source_skills_dir() -> Path:
    """Return the path to the source .agents/skills/ directory in the project."""
    return Path(__file__).resolve().parent.parent / ".agents" / "skills"


def get_skill_dirs(source_dir: Path) -> list[Path]:
    """Return list of skill source directories (each containing SKILL.md)."""
    if not source_dir.is_dir():
        return []
    skills = []
    for entry in source_dir.iterdir():
        if entry.is_dir() and (entry / "SKILL.md").exists():
            skills.append(entry)
    return skills


def resolve_agent_names(
    names: list[str] | None = None,
) -> dict[str, dict]:
    """Resolve agent names to their definitions.

    Args:
        names: List of agent names to include. If None or empty, return all.

    Returns:
        Filtered dict of {name: config} from AGENT_DEFINITIONS.
    """
    if not names:
        return dict(AGENT_DEFINITIONS)

    # Build a case-insensitive lookup
    name_lower_map: dict[str, str] = {
        k.lower(): k for k in AGENT_DEFINITIONS
    }

    resolved: dict[str, dict] = {}
    for name in names:
        key = name_lower_map.get(name.lower())
        if key:
            resolved[key] = AGENT_DEFINITIONS[key]
        else:
            # Fuzzy suggestion: find closest match
            import difflib
            close = difflib.get_close_matches(
                name.lower(), list(name_lower_map), n=1, cutoff=0.4
            )
            hint = f" Did you mean '{name_lower_map[close[0]]}'?" if close else ""
            print(
                f"Warning: Unknown agent '{name}'.{hint} "
                f"Available: {list(AGENT_DEFINITIONS)}"
            )

    return resolved


def detect_home_agents(
    home: Path,
    agent_filter: dict[str, dict] | None = None,
) -> dict[str, list[Path]]:
    """Detect which AI agents are installed in the user's home directory.

    Args:
        home: User's home directory.
        agent_filter: If provided, only check these agents (from resolve_agent_names).

    Returns dict: {agent_name: [skill_base_dir, ...]}
    Each skill_base_dir is the directory where skill subdirectories live
    (e.g. ~/.claude/skills/).
    """
    agents_to_check = agent_filter if agent_filter is not None else AGENT_DEFINITIONS
    detected: dict[str, list[Path]] = {}
    for agent_name, config in agents_to_check.items():
        found: list[Path] = []
        for home_dir in config["home_dirs"]:
            skill_path = home / home_dir / config["skill_subdir"]
            if skill_path.is_dir():
                found.append(skill_path)
        if found:
            detected[agent_name] = found
    return detected


def detect_project_agents(
    project_path: Path,
    agent_filter: dict[str, dict] | None = None,
) -> list[Path]:
    """Detect agent skill directories within a project.

    Args:
        project_path: Root of the project to scan.
        agent_filter: If provided, only check directories relevant to these agents.

    Returns list of skill base directories inside the project
    (e.g. /project/.agents/skills/).
    """
    # If a filter is given, only check the home_dirs of those agents
    if agent_filter is not None:
        dirs_to_check: set[str] = set()
        for config in agent_filter.values():
            dirs_to_check.update(config["home_dirs"])
    else:
        dirs_to_check = set(PROJECT_AGENT_DIRS)

    found: list[Path] = []
    for d in sorted(dirs_to_check):
        skill_dir = project_path / d / "skills"
        if skill_dir.is_dir() or (project_path / d).is_dir():
            found.append(skill_dir)
    return found


def install_skill_to_target(
    skill_source: Path,
    target_base: Path,
    dry_run: bool = False,
) -> tuple[str, Path]:
    """Copy a skill directory to a target agent skills directory.

    Returns (action, target_path).
    """
    skill_name = skill_source.name
    target = target_base / skill_name

    if not dry_run:
        if target.exists():
            shutil.rmtree(target)
        shutil.copytree(skill_source, target)

    return ("installed" if not target.exists() else "updated", target)


def install_skills(
    skill_dirs: list[Path],
    agent_targets: dict[str, list[Path]] | None = None,
    project_targets: list[Path] | None = None,
    dry_run: bool = False,
) -> None:
    """Install skills to all detected agent directories.

    Args:
        skill_dirs: List of source skill directories to install.
        agent_targets: Dict of {agent_name: [target_base_dirs]} for global installs.
        project_targets: List of target_base_dirs for project-level installs.
        dry_run: If True, only show what would be done.
    """
    if not skill_dirs:
        print("No skills found in source directory.")
        return

    total_results: list[tuple[str, str, str, Path]] = []

    if agent_targets:
        for agent_name, base_dirs in agent_targets.items():
            for base_dir in base_dirs:
                for skill_dir in skill_dirs:
                    action, target = install_skill_to_target(
                        skill_dir, base_dir, dry_run=dry_run
                    )
                    total_results.append((action, skill_dir.name, agent_name, target))

    if project_targets:
        for base_dir in project_targets:
            for skill_dir in skill_dirs:
                action, target = install_skill_to_target(
                    skill_dir, base_dir, dry_run=dry_run
                )
                total_results.append(
                    (action, skill_dir.name, "(project)", target)
                )

    if dry_run:
        print("[DRY RUN] The following would be performed:\n")
    else:
        print()

    if not total_results:
        print("No agent directories detected. Nothing to install.")
        return

    # Group results by agent for cleaner output
    for action, skill_name, agent_name, target in total_results:
        icon = "+" if action == "installed" else "~"
        print(f"  {icon} [{agent_name}] {skill_name} -> {target}")

    if not dry_run:
        print(f"\nDone. {len(total_results)} skill(s) installed/updated.")


def list_detected_agents(
    home: Path,
    project: Path | None = None,
    agent_filter: dict[str, dict] | None = None,
) -> None:
    """Print a summary of detected AI agents."""
    print(f"Home directory: {home}\n")
    print("Detected AI agents:\n")

    home_agents = detect_home_agents(home, agent_filter=agent_filter)
    if home_agents:
        for name, dirs in home_agents.items():
            desc = AGENT_DEFINITIONS[name]["description"]
            print(f"  [Global] {name} ({desc})")
            for d in dirs:
                print(f"    -> {d}")
        print()
    else:
        agents_label = ", ".join(agent_filter) if agent_filter else "any"
        print(f"  (no matching agents detected in home directory: {agents_label})\n")

    if project and project.is_dir():
        project_dirs = detect_project_agents(project, agent_filter=agent_filter)
        if project_dirs:
            print(f"  [Project] {project}")
            for d in project_dirs:
                installed_skills = list(d.glob("*/SKILL.md"))
                count = len(installed_skills)
                print(f"    -> {d}  ({count} skill(s) present)")
        else:
            print(f"  [Project] {project}  (no agent directories found)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="md2docx Skill Auto-Installer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python install_skill.py                    # Install to all detected agents globally
  python install_skill.py --dry-run          # Preview installation
  python install_skill.py --project ./myproj # Install to project only
  python install_skill.py --list             # List detected agents
  python install_skill.py --all              # Install globally AND to current project
  python install_skill.py --agent "Claude Code"  # Install only for a specific agent
        """,
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview what would be installed without making changes",
    )
    parser.add_argument(
        "--project",
        type=Path,
        default=None,
        help="Install skills to a specific project directory (creates .agents/skills/ etc.)",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Install both globally and to the current directory as a project",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List all detected AI agents without installing",
    )
    parser.add_argument(
        "--agent",
        nargs="+",
        default=None,
        metavar="NAME",
        help="Install only for specific agent(s). Use names from --list. "
        "Example: --agent \"Claude Code\" \"OpenAI Codex\"",
    )

    args = parser.parse_args()
    home = Path.home()
    source_dir = get_source_skills_dir()
    skill_dirs = get_skill_dirs(source_dir)

    # Resolve agent filter (None = all agents)
    agent_filter = resolve_agent_names(args.agent)

    if args.list:
        project = args.project or (Path.cwd() if args.all else None)
        list_detected_agents(home, project, agent_filter=agent_filter)
        print(f"\nSource skills: {[d.name for d in skill_dirs]}")
        return

    # Determine targets
    agent_targets: dict[str, list[Path]] | None = None
    project_targets: list[Path] | None = None

    if args.project:
        project_targets = detect_project_agents(
            args.project.resolve(), agent_filter=agent_filter
        )
    elif args.all:
        agent_targets = detect_home_agents(home, agent_filter=agent_filter)
        project_targets = detect_project_agents(
            Path.cwd(), agent_filter=agent_filter
        )
    else:
        agent_targets = detect_home_agents(home, agent_filter=agent_filter)

    install_skills(
        skill_dirs,
        agent_targets=agent_targets,
        project_targets=project_targets,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
