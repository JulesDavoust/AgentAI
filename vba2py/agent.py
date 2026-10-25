"""Configuration et lancement de l'agent (Claude Agent SDK) pour un projet VBA."""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

from claude_agent_sdk import (
    AssistantMessage,
    ClaudeAgentOptions,
    HookMatcher,
    ResultMessage,
    TextBlock,
    ToolUseBlock,
    query,
)

from .guardrails import Guardrails

PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"

DEFAULT_MODEL = "claude-opus-5-5"
AGENT_TOOLS = ["Read", "Glob", "Grep", "Write", "Edit", "Bash", "TodoWrite"]

USER_PROMPT = (
    "Traduis en Python le projet VBA qui se trouve dans `_source_vba/`, en suivant tes "
    "instructions : code Python, données fictives, batterie de tests qui passe, "
    "requirements.txt et RAPPORT.md."
)


@dataclass
class AgentRun:
    success: bool
    summary: str = ""
    turns: int = 0
    cost_usd: float | None = None
    denials: list = field(default_factory=list)


def build_system_prompt() -> str:
    system = (PROMPTS_DIR / "system.md").read_text(encoding="utf-8")
    mapping = (PROMPTS_DIR / "mapping.md").read_text(encoding="utf-8")
    rapport = (PROMPTS_DIR / "rapport_modele.md").read_text(encoding="utf-8")
    return system.replace("{{MAPPING}}", mapping).replace("{{RAPPORT_MODELE}}", "```markdown\n" + rapport + "```")


def build_options(
    project_dir: Path,
    journal_path: Path | None,
    model: str = DEFAULT_MODEL,
    max_turns: int = 60,
    max_budget_usd: float | None = 5.0,
) -> ClaudeAgentOptions:
    guardrails = Guardrails(project_dir, journal_path)
    python_bin = str(Path(sys.executable).parent)
    return ClaudeAgentOptions(
        # Outils intégrés disponibles : tout le reste (web, notebooks, sous-agents) n'existe pas pour l'agent.
        tools=AGENT_TOOLS,
        # Auto-approuvés, mais chaque appel passe d'abord par le hook des garde-fous.
        allowed_tools=AGENT_TOOLS,
        disallowed_tools=["WebFetch", "WebSearch", "NotebookEdit", "Agent", "Task"],
        # Tout ce qui n'est pas pré-approuvé est refusé, sans jamais demander.
        permission_mode="dontAsk",
        hooks={"PreToolUse": [HookMatcher(matcher=None, hooks=[guardrails.pre_tool_use_hook])]},
        # Isolation : aucun réglage, CLAUDE.md ou skill de la machine n'est chargé.
        setting_sources=[],
        skills=[],
        system_prompt={"type": "preset", "preset": "claude_code", "append": build_system_prompt()},
        cwd=str(project_dir),
        # `python` désigne l'environnement de vba2py, où les librairies autorisées sont installées.
        env={"PATH": python_bin + os.pathsep + os.environ.get("PATH", ""), "MPLBACKEND": "Agg"},
        model=model,
        max_turns=max_turns,
        max_budget_usd=max_budget_usd,
    )


async def translate_project(project_dir: Path, journal_path: Path | None, verbose: bool = True, **kwargs) -> AgentRun:
    options = build_options(project_dir, journal_path, **kwargs)
    run = AgentRun(success=False)
    async for message in query(prompt=USER_PROMPT, options=options):
        if isinstance(message, AssistantMessage) and verbose:
            for block in message.content:
                if isinstance(block, ToolUseBlock):
                    print(f"   → {block.name} {_describe(block.input)}")
                elif isinstance(block, TextBlock) and block.text.strip():
                    print(f"   {block.text.strip()[:300]}")
        elif isinstance(message, ResultMessage):
            run.success = not message.is_error
            run.summary = message.result or ""
            run.turns = message.num_turns
            run.cost_usd = message.total_cost_usd
            run.denials = list(message.permission_denials or [])
            if message.is_error:
                run.summary = run.summary or f"arrêt : {message.subtype}"
    return run


def _describe(tool_input: dict) -> str:
    for key in ("command", "file_path", "pattern"):
        if key in tool_input:
            return str(tool_input[key])[:120]
    return ""
