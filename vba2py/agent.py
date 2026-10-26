"""Configuration et lancement de l'agent (Claude Agent SDK) pour un projet VBA."""

from __future__ import annotations

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
from .outils import NOM_SERVEUR, creer_serveur, noms_outils

PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"

DEFAULT_MODEL = "claude-opus-5-5"
# Aucun terminal : l'exécution passe par les trois outils de vba2py (voir outils.py)
AGENT_TOOLS = ["Read", "Glob", "Grep", "Write", "Edit", "TodoWrite"]

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


# Estimation grossière du coût d'une traduction, tirée des premiers essais (avec marge) :
# coût fixe + coût par millier de lignes de VBA, pour un modèle de la gamme Sonnet.
COUT_FIXE_USD = 1.0
COUT_PAR_MILLE_LIGNES_USD = 1.5
FACTEURS_MODELE = {"claude-haiku": 0.25, "claude-sonnet": 1.0, "claude-opus": 2.0, "claude-fable": 5.0}


def lignes_de_vba(project_dir: Path) -> int:
    total = 0
    for fichier in (project_dir / "_source_vba").glob("*"):
        if fichier.suffix.lower() in {".bas", ".cls", ".frm"}:
            total += sum(1 for l in fichier.read_text(encoding="utf-8").splitlines()
                         if l.strip() and not l.lstrip().startswith(("'", "Attribute ")))
    return total


def estimer_budget(nb_lignes: int, modele: str) -> float:
    facteur = next((f for prefixe, f in FACTEURS_MODELE.items() if modele.startswith(prefixe)), 2.0)
    return round((COUT_FIXE_USD + COUT_PAR_MILLE_LIGNES_USD * nb_lignes / 1000) * facteur, 2)


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
    reference_vba: bool = False,
) -> ClaudeAgentOptions:
    guardrails = Guardrails(project_dir, journal_path)
    return ClaudeAgentOptions(
        # Outils intégrés disponibles : ni terminal, ni web, ni notebooks, ni sous-agents.
        tools=AGENT_TOOLS,
        # Outils d'exécution de vba2py (tests, syntaxe, script), exécutés par vba2py lui-même.
        mcp_servers={NOM_SERVEUR: creer_serveur(project_dir, guardrails, reference_vba)},
        strict_mcp_config=True,
        # Auto-approuvés, mais chaque appel passe d'abord par le hook des garde-fous.
        allowed_tools=AGENT_TOOLS + noms_outils(reference_vba),
        disallowed_tools=["Bash", "PowerShell", "WebFetch", "WebSearch", "NotebookEdit", "Agent", "Task"],
        # Tout ce qui n'est pas pré-approuvé est refusé, sans jamais demander.
        permission_mode="dontAsk",
        hooks={"PreToolUse": [HookMatcher(matcher=None, hooks=[guardrails.pre_tool_use_hook])]},
        # Isolation : aucun réglage, CLAUDE.md ou skill de la machine n'est chargé.
        setting_sources=[],
        skills=[],
        system_prompt={"type": "preset", "preset": "claude_code", "append": build_system_prompt()},
        cwd=str(project_dir),
        model=model,
        max_turns=max_turns,
        max_budget_usd=max_budget_usd,
    )


async def translate_project(project_dir: Path, journal_path: Path | None, verbose: bool = True, **kwargs) -> AgentRun:
    options = build_options(project_dir, journal_path, **kwargs)
    prompt = USER_PROMPT
    if kwargs.get("reference_vba"):
        prompt += (" L'outil `executer_macro_vba` est disponible : utilise-le pour comparer ta traduction "
                   "à la macro d'origine (voir « Comparaison avec la macro d'origine »).")
    run = AgentRun(success=False)
    async for message in query(prompt=prompt, options=options):
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
            if message.subtype == "error_max_budget_usd":
                run.summary = (run.summary + "\n\n" if run.summary else "") + \
                    "⚠️ Budget atteint : la traduction a été interrompue. Relancez avec un --budget plus élevé."
            elif message.is_error:
                run.summary = run.summary or f"arrêt : {message.subtype}"
    return run


def _describe(tool_input: dict) -> str:
    for key in ("file_path", "pattern", "script", "macro", "cibles", "fichiers"):
        if key in tool_input:
            return str(tool_input[key])[:120]
    return ""
