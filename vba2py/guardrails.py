"""Garde-fous de l'agent : chaque action est vérifiée par du code avant d'être exécutée.

Les instructions données au modèle ne suffisent pas : même si le modèle (ou un commentaire
malveillant caché dans le VBA) voulait sortir du cadre, ces règles l'en empêchent.

Règles :
- lecture : uniquement dans le dossier du projet ;
- écriture : uniquement dans le dossier du projet, jamais dans `_source_vba/`, et
  seulement des fichiers texte (.py, .md, .txt, .json, .csv, .toml, .ini, .cfg) ;
- code Python écrit : pas de réseau, de sous-processus, d'eval/exec, ni de COM Windows ;
- aucun terminal : l'agent n'a que trois outils d'exécution fournis par vba2py (lancer les tests,
  vérifier la syntaxe, exécuter un script .py du projet), dont les arguments sont contrôlés ici ;
- tout autre outil (terminal, web, notebooks, sous-agents...) est refusé.

Limite assumée : pytest exécute le code écrit par l'agent. Le filtrage du contenu réduit
le risque mais ne remplace pas une isolation complète (conteneur) pour des classeurs
d'origine inconnue.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from .bibliotheque import LIBRARY_DIRNAME
from .extract import SOURCE_DIRNAME

ALLOWED_WRITE_EXTENSIONS = {".py", ".md", ".txt", ".json", ".csv", ".toml", ".ini", ".cfg"}
HARMLESS_TOOLS = {"TodoWrite"}

_FORBIDDEN_MODULES = (
    "subprocess|socket|requests|urllib|urllib3|http|httpx|aiohttp|ftplib|smtplib|telnetlib|"
    "paramiko|ctypes|importlib|win32com|pythoncom|pywintypes|xlwings|pywinauto|multiprocessing"
)
_FORBIDDEN_CODE = [
    (re.compile(rf"^\s*(import|from)\s+({_FORBIDDEN_MODULES})\b", re.MULTILINE), "module interdit"),
    (re.compile(r"\bos\.(system|popen|exec\w*|spawn\w*|fork\w*|kill)\s*\("), "appel système interdit"),
    (re.compile(r"\bshutil\.rmtree\s*\("), "suppression récursive interdite"),
    (re.compile(r"(?<![\w.])(eval|exec|compile|__import__)\s*\("), "exécution dynamique de code interdite"),
]

# Outils d'exécution fournis par vba2py (serveur MCP interne « vba2py », voir outils.py)
OUTIL_TESTS = "mcp__vba2py__lancer_tests"
OUTIL_SYNTAXE = "mcp__vba2py__verifier_syntaxe"
OUTIL_SCRIPT = "mcp__vba2py__executer_script"

# Modules autorisés uniquement dans connecteurs/ (intégrations Office fidèles au VBA)
CONNECTOR_DIRNAME = "connecteurs"
_CONNECTOR_MODULES = {"win32com", "pythoncom", "pywintypes"}


class Guardrails:
    def __init__(self, project_dir: Path, journal_path: Path | None = None):
        self.project_dir = project_dir.resolve()
        self.source_dir = self.project_dir / SOURCE_DIRNAME
        self.library_dir = self.project_dir / LIBRARY_DIRNAME
        self.journal_path = journal_path

    # ------------------------------------------------------------------ décision

    def check(self, tool_name: str, tool_input: dict[str, Any]) -> tuple[bool, str]:
        """Retourne (autorisé, raison)."""
        if tool_name in HARMLESS_TOOLS:
            return True, "outil sans effet sur les fichiers"
        handler = {
            "Read": self._check_read,
            "Glob": self._check_glob,
            "Grep": self._check_grep,
            "Write": self._check_write,
            "Edit": self._check_edit,
            "MultiEdit": self._check_edit,
            OUTIL_TESTS: self._check_tests,
            OUTIL_SYNTAXE: self._check_syntaxe,
            OUTIL_SCRIPT: self._check_script,
        }.get(tool_name)
        if handler is None:
            return False, f"outil « {tool_name} » non autorisé pour cet agent"
        return handler(tool_input)

    def _check_read(self, tool_input: dict) -> tuple[bool, str]:
        path = tool_input.get("file_path", "")
        if not self._inside(path, self.project_dir):
            return False, f"lecture hors du dossier du projet refusée : {path}"
        return True, "lecture dans le projet"

    def _check_glob(self, tool_input: dict) -> tuple[bool, str]:
        return self._check_search(tool_input, [tool_input.get("pattern") or ""])

    def _check_grep(self, tool_input: dict) -> tuple[bool, str]:
        # Pour Grep, « pattern » est une expression régulière, pas un chemin.
        return self._check_search(tool_input, [tool_input.get("glob") or ""])

    def _check_search(self, tool_input: dict, path_patterns: list[str]) -> tuple[bool, str]:
        path = tool_input.get("path") or str(self.project_dir)
        if not self._inside(path, self.project_dir):
            return False, f"recherche hors du dossier du projet refusée : {path}"
        for value in path_patterns:
            if ".." in Path(value).parts or value.startswith(("/", "~")):
                return False, f"motif de chemin hors du projet refusé : {value}"
        return True, "recherche dans le projet"

    def _check_write_target(self, path: str) -> tuple[bool, str]:
        if not self._inside(path, self.project_dir):
            return False, f"écriture hors du dossier du projet refusée : {path}"
        if self._inside(path, self.source_dir):
            return False, f"le code VBA source est en lecture seule : {path}"
        if self._inside(path, self.library_dir):
            return False, f"la bibliothèque fournie par vba2py est en lecture seule : {path}"
        if Path(path).suffix.lower() not in ALLOWED_WRITE_EXTENSIONS:
            return False, f"type de fichier non autorisé en écriture : {path}"
        return True, "écriture dans le projet"

    def _check_write(self, tool_input: dict) -> tuple[bool, str]:
        path = tool_input.get("file_path", "")
        ok, reason = self._check_write_target(path)
        if not ok:
            return ok, reason
        return self._check_python_content(path, tool_input.get("content", ""))

    def _check_edit(self, tool_input: dict) -> tuple[bool, str]:
        path = tool_input.get("file_path", "")
        ok, reason = self._check_write_target(path)
        if not ok:
            return ok, reason
        new_strings = [tool_input.get("new_string", "")]
        new_strings += [e.get("new_string", "") for e in tool_input.get("edits", [])]
        return self._check_python_content(path, "\n".join(new_strings))

    def _check_python_content(self, path: str, content: str) -> tuple[bool, str]:
        if Path(path).suffix.lower() != ".py":
            return True, "fichier non Python"
        in_connector = self._inside(path, self.project_dir / CONNECTOR_DIRNAME)
        for pattern, label in _FORBIDDEN_CODE:
            for match in pattern.finditer(content):
                if in_connector and label == "module interdit" and match.group(2) in _CONNECTOR_MODULES:
                    continue
                return False, f"{label} dans {Path(path).name} : « {match.group(0).strip()} »"
        return True, "contenu Python conforme"

    def _check_tests(self, tool_input: dict) -> tuple[bool, str]:
        cibles = tool_input.get("cibles") or []
        for cible in cibles:
            if str(cible).startswith("-"):
                return False, f"cible de test invalide (les options ne sont pas acceptées) : {cible}"
        return self._check_paths([str(c).split("::", 1)[0] for c in cibles], "tests")

    def _check_syntaxe(self, tool_input: dict) -> tuple[bool, str]:
        fichiers = tool_input.get("fichiers") or []
        if not fichiers:
            return False, "aucun fichier à vérifier"
        if any(not str(f).endswith(".py") for f in fichiers):
            return False, "seuls des fichiers .py peuvent être vérifiés"
        return self._check_paths([str(f) for f in fichiers], "vérification de syntaxe")

    def _check_script(self, tool_input: dict) -> tuple[bool, str]:
        script = str(tool_input.get("script") or "")
        if not script.endswith(".py"):
            return False, f"seul un script .py du projet peut être exécuté : {script}"
        ok, raison = self._check_paths([script], "script")
        if not ok:
            return ok, raison
        if self._inside(script, self.source_dir) or self._inside(script, self.library_dir):
            return False, "exécution d'un fichier de _source_vba/ ou bibliotheque/ interdite"
        for argument in tool_input.get("arguments") or []:
            argument = str(argument)
            if (Path(argument).is_absolute() or ".." in Path(argument).parts) and not self._inside(argument, self.project_dir):
                return False, f"argument pointant hors du projet refusé : {argument}"
        return True, "script du projet"

    def _check_paths(self, paths: list[str], label: str) -> tuple[bool, str]:
        for p in paths:
            if not self._inside(p, self.project_dir):
                return False, f"{label} : chemin hors du projet refusé : {p}"
        return True, f"{label} dans le projet"

    def _inside(self, path: str, root: Path) -> bool:
        if not path:
            return False
        candidate = Path(path).expanduser()
        if not candidate.is_absolute():
            candidate = self.project_dir / candidate
        try:
            candidate.resolve().relative_to(root.resolve())
            return True
        except ValueError:
            return False

    # ------------------------------------------------------------------ intégration Agent SDK

    async def pre_tool_use_hook(self, hook_input: dict, tool_use_id: str | None, context: Any) -> dict:
        tool_name = hook_input.get("tool_name", "")
        tool_input = hook_input.get("tool_input", {}) or {}
        allowed, reason = self.check(tool_name, tool_input)
        self._log(tool_name, tool_input, allowed, reason)
        if allowed:
            return {}
        return {
            "hookSpecificOutput": {
                "hookEventName": "PreToolUse",
                "permissionDecision": "deny",
                "permissionDecisionReason": f"Refusé par les garde-fous : {reason}",
            }
        }

    def _log(self, tool_name: str, tool_input: dict, allowed: bool, reason: str) -> None:
        if self.journal_path is None:
            return
        summary = {k: (v[:200] + "…" if isinstance(v, str) and len(v) > 200 else v) for k, v in tool_input.items()}
        entry = {
            "date": datetime.now().isoformat(timespec="seconds"),
            "outil": tool_name,
            "decision": "autorise" if allowed else "refuse",
            "raison": reason,
            "entree": summary,
        }
        self.journal_path.parent.mkdir(parents=True, exist_ok=True)
        with self.journal_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
