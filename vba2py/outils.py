"""Outils d'exécution donnés à l'agent, à la place d'un terminal.

L'agent ne tape aucune commande : il demande à vba2py de lancer les tests, de vérifier la syntaxe
ou d'exécuter un script du projet, et vba2py le fait lui-même avec son propre Python, sans shell.
Cela fonctionne de la même façon sous Windows (sans Git Bash), macOS et Linux.

Les arguments sont vérifiés deux fois : par le hook des garde-fous, puis ici avant l'exécution.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Annotated

from claude_agent_sdk import create_sdk_mcp_server, tool

from .guardrails import OUTIL_SCRIPT, OUTIL_SYNTAXE, OUTIL_TESTS, Guardrails

NOM_SERVEUR = "vba2py"
NOMS_OUTILS = [OUTIL_TESTS, OUTIL_SYNTAXE, OUTIL_SCRIPT]
DELAI_TESTS = 900  # secondes : les tests sur Excel réel peuvent être lents
DELAI_SCRIPT = 300
TAILLE_SORTIE = 20_000  # caractères conservés (la fin, où pytest résume)


def environnement() -> dict[str, str]:
    env = dict(os.environ)
    env.update({"MPLBACKEND": "Agg", "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"})
    return env


def _executer(commande: list[str], dossier: Path, delai: int) -> tuple[int, str]:
    try:
        resultat = subprocess.run(
            commande, cwd=dossier, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=delai, env=environnement(),
        )
    except subprocess.TimeoutExpired:
        return -1, f"Arrêté : durée maximale dépassée ({delai} s)."
    sortie = (resultat.stdout or "") + (("\n" + resultat.stderr) if resultat.stderr else "")
    if len(sortie) > TAILLE_SORTIE:
        sortie = "[…début tronqué…]\n" + sortie[-TAILLE_SORTIE:]
    return resultat.returncode, sortie.strip()


def lancer_tests(dossier: Path, cibles: list[str] | None = None, marqueur: str | None = None,
                 mot_cle: str | None = None, arreter_au_premier_echec: bool = False) -> tuple[int, str]:
    commande = [sys.executable, "-m", "pytest", "-q", "-rs"]
    if arreter_au_premier_echec:
        commande.append("-x")
    if marqueur:
        commande += ["-m", marqueur]
    if mot_cle:
        commande += ["-k", mot_cle]
    commande += list(cibles or [])
    return _executer(commande, dossier, DELAI_TESTS)


def verifier_syntaxe(dossier: Path, fichiers: list[str]) -> tuple[int, str]:
    erreurs = []
    for fichier in fichiers:
        chemin = dossier / fichier
        if not chemin.exists():
            erreurs.append(f"{fichier} : fichier introuvable")
            continue
        try:  # même contrôle que py_compile, sans écrire de .pyc
            compile(chemin.read_bytes(), fichier, "exec", dont_inherit=True)
        except SyntaxError as exc:
            erreurs.append(f"{fichier}, ligne {exc.lineno} : {exc.msg}\n  {(exc.text or '').rstrip()}")
    if erreurs:
        return 1, "\n\n".join(erreurs)
    return 0, f"Syntaxe correcte : {len(fichiers)} fichier(s)."


def executer_script(dossier: Path, script: str, arguments: list[str] | None = None) -> tuple[int, str]:
    return _executer([sys.executable, script, *[str(a) for a in arguments or []]], dossier, DELAI_SCRIPT)


def _reponse(code: int, sortie: str) -> dict:
    texte = f"Code de retour : {code}\n\n{sortie or '(aucune sortie)'}"
    return {"content": [{"type": "text", "text": texte}], "is_error": code != 0}


def _refus(raison: str) -> dict:
    return {"content": [{"type": "text", "text": f"Refusé par les garde-fous : {raison}"}], "is_error": True}


def creer_serveur(project_dir: Path, guardrails: Guardrails):
    """Serveur MCP interne exposant les trois outils, liés au dossier du projet."""
    return create_sdk_mcp_server(name=NOM_SERVEUR, tools=creer_outils(project_dir, guardrails))


def creer_outils(project_dir: Path, guardrails: Guardrails) -> list:

    def controler(nom_outil: str, args: dict) -> str | None:
        ok, raison = guardrails.check(nom_outil, args)
        return None if ok else raison

    @tool(
        "lancer_tests",
        "Lance pytest dans le dossier du projet (équivalent de `python -m pytest -q`). "
        "Sans argument : toute la batterie. Renvoie la sortie de pytest et son code de retour.",
        {
            "type": "object",
            "properties": {
                "cibles": {"type": "array", "items": {"type": "string"},
                           "description": "Fichiers ou tests précis, ex. tests/test_x.py::test_y (facultatif)"},
                "marqueur": {"type": "string", "description": "Expression -m, ex. excel_reel (facultatif)"},
                "mot_cle": {"type": "string", "description": "Expression -k (facultatif)"},
                "arreter_au_premier_echec": {"type": "boolean", "description": "Équivalent de -x"},
            },
        },
    )
    async def outil_tests(args):
        refus = controler(OUTIL_TESTS, args)
        if refus:
            return _refus(refus)
        return _reponse(*lancer_tests(project_dir, args.get("cibles"), args.get("marqueur"),
                                      args.get("mot_cle"), bool(args.get("arreter_au_premier_echec"))))

    @tool(
        "verifier_syntaxe",
        "Vérifie la syntaxe de fichiers Python du projet (équivalent de `python -m py_compile`).",
        {"fichiers": Annotated[list[str], "Chemins des fichiers .py, relatifs au projet"]},
    )
    async def outil_syntaxe(args):
        refus = controler(OUTIL_SYNTAXE, args)
        if refus:
            return _refus(refus)
        return _reponse(*verifier_syntaxe(project_dir, args["fichiers"]))

    @tool(
        "executer_script",
        "Exécute un script Python du projet (équivalent de `python <script> <arguments>`), "
        "par exemple tests/generer_donnees_fictives.py.",
        {
            "type": "object",
            "properties": {
                "script": {"type": "string", "description": "Chemin du script .py, relatif au projet"},
                "arguments": {"type": "array", "items": {"type": "string"}, "description": "Arguments (facultatif)"},
            },
            "required": ["script"],
        },
    )
    async def outil_script(args):
        refus = controler(OUTIL_SCRIPT, args)
        if refus:
            return _refus(refus)
        return _reponse(*executer_script(project_dir, args["script"], args.get("arguments")))

    return [outil_tests, outil_syntaxe, outil_script]
