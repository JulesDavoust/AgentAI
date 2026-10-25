"""Contrôle et finalisation du livrable produit par l'agent.

- `interface.json` : contrat de la ligne de commande du script traduit. Il est conservé
  d'une traduction à l'autre pour que les commandes restent identiques.
- `RAPPORT.md` : structure imposée par prompts/rapport_modele.md. Les sections factuelles
  (commandes, tests, extraction) sont écrites ici par le code, pas par l'agent.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from .extract import SOURCE_DIRNAME

PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"
INTERFACE_FILENAME = "interface.json"
PREVIOUS_INTERFACE_FILENAME = "interface_precedente.json"
REPORT_FILENAME = "RAPPORT.md"


@dataclass
class Verification:
    erreurs: list[str] = field(default_factory=list)
    resultat_tests: str = ""

    @property
    def ok(self) -> bool:
        return not self.erreurs


# --------------------------------------------------------------------------- interface


def charger_interface(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    for cle in ("script", "fixture_principale", "arguments"):
        if cle not in data:
            raise ValueError(f"{path.name} : clé « {cle} » manquante")
    for arg in data["arguments"]:
        if "nom" not in arg or "remplace" not in arg:
            raise ValueError(f"{path.name} : chaque argument doit avoir « nom » et « remplace »")
    return data


def noms_arguments(interface: dict) -> list[str]:
    return [a["nom"] for a in interface["arguments"]]


def options_du_script(project_dir: Path, script: str) -> set[str] | None:
    """Options déclarées par argparse, lues dans `python script --help` (None si la commande échoue)."""
    result = subprocess.run(
        [sys.executable, script, "--help"], cwd=project_dir, capture_output=True, text=True, timeout=60
    )
    if result.returncode != 0:
        return None
    return set(re.findall(r"(?<![\w-])(--[a-z0-9][a-z0-9-]*)", result.stdout)) - {"--help"}


def verifier_interface(project_dir: Path, interface: dict, precedente: dict | None) -> list[str]:
    erreurs = []
    script = interface["script"]
    if not (project_dir / script).exists():
        return [f"interface.json : script introuvable ({script})"]
    if not (project_dir / interface["fixture_principale"]).exists():
        erreurs.append(f"interface.json : fixture principale introuvable ({interface['fixture_principale']})")

    noms = noms_arguments(interface)
    if not noms or noms[0] != "classeur":
        erreurs.append("interface.json : le premier argument doit être « classeur »")
    declarees = {n for n in noms if n.startswith("--")}
    options = options_du_script(project_dir, script)
    if options is None:
        erreurs.append(f"`python {script} --help` a échoué")
    else:
        if declarees - options:
            erreurs.append(f"options de interface.json absentes du script : {sorted(declarees - options)}")
        if options - declarees:
            erreurs.append(f"options du script absentes de interface.json : {sorted(options - declarees)}")

    if precedente is not None and noms_arguments(precedente) != noms:
        erreurs.append(
            "les arguments ont changé par rapport à la traduction précédente : "
            f"{noms_arguments(precedente)} → {noms}"
        )
    return erreurs


def commandes(interface: dict) -> str:
    """Section « Commandes » du rapport, générée de façon déterministe depuis interface.json."""
    script = interface["script"]
    noms = noms_arguments(interface)
    securite = ""
    if "--mode-mail" in noms:
        securite = " --mode-mail afficher"
    test_mail = " --mode-mail aucun" if "--mode-mail" in noms else ""
    dossier = " --dossier-sortie sortie_test" if "--dossier-sortie" in noms else ""
    dossier_reel = ' --dossier-sortie "C:\\chemin\\vers\\sorties"' if "--dossier-sortie" in noms else ""

    lignes = [
        "```powershell",
        "pip install -r requirements.txt",
        "",
        "# Tests sur données fictives",
        "python tests/generer_donnees_fictives.py",
        "python -m pytest -q",
        "",
        "# Liste complète des options",
        f"python {script} --help",
        "",
        "# Essai sur le classeur fictif (aucune action externe)",
        f"python {script} {interface['fixture_principale']}{dossier}{test_mail}",
        "",
        "# Sur un vrai classeur (chemins complets ; mail affiché sans être envoyé)",
        f'python {script} "C:\\chemin\\vers\\classeur.xlsm"{dossier_reel}{securite}',
        "```",
        "",
        "| Argument | Remplace dans le VBA | Valeur par défaut |",
        "|---|---|---|",
    ]
    for arg in interface["arguments"]:
        defaut = arg.get("defaut")
        defaut = "—" if defaut in (None, "") else f"`{defaut}`"
        lignes.append(f"| `{arg['nom']}` | {arg['remplace']} | {defaut} |")
    return "\n".join(lignes)


# --------------------------------------------------------------------------- rapport


def titres_attendus() -> list[str]:
    modele = (PROMPTS_DIR / "rapport_modele.md").read_text(encoding="utf-8")
    return re.findall(r"^## .+$", modele, re.MULTILINE)


def verifier_structure_rapport(rapport: str) -> list[str]:
    titres = re.findall(r"^## .+$", rapport, re.MULTILINE)
    attendus = titres_attendus()
    if titres != attendus:
        manquants = [t for t in attendus if t not in titres]
        en_trop = [t for t in titres if t not in attendus]
        detail = []
        if manquants:
            detail.append(f"manquants : {manquants}")
        if en_trop:
            detail.append(f"non prévus : {en_trop}")
        if not detail:
            detail.append("ordre différent du modèle")
        return ["RAPPORT.md ne suit pas le modèle (" + " ; ".join(detail) + ")"]
    return []


def remplir_section(rapport: str, nom: str, contenu: str) -> str:
    debut, fin = f"<!-- vba2py:debut {nom} -->", f"<!-- vba2py:fin {nom} -->"
    pattern = re.compile(re.escape(debut) + r".*?" + re.escape(fin), re.DOTALL)
    bloc = f"{debut}\n{contenu.strip()}\n{fin}"
    if pattern.search(rapport):
        return pattern.sub(lambda _: bloc, rapport)
    return rapport


def section_extraction(project_dir: Path) -> str:
    extraction = project_dir / SOURCE_DIRNAME / "EXTRACTION.md"
    if not extraction.exists():
        return "Rapport d'extraction absent."
    texte = extraction.read_text(encoding="utf-8")
    return re.sub(r"^# .*\n", "", texte, count=1).strip()


# --------------------------------------------------------------------------- vérification complète


def verifier_projet(project_dir: Path, interface_precedente: Path | None = None) -> Verification:
    verif = Verification()
    manquants = [f for f in (REPORT_FILENAME, "requirements.txt", "tests", INTERFACE_FILENAME) if not (project_dir / f).exists()]
    if manquants:
        verif.erreurs.append(f"éléments manquants : {', '.join(manquants)}")

    interface = None
    if (project_dir / INTERFACE_FILENAME).exists():
        try:
            interface = charger_interface(project_dir / INTERFACE_FILENAME)
            precedente = charger_interface(interface_precedente) if interface_precedente and interface_precedente.exists() else None
            verif.erreurs += verifier_interface(project_dir, interface, precedente)
        except (ValueError, json.JSONDecodeError) as exc:
            verif.erreurs.append(f"interface.json invalide : {exc}")

    if (project_dir / "tests").exists():
        result = subprocess.run(
            [sys.executable, "-m", "pytest", "-q", "tests"], cwd=project_dir, capture_output=True, text=True
        )
        derniere = (result.stdout.strip().splitlines() or [""])[-1]
        verif.resultat_tests = derniere
        if result.returncode != 0:
            verif.erreurs.append(f"tests en échec : {derniere}\n{result.stdout[-3000:]}")

    rapport_path = project_dir / REPORT_FILENAME
    if rapport_path.exists():
        rapport = rapport_path.read_text(encoding="utf-8")
        verif.erreurs += verifier_structure_rapport(rapport)
        statut = "✅" if not any(e.startswith("tests en échec") for e in verif.erreurs) else "❌"
        rapport = remplir_section(
            rapport,
            "tests",
            f"Vérification indépendante par vba2py : `python -m pytest -q tests` → {statut} **{verif.resultat_tests}**",
        )
        rapport = remplir_section(rapport, "extraction", section_extraction(project_dir))
        if interface is not None:
            rapport = remplir_section(rapport, "commandes", commandes(interface))
        rapport_path.write_text(rapport, encoding="utf-8")
    return verif
