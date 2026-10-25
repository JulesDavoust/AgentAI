"""vba2py : traduit des macros VBA Excel en Python, avec une batterie de tests.

Exemples :
    python main.py extraire examples/Ventes.xlsm -o sortie/
    python main.py traduire examples/Ventes.xlsm -o sortie/
    python main.py comparer examples/Ventes_vba/ examples/Ventes.xlsm
    python main.py verifier sortie/ventes
"""

from __future__ import annotations

import argparse
import asyncio
import subprocess
import sys
import warnings
from pathlib import Path

from vba2py.compare import compare, format_report
from vba2py.extract import ExtractionError, discover_projects, write_project_sources

warnings.filterwarnings("ignore", category=DeprecationWarning, module="oletools")


def cmd_extraire(args) -> list[Path]:
    projects = discover_projects(Path(args.entree))
    out_root = Path(args.sortie).resolve()
    project_dirs = []
    for project in projects:
        project_dir = write_project_sources(project, out_root)
        project_dirs.append(project_dir)
        kept = [m.name for m in project.modules if not m.is_empty]
        print(f"📄 {project.source.name} → {project_dir}")
        print(f"   {len(kept)} module(s) avec du code : {', '.join(kept) or 'aucun'}")
        for warning in project.warnings:
            print(f"   ⚠️  {warning}")
    return project_dirs


def cmd_traduire(args) -> int:
    from vba2py.agent import translate_project

    project_dirs = cmd_extraire(args)
    journals = Path(args.sortie).resolve() / "_journaux"
    status = 0
    for project_dir in project_dirs:
        print(f"\n🤖 Traduction de {project_dir.name}…")
        run = asyncio.run(
            translate_project(
                project_dir,
                journals / f"{project_dir.name}.jsonl",
                model=args.modele,
                max_turns=args.max_tours,
                max_budget_usd=args.budget,
            )
        )
        cost = f"{run.cost_usd:.2f} $" if run.cost_usd is not None else "inconnu"
        print(f"\n   Agent : {'terminé' if run.success else 'ÉCHEC'} ({run.turns} tours, coût {cost})")
        if run.summary:
            print("   " + run.summary.strip().replace("\n", "\n   "))
        # On ne se fie pas à ce que dit l'agent : les tests sont relancés ici.
        if verifier_projet(project_dir) != 0 or not run.success:
            status = 1
    print(f"\nJournal des actions de l'agent : {journals}")
    return status


def verifier_projet(project_dir: Path) -> int:
    print(f"\n🔎 Vérification indépendante de {project_dir.name}")
    missing = [f for f in ("RAPPORT.md", "requirements.txt", "tests") if not (project_dir / f).exists()]
    if missing:
        print(f"   ❌ Éléments manquants : {', '.join(missing)}")
        return 1
    result = subprocess.run(
        [sys.executable, "-m", "pytest", "-q", "tests"],
        cwd=project_dir,
        capture_output=True,
        text=True,
    )
    last_lines = (result.stdout.strip().splitlines() or [""])[-1]
    print(f"   {'✅' if result.returncode == 0 else '❌'} pytest : {last_lines}")
    if result.returncode != 0:
        print(result.stdout[-3000:])
    return result.returncode


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="vba2py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="commande", required=True)

    p = sub.add_parser("extraire", help="extraire le VBA et la structure (sans appel à l'IA)")
    p.add_argument("entree", help="classeur .xlsm/.xlsb/.xls ou dossier (classeurs ou modules exportés)")
    p.add_argument("-o", "--sortie", default="sortie")

    p = sub.add_parser("traduire", help="extraire puis traduire avec l'agent")
    p.add_argument("entree")
    p.add_argument("-o", "--sortie", default="sortie")
    p.add_argument("--modele", default="claude-opus-5-5", help="ex. claude-opus-5-5, claude-sonnet-5-5")
    p.add_argument("--max-tours", type=int, default=60)
    p.add_argument("--budget", type=float, default=5.0, help="budget maximum en dollars par projet")

    p = sub.add_parser("comparer", help="comparer un export manuel avec l'extraction automatique")
    p.add_argument("export", help="dossier des modules exportés depuis l'éditeur VBA")
    p.add_argument("classeur", help="le classeur .xlsm d'origine")

    p = sub.add_parser("verifier", help="relancer les tests d'un projet traduit")
    p.add_argument("projet")

    args = parser.parse_args(argv)
    try:
        if args.commande == "extraire":
            cmd_extraire(args)
            return 0
        if args.commande == "traduire":
            return cmd_traduire(args)
        if args.commande == "comparer":
            results = compare(Path(args.export), Path(args.classeur))
            print(format_report(results))
            return 0 if all(r.status == "identique" for r in results) else 1
        if args.commande == "verifier":
            return verifier_projet(Path(args.projet).resolve())
    except ExtractionError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
