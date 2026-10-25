"""Compare l'extraction oletools d'un classeur avec un export manuel depuis l'éditeur VBA.

Sert à vérifier que l'extraction automatique ne perd rien sur vos propres classeurs.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass
from pathlib import Path

from .extract import extract_from_export_dir, extract_from_workbook, strip_vba_headers


@dataclass
class ModuleComparison:
    name: str
    status: str  # "identique", "different", "absent_export", "absent_classeur"
    diff: str = ""


def compare(export_dir: Path, workbook: Path) -> list[ModuleComparison]:
    exported = {m.name.lower(): m for m in extract_from_export_dir(export_dir) if not m.is_empty}
    extracted_modules, _ = extract_from_workbook(workbook)
    extracted = {m.name.lower(): m for m in extracted_modules if not m.is_empty}

    results = []
    for key in sorted(exported.keys() | extracted.keys()):
        exp, ext = exported.get(key), extracted.get(key)
        if exp is None:
            results.append(ModuleComparison(ext.name, "absent_export"))
        elif ext is None:
            results.append(ModuleComparison(exp.name, "absent_classeur"))
        else:
            a = strip_vba_headers(exp.code).splitlines()
            b = strip_vba_headers(ext.code).splitlines()
            if a == b:
                results.append(ModuleComparison(exp.name, "identique"))
            else:
                diff = "\n".join(
                    difflib.unified_diff(a, b, "export", "oletools", lineterm="")
                )
                results.append(ModuleComparison(exp.name, "different", diff))
    return results


def format_report(results: list[ModuleComparison]) -> str:
    labels = {
        "identique": "✅ identique",
        "different": "❌ différent",
        "absent_export": "⚠️  absent de l'export (présent dans le classeur)",
        "absent_classeur": "⚠️  absent du classeur (présent dans l'export)",
    }
    lines = [f"{r.name:<30} {labels[r.status]}" for r in results]
    for r in results:
        if r.diff:
            lines += ["", f"--- Différences pour {r.name} ---", r.diff]
    ok = sum(r.status == "identique" for r in results)
    lines += ["", f"{ok}/{len(results)} module(s) identique(s)."]
    return "\n".join(lines)
