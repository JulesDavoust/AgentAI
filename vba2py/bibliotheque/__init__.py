"""Modules fournis par vba2py et copiés tels quels dans chaque projet traduit.

L'agent les utilise mais ne peut pas les modifier : ils sont testés une fois pour toutes ici.
"""

from pathlib import Path

LIBRARY_DIRNAME = "bibliotheque"
LIBRARY_SOURCE = Path(__file__).resolve().parent


def fichiers_bibliotheque() -> list[Path]:
    return sorted(p for p in LIBRARY_SOURCE.glob("*.py") if p.name != "__init__.py")


def installer_bibliotheque(project_dir: Path) -> Path:
    cible = project_dir / LIBRARY_DIRNAME
    cible.mkdir(parents=True, exist_ok=True)
    (cible / "__init__.py").write_text('"""Fourni par vba2py : ne pas modifier."""\n', encoding="utf-8")
    for fichier in fichiers_bibliotheque():
        (cible / fichier.name).write_bytes(fichier.read_bytes())
    return cible


def fichiers_modifies(project_dir: Path) -> list[str]:
    """Fichiers de la bibliothèque absents ou différents de la version fournie par vba2py."""
    cible = project_dir / LIBRARY_DIRNAME
    if not cible.exists():
        return []
    return [
        f.name for f in fichiers_bibliotheque()
        if not (cible / f.name).exists() or (cible / f.name).read_bytes() != f.read_bytes()
    ]
