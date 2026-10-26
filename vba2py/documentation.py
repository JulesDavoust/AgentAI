"""Documentation facultative jointe aux macros (spécifications, règles métier, exemples).

Deux rattachements possibles :
- `--doc <fichier ou dossier>` (répétable) : documentation **commune**, donnée à tous les projets
  de la commande ;
- un dossier `<NomDuClasseur>_doc/` placé à côté du classeur (ou `<dossier_exporté>_doc/` à côté
  d'un dossier de modules exportés) : documentation **rattachée** à ce seul projet.

Les documents sont copiés dans `_source_vba/documentation/` (lecture seule pour l'agent), avec un
index. Le code VBA prime toujours sur la documentation.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from .extract import SOURCE_DIRNAME, ExtractionError

DOC_EXTENSIONS = {".pdf", ".md", ".txt"}
DOC_SUFFIX = "_doc"
DOC_DIRNAME = "documentation"


@dataclass
class Document:
    chemin: Path
    portee: str  # "commune" ou "rattachée"
    nom: str = ""  # nom de la copie dans _source_vba/documentation/


def collecter(chemins: list[Path]) -> list[Path]:
    """Fichiers de documentation désignés par des fichiers ou dossiers (parcourus récursivement)."""
    fichiers: list[Path] = []
    refuses: list[str] = []
    for chemin in chemins:
        chemin = Path(chemin)
        if not chemin.exists():
            raise ExtractionError(f"documentation introuvable : {chemin}")
        candidats = [chemin] if chemin.is_file() else sorted(p for p in chemin.rglob("*") if p.is_file())
        for f in candidats:
            if f.name.startswith("."):
                continue
            if f.suffix.lower() in DOC_EXTENSIONS:
                fichiers.append(f.resolve())
            else:
                refuses.append(str(f))
    if refuses:
        raise ExtractionError(
            "format de documentation non pris en charge (acceptés : "
            f"{', '.join(sorted(DOC_EXTENSIONS))}) : {', '.join(refuses)}"
        )
    return list(dict.fromkeys(fichiers))  # sans doublons, ordre conservé


def dossier_rattache(source: Path) -> Path:
    nom = source.stem if source.is_file() else source.name
    return source.parent / f"{nom}{DOC_SUFFIX}"


def documentation_du_projet(source: Path, communs: list[Path]) -> list[Document]:
    documents = [Document(p, "commune") for p in communs]
    dossier = dossier_rattache(source)
    if dossier.is_dir():
        deja = {d.chemin for d in documents}
        documents += [Document(p, "rattachée") for p in collecter([dossier]) if p not in deja]
    return documents


def ecrire_documentation(project_dir: Path, documents: list[Document]) -> None:
    """Copie les documents dans _source_vba/documentation/ et écrit INDEX.md."""
    cible = project_dir / SOURCE_DIRNAME / DOC_DIRNAME
    if cible.exists():
        shutil.rmtree(cible)  # dossier généré par vba2py : reflète la commande en cours
    if not documents:
        return
    cible.mkdir(parents=True)
    pris: set[str] = set()
    for doc in documents:
        nom, n = doc.chemin.name, 2
        while nom.lower() in pris:  # deux documents de même nom venant de dossiers différents
            nom = f"{doc.chemin.stem}_{n}{doc.chemin.suffix}"
            n += 1
        pris.add(nom.lower())
        doc.nom = nom
        shutil.copyfile(doc.chemin, cible / nom)

    lignes = [
        "# Documentation fournie",
        "",
        "Le code VBA prime toujours : en cas de contradiction, traduire le code et signaler l'écart.",
        "",
        "| Document | Portée | Taille |",
        "|---|---|---|",
    ]
    for doc in documents:
        lignes.append(f"| {doc.nom} | {doc.portee} | {_taille(doc.chemin)} |")
    (cible / "INDEX.md").write_text("\n".join(lignes) + "\n", encoding="utf-8")


def _taille(chemin: Path) -> str:
    octets = chemin.stat().st_size
    return f"{octets / 1024:.0f} Ko" if octets < 1024 * 1024 else f"{octets / 1024 / 1024:.1f} Mo"
