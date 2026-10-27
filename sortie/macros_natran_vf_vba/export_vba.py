"""Traduction du module VBA `ExportVBA` : export des modules VBA d'un classeur (via Excel)."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path


def dossier_export(classeur) -> Path:
    """Dossier `<NomDuClasseur>_vba` créé à côté du classeur (`wb.Path` + `Left(wb.Name, InStrRev(...)-1)`)."""
    classeur = Path(classeur).resolve()
    return classeur.parent / (classeur.stem + "_vba")


def exporter_modules_vba(classeur, exporter=None) -> tuple[Path, int]:
    """VBA `ExporterModulesVBA`. `classeur` remplace ActiveWorkbook ; `exporter(chemin, dossier)` est
    injectable (par défaut `connecteurs.excel.exporter_modules_vba`). Retourne (dossier, nombre de modules).
    Prérequis Excel : « Accès approuvé au modèle d'objet du projet VBA »."""
    classeur = Path(classeur)
    if not classeur.exists():
        raise FileNotFoundError(str(classeur))  # équivalent de « Enregistrez d'abord le classeur »
    dossier = dossier_export(classeur)
    dossier.mkdir(exist_ok=True)
    if exporter is None:
        from connecteurs import excel as connecteur_excel
        exporter = connecteur_excel.exporter_modules_vba
    try:
        nombre = exporter(str(classeur.resolve()), str(dossier))
    except Exception as erreur:  # le message d'ErreurExcel est conservé en entier
        logging.error("Export impossible : %s\nVerifiez que « Acces approuve au modele d'objet du projet VBA » "
                      "est active.", erreur)
        raise
    logging.info("%s module(s) exporte(s) dans :\n%s", nombre, dossier)
    return dossier, nombre


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description="ExportVBA : ExporterModulesVBA")
    parser.add_argument("classeur", help="classeur dont les modules sont exportés (remplace ActiveWorkbook)")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    exporter_modules_vba(args.classeur)


if __name__ == "__main__":
    main()
