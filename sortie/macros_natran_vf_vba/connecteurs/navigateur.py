"""Connecteur navigateur : ouverture d'un fichier local (équivalent de ThisWorkbook.FollowHyperlink)."""

from __future__ import annotations

import webbrowser
from pathlib import Path


def ouvrir_fichier(chemin) -> None:
    """Équivalent de `OuvrirFichier` (VBA) : ouvre le fichier avec l'application associée.

    Lève OSError si aucun navigateur n'a pu être lancé.
    """
    adresse = Path(chemin).resolve().as_uri()
    if not webbrowser.open(adresse):
        raise OSError(f"Impossible d'ouvrir {adresse}")
