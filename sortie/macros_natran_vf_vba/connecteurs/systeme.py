"""Ouverture de fichiers (VBA ``ThisWorkbook.FollowHyperlink``)."""

from __future__ import annotations

import os
from pathlib import Path


def ouvrir_fichier(chemin: str) -> None:
    """VBA ``ThisWorkbook.FollowHyperlink chemin`` : ouvre le fichier avec l'application associée."""
    os.startfile(str(Path(chemin).resolve()))  # Windows uniquement
