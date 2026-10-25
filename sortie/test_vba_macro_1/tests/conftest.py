import sys
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RACINE))

from tests.generer_donnees_fictives import DOSSIER, generer  # noqa: E402


@pytest.fixture(scope="session")
def dossier_fixtures() -> Path:
    if not (DOSSIER / "classeur_complet.xlsx").exists() or not (DOSSIER / "classeur_vide.xlsx").exists():
        generer()
    return DOSSIER


@pytest.fixture
def classeur_complet(dossier_fixtures) -> Path:
    return dossier_fixtures / "classeur_complet.xlsx"


@pytest.fixture
def classeur_vide(dossier_fixtures) -> Path:
    return dossier_fixtures / "classeur_vide.xlsx"
