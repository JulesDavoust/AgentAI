import shutil
from pathlib import Path

import pytest

from connecteurs import outlook
from tests.generer_donnees_fictives import generer

FIXTURES = Path(__file__).parent / "fixtures"

# Fonctions publiques de connecteurs/ à bloquer (ajouter ici tout nouveau connecteur).
FONCTIONS_CONNECTEURS = [(outlook, "envoyer_mail")]


@pytest.fixture(scope="session", autouse=True)
def _fixtures_generees():
    if not (FIXTURES / "classeur_vide.xlsx").exists() or not (FIXTURES / "classeur_donnees.xlsx").exists():
        generer(FIXTURES)


@pytest.fixture(autouse=True)
def _bloquer_connecteurs(monkeypatch):
    """Toute fonction de connecteurs/ lève une erreur si elle est appelée."""
    for mod, nom in FONCTIONS_CONNECTEURS:
        def _interdit(*a, _n=f"{mod.__name__}.{nom}", **k):
            raise AssertionError(f"Appel réel au connecteur {_n} dans un test")
        monkeypatch.setattr(mod, nom, _interdit)


@pytest.fixture
def classeur_vide(tmp_path):
    dest = tmp_path / "vide.xlsx"
    shutil.copy(FIXTURES / "classeur_vide.xlsx", dest)
    return dest


@pytest.fixture
def classeur_donnees(tmp_path):
    dest = tmp_path / "donnees.xlsx"
    shutil.copy(FIXTURES / "classeur_donnees.xlsx", dest)
    return dest
