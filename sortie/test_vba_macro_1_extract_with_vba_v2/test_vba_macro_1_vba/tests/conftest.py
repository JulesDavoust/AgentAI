import sys
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from tests import generer_donnees_fictives  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures"
NOMS = ["classeur_complet.xlsx", "classeur_jour_present.xlsx", "classeur_vierge.xlsx"]


@pytest.fixture(scope="session", autouse=True)
def _fixtures_presentes():
    if not all((FIXTURES / n).exists() for n in NOMS):
        generer_donnees_fictives.generer(FIXTURES)


@pytest.fixture(autouse=True)
def _bloquer_connecteurs(monkeypatch):
    """Toute fonction de connecteurs/ appelée sans fausse fonction explicite lève une erreur."""
    from connecteurs import outlook

    def interdit(*args, **kwargs):
        raise AssertionError("Appel réel à un connecteur pendant un test")

    monkeypatch.setattr(outlook, "envoyer_mail", interdit)


@pytest.fixture
def classeur_complet():
    return FIXTURES / "classeur_complet.xlsx"


@pytest.fixture
def classeur_jour_present():
    return FIXTURES / "classeur_jour_present.xlsx"


@pytest.fixture
def classeur_vierge():
    return FIXTURES / "classeur_vierge.xlsx"
