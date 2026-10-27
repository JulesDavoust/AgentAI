import shutil
import sys
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from bibliotheque import excel_com  # noqa: E402
from connecteurs import excel as connecteur_excel  # noqa: E402
from connecteurs import outlook as connecteur_outlook  # noqa: E402
from tests import generer_donnees_fictives  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures"


def pytest_configure(config):
    config.addinivalue_line("markers", "excel_reel: test qui pilote un vrai Excel")


@pytest.fixture(scope="session", autouse=True)
def _donnees_fictives():
    if not (FIXTURES / "classeur_complet.xlsx").exists() or not (FIXTURES / "classeur_vide.xlsx").exists():
        generer_donnees_fictives.generer(FIXTURES)


@pytest.fixture(autouse=True)
def _bloquer_actions_reelles(request, monkeypatch):
    def interdit(*args, **kwargs):
        raise AssertionError("Action réelle interdite dans les tests (mail, Excel)")

    monkeypatch.setattr(connecteur_outlook, "envoyer_mail", interdit)
    if not request.node.get_closest_marker("excel_reel"):
        monkeypatch.setattr(connecteur_excel, "executer_dans_excel", interdit)
        monkeypatch.setattr(excel_com, "ouvrir_excel", interdit)


@pytest.fixture
def classeur_complet(tmp_path):
    """Copie modifiable de la fixture complète."""
    cible = tmp_path / "classeur_complet.xlsx"
    shutil.copy(FIXTURES / "classeur_complet.xlsx", cible)
    return cible


@pytest.fixture
def classeur_vide(tmp_path):
    cible = tmp_path / "classeur_vide.xlsx"
    shutil.copy(FIXTURES / "classeur_vide.xlsx", cible)
    return cible
