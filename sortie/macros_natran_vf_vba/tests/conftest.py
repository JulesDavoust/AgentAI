"""Fixtures pytest : données fictives, blocage des actions réelles (mail, Excel, ouverture de fichier)."""

from __future__ import annotations

import functools
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

FIXTURES = Path(__file__).resolve().parent / "fixtures"
FICHIERS = ("classeur_complet.xlsx", "classeur_donnees.xlsx", "classeur_mvs.xlsx", "csv/DAT_20261027.csv")

MAINTENANT = datetime(2026, 10, 27, 14, 30, 5)


def pytest_configure(config):
    config.addinivalue_line("markers", "excel_reel: test qui pilote le vrai Excel (ignoré s'il est absent)")


@pytest.fixture(scope="session", autouse=True)
def donnees_fictives():
    """Crée tests/fixtures/* si un fichier manque."""
    if not all((FIXTURES / f).exists() for f in FICHIERS):
        from tests.generer_donnees_fictives import generer
        generer()
    return FIXTURES


def garde(original, interdit):
    @functools.wraps(original)
    def _f(*a, **k):
        if k.get("session") is None:
            return interdit(*a, **k)
        return original(*a, **k)
    return _f


@pytest.fixture(autouse=True)
def bloquer_actions_reelles(request, monkeypatch):
    """Aucun test ne doit envoyer de mail, lancer Excel ou ouvrir un fichier."""
    def interdit(nom):
        def _f(*a, **k):
            raise AssertionError(f"Action réelle bloquée dans les tests : {nom}")
        return _f

    from connecteurs import outlook, systeme
    monkeypatch.setattr(outlook, "envoyer_mail", interdit("outlook.envoyer_mail"))
    monkeypatch.setattr(systeme, "ouvrir_fichier", interdit("systeme.ouvrir_fichier"))
    if request.node.get_closest_marker("excel_reel"):
        return
    from bibliotheque import excel_com
    from connecteurs import excel
    monkeypatch.setattr(excel_com, "ouvrir_excel", interdit("excel_com.ouvrir_excel"))
    # Les fonctions d'entrée sont bloquées sauf si le test injecte une fausse session (session=...).
    for nom in ("finaliser_initiale", "finaliser_carte", "finaliser_csv_v3", "ajuster_feuille"):
        monkeypatch.setattr(excel, nom, garde(getattr(excel, nom), interdit(f"excel.{nom}")))


@pytest.fixture
def maintenant():
    return lambda: MAINTENANT


@pytest.fixture
def copie(tmp_path):
    """copie("classeur_donnees.xlsx") -> chemin d'une copie modifiable dans tmp_path."""
    def _copie(nom):
        cible = tmp_path / Path(nom).name
        shutil.copy(FIXTURES / nom, cible)
        return cible
    return _copie


class FauxRandom:
    """Générateur dont ``random()`` renvoie toujours la même valeur (résultats calculables à la main)."""

    def __init__(self, valeur=0.5):
        self.valeur = valeur

    def random(self):
        return self.valeur


@pytest.fixture
def rng():
    return FauxRandom(0.5)
