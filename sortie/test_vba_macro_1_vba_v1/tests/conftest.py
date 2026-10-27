from __future__ import annotations

from pathlib import Path

import pytest

from tests import generer_donnees_fictives as gen


@pytest.fixture(scope="session", autouse=True)
def _fixtures_presentes():
    if not (gen.FIXTURES / "classeur_complet.xlsx").exists() or not (gen.FIXTURES / "classeur_vide.xlsx").exists():
        gen.generer()


@pytest.fixture
def classeur_complet() -> Path:
    return gen.FIXTURES / "classeur_complet.xlsx"


@pytest.fixture
def classeur_vide() -> Path:
    return gen.FIXTURES / "classeur_vide.xlsx"


@pytest.fixture(autouse=True)
def _bloquer_actions_reelles(monkeypatch):
    """Aucun test ne peut envoyer de mail ni lancer Excel."""
    def interdit(*args, **kwargs):
        raise RuntimeError("Action réelle interdite dans les tests (mail / Excel)")

    import bibliotheque.excel_com as excel_com
    from connecteurs import excel as c_excel, outlook as c_outlook

    monkeypatch.setattr(excel_com, "ouvrir_excel", interdit)
    monkeypatch.setattr(c_outlook, "envoyer_mail", interdit)
    monkeypatch.setattr(c_excel, "appliquer_etapes_excel", interdit)
