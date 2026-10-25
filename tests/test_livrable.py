"""Vérification du livrable de l'agent : interface.json, structure du rapport, sections remplies par le code."""

import json

import pytest

from vba2py.extract import SOURCE_DIRNAME
from vba2py.livrable import PROMPTS_DIR, commandes, verifier_projet

SCRIPT = '''import argparse

def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("classeur")
    p.add_argument("--dossier-sortie", default="sortie")
    p.add_argument("--mode-mail", choices=["envoyer", "afficher", "aucun"], default="envoyer")
    return p.parse_args(argv)

if __name__ == "__main__":
    main()
'''

INTERFACE = {
    "script": "macro.py",
    "procedure_principale": "Lancer",
    "fixture_principale": "tests/fixtures/classeur.xlsx",
    "arguments": [
        {"nom": "classeur", "remplace": "ThisWorkbook", "defaut": None},
        {"nom": "--dossier-sortie", "remplace": "SaveFolder()", "defaut": "sortie"},
        {"nom": "--mode-mail", "remplace": "SEND_MAIL", "defaut": "envoyer"},
    ],
}


def rapport_conforme() -> str:
    return (PROMPTS_DIR / "rapport_modele.md").read_text(encoding="utf-8").replace("{{PROJET}}", "macro")


@pytest.fixture
def projet(tmp_path):
    d = tmp_path / "macro"
    (d / SOURCE_DIRNAME).mkdir(parents=True)
    (d / SOURCE_DIRNAME / "EXTRACTION.md").write_text("# Extraction : x.xlsm\n\n| Module | Type |\n", encoding="utf-8")
    (d / "tests" / "fixtures").mkdir(parents=True)
    (d / "tests" / "fixtures" / "classeur.xlsx").write_bytes(b"")
    (d / "tests" / "test_ok.py").write_text("def test_ok():\n    assert True\n")
    (d / "macro.py").write_text(SCRIPT)
    (d / "requirements.txt").write_text("openpyxl\n")
    (d / "interface.json").write_text(json.dumps(INTERFACE), encoding="utf-8")
    (d / "RAPPORT.md").write_text(rapport_conforme(), encoding="utf-8")
    return d


def test_livrable_conforme_et_sections_remplies(projet):
    verif = verifier_projet(projet)
    assert verif.ok, verif.erreurs
    assert "1 passed" in verif.resultat_tests
    rapport = (projet / "RAPPORT.md").read_text(encoding="utf-8")
    assert "python macro.py tests/fixtures/classeur.xlsx --dossier-sortie sortie_test --mode-mail aucun" in rapport
    assert "| `--mode-mail` | SEND_MAIL | `envoyer` |" in rapport
    assert "| Module | Type |" in rapport  # extraction recopiée
    # idempotent : une seconde vérification ne duplique rien
    verifier_projet(projet)
    assert (projet / "RAPPORT.md").read_text(encoding="utf-8").count("Vérification indépendante") == 1


def test_bibliotheque_modifiee(projet):
    from vba2py.bibliotheque import fichiers_bibliotheque, installer_bibliotheque

    cible = installer_bibliotheque(projet)
    assert verifier_projet(projet).ok
    (cible / fichiers_bibliotheque()[0].name).write_text("# modifié", encoding="utf-8")
    assert any("bibliotheque/ modifiée" in e for e in verifier_projet(projet).erreurs)


def test_option_non_declaree(projet):
    interface = dict(INTERFACE, arguments=INTERFACE["arguments"][:2])
    (projet / "interface.json").write_text(json.dumps(interface), encoding="utf-8")
    erreurs = verifier_projet(projet).erreurs
    assert any("options du script absentes de interface.json : ['--mode-mail']" in e for e in erreurs)


def test_premier_argument_classeur(projet):
    interface = dict(INTERFACE, arguments=INTERFACE["arguments"][1:])
    (projet / "interface.json").write_text(json.dumps(interface), encoding="utf-8")
    assert any("premier argument" in e for e in verifier_projet(projet).erreurs)


def test_interface_precedente_respectee(projet, tmp_path):
    precedente = tmp_path / "precedente.json"
    precedente.write_text(json.dumps(INTERFACE), encoding="utf-8")
    assert verifier_projet(projet, precedente).ok

    autre = dict(INTERFACE, arguments=INTERFACE["arguments"][:1] + [{"nom": "--dossier", "remplace": "x"}])
    precedente.write_text(json.dumps(autre), encoding="utf-8")
    assert any("ont changé" in e for e in verifier_projet(projet, precedente).erreurs)


def test_rapport_non_conforme(projet):
    rapport = rapport_conforme().replace("## 6. TODO et fonctionnalités non couvertes", "## TODO")
    (projet / "RAPPORT.md").write_text(rapport, encoding="utf-8")
    [erreur] = verifier_projet(projet).erreurs
    assert "manquants" in erreur and "## TODO" in erreur


def test_tests_en_echec(projet):
    (projet / "tests" / "test_ko.py").write_text("def test_ko():\n    assert False\n")
    verif = verifier_projet(projet)
    assert any(e.startswith("tests en échec") for e in verif.erreurs)
    assert "❌" in (projet / "RAPPORT.md").read_text(encoding="utf-8")


def test_fichiers_manquants(tmp_path):
    (tmp_path / "vide").mkdir()
    [erreur] = verifier_projet(tmp_path / "vide").erreurs
    assert "RAPPORT.md" in erreur and "interface.json" in erreur


def test_commandes_sans_mail():
    interface = dict(INTERFACE, arguments=INTERFACE["arguments"][:1])
    texte = commandes(interface)
    assert "--mode-mail" not in texte and "python macro.py tests/fixtures/classeur.xlsx\n" in texte
