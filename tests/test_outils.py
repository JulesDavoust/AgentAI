"""Outils d'exécution de l'agent : exécutés par vba2py lui-même, sans terminal."""

import asyncio
from pathlib import Path
import sys

import pytest

from vba2py.guardrails import Guardrails
from vba2py.outils import creer_outils, creer_serveur, executer_script, lancer_tests, verifier_syntaxe


@pytest.fixture
def projet(tmp_path):
    d = tmp_path / "projet"
    (d / "tests").mkdir(parents=True)
    (d / "_source_vba").mkdir()
    (d / "tests" / "test_a.py").write_text(
        "import pytest\n\ndef test_ok():\n    assert True\n\n@pytest.mark.skip(reason='démo')\ndef test_saute():\n    pass\n"
    )
    (d / "tests" / "test_b.py").write_text("def test_ko():\n    assert 1 == 2\n")
    return d


def test_lancer_tests_tout_ou_cible(projet):
    code, sortie = lancer_tests(projet)
    assert code == 1 and "1 failed, 1 passed, 1 skipped" in sortie
    code, sortie = lancer_tests(projet, cibles=["tests/test_a.py::test_ok"])
    assert code == 0 and "1 passed" in sortie
    code, sortie = lancer_tests(projet, mot_cle="ko", arreter_au_premier_echec=True)
    assert code == 1 and "assert 1 == 2" in sortie


def test_verifier_syntaxe(projet):
    (projet / "bon.py").write_text("x = 1\n")
    (projet / "mauvais.py").write_text("def f(:\n")
    assert verifier_syntaxe(projet, ["bon.py"])[0] == 0
    code, sortie = verifier_syntaxe(projet, ["bon.py", "mauvais.py", "absent.py"])
    assert code == 1 and "mauvais.py" in sortie and "absent.py : fichier introuvable" in sortie


def test_executer_script_utilise_le_python_de_vba2py(projet):
    (projet / "tests" / "gen.py").write_text(
        "import sys, pathlib\nprint(sys.executable)\nprint(sys.argv[1:])\npathlib.Path('tests/fait.txt').write_text('é')\n"
    )
    code, sortie = executer_script(projet, "tests/gen.py", ["a", "b c"])
    assert code == 0
    assert sys.executable in sortie and "['a', 'b c']" in sortie  # arguments transmis sans shell
    assert (projet / "tests" / "fait.txt").read_text(encoding="utf-8") == "é"


def appeler(outils, nom, args):
    [t] = [t for t in outils if t.name == nom]
    resultat = asyncio.run(t.handler(args))
    return resultat["is_error"], resultat["content"][0]["text"]


def test_outils_refusent_puis_executent(projet):
    outils = creer_outils(projet, Guardrails(projet))
    assert sorted(t.name for t in outils) == ["executer_script", "lancer_tests", "verifier_syntaxe"]
    est_erreur, texte = appeler(outils, "executer_script", {"script": "_source_vba/x.py"})
    assert est_erreur and texte.startswith("Refusé par les garde-fous")
    est_erreur, texte = appeler(outils, "lancer_tests", {"cibles": ["tests/test_a.py"]})
    assert not est_erreur and "Code de retour : 0" in texte and "1 passed, 1 skipped" in texte
    est_erreur, texte = appeler(outils, "lancer_tests", {})
    assert est_erreur and "1 failed" in texte  # échec de test : signalé comme erreur à l'agent
    est_erreur, texte = appeler(outils, "verifier_syntaxe", {"fichiers": ["tests/test_b.py"]})
    assert not est_erreur and "Syntaxe correcte" in texte


def test_serveur_mcp_cree(projet):
    serveur = creer_serveur(projet, Guardrails(projet))
    assert serveur["name"] == "vba2py" and serveur["type"] == "sdk"


def test_outil_reference_seulement_sur_demande(projet):
    assert "executer_macro_vba" not in [t.name for t in creer_outils(projet, Guardrails(projet))]
    assert "executer_macro_vba" in [t.name for t in creer_outils(projet, Guardrails(projet), reference_vba=True)]


@pytest.fixture
def projet_vba(projet):
    (projet / "_source_vba" / "Module1.bas").write_text(
        'Attribute VB_Name = "Module1"\nPublic Sub Lancer()\nEnd Sub\nPrivate Sub Interne()\nEnd Sub\n', encoding="utf-8")
    (projet / "tests" / "fixtures").mkdir()
    (projet / "tests" / "fixtures" / "classeur.xlsx").write_bytes(b"PK")
    return projet


def test_garde_fous_de_la_reference(projet_vba):
    g = Guardrails(projet_vba)
    ok = {"macro": "Module1.Lancer", "fixture": "tests/fixtures/classeur.xlsx", "nom": "lancer"}
    outil = "mcp__vba2py__executer_macro_vba"
    assert g.check(outil, ok)[0]
    for modif in ({"macro": "Module1.Interne"}, {"macro": "Lancer"}, {"fixture": "../ailleurs.xlsx"},
                  {"fixture": "tests/fixtures/absent.xlsx"}, {"fixture": "tests/test_a.py"},
                  {"nom": "../evasion"}, {"nom": "Majuscules"}):
        assert not g.check(outil, {**ok, **modif})[0], modif


def test_macro_de_reference(projet_vba):
    from vba2py.outils import macro_de_reference

    def faux_lanceur(parametres, delai):
        Path(parametres["journal"]).write_bytes(b"DEBUT\tModule1.Lancer\nFIN\tmacro terminee\n")
        return "OK"

    code, texte = macro_de_reference(projet_vba, "Module1.Lancer", "tests/fixtures/classeur.xlsx", [], "lancer",
                                     lanceur=faux_lanceur)
    assert code == 0 and "Statut : OK" in texte and "tests/references/lancer/" in texte
    assert (projet_vba / "tests" / "references" / "lancer" / "resultat.json").exists()

    code, texte = macro_de_reference(projet_vba, "Module1.Lancer", "tests/fixtures/classeur.xlsx", [], "lancer",
                                     lanceur=lambda p, d: "ECHEC délai dépassé (300 s)")
    assert code == 1 and "ECHEC délai dépassé" in texte
