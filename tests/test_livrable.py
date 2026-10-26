"""Vérification du livrable de l'agent : interface.json, structure du rapport, sections remplies par le code."""

import copy
import json

import pytest

from vba2py.extract import SOURCE_DIRNAME
from vba2py.livrable import PROMPTS_DIR, charger_interface, commandes, nom_script, nom_sous_commande, verifier_projet

# Deux modules VBA : « Macro » (une macro lançable + un outil) et « Ecran » (deux macros lançables)
VBA_MACRO = '''Attribute VB_Name = "Macro"
Public Sub Lancer()
End Sub
Sub Outil()
End Sub
Private Sub Interne()
End Sub
Public Sub AvecParametre(x As Long)
End Sub
'''
VBA_ECRAN = '''Attribute VB_Name = "Ecran"
Public Sub Ouvrir_Ecran()
End Sub
Public Sub Effacer()
End Sub
'''

SCRIPT_MACRO = '''import argparse

def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("classeur")
    p.add_argument("--dossier-sortie", default="sortie")
    p.add_argument("--mode-mail", choices=["envoyer", "afficher", "aucun"], default="envoyer")
    return p.parse_args(argv)

if __name__ == "__main__":
    main()
'''

SCRIPT_ECRAN = '''import argparse

def main(argv=None):
    p = argparse.ArgumentParser()
    sous = p.add_subparsers(dest="commande", required=True)
    o = sous.add_parser("ouvrir-ecran")
    o.add_argument("classeur")
    e = sous.add_parser("effacer")
    e.add_argument("classeur")
    e.add_argument("--saisie", default="")
    return p.parse_args(argv)

if __name__ == "__main__":
    main()
'''

CLASSEUR = {"nom": "classeur", "remplace": "ThisWorkbook", "defaut": None}
INTERFACE = {
    "points_entree": [
        {"script": "macro.py", "sous_commande": None, "procedure": "Macro.Lancer",
         "justification": "bouton « Lancer »", "fixture": "tests/fixtures/classeur.xlsx",
         "arguments": [CLASSEUR,
                       {"nom": "--dossier-sortie", "remplace": "SaveFolder()", "defaut": "sortie"},
                       {"nom": "--mode-mail", "remplace": "SEND_MAIL", "defaut": "envoyer"}]},
        {"script": "ecran.py", "sous_commande": "ouvrir-ecran", "procedure": "Ecran.Ouvrir_Ecran",
         "justification": "non appelée ailleurs", "fixture": "tests/fixtures/classeur.xlsx", "arguments": [CLASSEUR]},
        {"script": "ecran.py", "sous_commande": "effacer", "procedure": "Ecran.Effacer",
         "justification": "commentaire « bouton »", "fixture": "tests/fixtures/classeur.xlsx",
         "arguments": [CLASSEUR, {"nom": "--saisie", "remplace": "InputBox", "defaut": ""}]},
    ],
    "procedures_internes": [{"procedure": "Macro.Outil", "raison": "appelée seulement par Lancer"}],
}


def rapport_conforme() -> str:
    return (PROMPTS_DIR / "rapport_modele.md").read_text(encoding="utf-8").replace("{{PROJET}}", "macro")


def ecrire_interface(projet, interface):
    (projet / "interface.json").write_text(json.dumps(interface), encoding="utf-8")


def erreurs_avec(projet, modifier):
    interface = copy.deepcopy(INTERFACE)
    modifier(interface)
    ecrire_interface(projet, interface)
    return verifier_projet(projet).erreurs


@pytest.fixture
def projet(tmp_path):
    d = tmp_path / "macro"
    source = d / SOURCE_DIRNAME
    source.mkdir(parents=True)
    (source / "EXTRACTION.md").write_text("# Extraction : x.xlsm\n\n| Module | Type |\n", encoding="utf-8")
    (source / "Macro.bas").write_text(VBA_MACRO, encoding="utf-8")
    (source / "Ecran.bas").write_text(VBA_ECRAN, encoding="utf-8")
    (source / "structure_classeur.json").write_text(
        json.dumps({"feuilles": [], "boutons": [{"feuille": "Accueil", "texte": "Lancer", "macro": "Lancer"}]}),
        encoding="utf-8",
    )
    (d / "tests" / "fixtures").mkdir(parents=True)
    (d / "tests" / "fixtures" / "classeur.xlsx").write_bytes(b"")
    (d / "tests" / "test_ok.py").write_text("def test_ok():\n    assert True\n")
    (d / "macro.py").write_text(SCRIPT_MACRO)
    (d / "ecran.py").write_text(SCRIPT_ECRAN)
    (d / "requirements.txt").write_text("openpyxl\n")
    ecrire_interface(d, INTERFACE)
    (d / "RAPPORT.md").write_text(rapport_conforme(), encoding="utf-8")
    return d


def test_noms_deduits_du_vba():
    assert nom_sous_commande("Demo_EcrireEtLire") == "demo-ecrire-et-lire"
    assert nom_sous_commande("SimuMainframe.MVS_Entree") == "mvs-entree"
    assert nom_sous_commande("GenererCSVDuJourV3") == "generer-csv-du-jour-v3"
    assert nom_script("CréationCSV") == "creation_csv.py"
    assert nom_script("TestInterfaceMVS") == "test_interface_mvs.py"
    assert nom_script("marco_25_10") == "marco_25_10.py"


def test_livrable_conforme_et_sections_remplies(projet):
    verif = verifier_projet(projet)
    assert verif.ok, verif.erreurs
    assert "1 passed" in verif.resultat_tests
    rapport = (projet / "RAPPORT.md").read_text(encoding="utf-8")
    assert "python macro.py tests/fixtures/classeur.xlsx --dossier-sortie sortie_test --mode-mail aucun" in rapport
    assert "### Ecran.Effacer" in rapport and "python ecran.py effacer --help" in rapport
    assert "| `--mode-mail` | SEND_MAIL | `envoyer` |" in rapport
    assert "- Macro.Outil : appelée seulement par Lancer" in rapport
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


def test_macro_lancable_oubliee(projet):
    erreurs = erreurs_avec(projet, lambda i: i["procedures_internes"].clear())
    assert erreurs == ["macro Macro.Outil : ni point d'entrée, ni procédure interne déclarée dans interface.json"]


def test_macro_de_bouton_obligatoirement_point_d_entree(projet):
    def modifier(i):
        i["points_entree"].pop(0)
        i["procedures_internes"].append({"procedure": "Macro.Lancer", "raison": "x"})
    assert any("reliée à un bouton" in e for e in erreurs_avec(projet, modifier))


def test_procedure_inconnue_ou_avec_parametres(projet):
    def modifier(i):
        i["procedures_internes"].append({"procedure": "Macro.AvecParametre", "raison": "x"})
    assert any("Macro.AvecParametre n'est pas une macro lançable" in e for e in erreurs_avec(projet, modifier))


def test_noms_de_script_et_sous_commandes_imposes(projet):
    def mauvais_script(i):
        i["points_entree"][0]["script"] = "lancement.py"
    assert any("le script doit s'appeler macro.py" in e for e in erreurs_avec(projet, mauvais_script))

    def mauvaise_sous_commande(i):
        i["points_entree"][1]["sous_commande"] = "ouvrir"
    assert any("la sous-commande doit être « ouvrir-ecran »" in e for e in erreurs_avec(projet, mauvaise_sous_commande))

    def sous_commande_inutile(i):
        i["points_entree"][0]["sous_commande"] = "lancer"
    assert any("un seul point d'entrée" in e for e in erreurs_avec(projet, sous_commande_inutile))


def test_options_comparees_au_help_de_chaque_sous_commande(projet):
    def modifier(i):
        i["points_entree"][2]["arguments"] = [CLASSEUR]
    erreurs = erreurs_avec(projet, modifier)
    assert erreurs == ["ecran.py effacer : options du script absentes de interface.json : ['--saisie']"]


def test_seul_classeur_positionnel_en_premier(projet):
    def modifier(i):
        i["points_entree"][0]["arguments"] = i["points_entree"][0]["arguments"][1:] + [CLASSEUR]
    assert any("seul « classeur » peut être positionnel" in e for e in erreurs_avec(projet, modifier))


def test_interface_precedente_respectee(projet, tmp_path):
    precedente = tmp_path / "precedente.json"
    ancienne = copy.deepcopy(INTERFACE)
    ancienne["points_entree"] = ancienne["points_entree"][:2]  # un point d'entrée ajouté depuis : accepté
    precedente.write_text(json.dumps(ancienne), encoding="utf-8")
    assert verifier_projet(projet, precedente).ok

    ancienne["points_entree"][0]["arguments"] = [CLASSEUR, {"nom": "--dossier", "remplace": "x"}]
    ancienne["points_entree"].append({"script": "x.py", "procedure": "Macro.Disparue", "arguments": []})
    precedente.write_text(json.dumps(ancienne), encoding="utf-8")
    erreurs = verifier_projet(projet, precedente).erreurs
    assert any("Macro.Lancer : la commande a changé" in e for e in erreurs)
    assert any("Macro.Disparue présent dans la traduction précédente : supprimé" in e for e in erreurs)


def test_ancien_format_converti(tmp_path):
    ancien = tmp_path / "ancien.json"
    ancien.write_text(json.dumps({
        "script": "macro.py", "procedure_principale": "LancerPipelineComplet",
        "fixture_principale": "tests/fixtures/classeur.xlsx", "arguments": [CLASSEUR],
    }), encoding="utf-8")
    interface = charger_interface(ancien)
    [point] = interface["points_entree"]
    assert (point["script"], point["sous_commande"], point["procedure"]) == ("macro.py", None, "LancerPipelineComplet")
    assert interface["procedures_internes"] == []


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


def test_commandes_sans_mail_ni_classeur():
    interface = {"points_entree": [{"script": "outil.py", "sous_commande": None, "procedure": "Outil.Faire",
                                    "fixture": None, "arguments": []}], "procedures_internes": []}
    texte = commandes(interface)
    assert "--mode-mail" not in texte and "python outil.py        # essai sur données fictives" in texte


CONFTEST_MARQUEUR = 'def pytest_configure(config):\n    config.addinivalue_line("markers", "excel_reel: Excel réel")\n'


@pytest.mark.parametrize(
    "test_excel, ok, attendu",
    [
        (None, False, "❌ aucun test"),
        ("import pytest\n@pytest.mark.excel_reel\n@pytest.mark.skip(reason='pas d Excel')\ndef test_x():\n    pass\n",
         True, "⚠️ non exécutés"),
        ("import pytest\n@pytest.mark.excel_reel\ndef test_x():\n    pass\n", True, "✅ **1 passed"),
        ("import pytest\n@pytest.mark.excel_reel\ndef test_x():\n    assert False\n", False, "❌ **1 failed"),
    ],
)
def test_etapes_excel_verifiees_sur_excel_reel(projet, test_excel, ok, attendu):
    (projet / "connecteurs").mkdir()
    (projet / "connecteurs" / "excel.py").write_text("", encoding="utf-8")
    (projet / "tests" / "conftest.py").write_text(CONFTEST_MARQUEUR, encoding="utf-8")
    if test_excel:
        (projet / "tests" / "test_excel_reel.py").write_text(test_excel, encoding="utf-8")
    verif = verifier_projet(projet)
    assert verif.ok is ok, verif.erreurs
    assert verif.resultat_excel.startswith(attendu)
    assert "Tests sur Excel réel" in (projet / "RAPPORT.md").read_text(encoding="utf-8")
