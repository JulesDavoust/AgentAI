import json

import pytest

from conftest import make_export_dir, make_xlsm
from vba2py.extract import (
    SOURCE_DIRNAME,
    ExtractionError,
    discover_projects,
    strip_vba_headers,
    write_project_sources,
)


def test_classeur_modules_extraits(xlsm):
    [project] = discover_projects(xlsm)
    assert project.origin == "classeur"
    assert project.relative_path.as_posix() == "ventes_2024"
    modules = {m.name: m for m in project.modules}
    assert "say_hello" in modules["Module1"].code
    assert not modules["Module1"].is_empty
    assert modules["ThisWorkbook"].is_empty  # module de document sans code


def test_structure_sans_valeurs_reelles(xlsm):
    [project] = discover_projects(xlsm)
    structure = project.workbook_structure
    [sheet] = structure["feuilles"]
    assert sheet["nom"] == "Ventes"
    assert [c["en_tete"] for c in sheet["colonnes"]] == ["Date", "Client", "Montant"]
    assert {"nom": "Total", "reference": "Ventes!$C$5"} in structure["plages_nommees"]
    assert structure["formules_exemples"][0]["formule"] == "=SUM(C2:C3)"
    # Les données réelles ne doivent jamais sortir du classeur
    assert "Client confidentiel" not in json.dumps(structure)
    assert "1234.5" not in json.dumps(structure)


def test_ecriture_des_sources(xlsm, tmp_path):
    [project] = discover_projects(xlsm)
    project_dir = write_project_sources(project, tmp_path / "sortie")
    source = project_dir / SOURCE_DIRNAME
    assert (source / "Module1.bas").read_text(encoding="utf-8").startswith('Attribute VB_Name = "Module1"\nSub say_hello()')
    assert not (source / "ThisWorkbook.cls").exists()  # modules vides ignorés
    assert (source / "structure_classeur.json").exists()
    rapport = (source / "EXTRACTION.md").read_text(encoding="utf-8")
    assert "| Module1 | module | 3 | extrait |" in rapport
    assert "| ThisWorkbook | classe/document | 0 | vide (ignoré) |" in rapport


def test_dossier_exporte_encodage_windows(tmp_path):
    export = make_export_dir(tmp_path / "Facturation", 'Sub Calculer()\r\n    \' Évaluer la TVA à 20 %\r\nEnd Sub\r\n')
    [project] = discover_projects(export)
    assert project.origin == "export"
    assert project.relative_path.as_posix() == "facturation"
    module1 = next(m for m in project.modules if m.name == "Module1")
    assert "Évaluer la TVA à 20 %" in module1.code
    assert "\r" not in module1.code
    assert any("Aucun classeur" in w for w in project.warnings)


def test_dossier_exporte_avec_classeur_pour_la_structure(tmp_path):
    export = make_export_dir(tmp_path / "Ventes_vba")
    make_xlsm(export / "Ventes.xlsm")
    [project] = discover_projects(export)
    assert project.origin == "export"  # le classeur ne sert qu'à la structure
    assert project.workbook_structure["feuilles"][0]["nom"] == "Ventes"


def test_arborescence_conservee(tmp_path):
    racine = tmp_path / "entree"
    make_xlsm((racine / "Finance").mkdir(parents=True) or racine / "Finance" / "Budget.xlsm")
    make_xlsm((racine / "RH" / "Paie").mkdir(parents=True) or racine / "RH" / "Paie" / "Salaires été.xlsm")
    make_export_dir(racine / "Commun" / "Outils_vba")
    (racine / "RH" / "notes.xlsx").write_bytes(b"")  # .xlsx : pas de macros, ignoré
    chemins = sorted(p.relative_path.as_posix() for p in discover_projects(racine))
    assert chemins == ["commun/outils_vba", "finance/budget", "rh/paie/salaires_ete"]


def test_strip_vba_headers():
    code = (
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n"
        'Attribute VB_Name = "Feuil1"\nAttribute VB_Exposed = True\n\nSub A()\nEnd Sub   \n\n'
    )
    assert strip_vba_headers(code) == "Sub A()\nEnd Sub"


def test_entree_invalide(tmp_path):
    with pytest.raises(ExtractionError):
        discover_projects(tmp_path / "inexistant.xlsm")
    texte = tmp_path / "notes.txt"
    texte.write_text("x")
    with pytest.raises(ExtractionError):
        discover_projects(texte)
    with pytest.raises(ExtractionError, match="Aucun classeur"):
        discover_projects(tmp_path)


def test_classeur_sans_macros(tmp_path):
    import openpyxl

    path = tmp_path / "vide.xls"
    openpyxl.Workbook().save(tmp_path / "vide.xlsx")
    path.write_bytes((tmp_path / "vide.xlsx").read_bytes())
    with pytest.raises(ExtractionError):
        discover_projects(path)


def test_boutons_et_feuilles_masquees(tmp_path):
    import xlsxwriter

    from conftest import VBA_BIN
    from vba2py.extract import _nom_macro

    path = tmp_path / "Boutons.xlsm"
    wb = xlsxwriter.Workbook(str(path))
    accueil = wb.add_worksheet("Accueil")
    accueil.insert_button("B2", {"macro": "say_hello", "caption": "Dire bonjour"})
    accueil.insert_button("B5", {"macro": "Module1.Autre", "caption": "Autre macro"})
    base = wb.add_worksheet("Base")
    base.write("A1", "x")
    base.hide()
    wb.add_vba_project(str(VBA_BIN))
    wb.close()

    [project] = discover_projects(path)
    structure = project.workbook_structure
    assert structure["boutons"] == [
        {"feuille": "Accueil", "texte": "Dire bonjour", "macro": "say_hello"},
        {"feuille": "Accueil", "texte": "Autre macro", "macro": "Module1.Autre"},
    ]
    assert {f["nom"]: f["visibilite"] for f in structure["feuilles"]} == {"Accueil": "visible", "Base": "masquée"}

    project_dir = write_project_sources(project, tmp_path / "sortie")
    extraction = (project_dir / SOURCE_DIRNAME / "EXTRACTION.md").read_text(encoding="utf-8")
    assert "| Accueil | Dire bonjour | say_hello |" in extraction
    assert "Feuilles masquées : Base" in extraction

    assert _nom_macro("[0]!LancerPipelineComplet") == "LancerPipelineComplet"
    assert _nom_macro("'Mon classeur.xlsm'!Module1.Macro") == "Module1.Macro"
    assert _nom_macro("Macro") == "Macro"
