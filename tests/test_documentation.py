"""Documentation facultative jointe aux macros (--doc et dossiers <classeur>_doc/)."""

import pytest

import main
from conftest import make_export_dir, make_xlsm
from vba2py.documentation import collecter, documentation_du_projet, ecrire_documentation
from vba2py.extract import SOURCE_DIRNAME, ExtractionError, discover_projects


def doc(chemin, contenu="x"):
    chemin.parent.mkdir(parents=True, exist_ok=True)
    chemin.write_text(contenu, encoding="utf-8")
    return chemin


def test_collecter_fichiers_et_dossiers(tmp_path):
    a = doc(tmp_path / "spec.md")
    doc(tmp_path / "docs" / "regles.txt")
    doc(tmp_path / "docs" / "sous" / "guide.pdf")
    doc(tmp_path / "docs" / ".cache")  # fichier caché ignoré
    fichiers = collecter([a, tmp_path / "docs", a])
    assert [f.name for f in fichiers] == ["spec.md", "regles.txt", "guide.pdf"]  # sans doublon


def test_collecter_refuse_formats_et_chemins_inconnus(tmp_path):
    doc(tmp_path / "docs" / "spec.docx")
    with pytest.raises(ExtractionError, match="spec.docx"):
        collecter([tmp_path / "docs"])
    with pytest.raises(ExtractionError, match="introuvable"):
        collecter([tmp_path / "absent.pdf"])


def test_documentation_commune_et_rattachee(tmp_path):
    classeur = make_xlsm(tmp_path / "Ventes.xlsm")
    doc(tmp_path / "Ventes_doc" / "regles.md")
    export = make_export_dir(tmp_path / "Paie_vba")
    doc(tmp_path / "Paie_vba_doc" / "paie.txt")
    commun = doc(tmp_path / "commun.md")

    docs = documentation_du_projet(classeur, [commun])
    assert [(d.chemin.name, d.portee) for d in docs] == [("commun.md", "commune"), ("regles.md", "rattachée")]
    assert [d.chemin.name for d in documentation_du_projet(export, [])] == ["paie.txt"]
    assert documentation_du_projet(make_xlsm(tmp_path / "Autre.xlsm"), []) == []


def test_dossiers_doc_ignores_par_la_decouverte(tmp_path):
    make_xlsm(tmp_path / "Ventes.xlsm")
    (tmp_path / "Ventes_doc").mkdir()
    make_xlsm(tmp_path / "Ventes_doc" / "specification.xlsm")  # pas traité comme un projet
    assert [p.relative_path.as_posix() for p in discover_projects(tmp_path)] == ["ventes"]


def test_copie_index_et_homonymes(tmp_path):
    from vba2py.documentation import Document

    a = doc(tmp_path / "a" / "guide.md", "A")
    b = doc(tmp_path / "b" / "guide.md", "B")
    projet = tmp_path / "sortie" / "ventes"
    ecrire_documentation(projet, [Document(a, "commune"), Document(b, "rattachée")])
    cible = projet / SOURCE_DIRNAME / "documentation"
    assert sorted(p.name for p in cible.iterdir()) == ["INDEX.md", "guide.md", "guide_2.md"]
    assert (cible / "guide_2.md").read_text(encoding="utf-8") == "B"
    assert "| guide_2.md | rattachée |" in (cible / "INDEX.md").read_text(encoding="utf-8")

    ecrire_documentation(projet, [])  # nouvelle commande sans documentation : rien ne traîne
    assert not cible.exists()


def test_commande_extraire_avec_doc(tmp_path, capsys):
    (tmp_path / "entree").mkdir()
    make_xlsm(tmp_path / "entree" / "Ventes.xlsm")
    doc(tmp_path / "entree" / "Ventes_doc" / "regles.md")
    commun = doc(tmp_path / "commun.pdf")
    assert main.main(["extraire", str(tmp_path / "entree"), "-o", str(tmp_path / "sortie"), "--doc", str(commun)]) == 0
    source = tmp_path / "sortie" / "ventes" / SOURCE_DIRNAME
    assert (source / "documentation" / "commun.pdf").exists()
    assert (source / "documentation" / "regles.md").exists()
    extraction = (source / "EXTRACTION.md").read_text(encoding="utf-8")
    assert "| commun.pdf | commune |" in extraction and "| regles.md | rattachée |" in extraction
    assert "Documentation : commun.pdf (commune), regles.md (rattachée)" in capsys.readouterr().out


def test_commande_sans_doc(tmp_path):
    make_xlsm(tmp_path / "Ventes.xlsm")
    assert main.main(["extraire", str(tmp_path / "Ventes.xlsm"), "-o", str(tmp_path / "sortie")]) == 0
    source = tmp_path / "sortie" / "ventes" / SOURCE_DIRNAME
    assert not (source / "documentation").exists()
    assert "Aucune documentation fournie." in (source / "EXTRACTION.md").read_text(encoding="utf-8")


def test_commande_doc_format_refuse(tmp_path, capsys):
    make_xlsm(tmp_path / "Ventes.xlsm")
    mauvais = doc(tmp_path / "spec.docx")
    assert main.main(["extraire", str(tmp_path / "Ventes.xlsm"), "--doc", str(mauvais), "-o", str(tmp_path / "s")]) == 2
    assert "non pris en charge" in capsys.readouterr().err
