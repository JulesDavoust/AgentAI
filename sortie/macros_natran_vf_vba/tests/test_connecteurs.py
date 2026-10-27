"""Étapes Excel (avec EnregistreurCOM), pipeline d'entrée avec fausse session, Outlook, ExportVBA."""

import sys
import types
from pathlib import Path

import pytest

import export_vba
from bibliotheque.excel_com import EnregistreurCOM, xl
from connecteurs import excel as cx
from connecteurs import outlook


def indice(wb, fin):
    """Position de la première instruction dont le chemin se termine par `fin`."""
    for i, (_, chemin, _) in enumerate(wb.journal):
        if chemin.endswith(fin):
            return i
    raise AssertionError(f"instruction absente : {fin}")


FILTRE = "wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode"


def test_creer_visuel_excel_toutes_les_instructions_dans_l_ordre():
    wb = EnregistreurCOM("wb")
    cx.creer_visuel_excel(wb, "Feuille1", "tblDonnees")
    assert wb.appels("PivotCaches().Create") == [((), {"SourceType": xl.xlDatabase, "SourceData": "tblDonnees"})]
    (_, kw), = wb.appels("CreatePivotTable")
    assert kw["TableName"] == "TCD_Volume"
    assert wb.affectations("PivotFields('Date').Orientation") == [xl.xlRowField]
    assert wb.affectations("PivotFields('Region').Orientation") == [xl.xlColumnField]
    (args, _), = wb.appels("AddDataField")
    assert args[1:] == ("Volume (GWh)", xl.xlSum)
    assert wb.affectations(".NumberFormat") == ["#,##0.0", "dd/mm"]
    (_, kw), = wb.appels("ChartObjects().Add")
    assert (kw["Width"], kw["Height"]) == (520, 300)
    assert wb.affectations("Chart.ChartType") == [xl.xlLine]
    assert wb.affectations("Chart.HasTitle") == [True]
    assert wb.affectations("ChartTitle.Text") == ["Volumes distribues par region (GWh)"]
    assert len(wb.appels("SetSourceData")) == 1 and len(wb.appels("AutoFit")) == 1 and len(wb.appels("Activate")) == 1
    ordre = [indice(wb, f) for f in ("PivotCaches().Create", "CreatePivotTable", "PivotFields('Date').Orientation",
                                     "PivotFields('Region').Orientation", "AddDataField", "ChartObjects().Add",
                                     "Chart.SetSourceData", "Chart.ChartType", "Chart.HasTitle",
                                     "ChartTitle.Text", "AutoFit", "Activate")]
    assert ordre == sorted(ordre)


def test_filtre_excel_avec_region():
    wb = EnregistreurCOM("wb", reponses={FILTRE: True})
    cx.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "Nord")
    assert len(wb.appels("AutoFilter.ShowAllData")) == 1
    assert wb.appels("Range.AutoFilter") == [((), {"Field": 2, "Criteria1": "Nord"})]
    assert indice(wb, "ShowAllData") < indice(wb, "Range.AutoFilter")


def test_filtre_excel_sans_filtre_existant_ni_region():
    wb = EnregistreurCOM("wb", reponses={FILTRE: False})
    cx.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "")
    assert wb.appels("ShowAllData") == [] and wb.appels("Range.AutoFilter") == []


def test_ajuster_colonnes_excel():
    wb = EnregistreurCOM("wb")
    cx.ajuster_colonnes_excel(wb, "Carte_V3", "A:H", 24)
    assert wb.appels("Columns")[0][0] == ("A:H",) and len(wb.appels("AutoFit")) == 1
    assert wb.appels("Worksheets") == [(("Carte_V3",), {})]
    assert wb.affectations("ColumnWidth") == [24]


@pytest.mark.parametrize("fonction,args,attendu", [
    (cx.traiter_classeur_initiale, dict(feuille_donnees="D", feuille_visuel="V", nom_table="t", region="Sud",
                                        ajuster_donnees=True), ["A:E"]),
    (cx.traiter_classeur_carte, dict(feuille_donnees="D", feuille_visuel="V", nom_table="t", region="",
                                     ajuster_donnees=False), ["A:H"]),
    (cx.traiter_classeur_csv, dict(feuille_donnees="D", feuille_visuel="V"), ["A:F", "A:H"]),
])
def test_entrees_excel_ouvrent_une_session_sur_un_chemin_absolu(monkeypatch, tmp_path, fonction, args, attendu):
    fichier = tmp_path / "c.xlsx"
    fichier.write_bytes(b"x")
    faux_wb = EnregistreurCOM("wb", reponses={"wb.Worksheets('D').ListObjects('t').AutoFilter.FilterMode": False})
    ouverts = []

    class Session:
        def __enter__(self_):
            return faux_wb

        def __exit__(self_, *exc):
            return False

    monkeypatch.setattr(cx, "session_excel", lambda chemin, **kw: ouverts.append(chemin) or Session())
    fonction(str(fichier), **args)
    assert ouverts == [str(fichier)]
    colonnes = [a[0][0] for a in faux_wb.appels("Columns")]
    assert colonnes[:len(attendu)] == attendu or all(c in colonnes for c in attendu)


def test_exporter_composants_excel_sans_composant():
    wb = EnregistreurCOM("wb")
    assert cx.exporter_composants_excel(wb, "dossier") == 0


def test_exporter_composants_excel_types():
    class Composant:
        def __init__(self, nom, type_):
            self.Name, self.Type, self.exportes = nom, type_, []
            self.Export = self.exportes.append

    class Collection:
        VBComponents = [Composant("Mod1", 1), Composant("Feuil1", 100), Composant("Classe", 2),
                        Composant("Form", 3), Composant("Autre", 9)]

    wb = types.SimpleNamespace(VBProject=Collection)
    assert cx.exporter_composants_excel(wb, "D") == 4
    noms = [c.exportes for c in Collection.VBComponents]
    assert noms == [["D\\Mod1.bas"], ["D\\Feuil1.cls"], ["D\\Classe.cls"], ["D\\Form.frm"], []]


def test_export_vba_dossier_et_erreur(tmp_path):
    classeur = tmp_path / "Ventes.xlsm"
    classeur.write_bytes(b"x")
    appels = []

    def faux_export(chemin, dossier):
        assert Path(chemin).is_absolute() and Path(dossier).exists()
        appels.append(dossier)
        return 3

    dossier, nombre = export_vba.exporter_modules_vba(classeur, faux_export)
    assert dossier == tmp_path.resolve() / "Ventes_vba" and nombre == 3 and appels == [str(dossier)]

    def export_refuse(chemin, dossier):
        raise RuntimeError("acces refuse")

    with pytest.raises(RuntimeError, match="acces refuse"):
        export_vba.exporter_modules_vba(classeur, export_refuse)
    with pytest.raises(FileNotFoundError):
        export_vba.exporter_modules_vba(tmp_path / "absent.xlsm", faux_export)


def test_outlook_envoyer_et_afficher(monkeypatch, tmp_path):
    """Le connecteur reproduit les propriétés COM de la macro ; win32com est simulé."""
    from tests.conftest import ENVOYER_MAIL_REEL as reel

    pj = tmp_path / "a.txt"
    pj.write_text("x")
    journal = []

    class Mail:
        def __init__(self):
            self.Attachments = types.SimpleNamespace(Add=lambda chemin: journal.append(("pj", chemin)))

        def __setattr__(self, nom, valeur):
            journal.append((nom, valeur))
            object.__setattr__(self, nom, valeur)

        def Send(self):
            journal.append("Send")

        def Display(self):
            journal.append("Display")

    mail = Mail()
    faux_client = types.SimpleNamespace(Dispatch=lambda nom: types.SimpleNamespace(CreateItem=lambda n: mail))
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=faux_client))
    monkeypatch.setitem(sys.modules, "win32com.client", faux_client)

    reel("a@b.fr", "c@d.fr", "Objet", "Corps", [str(pj)], "envoyer")
    assert journal[-6:] == [("To", "a@b.fr"), ("CC", "c@d.fr"), ("Subject", "Objet"), ("Body", "Corps"),
                            ("pj", str(pj.resolve())), "Send"] or journal[:6] == [
        ("To", "a@b.fr"), ("CC", "c@d.fr"), ("Subject", "Objet"), ("Body", "Corps"), ("pj", str(pj.resolve())), "Send"]
    journal.clear()
    reel("a@b.fr", "", "Objet", "Corps", [], "afficher")
    assert journal[-1] == "Display" and "Send" not in journal
    with pytest.raises(ValueError, match="Mode de mail inconnu"):
        reel("a@b.fr", "", "o", "c", [], "aucun")
