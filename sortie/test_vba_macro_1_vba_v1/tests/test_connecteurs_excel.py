from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from bibliotheque.excel_com import EnregistreurCOM, xl
from connecteurs import excel as ce

MAINTENANT = datetime(2026, 10, 25, 11, 15)
# La fixture autouse de conftest remplace l'entrée par une fonction qui lève une erreur ;
# on garde ici la vraie fonction (elle ne reçoit qu'un faux Excel dans ces tests).
appliquer_etapes_excel = ce.appliquer_etapes_excel


def test_autofit_donnees():
    wb = EnregistreurCOM("wb")
    ce.autofit_donnees_excel(wb, "Feuille2")
    assert wb.instructions() == ["wb.Worksheets('Feuille2')", "wb.Worksheets('Feuille2').Columns('A:E')",
                                 "wb.Worksheets('Feuille2').Columns('A:E').AutoFit()"]


def test_filtre_avec_saisie():
    wb = EnregistreurCOM("wb")
    ce.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "Nord")
    ins = wb.instructions()
    # FilterMode (objet enregistreur, donc vrai) -> ShowAllData puis AutoFilter sur la colonne 2
    assert any(i.endswith("AutoFilter.ShowAllData()") for i in ins)
    assert wb.appels("Range.AutoFilter") == [((), {"Field": 2, "Criteria1": "Nord"})]
    assert ins.index(next(i for i in ins if "ShowAllData" in i)) < ins.index(next(i for i in ins if "Criteria1" in i))


def test_filtre_sans_saisie_retire_seulement_le_filtre():
    wb = EnregistreurCOM("wb")
    ce.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", None)
    assert wb.appels("Range.AutoFilter") == []
    assert len(wb.appels("AutoFilter.ShowAllData")) == 1


def test_filtre_aucun_filtre_actif():
    wb = EnregistreurCOM("wb", reponses={"wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode": False})
    ce.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", None)
    assert wb.appels("ShowAllData") == []


def test_creer_visuel_instructions_dans_l_ordre():
    wb = EnregistreurCOM("wb")
    ce.creer_visuel_excel(wb, "Feuille1", "Feuille2", "tblDonnees", MAINTENANT)
    v = "wb.Worksheets('Feuille1')"
    ins = wb.instructions()

    def pos(fragment):
        return next(i for i, x in enumerate(ins) if fragment in x)

    # titre et KPI
    assert wb.affectations("Range('A1').Value") == ["Suivi distribution gaz - mis a jour le 25/10/2026 11:15"]
    assert wb.affectations("Range('A1').Font.Bold") == [True]
    assert wb.affectations("Range('A1').Font.Size") == [14]
    assert wb.affectations("Range('A2').Value") == ["Volume total (GWh)"]
    assert wb.affectations("Range('B2').Formula") == ["=SUM(tblDonnees[Volume_GWh])"]
    assert wb.affectations("Range('D2').Value") == ["Moyenne / jour (GWh)"]
    assert wb.affectations("Range('E2').Formula") == [
        "=SUM(tblDonnees[Volume_GWh])/SUMPRODUCT(1/COUNTIF(tblDonnees[Date],tblDonnees[Date]))"]
    assert wb.affectations("Range('G2').Value") == ["Region du pic"]
    assert wb.affectations("Range('A2,D2,G2').Font.Bold") == [True]
    assert wb.affectations("Range('B2,E2').NumberFormat") == ["#,##0.0"]
    assert wb.affectations("Range('H2').Formula") == [
        "=INDEX(tblDonnees[Region],MATCH(MAX(tblDonnees[Volume_GWh]),tblDonnees[Volume_GWh],0))"]
    # nettoyage avant tout
    assert pos(f"{v}.Cells.Clear()") < pos("Range('A1').Value")
    # tableau croisé
    assert wb.appels("PivotCaches().Create") == [((), {"SourceType": xl.xlDatabase, "SourceData": "tblDonnees"})]
    (_, kw), = wb.appels("CreatePivotTable")
    assert kw["TableName"] == "TCD_Volume"
    assert wb.affectations("PivotFields('Date').Orientation") == [xl.xlRowField]
    assert wb.affectations("PivotFields('Region').Orientation") == [xl.xlColumnField]
    assert wb.appels("AddDataField") [0][0][1:] == ("Volume (GWh)", xl.xlSum)
    # B2,E2 puis le champ de données ; puis le champ Date en dd/mm
    assert wb.affectations("NumberFormat").count("#,##0.0") == 2 and "dd/mm" in wb.affectations("NumberFormat")
    # graphique
    (_, kw), = wb.appels("ChartObjects().Add")
    assert (kw["Width"], kw["Height"]) == (520, 300)
    assert wb.affectations("Chart.ChartType") == [xl.xlLine]
    assert wb.affectations("Chart.HasTitle") == [True]
    assert wb.affectations("ChartTitle.Text") == ["Volumes distribues par region (GWh)"]
    assert len(wb.appels("Chart.SetSourceData")) == 1
    # fin
    assert wb.appels("Columns('A:H').AutoFit") == [((), {})]
    assert len(wb.appels("Activate")) == 1
    assert pos("CreatePivotTable") < pos("ChartObjects().Add") < pos("AutoFit") < pos("Activate")


class _UneFois:
    """Count qui vaut 1 à la première comparaison, puis 0 (l'objet a été supprimé)."""

    def __init__(self):
        self.restant = 1

    def __gt__(self, autre):
        vrai = self.restant > autre
        self.restant = 0
        return vrai


def test_creer_visuel_supprime_anciens_objets():
    v = "wb.Worksheets('Feuille1')"
    wb = EnregistreurCOM("wb", reponses={f"{v}.ChartObjects().Count": _UneFois(),
                                         f"{v}.PivotTables().Count": _UneFois()})
    ce.creer_visuel_excel(wb, "Feuille1", "Feuille2", "tblDonnees", MAINTENANT)
    assert len(wb.appels("ChartObjects(1).Delete")) == 1
    assert len(wb.appels("PivotTables(1).TableRange2.Clear")) == 1


def test_entree_excel_une_session(tmp_path):
    fichier = tmp_path / "f.xlsx"
    fichier.write_bytes(b"x")
    faux_excel = EnregistreurCOM("excel")
    appliquer_etapes_excel(fichier, sheet_visuel="Feuille1", sheet_data="Feuille2", table_name="tblDonnees",
                              saisie_filtre="Est", structure_creee=True, maintenant=MAINTENANT, excel=faux_excel)
    ouvertures = faux_excel.appels("Workbooks.Open")
    assert ouvertures == [((str(fichier.resolve()),), {})]
    assert Path(ouvertures[0][0][0]).is_absolute()
    ins = faux_excel.instructions()
    autofit_a_e = next(i for i, x in enumerate(ins) if "Columns('A:E').AutoFit" in x)
    filtre = next(i for i, x in enumerate(ins) if "Criteria1='Est'" in x)
    visuel = next(i for i, x in enumerate(ins) if "PivotCaches" in x)
    enreg = next(i for i, x in enumerate(ins) if x.endswith(".Save()"))
    assert autofit_a_e < filtre < visuel < enreg


def test_entree_excel_sans_structure_pas_d_autofit_donnees(tmp_path):
    fichier = tmp_path / "f.xlsx"
    fichier.write_bytes(b"x")
    faux_excel = EnregistreurCOM("excel")
    appliquer_etapes_excel(fichier, sheet_visuel="Feuille1", sheet_data="Feuille2", table_name="t",
                              saisie_filtre=None, structure_creee=False, maintenant=MAINTENANT, excel=faux_excel)
    assert not any("Columns('A:E').AutoFit" in i for i in faux_excel.instructions())


def test_entree_excel_fichier_absent(tmp_path):
    with pytest.raises(FileNotFoundError):
        appliquer_etapes_excel(tmp_path / "absent.xlsx", sheet_visuel="a", sheet_data="b", table_name="t",
                                  saisie_filtre=None, structure_creee=False, maintenant=MAINTENANT,
                                  excel=EnregistreurCOM("excel"))
