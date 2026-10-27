from datetime import datetime

import pytest

from bibliotheque.excel_com import EnregistreurCOM, xl
from connecteurs import excel as cx
from connecteurs import outlook

MAINTENANT = datetime(2026, 10, 25, 9, 30)
ORIGINAL_ENVOYER = outlook.envoyer_mail
ORIGINAL_EXECUTER = cx.executer_dans_excel  # conftest les remplace pendant les tests


def test_filtre_avec_region():
    wb = EnregistreurCOM("wb")
    cx.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "Sud")
    instr = wb.instructions()
    assert instr[0] == "wb.Worksheets('Feuille2')"
    assert instr[1] == "wb.Worksheets('Feuille2').ListObjects('tblDonnees')"
    assert instr[2].endswith("AutoFilter.ShowAllData()")
    assert instr[3].endswith("Range.AutoFilter(Field=2, Criteria1='Sud')")
    assert len(instr) == 4


def test_filtre_vide_ou_sans_filtre_actif():
    wb = EnregistreurCOM("wb", reponses={
        "wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode": False})
    cx.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "")
    # pas de filtre actif, saisie vide : ni ShowAllData ni AutoFilter
    assert wb.appels("ShowAllData") == [] and wb.appels("AutoFilter") == []


def test_creer_visuel_excel_instructions():
    wb = EnregistreurCOM("wb")
    cx.creer_visuel_excel(wb, "Feuille1", "Feuille2", "tblDonnees", MAINTENANT)
    v = "wb.Worksheets('Feuille1')"
    instr = wb.instructions()
    # nettoyage (Count = 0 : aucune suppression), puis Cells.Clear
    assert f"{v}.Cells.Clear()" in instr
    assert instr.index(f"{v}.Cells.Clear()") < next(i for i, s in enumerate(instr) if "Formula" in s)
    assert wb.affectations("Range('A1').Value") == ["Suivi distribution gaz - mis a jour le 25/10/2026 09:30"]
    assert wb.affectations("Range('A1').Font.Bold") == [True]
    assert wb.affectations("Range('A1').Font.Size") == [14]
    assert wb.affectations("Range('B2').Formula") == ["=SUM(tblDonnees[Volume_GWh])"]
    assert wb.affectations("Range('E2').Formula") == [
        "=SUM(tblDonnees[Volume_GWh])/SUMPRODUCT(1/COUNTIF(tblDonnees[Date],tblDonnees[Date]))"]
    assert wb.affectations("Range('H2').Formula") == [
        "=INDEX(tblDonnees[Region],MATCH(MAX(tblDonnees[Volume_GWh]),tblDonnees[Volume_GWh],0))"]
    assert wb.affectations("Range('A2,D2,G2').Font.Bold") == [True]
    assert wb.affectations("Range('B2,E2').NumberFormat") == ["#,##0.0"]
    assert wb.appels("PivotCaches().Create") == [((), {"SourceType": xl.xlDatabase, "SourceData": "tblDonnees"})]
    assert wb.appels("CreatePivotTable")[0][1]["TableName"] == "TCD_Volume"
    assert wb.affectations("PivotFields('Date').Orientation") == [xl.xlRowField]
    assert wb.affectations("PivotFields('Region').Orientation") == [xl.xlColumnField]
    assert wb.appels("AddDataField")[0][0][1:] == ("Volume (GWh)", xl.xlSum)
    assert wb.affectations("NumberFormat") .count("#,##0.0") == 2  # B2,E2 et le champ de données
    assert wb.affectations("PivotFields('Date').NumberFormat") == ["dd/mm"]
    assert wb.appels("ChartObjects().Add")[0][1]["Width"] == 520
    assert wb.appels("ChartObjects().Add")[0][1]["Height"] == 300
    assert wb.affectations("ChartType") == [xl.xlLine]
    assert wb.affectations("HasTitle") == [True]
    assert wb.affectations("ChartTitle.Text") == ["Volumes distribues par region (GWh)"]
    assert wb.appels("Columns('A:H').AutoFit") == [((), {})]
    assert instr[-1] == f"{v}.Activate()"
    # ordre : KPI < TCD < graphique
    pos = lambda cle: next(i for i, s in enumerate(instr) if cle in s)  # noqa: E731
    assert pos("Formula") < pos("PivotCaches") < pos("ChartObjects().Add") < pos("AutoFit")


def test_creer_visuel_excel_nettoyage():
    wb = EnregistreurCOM("wb", reponses={"wb.Worksheets('Feuille1').ChartObjects().Count": 0})
    cx.creer_visuel_excel(wb, "Feuille1", "Feuille2", "tblDonnees", MAINTENANT)
    assert wb.appels("ChartObjects(1).Delete") == []


def test_sauvegarder_copie_et_log():
    wb = EnregistreurCOM("wb", reponses={
        "wb.Worksheets('Log').Rows.Count": 1048576,
        "wb.Worksheets('Log').Cells(1048576, 1).End(-4162).Row": 5})
    cx.sauvegarder_copie_excel(wb, "C:/x/y.xlsx")
    cx.ecrire_log_excel(wb, "Log", "INFO", "msg", MAINTENANT)
    assert wb.appels("SaveCopyAs") == [(("C:/x/y.xlsx",), {})]
    assert wb.affectations("Cells(6, 1).Value") == [MAINTENANT]
    assert wb.affectations("Cells(6, 2).Value") == ["INFO"]
    assert wb.affectations("Cells(6, 3).Value") == ["msg"]


def test_executer_dans_excel_utilise_session(monkeypatch, tmp_path):
    f = tmp_path / "c.xlsx"
    f.write_text("x")
    vus = []

    class Session:
        def __init__(self, chemin):
            vus.append(chemin)

        def __enter__(self):
            return "wb"

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(cx, "session_excel", Session)
    assert ORIGINAL_EXECUTER(f, lambda wb: wb + "!") == "wb!"
    assert vus == [f]


def test_outlook_mode_invalide():
    # la vraie fonction (conftest la remplace pendant les tests) valide le mode avant tout appel COM
    with pytest.raises(ValueError):
        ORIGINAL_ENVOYER("a@b.c", "", "s", "c", [], "aucun")
