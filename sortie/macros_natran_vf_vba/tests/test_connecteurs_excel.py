"""Étapes Excel (connecteurs/excel.py) : chaque instruction du bloc VBA a son équivalent, dans l'ordre."""

from contextlib import contextmanager
from pathlib import Path

import pytest

from bibliotheque.excel_com import EnregistreurCOM, xl
from connecteurs import excel as ce


def chemins(wb):
    """Chemins des appels, dans l'ordre (les affectations sont repérées par « = »)."""
    sorties = []
    for t, c, v in wb.journal:
        sorties.append(c if t == "appel" else f"{c} = {v!r}")
    return sorties


def position(wb, fragment, apres=-1):
    for i, ligne in enumerate(chemins(wb)):
        if i > apres and fragment in ligne:
            return i
    raise AssertionError(f"{fragment!r} absent après la position {apres} : {chemins(wb)}")


def test_creer_visuel_excel_reprend_chaque_instruction_dans_l_ordre():
    wb = EnregistreurCOM("wb")
    ce.creer_visuel_excel(wb)
    args, kw = wb.appels("PivotCaches().Create")[0]
    assert (args, kw) == ((), {"SourceType": xl.xlDatabase, "SourceData": "tblDonnees"})
    assert wb.appels("CreatePivotTable")[0][1]["TableName"] == "TCD_Volume"
    assert wb.affectations("Orientation") == [xl.xlRowField, xl.xlColumnField]
    add = wb.appels("AddDataField")[0][0]
    assert add[1:] == ("Volume (GWh)", xl.xlSum)
    assert wb.affectations("NumberFormat") == ["#,##0.0", "dd/mm"]
    assert wb.appels("ChartObjects().Add")[0][1]["Width"] == 520 and wb.appels("ChartObjects().Add")[0][1]["Height"] == 300
    assert wb.affectations("ChartType") == [xl.xlLine] and wb.affectations("HasTitle") == [True]
    assert wb.affectations("ChartTitle.Text") == ["Volumes distribues par region (GWh)"]
    assert wb.appels("Columns")[-1][0] == ("A:H",)
    ordre = ["PivotCaches().Create", "CreatePivotTable", ".Orientation", "AddDataField", "ChartObjects().Add",
             "SetSourceData", ".ChartType", ".HasTitle", "ChartTitle.Text", "AutoFit", "Activate"]
    pos = -1
    for fragment in ordre:
        pos = position(wb, fragment, pos)


class _FauxOle:
    def __init__(self):
        self.appels = []

    def Invoke(self, *args):
        self.appels.append(args)


class _FauxObjetCom:
    """Objet COM à liaison précoce : _prop_map_put_ = (dispid, LCID, DISPATCH_PROPERTYPUT, 0)."""
    _prop_map_put_ = {"NumberFormat": ((193, 0, 4, 0), ())}

    def __init__(self):
        self._oleobj_ = _FauxOle()


def test_format_de_nombre_est_envoye_avec_le_lcid_anglais():
    """VBA envoie « dd/mm » en anglais ; Python (LCID par défaut) ferait attendre « jj/mm » à un Excel français."""
    from bibliotheque.excel_com import TraceurCOM
    com = _FauxObjetCom()
    ce.affecter_format_nombre(TraceurCOM(com, "o13"), "dd/mm")
    assert com._oleobj_.appels == [(193, 0x0409, 4, 0, "dd/mm")]


def test_format_de_nombre_sur_faux_objet_est_une_simple_affectation():
    wb = EnregistreurCOM("wb")
    ce.affecter_format_nombre(wb.Worksheets("F").Range("A1"), "#,##0.0")
    assert wb.affectations("NumberFormat") == ["#,##0.0"]


def test_appliquer_filtre_excel_avec_filtre_actif():
    wb = EnregistreurCOM("wb", reponses={
        "wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode": True})
    ce.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", "Nord")
    assert position(wb, "ShowAllData") < position(wb, "Range.AutoFilter")       # ShowAllData puis AutoFilter
    assert wb.appels("Range.AutoFilter")[0][1] == {"Field": 2, "Criteria1": "Nord"}


def test_appliquer_filtre_excel_sans_saisie():
    wb = EnregistreurCOM("wb", reponses={
        "wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode": False})
    ce.appliquer_filtre_excel(wb, "Feuille2", "tblDonnees", None)
    assert wb.appels("ShowAllData") == [] and wb.appels("Range.AutoFilter") == []


def test_ajuster_colonnes_et_largeur():
    wb = EnregistreurCOM("wb")
    ce.ajuster_colonnes(wb, "Carte_V3", "A:H", largeur_colonne_a=24)
    assert wb.appels("AutoFit") == [((), {})] and wb.affectations("ColumnWidth") == [24]
    assert position(wb, "AutoFit") < position(wb, "ColumnWidth")


class Session:
    """Fausse fonction d'entrée : enregistre le chemin reçu, fournit un EnregistreurCOM."""

    def __init__(self, reponses=None):
        self.wb = EnregistreurCOM("wb", reponses=reponses)
        self.chemins = []

    @contextmanager
    def __call__(self, chemin):
        self.chemins.append(chemin)
        yield self.wb


def test_finaliser_initiale_ordre_des_etapes(tmp_path):
    s = Session()
    ce.finaliser_initiale(tmp_path / "c.xlsm", True, "Sud", session=s)
    assert s.chemins == [tmp_path / "c.xlsm"]
    wb = s.wb
    a = position(wb, "AutoFit")                       # 1. AutoFit des données (structure créée)
    b = position(wb, "Range.AutoFilter", a)           # 2. filtre
    c = position(wb, "PivotCaches().Create", b)       # 3. tableau croisé
    assert a < b < c
    assert "Feuille2" in chemins(wb)[a]


def test_finaliser_initiale_sans_structure_creee_ni_filtre(tmp_path):
    s = Session(reponses={"wb.Worksheets('Feuille2').ListObjects('tblDonnees').AutoFilter.FilterMode": False})
    ce.finaliser_initiale(tmp_path / "c.xlsm", False, None, session=s)
    assert "Feuille2').Columns" not in " ".join(chemins(s.wb))   # pas d'AutoFit des données
    assert s.wb.appels("PivotCaches().Create")


def test_finaliser_carte_et_csv_v3(tmp_path):
    s = Session()
    ce.finaliser_carte(tmp_path / "c.xlsm", True, "Est", session=s)
    assert s.wb.affectations("ColumnWidth") == [24] and position(s.wb, "Activate") > position(s.wb, "ColumnWidth")
    s2 = Session()
    ce.finaliser_csv_v3(tmp_path / "c.xlsm", session=s2)
    assert [a for a in chemins(s2.wb) if "Worksheets" in a and a.endswith("Columns")] and \
        s2.wb.affectations("ColumnWidth") == [24]


def test_ajuster_feuille(tmp_path):
    s = Session()
    ce.ajuster_feuille(tmp_path / "c.xlsx", "Extraction_MVS", "A:F", True, session=s)
    assert s.wb.appels("AutoFit") and s.wb.appels("Activate")


def test_entree_par_defaut_utilise_la_session_excel_de_la_bibliotheque(tmp_path, monkeypatch):
    """Sans fonction injectée, l'entrée passe par excel_com.session_excel (ici remplacée)."""
    from bibliotheque import excel_com
    s = Session()
    monkeypatch.setattr(excel_com, "session_excel", s)
    ce.finaliser_csv_v3.__wrapped__(tmp_path / "x.xlsx")     # fonction d'origine (la fixture autouse la bloque)
    assert s.chemins == [tmp_path / "x.xlsx"]
