"""Module bibliotheque/excel_com.py, testé avec de faux objets COM (Excel n'est pas disponible ici)."""

import sys

import pytest

from vba2py.bibliotheque import excel_com
from vba2py.bibliotheque.excel_com import Graphique, creer_tableau_et_graphique_croises, preparer_feuille


class Objet:
    """Objet COM factice : enregistre les affectations et les appels."""

    def __init__(self, journal, nom, **attrs):
        object.__setattr__(self, "_journal", journal)
        object.__setattr__(self, "_nom", nom)
        for k, v in attrs.items():
            object.__setattr__(self, k, v)

    def __setattr__(self, cle, valeur):
        self._journal.append((f"{self._nom}.{cle}", valeur))
        object.__setattr__(self, cle, valeur)


class Collection:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)


class FauxExcel:
    def __init__(self, tableaux_existants=(), graphiques_existants=()):
        self.journal = []
        j = self.journal
        self.fermetures = []
        tableau_range = Objet(j, "TableRange2", Address="$A$5:$G$22", Clear=lambda: j.append(("TableRange2.Clear()", None)))
        self.pt = Objet(j, "pt", Name="TCD_Volume", TableRange1="<TableRange1>", TableRange2=tableau_range,
                        PivotFields=lambda nom: self.champs.setdefault(nom, Objet(j, f"champ[{nom}]")),
                        AddDataField=self._add_data_field)
        self.champs = {}
        cache = Objet(j, "cache", CreatePivotTable=lambda dest, nom: j.append(("CreatePivotTable", (dest._nom, nom))) or self.pt)
        self.graphiques = list(graphiques_existants)
        self.tableaux = list(tableaux_existants)
        self.chart = Objet(j, "chart", SetSourceData=lambda src: j.append(("SetSourceData", src)),
                           ChartTitle=Objet(j, "ChartTitle"))

        def chart_objects(i=None):
            if i is not None:
                return self.graphiques[i - 1]
            coll = Collection(self.graphiques)
            coll.Add = lambda *a: j.append(("ChartObjects.Add", a)) or Objet(j, "co", Chart=self.chart)
            return coll

        ws = Objet(j, "ws",
                   Range=lambda a: Objet(j, f"Range[{a}]", Left=100.0, Top=50.0),
                   ChartObjects=chart_objects,
                   PivotTables=lambda i=None: self.tableaux[i - 1] if i else Collection(self.tableaux),
                   Columns=lambda c: Objet(j, f"Columns[{c}]", AutoFit=lambda: j.append(("AutoFit", c))))
        self.wb = Objet(j, "wb", Worksheets=lambda nom: j.append(("Worksheets", nom)) or ws,
                        PivotCaches=lambda: Objet(j, "caches", Create=lambda t, s: j.append(("PivotCaches.Create", (t, s))) or cache),
                        Save=lambda: j.append(("Save", None)), Close=lambda s: self.fermetures.append(("Close", s)))
        self.Workbooks = Objet(j, "Workbooks", Open=lambda p: j.append(("Open", p)) or self.wb)
        self.Quit = lambda: self.fermetures.append(("Quit", None))

    def _add_data_field(self, champ, libelle, fonction):
        self.journal.append(("AddDataField", (champ._nom, libelle, fonction)))
        return Objet(self.journal, "champ_donnees")

    def appels(self, nom):
        return [v for k, v in self.journal if k == nom]


PARAMS = dict(feuille="Feuille1", nom="TCD_Volume", table_source="tblDonnees", ancre="A5",
              ligne="Date", colonne="Region", valeur="Volume_GWh", libelle_valeur="Volume (GWh)",
              format_valeurs="#,##0.0", format_lignes="dd/mm",
              graphique=Graphique(ancre="J5", titre="Volumes distribues par region (GWh)"), autofit="A:H")


@pytest.fixture
def classeur(tmp_path):
    chemin = tmp_path / "copie.xlsm"
    chemin.write_bytes(b"")
    return chemin


def test_memes_appels_que_la_macro(classeur):
    excel = FauxExcel()
    plage = creer_tableau_et_graphique_croises(classeur, excel=excel, **PARAMS)
    assert plage == "$A$5:$G$22"
    assert excel.appels("Open") == [str(classeur.resolve())]
    assert excel.appels("PivotCaches.Create") == [(1, "tblDonnees")]  # xlDatabase
    assert excel.appels("CreatePivotTable") == [("Range[A5]", "TCD_Volume")]
    assert excel.appels("champ[Date].Orientation") == [1]  # xlRowField
    assert excel.appels("champ[Region].Orientation") == [2]  # xlColumnField
    assert excel.appels("AddDataField") == [("champ[Volume_GWh]", "Volume (GWh)", -4157)]  # xlSum
    assert excel.appels("champ_donnees.NumberFormat") == ["#,##0.0"]
    assert excel.appels("champ[Date].NumberFormat") == ["dd/mm"]
    assert excel.appels("ChartObjects.Add") == [(100.0, 50.0, 520, 300)]  # Left/Top de J5
    assert excel.appels("SetSourceData") == ["<TableRange1>"]
    assert excel.appels("chart.ChartType") == [4]  # xlLine
    assert excel.appels("ChartTitle.Text") == ["Volumes distribues par region (GWh)"]
    assert excel.appels("AutoFit") == ["A:H"]
    assert excel.appels("Save") == [None]
    assert excel.fermetures == [("Close", False)]  # instance fournie : pas de Quit


def test_instance_creee_puis_fermee(classeur, monkeypatch):
    excel = FauxExcel()
    monkeypatch.setattr(excel_com, "ouvrir_excel", lambda: excel)
    creer_tableau_et_graphique_croises(classeur, **PARAMS)
    assert excel.fermetures == [("Close", False), ("Quit", None)]


def test_fermeture_meme_en_cas_d_erreur(classeur, monkeypatch):
    excel = FauxExcel()
    monkeypatch.setattr(excel_com, "ouvrir_excel", lambda: excel)
    excel.wb.PivotCaches = lambda: (_ for _ in ()).throw(RuntimeError("source introuvable"))
    with pytest.raises(RuntimeError):
        creer_tableau_et_graphique_croises(classeur, **PARAMS)
    assert excel.fermetures == [("Close", False), ("Quit", None)]


def test_suppression_de_l_existant(classeur):
    j = []
    lie = Objet(j, "lie", Chart=Objet(j, "c1", PivotLayout=Objet(j, "pl", PivotTable=Objet(j, "p", Name="TCD_Volume"))),
                Delete=lambda: j.append(("Delete", "lie")))
    autre = Objet(j, "autre", Chart=Objet(j, "c2"), Delete=lambda: j.append(("Delete", "autre")))
    excel = FauxExcel(graphiques_existants=[lie, autre])
    excel.tableaux.append(excel.pt)
    creer_tableau_et_graphique_croises(classeur, excel=excel, **PARAMS)
    assert ("Delete", "lie") in j and ("Delete", "autre") not in j
    assert excel.appels("TableRange2.Clear()") == [None]


def test_chemin_relatif_rendu_absolu(classeur, monkeypatch):
    monkeypatch.chdir(classeur.parent)
    excel = FauxExcel()
    creer_tableau_et_graphique_croises("copie.xlsm", excel=excel, **PARAMS)
    assert excel.appels("Open") == [str(classeur.resolve())]


def test_erreurs(tmp_path, classeur):
    with pytest.raises(FileNotFoundError):
        creer_tableau_et_graphique_croises(tmp_path / "absent.xlsm", excel=FauxExcel(), **PARAMS)
    with pytest.raises(ValueError):
        creer_tableau_et_graphique_croises(classeur, excel=FauxExcel(), **dict(PARAMS, fonction="stdDev"))


def test_sans_colonne_ni_graphique(classeur):
    excel = FauxExcel()
    params = dict(PARAMS, colonne=None, graphique=None, autofit=None, format_valeurs=None, format_lignes=None)
    creer_tableau_et_graphique_croises(classeur, excel=excel, **params)
    assert "champ[Region].Orientation" not in [k for k, _ in excel.journal]
    assert excel.appels("ChartObjects.Add") == []


def test_preparer_feuille():
    from openpyxl import Workbook
    from openpyxl.chart import LineChart

    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Titre"
    ws.add_chart(LineChart(), "J5")
    ws._pivots.append(object())
    preparer_feuille(ws)
    assert ws["A1"].value is None and ws._charts == [] and ws._pivots == []


@pytest.mark.skipif(sys.platform != "win32", reason="Excel réel : Windows uniquement")
def test_excel_reel(tmp_path):
    pytest.importorskip("win32com")
    import datetime as dt

    from openpyxl import Workbook, load_workbook
    from openpyxl.worksheet.table import Table

    wb = Workbook()
    wb.active.title = "Feuille1"
    d = wb.create_sheet("Feuille2")
    d.append(["Date", "Region", "Volume_GWh"])
    for jour in (1, 2):
        for r, v in (("Nord", 10.0), ("Sud", 20.0)):
            d.append([dt.datetime(2026, 10, jour), r, v * jour])
    d.add_table(Table(displayName="tblDonnees", ref=f"A1:C{d.max_row}"))
    wb.save(tmp_path / "reel.xlsx")
    creer_tableau_et_graphique_croises(tmp_path / "reel.xlsx", **PARAMS)
    ws = load_workbook(tmp_path / "reel.xlsx")["Feuille1"]
    assert [p.name for p in ws._pivots] == ["TCD_Volume"]
    assert ws._charts and ws._charts[0].pivotSource is not None
