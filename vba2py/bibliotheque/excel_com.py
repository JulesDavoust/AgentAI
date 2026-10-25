"""Tableau croisé dynamique et graphique croisé créés par Excel lui-même (pilotage COM).

Module fourni par vba2py et copié tel quel dans chaque projet traduit : ne pas le modifier.

Le reste de la traduction utilise openpyxl. Excel n'intervient qu'à la fin, sur le fichier
déjà enregistré, pour ce qu'openpyxl ne sait pas faire : créer un tableau croisé dynamique et
son graphique, avec exactement les mêmes appels que la macro. Le fichier contient ensuite des
valeurs calculées par Excel (visibles même en Mode protégé).

Prérequis : Windows, Excel installé, pywin32 ; le fichier ne doit pas être ouvert dans Excel.
Une instance d'Excel séparée et invisible est utilisée, sans exécuter les macros du classeur.

Correspondance VBA :
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, "tblDonnees")
    Set pt = pc.CreatePivotTable(wsV.Range("A5"), "TCD_Volume")
    .PivotFields("Date").Orientation = xlRowField ...
    Set co = wsV.ChartObjects.Add(...) : co.Chart.SetSourceData pt.TableRange1
→   1. openpyxl : preparer_feuille(ws)  (remplace le nettoyage ChartObjects/PivotTables/Cells.Clear)
    2. openpyxl : wb.save(chemin)
    3. creer_tableau_et_graphique_croises(chemin, feuille="Feuille1", nom="TCD_Volume", ...)
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

# Constantes Excel (valeurs de l'énumération VBA)
XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
FONCTIONS = {"sum": -4157, "count": -4112, "average": -4106, "min": -4139, "max": -4136}
TYPES_GRAPHIQUE = {"ligne": 4, "barres": 51}  # xlLine, xlColumnClustered
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


@dataclass
class Graphique:
    ancre: str
    titre: str | None = None
    type_graphique: str = "ligne"
    largeur_pt: float = 520
    hauteur_pt: float = 300


def preparer_feuille(ws) -> None:
    """Côté openpyxl, avant l'enregistrement : équivalent du nettoyage de la macro.

    `Do While ws.ChartObjects.Count > 0 ... Delete`, `PivotTables(1).TableRange2.Clear`,
    `ws.Cells.Clear` : tableaux croisés, graphiques, images et cellules sont retirés ; Excel
    recréera le tableau et le graphique ensuite.
    """
    ws._pivots = []
    ws._charts = []
    ws._images = []
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
            cell.style = "Normal"


def ouvrir_excel():
    """Instance d'Excel séparée et invisible (ne touche pas aux classeurs déjà ouverts)."""
    import win32com.client  # import paresseux : le reste du projet s'importe sans pywin32

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE  # pas de Workbook_Open
    return excel


def creer_tableau_et_graphique_croises(
    chemin_classeur,
    *,
    feuille: str,
    nom: str,
    table_source: str,
    ancre: str,
    ligne: str,
    valeur: str,
    colonne: str | None = None,
    fonction: str = "sum",
    libelle_valeur: str | None = None,
    format_valeurs: str | None = None,
    format_lignes: str | None = None,
    graphique: Graphique | None = None,
    autofit: str | None = None,
    excel=None,
) -> str:
    """Ouvre le classeur dans Excel, (re)crée le tableau croisé et son graphique, enregistre.

    Retourne la plage occupée par le tableau croisé (ex. « $A$5:$G$22 »).
    `excel` : instance déjà ouverte (sinon une instance est créée puis fermée).
    `autofit` : colonnes à ajuster après création (ex. "A:H", comme `Columns("A:H").AutoFit`).
    """
    if fonction not in FONCTIONS:
        raise ValueError(f"fonction d'agrégation inconnue : {fonction}")
    if graphique is not None and graphique.type_graphique not in TYPES_GRAPHIQUE:
        raise ValueError(f"type de graphique inconnu : {graphique.type_graphique}")
    chemin = str(Path(chemin_classeur).resolve())  # Excel ne connaît pas le dossier courant de Python
    if not Path(chemin).exists():
        raise FileNotFoundError(chemin)

    instance_creee = excel is None
    excel = excel or ouvrir_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(chemin)
        ws = wb.Worksheets(feuille)
        _supprimer_existant(ws, nom)

        cache = wb.PivotCaches().Create(XL_DATABASE, table_source)
        pt = cache.CreatePivotTable(ws.Range(ancre), nom)
        pt.PivotFields(ligne).Orientation = XL_ROW_FIELD
        if colonne:
            pt.PivotFields(colonne).Orientation = XL_COLUMN_FIELD
        champ = pt.AddDataField(pt.PivotFields(valeur), libelle_valeur or valeur, FONCTIONS[fonction])
        if format_valeurs:
            champ.NumberFormat = format_valeurs
        if format_lignes:
            pt.PivotFields(ligne).NumberFormat = format_lignes

        if graphique is not None:
            cible = ws.Range(graphique.ancre)
            objet = ws.ChartObjects().Add(cible.Left, cible.Top, graphique.largeur_pt, graphique.hauteur_pt)
            chart = objet.Chart
            chart.SetSourceData(pt.TableRange1)
            chart.ChartType = TYPES_GRAPHIQUE[graphique.type_graphique]
            if graphique.titre:
                chart.HasTitle = True
                chart.ChartTitle.Text = graphique.titre
        if autofit:
            ws.Columns(autofit).AutoFit()

        plage = pt.TableRange2.Address
        wb.Save()
        return plage
    finally:
        if wb is not None:
            wb.Close(False)
        if instance_creee:
            excel.Quit()


def _supprimer_existant(ws, nom: str) -> None:
    """Supprime le tableau croisé `nom` et les graphiques qui lui sont liés, s'ils existent."""
    for i in range(ws.ChartObjects().Count, 0, -1):
        objet = ws.ChartObjects(i)
        try:
            lie = objet.Chart.PivotLayout.PivotTable.Name == nom
        except Exception:  # graphique ordinaire : pas de PivotLayout
            lie = False
        if lie:
            objet.Delete()
    for i in range(ws.PivotTables().Count, 0, -1):
        pt = ws.PivotTables(i)
        if pt.Name == nom:
            pt.TableRange2.Clear()
