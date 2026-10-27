"""Instructions que openpyxl ne sait pas faire, exécutées par Excel (tableau croisé, graphique du
tableau croisé, filtre réellement appliqué, ``AutoFit``). Une fonction par étape VBA, avec le
même modèle objet qu'en VBA, et une fonction d'entrée par pipeline (une seule session Excel).
"""

from __future__ import annotations

from bibliotheque import excel_com
from bibliotheque.excel_com import TraceurCOM, xl

TABLE_MACRO_INITIALE = "tblDonnees"
FEUILLE_DONNEES = "Feuille2"
FEUILLE_VISUEL = "Feuille1"
LCID_ANGLAIS_US = 0x0409


def affecter_format_nombre(objet, code: str) -> None:
    """VBA ``objet.NumberFormat = code``.

    VBA transmet toujours les codes de format en anglais (``dd/mm/yyyy``, ``#,##0.0``). Un appel COM
    depuis Python utilise par défaut la langue de Windows : sur un Excel français, Excel attend alors
    ``jj/mm/aaaa`` et refuse ``dd/mm`` (« Impossible de définir la propriété NumberFormat »), ou lit
    ``#,##0.0`` avec la virgule comme séparateur décimal. L'affectation est donc envoyée avec le LCID
    anglais, comme VBA. Avec un faux objet (tests), c'est une simple affectation.
    """
    if not isinstance(objet, TraceurCOM):
        objet.NumberFormat = code
        return
    com = object.__getattribute__(objet, "_objet")
    chemin = object.__getattribute__(objet, "_chemin")
    table = getattr(com, "_prop_map_put_", {})

    def action():
        if "NumberFormat" not in table:  # liaison tardive : pas de LCID à remplacer
            com.NumberFormat = code
            return
        args, args_par_defaut = table["NumberFormat"]
        com._oleobj_.Invoke(args[0], LCID_ANGLAIS_US, *args[2:], code, *args_par_defaut)

    objet._executer(f"{chemin}.NumberFormat = {code!r}", action)


# ----------------------------------------------------------------------------- étapes

def ajuster_colonnes(wb, nom_feuille: str, colonnes: str, largeur_colonne_a: float | None = None) -> None:
    """VBA ``ws.Columns("A:H").AutoFit`` (puis ``ws.Columns("A").ColumnWidth = n`` si demandé)."""
    ws = wb.Worksheets(nom_feuille)
    ws.Columns(colonnes).AutoFit()
    if largeur_colonne_a is not None:
        ws.Columns("A").ColumnWidth = largeur_colonne_a


def activer_feuille(wb, nom_feuille: str) -> None:
    """VBA ``ws.Activate``"""
    wb.Worksheets(nom_feuille).Activate()


def appliquer_filtre_excel(wb, nom_feuille: str, nom_table: str, region: str | None) -> None:
    """VBA ``AppliquerFiltre`` / ``AppliquerFiltreV2`` : ``ShowAllData`` si un filtre est actif,
    puis ``lo.Range.AutoFilter Field:=2, Criteria1:=saisie`` si une région valide a été saisie."""
    lo = wb.Worksheets(nom_feuille).ListObjects(nom_table)
    if lo.AutoFilter.FilterMode:
        lo.AutoFilter.ShowAllData()
    if region:
        lo.Range.AutoFilter(Field=2, Criteria1=region)


def creer_visuel_excel(wb) -> None:
    """VBA ``CreerVisuel`` (partie Excel) : tableau croisé TCD_Volume, graphique en courbes,
    ajustement des colonnes A:H, activation de la feuille."""
    ws_v = wb.Worksheets(FEUILLE_VISUEL)
    pc = wb.PivotCaches().Create(SourceType=xl.xlDatabase, SourceData=TABLE_MACRO_INITIALE)
    pt = pc.CreatePivotTable(TableDestination=ws_v.Range("A5"), TableName="TCD_Volume")
    pt.PivotFields("Date").Orientation = xl.xlRowField
    pt.PivotFields("Region").Orientation = xl.xlColumnField
    champ = pt.AddDataField(pt.PivotFields("Volume_GWh"), "Volume (GWh)", xl.xlSum)
    affecter_format_nombre(champ, "#,##0.0")
    affecter_format_nombre(pt.PivotFields("Date"), "dd/mm")

    co = ws_v.ChartObjects().Add(Left=ws_v.Range("J5").Left, Top=ws_v.Range("J5").Top, Width=520, Height=300)
    co.Chart.SetSourceData(pt.TableRange1)
    co.Chart.ChartType = xl.xlLine
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Volumes distribues par region (GWh)"

    ws_v.Columns("A:H").AutoFit()
    ws_v.Activate()


# ----------------------------------------------------------------------------- entrées (une session)

def _session(session):
    return session or excel_com.session_excel


def finaliser_initiale(chemin, ajuster_donnees: bool, region: str | None = None, *, session=None) -> None:
    """MacroInitiale : AutoFit des données (si la structure vient d'être créée), filtre, visuel."""
    with _session(session)(chemin) as wb:
        if ajuster_donnees:
            ajuster_colonnes(wb, FEUILLE_DONNEES, "A:E")
        appliquer_filtre_excel(wb, FEUILLE_DONNEES, TABLE_MACRO_INITIALE, region)
        creer_visuel_excel(wb)


def finaliser_carte(chemin, ajuster_donnees: bool, region: str | None = None, *, session=None) -> None:
    """MacroCarte : AutoFit des données, filtre, puis mise en forme de la synthèse (Feuille1)."""
    with _session(session)(chemin) as wb:
        if ajuster_donnees:
            ajuster_colonnes(wb, FEUILLE_DONNEES, "A:E")
        appliquer_filtre_excel(wb, FEUILLE_DONNEES, TABLE_MACRO_INITIALE, region)
        ajuster_colonnes(wb, FEUILLE_VISUEL, "A:H", largeur_colonne_a=24)
        activer_feuille(wb, FEUILLE_VISUEL)


def finaliser_csv_v3(chemin, *, session=None) -> None:
    """UtilisationCSV : AutoFit de Donnees_V3 (A:F) puis de la synthèse Carte_V3 (A:H, A = 24)."""
    with _session(session)(chemin) as wb:
        ajuster_colonnes(wb, "Donnees_V3", "A:F")
        ajuster_colonnes(wb, "Carte_V3", "A:H", largeur_colonne_a=24)
        activer_feuille(wb, "Carte_V3")


def ajuster_feuille(chemin, nom_feuille: str, colonnes: str, activer: bool = False, *, session=None) -> None:
    """SimuMainframe / TestInterfaceMVS : AutoFit d'une feuille (et ``Activate`` si demandé)."""
    with _session(session)(chemin) as wb:
        ajuster_colonnes(wb, nom_feuille, colonnes)
        if activer:
            activer_feuille(wb, nom_feuille)
