"""Étapes exécutées par Excel lui-même (ce qu'openpyxl ne sait pas faire).

Filtre de tableau réellement appliqué, tableau croisé dynamique, graphique du TCD, AutoFit,
SaveCopyAs et journal écrit dans le classeur ouvert. Chaque fonction d'étape reprend les
instructions VBA dans le même ordre.
"""

from __future__ import annotations

from datetime import datetime

from bibliotheque.excel_com import session_excel, xl


def executer_dans_excel(chemin, fonction):
    """Ouvre `chemin` dans Excel (une seule session), appelle `fonction(wb)`, enregistre, ferme."""
    with session_excel(chemin) as wb:
        return fonction(wb)


def format_excel(wb, code: str) -> str:
    """Code de format anglais (VBA : "dd/mm/yyyy", "#,##0.0") traduit dans la langue d'Excel.

    Appelé depuis Python, `NumberFormat` est interprété avec la langue d'Excel (en français,
    "dd" est refusé et "yyyy" est pris pour du texte) : on traduit jour, année et séparateurs
    d'après `Application.International`. Sans valeur lisible (tests), le code est inchangé.
    """
    try:
        # Application.International renvoie le tableau complet (xlXxx = rang, base 1)
        international = wb.Application.International
        jour, annee = international[21 - 1], international[19 - 1]
        milliers, decimal = international[4 - 1], international[3 - 1]
    except Exception:  # langue indéterminée (tests, faux objet) : on garde le code anglais
        return code
    if not all(isinstance(v, str) and v for v in (jour, annee, milliers, decimal)):
        return code
    if any(c in code for c in "dmyhs"):  # format de date : seules les lettres changent
        return code.replace("d", jour).replace("y", annee)
    return code.replace(",", "\0").replace(".", decimal).replace("\0", milliers)


def adresse_excel(wb, adresse: str) -> str:
    """Adresse multiple VBA ("A2,D2,G2") avec le séparateur de liste de la langue d'Excel.

    Depuis Python, Excel lit l'union de cellules avec son séparateur local (";" en français).
    Sans valeur lisible (tests, faux objet), l'adresse est inchangée.
    """
    try:
        separateur = wb.Application.International[5 - 1]  # xlListSeparator
    except Exception:
        return adresse
    return adresse.replace(",", separateur) if isinstance(separateur, str) and separateur else adresse


def appliquer_filtre_excel(wb, sheet_data: str, table_name: str, region: str) -> None:
    """VBA AppliquerFiltre (parties Excel). `region` vide = aucun filtre."""
    lo = wb.Worksheets(sheet_data).ListObjects(table_name)
    if lo.AutoFilter.FilterMode:
        lo.AutoFilter.ShowAllData()
    if region == "":
        return
    lo.Range.AutoFilter(Field=2, Criteria1=region)


def creer_visuel_excel(wb, sheet_visuel: str, sheet_data: str, table_name: str,
                       maintenant: datetime) -> None:
    """VBA CreerVisuel : KPI, tableau croisé TCD_Volume et graphique."""
    wsV = wb.Worksheets(sheet_visuel)

    # Nettoyage
    while wsV.ChartObjects().Count > 0:
        wsV.ChartObjects(1).Delete()
    while wsV.PivotTables().Count > 0:
        wsV.PivotTables(1).TableRange2.Clear()
    wsV.Cells.Clear()

    # Titre + KPI
    a1 = wsV.Range("A1")
    a1.Value = "Suivi distribution gaz - mis a jour le " + maintenant.strftime("%d/%m/%Y %H:%M")
    a1.Font.Bold = True
    a1.Font.Size = 14
    wsV.Range("A2").Value = "Volume total (GWh)"
    wsV.Range("B2").Formula = f"=SUM({table_name}[Volume_GWh])"
    wsV.Range("D2").Value = "Moyenne / jour (GWh)"
    wsV.Range("E2").Formula = (f"=SUM({table_name}[Volume_GWh])/"
                               f"SUMPRODUCT(1/COUNTIF({table_name}[Date],{table_name}[Date]))")
    wsV.Range("G2").Value = "Region du pic"
    wsV.Range(adresse_excel(wb, "A2,D2,G2")).Font.Bold = True
    wsV.Range(adresse_excel(wb, "B2,E2")).NumberFormat = format_excel(wb, "#,##0.0")
    wsV.Range("H2").Formula = (f"=INDEX({table_name}[Region],"
                               f"MATCH(MAX({table_name}[Volume_GWh]),{table_name}[Volume_GWh],0))")

    # TCD : lignes = Date, colonnes = Region, valeurs = somme volume
    pc = wb.PivotCaches().Create(SourceType=xl.xlDatabase, SourceData=table_name)
    pt = pc.CreatePivotTable(TableDestination=wsV.Range("A5"), TableName="TCD_Volume")
    pt.PivotFields("Date").Orientation = xl.xlRowField
    pt.PivotFields("Region").Orientation = xl.xlColumnField
    champ = pt.AddDataField(pt.PivotFields("Volume_GWh"), "Volume (GWh)", xl.xlSum)
    champ.NumberFormat = format_excel(wb, "#,##0.0")
    pt.PivotFields("Date").NumberFormat = format_excel(wb, "dd/mm")

    # Graphique à côté du TCD
    co = wsV.ChartObjects().Add(Left=wsV.Range("J5").Left, Top=wsV.Range("J5").Top,
                                Width=520, Height=300)
    graphique = co.Chart
    graphique.SetSourceData(pt.TableRange1)
    graphique.ChartType = xl.xlLine
    graphique.HasTitle = True
    graphique.ChartTitle.Text = "Volumes distribues par region (GWh)"

    wsV.Columns("A:H").AutoFit()
    wsV.Activate()


def sauvegarder_copie_excel(wb, chemin: str) -> None:
    """VBA SauvegarderVersion : ThisWorkbook.SaveCopyAs chemin (chemin absolu)."""
    wb.SaveCopyAs(chemin)


def ecrire_log_excel(wb, sheet_log: str, niveau: str, msg: str, horodatage: datetime) -> None:
    """VBA LogMsg, sur le classeur ouvert dans Excel (la feuille de log existe déjà)."""
    ws = wb.Worksheets(sheet_log)
    if ws.Range("A1").Value in (None, ""):
        ws.Range("A1:C1").Value = ("Horodatage", "Niveau", "Message")
    n = int(ws.Cells(ws.Rows.Count, 1).End(xl.xlUp).Row) + 1
    ws.Cells(n, 1).Value = horodatage
    ws.Cells(n, 1).NumberFormat = format_excel(wb, "dd/mm/yyyy hh:mm:ss")
    ws.Cells(n, 2).Value = niveau
    ws.Cells(n, 3).Value = msg
