"""Étapes exécutées par Excel lui-même (filtre de tableau, tableau croisé, graphique croisé, AutoFit).

openpyxl ne sait pas appliquer réellement un filtre de tableau, créer un tableau croisé dynamique,
un graphique basé dessus, ni mesurer les largeurs de colonnes : ces instructions VBA sont donc
reprises une à une via `bibliotheque.excel_com`.

Prérequis : Windows, Excel installé, pywin32, fichier fermé dans Excel.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from bibliotheque.excel_com import session_excel, xl


def autofit_donnees_excel(wb, sheet_data: str) -> None:
    """VBA EnsureStructure : `wsD.Columns("A:E").AutoFit`."""
    ws_d = wb.Worksheets(sheet_data)
    ws_d.Columns("A:E").AutoFit()


def appliquer_filtre_excel(wb, sheet_data: str, table_name: str, saisie: str | None) -> None:
    """VBA AppliquerFiltre : retire le filtre existant puis filtre la colonne 2 sur `saisie`.

    `saisie` vaut None quand l'utilisateur n'a rien saisi (Annuler) ou a saisi une région inconnue.
    """
    lo = wb.Worksheets(sheet_data).ListObjects(table_name)
    if lo.AutoFilter.FilterMode:
        lo.AutoFilter.ShowAllData()
    if saisie is None:
        return
    lo.Range.AutoFilter(Field=2, Criteria1=saisie)


def creer_visuel_excel(wb, sheet_visuel: str, sheet_data: str, table_name: str,
                       maintenant: datetime) -> None:
    """VBA CreerVisuel : nettoyage, KPI, tableau croisé TCD_Volume et graphique en courbes."""
    ws_v = wb.Worksheets(sheet_visuel)
    ws_d = wb.Worksheets(sheet_data)
    lo = ws_d.ListObjects(table_name)  # noqa: F841 (Set lo = ... dans la macro)

    # Nettoyage
    while ws_v.ChartObjects().Count > 0:
        ws_v.ChartObjects(1).Delete()
    while ws_v.PivotTables().Count > 0:
        ws_v.PivotTables(1).TableRange2.Clear()
    ws_v.Cells.Clear()

    # Titre + KPI (calculés sur toute la table, pas sur le filtre)
    a1 = ws_v.Range("A1")
    a1.Value = "Suivi distribution gaz - mis a jour le " + maintenant.strftime("%d/%m/%Y %H:%M")
    a1.Font.Bold = True
    a1.Font.Size = 14
    ws_v.Range("A2").Value = "Volume total (GWh)"
    ws_v.Range("B2").Formula = f"=SUM({table_name}[Volume_GWh])"
    ws_v.Range("D2").Value = "Moyenne / jour (GWh)"
    ws_v.Range("E2").Formula = (
        f"=SUM({table_name}[Volume_GWh])/SUMPRODUCT(1/COUNTIF({table_name}[Date],{table_name}[Date]))"
    )
    ws_v.Range("G2").Value = "Region du pic"
    ws_v.Range("A2,D2,G2").Font.Bold = True
    ws_v.Range("B2,E2").NumberFormat = "#,##0.0"
    ws_v.Range("H2").Formula = (
        f"=INDEX({table_name}[Region],MATCH(MAX({table_name}[Volume_GWh]),{table_name}[Volume_GWh],0))"
    )

    # TCD : lignes = Date, colonnes = Region, valeurs = somme volume
    pc = wb.PivotCaches().Create(SourceType=xl.xlDatabase, SourceData=table_name)
    pt = pc.CreatePivotTable(TableDestination=ws_v.Range("A5"), TableName="TCD_Volume")
    pt.PivotFields("Date").Orientation = xl.xlRowField
    pt.PivotFields("Region").Orientation = xl.xlColumnField
    champ = pt.AddDataField(pt.PivotFields("Volume_GWh"), "Volume (GWh)", xl.xlSum)
    champ.NumberFormat = "#,##0.0"
    pt.PivotFields("Date").NumberFormat = "dd/mm"

    # Graphique à côté du TCD
    co = ws_v.ChartObjects().Add(Left=ws_v.Range("J5").Left, Top=ws_v.Range("J5").Top,
                                 Width=520, Height=300)
    co.Chart.SetSourceData(pt.TableRange1)
    co.Chart.ChartType = xl.xlLine
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Volumes distribues par region (GWh)"

    ws_v.Columns("A:H").AutoFit()
    ws_v.Activate()


def appliquer_etapes_excel(chemin: str | Path, *, sheet_visuel: str, sheet_data: str,
                           table_name: str, saisie_filtre: str | None,
                           structure_creee: bool, maintenant: datetime, excel=None) -> None:
    """Une seule session Excel sur la copie enregistrée : AutoFit initial, filtre, visuel."""
    with session_excel(chemin, excel=excel) as wb:
        if structure_creee:
            autofit_donnees_excel(wb, sheet_data)
        appliquer_filtre_excel(wb, sheet_data, table_name, saisie_filtre)
        creer_visuel_excel(wb, sheet_visuel, sheet_data, table_name, maintenant)
