"""Diagnostic (facultatif) : rejoue l'étape « tableau croisé » de CreerVisuel dans Excel, variante par
variante, et affiche ce qu'Excel accepte ou refuse. À lancer à la main si Excel refuse une instruction :

    python tests/diagnostic_excel.py
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RACINE))

import openpyxl  # noqa: E402

import macro_initiale as mi  # noqa: E402
from bibliotheque.excel_com import ErreurExcel, session_excel, xl  # noqa: E402


def preparer(dossier: Path) -> Path:
    wb = openpyxl.load_workbook(RACINE / "tests" / "fixtures" / "classeur_donnees.xlsx")
    mi.creer_visuel_cellules(wb, datetime(2026, 10, 27, 14, 30))
    chemin = dossier / "diag.xlsx"
    wb.save(chemin)
    return chemin


def variante(nom, chemin, action):
    copie = chemin.with_name(f"diag_{abs(hash(nom))}.xlsx")
    shutil.copy(chemin, copie)
    try:
        with session_excel(copie) as wb:
            ws = wb.Worksheets("Feuille1")
            pc = wb.PivotCaches().Create(SourceType=xl.xlDatabase, SourceData="tblDonnees")
            pt = pc.CreatePivotTable(TableDestination=ws.Range("A5"), TableName="TCD_Volume")
            action(wb, pt)
        print(f"OK      {nom}")
    except ErreurExcel as e:
        print(f"REFUSE  {nom} : {str(e).splitlines()[1] if len(str(e).splitlines()) > 1 else e}")
    except Exception as e:  # noqa: BLE001
        print(f"ERREUR  {nom} : {e!r}")


def base(pt):
    pt.PivotFields("Date").Orientation = xl.xlRowField
    pt.PivotFields("Region").Orientation = xl.xlColumnField
    pt.AddDataField(pt.PivotFields("Volume_GWh"), "Volume (GWh)", xl.xlSum)


def a_originale(wb, pt):
    base(pt)
    pt.PivotFields("Date").NumberFormat = "dd/mm"


def b_avant_donnees(wb, pt):
    pt.PivotFields("Date").Orientation = xl.xlRowField
    pt.PivotFields("Date").NumberFormat = "dd/mm"


def c_infos(wb, pt):
    base(pt)
    champ = pt.PivotFields("Date")
    print("   DataType Date =", champ.DataType, "| Orientation =", champ.Orientation,
          "| 1er item =", champ.PivotItems(1).Name, "| Layout =", pt.RowAxisLayout if hasattr(pt, "RowAxisLayout") else "?")


def d_refresh(wb, pt):
    base(pt)
    pt.RefreshTable()
    pt.PivotFields("Date").NumberFormat = "dd/mm"


def e_format_long(wb, pt):
    base(pt)
    pt.PivotFields("Date").NumberFormat = "dd/mm/yyyy"


def f_range_row(wb, pt):
    base(pt)
    pt.RowRange.NumberFormat = "dd/mm"


def g_tabulaire(wb, pt):
    pt.RowAxisLayout(xl.xlTabularRow)
    base(pt)
    pt.PivotFields("Date").NumberFormat = "dd/mm"


def i_general(wb, pt):
    base(pt)
    champ = pt.PivotFields("Date")
    print("   NumberFormat lu =", repr(champ.NumberFormat))
    champ.NumberFormat = "General"


def formats_de_nombre():
    """Quels codes de format Excel accepte-t-il via COM ? (anglais « dd/mm/yyyy » ou français « jj/mm/aaaa »)"""
    from bibliotheque.excel_com import ouvrir_excel
    excel = ouvrir_excel()
    try:
        print("   Version =", excel.Version, "| séparateur décimal =", repr(excel.International[2]))
        wb = excel.Workbooks.Add()
        cellule = wb.Worksheets(1).Range("A1")
        for propriete, code in (("NumberFormat", "dd/mm/yyyy"), ("NumberFormat", "dd/mm"), ("NumberFormat", "0.00"),
                                ("NumberFormat", "#,##0.0"), ("NumberFormat", "jj/mm/aaaa"),
                                ("NumberFormatLocal", "jj/mm/aaaa"), ("NumberFormatLocal", "dd/mm/yyyy")):
            try:
                setattr(cellule, propriete, code)
                print(f"OK      liaison précoce {propriete} = {code!r}  ->  relu : {cellule.NumberFormat!r}")
            except Exception:  # noqa: BLE001
                print(f"REFUSE  liaison précoce {propriete} = {code!r}")
        wb.Close(False)
    finally:
        excel.Quit()


def classeur_natif():
    """Même scénario, mais classeur et tableau créés par Excel lui-même (pas par openpyxl)."""
    from bibliotheque.excel_com import ouvrir_excel
    excel = ouvrir_excel()
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1:C1").Value = ("Date", "Region", "Volume_GWh")
        for i in range(6):
            ws.Cells(i + 2, 1).Value = datetime(2026, 10, 20 + i)
            ws.Cells(i + 2, 2).Value = "Nord" if i % 2 else "Sud"
            ws.Cells(i + 2, 3).Value = 100.0 + i
        ws.Columns("A").NumberFormat = "dd/mm/yyyy"
        lo = ws.ListObjects.Add(xl.xlSrcRange, ws.Range("A1").CurrentRegion, None, xl.xlYes)
        lo.Name = "tblDonnees"
        ws2 = wb.Worksheets.Add()
        pc = wb.PivotCaches().Create(SourceType=xl.xlDatabase, SourceData="tblDonnees")
        pt = pc.CreatePivotTable(TableDestination=ws2.Range("A5"), TableName="TCD")
        pt.PivotFields("Date").Orientation = xl.xlRowField
        pt.PivotFields("Region").Orientation = xl.xlColumnField
        pt.AddDataField(pt.PivotFields("Volume_GWh"), "Volume (GWh)", xl.xlSum)
        try:
            pt.PivotFields("Date").NumberFormat = "dd/mm"
            print("OK      h. classeur natif Excel")
        except Exception as e:  # noqa: BLE001
            print("REFUSE  h. classeur natif Excel :", str(e)[:120])
        wb.Close(False)
    finally:
        excel.Quit()


def lenteurs():
    """Mesure l'ouverture, l'AutoFit, l'enregistrement et la fermeture selon le contenu du classeur
    (barres de données, lien hypertexte) : une session Excel lente se voit ici."""
    import time

    from openpyxl.formatting.rule import DataBarRule

    from bibliotheque.excel_com import ouvrir_excel
    with tempfile.TemporaryDirectory() as t:
        fichiers = {}
        for nom in ("simple", "databar", "lien", "databar+lien"):
            wb = openpyxl.Workbook()
            ws = wb.active
            for i in range(1, 6):
                ws.cell(row=i, column=1, value=f"ligne {i}")
                ws.cell(row=i, column=2, value=i * 10.5)
            if "databar" in nom:
                ws.conditional_formatting.add("B1:B5", DataBarRule(start_type="min", end_type="max", color="638EC6"))
            if "lien" in nom:
                ws["A8"].hyperlink = "c.html"
                ws["A8"].value = "Ouvrir"
                ws["A8"].style = "Hyperlink"
            fichiers[nom] = Path(t) / f"{nom.replace('+', '_')}.xlsx"
            wb.save(fichiers[nom])
        for nom, chemin in fichiers.items():
            etapes = []
            debut = time.time()
            excel = ouvrir_excel()
            etapes.append(("demarrage", time.time() - debut))
            try:
                for etape, action in (("ouvrir", lambda: excel.Workbooks.Open(str(chemin))),):
                    d = time.time()
                    wb = action()
                    etapes.append((etape, time.time() - d))
                for etape, action in (("autofit", lambda: wb.Worksheets(1).Columns("A:B").AutoFit()),
                                      ("enregistrer", lambda: wb.Save()), ("fermer", lambda: wb.Close(False))):
                    d = time.time()
                    action()
                    etapes.append((etape, time.time() - d))
            finally:
                d = time.time()
                excel.Quit()
                etapes.append(("quit", time.time() - d))
            print(f"{nom:14s}", "  ".join(f"{e}={s:5.1f}s" for e, s in etapes))


def lenteurs_v3():
    """Chronomètre chaque instruction de finaliser_csv_v3 sur un classeur préparé comme le pipeline."""
    import time

    import utilisation_csv as uc
    from bibliotheque.excel_com import ouvrir_excel
    with tempfile.TemporaryDirectory() as t:
        wb = openpyxl.load_workbook(RACINE / "tests" / "fixtures" / "classeur_complet.xlsx")
        csv = RACINE / "tests" / "fixtures" / "csv" / "DAT_20261027.csv"
        now = datetime(2026, 10, 27, 14, 30)
        uc.importer_csv(wb, csv, now)
        uc.creer_synthese(wb, "c.html", csv, now)
        chemin = Path(t) / "v3.xlsx"
        wb.save(chemin)
        excel = ouvrir_excel()
        try:
            wbx = excel.Workbooks.Open(str(chemin))
            etapes = (
                ("autofit Donnees", lambda: wbx.Worksheets("Donnees_V3").Columns("A:F").AutoFit()),
                ("autofit Carte", lambda: wbx.Worksheets("Carte_V3").Columns("A:H").AutoFit()),
                ("largeur A", lambda: setattr(wbx.Worksheets("Carte_V3").Columns("A"), "ColumnWidth", 24)),
                ("activate", lambda: wbx.Worksheets("Carte_V3").Activate()),
                ("save", lambda: wbx.Save()),
                ("close", lambda: wbx.Close(False)),
            )
            for nom, action in etapes:
                d = time.time()
                action()
                print(f"{nom:18s} {time.time() - d:5.1f}s")
        finally:
            d = time.time()
            excel.Quit()
            print(f"{'quit':18s} {time.time() - d:5.1f}s")


def tableau_croise():
    with tempfile.TemporaryDirectory() as t:
        source = preparer(Path(t))
        for n, f in (("a. instruction d'origine", a_originale), ("b. avant AddDataField", b_avant_donnees),
                     ("c. informations sur le champ", c_infos), ("d. après RefreshTable", d_refresh),
                     ("e. format dd/mm/yyyy", e_format_long), ("f. RowRange.NumberFormat", f_range_row),
                     ("g. disposition tabulaire", g_tabulaire),
                     ("i. format General", i_general)):
            variante(n, source, f)


if __name__ == "__main__":
    # python tests/diagnostic_excel.py [formats | tableau | lenteurs | lenteurs_v3]
    choix = sys.argv[1] if len(sys.argv) > 1 else "formats"
    {"formats": formats_de_nombre, "tableau": tableau_croise, "lenteurs": lenteurs,
     "lenteurs_v3": lenteurs_v3}[choix]()
