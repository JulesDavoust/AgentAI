"""Traduction du module VBA ``marco_25_10`` (suivi distribution gaz).

Pipeline : maj données -> filtre -> visuel -> sauvegarde -> mail.
Feuille2 : données (tableau ``tblDonnees``) ; Feuille1 : visuel (KPI + TCD + graphique).
"""
from __future__ import annotations

import argparse
import logging
import random
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Font  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.filters import AutoFilter, Filters, FilterColumn  # noqa: E402
from openpyxl.worksheet.table import Table, TableStyleInfo  # noqa: E402

logger = logging.getLogger(__name__)

# --- CONFIGURATION ---------------------------------------------------------
SHEET_VISUEL = "Feuille1"
SHEET_DATA = "Feuille2"
SHEET_LOG = "Log"
TABLE_NAME = "tblDonnees"
FILE_TITLE = "VBA_Natran_Test"
FILE_SUFFIX = "VBATF"
MAIL_TO = "jdavoust201@headmind.com"
MAIL_CC = ""
SEND_MAIL = True
ENTETES = ["Date", "Region", "Volume_GWh", "Temperature_C", "Statut"]


def regions() -> list[str]:
    """VBA : fonction ``Regions``."""
    return ["Nord", "Sud", "Est", "Ouest", "Centre"]


# --- OUTILS ------------------------------------------------------------------
def get_or_create_sheet(wb: Workbook, nom: str):
    """VBA : ``GetOrCreateSheet`` (crée la feuille à la fin si absente)."""
    if nom in wb.sheetnames:
        return wb[nom]
    return wb.create_sheet(title=nom)


def log_msg(wb: Workbook, niveau: str, msg: str, maintenant: datetime | None = None) -> None:
    """VBA : ``LogMsg`` (ajoute une ligne dans la feuille Log)."""
    maintenant = maintenant or datetime.now()
    ws = get_or_create_sheet(wb, SHEET_LOG)
    if ws["A1"].value in (None, ""):
        for c, titre in enumerate(["Horodatage", "Niveau", "Message"], start=1):
            ws.cell(row=1, column=c, value=titre)
    # équivalent de Cells(Rows.Count, 1).End(xlUp).Row + 1
    derniere = ws.max_row
    while derniere > 1 and ws.cell(row=derniere, column=1).value in (None, ""):
        derniere -= 1
    n = derniere + 1
    ws.cell(row=n, column=1, value=maintenant).number_format = "dd/mm/yyyy hh:mm:ss"
    ws.cell(row=n, column=2, value=niveau)
    ws.cell(row=n, column=3, value=msg)


def _en_date(v) -> date | None:
    """Convertit une valeur de cellule en ``date`` (``CLng`` de la date VBA)."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _autofit(ws, colonnes: str | None = None) -> None:
    """VBA : ``Columns(...).AutoFit`` (approximation par la longueur du texte)."""
    for col in ws.iter_cols():
        lettre = get_column_letter(col[0].column)
        if colonnes and lettre not in colonnes:
            continue
        largeur = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        if largeur:
            ws.column_dimensions[lettre].width = largeur + 2


def _vider_feuille(ws) -> None:
    """VBA : ``ws.Cells.Clear``."""
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)
    ws._charts = []


def _derniere_ligne_table(ws) -> int:
    """Dernière ligne (1-indexée) de la plage du tableau ``tblDonnees``."""
    ref = ws.tables[TABLE_NAME].ref  # ex. "A1:E71"
    return int(''.join(ch for ch in ref.split(":")[1] if ch.isdigit()))


def lire_table(wb: Workbook) -> pd.DataFrame:
    """Lit ``tblDonnees`` dans un DataFrame (colonnes = ENTETES, Date en ``date``)."""
    ws = wb[SHEET_DATA]
    lignes = []
    for r in range(2, _derniere_ligne_table(ws) + 1):
        lignes.append([ws.cell(row=r, column=c).value for c in range(1, 6)])
    df = pd.DataFrame(lignes, columns=ENTETES)
    df["Date"] = df["Date"].map(_en_date)
    return df


# --- 1. STRUCTURE / DONNEES -----------------------------------------------------
def simuler_volume(reg: str, rng: random.Random) -> float:
    """VBA : ``SimulerVolume`` (``Rnd`` remplacé par ``rng.random``)."""
    base = {"Nord": 120, "Sud": 95, "Est": 80, "Ouest": 70}.get(reg, 60)
    return round(base * (0.85 + rng.random() * 0.3), 1)


def simuler_temperature(rng: random.Random) -> float:
    """VBA : ``SimulerTemperature``."""
    return round(2 + rng.random() * 14, 1)


def ensure_structure(wb: Workbook, rng: random.Random, aujourd_hui: date | None = None,
                     maintenant: datetime | None = None) -> bool:
    """VBA : ``EnsureStructure``. Crée feuilles + 14 jours d'historique si le tableau est absent.

    Retourne True si la structure a été créée.
    """
    aujourd_hui = aujourd_hui or date.today()
    ws_d = get_or_create_sheet(wb, SHEET_DATA)
    get_or_create_sheet(wb, SHEET_VISUEL)
    if TABLE_NAME in ws_d.tables:
        return False

    _vider_feuille(ws_d)
    for c, titre in enumerate(ENTETES, start=1):
        ws_d.cell(row=1, column=c, value=titre)
    ligne = 2
    for i in range(14, 0, -1):
        for r in regions():
            ws_d.cell(row=ligne, column=1, value=aujourd_hui - timedelta(days=i))
            ws_d.cell(row=ligne, column=2, value=r)
            ws_d.cell(row=ligne, column=3, value=simuler_volume(r, rng))
            ws_d.cell(row=ligne, column=4, value=simuler_temperature(rng))
            ws_d.cell(row=ligne, column=5, value="Historique")
            ligne += 1
    tbl = Table(displayName=TABLE_NAME, ref=f"A1:E{ligne - 1}")
    tbl.autoFilter = AutoFilter(ref=tbl.ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_d.add_table(tbl)
    for r in range(2, ligne):
        ws_d.cell(row=r, column=1).number_format = "dd/mm/yyyy"
    _autofit(ws_d, "ABCDE")
    log_msg(wb, "INFO", "Structure initiale creee", maintenant)
    return True


def ligne_existe(wb: Workbook, d: date, reg: str) -> bool:
    """VBA : ``LigneExiste`` (comparaison de région sensible à la casse, comme ``=`` VBA)."""
    ws = wb[SHEET_DATA]
    for r in range(2, _derniere_ligne_table(ws) + 1):
        if _en_date(ws.cell(row=r, column=1).value) == d and ws.cell(row=r, column=2).value == reg:
            return True
    return False


def ajouter_donnees_du_jour(wb: Workbook, rng: random.Random, aujourd_hui: date | None = None,
                            maintenant: datetime | None = None) -> int:
    """VBA : ``AjouterDonneesDuJour``. Ajoute une ligne par région absente ; retourne le nombre ajouté."""
    aujourd_hui = aujourd_hui or date.today()
    ws = wb[SHEET_DATA]
    nb = 0
    for r in regions():
        if not ligne_existe(wb, aujourd_hui, r):
            n = _derniere_ligne_table(ws) + 1
            ws.cell(row=n, column=1, value=aujourd_hui).number_format = "dd/mm/yyyy"
            ws.cell(row=n, column=2, value=r)
            ws.cell(row=n, column=3, value=simuler_volume(r, rng))
            ws.cell(row=n, column=4, value=simuler_temperature(rng))
            ws.cell(row=n, column=5, value="Phase de test")
            ws.tables[TABLE_NAME].ref = f"A1:E{n}"
            if ws.tables[TABLE_NAME].autoFilter is None:
                ws.tables[TABLE_NAME].autoFilter = AutoFilter()
            ws.tables[TABLE_NAME].autoFilter.ref = f"A1:E{n}"
            nb += 1
    log_msg(wb, "INFO", f"{nb} ligne(s) ajoutee(s) pour le {aujourd_hui:%d/%m/%Y}", maintenant)
    return nb


# --- 2. FILTRE --------------------------------------------------------------------
def est_region_valide(s: str) -> bool:
    """VBA : ``EstRegionValide`` (comparaison insensible à la casse, ``vbTextCompare``)."""
    return any(r.casefold() == s.casefold() for r in regions())


def appliquer_filtre(wb: Workbook, saisie: str | None, maintenant: datetime | None = None) -> str | None:
    """VBA : ``AppliquerFiltre``. ``saisie`` remplace l'``InputBox`` (None/vide = aucun filtre).

    Retourne le filtre appliqué, ou None. Une région inconnue ne filtre rien (le ``MsgBox`` devient un warning).
    """
    ws = wb[SHEET_DATA]
    tbl = ws.tables[TABLE_NAME]
    # ShowAllData : on retire un filtre existant et on réaffiche les lignes
    if tbl.autoFilter is None:
        tbl.autoFilter = AutoFilter(ref=tbl.ref)
    tbl.autoFilter.filterColumn = []
    for r in range(2, _derniere_ligne_table(ws) + 1):
        ws.row_dimensions[r].hidden = False

    saisie = (saisie or "").strip()
    if saisie == "":
        return None
    if not est_region_valide(saisie):
        logger.warning("Region inconnue : %s", saisie)
        return None
    tbl.autoFilter.filterColumn = [FilterColumn(colId=1, filters=Filters(filter=[saisie]))]
    for r in range(2, _derniere_ligne_table(ws) + 1):
        valeur = str(ws.cell(row=r, column=2).value or "")
        ws.row_dimensions[r].hidden = valeur.casefold() != saisie.casefold()
    log_msg(wb, "INFO", f"Filtre applique : {saisie}", maintenant)
    return saisie


# --- 3. VISUEL ----------------------------------------------------------------------
def calculer_kpi(df: pd.DataFrame) -> dict:
    """KPI de Feuille1 (B2, E2, H2) calculés en Python sur toute la table (pas sur le filtre)."""
    if df.empty:
        return {"volume_total": 0.0, "moyenne_par_jour": None, "region_pic": None}
    total = float(df["Volume_GWh"].sum())
    nb_jours = df["Date"].nunique()
    return {
        "volume_total": total,
        "moyenne_par_jour": total / nb_jours,
        # MATCH(MAX(...), ..., 0) : première occurrence du maximum
        "region_pic": df.loc[df["Volume_GWh"].idxmax(), "Region"],
    }


def tableau_croise(df: pd.DataFrame) -> pd.DataFrame:
    """Équivalent du TCD ``TCD_Volume`` : lignes = Date, colonnes = Region, valeurs = somme du volume."""
    return df.pivot_table(index="Date", columns="Region", values="Volume_GWh", aggfunc="sum")


def creer_visuel(wb: Workbook, dossier_png: str | Path, maintenant: datetime | None = None) -> Path:
    """VBA : ``CreerVisuel``. Écrit KPI + TCD (valeurs) dans Feuille1 et exporte le graphique en PNG.

    Retourne le chemin du PNG.
    """
    maintenant = maintenant or datetime.now()
    ws_v = wb[SHEET_VISUEL]
    df = lire_table(wb)
    _vider_feuille(ws_v)

    ws_v["A1"] = f"Suivi distribution gaz - mis a jour le {maintenant:%d/%m/%Y %H:%M}"
    ws_v["A1"].font = Font(bold=True, size=14)
    ws_v["A2"] = "Volume total (GWh)"
    ws_v["B2"] = f"=SUM({TABLE_NAME}[Volume_GWh])"
    ws_v["D2"] = "Moyenne / jour (GWh)"
    ws_v["E2"] = (f"=SUM({TABLE_NAME}[Volume_GWh])/SUMPRODUCT(1/COUNTIF({TABLE_NAME}[Date],{TABLE_NAME}[Date]))")
    ws_v["G2"] = "Region du pic"
    for c in ("A2", "D2", "G2"):
        ws_v[c].font = Font(bold=True)
    ws_v["B2"].number_format = ws_v["E2"].number_format = "#,##0.0"
    ws_v["H2"] = (f"=INDEX({TABLE_NAME}[Region],MATCH(MAX({TABLE_NAME}[Volume_GWh]),{TABLE_NAME}[Volume_GWh],0))")

    # TCD écrit en valeurs (openpyxl ne crée pas de tableau croisé dynamique)
    tcd = tableau_croise(df)
    ws_v["A5"] = "Volume (GWh)"
    for j, reg in enumerate(tcd.columns, start=2):
        ws_v.cell(row=5, column=j, value=reg)
    ws_v.cell(row=5, column=len(tcd.columns) + 2, value="Total general")
    for i, (d, ligne) in enumerate(tcd.iterrows(), start=6):
        ws_v.cell(row=i, column=1, value=d).number_format = "dd/mm"
        for j, v in enumerate(ligne, start=2):
            ws_v.cell(row=i, column=j, value=None if pd.isna(v) else float(v)).number_format = "#,##0.0"
        ws_v.cell(row=i, column=len(tcd.columns) + 2, value=float(ligne.sum())).number_format = "#,##0.0"
    _autofit(ws_v, "ABCDEFGH")

    # Graphique (matplotlib -> PNG)
    dossier_png = Path(dossier_png)
    dossier_png.mkdir(parents=True, exist_ok=True)
    chemin_png = dossier_png / "volumes_par_region.png"
    fig, ax = plt.subplots(figsize=(7, 4))
    for reg in tcd.columns:
        ax.plot(list(tcd.index), tcd[reg].tolist(), label=reg)
    ax.set_title("Volumes distribues par region (GWh)")
    if len(tcd.columns):
        ax.legend()
    fig.autofmt_xdate()
    fig.savefig(chemin_png)
    plt.close(fig)

    log_msg(wb, "INFO", "Visuel regenere", maintenant)
    return chemin_png


# --- 4. SAUVEGARDE --------------------------------------------------------------------
def sauvegarder_version(wb: Workbook, dossier: str | Path, maintenant: datetime | None = None,
                        maintenant_log: datetime | None = None) -> Path:
    """VBA : ``SauvegarderVersion`` (``SaveCopyAs``). Extension .xlsx (pas de macros côté Python)."""
    maintenant = maintenant or datetime.now()
    dossier = Path(dossier)
    dossier.mkdir(parents=True, exist_ok=True)
    chemin = dossier / f"{FILE_TITLE}_{maintenant:%Y%m%d_%H%M%S}_{FILE_SUFFIX}.xlsx"
    log_msg(wb, "INFO", f"Sauvegarde : {chemin}", maintenant_log or maintenant)
    wb.save(chemin)
    return chemin


# --- 5. MAIL ----------------------------------------------------------------------------
def composer_mail(chemin_pj: str | Path, aujourd_hui: date | None = None) -> EmailMessage:
    """VBA : ``EnvoyerMail`` (partie composition uniquement ; l'envoi Outlook n'est pas traduit)."""
    aujourd_hui = aujourd_hui or date.today()
    msg = EmailMessage()
    msg["To"] = MAIL_TO
    if MAIL_CC:
        msg["Cc"] = MAIL_CC
    msg["Subject"] = f"[PHASE DE TEST] Mise a jour suivi distribution - {aujourd_hui:%d/%m/%Y}"
    msg.set_content(
        "Bonjour,\r\n\r\n"
        "Ceci est un envoi automatique en phase de test.\r\n"
        "Le graphique et les donnees ont ete mis a jour.\r\n"
        f"Fichier : {Path(chemin_pj).name}\r\n\r\n"
        "Cordialement,\r\nRobot VBATF"
    )
    return msg


def envoyer_mail(chemin_pj: str | Path, dossier_brouillon: str | Path,
                 aujourd_hui: date | None = None) -> Path:
    """VBA : ``EnvoyerMail``. Pas d'Outlook : le message est écrit en brouillon ``.eml``.

    # TODO(vba2py): envoi réel (Outlook CreateObject / .Send) non traduit ; aucun accès réseau autorisé.
    """
    msg = composer_mail(chemin_pj, aujourd_hui)
    chemin_pj = Path(chemin_pj)
    msg.add_attachment(chemin_pj.read_bytes(), maintype="application",
                       subtype="octet-stream", filename=chemin_pj.name)
    dossier_brouillon = Path(dossier_brouillon)
    dossier_brouillon.mkdir(parents=True, exist_ok=True)
    chemin_eml = dossier_brouillon / (chemin_pj.stem + ".eml")
    chemin_eml.write_bytes(bytes(msg))
    return chemin_eml


# --- POINT D'ENTREE ----------------------------------------------------------------------
def lancer_pipeline_complet(chemin_classeur: str | Path, dossier_sauvegarde: str | Path,
                            region_filtre: str | None = None, rng: random.Random | None = None,
                            maintenant: datetime | None = None) -> Path:
    """VBA : ``LancerPipelineComplet``. Retourne le chemin de la copie sauvegardée.

    Si ``chemin_classeur`` n'existe pas, un classeur vierge est utilisé (la structure est recréée).
    En cas d'erreur : journalisée (feuille Log + logging) puis relancée (le VBA affichait un MsgBox).
    """
    rng = rng or random.Random()
    maintenant = maintenant or datetime.now()
    chemin_classeur = Path(chemin_classeur)
    wb = load_workbook(chemin_classeur) if chemin_classeur.exists() else Workbook()
    try:
        ensure_structure(wb, rng, maintenant.date(), maintenant)
        ajouter_donnees_du_jour(wb, rng, maintenant.date(), maintenant)
        appliquer_filtre(wb, region_filtre, maintenant)
        creer_visuel(wb, dossier_sauvegarde, maintenant)
        chemin = sauvegarder_version(wb, dossier_sauvegarde, maintenant)
        envoyer_mail(chemin, dossier_sauvegarde, maintenant.date())
        log_msg(wb, "INFO", f"Pipeline termine : {chemin}", maintenant)
        wb.save(chemin)  # la copie contient aussi les derniers messages du journal
        logger.info("Pipeline termine : %s", chemin)
        return chemin
    except Exception as exc:
        log_msg(wb, "ERREUR", f"{type(exc).__name__} - {exc}", maintenant)
        logger.error("Erreur : %s", exc)
        raise


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pipeline suivi distribution gaz (traduction de marco_25_10.bas)")
    parser.add_argument("classeur", help="classeur .xlsx source (créé s'il n'existe pas)")
    parser.add_argument("dossier", help="dossier de sauvegarde (copie, PNG, brouillon .eml)")
    parser.add_argument("--region", default=None, help="filtre région (vide = aucun)")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO)
    print(lancer_pipeline_complet(args.classeur, args.dossier, args.region))
