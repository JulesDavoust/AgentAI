"""Traduction du module VBA ``CréationCSV`` : génère le CSV quotidien de volumes par ville
(``DAT_aaaammjj.csv``, séparateur ``;``, Windows-1252, fins de ligne Windows).

Le classeur (``ThisWorkbook``) ne sert qu'au journal (feuille Log) : il est enregistré en place.
"""

from __future__ import annotations

import argparse
import logging
import random
import sys
from datetime import date, datetime
from pathlib import Path

import commun
from commun import Resultat

V3_CSV_FOLDER = "csv"  # VBA : "C:\Users\jdavoust\natran\VBATF_CSV\" (chemin personnel remplacé)
V3_CSV_PREFIX = "DAT_"
V3_DATE_FORMAT = "%Y%m%d"  # "yyyymmdd"
SEP = ";"
SHEET_LOG = "Log"
INJECTER_LIGNE_INVALIDE = False  # True = ajoute une ligne invalide

# Nom, latitude, longitude, volume de base (GWh)
VILLES = (
    ("Lille", 50.63, 3.06, 95), ("Paris", 48.86, 2.35, 150), ("Strasbourg", 48.57, 7.75, 75),
    ("Nantes", 47.22, -1.55, 70), ("Bordeaux", 44.84, -0.58, 85), ("Toulouse", 43.6, 1.44, 80),
    ("Marseille", 43.3, 5.37, 100), ("Lyon", 45.76, 4.84, 110), ("Caen", 49.18, -0.37, 55),
    ("Dijon", 47.32, 5.04, 45),
)


def villes() -> tuple:
    """VBA ``Villes``"""
    return VILLES


def v3_nom_csv(d: date, dossier=V3_CSV_FOLDER) -> Path:
    """VBA ``V3_NomCSV`` : chemin complet du CSV du jour ``d``."""
    return Path(dossier) / f"{V3_CSV_PREFIX}{d.strftime(V3_DATE_FORMAT)}.csv"


def v3_creer_dossier(chemin) -> None:
    """VBA ``V3_CreerDossier`` : crée le dossier et ses parents manquants."""
    Path(chemin).mkdir(parents=True, exist_ok=True)


def log_msg(wb, niveau: str, msg: str, maintenant: datetime | None = None) -> None:
    """VBA ``LogMsg`` de CréationCSV (variante avec ``On Error Resume Next`` : même effet)."""
    ws = commun.get_or_create_sheet(wb, SHEET_LOG)
    if ws["A1"].value in (None, ""):
        for i, titre in enumerate(("Horodatage", "Niveau", "Message"), 1):
            ws.cell(row=1, column=i, value=titre)
    n = commun.derniere_ligne(ws, 1) + 1
    ws.cell(row=n, column=1, value=maintenant or datetime.now()).number_format = "dd/mm/yyyy hh:mm:ss"
    ws.cell(row=n, column=2, value=niveau)
    ws.cell(row=n, column=3, value=msg)
    commun.journal.log(logging.ERROR if niveau == "ERREUR" else logging.INFO, "[%s] %s", niveau, msg)


def nb(v) -> str:
    """VBA ``nb`` : ``Trim$(Str$(CDbl(v)))``"""
    return commun.str_nb(v)


def ecrire_csv(wb, d: date, dossier=V3_CSV_FOLDER, rng=None, maintenant=datetime.now) -> Path:
    """VBA ``EcrireCSV`` : écrit le CSV de la date ``d`` et retourne son chemin."""
    rng = rng or random.Random()
    v3_creer_dossier(dossier)
    chemin = v3_nom_csv(d, dossier)
    if chemin.is_file():
        log_msg(wb, "INFO", f"CSV existant ecrase : {chemin}", maintenant())

    jour = d.strftime("%Y-%m-%d")
    with open(chemin, "w", encoding="cp1252", errors="replace", newline="\r\n") as f:
        f.write(SEP.join(("Date", "Ville", "Latitude", "Longitude", "Volume_GWh", "Temperature_C")) + "\n")
        for nom, lat, lon, base in villes():
            vol = round(base * (0.85 + rng.random() * 0.3), 1)
            temp = round(2 + rng.random() * 14, 1)
            f.write(SEP.join((jour, nom, nb(lat), nb(lon), nb(vol), nb(temp))) + "\n")
        if INJECTER_LIGNE_INVALIDE:
            f.write(SEP.join((jour, "VilleInconnue", "abc", "", "xx", "")) + "\n")

    log_msg(wb, "INFO", f"CSV genere : {chemin}", maintenant())
    return chemin


def _executer(classeur, action, maintenant) -> Resultat:
    """Corps commun de ``GenererCSVDuJourV3`` / ``GenererCSVDateV3`` (``On Error GoTo ErrHandler``)."""
    wb = commun.charger_classeur(classeur)
    try:
        resultat = action(wb)
    except Exception as erreur:
        description = getattr(erreur, "description", str(erreur))
        log_msg(wb, "ERREUR", f"Generation CSV : {commun.numero_erreur(erreur)} - {description}", maintenant())
        commun.journal.error("Erreur : %s", description)  # MsgBox vbCritical
        resultat = Resultat(False, [], f"Erreur : {description}")
    wb.save(classeur)  # le classeur ouvert garde son journal : enregistré en place
    return resultat


def generer_csv_du_jour_v3(classeur, v3_csv_folder=V3_CSV_FOLDER, *, rng=None, maintenant=datetime.now) -> Resultat:
    """VBA ``GenererCSVDuJourV3`` : CSV de la date du jour."""
    def action(wb):
        chemin = ecrire_csv(wb, maintenant().date(), v3_csv_folder, rng, maintenant)
        message = f"CSV genere :\n{chemin}"
        commun.journal.info(message)  # MsgBox vbInformation
        return Resultat(True, [chemin], message)
    return _executer(classeur, action, maintenant)


def generer_csv_date_v3(classeur, v3_csv_folder=V3_CSV_FOLDER, saisie=None, *, rng=None,
                        maintenant=datetime.now) -> Resultat:
    """VBA ``GenererCSVDateV3`` : CSV d'une date saisie (jj/mm/aaaa). ``saisie`` remplace l'InputBox
    (défaut de l'InputBox : la date du jour)."""
    def action(wb):
        texte = (commun.fmt_date(maintenant()) if saisie is None else saisie).strip()
        if texte == "":
            return Resultat(True, [], "")
        if not commun.is_date(texte):
            message = f"Date invalide : {texte}"
            commun.journal.warning(message)  # MsgBox vbExclamation
            return Resultat(False, [], message)
        chemin = ecrire_csv(wb, commun.c_date(texte), v3_csv_folder, rng, maintenant)
        message = f"CSV genere :\n{chemin}"
        commun.journal.info(message)
        return Resultat(True, [chemin], message)
    return _executer(classeur, action, maintenant)


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="VBA CréationCSV : génération du CSV quotidien")
    sous = p.add_subparsers(dest="commande", required=True)
    s1 = sous.add_parser("generer-csv-du-jour-v3", help="GenererCSVDuJourV3 : CSV du jour")
    s2 = sous.add_parser("generer-csv-date-v3", help="GenererCSVDateV3 : CSV d'une date saisie")
    for s in (s1, s2):
        s.add_argument("classeur", help="classeur (remplace ThisWorkbook ; reçoit le journal Log)")
        s.add_argument("--v3-csv-folder", default=V3_CSV_FOLDER, help="dossier des CSV (V3_CSV_FOLDER)")
    s2.add_argument("--saisie", default=None, help="date jj/mm/aaaa (InputBox ; défaut : date du jour)")
    a = p.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    if a.commande == "generer-csv-du-jour-v3":
        res = generer_csv_du_jour_v3(a.classeur, a.v3_csv_folder)
    else:
        res = generer_csv_date_v3(a.classeur, a.v3_csv_folder, a.saisie)
    return 0 if res.ok else 1


if __name__ == "__main__":
    sys.exit(main())
