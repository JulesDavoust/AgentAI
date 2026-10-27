"""Traduction du module VBA ``TestInterfaceMVS`` : API « terminal » (Init_MVS, Lire, Ecrire, Touche,
Recherche...) au-dessus de l'écran simulé par ``SimuMainframe``, et ses démonstrations.

L'état du terminal (variables de module ``ecran``, ``curLigne``, ``curCol``) est regroupé dans la classe
``TerminalMVS`` : il est recréé à chaque commande. Le classeur est enregistré en place.
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import commun
import simu_mainframe as simu
from bibliotheque.excel_com import preparer_feuille
from commun import ErreurVBA
from connecteurs import excel as connecteur_excel
from simu_mainframe import (C_CHAMP, C_MSG, L_COM, L_DATE, L_ENC, L_ISIN, L_LIB, L_MSG, L_STAT, LEN_COM, LEN_DATE,
                            LEN_ENC, LEN_ISIN, LEN_LIB, LEN_MSG, LEN_STAT, NB_COLS, NB_LIGNES)

TIMEOUT = 200
SHEET_MVS = "MVS"
AFFICHER_CURSEUR = True  # sélectionne la cellule courante (visuel)


def attendre(temps: int = TIMEOUT) -> None:
    """VBA ``Attendre`` (``DoEvents``) : sans effet en Python."""


def verifier_coord(ligne: int, colonne: int) -> None:
    """VBA ``VerifierCoord`` : ``Err.Raise vbObjectError + 20`` hors de l'écran."""
    if ligne < 1 or ligne > NB_LIGNES or colonne < 1 or colonne > NB_COLS:
        raise ErreurVBA(commun.VB_OBJECT_ERROR + 20, f"Coordonnees hors ecran : ({ligne},{colonne})")


class TerminalMVS:
    """API « terminal » sur la feuille MVS d'un classeur."""

    def __init__(self, wb, date_du_jour=date.today):
        self.wb = wb
        self.ecran = None
        self.cur_ligne = 0
        self.cur_col = 0
        self.date_du_jour = date_du_jour

    # --- API terminal
    def init_mvs(self) -> bool:
        """VBA ``Init_MVS``"""
        self.ecran = self.wb[SHEET_MVS] if SHEET_MVS in self.wb.sheetnames else None
        if self.ecran is None:
            commun.journal.error("Impossible de se connecter a MVS : feuille '%s' absente.\n"
                                 "Clique d'abord sur le bouton 'Ouvrir le mainframe'.", SHEET_MVS)
            return False
        if self.modele_session() != 3:
            commun.journal.error("Votre terminal doit etre parametre en Modele 3 (32x80)")
            return False
        self.cur_ligne, self.cur_col = 1, 1
        return True

    def modele_session(self) -> int:
        """VBA ``Modele_Session``"""
        return 0 if self.ecran is None else 3

    def buffer_ligne(self, l: int) -> str:
        """VBA ``BufferLigne`` (version de TestInterfaceMVS, sur ``ecran``)."""
        return simu.buffer_ligne(self.ecran, l)

    def normaliser_ligne(self, l: int) -> None:
        """VBA ``NormaliserLigne`` (version de TestInterfaceMVS, sur ``ecran``)."""
        simu.normaliser_ligne(self.ecran, l)

    def lire(self, ligne: int, colonne: int, taille: int) -> str:
        """VBA ``Lire`` : ``taille`` caractères à partir de (ligne, colonne), complétés par des espaces."""
        verifier_coord(ligne, colonne)
        s = self.buffer_ligne(ligne)[colonne - 1:colonne - 1 + taille]
        attendre(TIMEOUT)
        return (s + " " * taille)[:taille]

    def se_positionner(self, ligne: int, colonne: int) -> None:
        """VBA ``Se_Positionner``"""
        verifier_coord(ligne, colonne)
        self.cur_ligne, self.cur_col = ligne, colonne
        if AFFICHER_CURSEUR:
            commun.activer_feuille(self.wb, self.ecran)
            commun.selectionner(self.ecran, f"{get_column_letter(colonne)}{ligne}")
        attendre(TIMEOUT)

    def ecrire(self, ligne: int, colonne: int, chaine: str) -> None:
        """VBA ``Ecrire`` : un caractère par cellule (un espace efface la cellule)."""
        self.se_positionner(ligne, colonne)
        self.normaliser_ligne(ligne)
        for k, ch in enumerate(chaine):
            if colonne + k > NB_COLS:
                break
            cell = self.ecran.cell(row=ligne, column=colonne + k)
            cell.value = None if ch == " " else ch
        self.cur_col = min(colonne + len(chaine), NB_COLS)
        attendre(TIMEOUT)

    def touche(self, cle: str) -> None:
        """VBA ``Touche`` : @E = ENTREE, @C = CLEAR, @3 = PF3, @F = efface la fin de ligne, texte = saisie."""
        k = cle.upper()
        if k == "@E":
            self.appeler_mvs("MVS_Entree")
        elif k == "@C":
            self.appeler_mvs("MVS_Effacer")
        elif k == "@3":
            self.appeler_mvs("MVS_PF3")
        elif k == "@F":
            self.normaliser_ligne(self.cur_ligne)
            for c in range(self.cur_col, NB_COLS + 1):
                self.ecran.cell(row=self.cur_ligne, column=c).value = None
        elif k in ("", "@T"):
            pass
        elif cle[:1] == "@":
            logging.debug("Touche non geree : %s", cle)
        else:
            self.ecrire(self.cur_ligne, self.cur_col, cle)
        attendre(TIMEOUT)

    def recherche(self, texte: str) -> bool:
        """VBA ``Recherche`` : le texte figure-t-il sur l'écran (insensible à la casse) ?"""
        return any(texte.lower() in self.buffer_ligne(l).lower() for l in range(1, NB_LIGNES + 1))

    # --- outils
    def appeler_mvs(self, nom_macro: str) -> None:
        """VBA ``AppelerMVS`` : ``Application.Run`` d'une macro de SimuMainframe (même classeur) ;
        erreur signalée sans interrompre (``On Error Resume Next``)."""
        macros = {
            "MVS_Entree": lambda: simu.mvs_entree(self.wb, self.date_du_jour()),
            "MVS_Effacer": lambda: simu.mvs_effacer(self.wb),
            "MVS_PF3": lambda: simu.mvs_pf3(self.wb),
        }
        try:
            if nom_macro not in macros:
                raise ErreurVBA(1004, f"Impossible d'exécuter la macro '{nom_macro}'.")
            macros[nom_macro]()
        except Exception as erreur:
            commun.journal.error("Macro '%s' introuvable ou en erreur : %s\n"
                                 "Le module de simulation du mainframe (SimuMainframe) est-il present ?",
                                 nom_macro, getattr(erreur, "description", erreur))

    def selectionner_zone(self, l: int, c: int, n: int) -> None:
        """VBA ``SelectionnerZone``"""
        commun.activer_feuille(self.wb, self.ecran)
        commun.selectionner(self.ecran, f"{get_column_letter(c)}{l}:{get_column_letter(c + n - 1)}{l}")

    def consulter_isin(self, isin: str) -> None:
        """VBA ``ConsulterISIN``"""
        self.se_positionner(L_ISIN, C_CHAMP)
        self.touche("@F")
        self.ecrire(L_ISIN, C_CHAMP, isin)
        self.touche("@E")


def demander_entier(saisie, mini: int, maxi: int):
    """VBA ``DemanderEntier`` pour une saisie donnée : retourne l'entier, None si la saisie est vide
    (= Annuler). Une valeur invalide lève ``ValueError`` (VBA : MsgBox puis nouvelle saisie)."""
    s = ("" if saisie is None else str(saisie)).strip()
    if s == "":
        return None
    if commun.is_numeric(s):
        valeur = float(s.replace(",", "."))
        if mini <= valeur <= maxi and valeur == int(valeur):
            return int(valeur)
    raise ValueError(f"Valeur invalide : saisir un entier entre {mini} et {maxi}.")


def _ouvrir(classeur, date_du_jour=date.today):
    wb = commun.charger_classeur(classeur)
    return wb, TerminalMVS(wb, date_du_jour)


# ----------------------------------------------------------------------------- bouton 2

def interagir_mainframe(classeur, ligne=None, colonne=None, longueur="10", texte=None, *,
                        date_du_jour=date.today) -> commun.Resultat:
    """VBA ``InteragirMainframe``. Les InputBox deviennent ``ligne``, ``colonne``, ``longueur`` (défaut
    « 10 ») et ``texte`` ; le choix LIRE / ECRIRE du MsgBox se déduit de ``texte`` (absent = lecture,
    fourni = écriture suivie d'ENTREE). Une saisie vide = Annuler (la macro s'arrête)."""
    wb, t = _ouvrir(classeur, date_du_jour)
    if SHEET_MVS not in wb.sheetnames:
        message = "Le mainframe n'est pas encore ouvert.\nClique d'abord sur le bouton 'Ouvrir le mainframe'."
        commun.journal.warning(message)
        return commun.Resultat(False, [], message)
    if not t.init_mvs():
        return commun.Resultat(False, [], "Init_MVS a echoue")
    ecriture = texte is not None
    try:
        lig = demander_entier(ligne, 1, NB_LIGNES)
        if lig is None:
            return commun.Resultat(True, [], "")
        col = demander_entier(colonne, 1, NB_COLS)
        if col is None:
            return commun.Resultat(True, [], "")
        max_long = NB_COLS - col + 1
        if not ecriture:
            lon = demander_entier(longueur, 1, max_long)
            if lon is None:
                return commun.Resultat(True, [], "")
            lu = t.lire(lig, col, lon)
            t.selectionner_zone(lig, col, lon)
            message = f"Ligne {lig}, colonne {col}, longueur {lon}\n\nContenu : [{lu}]"
        else:
            if texte == "":
                return commun.Resultat(True, [], "")
            if len(texte) > max_long:
                texte = texte[:max_long]
                commun.journal.warning("Texte tronque a %s caracteres (fin de ligne).", max_long)
            t.ecrire(lig, col, texte)
            t.selectionner_zone(lig, col, len(texte))
            t.touche("@E")
            message = "Reponse du mainframe :\n" + t.lire(L_MSG, C_MSG, LEN_MSG).strip()
    except ValueError as erreur:
        commun.journal.warning(str(erreur))
        return commun.Resultat(False, [], str(erreur))
    commun.journal.info(message)
    # wb.save(classeur)
    return commun.Resultat(True, [], message)


# ----------------------------------------------------------------------------- démonstrations

def demo_ecrire_et_lire(classeur, *, date_du_jour=date.today, maintenant=datetime.now) -> commun.Resultat:
    """VBA ``Demo_EcrireEtLire`` : consulte FR0010000002, lit le libellé, écrit un commentaire, ENTREE."""
    wb, t = _ouvrir(classeur, date_du_jour)
    if not t.init_mvs():
        return commun.Resultat(False, [], "Init_MVS a echoue")
    t.consulter_isin("FR0010000002")
    libelle = t.lire(L_LIB, C_CHAMP, LEN_LIB).strip()
    statut = t.lire(L_STAT, C_CHAMP, LEN_STAT).strip()
    encours = t.lire(L_ENC, C_CHAMP, LEN_ENC).strip()
    lecture = f"Lecture mainframe :\n{libelle}\nStatut : {statut}\nEncours (MEUR) : {encours}"
    commun.journal.info(lecture)

    t.se_positionner(L_COM, C_CHAMP)
    t.touche("@F")
    t.ecrire(L_COM, C_CHAMP, "MAJ PAR VBA " + maintenant().strftime("%H:%M:%S"))
    t.touche("@E")
    reponse = "Reponse du mainframe : " + t.lire(L_MSG, C_MSG, LEN_MSG).strip()
    commun.journal.info(reponse)
    wb.save(classeur)
    return commun.Resultat(True, [], lecture + "\n" + reponse)


def demo_lire_saisie_manuelle(classeur, *, date_du_jour=date.today) -> commun.Resultat:
    """VBA ``Demo_LireSaisieManuelle``"""
    wb, t = _ouvrir(classeur, date_du_jour)
    if not t.init_mvs():
        return commun.Resultat(False, [], "Init_MVS a echoue")
    message = (f"ISIN saisi : {t.lire(L_ISIN, C_CHAMP, LEN_ISIN).strip()}\n"
               f"Commentaire : {t.lire(L_COM, C_CHAMP, LEN_COM).strip()}\n"
               f"Message : {t.lire(L_MSG, C_MSG, LEN_MSG).strip()}")
    commun.journal.info(message)
    return commun.Resultat(True, [], message)


def demo_verifier_existence_fsic(classeur, isin="FR0010000003", *, date_du_jour=date.today) -> commun.Resultat:
    """VBA ``Demo_VerifierExistenceFSIC`` : ``isin`` remplace l'InputBox (défaut : FR0010000003)."""
    wb, t = _ouvrir(classeur, date_du_jour)
    isin = isin.strip().upper()
    if isin == "":
        return commun.Resultat(True, [], "")
    if not t.init_mvs():
        return commun.Resultat(False, [], "Init_MVS a echoue")
    t.consulter_isin(isin)
    if t.recherche("INCONNU"):
        message = f"{isin} n'existe pas dans FSIC."
    else:
        message = f"{isin} existe dans FSIC :\n" + t.lire(L_LIB, C_CHAMP, LEN_LIB).strip()
    commun.journal.info(message)
    wb.save(classeur)
    return commun.Resultat(True, [], message)


def demo_extraction_vers_feuille(classeur, *, date_du_jour=date.today, excel=None) -> commun.Resultat:
    """VBA ``Demo_ExtractionVersFeuille`` : interroge 7 ISIN (le dernier est inconnu) et remplit la feuille
    Extraction_MVS. L'AutoFit et l'activation de la feuille sont faits par Excel."""
    liste = ("FR0010000001", "FR0010000002", "FR0010000003", "FR0010000004", "FR0010000005", "FR0010000006",
             "FR0000000000")
    wb, t = _ouvrir(classeur, date_du_jour)
    if not t.init_mvs():
        return commun.Resultat(False, [], "Init_MVS a echoue")
    ws_x = commun.get_or_create_sheet(wb, "Extraction_MVS")
    preparer_feuille(ws_x)  # wsX.Cells.Clear
    for i, titre in enumerate(("ISIN", "Libelle", "Statut", "Encours_MEUR", "Date_MAJ", "Resultat"), 1):
        c = ws_x.cell(row=1, column=i, value=titre)
        c.font = Font(bold=True)
    for i, isin in enumerate(liste):
        r = i + 2
        t.consulter_isin(isin)
        ws_x.cell(row=r, column=1, value=isin)
        if t.recherche("INCONNU"):
            ws_x.cell(row=r, column=6, value="INCONNU")
        else:
            ws_x.cell(row=r, column=2, value=t.lire(L_LIB, C_CHAMP, LEN_LIB).strip())
            ws_x.cell(row=r, column=3, value=t.lire(L_STAT, C_CHAMP, LEN_STAT).strip())
            ws_x.cell(row=r, column=4, value=commun.vba_val(t.lire(L_ENC, C_CHAMP, LEN_ENC).strip()))
            ws_x.cell(row=r, column=5, value=t.lire(L_DATE, C_CHAMP, LEN_DATE).strip())
            ws_x.cell(row=r, column=6, value="OK")
    commun.activer_feuille(wb, ws_x)
    wb.save(classeur)
    (excel or connecteur_excel.ajuster_feuille)(classeur, "Extraction_MVS", "A:F", True)
    message = f"{len(liste)} ISIN interroges sur le mainframe."
    commun.journal.info(message)
    return commun.Resultat(True, [], message)


# ----------------------------------------------------------------------------- ligne de commande

def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="VBA TestInterfaceMVS : interaction avec le mainframe simulé")
    sous = p.add_subparsers(dest="commande", required=True)
    s = {}
    for nom, aide in (("interagir-mainframe", "InteragirMainframe : lire ou écrire une zone"),
                      ("demo-ecrire-et-lire", "Demo_EcrireEtLire"),
                      ("demo-lire-saisie-manuelle", "Demo_LireSaisieManuelle"),
                      ("demo-verifier-existence-fsic", "Demo_VerifierExistenceFSIC"),
                      ("demo-extraction-vers-feuille", "Demo_ExtractionVersFeuille")):
        s[nom] = sous.add_parser(nom, help=aide)
        s[nom].add_argument("classeur", help="classeur (remplace ThisWorkbook ; enregistré en place)")
    ia = s["interagir-mainframe"]
    ia.add_argument("--ligne", default=None, help="ligne 1 à 32 (InputBox ; vide = annuler)")
    ia.add_argument("--colonne", default=None, help="colonne 1 à 80 (InputBox)")
    ia.add_argument("--longueur", default="10", help="longueur à lire (InputBox ; mode lecture)")
    ia.add_argument("--texte", default=None, help="texte à écrire (InputBox ; absent = lecture)")
    s["demo-verifier-existence-fsic"].add_argument("--isin", default="FR0010000003",
                                                   help="code ISIN à vérifier (InputBox)")
    a = p.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    if a.commande == "interagir-mainframe":
        res = interagir_mainframe(a.classeur, a.ligne, a.colonne, a.longueur, a.texte)
    elif a.commande == "demo-ecrire-et-lire":
        res = demo_ecrire_et_lire(a.classeur)
    elif a.commande == "demo-lire-saisie-manuelle":
        res = demo_lire_saisie_manuelle(a.classeur)
    elif a.commande == "demo-verifier-existence-fsic":
        res = demo_verifier_existence_fsic(a.classeur, a.isin)
    else:
        res = demo_extraction_vers_feuille(a.classeur)
    return 0 if res.ok else 1


if __name__ == "__main__":
    sys.exit(main())
