"""MsgBox et InputBox de VBA, sous forme de questions dans le terminal.

Module fourni par vba2py et copié tel quel dans chaque projet traduit : ne pas le modifier.

    choix = msgbox("Continuer ?", vbYesNo + vbQuestion, "Titre")      # VBA : MsgBox(...)
    if choix == vbNo: ...
    saisie = inputbox("Region ?", "Filtre", "Nord")                   # VBA : InputBox(...)
    n = application_inputbox("Nombre ?", type_=1)                     # VBA : Application.InputBox(Type:=1)

Comportement, identique à VBA :
- `msgbox` sans choix (vbOKOnly) : affiche le message et attend Entrée (comme le clic sur OK) ;
- `msgbox` avec choix : redemande tant que la réponse n'est pas l'un des boutons proposés ;
  renvoie la constante du bouton (vbYes, vbNo, vbCancel...) ;
- `inputbox` : Entrée seule renvoie la valeur par défaut (comme OK sans modifier le champ) ;
  « Annuler » (taper `!annuler`, ou fin de saisie Ctrl+Z / Ctrl+D) renvoie "" ;
- `application_inputbox` : redemande tant que la saisie ne correspond pas au type demandé ;
  « Annuler » renvoie False.

Tests : `with simuler(["oui", "Nord"]) as journal:` fournit les réponses dans l'ordre, sans
aucune saisie au clavier, et `journal` contient chaque dialogue affiché (texte, titre, réponse).
"""

from __future__ import annotations

import sys
from contextlib import contextmanager
from dataclasses import dataclass, field

# Boutons et icônes (valeurs VBA)
vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel = 0, 1, 2, 3, 4, 5
vbCritical, vbQuestion, vbExclamation, vbInformation = 16, 32, 48, 64
vbDefaultButton1, vbDefaultButton2, vbDefaultButton3 = 0, 256, 512
# Valeurs renvoyées
vbOK, vbCancel, vbAbort, vbRetry, vbIgnore, vbYes, vbNo = 1, 2, 3, 4, 5, 6, 7

ANNULER = "!annuler"
_BOUTONS = {
    vbOKOnly: [("ok", vbOK)],
    vbOKCancel: [("ok", vbOK), ("annuler", vbCancel)],
    vbAbortRetryIgnore: [("abandonner", vbAbort), ("reessayer", vbRetry), ("ignorer", vbIgnore)],
    vbYesNoCancel: [("oui", vbYes), ("non", vbNo), ("annuler", vbCancel)],
    vbYesNo: [("oui", vbYes), ("non", vbNo)],
    vbRetryCancel: [("reessayer", vbRetry), ("annuler", vbCancel)],
}
_SYNONYMES = {"o": "oui", "y": "oui", "yes": "oui", "n": "non", "no": "non", "a": "annuler",
              "cancel": "annuler", "r": "reessayer", "retry": "reessayer", "i": "ignorer",
              "ignore": "ignorer", "abort": "abandonner"}


class AucuneReponsePrevue(RuntimeError):
    """En mode simulé, un dialogue a été affiché alors que la liste de réponses est épuisée."""


@dataclass
class Dialogue:
    genre: str  # "msgbox", "inputbox" ou "application_inputbox"
    texte: str
    titre: str
    reponse: object


@dataclass
class _Simulation:
    reponses: list
    journal: list[Dialogue] = field(default_factory=list)


_simulation: _Simulation | None = None


@contextmanager
def simuler(reponses: list):
    """Réponses fournies à l'avance (tests) ; un MsgBox sans choix n'en consomme pas."""
    global _simulation
    precedente = _simulation
    _simulation = _Simulation(list(reponses))
    try:
        yield _simulation.journal
    finally:
        _simulation = precedente


def _lire(invite: str) -> str | None:
    """Une ligne saisie au clavier, ou une réponse prévue ; None = fin de saisie (Annuler)."""
    if _simulation is not None:
        if not _simulation.reponses:
            raise AucuneReponsePrevue(f"aucune réponse prévue pour : {invite.strip()}")
        return _simulation.reponses.pop(0)
    try:
        return input(invite)
    except EOFError:
        return None


def _afficher(texte: str, titre: str) -> None:
    if _simulation is None:
        print()
        if titre:
            print(f"[{titre}]")
        print(texte)


def msgbox(texte, boutons: int = vbOKOnly, titre: str = "") -> int:
    """VBA `MsgBox` : renvoie vbOK, vbYes, vbNo, vbCancel..."""
    texte, titre = str(texte), str(titre or "")
    _afficher(texte, titre)
    choix = _BOUTONS.get(boutons & 7, _BOUTONS[vbOKOnly])
    if len(choix) == 1:
        if _simulation is None and sys.stdin.isatty():
            _lire("[Entrée = OK] ")
        if _simulation is not None:
            _simulation.journal.append(Dialogue("msgbox", texte, titre, vbOK))
        return vbOK
    noms = [nom for nom, _ in choix]
    invite = " / ".join(noms) + " ? "
    while True:
        saisie = _lire(invite)
        if saisie is None:  # fin de saisie : comme fermer la fenêtre
            valeur = dict(choix).get("annuler", choix[-1][1])
            break
        saisie = _SYNONYMES.get(saisie.strip().lower(), saisie.strip().lower())
        if saisie in noms:
            valeur = dict(choix)[saisie]
            break
        if _simulation is not None:
            raise ValueError(f"réponse prévue « {saisie} » impossible pour ce MsgBox (choix : {noms})")
        print(f"Répondre par : {', '.join(noms)}")
    if _simulation is not None:
        _simulation.journal.append(Dialogue("msgbox", texte, titre, valeur))
    return valeur


def inputbox(invite, titre: str = "", defaut="") -> str:
    """VBA `InputBox` : Entrée seule = valeur par défaut ; Annuler = ""."""
    invite, titre, defaut = str(invite), str(titre or ""), "" if defaut is None else str(defaut)
    _afficher(invite, titre)
    indication = f" (Entrée = « {defaut} », {ANNULER} = Annuler) : " if defaut else f" ({ANNULER} = Annuler) : "
    saisie = _lire(indication)
    if saisie is None or saisie.strip() == ANNULER:
        valeur = ""
    elif saisie == "":
        valeur = defaut
    else:
        valeur = saisie
    if _simulation is not None:
        _simulation.journal.append(Dialogue("inputbox", invite, titre, valeur))
    return valeur


def application_inputbox(invite, titre: str = "", defaut="", type_: int = 2):
    """VBA `Application.InputBox(..., Type:=n)` : 1 nombre, 2 texte, 4 booléen ; Annuler = False."""
    invite, titre = str(invite), str(titre or "")
    _afficher(invite, titre)
    defaut = "" if defaut is None else str(defaut)
    while True:
        saisie = _lire(f" (Entrée = « {defaut} », {ANNULER} = Annuler) : " if defaut else f" ({ANNULER} = Annuler) : ")
        if saisie is None or saisie.strip() == ANNULER:
            valeur = False
            break
        texte = defaut if saisie == "" else saisie
        try:
            valeur = _convertir(texte, type_)
            break
        except ValueError:
            if _simulation is not None:
                raise
            print("Valeur incorrecte, recommencer.")
    if _simulation is not None:
        _simulation.journal.append(Dialogue("application_inputbox", invite, titre, valeur))
    return valeur


def _convertir(texte: str, type_: int):
    if type_ & 1:  # nombre (virgule ou point décimal, comme Excel en français)
        nombre = float(texte.strip().replace(" ", "").replace(" ", "").replace(",", "."))
        return int(nombre) if nombre.is_integer() else nombre
    if type_ & 4:
        if texte.strip().lower() in {"vrai", "true"}:
            return True
        if texte.strip().lower() in {"faux", "false"}:
            return False
        raise ValueError(texte)
    return texte
