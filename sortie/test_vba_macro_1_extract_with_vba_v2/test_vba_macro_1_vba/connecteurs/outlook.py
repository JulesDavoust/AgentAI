"""Connecteur Outlook (COM) : seul endroit du projet qui importe win32com."""

from typing import Sequence


def envoyer_mail(
    destinataire: str,
    cc: str,
    objet: str,
    corps: str,
    pieces_jointes: Sequence[str],
    mode: str = "envoyer",
) -> None:
    """Équivalent de CreateObject("Outlook.Application") + CreateItem(0) de EnvoyerMail (marco_25_10.bas).

    mode : "envoyer" -> .Send() ; "afficher" -> .Display() (brouillon visible, rien n'est envoyé).
    Prérequis : Windows, Outlook installé et connecté au compte.
    """
    if mode not in ("envoyer", "afficher"):
        raise ValueError(f"Mode de mail inconnu : {mode!r} (attendu : envoyer ou afficher)")
    import win32com.client  # import paresseux : le reste du projet s'importe sans pywin32

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = destinataire
    mail.CC = cc
    mail.Subject = objet
    mail.Body = corps
    for pj in pieces_jointes:
        mail.Attachments.Add(str(pj))
    if mode == "envoyer":
        mail.Send()
    else:
        mail.Display()
