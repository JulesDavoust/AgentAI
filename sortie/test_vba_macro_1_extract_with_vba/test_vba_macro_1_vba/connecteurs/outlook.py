"""Connecteur Outlook (COM). Seul endroit où win32com est importé."""
from pathlib import Path
from typing import Iterable

MODES_MAIL = ("envoyer", "afficher", "aucun")


def envoyer_mail(
    destinataire: str,
    cc: str,
    sujet: str,
    corps: str,
    pieces_jointes: Iterable[Path | str] = (),
    mode: str = "envoyer",
) -> None:
    """Équivalent de ``EnvoyerMail`` (VBA) : CreateObject("Outlook.Application"),
    CreateItem(0), To/CC/Subject/Body, Attachments.Add puis .Send ou .Display.

    mode : "envoyer" (.Send) ou "afficher" (.Display). Prérequis : Windows,
    Outlook installé et connecté au compte.
    """
    if mode not in ("envoyer", "afficher"):
        raise ValueError(f"Mode de mail inconnu : {mode!r}")
    import win32com.client  # import paresseux : Windows uniquement

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = destinataire
    mail.CC = cc
    mail.Subject = sujet
    mail.Body = corps
    for pj in pieces_jointes:
        mail.Attachments.Add(str(pj))
    if mode == "envoyer":
        mail.Send()
    else:
        mail.Display()
