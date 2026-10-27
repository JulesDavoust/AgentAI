"""Connecteur Outlook (VBA : EnvoyerMail, late binding CreateObject("Outlook.Application")).

Prérequis : Windows, Outlook installé et connecté au compte, pywin32.
"""

from __future__ import annotations

from pathlib import Path


def envoyer_mail(destinataire: str, cc: str, objet: str, corps: str,
                 piece_jointe: str | Path, envoyer: bool) -> None:
    """VBA EnvoyerMail : crée le mail, joint le fichier, puis `.Send` ou `.Display`.

    `envoyer=True` -> `.Send()` ; `False` -> `.Display()` (brouillon affiché).
    La pièce jointe est transmise à Outlook avec un chemin absolu.
    """
    import win32com.client  # import paresseux : seul endroit où win32com est utilisé pour Outlook

    ol = win32com.client.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)
    mail.To = destinataire
    mail.CC = cc
    mail.Subject = objet
    mail.Body = corps
    mail.Attachments.Add(str(Path(piece_jointe).resolve()))
    if envoyer:
        mail.Send()
    else:
        mail.Display()
