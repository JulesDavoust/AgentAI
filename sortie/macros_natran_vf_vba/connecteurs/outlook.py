"""Mails Outlook (VBA ``CreateObject("Outlook.Application")``). Seul endroit où win32com est importé."""

from __future__ import annotations

from pathlib import Path


def envoyer_mail(destinataire: str, cc: str, sujet: str, corps: str, pieces_jointes, mode: str = "envoyer") -> None:
    """VBA ``EnvoyerMail`` : ``CreateItem(0)``, ``.To/.CC/.Subject/.Body``, ``.Attachments.Add``,
    puis ``.Send`` (mode ``envoyer``) ou ``.Display`` (mode ``afficher``).

    Prérequis : Windows, Outlook installé et connecté au compte. Les chemins sont rendus absolus.
    """
    import win32com.client  # import paresseux

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = destinataire
    mail.CC = cc
    mail.Subject = sujet
    mail.Body = corps
    for chemin in pieces_jointes:
        mail.Attachments.Add(str(Path(chemin).resolve()))
    if mode == "envoyer":
        mail.Send()
    else:
        mail.Display()
