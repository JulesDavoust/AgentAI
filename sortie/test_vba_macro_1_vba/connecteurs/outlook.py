"""Connecteur Outlook : seul endroit où win32com est importé pour les mails."""

from __future__ import annotations

from pathlib import Path

MODES = ("envoyer", "afficher")


def envoyer_mail(destinataire: str, cc: str, sujet: str, corps: str,
                 pieces_jointes: list[str], mode: str = "envoyer") -> None:
    """VBA EnvoyerMail : CreateObject("Outlook.Application"), CreateItem(0), .Send ou .Display."""
    if mode not in MODES:
        raise ValueError(f"Mode de mail inconnu : {mode!r} (attendu : {', '.join(MODES)})")
    import win32com.client  # import paresseux : Windows + Outlook uniquement

    ol = win32com.client.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)
    mail.To = destinataire
    mail.CC = cc
    mail.Subject = sujet
    mail.Body = corps
    for pj in pieces_jointes:
        mail.Attachments.Add(str(Path(pj).resolve()))  # Outlook ne connaît pas le dossier courant
    if mode == "envoyer":
        mail.Send()
    else:
        mail.Display()
