"""Référence VBA : exécuter la macro d'origine sur un classeur fictif, sans effet réel.

Option `--reference-vba` de `traduire` (désactivée par défaut). vba2py lance la macro **d'origine**
dans une instance d'Excel séparée et invisible, sur une copie d'un classeur fictif du projet, et
enregistre le résultat comme **référence** : l'agent compare ensuite sa traduction à ce que fait
réellement la macro sur les mêmes données.

Avant l'exécution, le code VBA est **neutralisé** (le code d'origine n'est pas modifié, seule la
copie importée dans le classeur fictif l'est) :
- `MsgBox` / `InputBox` : aucune fenêtre (elle bloquerait l'Excel invisible) ; le texte est noté
  dans le journal et la réponse vient d'une liste fournie à l'avance ;
- `CreateObject` / `GetObject` d'Outlook, Word, PowerPoint, Access : refusé par une erreur VBA
  (aucun mail ne part) — la macro suit alors sa propre gestion d'erreur ;
- `FollowHyperlink`, `Shell`, `SendKeys` : rien n'est ouvert ni exécuté, l'appel est noté ;
- chemins absolus écrits dans le code (`"C:\\..."`, `"\\\\serveur\\..."`) : redirigés vers un dossier
  du bac à sable ; tout ce que la macro y écrit est recopié dans la référence.

Prérequis : Windows, Excel, pywin32, et l'option Excel « Accès approuvé au modèle d'objet du
projet VBA » (pour importer les modules dans le classeur fictif).
"""

from __future__ import annotations

import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

DELAI_MACRO = 300  # secondes
NOM_HARNAIS = "vba2py_Harnais"
APPLICATIONS_NEUTRALISEES = ("outlook.", "word.", "powerpoint.", "access.")

# --------------------------------------------------------------------------- neutralisation du code


_CHAINE = re.compile(r'"(?:[^"]|"")*"')
_CHEMIN_ABSOLU = re.compile(r'^"(?P<chemin>(?:[A-Za-z]:\\|\\\\)(?:[^"]|"")*)"$')

REMPLACEMENTS = [
    (re.compile(r"\bApplication\.InputBox\b", re.IGNORECASE), "vba2py_InputBox"),
    (re.compile(r"(?<![\w.])InputBox\$?(?=\W)", re.IGNORECASE), "vba2py_InputBox"),
    (re.compile(r"(?<![\w.])MsgBox(?=\W)", re.IGNORECASE), "vba2py_MsgBox"),
    (re.compile(r"(?<![\w.])(?:VBA\.)?CreateObject(?=\s*\()", re.IGNORECASE), "vba2py_CreateObject"),
    (re.compile(r"(?<![\w.])(?:VBA\.)?GetObject(?=\s*\()", re.IGNORECASE), "vba2py_GetObject"),
    (re.compile(r"(?<![\w])[\w.]*\.FollowHyperlink\b", re.IGNORECASE), "vba2py_FollowHyperlink"),
    (re.compile(r"(?<![\w.])Shell(?=\s*[\s(])", re.IGNORECASE), "vba2py_Shell"),
    (re.compile(r"(?<![\w.])(?:Application\.)?SendKeys\b", re.IGNORECASE), "vba2py_SendKeys"),
]


def chemin_bac_a_sable(chemin_vba: str, dossier_chemins: Path) -> str:
    """« C:\\Users\\x\\data\\ » → « <bac à sable>\\chemins\\C\\Users\\x\\data\\ »."""
    if chemin_vba.startswith("\\\\"):
        relatif = "UNC\\" + chemin_vba[2:]
    else:
        relatif = chemin_vba[0].upper() + "\\" + chemin_vba[3:]
    return str(dossier_chemins).rstrip("\\/") + "\\" + relatif


def _debut_commentaire(ligne: str) -> int:
    """Position de l'apostrophe qui ouvre un commentaire (hors chaînes), ou -1."""
    dans_chaine = False
    for i, c in enumerate(ligne):
        if c == '"':
            dans_chaine = not dans_chaine  # "" dans une chaîne : deux bascules, sans effet
        elif c == "'" and not dans_chaine:
            return i
    return -1


def neutraliser(code: str, dossier_chemins: Path) -> tuple[str, list[str]]:
    """Code VBA neutralisé et liste des modifications faites (pour le journal).

    Les chaînes de caractères et les commentaires ne sont jamais modifiés, sauf les chaînes qui
    sont des chemins absolus, redirigées vers le bac à sable.
    """
    modifications: list[str] = []
    lignes = []
    for ligne in code.replace("\r\n", "\n").split("\n"):
        if ligne.lstrip().startswith("Attribute ") or ligne.lstrip().upper().startswith("REM "):
            lignes.append(ligne)
            continue
        fin_code = _debut_commentaire(ligne)
        code_ligne, commentaire = (ligne, "") if fin_code < 0 else (ligne[:fin_code], ligne[fin_code:])
        morceaux = []
        position = 0
        for chaine in _CHAINE.finditer(code_ligne):
            morceaux.append(_neutraliser_code(code_ligne[position:chaine.start()], modifications))
            m = _CHEMIN_ABSOLU.match(chaine.group(0))
            if m:
                nouveau = chemin_bac_a_sable(m.group("chemin").replace('""', '"'), dossier_chemins)
                modifications.append(f"chemin redirigé : {m.group('chemin')} → {nouveau}")
                morceaux.append('"' + nouveau.replace('"', '""') + '"')
            else:
                morceaux.append(chaine.group(0))
            position = chaine.end()
        morceaux.append(_neutraliser_code(code_ligne[position:], modifications))
        lignes.append("".join(morceaux) + commentaire)
    return "\r\n".join(lignes), modifications


def _neutraliser_code(fragment: str, modifications: list[str]) -> str:
    for motif, remplacement in REMPLACEMENTS:
        fragment, n = motif.subn(remplacement, fragment)
        if n:
            modifications.append(f"{remplacement} ({n})")
    return fragment


# --------------------------------------------------------------------------- module VBA du harnais

HARNAIS_DEBUT = r'''Attribute VB_Name = "vba2py_Harnais"
Option Explicit
' Module ajouté par vba2py dans la copie de travail : remplace les dialogues et les appels externes.

Private vba2py_Journal As String
Private vba2py_Reponses As Variant
Private vba2py_Position As Long

Public Sub vba2py_Noter(ByVal genre As String, ByVal texte As String)
    Dim f As Integer
    f = FreeFile
    Open vba2py_Journal For Append As #f
    Print #f, genre & vbTab & Replace(Replace(texte, vbCr, " "), vbLf, " ")
    Close #f
End Sub

Private Function vba2py_ReponseSuivante(ByRef trouvee As Boolean) As String
    trouvee = False
    If IsArray(vba2py_Reponses) Then
        If vba2py_Position <= UBound(vba2py_Reponses) Then
            vba2py_ReponseSuivante = vba2py_Reponses(vba2py_Position)
            vba2py_Position = vba2py_Position + 1
            trouvee = True
        End If
    End If
End Function

Public Function vba2py_MsgBox(ByVal Prompt As Variant, Optional ByVal Buttons As Long = 0, _
        Optional ByVal Title As Variant, Optional ByVal HelpFile As Variant, Optional ByVal Context As Variant) As VbMsgBoxResult
    Dim choix As Long, r As String, trouvee As Boolean
    choix = Buttons And 7
    vba2py_Noter "MSGBOX", CStr(Prompt)
    Select Case choix
        Case vbOKCancel: vba2py_MsgBox = vbOK
        Case vbAbortRetryIgnore: vba2py_MsgBox = vbIgnore
        Case vbYesNoCancel, vbYesNo: vba2py_MsgBox = vbYes
        Case vbRetryCancel: vba2py_MsgBox = vbRetry
        Case Else: vba2py_MsgBox = vbOK
    End Select
    If choix <> vbOKOnly Then
        r = LCase$(Trim$(vba2py_ReponseSuivante(trouvee)))
        If trouvee Then
            Select Case r
                Case "oui", "yes": vba2py_MsgBox = vbYes
                Case "non", "no": vba2py_MsgBox = vbNo
                Case "annuler", "cancel": vba2py_MsgBox = vbCancel
                Case "ok": vba2py_MsgBox = vbOK
                Case "abandonner", "abort": vba2py_MsgBox = vbAbort
                Case "reessayer", "retry": vba2py_MsgBox = vbRetry
                Case "ignorer", "ignore": vba2py_MsgBox = vbIgnore
            End Select
        End If
        vba2py_Noter "REPONSE", CStr(vba2py_MsgBox)
    End If
End Function

Public Function vba2py_InputBox(ByVal Prompt As Variant, Optional ByVal Title As Variant, Optional ByVal Default As Variant, _
        Optional ByVal XPos As Variant, Optional ByVal YPos As Variant, Optional ByVal HelpFile As Variant, _
        Optional ByVal Context As Variant) As Variant
    Dim r As String, trouvee As Boolean
    vba2py_Noter "INPUTBOX", CStr(Prompt)
    r = vba2py_ReponseSuivante(trouvee)
    If trouvee Then
        vba2py_InputBox = r
    ElseIf IsMissing(Default) Then
        vba2py_InputBox = ""
    Else
        vba2py_InputBox = CStr(Default)
    End If
    vba2py_Noter "REPONSE", CStr(vba2py_InputBox)
End Function

Private Function vba2py_ApplicationNeutralisee(ByVal classe As String) As Boolean
    Dim prefixe As Variant
    For Each prefixe In Array("outlook.", "word.", "powerpoint.", "access.")
        If LCase$(Left$(classe, Len(prefixe))) = prefixe Then vba2py_ApplicationNeutralisee = True
    Next prefixe
End Function

Public Function vba2py_CreateObject(ByVal classe As String, Optional ByVal serveur As String = "") As Object
    If vba2py_ApplicationNeutralisee(classe) Then
        vba2py_Noter "NEUTRALISE", "CreateObject(" & classe & ")"
        Err.Raise vbObjectError + 513, "vba2py", classe & " neutralisé par vba2py (référence sans effet réel)"
    End If
    Set vba2py_CreateObject = CreateObject(classe)
End Function

Public Function vba2py_GetObject(Optional ByVal chemin As Variant, Optional ByVal classe As Variant) As Object
    If Not IsMissing(classe) Then
        If vba2py_ApplicationNeutralisee(CStr(classe)) Then
            vba2py_Noter "NEUTRALISE", "GetObject(" & classe & ")"
            Err.Raise vbObjectError + 513, "vba2py", classe & " neutralisé par vba2py (référence sans effet réel)"
        End If
        Set vba2py_GetObject = GetObject(chemin, classe)
    Else
        Set vba2py_GetObject = GetObject(chemin)
    End If
End Function

Public Sub vba2py_FollowHyperlink(ByVal Address As String, Optional ByVal SubAddress As Variant, Optional ByVal NewWindow As Variant, _
        Optional ByVal AddHistory As Variant, Optional ByVal ExtraInfo As Variant, Optional ByVal Method As Variant, _
        Optional ByVal HeaderInfo As Variant)
    vba2py_Noter "NEUTRALISE", "FollowHyperlink " & Address
End Sub

Public Function vba2py_Shell(ByVal PathName As String, Optional ByVal WindowStyle As Variant) As Double
    vba2py_Noter "NEUTRALISE", "Shell " & PathName
    vba2py_Shell = 0
End Function

Public Sub vba2py_SendKeys(ByVal Touches As String, Optional ByVal Attendre As Variant)
    vba2py_Noter "NEUTRALISE", "SendKeys " & Touches
End Sub

Public Function vba2py_Lancer(ByVal nomMacro As String, ByVal cheminJournal As String, ByVal reponses As String) As String
    vba2py_Journal = cheminJournal
    vba2py_Position = 0
    If reponses <> "" Then vba2py_Reponses = Split(reponses, Chr$(30)) Else vba2py_Reponses = Empty
    vba2py_Noter "DEBUT", nomMacro
    On Error GoTo Erreur
    Select Case nomMacro
'''

HARNAIS_FIN = r'''        Case Else
            Err.Raise vbObjectError + 514, "vba2py", "macro inconnue : " & nomMacro
    End Select
    vba2py_Noter "FIN", "macro terminée"
    vba2py_Lancer = "OK"
    Exit Function
Erreur:
    vba2py_Noter "ERREUR", Err.Number & " - " & Err.Description
    vba2py_Lancer = "ERREUR " & Err.Number & " - " & Err.Description
End Function
'''


def module_harnais(macros: list[str]) -> str:
    """Module VBA du harnais : appel direct de chaque macro (une erreur non gérée remonte au harnais)."""
    cas = []
    for macro in macros:
        if not re.fullmatch(r"[^\W\d]\w*\.[^\W\d]\w*", macro):
            raise ValueError(f"nom de macro invalide : {macro}")
        cas.append(f'        Case "{macro}"\n            {macro}\n')
    return (HARNAIS_DEBUT + "".join(cas) + HARNAIS_FIN).replace("\r\n", "\n").replace("\n", "\r\n")


# --------------------------------------------------------------------------- préparation et exécution


@dataclass
class ResultatReference:
    statut: str  # "OK", "ERREUR ..." (erreur VBA non gérée par la macro), ou "ECHEC ..." (problème d'exécution)
    dossier: Path
    journal: list[str]


def preparer(modules: list[Path], dossier_travail: Path) -> tuple[list[Path], list[str]]:
    """Copies neutralisées des modules standard (.bas, Windows-1252) + module du harnais."""
    dossier_modules = dossier_travail / "modules"
    dossier_modules.mkdir(parents=True, exist_ok=True)
    chemins = dossier_travail / "chemins"
    fichiers, notes = [], []
    for module in modules:
        code, modifications = neutraliser(module.read_text(encoding="utf-8"), chemins)
        cible = dossier_modules / module.name
        cible.write_bytes(code.encode("cp1252", errors="replace"))
        fichiers.append(cible)
        notes += [f"{module.stem} : {m}" for m in modifications]
    return fichiers, notes


def executer(
    project_dir: Path,
    macro: str,
    fixture: str,
    reponses: list[str],
    dossier_reference: Path,
    macros_connues: list[str],
    delai: int = DELAI_MACRO,
    lanceur=None,
) -> ResultatReference:
    """Exécute `macro` d'origine sur une copie de `fixture` et enregistre la référence.

    `lanceur` : fonction (paramètres -> statut), remplacée dans les tests.
    """
    sources = sorted((project_dir / "_source_vba").glob("*.bas"))
    if any(f.suffix.lower() == ".frm" for f in (project_dir / "_source_vba").glob("*")):
        raise ValueError("projet avec UserForms : la référence VBA ne les prend pas encore en charge")
    if dossier_reference.exists():
        shutil.rmtree(dossier_reference)
    dossier_reference.mkdir(parents=True)

    with tempfile.TemporaryDirectory(prefix="vba2py_reference_") as tmp:
        travail = Path(tmp)
        modules, notes = preparer(sources, travail)
        harnais = travail / "modules" / f"{NOM_HARNAIS}.bas"
        harnais.write_bytes(module_harnais(macros_connues).encode("cp1252"))
        journal = travail / "journal.txt"
        journal.write_bytes(b"")
        parametres = {
            "fixture": str((project_dir / fixture).resolve()),
            "classeur": str(travail / "classeur_reference.xlsm"),
            "modules": [str(m) for m in modules] + [str(harnais)],
            "macro": macro,
            "journal": str(journal),
            "reponses": reponses,
            "pid": str(travail / "excel.pid"),
        }
        statut = (lanceur or _lancer_processus)(parametres, delai)

        lignes = journal.read_bytes().decode("cp1252", errors="replace").splitlines()
        if Path(parametres["classeur"]).exists():
            shutil.copyfile(parametres["classeur"], dossier_reference / "classeur_apres.xlsm")
        if (travail / "chemins").exists():
            shutil.copytree(travail / "chemins", dossier_reference / "fichiers_ecrits")

    (dossier_reference / "journal.txt").write_text(
        "\n".join(["# Modifications de neutralisation"] + notes + ["", "# Journal d'exécution"] + lignes) + "\n",
        encoding="utf-8",
    )
    (dossier_reference / "resultat.json").write_text(
        json.dumps({"macro": macro, "fixture": fixture, "reponses": reponses, "statut": statut},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return ResultatReference(statut=statut, dossier=dossier_reference, journal=lignes)


def _lancer_processus(parametres: dict, delai: int) -> str:
    """Exécution dans un processus séparé : en cas de blocage, il est arrêté avec son Excel."""
    fichier = Path(parametres["journal"]).with_name("parametres.json")
    fichier.write_text(json.dumps(parametres), encoding="utf-8")
    racine = Path(__file__).resolve().parent.parent
    try:
        resultat = subprocess.run(
            [sys.executable, "-m", "vba2py.reference_vba", str(fichier)],
            cwd=racine, capture_output=True, text=True, timeout=delai,
        )
    except subprocess.TimeoutExpired:
        _arreter_excel(Path(parametres["pid"]))
        return f"ECHEC délai dépassé ({delai} s) : la macro attend peut-être une action non neutralisée"
    sortie = (resultat.stdout or "").strip().splitlines()
    if resultat.returncode != 0 or not sortie:
        return "ECHEC " + ((resultat.stderr or "").strip().splitlines() or ["erreur inconnue"])[-1]
    return sortie[-1]


def _arreter_excel(fichier_pid: Path) -> None:
    try:
        os.kill(int(fichier_pid.read_text()), signal.SIGTERM)
    except Exception:
        pass


def executer_dans_excel(parametres: dict, excel=None) -> str:
    """Partie exécutée dans le processus séparé (Windows). `excel` : faux objet dans les tests."""
    if excel is None:
        import win32com.client
        import win32process

        excel = win32com.client.DispatchEx("Excel.Application")
        try:
            _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
            Path(parametres["pid"]).write_text(str(pid))
        except Exception:
            pass
        instance_creee = True
    else:
        instance_creee = False
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # macros autorisées : on exécute la macro d'origine, neutralisée
    wb = None
    try:
        source = excel.Workbooks.Open(parametres["fixture"])
        source.SaveAs(parametres["classeur"], 52)  # xlOpenXMLWorkbookMacroEnabled
        wb = source
        try:
            composants = wb.VBProject.VBComponents
        except Exception:
            return ("ECHEC accès au projet VBA refusé : activer dans Excel « Accès approuvé au modèle "
                    "d'objet du projet VBA » (Fichier > Options > Centre de gestion de la confidentialité)")
        for module in parametres["modules"]:
            composants.Import(module)
        statut = excel.Run(f"'{Path(parametres['classeur']).name}'!{NOM_HARNAIS}.vba2py_Lancer",
                           parametres["macro"], parametres["journal"], chr(30).join(parametres["reponses"]))
        wb.Save()
        return str(statut)
    finally:
        if wb is not None:
            wb.Close(False)
        if instance_creee:
            excel.Quit()


if __name__ == "__main__":
    params = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    try:
        print(executer_dans_excel(params))
    except Exception as exc:  # erreur d'exécution (Excel, fichier) : renvoyée au processus parent
        print(f"ECHEC {type(exc).__name__} : {exc}")
