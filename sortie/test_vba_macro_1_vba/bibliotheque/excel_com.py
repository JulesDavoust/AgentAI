"""Piloter Excel depuis Python pour ce qu'openpyxl ne sait pas faire.

Module fourni par vba2py et copié tel quel dans chaque projet traduit : ne pas le modifier.

Principe : le code VBA qui utilise une fonction absente d'openpyxl (tableau croisé dynamique,
graphique croisé, segments, recalcul des formules, export PDF, impression...) est traduit
**instruction par instruction** vers le même modèle objet Excel, depuis Python. Ce module fournit
seulement la partie commune et délicate :

- `session_excel(chemin)` : ouvre le classeur dans une instance d'Excel séparée et invisible,
  sans exécuter ses macros, enregistre puis referme tout, même en cas d'erreur ; si Excel refuse
  une instruction, `ErreurExcel` indique laquelle ;
- `excel_disponible()` : vrai sous Windows avec Excel installé (tests sur Excel réel) ;
- `xl` : les constantes Excel (`xl.xlRowField`, `xl.xlSum`...) ;
- `EnregistreurCOM` : un faux objet Excel pour les tests, qui note chaque instruction reçue ;
- `preparer_feuille(ws)` : côté openpyxl, retire tableaux croisés et graphiques avant qu'Excel
  ne les recrée.

Exemple (VBA → Python) :
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, "tblDonnees")
    Set pt = pc.CreatePivotTable(wsV.Range("A5"), "TCD_Volume")
    pt.PivotFields("Date").Orientation = xlRowField
→
    with session_excel(chemin) as wb:
        pc = wb.PivotCaches().Create(xl.xlDatabase, "tblDonnees")
        pt = pc.CreatePivotTable(wb.Worksheets("Feuille1").Range("A5"), "TCD_Volume")
        pt.PivotFields("Date").Orientation = xl.xlRowField

Règles de traduction des appels :
- une collection s'appelle avec des parenthèses : `wb.PivotCaches()`, `ws.ChartObjects()` ;
  un élément : `ws.ChartObjects(1)`, `pt.PivotFields("Date")` ;
- une méthode sans argument garde ses parenthèses : `wb.Save()`, `rng.Clear()` ;
- arguments nommés VBA → arguments nommés Python identiques (`Periods:=...` → `Periods=...`) ;
- `Array(True, False)` → tuple `(True, False)`.

Prérequis à l'exécution : Windows, Excel installé, pywin32 ; le fichier ne doit pas être ouvert
dans Excel.
"""

from __future__ import annotations

import datetime
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Constantes Excel courantes (valeurs du modèle objet, identiques sous VBA). Les autres sont
# lues dans la bibliothèque de types d'Excel quand elle est disponible.
CONSTANTES = {
    # Tableaux croisés
    "xlDatabase": 1, "xlExternal": 2, "xlConsolidation": 3,
    "xlHidden": 0, "xlRowField": 1, "xlColumnField": 2, "xlPageField": 3, "xlDataField": 4,
    "xlSum": -4157, "xlCount": -4112, "xlAverage": -4106, "xlMax": -4136, "xlMin": -4139,
    "xlProduct": -4149, "xlCountNums": -4113, "xlStDev": -4155, "xlStDevP": -4156,
    "xlVar": -4164, "xlVarP": -4165,
    "xlNormal": -4143, "xlPercentOfTotal": 8, "xlPercentOfRow": 6, "xlPercentOfColumn": 7,
    "xlPercentOf": 5, "xlDifferenceFrom": 2, "xlPercentDifferenceFrom": 4, "xlRunningTotal": 5,
    "xlCompactRow": 0, "xlTabularRow": 1, "xlOutlineRow": 2,
    "xlPivotTableVersion12": 3, "xlPivotTableVersion14": 4, "xlPivotTableVersion15": 5,
    "xlAscending": 1, "xlDescending": 2, "xlManual": -4135,
    # Graphiques
    "xlLine": 4, "xlLineMarkers": 65, "xlColumnClustered": 51, "xlColumnStacked": 52,
    "xlBarClustered": 57, "xlBarStacked": 58, "xlPie": 5, "xlDoughnut": -4120, "xlArea": 1,
    "xlXYScatter": -4169, "xlXYScatterLines": 74, "xlRows": 1, "xlColumns": 2,
    "xlCategory": 1, "xlValue": 2, "xlLegendPositionBottom": -4107, "xlLegendPositionRight": -4152,
    # Plages, tableaux, recherche
    "xlUp": -4162, "xlDown": -4121, "xlToLeft": -4159, "xlToRight": -4161,
    "xlYes": 1, "xlNo": 2, "xlGuess": 0, "xlSrcRange": 1,
    "xlCellTypeVisible": 12, "xlCellTypeBlanks": 4, "xlCellTypeConstants": 2, "xlCellTypeFormulas": -4123,
    "xlPasteValues": -4163, "xlPasteFormats": -4122, "xlPasteAll": -4104,
    "xlWhole": 1, "xlPart": 2, "xlValues": -4163, "xlFormulas": -4123,
    "xlAnd": 1, "xlOr": 2, "xlFilterValues": 7,
    # Mise en forme conditionnelle
    "xlCellValue": 1, "xlExpression": 2, "xlBetween": 1, "xlEqual": 3, "xlNotEqual": 4,
    "xlGreater": 5, "xlLess": 6, "xlGreaterEqual": 7, "xlLessEqual": 8,
    # Calcul, export, impression
    "xlCalculationAutomatic": -4105, "xlCalculationManual": -4135,
    "xlTypePDF": 0, "xlTypeXPS": 1, "xlQualityStandard": 0,
    "xlPortrait": 1, "xlLandscape": 2,
    # Formats de fichier (SaveAs)
    "xlOpenXMLWorkbook": 51, "xlOpenXMLWorkbookMacroEnabled": 52, "xlCSV": 6, "xlCSVUTF8": 62,
}


class _Constantes:
    def __getattr__(self, nom: str):
        if nom in CONSTANTES:
            return CONSTANTES[nom]
        try:
            from win32com.client import constants  # disponible après EnsureDispatch
            return getattr(constants, nom)
        except Exception:
            raise AttributeError(
                f"Constante Excel inconnue : {nom}. Utilise sa valeur numérique (documentation "
                f"Microsoft) avec un commentaire, ex. `-4157  # xlSum`."
            ) from None


xl = _Constantes()


# --------------------------------------------------------------------------- session Excel


def ouvrir_excel():
    """Instance d'Excel séparée et invisible (ne touche pas aux classeurs déjà ouverts)."""
    import win32com.client  # import paresseux : le reste du projet s'importe sans pywin32

    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        # Liaison précoce : arguments nommés et constantes de la bibliothèque de types
        excel = win32com.client.gencache.EnsureDispatch(excel)
    except Exception:
        pass
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE  # pas de Workbook_Open
    return excel


@contextmanager
def session_excel(chemin_classeur, *, enregistrer: bool = True, excel=None):
    """Ouvre `chemin_classeur` dans Excel et fournit le classeur (`Workbook`), tracé.

    Chaque instruction envoyée à Excel est notée : si Excel en refuse une, l'erreur levée
    (`ErreurExcel`) indique laquelle, avec une explication du code d'erreur.
    À la sortie du bloc : enregistrement (si `enregistrer`), fermeture du classeur, puis de
    l'instance d'Excel si elle a été créée ici. En cas d'erreur, rien n'est enregistré mais
    tout est refermé.
    `excel` : instance déjà ouverte, ou faux objet dans les tests.
    """
    chemin = str(Path(chemin_classeur).resolve())  # Excel ne connaît pas le dossier courant de Python
    if not Path(chemin).exists():
        raise FileNotFoundError(chemin)
    instance_creee = excel is None
    excel = excel if excel is not None else ouvrir_excel()
    trace: list[str] = []
    wb = None
    try:
        wb = TraceurCOM(excel, "excel", trace).Workbooks.Open(chemin)
        yield wb
        if enregistrer:
            wb.Save()
    finally:
        if wb is not None:
            _deballer(wb).Close(False)
        if instance_creee:
            excel.Quit()


def excel_disponible() -> bool:
    """Vrai si Excel peut être piloté ici (Windows + pywin32 + Excel installé)."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------- traceur

# Codes d'erreur COM fréquents (scode) et leur explication
ERREURS_COM = {
    -2146827284: "Excel a refusé l'opération (0x800A03EC) : nom de feuille, de tableau ou de champ "
                 "introuvable, argument ou plage invalide, ou opération impossible dans l'état actuel "
                 "du classeur.",
    -2147352570: "Nom inconnu (DISP_E_UNKNOWNNAME) : propriété ou méthode mal orthographiée.",
    -2147352573: "Membre introuvable (DISP_E_MEMBERNOTFOUND) : souvent une propriété appelée comme "
                 "une méthode, ou l'inverse.",
    -2147352571: "Type d'argument incorrect (DISP_E_TYPEMISMATCH).",
    -2147352561: "Argument obligatoire manquant (DISP_E_PARAMNOTOPTIONAL).",
    -2147352562: "Nombre d'arguments incorrect (DISP_E_BADPARAMCOUNT).",
    -2147417848: "Excel s'est fermé ou ne répond plus (RPC_E_DISCONNECTED).",
    -2147418111: "Excel est occupé et a rejeté l'appel (RPC_E_CALL_REJECTED) : réessayer.",
}


class ErreurExcel(RuntimeError):
    """Une instruction envoyée à Excel a échoué ; le message indique laquelle."""

    def __init__(self, instruction: str, erreur: Exception, precedentes: list[str]):
        self.instruction = instruction
        self.code = _code_com(erreur)
        explication = ERREURS_COM.get(self.code, "Code d'erreur Excel non répertorié.")
        detail = _description_com(erreur)
        lignes = [f"Excel a refusé : {instruction}", f"  {explication}"]
        if detail:
            lignes.append(f"  Message d'Excel : {detail}")
        if precedentes:
            lignes.append("  Instructions précédentes :")
            lignes += [f"    {p}" for p in precedentes]
        super().__init__("\n".join(lignes))


def _code_com(erreur: Exception) -> int | None:
    args = getattr(erreur, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6 and args[2][5]:
        return args[2][5]  # scode détaillé (excepinfo)
    return args[0] if args and isinstance(args[0], int) else None


def _description_com(erreur: Exception) -> str:
    args = getattr(erreur, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2])
    return ""


def _est_erreur_com(erreur: Exception) -> bool:
    return type(erreur).__name__ == "com_error"


def _deballer(valeur):
    if isinstance(valeur, TraceurCOM):
        return object.__getattribute__(valeur, "_objet")
    if isinstance(valeur, tuple):
        return tuple(_deballer(v) for v in valeur)
    if isinstance(valeur, list):
        return [_deballer(v) for v in valeur]
    return valeur


def _formater(valeur) -> str:
    if isinstance(valeur, TraceurCOM):
        return object.__getattribute__(valeur, "_chemin")
    return repr(valeur)


_VALEURS_SIMPLES = (str, bytes, int, float, bool, type(None), tuple, list, dict, datetime.date, datetime.time)


class TraceurCOM:
    """Enveloppe un objet Excel réel et note chaque instruction (utilisé par `session_excel`)."""

    def __init__(self, objet, chemin: str, trace: list[str]):
        object.__setattr__(self, "_objet", objet)
        object.__setattr__(self, "_chemin", chemin)
        object.__setattr__(self, "_trace", trace)

    def _envelopper(self, valeur, chemin: str):
        # Valeurs simples (texte, nombre, date, tableau de valeurs) : renvoyées telles quelles ;
        # tout le reste (objets et méthodes Excel) reste tracé.
        if isinstance(valeur, _VALEURS_SIMPLES):
            return valeur
        return TraceurCOM(valeur, chemin, self._trace)

    def _executer(self, instruction: str, action):
        self._trace.append(instruction)
        try:
            return action()
        except Exception as erreur:
            if _est_erreur_com(erreur):
                raise ErreurExcel(instruction, erreur, self._trace[-6:-1]) from erreur
            raise

    def __getattr__(self, nom: str):
        if nom.startswith("__"):
            raise AttributeError(nom)
        chemin = f"{self._chemin}.{nom}"
        try:
            valeur = getattr(self._objet, nom)
        except Exception as erreur:
            if _est_erreur_com(erreur):
                raise ErreurExcel(chemin, erreur, self._trace[-5:]) from erreur
            raise
        return self._envelopper(valeur, chemin)

    def __setattr__(self, nom: str, valeur):
        instruction = f"{self._chemin}.{nom} = {_formater(valeur)}"
        self._executer(instruction, lambda: setattr(self._objet, nom, _deballer(valeur)))

    def __call__(self, *args, **kwargs):
        arguments = [_formater(a) for a in args] + [f"{k}={_formater(v)}" for k, v in kwargs.items()]
        instruction = f"{self._chemin}({', '.join(arguments)})"
        resultat = self._executer(
            instruction,
            lambda: self._objet(*_deballer(args), **{k: _deballer(v) for k, v in kwargs.items()}),
        )
        return self._envelopper(resultat, instruction)

    def __iter__(self):
        for i, element in enumerate(self._objet):
            yield self._envelopper(element, f"{self._chemin}[{i}]")

    def __bool__(self):
        return bool(self._objet)

    def __repr__(self):
        return f"<Excel {self._chemin}>"


# --------------------------------------------------------------------------- côté openpyxl


def preparer_feuille(ws) -> None:
    """Équivalent openpyxl du nettoyage de la macro avant recréation par Excel.

    `ChartObjects(i).Delete`, `PivotTables(i).TableRange2.Clear`, `Cells.Clear` : tableaux
    croisés, graphiques, images et cellules de la feuille sont retirés.
    """
    ws._pivots = []
    ws._charts = []
    ws._images = []
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
            cell.style = "Normal"


# --------------------------------------------------------------------------- tests


class EnregistreurCOM:
    """Faux objet Excel pour les tests : chaque instruction reçue est notée dans `journal`.

    Chaque objet porte un chemin lisible, par exemple
    `wb.Worksheets('Feuille1').PivotTables('TCD').PivotFields('Date').Orientation`.

        wb = EnregistreurCOM("wb", reponses={"wb.Worksheets('Feuille1').ChartObjects().Count": 0})
        construire_visuel(wb)
        assert wb.affectations("PivotFields('Date').Orientation") == [xl.xlRowField]

    `reponses` : valeur à renvoyer pour un chemin donné (attribut lu ou résultat d'appel) ;
    par défaut un nouvel enregistreur, `Count` vaut 0 et l'itération est vide.
    """

    def __init__(self, chemin: str = "excel", reponses: dict | None = None, _journal: list | None = None):
        object.__setattr__(self, "_chemin", chemin)
        object.__setattr__(self, "_reponses", reponses or {})
        object.__setattr__(self, "journal", _journal if _journal is not None else [])

    def _enfant(self, chemin: str):
        if chemin in self._reponses:
            return self._reponses[chemin]
        if chemin.endswith(".Count"):
            return 0
        return EnregistreurCOM(chemin, self._reponses, self.journal)

    def __getattr__(self, nom: str):
        if nom.startswith("__"):
            raise AttributeError(nom)
        return self._enfant(f"{self._chemin}.{nom}")

    def __setattr__(self, nom: str, valeur: Any):
        self.journal.append(("affectation", f"{self._chemin}.{nom}", valeur))

    def __call__(self, *args, **kwargs):
        arguments = [repr(a) for a in args] + [f"{k}={v!r}" for k, v in kwargs.items()]
        chemin = f"{self._chemin}({', '.join(arguments)})"
        self.journal.append(("appel", self._chemin, (args, kwargs)))
        return self._enfant(chemin)

    def __iter__(self):
        return iter(())

    def __repr__(self):
        return f"<{self._chemin}>"

    # ---- lecture du journal

    def appels(self, fin_de_chemin: str) -> list[tuple]:
        """Arguments (args, kwargs) des appels dont le chemin se termine par `fin_de_chemin`."""
        return [v for t, c, v in self.journal if t == "appel" and c.endswith(fin_de_chemin)]

    def affectations(self, fin_de_chemin: str) -> list:
        """Valeurs affectées aux propriétés dont le chemin se termine par `fin_de_chemin`."""
        return [v for t, c, v in self.journal if t == "affectation" and c.endswith(fin_de_chemin)]

    def instructions(self) -> list[str]:
        """Le journal sous forme lisible, dans l'ordre (utile pour comparer à la macro)."""
        lignes = []
        for t, c, v in self.journal:
            if t == "affectation":
                lignes.append(f"{c} = {v!r}")
            else:
                args, kwargs = v
                arguments = [repr(a) for a in args] + [f"{k}={x!r}" for k, x in kwargs.items()]
                lignes.append(f"{c}({', '.join(arguments)})")
        return lignes
