# Correspondances VBA → Python

Librairies autorisées dans le code produit : **bibliothèque standard Python, openpyxl, pandas,
matplotlib, pytest**, les modules fournis dans `bibliotheque/`, et **pywin32** uniquement dans le
dossier `connecteurs/` (voir « Intégrations externes »). Interdits partout : `xlwings`,
`subprocess`, `eval`/`exec`, et tout accès réseau en dehors de `connecteurs/`.

## Le classeur de la macro (`ThisWorkbook`) : toujours dans Excel

**Règle** : tout ce que la macro lit ou modifie dans `ThisWorkbook` / `ActiveWorkbook` (feuilles,
cellules, tableaux, formats, graphiques, tableaux croisés, enregistrement) passe **par Excel**,
avec le **même modèle objet** que la macro, instruction par instruction :

```python
from bibliotheque.excel_com import session_classeur, xl, rgb, en_datetime, pour_excel

def lancer_pipeline_complet(classeur, ..., sans_enregistrer=False, ouvrir_classeur=session_classeur):
    with ouvrir_classeur(classeur, enregistrer=not sans_enregistrer) as wb:   # wb = ThisWorkbook
        ensure_structure(wb)
        ...
```

- Si le classeur est **déjà ouvert dans Excel**, la commande agit directement dessus, comme la
  macro (l'utilisateur voit le résultat, saisit dans les cellules, enregistre lui-même). S'il est
  **fermé**, il est ouvert, modifié puis **enregistré** (option `--sans-enregistrer` pour ne pas
  l'enregistrer). C'est ce qui rend la traduction fidèle : feuilles, Log et historique restent
  dans le classeur d'un lancement à l'autre, et rien n'est perdu (boutons, macros, objets
  qu'openpyxl ne sait pas conserver).
- Chaque procédure VBA devient une fonction Python qui reçoit `wb` (et non un chemin), dans le
  module Python de la macro. Les calculs purs (simulations, validations, construction de
  textes, HTML, CSV) sont dans des fonctions séparées, testables sans Excel.
- `session_classeur` est un **paramètre injectable** (`ouvrir_classeur`) : les tests passent un
  `EnregistreurCOM` ou une fausse session.
- **openpyxl ne sert plus qu'aux tests** : créer les classeurs fictifs et relire les résultats.
  Ne jamais écrire `ThisWorkbook` avec openpyxl (une macro qui lit un classeur fermé **autre**
  que le sien peut le lire avec openpyxl).

| VBA | Python (Excel) |
|---|---|
| `ThisWorkbook`, `ActiveWorkbook` | `wb` fourni par `session_classeur` |
| `Workbooks.Open chemin` (autre classeur) | `autre = wb.Application.Workbooks.Open(str(Path(chemin).resolve()))`, refermé comme dans la macro |
| `ThisWorkbook.SaveCopyAs chemin` | `wb.SaveCopyAs(str(Path(chemin).resolve()))` (chemin complet obligatoire) |
| `ThisWorkbook.Save`, `SaveAs` | `wb.Save()`, `wb.SaveAs(chemin_absolu, format)` |
| `Worksheets("X")`, `Sheets(1)`, `Worksheets.Add(After:=...)` | `wb.Worksheets("X")`, `wb.Sheets(1)`, `wb.Worksheets.Add(After=...)` (index 1 comme en VBA) |
| `Range("A1").Value`, `Cells(l, c).Value`, `.Formula`, `.NumberFormat` | `ws.Range("A1").Value`, `ws.Cells(l, c).Value`, `.Formula` (en anglais, Excel calcule), `.NumberFormat` |
| `Range("A1:C1").Value = Array(a, b, c)` | `ws.Range("A1:C1").Value = (a, b, c)` ; plusieurs lignes : tuple de tuples |
| `Cells(Rows.Count, 1).End(xlUp).Row` | `ws.Cells(ws.Rows.Count, 1).End(xl.xlUp).Row` |
| `ListObjects("t")`, `.ListRows.Add`, `ListObjects.Add(xlSrcRange, ...)` | mêmes appels : `ws.ListObjects("t")`, `.ListRows.Add()`, `ws.ListObjects.Add(xl.xlSrcRange, ...)` |
| `AutoFilter`, `Sort`, `AutoFit`, `WorksheetFunction.X` | mêmes appels (`wb.Application.WorksheetFunction.Sum(...)`) |
| Mise en forme (`Font.Bold`, `Interior.Color = RGB(...)`, `ColumnWidth`) | mêmes propriétés ; `RGB(r, v, b)` → `rgb(r, v, b)` |
| `ChartObjects`, `PivotCaches`, `PivotTables`, `Shapes`, `Hyperlinks.Add` | mêmes appels |
| `ws.Activate`, `Range.Select`, `ActiveWindow.DisplayGridlines`, `ActiveWindow.Zoom` | mêmes appels (`wb.Application.ActiveWindow.Zoom = 100`) : visibles quand le classeur est ouvert |
| `ws.Visible = xlSheetHidden` | `ws.Visible = xl.xlSheetHidden` |
| `Application.ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, `Calculation` | mêmes réglages sur `wb.Application`, remis comme la macro le fait |
| Date lue dans une cellule / écrite dans une cellule | `en_datetime(valeur)` / `ws.Cells(l, c).Value = pour_excel(une_date)` |
| `MsgBox` / `InputBox` | voir les lignes ci-dessous (`bibliotheque.dialogues`) |

Traduction des appels : collection avec parenthèses (`wb.Worksheets("X")`, `ws.ChartObjects()`),
élément avec son index ou son nom, méthode sans argument avec `()` (`.Delete()`, `.Clear()`),
propriété sans parenthèses (`.Value`, `.Count`), arguments nommés identiques (`After:=x` →
`After=x`), `Array(...)` → tuple, constantes `xlUp` → `xl.xlUp`.

## Autres correspondances

| VBA | Python |
|---|---|
| Graphique exporté hors du classeur (image jointe à un mail, PNG) | `matplotlib`, export PNG dans `--dossier-sortie` (`plt.savefig`, puis `plt.close`) |
| `MsgBox texte, boutons, titre` | `from bibliotheque.dialogues import msgbox, vbYesNo, vbQuestion, vbYes...` puis `msgbox(texte, boutons, titre)` : même texte, mêmes boutons, même titre, même valeur renvoyée. Un `MsgBox` d'information reste un `msgbox` (il attend la validation comme en VBA), pas un simple log |
| `InputBox(invite, titre, defaut)` | `inputbox(invite, titre, defaut)` : Entrée = valeur par défaut, Annuler = `""`, comme VBA. Si l'interface prévoit une option pour cette saisie (`--saisie`) et qu'elle est fournie, sa valeur est utilisée **à la place** de la question |
| `Application.InputBox(..., Type:=n)` | `application_inputbox(invite, titre, defaut, type_=n)` : même validation du type, Annuler = `False` |
| Boucle de validation autour d'un `InputBox` (redemander tant que la saisie est invalide) | la même boucle en Python, avec les mêmes `msgbox` d'erreur |
| `UserForm` | pas encore pris en charge : `# TODO(vba2py)` et section « Écarts » |
| `Debug.Print` | `logging.debug` |
| `Workbook_Open`, `Worksheet_Change` et autres événements | fonctions nommées (`au_demarrage_classeur()`, `sur_modification_feuille(cellules)`) appelées explicitement ; signaler dans le rapport qu'il n'y a plus de déclenchement automatique |
| `Application.Run "Macro"` / `Application.Run "Module.Macro"` (macro **du même classeur**) | simple appel Python de la fonction traduite (import si elle est dans un autre module) ; si le nom est construit dynamiquement, un dictionnaire `{nom VBA: fonction}` |
| `Application.Run "Autre.xlsm!Macro"` (macro d'**un autre classeur**) | jamais exécuté : `# TODO(vba2py)` + section « Écarts » |
| `ThisWorkbook.FollowHyperlink chemin` (ouvrir un fichier ou une page) | `os.startfile(str(Path(chemin).resolve()))` dans `connecteurs/systeme.py` ; fonction injectable, bloquée dans les tests (fixture `autouse`) ; une constante comme `OPEN_MAP` devient une option (`--open-map`) |

## Fichiers, dates, chaînes

| VBA | Python |
|---|---|
| `Dir`, `FileSystemObject`, `MkDir`, `Kill`, `FileCopy` | `pathlib.Path` (`glob`, `mkdir`, `unlink`), `shutil.copy` |
| `Open chemin For Output As #f` + `Print #f, ligne` | `open(chemin, "w", encoding="cp1252", newline="\r\n")` puis `f.write(ligne + "\n")` : VBA écrit en **Windows-1252** avec des fins de ligne **Windows**, même si le contenu annonce `charset='utf-8'` (le reproduire et le signaler dans le rapport). `For Append` → mode `"a"` |
| `Open chemin For Input As #f` + `Line Input #f, ligne` | `open(chemin, encoding="cp1252", newline="")` et lecture ligne par ligne sans le saut de ligne ; `EOF(f)` → fin de l'itération |
| `Write #f, ...` | valeurs séparées par des virgules, textes entre guillemets, dates `#aaaa-mm-jj#` : reproduire exactement |
| `Date`, `Now`, `DateAdd`, `DateDiff`, `DateSerial` | `datetime`, `timedelta` |
| Dates stockées en numéro de série Excel | `datetime(1899, 12, 30) + timedelta(days=n)` |
| `Format(x, "0.00")`, `Format(d, "dd/mm/yyyy")` | f-strings / `strftime` |
| `Left`, `Mid`, `Right`, `Len`, `InStr` | tranches de chaînes ; **`Mid` et `InStr` sont 1-indexés, `InStr` renvoie 0 si absent** |
| `Like` | `fnmatch.fnmatchcase` ou `re` |
| `Split`, `Join`, `Replace`, `Trim`, `UCase` | `str.split`, `str.join`, `str.replace`, `str.strip`, `str.upper` |
| `Scripting.Dictionary`, `Collection` | `dict`, `list` (attention : `Dictionary` peut être insensible à la casse avec `CompareMode`) |
| `Type ... End Type`, modules de classe | `@dataclass`, classes Python |

## Intégrations externes (Outlook, Word, Access…)

Une macro qui pilote une autre application Office doit garder **le même comportement réel**.
Ne la remplace jamais par une simulation (fichier .eml, log, brouillon) sans le dire.

| VBA | Python |
|---|---|
| `CreateObject("Outlook.Application")`, `.CreateItem(0)` | `win32com.client.Dispatch("Outlook.Application")`, `.CreateItem(0)`, dans `connecteurs/outlook.py` |
| `.To`, `.CC`, `.Subject`, `.Body`, `.HTMLBody`, `.Attachments.Add` | mêmes propriétés et méthodes sur l'objet COM |
| `.Send` / `.Display` | `.Send()` / `.Display()` : les deux modes doivent exister |

Règles :
- Tout appel COM ou réseau est isolé dans `connecteurs/<application>.py`. C'est le seul endroit
  où `win32com` est importé, et **à l'intérieur des fonctions** (import paresseux), pour que le
  reste du code s'importe sous Linux et dans les tests.
- Le module métier ne fait jamais l'appel lui-même : il reçoit une fonction `envoyer_mail`
  en paramètre, dont la valeur par défaut est celle du connecteur. Les tests passent une fausse
  fonction qui enregistre les appels.
- Les valeurs codées en dur dans le VBA (destinataires, constante `SEND_MAIL`…) deviennent des
  options de la ligne de commande (`--destinataire`, `--mode-mail envoyer|afficher|aucun`). Leur
  valeur par défaut reproduit le VBA.
- **Tout chemin transmis à une autre application est absolu** : `str(Path(chemin).resolve())`.
  Outlook, Excel ou Word tournent dans leur propre processus et ne connaissent pas le dossier
  courant de Python (erreur « Ce chemin d'accès n'existe pas »). Le VBA utilisait des chemins
  complets ; le Python doit faire de même, même si l'utilisateur passe un chemin relatif.
- `requirements.txt` contient alors `pywin32; sys_platform == "win32"`.
- Le rapport indique les prérequis (Windows, Outlook installé et connecté au compte) et comment
  tester sans rien envoyer (`--mode-mail afficher`).

## Excel piloté depuis Python : règles communes

- Exceptions (jamais exécutées) : `Application.Run` vers un autre classeur, `SendKeys`, `Shell`
  → `# TODO(vba2py)` et section « Écarts ».
- Si Excel refuse une instruction, la session lève `ErreurExcel`, dont le message cite
  l'instruction refusée, le message d'Excel et les objets utilisés : ne l'avale pas, journalise
  `str(erreur)` en entier (et voir « Quand une instruction traduite échoue »).
- `conftest.py` bloque `bibliotheque.excel_com.ouvrir_excel` et
  `bibliotheque.excel_com.trouver_classeur_ouvert` (fixture `autouse`, qui renvoie `None` pour
  le second), sauf pour les tests marqués `excel_reel`.
- `requirements.txt` contient `pywin32; sys_platform == "win32"` ; le rapport indique le
  prérequis (Windows, Excel installé) et les deux modes (classeur ouvert / fermé).

## Pièges sémantiques à respecter

- **`Val`, `IsNumeric`, `CDate`, `Format` et les réglages régionaux** :
  - `Val("12,5")` vaut 12 et `Val("1 234")` vaut 1234 : `Val` lit **toujours** le point comme
    séparateur décimal, ignore les espaces et s'arrête au premier caractère invalide
    (`Val("abc")` = 0). Écrire une fonction `val()` qui reproduit ce comportement ;
  - `IsNumeric` et `CDate` dépendent des réglages régionaux de Windows (ici français, séparateur
    décimal `","`) : `IsNumeric("12,5")` est vrai, c'est pourquoi certaines macros remplacent le
    point par `Application.International(xlDecimalSeparator)` avant de tester ;
    `CDate("05/10/2026")` est le **5 octobre**. Reproduire ces règles explicitement (ne pas
    dépendre de la locale de la machine Python) ;
  - `Format(d, "dd/mm/yyyy")` → `d.strftime("%d/%m/%Y")` ; `Format(x, "0.00")` donne un point
    décimal en code mais une virgule à l'affichage français : reproduire ce que la macro écrit
    réellement dans le fichier ou la cellule.
- **`Err.Raise vbObjectError + n, , message`** : une exception Python dédiée (classe nommée) avec
  le même message ; `Err.Number` / `Err.Description` → attributs de l'exception.
- **Variables de module** (`Private ecran As Worksheet` en tête de module) : en VBA elles gardent
  leur valeur entre deux macros tant que le classeur reste ouvert. En Python, chaque commande
  repart de zéro : regroupe cet état dans une classe, et si une macro dépend d'une valeur laissée
  par une macro précédente, signale-le dans « Écarts ».
- **Code dupliqué entre modules** (`LogMsg`, `GetOrCreateSheet`... recopiés dans chaque
  module) : ne les regroupe dans un module commun que s'ils sont **strictement identiques** ;
  sinon chaque module garde sa version. Indique le regroupement dans « Choix de traduction ».
- **`ByRef` est le mode par défaut en VBA** : si une procédure modifie un argument, la version Python doit retourner la nouvelle valeur (ou modifier un objet mutable) et l'appelant doit l'utiliser.
- **Tableaux** : `Dim t(1 To 10)` ou `Option Base 1` → adapter les indices ; `UBound`/`LBound` → `len`.
- **Division** : `/` donne un décimal, `\` est la division entière (`//` mais tronquée vers zéro pour les négatifs : utiliser `int(a / b)`), `Mod` garde le signe du dividende (`math.fmod`).
- **Arrondis** : `Round` et `CInt`/`CLng` font un arrondi bancaire (comme `round()` Python) ; `Int` arrondit vers le bas, `Fix` tronque vers zéro.
- **Entiers** : `Integer` VBA est sur 16 bits (dépassement à 32 767) — signaler si le code pourrait en dépendre.
- **Comparaisons de chaînes** : `Option Compare Text` → comparaisons insensibles à la casse.
- **`Empty`, `Null`, `Nothing`, `Variant`** : une cellule vide vaut `None` (lue par Excel ou openpyxl) ; `""` et `0` ne sont pas équivalents à `None` en Python.
- **Gestion d'erreurs** : `On Error Resume Next` → `try/except` ciblé autour de l'instruction concernée ; `On Error GoTo Label` → `try/except/finally`. Ne jamais avaler toutes les erreurs silencieusement.
- **Booléens** : `True` vaut `-1` en VBA ; rare, mais à signaler si le code fait de l'arithmétique dessus.
