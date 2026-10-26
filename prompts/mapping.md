# Correspondances VBA → Python

Seules ces librairies sont autorisées dans le code produit : **bibliothèque standard Python,
openpyxl, pandas, matplotlib, streamlit, pytest**, et **pywin32** uniquement dans le dossier
`connecteurs/` (voir « Intégrations externes »).
Interdits partout : `xlwings` (Excel n'est plus utilisé, openpyxl le remplace), `subprocess`,
`eval`/`exec`, et tout accès réseau en dehors de `connecteurs/`.

## Objets Excel

| VBA | Python |
|---|---|
| `Workbooks.Open`, `ThisWorkbook`, `ActiveWorkbook` | `openpyxl.load_workbook(chemin)` ; le chemin est **un paramètre** de la fonction, jamais un classeur « actif » implicite |
| `Workbook.Save` / `SaveAs` | `wb.save(chemin)` (chemin en paramètre) |
| `Worksheets("X")`, `Sheets(1)` | `wb["X"]`, `wb.worksheets[0]` (attention : index 0 en Python) |
| `Range("A1").Value`, `Cells(l, c).Value` | `ws["A1"].value`, `ws.cell(row=l, column=c).value` (lignes/colonnes restent 1-indexées dans openpyxl) |
| `Range("A1").Formula` | `ws["A1"] = "=SOMME(...)"` → écrire la formule **en anglais** (`=SUM(...)`) ; openpyxl ne calcule pas les formules |
| `Cells(Rows.Count, 1).End(xlUp).Row` | dernière ligne non vide calculée explicitement (boucle ou `ws.max_row` puis remontée) |
| `UsedRange`, `CurrentRegion` | `ws.iter_rows(...)` / `pandas.read_excel(..., sheet_name=...)` |
| Boucles sur des plages, filtres, tris, `WorksheetFunction.Sum/VLookup/Match/CountIf` | `pandas` (`DataFrame`, `merge`, `groupby`, `loc`) ou Python pur pour les cas simples |
| `AutoFilter`, `Sort` | `DataFrame.query` / `sort_values`, puis réécriture avec openpyxl si le résultat doit rester dans le classeur |
| Mise en forme (`Font`, `Interior.Color`, `NumberFormat`) | `openpyxl.styles` (`Font`, `PatternFill`, `cell.number_format`) |
| `ChartObjects.Add`, graphiques dans une feuille | **vrai graphique Excel** avec `openpyxl.chart` (`LineChart`, `BarChart`... + `Reference` sur les cellules sources), au même emplacement que dans la macro (`ws.add_chart(chart, "J5")`), avec le même titre et le même type. Ne jamais remplacer un graphique de feuille par une image |
| Graphique exporté hors du classeur (image jointe à un mail, PNG) | `matplotlib`, export PNG dans `--dossier-sortie` (`plt.savefig`, puis `plt.close`) |
| `UserForm`, `InputBox` | **Streamlit** (`st.text_input`, `st.selectbox`, `st.button`...). La logique métier reste dans des fonctions pures, testables sans Streamlit |
| `MsgBox` (information) | `logging.info` ou valeur de retour ; `MsgBox` avec choix (vbYesNo) → paramètre booléen de la fonction |
| `Debug.Print` | `logging.debug` |
| `Application.ScreenUpdating`, `Calculation`, `DisplayAlerts`, `EnableEvents` | à supprimer (aucun équivalent nécessaire) — le noter dans le rapport |
| `Workbook_Open`, `Worksheet_Change` et autres événements | fonctions nommées (`au_demarrage_classeur()`, `sur_modification_feuille(cellules)`) appelées explicitement ; signaler dans le rapport qu'il n'y a plus de déclenchement automatique |
| `Application.Run "Macro"` / `Application.Run "Module.Macro"` (macro **du même classeur**) | simple appel Python de la fonction traduite (import si elle est dans un autre module) ; si le nom est construit dynamiquement, un dictionnaire `{nom VBA: fonction}` |
| `Application.Run "Autre.xlsm!Macro"` (macro d'**un autre classeur**) | jamais exécuté : `# TODO(vba2py)` + section « Écarts » |
| `ws.Activate`, `Range.Select`, `ActiveWindow.DisplayGridlines`, `ActiveWindow.Zoom` | état dans lequel la macro laisse le classeur, enregistré avec openpyxl : `wb.active = ws`, `ws.sheet_view.selection[0].activeCell` / `.sqref`, `ws.sheet_view.showGridLines`, `ws.sheet_view.zoomScale` (pour la feuille concernée) |
| `ws.Visible = xlSheetHidden` / `xlSheetVeryHidden` / `xlSheetVisible` | `ws.sheet_state = "hidden"` / `"veryHidden"` / `"visible"` (voir aussi « Feuilles masquées » dans `EXTRACTION.md`) |
| `ws.Hyperlinks.Add Anchor:=cell, Address:=..., TextToDisplay:=...` | `cell.hyperlink = adresse` ; `cell.value = texte` ; `cell.style = "Hyperlink"` |
| `ws.Hyperlinks.Delete` | `cell.hyperlink = None` pour chaque cellule qui en porte un |
| `ThisWorkbook.FollowHyperlink chemin` (ouvrir un fichier ou une page) | `os.startfile(str(Path(chemin).resolve()))` dans `connecteurs/systeme.py` ; fonction injectable, bloquée dans les tests (fixture `autouse`) ; une constante comme `OPEN_MAP` devient une option (`--open-map`) |

## Fichiers, dates, chaînes

| VBA | Python |
|---|---|
| `Dir`, `FileSystemObject`, `MkDir`, `Kill`, `FileCopy` | `pathlib.Path` (`glob`, `mkdir`, `unlink`), `shutil.copy` |
| `Open chemin For Output As #f` + `Print #f, ligne` | `open(chemin, "w", encoding="cp1252", newline="\r\n")` puis `f.write(ligne + "\n")` : VBA écrit en **Windows-1252** avec des fins de ligne **Windows**, même si le contenu annonce `charset='utf-8'` (le reproduire et le signaler dans le rapport). `For Append` → mode `"a"` |
| `Open chemin For Input As #f` + `Line Input #f, ligne` | `open(chemin, encoding="cp1252", newline="")` et lecture ligne par ligne sans le saut de ligne ; `EOF(f)` → fin de l'itération |
| `Write #f, ...` | valeurs séparées par des virgules, textes entre guillemets, dates `#aaaa-mm-jj#` : reproduire exactement |
| `Date`, `Now`, `DateAdd`, `DateDiff`, `DateSerial` | `datetime`, `timedelta` |
| Dates stockées en numéro de série Excel | `openpyxl` les convertit déjà ; sinon `datetime(1899, 12, 30) + timedelta(days=n)` |
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

## Ce qu'openpyxl ne sait pas faire : Excel piloté depuis Python

**Règle** : openpyxl par défaut. Quand une instruction VBA n'a pas d'équivalent **fidèle** en
openpyxl, elle est exécutée par Excel lui-même : traduis-la **instruction par instruction**,
avec le même modèle objet Excel, dans `connecteurs/excel.py`, en utilisant
`bibliotheque.excel_com` (`session_excel`, `xl`, `preparer_feuille`). Jamais de substitut
approximatif : pas de valeurs figées calculées avec pandas à la place d'un tableau croisé, pas
d'image matplotlib à la place d'un graphique Excel, pas de formule laissée non calculée quand la
macro relit son résultat.

Cas typiques (liste non exhaustive : applique la règle à tout le reste) :

| VBA | Pourquoi Excel |
|---|---|
| `PivotCaches`, `PivotTables`, `PivotFields`, `AddDataField`, `CalculatedFields`, `Group`, `CurrentPage`, `SlicerCaches` | openpyxl ne crée ni ne modifie les tableaux croisés |
| Graphique d'un tableau croisé (`SetSourceData pt.TableRange1`), `Chart.Export`, `CopyPicture` | dépend du tableau croisé, ou rendu fait par Excel |
| Relire une valeur calculée par une formule (`Range("B2").Value` après `.Formula`), `Calculate` | openpyxl ne calcule pas les formules |
| `ExportAsFixedFormat` (PDF), `PrintOut`, `PageSetup` avancé | rendu et impression |
| `RefreshAll`, `QueryTables`, `ListObject.QueryTable`, connexions de données | actualisation de sources externes |
| `GoalSeek`, Solveur | moteur de calcul d'Excel |
| `AutoFilter` / `Sort` dont l'état doit rester réellement appliqué dans le classeur, `AutoFit` exact | openpyxl n'écrit que des critères, ne masque/trie/mesure rien lui-même |

Exceptions (jamais exécutées, même via Excel) : `Application.Run` vers un autre classeur,
`SendKeys`, `Shell` → `# TODO(vba2py)` et section « Écarts ».

Organisation :
- `connecteurs/excel.py` contient une fonction par étape VBA concernée (ex.
  `creer_visuel_excel(wb, ...)`), qui reçoit le `Workbook` Excel et reprend les instructions dans
  le même ordre que la macro, plus une fonction d'entrée qui ouvre la session :
  `with session_excel(chemin) as wb: creer_visuel_excel(wb, ...)`. Une seule session par fichier :
  regroupe les étapes Excel consécutives.
- Ordre du pipeline : étapes openpyxl → `wb.save(chemin)` → session Excel sur ce fichier → étapes
  suivantes (si une étape openpyxl suit, recharge le fichier avec `load_workbook`). Si la macro
  crée le tableau puis fait `SaveCopyAs`, c'est la copie enregistrée qui passe dans Excel.
- Le nettoyage d'une feuille qu'Excel va remplir (`ChartObjects(i).Delete`,
  `PivotTables(i).TableRange2.Clear`, `Cells.Clear`) se fait avec `preparer_feuille(ws)` côté
  openpyxl ; les mêmes instructions peuvent aussi être traduites dans l'étape Excel.
- Traduction des appels : collection avec parenthèses (`wb.PivotCaches()`, `ws.ChartObjects()`),
  élément avec son index ou son nom (`pt.PivotFields("Date")`), méthode sans argument avec `()`
  (`.Delete()`, `.Clear()`), arguments nommés identiques (`Periods:=` → `Periods=`),
  `Array(...)` → tuple, constantes `xlRowField` → `xl.xlRowField`.
- La fonction d'entrée est un **paramètre injectable** de la fonction métier (comme
  `envoyer_mail`) ; `conftest.py` bloque `excel_com.ouvrir_excel` (fixture `autouse`), sauf pour
  les tests marqués `excel_reel`.
- Si Excel refuse une instruction, `session_excel` lève `ErreurExcel`, dont le message cite
  l'instruction refusée et explique le code d'erreur : ne l'avale pas, journalise `str(erreur)`
  en entier.
- `requirements.txt` contient `pywin32; sys_platform == "win32"` ; le rapport indique le
  prérequis (Windows, Excel installé, fichier fermé dans Excel).

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
- **`Empty`, `Null`, `Nothing`, `Variant`** : une cellule vide vaut `None` en openpyxl ; `""` et `0` ne sont pas équivalents à `None` en Python.
- **Gestion d'erreurs** : `On Error Resume Next` → `try/except` ciblé autour de l'instruction concernée ; `On Error GoTo Label` → `try/except/finally`. Ne jamais avaler toutes les erreurs silencieusement.
- **Booléens** : `True` vaut `-1` en VBA ; rare, mais à signaler si le code fait de l'arithmétique dessus.
