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
| `PivotCaches.Create(...)`, `CreatePivotTable`, `PivotFields(...).Orientation`, `AddDataField` | `from bibliotheque.tableau_croise import assurer_tableau_croise` : conserve le tableau croisé s'il existe déjà, sinon le crée (interactif, recalculé par Excel à l'ouverture). Voir la docstring du module pour la correspondance des paramètres. **Jamais** de tableau de valeurs calculé avec pandas à la place |
| Graphique d'un tableau croisé (`SetSourceData pt.TableRange1`) | `assurer_graphique_croise(wb, resultat, ancre=..., titre=..., type_graphique="ligne"/"barres", largeur_pt=..., hauteur_pt=..., nom_fichier=<nom du classeur>)` : vrai graphique croisé dynamique, conservé s'il existe |
| `ws.Cells.Clear`, `ChartObjects(i).Delete`, `PivotTables(i).TableRange2.Clear` sur une feuille qui porte un tableau croisé | `effacer_hors_tableaux_croises(ws)` : le tableau et son graphique ne sont pas supprimés (ils sont conservés et recalculés). Ne jamais écrire dans la plage d'un tableau croisé, ni vider `ws._pivots` ou `ws._charts` à la main |
| `UserForm`, `InputBox` | **Streamlit** (`st.text_input`, `st.selectbox`, `st.button`...). La logique métier reste dans des fonctions pures, testables sans Streamlit |
| `MsgBox` (information) | `logging.info` ou valeur de retour ; `MsgBox` avec choix (vbYesNo) → paramètre booléen de la fonction |
| `Debug.Print` | `logging.debug` |
| `Application.ScreenUpdating`, `Calculation`, `DisplayAlerts`, `EnableEvents` | à supprimer (aucun équivalent nécessaire) — le noter dans le rapport |
| `Workbook_Open`, `Worksheet_Change` et autres événements | fonctions nommées (`au_demarrage_classeur()`, `sur_modification_feuille(cellules)`) appelées explicitement ; signaler dans le rapport qu'il n'y a plus de déclenchement automatique |
| `Application.Run "Autre.xlsm!Macro"` | import Python si le module est traduit dans la même sortie, sinon `# TODO` + rapport |

## Fichiers, dates, chaînes

| VBA | Python |
|---|---|
| `Dir`, `FileSystemObject`, `MkDir`, `Kill`, `FileCopy` | `pathlib.Path` (`glob`, `mkdir`, `unlink`), `shutil.copy` |
| `Open ... For Input/Output` | `open(..., encoding=...)` ou `csv` |
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

## Pièges sémantiques à respecter

- **Tableaux croisés** : si `assurer_tableau_croise` lève `FormeNonPriseEnCharge` (plusieurs champs
  en lignes, plusieurs valeurs...), ne la contourne pas : `# TODO(vba2py)` et section « Écarts ».
  Le projet utilisant `bibliotheque/`, `requirements.txt` fixe `openpyxl==3.1.5`.
- **`ByRef` est le mode par défaut en VBA** : si une procédure modifie un argument, la version Python doit retourner la nouvelle valeur (ou modifier un objet mutable) et l'appelant doit l'utiliser.
- **Tableaux** : `Dim t(1 To 10)` ou `Option Base 1` → adapter les indices ; `UBound`/`LBound` → `len`.
- **Division** : `/` donne un décimal, `\` est la division entière (`//` mais tronquée vers zéro pour les négatifs : utiliser `int(a / b)`), `Mod` garde le signe du dividende (`math.fmod`).
- **Arrondis** : `Round` et `CInt`/`CLng` font un arrondi bancaire (comme `round()` Python) ; `Int` arrondit vers le bas, `Fix` tronque vers zéro.
- **Entiers** : `Integer` VBA est sur 16 bits (dépassement à 32 767) — signaler si le code pourrait en dépendre.
- **Comparaisons de chaînes** : `Option Compare Text` → comparaisons insensibles à la casse.
- **`Empty`, `Null`, `Nothing`, `Variant`** : une cellule vide vaut `None` en openpyxl ; `""` et `0` ne sont pas équivalents à `None` en Python.
- **Gestion d'erreurs** : `On Error Resume Next` → `try/except` ciblé autour de l'instruction concernée ; `On Error GoTo Label` → `try/except/finally`. Ne jamais avaler toutes les erreurs silencieusement.
- **Booléens** : `True` vaut `-1` en VBA ; rare, mais à signaler si le code fait de l'arithmétique dessus.
