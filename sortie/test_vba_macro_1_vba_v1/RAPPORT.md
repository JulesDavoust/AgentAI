# Rapport de traduction : test_vba_macro_1_vba

## 1. Écarts de comportement avec la macro d'origine

- Classeur d'origine non modifié : la macro travaillait dans `ThisWorkbook` puis faisait `SaveCopyAs`. Le Python enregistre directement les modifications dans la copie datée (`--dossier-sortie`) et ne touche jamais au classeur source. Le contenu de la copie est le même.
- Extension de la copie : celle du classeur source (la macro écrivait toujours `.xlsm`). Un `.xlsx` en entrée donne un `.xlsx`.
- Journal (feuille `Log`) : les lignes « Filtre applique » et « Visuel regenere » sont écrites avant l'étape Excel (Excel ne reçoit que la copie déjà enregistrée). Les messages « Sauvegarde », « Mail envoye » et « Pipeline termine » ne sont pas dans la copie (comme dans la macro, où ils étaient écrits après `SaveCopyAs`) : ils vont dans `logging`.
- `Rnd()` : générateur `random.Random` ; les valeurs simulées diffèrent d'une exécution à l'autre (comme en VBA avec `Randomize`).
- `MsgBox` (fin de pipeline, région inconnue, erreur) : remplacés par `logging`. Une erreur du pipeline est journalisée, la fonction renvoie `None` et le programme sort avec le code 1.
- Aucun déclenchement de mail avec `--mode-mail aucun` (option ajoutée par la convention d'interface).

## 2. Correspondance des modules

| Module VBA | Fichier Python | Procédure VBA | Fonction Python |
|---|---|---|---|
| marco_25_10 | marco_25_10.py | LancerPipelineComplet | lancer_pipeline_complet |
| marco_25_10 | marco_25_10.py | EnsureStructure | ensure_structure |
| marco_25_10 | marco_25_10.py | AjouterDonneesDuJour / LigneExiste | ajouter_donnees_du_jour / ligne_existe |
| marco_25_10 | marco_25_10.py | SimulerVolume / SimulerTemperature | simuler_volume / simuler_temperature |
| marco_25_10 | marco_25_10.py | AppliquerFiltre / EstRegionValide | choisir_filtre / est_region_valide ; étape Excel `appliquer_filtre_excel` |
| marco_25_10 | connecteurs/excel.py | CreerVisuel | creer_visuel_excel |
| marco_25_10 | marco_25_10.py | SauvegarderVersion | chemin_sauvegarde |
| marco_25_10 | marco_25_10.py + connecteurs/outlook.py | EnvoyerMail | envoyer_mail_suivi / outlook.envoyer_mail |
| marco_25_10 | marco_25_10.py | GetOrCreateSheet / LogMsg | get_or_create_sheet / log_msg |
| ThisWorkbook, Feuil1 à Feuil4 | (aucun) | modules vides | — |

## 3. Entrées attendues

`classeur` : remplace `ThisWorkbook`. Fichier `.xlsx` ou `.xlsm` avec idéalement les feuilles `Feuille2` (tableau `tblDonnees` : Date, Region, Volume_GWh, Temperature_C, Statut), `Feuille1` (visuel) et `Log` ; les feuilles et le tableau manquants sont créés avec 14 jours d'historique simulé. Pas de valeur par défaut.

`--dossier-sortie` : remplace `SaveFolder()`. Dossier où la copie datée est créée (créé si absent). Défaut : `sortie`.

`--sheet-visuel` : remplace `SHEET_VISUEL`, feuille du visuel. Défaut : `Feuille1`.

`--sheet-data` : remplace `SHEET_DATA`, feuille des données. Défaut : `Feuille2`.

`--sheet-log` : remplace `SHEET_LOG`, feuille du journal. Défaut : `Log`.

`--table-name` : remplace `TABLE_NAME`, nom du tableau structuré. Défaut : `tblDonnees`.

`--file-title` : remplace `FILE_TITLE`, début du nom de la copie. Défaut : `VBA_Natran_Test`.

`--file-suffix` : remplace `FILE_SUFFIX`, fin du nom de la copie. Défaut : `VBATF`.

`--mail-to` : remplace `MAIL_TO`, destinataire. Défaut : `jdavoust201@headmind.com`.

`--mail-cc` : remplace `MAIL_CC`, copie. Défaut : vide.

`--mode-mail` : remplace `SEND_MAIL` (`True` = `envoyer`, `False` = `afficher`) ; `aucun` n'envoie rien. Défaut : `envoyer` (valeur du VBA).

`--saisie` : remplace l'`InputBox` (variable `saisie`) : région à filtrer (Nord, Sud, Est, Ouest, Centre, casse libre) ; vide = aucun filtre. Défaut : vide.

## 4. Commandes

<!-- vba2py:debut commandes -->
```powershell
pip install -r requirements.txt

# Tests sur données fictives
python tests/generer_donnees_fictives.py
python -m pytest -q

# Liste complète des options
python marco_25_10.py --help

# Essai sur le classeur fictif (aucune action externe)
python marco_25_10.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie_test --mode-mail aucun

# Sur un vrai classeur (chemins complets ; mail affiché sans être envoyé)
python marco_25_10.py "C:\chemin\vers\classeur.xlsm" --dossier-sortie "C:\chemin\vers\sorties" --mode-mail afficher
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--dossier-sortie` | SaveFolder() | `sortie` |
| `--sheet-visuel` | SHEET_VISUEL | `Feuille1` |
| `--sheet-data` | SHEET_DATA | `Feuille2` |
| `--sheet-log` | SHEET_LOG | `Log` |
| `--table-name` | TABLE_NAME | `tblDonnees` |
| `--file-title` | FILE_TITLE | `VBA_Natran_Test` |
| `--file-suffix` | FILE_SUFFIX | `VBATF` |
| `--mail-to` | MAIL_TO | `jdavoust201@headmind.com` |
| `--mail-cc` | MAIL_CC | — |
| `--mode-mail` | SEND_MAIL | `envoyer` |
| `--saisie` | InputBox (saisie) | — |
<!-- vba2py:fin commandes -->

## 5. Choix de traduction

- `ListObjects.Add` / `ListRows.Add` → `openpyxl.worksheet.table.Table`, `ref` étendue à chaque ligne ajoutée.
- `Date`, `Now` → `datetime` via une `horloge` injectable (tests déterministes) ; `CLng(date)` → comparaison de dates.
- `Rnd`, `Round` → `random.Random`, `round` (arrondi bancaire, comme VBA).
- `Application.ScreenUpdating` / `DisplayAlerts` : supprimés (aucun équivalent nécessaire).
- `AutoFilter` (état réellement appliqué dans la copie), `AutoFit`, tableau croisé `TCD_Volume`, graphique `xlLine` croisé, formules KPI calculées : exécutés par Excel via `bibliotheque.excel_com` (`connecteurs/excel.py`), instruction par instruction, en une seule session sur la copie enregistrée.
- Nettoyage de la feuille visuel : `preparer_feuille` côté openpyxl, puis les mêmes instructions dans l'étape Excel.
- `Outlook.Application` : `connecteurs/outlook.py` (import paresseux de `win32com`), pièce jointe en chemin absolu, `Send` / `Display` selon `--mode-mail`.
- `On Error GoTo MailErr` → `try/except` autour du seul envoi ; `On Error GoTo ErrHandler` → `try/except` global qui journalise.
- `Trim$` → `strip(" ")` ; `StrComp(vbTextCompare)` → `casefold` ; `=` entre chaînes (sensible à la casse) → `==`.
- Prérequis : Windows, Excel installé et fichier fermé dans Excel, Outlook installé et connecté au compte. Pour tester sans rien envoyer : `--mode-mail afficher` ou `aucun`.

## 6. TODO et fonctionnalités non couvertes

Aucun.

## 7. Éléments suspects

- Adresse e-mail codée en dur : `jdavoust201@headmind.com` (`MAIL_TO`, commentaire « A modifier avec mail HDP »). Conservée comme valeur par défaut de `--mail-to`.
- Chemin personnel codé en dur : `C:\Users\Jules\Documents\Jules\Projects\AgentAI\examples\sortiesVBA\` (`SaveFolder()`). Non repris : le défaut de `--dossier-sortie` est `sortie`.
- Envoi automatique : `SEND_MAIL = True`, donc la macro (et le Python par défaut) envoie réellement le mail, sans confirmation. Utiliser `--mode-mail afficher` pour tester.
- Aucune instruction adressée à un agent trouvée dans le VBA.

## 8. Comparer avec la macro d'origine dans Excel

1. Ouvrir `tests/fixtures/classeur_complet.xlsx` dans Excel, l'enregistrer en `.xlsm`, y importer `_source_vba/marco_25_10.bas`.
2. Dans le VBA, mettre `SEND_MAIL = False` et un `SaveFolder()` de test, puis lancer `LancerPipelineComplet` ; saisir `Sud` dans l'InputBox.
3. Fermer le classeur, puis lancer : `python marco_25_10.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie --mode-mail afficher --saisie Sud`.
4. Comparer la copie VBA et la copie Python : tableau `tblDonnees` (5 lignes « Phase de test » ajoutées, valeurs aléatoires donc différentes), filtre sur la région, KPI `B2`, `E2`, `H2` de `Feuille1`, tableau croisé `TCD_Volume` en `A5`, graphique en courbes en `J5`, feuille `Log`.
5. Refaire avec `tests/fixtures/classeur_vide.xlsx` : le tableau doit être créé avec 14 jours d'historique (70 lignes) puis 5 lignes du jour.

## 9. Tests

<!-- vba2py:debut tests -->
Vérification indépendante par vba2py : `python -m pytest -q tests` → ✅ **37 passed in 0.20s**
<!-- vba2py:fin tests -->

## 10. Extraction

<!-- vba2py:debut extraction -->
- Origine : export
- Dossier de sortie : `test_vba_macro_1_vba`

| Module | Type | Lignes de code | Statut |
|---|---|---|---|
| Feuil1 | classe/document | 0 | vide (ignoré) |
| Feuil2 | classe/document | 0 | vide (ignoré) |
| Feuil3 | classe/document | 0 | vide (ignoré) |
| Feuil4 | classe/document | 0 | vide (ignoré) |
| marco_25_10 | module | 288 | extrait |
| ThisWorkbook | classe/document | 0 | vide (ignoré) |
<!-- vba2py:fin extraction -->
