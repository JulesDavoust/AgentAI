# Rapport de traduction : macros_natran_vf_vba

## 1. Écarts de comportement avec la macro d'origine

- **Boîtes de dialogue → options de la ligne de commande.** Chaque `InputBox` devient une option (`--saisie`, `--ligne`, `--isin`...), chaque `MsgBox` d'information un message `logging`. Une saisie vide vaut « Annuler » comme en VBA. Conséquences : une saisie invalide (ex. ligne « 40 ») arrête la commande avec un message au lieu de redemander la valeur ; dans `InteragirMainframe`, le choix Lire / Écrire (MsgBox Oui/Non/Annuler) se déduit de `--texte` (absent = lecture ; fourni = écriture) et l'écriture est toujours validée par ENTREE (réponse « Oui » à la question de validation).
- **Pipelines gaz, carte et CSV : le classeur d'origine n'est jamais modifié.** En VBA, `ThisWorkbook` reste ouvert et garde ses feuilles (données, Log, visuel) tant qu'on ne le ferme pas ; seule la copie `SaveCopyAs` est enregistrée. Ici, `classeur` est lu, modifié en mémoire et seule la copie datée est écrite dans `--dossier-sortie`. Conséquences : chaque lancement repart du classeur de départ (l'historique de 14 jours est simulé de nouveau, il ne s'accumule pas d'un lancement à l'autre) et les lignes du Log écrites après la copie (« Sauvegarde », « Mail envoye », « Pipeline termine ») ne sont que journalisées (console), comme elles ne figurent pas dans la copie en VBA.
- **Commandes qui gardent un état dans le classeur** (`ouvrir-mainframe`, `mvs-*`, `interagir-mainframe`, `demo-*`, `generer-csv-*`) : le classeur passé en argument joue le rôle du classeur resté ouvert et est **enregistré en place**. Travailler sur une copie. openpyxl ne conserve pas les boutons / formes du classeur (les macros VBA d'un `.xlsm` sont conservées, pas les boutons de la feuille Feuil1).
- **Variables de module de `TestInterfaceMVS`** (`ecran`, `curLigne`, `curCol`) : regroupées dans la classe `TerminalMVS`, recréée à chaque commande (`Init_MVS` est rappelé, curseur en (1,1)). Aucune macro ne dépend d'une valeur laissée par la précédente, hormis l'écran lui-même, qui est dans le classeur.
- **Format et nom de la copie.** VBA enregistre toujours en `.xlsm` ; Python reprend l'extension du classeur de départ (`.xlsm` si le classeur est un `.xlsm`, `.xlsx` pour les fichiers de test) pour ne pas produire un fichier dont l'extension ne correspond pas au contenu.
- **Mail réel par défaut.** `--mode-mail` vaut `envoyer` (comme `SEND_MAIL = True`) : lancer une commande sans option envoie vraiment un mail aux adresses codées en dur dans le VBA. Pour essayer sans rien envoyer : `--mode-mail afficher` (brouillon Outlook) ou `--mode-mail aucun`.
- **Fichiers HTML** : écrits en Windows-1252 avec `<meta charset='utf-8'>`, comme le VBA (`Print #f`) ; un caractère non représentable est remplacé par `?`.
- **Fonctions régionales VBA reproduites explicitement** : `IsDate` / `CDate` acceptent `jj/mm/aaaa`, `jj-mm-aaaa`, `jj.mm.aaaa` (années à 2 ou 4 chiffres) et `aaaa-mm-jj` ; `IsNumeric` accepte le format français (`12,5`), l'exposant et `&H`/`&O` ; `Val` ne lit pas l'hexadécimal. Les autres formats de date que VBA accepterait (« 5 octobre 2026 », heures) sont refusés.
- **Barres de données** (`FormatConditions.AddDatabar`) : `DataBarRule` d'openpyxl (min/max automatiques, bleu 638EC6) ; l'aspect peut différer légèrement de la barre par défaut d'Excel.
- **Lenteur occasionnelle à la fermeture d'Excel** (observée sur Excel 16) : lors des tests sur Excel réel, la fermeture de la session a parfois pris ~60 s et affiché `Windows fatal exception: code 0x800706be` / `0x80010108`. Le fichier produit est correct (relu et vérifié par les tests) ; cela vient de la libération des objets COM après `Quit` dans `bibliotheque/excel_com.py` (non modifiable). Aucune instruction n'est refusée.
- Aucun écart non résolu.

## 2. Correspondance des modules

| Module VBA | Fichier Python | Procédure VBA | Fonction Python |
|---|---|---|---|
| (communs) | `commun.py` | `LogMsg`, `GetOrCreateSheet`, `Num`, `Ajoute`, `OuvrirFichier`, blocs HTML identiques, bloc `EnvoyerMail` | `log_msg`, `get_or_create_sheet`, `str_nb`, `HTML_*`, `ouvrir_fichier_vba`, `expedier_mail` |
| MacroInitiale | `macro_initiale.py` | `LancerPipelineComplet` | `lancer_pipeline_complet` |
| MacroInitiale | `macro_initiale.py` | `EnsureStructure` | `ensure_structure` |
| MacroInitiale | `macro_initiale.py` | `AjouterDonneesDuJour`, `LigneExiste` | `ajouter_donnees_du_jour`, `ligne_existe` |
| MacroInitiale | `macro_initiale.py` | `SimulerVolume`, `SimulerTemperature`, `Regions` | `simuler_volume`, `simuler_temperature`, `regions` |
| MacroInitiale | `macro_initiale.py` | `AppliquerFiltre`, `EstRegionValide` | `appliquer_filtre` (+ `connecteurs.excel.appliquer_filtre_excel`), `est_region_valide` |
| MacroInitiale | `macro_initiale.py` | `CreerVisuel` | `creer_visuel_cellules` + `connecteurs.excel.creer_visuel_excel` |
| MacroInitiale | `macro_initiale.py` | `SauvegarderVersion`, `EnvoyerMail` | `sauvegarder_version`, `envoyer_mail_pipeline` |
| MacroCarte | `macro_carte.py` | `LancerPipelineCarteV2` | `lancer_pipeline_carte_v2` |
| MacroCarte | `macro_carte.py` | `EnsureStructure`, `AjouterDonnees`, `LigneExiste`, `SimulerVolume`, `SimulerTemperature`, `EstRegionValide` | `ensure_structure`, `ajouter_donnees`, `ligne_existe`, `simuler_volume`, `simuler_temperature`, `est_region_valide` |
| MacroCarte | `macro_carte.py` | `AppliquerFiltreV2`, `RegionFiltree` | `appliquer_filtre_v2` (renvoie la région saisie) |
| MacroCarte | `macro_carte.py` | `CoordRegion`, `GenererCarteHTML` | `coord_region`, `generer_carte_html` |
| MacroCarte | `macro_carte.py` | `CreerSynthese` | `creer_synthese` (+ AutoFit dans `connecteurs.excel.finaliser_carte`) |
| MacroCarte | `macro_carte.py` | `SauvegarderVersion`, `EnsureFolder`, `EnvoyerMail`, `OuvrirFichier` | `sauvegarder_version`, `envoyer_mail_pipeline`, `commun.ouvrir_fichier_vba` |
| CréationCSV | `creation_csv.py` | `GenererCSVDuJourV3` | `generer_csv_du_jour_v3` |
| CréationCSV | `creation_csv.py` | `GenererCSVDateV3` | `generer_csv_date_v3` |
| CréationCSV | `creation_csv.py` | `EcrireCSV`, `nb`, `Villes`, `V3_NomCSV`, `V3_CreerDossier`, `LogMsg` | `ecrire_csv`, `nb`, `villes`, `v3_nom_csv`, `v3_creer_dossier`, `log_msg` |
| UtilisationCSV | `utilisation_csv.py` | `LancerPipelineCSVV3` | `lancer_pipeline_csv_v3` |
| UtilisationCSV | `utilisation_csv.py` | `ImporterCSV`, `LigneValide`, `EstNombre`, `NomCSV` | `importer_csv`, `ligne_valide`, `est_nombre`, `nom_csv` |
| UtilisationCSV | `utilisation_csv.py` | `DemanderVille`, `GenererCarteHTML` | `demander_ville`, `generer_carte_html` |
| UtilisationCSV | `utilisation_csv.py` | `CreerSynthese`, `SauvegarderVersion`, `EnvoyerMail`, `CreerDossier` | `creer_synthese`, `sauvegarder_version`, `envoyer_mail_pipeline`, `Path.mkdir` |
| SimuMainframe | `simu_mainframe.py` | `OuvrirMainframe` | `ouvrir_mainframe` |
| SimuMainframe | `simu_mainframe.py` | `CreerEcranMVS`, `ReinitialiserBaseMVS`, `DessinerEcran`, `StyleSaisie` | `creer_ecran_mvs`, `reinitialiser_base_mvs`, `dessiner_ecran`, `style_saisie` |
| SimuMainframe | `simu_mainframe.py` | `MVS_Entree`, `MVS_Effacer`, `MVS_PF3` | `mvs_entree`, `mvs_effacer`, `mvs_pf3` |
| SimuMainframe | `simu_mainframe.py` | `Poser`, `Message`, `EffacerSorties`, `EcrireChamp`, `BufferLigne`, `LireChamp`, `NormaliserLigne`, `TrouverISIN`, `FeuilleExiste` | `poser`, `message`, `effacer_sorties`, `ecrire_champ`, `buffer_ligne`, `lire_champ`, `normaliser_ligne`, `trouver_isin`, `feuille_existe` |
| TestInterfaceMVS | `test_interface_mvs.py` | `Init_MVS`, `Modele_Session`, `Attendre`, `Lire`, `Ecrire`, `Touche`, `Recherche`, `Se_Positionner` | méthodes de `TerminalMVS` (`init_mvs`, `modele_session`, `lire`, `ecrire`, `touche`, `recherche`, `se_positionner`), `attendre` |
| TestInterfaceMVS | `test_interface_mvs.py` | `InteragirMainframe`, `DemanderEntier`, `SelectionnerZone` | `interagir_mainframe`, `demander_entier`, `TerminalMVS.selectionner_zone` |
| TestInterfaceMVS | `test_interface_mvs.py` | `Demo_EcrireEtLire`, `Demo_LireSaisieManuelle`, `Demo_VerifierExistenceFSIC`, `Demo_ExtractionVersFeuille` | `demo_ecrire_et_lire`, `demo_lire_saisie_manuelle`, `demo_verifier_existence_fsic`, `demo_extraction_vers_feuille` |
| TestInterfaceMVS | `test_interface_mvs.py` | `ConsulterISIN`, `AppelerMVS`, `VerifierCoord`, `BufferLigne`, `NormaliserLigne`, `FeuilleMVSExiste` | `TerminalMVS.consulter_isin`, `TerminalMVS.appeler_mvs`, `verifier_coord`, `TerminalMVS.buffer_ligne`, `TerminalMVS.normaliser_ligne`, test `SHEET_MVS in wb.sheetnames` |
| Feuil1, Feuil9, ThisWorkbook | — | aucun code | — |

## 3. Entrées attendues

Tous les points d'entrée prennent en premier argument le classeur (`classeur`). Les classeurs des pipelines gaz et carte n'ont besoin d'aucune feuille particulière (Feuille1, Feuille2, Log et le tableau `tblDonnees` sont créés si absents). Le pipeline CSV exige le CSV du jour (voir `--csv-folder`). Les commandes MVS attendent la feuille masquée `MVS_Base` (A1:F7 : ISIN, Libelle, Statut, Encours_MEUR, Date_MAJ, Commentaire) et, sauf `ouvrir-mainframe`, la feuille `MVS` déjà dessinée.

### MacroInitiale.LancerPipelineComplet
Pipeline gaz : crée la structure et 14 jours d'historique si besoin, ajoute les données du jour, filtre, crée le visuel (KPI, tableau croisé TCD_Volume, graphique en courbes), enregistre une copie datée et envoie le mail. Prérequis : Windows, Excel installé (tableau croisé, graphique, filtre, AutoFit) et Outlook connecté (mail), fichier fermé dans Excel.
- `classeur` : remplace `ThisWorkbook`. Classeur `.xlsm` / `.xlsx` ; aucune valeur par défaut.
- `--dossier-sortie` : remplace `SaveFolder()` (`C:\Users\jdavoust\natran\sortiesVBA\`). Dossier de la copie datée. Défaut : `sortie`.
- `--mail-to` : remplace `MAIL_TO`. Défaut : `jdavoust201@headmind.com`.
- `--mail-cc` : remplace `MAIL_CC`. Défaut : vide.
- `--mode-mail` : remplace `SEND_MAIL` (`envoyer` = `.Send`, `afficher` = `.Display`, `aucun` = pas de mail). Défaut : `envoyer`.
- `--saisie` : remplace l'`InputBox` de `AppliquerFiltre` (région à filtrer parmi Nord, Sud, Est, Ouest, Centre ; vide = aucun filtre ; casse libre). Défaut : vide.

### MacroCarte.LancerPipelineCarteV2
Même pipeline avec une carte HTML interactive (hors ligne), une synthèse dans Feuille1 (KPI, tableau du dernier jour, lien), copie datée, mail (classeur + carte) et ouverture de la carte.
- `classeur` : remplace `ThisWorkbook`. Aucune valeur par défaut.
- `--dossier-sortie` : remplace `SaveFolder()`. Dossier de la copie et de la carte. Défaut : `sortie`.
- `--mail-to` : remplace `MAIL_TO`. Défaut : `jdavoust201@headmind.com`.
- `--mail-cc` : remplace `MAIL_CC`. Défaut : vide.
- `--mode-mail` : remplace `SEND_MAIL`. Défaut : `envoyer`.
- `--open-map` : remplace `OPEN_MAP` (`oui` / `non` : ouvrir la carte dans le navigateur). Défaut : `oui`.
- `--saisie` : remplace l'`InputBox` de `AppliquerFiltreV2` (région mise en avant ; vide = toutes). Défaut : vide.

### CréationCSV.GenererCSVDuJourV3
Écrit `DAT_aaaammjj.csv` (date du jour) : 10 villes, `;`, Windows-1252, fins de ligne Windows, valeurs aléatoires. Journal dans la feuille Log du classeur (enregistré en place).
- `classeur` : remplace `ThisWorkbook` (reçoit le Log).
- `--v3-csv-folder` : remplace `V3_CSV_FOLDER` (`C:\Users\jdavoust\natran\VBATF_CSV\`). Dossier du CSV, créé avec ses parents. Défaut : `csv`.

### CréationCSV.GenererCSVDateV3
Idem pour une date choisie.
- `classeur`, `--v3-csv-folder` : comme ci-dessus.
- `--saisie` : remplace l'`InputBox` de date (`jj/mm/aaaa`). Défaut de l'InputBox : la date du jour (option absente) ; vide = rien n'est fait.

### UtilisationCSV.LancerPipelineCSVV3
Importe le CSV du jour (lignes invalides ignorées et journalisées), génère la carte des villes et la feuille Carte_V3, enregistre une copie, envoie le mail (classeur + carte + CSV) et ouvre la carte. Le CSV du jour doit exister (`python creation_csv.py generer-csv-du-jour-v3` avec le même dossier).
- `classeur` : remplace `ThisWorkbook`.
- `--dossier-sortie` : remplace `SaveFolder()`. Défaut : `sortie`.
- `--csv-folder` : remplace `CSV_FOLDER`. Dossier où lire le CSV du jour (même défaut que `--v3-csv-folder`). Défaut : `csv`.
- `--mail-to` : remplace `MAIL_TO`. Défaut : `tom.ferrandis@hotmail.com`.
- `--mail-cc` : remplace `MAIL_CC`. Défaut : vide.
- `--mode-mail` : remplace `SEND_MAIL`. Défaut : `envoyer`.
- `--open-map` : remplace `OPEN_MAP`. Défaut : `oui`.
- `--saisie` : remplace l'`InputBox` de `DemanderVille` (nom d'une ville du CSV ; vide ou inconnue = toutes). Défaut : vide.

### SimuMainframe.OuvrirMainframe
Crée l'écran MVS (et la base MVS_Base masquée si absente), sinon affiche la feuille MVS. Classeur enregistré en place.
- `classeur` : remplace `ThisWorkbook`.

### SimuMainframe.MVS_Entree
Touche ENTREE : consulte l'ISIN saisi sur l'écran ou met à jour le commentaire. Saisir d'abord l'ISIN dans la feuille MVS (cellules AA6…) ou avec `interagir-mainframe`.
- `classeur` : remplace `ThisWorkbook` (enregistré en place).

### SimuMainframe.MVS_Effacer
Touche CLEAR : efface l'ISIN et les champs affichés.
- `classeur` : remplace `ThisWorkbook` (enregistré en place).

### SimuMainframe.MVS_PF3
Touche PF3 : efface l'écran et affiche « PF3 - SESSION TERMINEE (SIMULATION) ».
- `classeur` : remplace `ThisWorkbook` (enregistré en place).

### TestInterfaceMVS.InteragirMainframe
Lit une zone de l'écran, ou y écrit du texte puis valide par ENTREE (la réponse du mainframe est affichée).
- `classeur` : remplace `ThisWorkbook` (enregistré en place).
- `--ligne` : remplace l'`InputBox` de ligne (entier 1 à 32). Défaut : aucun (la commande s'arrête, comme « Annuler »).
- `--colonne` : remplace l'`InputBox` de colonne (entier 1 à 80). Défaut : aucun.
- `--longueur` : remplace l'`InputBox` de longueur à lire (1 à 81 − colonne), utilisée en lecture. Défaut : `10` (valeur par défaut de l'InputBox).
- `--texte` : remplace l'`InputBox` du texte à écrire. Absent = mode lecture ; présent = écriture suivie d'ENTREE ; tronqué en fin de ligne. Défaut : aucun.

### TestInterfaceMVS.Demo_EcrireEtLire
Consulte `FR0010000002`, lit libellé / statut / encours, écrit « MAJ PAR VBA hh:mm:ss » dans le commentaire et valide.
- `classeur` : remplace `ThisWorkbook` (enregistré en place).

### TestInterfaceMVS.Demo_LireSaisieManuelle
Affiche l'ISIN, le commentaire et le message actuellement à l'écran.
- `classeur` : remplace `ThisWorkbook`.

### TestInterfaceMVS.Demo_VerifierExistenceFSIC
Consulte un ISIN et indique s'il existe.
- `classeur` : remplace `ThisWorkbook` (enregistré en place).
- `--isin` : remplace l'`InputBox` (variable `isin`). Défaut : `FR0010000003` ; vide = rien n'est fait.

### TestInterfaceMVS.Demo_ExtractionVersFeuille
Interroge 7 ISIN (le dernier est inconnu) et remplit la feuille `Extraction_MVS` (AutoFit et activation par Excel).
- `classeur` : remplace `ThisWorkbook` (enregistré en place).

## 4. Commandes

<!-- vba2py:debut commandes -->
```powershell
pip install -r requirements.txt
python tests/generer_donnees_fictives.py
python -m pytest -q
```

### MacroInitiale.LancerPipelineComplet

Pourquoi une commande : bouton « Test »

```powershell
python macro_initiale.py --help
python macro_initiale.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie_test --mode-mail aucun        # essai sur données fictives
python macro_initiale.py "C:\chemin\vers\classeur.xlsm" --dossier-sortie "C:\chemin\vers\sorties" --mode-mail afficher        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--dossier-sortie` | SaveFolder() | `sortie` |
| `--mail-to` | MAIL_TO | `jdavoust201@headmind.com` |
| `--mail-cc` | MAIL_CC | — |
| `--mode-mail` | SEND_MAIL | `envoyer` |
| `--saisie` | InputBox de AppliquerFiltre (variable saisie) | — |

### MacroCarte.LancerPipelineCarteV2

Pourquoi une commande : bouton « Test 2 (Carte) »

```powershell
python macro_carte.py --help
python macro_carte.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie_test --mode-mail aucun        # essai sur données fictives
python macro_carte.py "C:\chemin\vers\classeur.xlsm" --dossier-sortie "C:\chemin\vers\sorties" --mode-mail afficher        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--dossier-sortie` | SaveFolder() | `sortie` |
| `--mail-to` | MAIL_TO | `jdavoust201@headmind.com` |
| `--mail-cc` | MAIL_CC | — |
| `--mode-mail` | SEND_MAIL | `envoyer` |
| `--open-map` | OPEN_MAP | `oui` |
| `--saisie` | InputBox de AppliquerFiltreV2 (variable saisie) | — |

### CréationCSV.GenererCSVDuJourV3

Pourquoi une commande : bouton « Création CSV »

```powershell
python creation_csv.py generer-csv-du-jour-v3 --help
python creation_csv.py generer-csv-du-jour-v3 tests/fixtures/classeur_complet.xlsx        # essai sur données fictives
python creation_csv.py generer-csv-du-jour-v3 "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--v3-csv-folder` | V3_CSV_FOLDER | `csv` |

### CréationCSV.GenererCSVDateV3

Pourquoi une commande : macro publique appelée par aucune autre procédure (InputBox de date)

```powershell
python creation_csv.py generer-csv-date-v3 --help
python creation_csv.py generer-csv-date-v3 tests/fixtures/classeur_complet.xlsx        # essai sur données fictives
python creation_csv.py generer-csv-date-v3 "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--v3-csv-folder` | V3_CSV_FOLDER | `csv` |
| `--saisie` | InputBox de GenererCSVDateV3 (variable saisie) | — |

### UtilisationCSV.LancerPipelineCSVV3

Pourquoi une commande : bouton « Macro avec Intégration »

```powershell
python utilisation_csv.py --help
python utilisation_csv.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie_test --mode-mail aucun        # essai sur données fictives
python utilisation_csv.py "C:\chemin\vers\classeur.xlsm" --dossier-sortie "C:\chemin\vers\sorties" --mode-mail afficher        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--dossier-sortie` | SaveFolder() | `sortie` |
| `--csv-folder` | CSV_FOLDER | `csv` |
| `--mail-to` | MAIL_TO | `tom.ferrandis@hotmail.com` |
| `--mail-cc` | MAIL_CC | — |
| `--mode-mail` | SEND_MAIL | `envoyer` |
| `--open-map` | OPEN_MAP | `oui` |
| `--saisie` | InputBox de DemanderVille (variable saisie) | — |

### SimuMainframe.OuvrirMainframe

Pourquoi une commande : bouton « Création MVS »

```powershell
python simu_mainframe.py ouvrir-mainframe --help
python simu_mainframe.py ouvrir-mainframe tests/fixtures/classeur_complet.xlsx        # essai sur données fictives
python simu_mainframe.py ouvrir-mainframe "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### SimuMainframe.MVS_Entree

Pourquoi une commande : commentaire « TOUCHES DU TERMINAL (boutons optionnels sur la feuille MVS) »

```powershell
python simu_mainframe.py mvs-entree --help
python simu_mainframe.py mvs-entree tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python simu_mainframe.py mvs-entree "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### SimuMainframe.MVS_Effacer

Pourquoi une commande : commentaire « TOUCHES DU TERMINAL (boutons optionnels sur la feuille MVS) »

```powershell
python simu_mainframe.py mvs-effacer --help
python simu_mainframe.py mvs-effacer tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python simu_mainframe.py mvs-effacer "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### SimuMainframe.MVS_PF3

Pourquoi une commande : commentaire « TOUCHES DU TERMINAL (boutons optionnels sur la feuille MVS) »

```powershell
python simu_mainframe.py mvs-pf3 --help
python simu_mainframe.py mvs-pf3 tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python simu_mainframe.py mvs-pf3 "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### TestInterfaceMVS.InteragirMainframe

Pourquoi une commande : bouton « Interra MVS »

```powershell
python test_interface_mvs.py interagir-mainframe --help
python test_interface_mvs.py interagir-mainframe tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python test_interface_mvs.py interagir-mainframe "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--ligne` | InputBox de DemanderEntier : ligne (1 à 32) | — |
| `--colonne` | InputBox de DemanderEntier : colonne (1 à 80) | — |
| `--longueur` | InputBox de DemanderEntier : longueur à lire (mode lecture) | `10` |
| `--texte` | InputBox texte à écrire (mode écriture ; absent = lecture) | — |

### TestInterfaceMVS.Demo_EcrireEtLire

Pourquoi une commande : commentaire « DEMOS (boutons optionnels) »

```powershell
python test_interface_mvs.py demo-ecrire-et-lire --help
python test_interface_mvs.py demo-ecrire-et-lire tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python test_interface_mvs.py demo-ecrire-et-lire "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### TestInterfaceMVS.Demo_LireSaisieManuelle

Pourquoi une commande : commentaire « DEMOS (boutons optionnels) »

```powershell
python test_interface_mvs.py demo-lire-saisie-manuelle --help
python test_interface_mvs.py demo-lire-saisie-manuelle tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python test_interface_mvs.py demo-lire-saisie-manuelle "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### TestInterfaceMVS.Demo_VerifierExistenceFSIC

Pourquoi une commande : commentaire « DEMOS (boutons optionnels) »

```powershell
python test_interface_mvs.py demo-verifier-existence-fsic --help
python test_interface_mvs.py demo-verifier-existence-fsic tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python test_interface_mvs.py demo-verifier-existence-fsic "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |
| `--isin` | InputBox de Demo_VerifierExistenceFSIC (variable isin) | `FR0010000003` |

### TestInterfaceMVS.Demo_ExtractionVersFeuille

Pourquoi une commande : commentaire « DEMOS (boutons optionnels) »

```powershell
python test_interface_mvs.py demo-extraction-vers-feuille --help
python test_interface_mvs.py demo-extraction-vers-feuille tests/fixtures/classeur_mvs.xlsx        # essai sur données fictives
python test_interface_mvs.py demo-extraction-vers-feuille "C:\chemin\vers\classeur.xlsm"        # sur un vrai classeur
```

| Argument | Remplace dans le VBA | Valeur par défaut |
|---|---|---|
| `classeur` | ThisWorkbook | — |

### Procédures internes (sans commande)

- MacroInitiale.AjouterDonneesDuJour : appelée seulement par LancerPipelineComplet
- MacroInitiale.AppliquerFiltre : appelée seulement par LancerPipelineComplet
- MacroInitiale.CreerVisuel : appelée seulement par LancerPipelineComplet
- MacroCarte.AppliquerFiltreV2 : appelée seulement par LancerPipelineCarteV2
- SimuMainframe.CreerEcranMVS : appelée seulement par OuvrirMainframe
- SimuMainframe.ReinitialiserBaseMVS : appelée seulement par CreerEcranMVS
<!-- vba2py:fin commandes -->

## 5. Choix de traduction

- `Application.ScreenUpdating`, `DisplayAlerts` : supprimés (aucun équivalent nécessaire).
- `Cells(Rows.Count, 1).End(xlUp).Row` → `commun.derniere_ligne` (remontée explicite).
- `ListObjects.Add`, `ListRows.Add`, `.Name`, `.TableStyle` → `Table` openpyxl (`commun.creer_table`, `ajouter_ligne_table` : la plage et l'`autoFilter` suivent la dernière ligne ; dates écrites comme vraies dates, format `dd/mm/yyyy`).
- `Cells.Clear`, `ChartObjects.Delete`, `PivotTables.TableRange2.Clear`, `Hyperlinks.Delete` → `preparer_feuille` + suppression des liens et de la mise en forme conditionnelle.
- `PivotCaches.Create` / `CreatePivotTable` / `PivotFields` / `AddDataField`, graphique `SetSourceData pt.TableRange1`, `ChartType = xlLine` → exécutés par Excel dans `connecteurs/excel.py` (`creer_visuel_excel`), instruction par instruction, dans le même ordre. **Pas d'image ni de valeurs figées.**
- `lo.AutoFilter.ShowAllData` / `lo.Range.AutoFilter Field:=2, Criteria1:=saisie` → Excel (`appliquer_filtre_excel`), pour que le filtre soit réellement appliqué (lignes masquées). `RegionFiltree` (relecture du critère) devient la valeur saisie.
- `Columns(...).AutoFit`, `Columns("A").ColumnWidth = 24`, `ws.Activate` sur les feuilles créées → Excel (`ajuster_colonnes`, `activer_feuille`), regroupés dans **une seule session par fichier** (`finaliser_initiale`, `finaliser_carte`, `finaliser_csv_v3`, `ajuster_feuille`). Ordre : étapes openpyxl → enregistrement de la copie → session Excel sur la copie → mail.
- `.NumberFormat = "dd/mm"` sur un champ de tableau croisé → `affecter_format_nombre` : envoyé avec le **LCID anglais** (0x0409). Raison (constatée sur l'Excel de cette machine) : VBA transmet toujours les codes de format en anglais, alors qu'un appel COM depuis Python utilise la langue de Windows ; Excel (français) acceptait `jj/mm/aaaa` mais refusait `dd/mm` (« Impossible de définir la propriété NumberFormat de la classe PivotField »). Le script `tests/diagnostic_excel.py` reproduit l'analyse.
- `FormatConditions.AddDatabar` → `DataBarRule` ; `Interior.Color = RGB(221,235,247)` → remplissage `DDEBF7`.
- `Hyperlinks.Add Anchor, Address, TextToDisplay` → `cell.hyperlink`, `cell.value`, style `Hyperlink`.
- `ThisWorkbook.SaveCopyAs` → `wb.save(copie)` ; `MkDir` → `Path.mkdir(parents=True)` (VBA échouerait si un dossier parent manque).
- `Open ... For Output` / `Print #f` → Windows-1252, CRLF (CSV) ou lignes LF + CRLF final (HTML) ; `Line Input` → lecture Windows-1252, séparateurs CR / LF / CRLF.
- `Val`, `IsNumeric`, `CDate`/`IsDate`, `Str$`, `CStr`, `CInt`, `DateSerial` → fonctions explicites de `commun.py` (point décimal pour `Val` et `Str$`, virgule pour `IsNumeric` et `CStr`, débordement des mois pour `DateSerial`), indépendantes de la locale de la machine.
- `Rnd` / `Randomize` → `random.Random` (injectable dans les tests). `Round` → `round` (arrondi bancaire).
- `Err.Raise vbObjectError + n` → `ErreurVBA(numero, description)` ; `On Error GoTo ErrHandler` → `try/except` qui consigne `Err.Number - Err.Description` dans le Log et en console, puis renvoie un résultat en échec (code de sortie 1). `MailErr` et `OuvrirFichier` (`On Error Resume Next`) consignent l'erreur sans interrompre.
- `CreateObject("Outlook.Application")` → `connecteurs/outlook.py` (win32com importé à l'intérieur de la fonction ; pièces jointes en chemin absolu). `FollowHyperlink` → `connecteurs/systeme.py` (`os.startfile`).
- `Application.Run nomMacro` (même classeur) → dictionnaire `{nom VBA: fonction}` dans `TerminalMVS.appeler_mvs`.
- `ActiveWindow.DisplayGridlines`, `Zoom`, `Select`, `Activate`, `Visible = xlSheetHidden/xlSheetVisible` → `sheet_view.showGridLines`, `zoomScale`, `selection`, `wb.active`, `sheet_state`. `ws.Cells.Interior.Color` (feuille entière) → style des colonnes A:XFD et des cellules de la zone 36×80.
- Code dupliqué : `LogMsg` (MacroInitiale, MacroCarte, UtilisationCSV), `GetOrCreateSheet`, `Num`/`nb`, les blocs HTML identiques (entête CSS, SVG, contour de la France), le bloc d'envoi de mail et `OuvrirFichier` sont regroupés dans `commun.py` car strictement identiques. Gardés séparés car différents : `LogMsg` de CréationCSV, `EnsureStructure`/`AjouterDonnees`/`SimulerVolume`... de MacroCarte (statut « Phase de test V2 »), les `BufferLigne` / `NormaliserLigne` de SimuMainframe (paramètre `ws`) et de TestInterfaceMVS (variable de module), `CreerDossier` (UtilisationCSV) et `V3_CreerDossier` (CréationCSV).
- Un seul module commun (`commun.py`) et un dossier `connecteurs/`. `InputBox` / `UserForm` : l'interface imposée est la ligne de commande ; Streamlit n'est pas utilisé.
- Entiers : le type `Integer` de VBA (16 bits) n'a pas d'effet ici (valeurs ≤ 80).

## 6. TODO et fonctionnalités non couvertes

Aucun.

## 7. Éléments suspects

Aucune instruction adressée à un agent n'a été trouvée dans le VBA. En revanche :
- **Adresses mail codées en dur**, conservées comme valeurs par défaut (comportement fidèle) : `jdavoust201@headmind.com` (MacroInitiale, MacroCarte) et `tom.ferrandis@hotmail.com` (UtilisationCSV, adresse personnelle différente des deux autres modules, commentée « A modifier avec mail HDP »).
- **Envoi automatique** : `SEND_MAIL = True` dans les trois pipelines, avec un corps « envoi automatique en phase de test » signé « Robot VBATF » ; par défaut, la commande envoie réellement le mail.
- **Chemins personnels** codés en dur, remplacés par des dossiers relatifs (`sortie`, `csv`) : `C:\Users\jdavoust\natran\sortiesVBA\` (3 modules) et `C:\Users\jdavoust\natran\VBATF_CSV\` (CréationCSV, UtilisationCSV).
- `Application.Run nomMacro` dans TestInterfaceMVS appelle des macros du même classeur (MVS_Entree, MVS_Effacer, MVS_PF3) : traduit par des appels directs, rien d'externe.

## 8. Comparer avec la macro d'origine dans Excel

Les fichiers de `tests/fixtures/` sont des `.xlsx` : pour y lancer les macros VBA, ouvrir le classeur d'origine `macros_natran_vf.xlsm` (ou importer les modules de `_source_vba/` dans une copie du fixture enregistrée en `.xlsm`). Dans les modules VBA, passer `SEND_MAIL` à `False` (brouillon) ou commenter l'envoi, et mettre un dossier de test dans `SaveFolder()`.

1. **Pipeline gaz.** Copie de `tests/fixtures/classeur_complet.xlsx` en `.xlsm` avec les modules VBA, lancer `LancerPipelineComplet` (saisir `Nord` à la demande de filtre). Puis `python macro_initiale.py tests/fixtures/classeur_complet.xlsx --dossier-sortie sortie --mode-mail aucun --saisie Nord`. Comparer dans la copie datée : feuilles Feuille1 (titre, KPI B2/E2/H2, tableau croisé TCD_Volume en A5, graphique en J5 « Volumes distribues par region (GWh) »), Feuille2 (tableau tblDonnees, 14 jours + jour courant, filtre sur Nord) et Log. Les valeurs simulées sont aléatoires : comparer la forme, pas les nombres.
2. **Pipeline carte.** Avec `tests/fixtures/classeur_donnees.xlsx` : `LancerPipelineCarteV2` puis `python macro_carte.py tests/fixtures/classeur_donnees.xlsx --mode-mail aucun --open-map non`. Comparer le HTML (ouvrir les deux cartes : mêmes cercles, sélecteurs, total), Feuille1 (lien, barres de données, formules SUMIFS) et la largeur de la colonne A (24).
3. **CSV.** Lancer `GenererCSVDuJourV3`, puis `python creation_csv.py generer-csv-du-jour-v3 tests/fixtures/classeur_complet.xlsx --v3-csv-folder csv` : comparer les fichiers `DAT_aaaammjj.csv` (en-tête, 10 villes, séparateur `;`, fins de ligne CRLF). Ensuite `LancerPipelineCSVV3` (ville `Lyon`) et `python utilisation_csv.py tests/fixtures/classeur_complet.xlsx --csv-folder csv --mode-mail aucun --open-map non --saisie Lyon` : comparer Donnees_V3, Carte_V3 et la carte HTML. `tests/fixtures/csv/DAT_20261027.csv` contient 4 lignes valides et 4 invalides pour vérifier le Log (« Ligne ignoree »).
4. **Mainframe.** Avec `tests/fixtures/classeur_complet.xlsx` : `OuvrirMainframe`, saisir `FR0010000002` en AA6, `MVS_Entree`. Python : `python simu_mainframe.py ouvrir-mainframe <copie>.xlsx`, saisir l'ISIN avec `python test_interface_mvs.py interagir-mainframe <copie>.xlsx --ligne 6 --colonne 27 --texte FR0010000002`, puis `python simu_mainframe.py mvs-entree <copie>.xlsx`. Comparer l'écran (libellé, statut, encours, date, commentaire, message « FSIC - ENREGISTREMENT TROUVE »). Idem pour `Demo_EcrireEtLire`, `Demo_VerifierExistenceFSIC` (`--isin FR0000000000` : inconnu) et `Demo_ExtractionVersFeuille` (feuille Extraction_MVS, 7 lignes, dernière « INCONNU »), avec `tests/fixtures/classeur_mvs.xlsx` (écran déjà dessiné).
5. **Tests sur Excel réel.** `python -m pytest -m excel_reel` (nécessite Excel, fichiers fermés) crée le tableau croisé, le graphique, le filtre et les largeurs dans des classeurs de test.

## 9. Tests

<!-- vba2py:debut tests -->
Vérification indépendante par vba2py : `python -m pytest -q tests` → ✅ **178 passed in 72.61s (0:01:12)**

Tests sur Excel réel (`python -m pytest -m excel_reel`) : ✅ **4 passed, 174 deselected in 67.74s (0:01:07)**
<!-- vba2py:fin tests -->

## 10. Extraction

<!-- vba2py:debut extraction -->
- Origine : export
- Dossier de sortie : `macros_natran_vf_vba`

| Module | Type | Lignes de code | Statut |
|---|---|---|---|
| CréationCSV | module | 127 | extrait |
| Feuil1 | classe/document | 0 | vide (ignoré) |
| Feuil9 | classe/document | 0 | vide (ignoré) |
| MacroCarte | module | 431 | extrait |
| MacroInitiale | module | 278 | extrait |
| SimuMainframe | module | 320 | extrait |
| TestInterfaceMVS | module | 355 | extrait |
| ThisWorkbook | classe/document | 0 | vide (ignoré) |
| UtilisationCSV | module | 406 | extrait |

## Macros reliées à des boutons

| Feuille | Bouton | Macro |
|---|---|---|
| Feuil1 | Test | LancerPipelineComplet |
| Feuil1 | Test 2 (Carte) | LancerPipelineCarteV2 |
| Feuil1 | Création CSV | GenererCSVDuJourV3 |
| Feuil1 | Macro avec Intégration | LancerPipelineCSVV3 |
| Feuil1 | Création MVS | OuvrirMainframe |
| Feuil1 | Interra MVS | InteragirMainframe |

Feuilles masquées : MVS_Base
<!-- vba2py:fin extraction -->
