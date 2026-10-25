# Rapport de traduction VBA → Python

## ⚠️ Écarts de comportement avec la macro d'origine

1. **Tableau croisé dynamique (TCD)** : openpyxl ne sait pas créer de TCD. `creer_visuel` écrit à partir de `A5` une **copie statique** (valeurs calculées avec pandas, mêmes lignes/colonnes/totaux, format `dd/mm`). Elle n'est pas actualisable dans Excel. `# TODO(vba2py)` non nécessaire : aucune librairie autorisée ne le permet.
2. **Graphique** : la macro crée un graphique Excel natif ; Python génère un **PNG matplotlib** (`graphique_volumes.png`) intégré en `J5` de `Feuille1`. Le graphique VBA incluait aussi les totaux du TCD ; le PNG ne trace que les 5 régions.
3. **Formules KPI** (B2, E2, H2) : écrites telles quelles (références structurées `tblDonnees[...]`) ; openpyxl ne les calcule pas, elles se calculent à l'ouverture dans Excel. `calculer_kpi()` fournit les mêmes valeurs côté Python.
4. **Extension de la copie** : `.xlsm` seulement si le classeur source est un `.xlsm` (chargé avec `keep_vba=True`) ; sinon `.xlsx` (un `.xlsm` sans projet VBA serait illisible par Excel).
5. **Filtre** : le filtre automatique est posé sur la table et les lignes non retenues sont masquées (openpyxl ne ré-évalue pas les filtres ; Excel peut les ré-appliquer à l'ouverture).
6. **Largeurs de colonnes** (`AutoFit`) : approximation par longueur de texte.
7. **Classeur source** : comme `SaveCopyAs`, seule la copie est écrite ; le classeur source n'est modifié qu'avec `--enregistrer-classeur`.
8. **Nombres aléatoires** : `Rnd()` → `random.Random` (valeurs différentes de celles d'Excel, mêmes fourchettes).

## Correspondance des modules

| Module VBA | Fichier Python | Procédures |
|---|---|---|
| `marco_25_10.bas` | `marco_25_10.py` | `LancerPipelineComplet` → `lancer_pipeline_complet` ; `EnsureStructure` → `ensure_structure` ; `AjouterDonneesDuJour` → `ajouter_donnees_du_jour` ; `LigneExiste` → `ligne_existe` ; `SimulerVolume/Temperature` → `simuler_volume/temperature` ; `AppliquerFiltre` → `appliquer_filtre` ; `EstRegionValide` → `est_region_valide` ; `CreerVisuel` → `creer_visuel` ; `SauvegarderVersion` → `sauvegarder_version` ; `EnvoyerMail` → `envoyer_mail_pj` ; `GetOrCreateSheet`, `LogMsg`, `Regions`, `SaveFolder` → `get_or_create_sheet`, `log_msg`, `regions`, `save_folder` |
| (appel Outlook) | `connecteurs/outlook.py` | `envoyer_mail(...)` : COM, modes `envoyer` / `afficher` |
| `Feuil1-4`, `ThisWorkbook` | — | vides, ignorés |

## Choix de traduction

- `InputBox` → paramètre `region_filtre` (`--region`) ; vide = Annuler = aucun filtre ; région inconnue → avertissement `logging`, pas de filtre.
- `MsgBox` → `logging.info` / `logging.warning` / `logging.error`.
- `On Error GoTo ErrHandler` → `try/except Exception` : l'erreur est journalisée dans la feuille `Log` et la fonction retourne `None` (comme le VBA, le pipeline ne plante pas).
- `On Error GoTo MailErr` → erreur d'envoi journalisée dans `Log`, pipeline non interrompu.
- `SEND_MAIL`, `MAIL_TO`, `MAIL_CC`, `SaveFolder` → options `--mode-mail envoyer|afficher|aucun`, `--destinataire`, `--cc`, `--dossier-sauvegarde` (défauts = valeurs du VBA).
- `Application.ScreenUpdating` / `DisplayAlerts` : supprimés (sans équivalent).
- `LigneExiste` : comparaison de région sensible à la casse (`=` VBA sans `Option Compare Text`) ; `EstRegionValide` insensible (`vbTextCompare`).
- Le mail et le graphique reproduisent objet, corps (`\r\n`), pièce jointe du VBA.
- Dans la copie sauvegardée, les logs « Sauvegarde » et « Pipeline termine » sont absents (écrits après `SaveCopyAs`, comme en VBA).
- Dates : en VBA `Date`/`Now` ; ici paramètre `maintenant` (testable).
- Entiers 16 bits, `True = -1`, `ByRef` : non concernés.

## TODO / non couvert

- Aucun `# TODO(vba2py)` dans le code. Fonctionnalités non reproduites : voir les écarts ci-dessus.

## Éléments suspects

Aucun. Les commentaires du VBA sont des commentaires de maintenance (« A modifier avec mail HDP », etc.), sans consigne. À noter : l'adresse `jdavoust201@headmind.com` et le chemin `C:\Users\Jules\...` sont codés en dur dans le VBA ; conservés comme valeurs par défaut des options.

## Prérequis mail

Windows, Outlook installé et connecté au compte, `pywin32`. Pour tester **sans rien envoyer** : `--mode-mail afficher` (ouvre le brouillon) ou `--mode-mail aucun`. ⚠️ Le défaut est `envoyer`, comme `SEND_MAIL = True` en VBA.

## Résultat des tests

`python -m pytest -q` : **46 passés, 0 échoué.** Aucun test n'appelle Outlook (fixture `autouse` qui bloque `connecteurs.outlook.envoyer_mail`).

## Utilisation

```
pip install -r requirements.txt
python marco_25_10.py tests/fixtures/classeur_complet.xlsx --dossier-sauvegarde sortie --region Nord --mode-mail aucun
python tests/generer_donnees_fictives.py     # (re)génère tests/fixtures/
python -m pytest -q
```

## Comparer avec la macro d'origine

1. Ouvrir `tests/fixtures/classeur_complet.xlsx` (feuilles Dashboard, Feuille2 avec table `tblDonnees` sur 3 jours × 5 régions, Feuille1, Log), l'enregistrer en `.xlsm` et y importer `marco_25_10.bas`.
2. Lancer `LancerPipelineComplet` (adapter `SaveFolder` et `MAIL_TO`, mettre `SEND_MAIL = False`).
3. Comparer avec la sortie Python : mêmes lignes ajoutées (5 lignes « Phase de test »), mêmes KPI attendus sur les données de départ (total 1350, moyenne/jour 450, région du pic Nord), mêmes noms de fichiers, objet et corps du mail. Les volumes simulés diffèrent (aléatoire).
