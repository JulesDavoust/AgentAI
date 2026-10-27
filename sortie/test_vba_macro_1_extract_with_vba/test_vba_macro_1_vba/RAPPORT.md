# Rapport de traduction VBA → Python

## ⚠️ Écarts de comportement avec la macro d'origine

1. **Tableau croisé dynamique (`TCD_Volume`)** : openpyxl ne sait pas créer de TCD. Le TCD est écrit en A5 **en valeurs statiques** (calculées avec pandas, mêmes lignes/colonnes/totaux), non actualisable dans Excel. `# TODO(vba2py)` dans `creer_visuel`.
2. **Graphique** : le graphique Excel lié au TCD est remplacé par un PNG matplotlib (`<dossier>/graphiques/graphique_volumes.png`), inséré en J5 de Feuille1 dans la copie sauvegardée (image, pas un objet graphique Excel).
3. **`InputBox` du filtre** : remplacé par l'option `--region` / le paramètre `region` (pas d'interface Streamlit, la saisie en ligne de commande suffit ici). Le filtre est réellement posé sur le tableau (critère d'autofiltre + lignes masquées), mais Excel ne le réévalue pas à l'ouverture.
4. **Persistance** : en VBA, `SaveCopyAs` n'enregistre que la copie ; le classeur d'origine reste ouvert en mémoire avec les nouvelles données. En Python, le classeur source n'est **pas** modifié par défaut (idem) ; `--enregistrer-source` permet de le réécrire pour conserver les données entre deux exécutions.
5. **Extension `.xlsm`** : la copie garde l'extension du classeur source (`keep_vba=True` pour un `.xlsm`). Si le classeur est `.xlsx`, la copie est `.xlsx`. Si le classeur n'existe pas, il est créé (Feuil1..Feuil4) en `.xlsx`.
6. **Formules** : écrites mais non calculées (openpyxl) ; Excel les calcule à l'ouverture. Références structurées `tblDonnees[...]` conservées.
7. **`Rnd()`** : générateur Python (`random.Random`), pas la même suite que VBA ; les valeurs simulées ne sont donc pas reproductibles à l'identique (mêmes bornes).

## Module VBA → fichier Python

| Module VBA | Fichier Python | Procédures traduites |
|---|---|---|
| `marco_25_10` (288 lignes) | `marco_25_10.py` | `LancerPipelineComplet` → `lancer_pipeline_complet`, `EnsureStructure`, `AjouterDonneesDuJour`, `LigneExiste`, `SimulerVolume`, `SimulerTemperature`, `AppliquerFiltre`, `EstRegionValide`, `CreerVisuel` → `creer_visuel` (+ `construire_tcd`, `exporter_graphique`), `SauvegarderVersion`, `EnvoyerMail` → `envoyer_mail_pipeline`, `GetOrCreateSheet`, `LogMsg`, `Regions`, `SaveFolder` (constante `SAVE_FOLDER`) |
| (appel Outlook) | `connecteurs/outlook.py` | `envoyer_mail` (COM : `.To/.CC/.Subject/.Body/.Attachments.Add/.Send/.Display`) |

Feuil1–Feuil4 et ThisWorkbook sont vides : aucun fichier.

## Choix de traduction

- Le classeur est un paramètre (`ThisWorkbook` → `charger_ou_creer_classeur(chemin)`) ; dossier de sauvegarde, dossier PNG, destinataire, CC et mode mail sont des options CLI dont les valeurs par défaut reproduisent le VBA (`SEND_MAIL = True` → `--mode-mail envoyer`, destinataire `jdavoust201@headmind.com`, dossier `C:\Users\Jules\...\sortiesVBA\`).
- Le module métier ne touche pas à COM : il reçoit `envoyer_mail` (défaut : connecteur Outlook, résolu à l'appel). Mode `aucun` ajouté pour ne rien envoyer.
- `On Error GoTo ErrHandler` → `try/except Exception` dans le pipeline : erreur journalisée dans la feuille Log, retour `None`. `MailErr` → erreur du mail journalisée sans interrompre.
- `MsgBox` → `logging`. `Application.ScreenUpdating/DisplayAlerts` supprimés (sans équivalent).
- `Date`/`Now` injectables (`aujourdhui`, `maintenant`) pour des tests déterministes.
- `LigneExiste` : comparaison de la partie date (`CLng`) et de la région sensible à la casse (`=` VBA par défaut) ; `EstRegionValide` insensible à la casse (`vbTextCompare`).
- `Round` : `round()` Python (arrondi bancaire, comme VBA). Pas d'arithmétique sur `Integer` 16 bits ni sur booléens.
- Ordre des colonnes du TCD : alphabétique (comme Excel).
- Les erreurs de `wb.save` (dossier impossible) sont journalisées dans le classeur en mémoire (non sauvegardé sauf `--enregistrer-source`).

## TODO / non couvert

- TCD natif et graphique Excel natif (voir écarts).
- Pas de déclenchement automatique : aucune procédure d'événement dans le VBA, donc rien à signaler de ce côté.

## Éléments suspects

Aucun. Les commentaires du VBA (« A modifier avec mail HDP », « A modifier avec path de chancun de vs 2 ») sont de simples notes de configuration. Remarque : le VBA envoie un vrai mail par défaut (`SEND_MAIL = True`) ; utiliser `--mode-mail afficher` ou `aucun` pour tester.

## Résultat des tests

`python -m pytest -q` : **33 passés, 0 échoué** (1 correction de code en un cycle : `pd.Timedelta` → `timedelta`). Aucun test n'a été modifié pour passer.

## Utilisation

Prérequis pour le mail réel : Windows, Outlook installé et connecté au compte, `pip install -r requirements.txt`.

```
python marco_25_10.py classeur.xlsm --dossier sortie --region Nord --mode-mail afficher
python -m pytest -q
python tests/generer_donnees_fictives.py      # régénère tests/fixtures/
```

## Comparer avec la macro d'origine

`tests/fixtures/classeur_vide.xlsx` (Feuil1–4 vides : la macro crée tout, 14 jours d'historique aléatoires) et `tests/fixtures/classeur_donnees.xlsx` (Feuille2 avec `tblDonnees`, 15 lignes du 20 au 22/10/2025 : volume = base région + n° du jour − 20, température 5 à 9). Ouvrir l'un d'eux dans Excel, y importer `_source_vba/marco_25_10.bas`, lancer `LancerPipelineComplet` (avec `SEND_MAIL = False` pour ne rien envoyer), puis lancer `python marco_25_10.py <même classeur> --mode-mail aucun` et comparer : lignes ajoutées (5 « Phase de test »), KPI (B2/E2/H2), TCD en A5, nom du fichier sauvegardé. Les valeurs simulées étant aléatoires, comparer sur `classeur_donnees.xlsx` avec les lignes déjà présentes.
