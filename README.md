# AgentAI — vba2py

Agent qui traduit des macros VBA Excel en Python et livre, pour chaque classeur, le code
traduit, une batterie de tests sur données fictives et un rapport.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate          # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

Clé API (créée sur console.anthropic.com) :

```bash
export ANTHROPIC_API_KEY=sk-ant-...          # macOS / Linux
$env:ANTHROPIC_API_KEY = "sk-ant-..."        # Windows PowerShell (session en cours)
setx ANTHROPIC_API_KEY "sk-ant-..."          # Windows, permanent (rouvrir le terminal)
```

## Utilisation

```bash
# 1. Extraction seule (sans IA, gratuit) : vérifier ce qui a été trouvé
python main.py extraire examples/Ventes.xlsm -o sortie

# 2. Traduction complète : extraction → agent → vérification indépendante des tests
python main.py traduire examples/Ventes.xlsm -o sortie
python main.py traduire examples/ -o sortie            # un dossier entier, arborescence conservée
python main.py traduire examples/Ventes.xlsm -o sortie --modele claude-sonnet-5-5 --budget 2

# 3. Relancer les tests d'un projet traduit
python main.py verifier sortie/ventes

# 4. Contrôler l'extraction automatique contre un export manuel depuis Excel
python main.py comparer examples/Ventes_vba examples/Ventes.xlsm
```

### Entrées acceptées

| Entrée | Extraction |
|---|---|
| Classeur `.xlsm`, `.xlsb`, `.xls`, `.xlam` | automatique avec `oletools` (analyse statique, aucune macro exécutée) |
| Dossier de modules exportés (`.bas`, `.cls`, `.frm`) | lecture directe ; le plus fidèle, surtout pour les UserForms. Un classeur posé dans le même dossier sert à lire la structure des feuilles |

Pour exporter les modules depuis Excel : importez `tools/ExportVBA.bas` (idéalement dans
PERSONAL.XLSB), activez le classeur voulu et lancez la macro `ExporterModulesVBA`.

## Résultat

```
sortie/
├── _journaux/ventes.jsonl        toutes les actions de l'agent (autorisées et refusées)
└── ventes/
    ├── _source_vba/              VBA extrait + structure du classeur (lecture seule pour l'agent)
    ├── module1.py                un fichier Python par module VBA
    ├── requirements.txt
    ├── interface.json            contrat de la ligne de commande du script
    ├── RAPPORT.md                structure fixe (prompts/rapport_modele.md)
    └── tests/
        ├── generer_donnees_fictives.py
        ├── fixtures/             classeurs fictifs, à ouvrir aussi dans Excel pour comparer avec la macro d'origine
        └── test_module1.py
```

## Stabilité d'une traduction à l'autre

- **Rapport** : titres et ordre imposés par `prompts/rapport_modele.md`. Les sections
  « Commandes », « Tests » et « Extraction » sont écrites par vba2py, pas par l'agent.
- **Commandes** : les noms d'arguments sont déduits du VBA par des règles fixes, et décrits
  dans `interface.json`. Après une traduction réussie, ce contrat est enregistré dans
  `interfaces/<projet>.json` (à committer) ; les traductions suivantes du même projet doivent
  le respecter.
- **Vérification indépendante** (`traduire` et `verifier`) : tests relancés, `--help` du script
  comparé à `interface.json`, structure du rapport contrôlée.

## Sécurité

- Le VBA est traité comme une donnée : l'agent a pour consigne de ne jamais suivre
  d'instructions trouvées dans le code ou ses commentaires.
- Chaque action de l'agent est contrôlée par du code (`vba2py/guardrails.py`) avant exécution :
  lecture et écriture limitées au dossier du projet, `_source_vba/` en lecture seule,
  commandes limitées à `pytest`, `py_compile` et aux scripts du projet, pas d'accès web,
  imports dangereux refusés dans le code écrit (réseau, sous-processus, `eval`, COM Windows).
- Les réglages de la machine (`~/.claude`, CLAUDE.md, skills) ne sont pas chargés.
- Limites par projet : nombre de tours (`--max-tours`) et budget en dollars (`--budget`).
- Limite assumée : `pytest` exécute le code écrit par l'agent. Pour des classeurs d'origine
  inconnue, lancez l'outil dans un conteneur.

## Structure de l'outil

```
main.py                 ligne de commande
vba2py/extract.py       extraction (oletools / exports) + structure du classeur (openpyxl)
vba2py/compare.py       comparaison export manuel ↔ extraction automatique
vba2py/guardrails.py    contrôle de chaque action de l'agent + journal
vba2py/agent.py         configuration Claude Agent SDK
prompts/system.md       instructions de l'agent
prompts/mapping.md      technologies autorisées et correspondances VBA → Python
tools/ExportVBA.bas     macro d'export des modules depuis Excel
tests/                  tests de l'outil (sans appel à l'IA) : python -m pytest
```

## Ce qu'openpyxl ne sait pas faire : Excel piloté depuis Python (branche `pivot-excel`)

Règle de traduction : openpyxl par défaut ; tout ce qu'openpyxl ne sait pas faire fidèlement
(tableaux croisés dynamiques sous toutes leurs formes, graphiques croisés, segments, recalcul
des formules, export PDF, impression, actualisation de données...) est exécuté **par Excel
lui-même**. L'agent traduit ces instructions VBA une à une vers le même modèle objet Excel,
dans `connecteurs/excel.py` : aucune liste de cas à prévoir, toute fonction qu'une macro
utilise est disponible.

Le module testé `bibliotheque/excel_com.py`, fourni à chaque projet (l'agent ne peut pas le
modifier), apporte la partie commune :
- `session_excel(chemin)` : instance d'Excel séparée et invisible, macros du classeur
  désactivées, chemin complet, enregistrement puis fermeture garantie même en cas d'erreur ;
- `xl` : constantes Excel (`xl.xlRowField`, `xl.xlSum`...) ;
- `EnregistreurCOM` : faux Excel pour les tests, qui note chaque instruction reçue ; les tests
  comparent cette séquence aux instructions de la macro ;
- `preparer_feuille(ws)` : nettoyage côté openpyxl avant qu'Excel ne remplisse la feuille.

Prérequis à l'exécution des scripts traduits : Windows, Excel installé, `pywin32`, fichier
fermé dans Excel. Le test sur Excel réel (`tests/test_excel_com.py::test_excel_reel`) ne
s'exécute que sous Windows.

### Vérification sur Excel réel

Les tests de l'agent remplacent Excel par un enregistreur d'instructions : ils vérifient que la
traduction suit la macro, mais pas qu'Excel accepte chaque instruction. Chaque projet qui pilote
Excel contient donc aussi `tests/test_excel_reel.py` (marqueur `excel_reel`), qui exécute les
étapes Excel pour de vrai sur un classeur fictif :

- quand vba2py tourne sur un poste Windows avec Excel, ces tests s'exécutent automatiquement
  pendant la traduction : l'agent voit l'instruction refusée et la corrige avant de livrer ;
- sinon ils sont ignorés, et la vérification l'indique (« ⚠️ non exécutés ») ; les lancer
  ensuite sous Windows : `python -m pytest -m excel_reel`.

En cas de refus d'Excel, l'erreur `ErreurExcel` indique l'instruction exacte, par exemple :

```
Excel a refusé : wb.PivotCaches().Create(SourceType=1, SourceData='tblDonnees').CreatePivotTable(...)
  Excel a refusé l'opération (0x800A03EC) : nom de feuille, de tableau ou de champ introuvable, ...
  Instructions précédentes : ...
```

## Documentation jointe (facultative)

Des documents décrivant les macros (spécifications, règles métier, exemples de résultats)
peuvent être fournis à l'agent pour préciser l'intention :

```powershell
python main.py traduire examples\Ventes.xlsm -o sortie --doc docs\specification.pdf --doc docs\regles.md
python main.py traduire examples\ -o sortie --doc docs\commun\        # dossier : tous ses .pdf/.md/.txt
```

- `--doc` (répétable, fichier ou dossier) : documentation **commune** à tous les classeurs de la
  commande ;
- un dossier `<NomDuClasseur>_doc/` à côté d'un classeur (ou `<dossier_exporté>_doc/`) :
  documentation **rattachée** à ce seul classeur, ajoutée automatiquement ;
- formats acceptés : `.pdf`, `.md`, `.txt` ;
- **le code VBA prime toujours** : en cas de contradiction, la traduction suit le code et l'écart
  est signalé dans la section « Documentation utilisée » du rapport ; les exemples chiffrés de
  la documentation deviennent des tests (qui vérifient le résultat du code) ;
- comme le VBA, la documentation est une donnée : l'agent ne suit aucune consigne qu'elle
  contiendrait.
