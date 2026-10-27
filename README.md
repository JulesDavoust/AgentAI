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
- **Commandes** : chaque macro que l'utilisateur lance (reliée à un bouton, ou non appelée par
  une autre procédure) devient une commande Python, qui remplace aussi les boutons. Un script
  par module VBA (`CréationCSV` → `creation_csv.py`) et, s'il contient plusieurs macros, une
  sous-commande par macro (`Demo_EcrireEtLire` → `demo-ecrire-et-lire`). Les noms de scripts,
  de sous-commandes et d'arguments sont déduits du VBA par des règles fixes, et décrits dans
  `interface.json` ; chaque macro lançable (`Sub` publique sans paramètre) y est soit un point
  d'entrée, soit une procédure interne justifiée. Après une traduction réussie, ce contrat est enregistré dans
  `interfaces/<projet>.json` (à committer) ; les traductions suivantes du même projet doivent
  le respecter.
- **Vérification indépendante** (`traduire` et `verifier`) : tests relancés, `--help` du script
  comparé à `interface.json`, structure du rapport contrôlée.

## Sécurité

- Le VBA est traité comme une donnée : l'agent a pour consigne de ne jamais suivre
  d'instructions trouvées dans le code ou ses commentaires.
- Chaque action de l'agent est contrôlée par du code (`vba2py/guardrails.py`) avant exécution :
  lecture et écriture limitées au dossier du projet, `_source_vba/` en lecture seule,
  pas d'accès web, imports dangereux refusés dans le code écrit (réseau, sous-processus,
  `eval`, COM Windows).
- **Aucun terminal** : l'agent ne tape aucune commande. Il dispose de trois outils fournis par
  vba2py (`vba2py/outils.py`), que vba2py exécute lui-même avec son propre Python, sans shell :
  `lancer_tests` (pytest, éventuellement sur des cibles, un marqueur ou un mot-clé),
  `verifier_syntaxe` (fichiers `.py`) et `executer_script` (un script `.py` du projet). Leurs
  arguments sont contrôlés par les garde-fous. Aucun prérequis de terminal : pas besoin de Git
  Bash sous Windows.
- Les réglages de la machine (`~/.claude`, CLAUDE.md, skills) ne sont pas chargés.
- Limites par projet : nombre de tours (`--max-tours`) et budget en dollars (`--budget`).
  Avant chaque traduction, vba2py compte les lignes de VBA et prévient si le budget paraît
  insuffisant pour le modèle choisi (estimation grossière : environ 1 $ + 1,5 $ par millier de
  lignes avec Sonnet, le double avec Opus) ; un agent à court de budget s'arrête avant d'avoir
  corrigé ses tests.
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

## Le classeur de la macro : toujours dans Excel

Tout ce que la macro fait dans `ThisWorkbook` est traduit instruction par instruction vers le
même modèle objet Excel, et exécuté **par Excel** (`bibliotheque/excel_com.py`, fourni à chaque
projet). openpyxl ne sert plus qu'aux tests (classeurs fictifs, relecture des résultats).

- **Classeur ouvert dans Excel** : la commande agit directement dessus, comme la macro ; tu vois
  le résultat, tu peux saisir dans les cellules (écran simulé, par exemple) et tu enregistres
  quand tu veux. La commande n'enregistre jamais un classeur ouvert.
- **Classeur fermé** : il est ouvert dans un Excel invisible (macros désactivées), modifié, puis
  **enregistré** — comme « lancer la macro puis enregistrer » : feuilles, Log et historique
  s'accumulent d'un lancement à l'autre, boutons et macros sont conservés.
  `--sans-enregistrer` : modifications faites sans enregistrer le classeur.
- Un classeur ouvert ailleurs (autre utilisateur, synchronisation) est signalé clairement au
  lieu d'une erreur « Permission denied ».
- Les éléments propres à Excel (tableaux croisés, graphiques, segments, recalcul des formules,
  export PDF...) suivent naturellement : ce sont les mêmes instructions que dans la macro.

Prérequis à l'exécution des scripts traduits : Windows, Excel installé, `pywin32`.

## Boîtes de dialogue

`MsgBox` et `InputBox` deviennent des questions dans le terminal (`bibliotheque/dialogues.py`),
avec les mêmes boutons, la même valeur renvoyée et la même validation que VBA :

```
[Mainframe - Lire / Modifier]
Que veux-tu faire sur le mainframe ?
oui / non / annuler ? non
```

Entrée seule = valeur par défaut de l'`InputBox` ; `!annuler` = bouton Annuler. Quand une
commande a une option pour une saisie (ex. `--saisie Nord`), elle remplace la question. Les tests
fournissent les réponses à l'avance (`dialogues.simuler([...])`).

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
