# AgentAI — vba2py

Agent qui traduit des macros VBA Excel en Python et livre, pour chaque classeur, le code
traduit, une batterie de tests sur données fictives et un rapport.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate          # Windows : .venv\Scripts\activate
pip install -r requirements.txt
export ANTHROPIC_API_KEY=...       # clé créée sur console.anthropic.com
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
    ├── RAPPORT.md                choix de traduction, TODO, résultats des tests
    └── tests/
        ├── generer_donnees_fictives.py
        ├── fixtures/             classeurs fictifs, à ouvrir aussi dans Excel pour comparer avec la macro d'origine
        └── test_module1.py
```

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
