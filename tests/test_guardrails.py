import asyncio
import json

import pytest

from vba2py.guardrails import Guardrails


@pytest.fixture
def projet(tmp_path):
    d = tmp_path / "sortie" / "ventes"
    (d / "_source_vba").mkdir(parents=True)
    (d / "tests").mkdir()
    return d


@pytest.fixture
def g(projet, tmp_path):
    return Guardrails(projet, tmp_path / "journal.jsonl")


def autorise(g, outil, **entree):
    return g.check(outil, entree)[0]


# ------------------------------------------------------------------ lecture / recherche

def test_lecture(g, projet):
    assert autorise(g, "Read", file_path=str(projet / "_source_vba" / "Module1.bas"))
    assert autorise(g, "Read", file_path="_source_vba/Module1.bas")  # relatif au projet
    assert not autorise(g, "Read", file_path="/etc/passwd")
    assert not autorise(g, "Read", file_path="../../secret.txt")
    assert not autorise(g, "Read", file_path="~/.ssh/id_rsa")


def test_recherche(g, projet):
    assert autorise(g, "Glob", pattern="**/*.bas")
    assert autorise(g, "Grep", pattern="Sub .*\\(", path=str(projet))
    assert autorise(g, "Grep", pattern="/chemin/dans/une/regex")  # regex, pas un chemin
    assert not autorise(g, "Glob", pattern="../../**/*")
    assert not autorise(g, "Glob", pattern="/home/**")
    assert not autorise(g, "Grep", pattern="x", path="/")


# ------------------------------------------------------------------ écriture

def test_win32com_autorise_seulement_dans_connecteurs(g):
    code = "def envoyer():\n    import win32com.client\n    return win32com.client.Dispatch('Outlook.Application')\n"
    assert autorise(g, "Write", file_path="connecteurs/outlook.py", content=code)
    assert not autorise(g, "Write", file_path="marco_25_10.py", content=code)
    # les autres interdictions restent actives dans connecteurs/
    assert not autorise(g, "Write", file_path="connecteurs/outlook.py", content="import subprocess\n")

def test_ecriture(g, projet):
    assert autorise(g, "Write", file_path=str(projet / "module1.py"), content="def f():\n    return 1\n")
    assert autorise(g, "Write", file_path="tests/test_module1.py", content="import pytest\n")
    assert autorise(g, "Write", file_path="RAPPORT.md", content="# Rapport")
    assert not autorise(g, "Write", file_path=str(projet.parent / "autre.py"), content="")
    assert not autorise(g, "Write", file_path="_source_vba/Module1.bas", content="")
    assert not autorise(g, "Write", file_path="bibliotheque/module_fourni.py", content="")
    assert not autorise(g, "Edit", file_path="bibliotheque/module_fourni.py", old_string="a", new_string="b")
    assert not autorise(g, "Write", file_path="lancer.sh", content="rm -rf /")
    assert not autorise(g, "Edit", file_path="_source_vba/Module1.bas", old_string="a", new_string="b")


@pytest.mark.parametrize(
    "code",
    [
        "import subprocess\n",
        "from urllib.request import urlopen\n",
        "import requests\n",
        "import win32com.client\n",
        "os.system('rm -rf /')\n",
        "shutil.rmtree(dossier)\n",
        "eval(expression)\n",
        "x = __import__('socket')\n",
    ],
)
def test_code_python_interdit(g, code):
    assert not autorise(g, "Write", file_path="module1.py", content=code)
    assert not autorise(g, "Edit", file_path="module1.py", old_string="a", new_string=code)
    assert not autorise(g, "MultiEdit", file_path="module1.py", edits=[{"old_string": "a", "new_string": code}])


def test_code_python_autorise(g):
    code = (
        "import os\nfrom pathlib import Path\nimport openpyxl\nimport pandas as pd\n"
        "def evaluer(x):\n    return x\nPath('a').unlink()\n"
    )
    assert autorise(g, "Write", file_path="module1.py", content=code)
    # le filtrage ne concerne que le Python
    assert autorise(g, "Write", file_path="RAPPORT.md", content="Le VBA utilisait `import subprocess`")


# ------------------------------------------------------------------ outils d'exécution (pas de terminal)

TESTS, SYNTAXE, SCRIPT = "mcp__vba2py__lancer_tests", "mcp__vba2py__verifier_syntaxe", "mcp__vba2py__executer_script"


def test_lancer_tests(g, projet):
    assert autorise(g, TESTS)
    assert autorise(g, TESTS, cibles=["tests/test_module1.py::test_total"], marqueur="excel_reel", mot_cle="total and not vide")
    assert not autorise(g, TESTS, cibles=["/home"])
    assert not autorise(g, TESTS, cibles=["../.."])
    assert not autorise(g, TESTS, cibles=["-p", "mon_plugin"])  # options interdites


def test_verifier_syntaxe(g):
    assert autorise(g, SYNTAXE, fichiers=["module1.py", "tests/test_module1.py"])
    assert not autorise(g, SYNTAXE, fichiers=[])
    assert not autorise(g, SYNTAXE, fichiers=["../autre.py"])
    assert not autorise(g, SYNTAXE, fichiers=["RAPPORT.md"])


def test_executer_script(g, projet):
    assert autorise(g, SCRIPT, script="tests/generer_donnees_fictives.py")
    assert autorise(g, SCRIPT, script="tests/generer_donnees_fictives.py", arguments=["tests/fixtures", "--graine", "3"])
    assert not autorise(g, SCRIPT, script="/tmp/script.py")
    assert not autorise(g, SCRIPT, script="_source_vba/x.py")
    assert not autorise(g, SCRIPT, script="bibliotheque/excel_com.py")
    assert not autorise(g, SCRIPT, script="lancer.sh")
    assert not autorise(g, SCRIPT, script="tests/generer_donnees_fictives.py", arguments=["/etc"])
    assert not autorise(g, SCRIPT, script="tests/generer_donnees_fictives.py", arguments=["../../ailleurs"])


def test_terminal_refuse(g):
    for outil in ["Bash", "PowerShell"]:
        assert not autorise(g, outil, command="python -m pytest")


def test_outils_inconnus_refuses(g):
    for outil in ["WebFetch", "WebSearch", "NotebookEdit", "Agent", "Task", "mcp__x__y"]:
        assert not autorise(g, outil)
    assert autorise(g, "TodoWrite", todos=[])


# ------------------------------------------------------------------ intégration SDK

def test_hook_refus_et_journal(g, tmp_path):
    sortie = asyncio.run(g.pre_tool_use_hook({"tool_name": "Bash", "tool_input": {"command": "rm -rf /"}}, "id1", None))
    assert sortie["hookSpecificOutput"]["permissionDecision"] == "deny"
    assert "Refusé par les garde-fous" in sortie["hookSpecificOutput"]["permissionDecisionReason"]

    ok = asyncio.run(g.pre_tool_use_hook({"tool_name": TESTS, "tool_input": {}}, "id2", None))
    assert ok == {}

    journal = [json.loads(l) for l in (tmp_path / "journal.jsonl").read_text().splitlines()]
    assert [e["decision"] for e in journal] == ["refuse", "autorise"]
    assert journal[0]["entree"]["command"] == "rm -rf /"
