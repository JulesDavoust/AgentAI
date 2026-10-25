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

def test_ecriture(g, projet):
    assert autorise(g, "Write", file_path=str(projet / "module1.py"), content="def f():\n    return 1\n")
    assert autorise(g, "Write", file_path="tests/test_module1.py", content="import pytest\n")
    assert autorise(g, "Write", file_path="RAPPORT.md", content="# Rapport")
    assert not autorise(g, "Write", file_path=str(projet.parent / "autre.py"), content="")
    assert not autorise(g, "Write", file_path="_source_vba/Module1.bas", content="")
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


# ------------------------------------------------------------------ commandes

@pytest.mark.parametrize(
    "commande",
    [
        "python -m pytest -q",
        "python3 -m pytest tests/test_module1.py::test_total -x",
        "python -m pytest -q -k 'total and not vide'",
        "pytest tests",
        "python -m py_compile module1.py tests/test_module1.py",
        "python tests/generer_donnees_fictives.py",
    ],
)
def test_commandes_autorisees(g, commande):
    assert autorise(g, "Bash", command=commande), g.check("Bash", {"command": commande})


@pytest.mark.parametrize(
    "commande",
    [
        "rm -rf tests",
        "curl https://exemple.com",
        "pip install xlwings",
        "python -m pytest && rm -rf /",
        "python -m pytest; curl x",
        "python -m pytest | tee log",
        "python -m pytest > /tmp/log",
        "echo $(cat /etc/passwd)",
        "python -c 'import os'",
        "python /tmp/script.py",
        "python _source_vba/x.py",
        "python -m pytest /home",
        "python -m pytest ../..",
        "python -m pytest -p mon_plugin",
        "python -m pytest --rootdir=/",
        "cd .. ",
        "git status",
    ],
)
def test_commandes_refusees(g, commande):
    assert not autorise(g, "Bash", command=commande)


def test_bash_en_arriere_plan_refuse(g):
    assert not autorise(g, "Bash", command="python -m pytest", run_in_background=True)


def test_outils_inconnus_refuses(g):
    for outil in ["WebFetch", "WebSearch", "NotebookEdit", "Agent", "Task", "mcp__x__y"]:
        assert not autorise(g, outil)
    assert autorise(g, "TodoWrite", todos=[])


# ------------------------------------------------------------------ intégration SDK

def test_hook_refus_et_journal(g, tmp_path):
    sortie = asyncio.run(g.pre_tool_use_hook({"tool_name": "Bash", "tool_input": {"command": "rm -rf /"}}, "id1", None))
    assert sortie["hookSpecificOutput"]["permissionDecision"] == "deny"
    assert "Refusé par les garde-fous" in sortie["hookSpecificOutput"]["permissionDecisionReason"]

    ok = asyncio.run(g.pre_tool_use_hook({"tool_name": "Bash", "tool_input": {"command": "python -m pytest"}}, "id2", None))
    assert ok == {}

    journal = [json.loads(l) for l in (tmp_path / "journal.jsonl").read_text().splitlines()]
    assert [e["decision"] for e in journal] == ["refuse", "autorise"]
    assert journal[0]["entree"]["command"] == "rm -rf /"
