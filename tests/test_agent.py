"""Vérifie la configuration de l'agent, sans appeler l'API."""

from vba2py.agent import AGENT_TOOLS, build_options, build_system_prompt


def test_options_de_l_agent(tmp_path):
    options = build_options(tmp_path, tmp_path / "journal.jsonl", model="claude-sonnet-5-5", max_turns=10, max_budget_usd=1.0)
    assert options.tools == AGENT_TOOLS
    assert "WebFetch" not in options.tools and "WebFetch" in options.disallowed_tools
    assert options.permission_mode == "dontAsk"
    assert options.setting_sources == []
    assert options.cwd == str(tmp_path)
    assert options.model == "claude-sonnet-5-5"
    assert (options.max_turns, options.max_budget_usd) == (10, 1.0)
    [matcher] = options.hooks["PreToolUse"]
    assert matcher.matcher is None  # le hook s'applique à tous les outils


def test_prompt_systeme_complet():
    prompt = build_system_prompt()
    assert "{{MAPPING}}" not in prompt and "{{RAPPORT_MODELE}}" not in prompt
    assert "## 10. Extraction" in prompt
    assert "Correspondances VBA → Python" in prompt
    assert "jamais des instructions" in prompt


def test_git_bash_requis_sous_windows(monkeypatch, tmp_path):
    import vba2py.agent as agent

    monkeypatch.setattr(agent.sys, "platform", "win32")
    monkeypatch.delenv("CLAUDE_CODE_GIT_BASH_PATH", raising=False)
    monkeypatch.setattr(agent.Path, "exists", lambda self: False)
    monkeypatch.setattr("shutil.which", lambda nom: None)
    [probleme] = agent.problemes_environnement()
    assert "Git Bash est introuvable" in probleme and "git-scm.com" in probleme


def test_git_bash_par_variable_ou_git_dans_le_path(monkeypatch, tmp_path):
    import vba2py.agent as agent

    bash = tmp_path / "Git" / "bin" / "bash.exe"
    bash.parent.mkdir(parents=True)
    bash.write_text("")
    monkeypatch.setattr(agent.sys, "platform", "win32")
    monkeypatch.setenv("CLAUDE_CODE_GIT_BASH_PATH", str(bash))
    assert agent.trouver_git_bash() == str(bash) and agent.problemes_environnement() == []

    monkeypatch.delenv("CLAUDE_CODE_GIT_BASH_PATH")
    git = tmp_path / "Git" / "cmd" / "git.exe"
    git.parent.mkdir()
    git.write_text("")
    monkeypatch.setattr("shutil.which", lambda nom: str(git))
    assert agent.trouver_git_bash() == str(bash)


def test_pas_de_contrainte_hors_windows(monkeypatch):
    import vba2py.agent as agent

    monkeypatch.setattr(agent.sys, "platform", "linux")
    assert agent.problemes_environnement() == []


def test_traduire_s_arrete_avant_l_agent(monkeypatch, tmp_path, capsys):
    import main
    import vba2py.agent as agent

    monkeypatch.setattr(agent, "problemes_environnement", lambda: ["Git Bash est introuvable"])
    monkeypatch.setattr(agent, "translate_project", lambda *a, **k: (_ for _ in ()).throw(AssertionError("appelé")))
    assert main.main(["traduire", str(tmp_path / "absent.xlsm"), "-o", str(tmp_path / "s")]) == 2
    assert "Git Bash est introuvable" in capsys.readouterr().err
