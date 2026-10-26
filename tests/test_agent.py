"""Vérifie la configuration de l'agent, sans appeler l'API."""

from vba2py.agent import AGENT_TOOLS, build_options, build_system_prompt


def test_options_de_l_agent(tmp_path):
    options = build_options(tmp_path, tmp_path / "journal.jsonl", model="claude-sonnet-5-5", max_turns=10, max_budget_usd=1.0)
    assert options.tools == AGENT_TOOLS
    assert "Bash" not in options.tools and {"Bash", "PowerShell", "WebFetch"} <= set(options.disallowed_tools)
    assert list(options.mcp_servers) == ["vba2py"]
    assert "mcp__vba2py__lancer_tests" in options.allowed_tools
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
    assert '"points_entree"' in prompt and "Application.Run" in prompt and 'encoding="cp1252"' in prompt


def test_estimation_du_budget(tmp_path):
    from vba2py.agent import estimer_budget, lignes_de_vba

    source = tmp_path / "_source_vba"
    source.mkdir()
    (source / "Module1.bas").write_text('Attribute VB_Name = "Module1"\n\' commentaire\nSub A()\n\nEnd Sub\n', encoding="utf-8")
    (source / "EXTRACTION.md").write_text("ignoré\n", encoding="utf-8")
    assert lignes_de_vba(tmp_path) == 2
    assert estimer_budget(1000, "claude-sonnet-5-5") == 2.5
    assert estimer_budget(1000, "claude-opus-5-5") == 5.0
    assert estimer_budget(0, "modele-inconnu") == 2.0


def test_traduire_previent_si_budget_insuffisant(monkeypatch, tmp_path, capsys):
    import sys

    import main
    import vba2py.agent as agent

    sys.path.insert(0, str(main.Path(__file__).parent))
    from conftest import make_xlsm

    make_xlsm(tmp_path / "Ventes.xlsm")

    async def faux_agent(*args, **kwargs):
        return agent.AgentRun(success=True, summary="fini", turns=1, cost_usd=0.1)

    monkeypatch.setattr(agent, "translate_project", faux_agent)
    monkeypatch.setattr(agent, "estimer_budget", lambda nb, modele: 9.0)
    main.main(["traduire", str(tmp_path / "Ventes.xlsm"), "-o", str(tmp_path / "s"), "--budget", "2"])
    sortie = capsys.readouterr().out
    assert "Budget probablement insuffisant : environ 9.00 $ conseillés" in sortie
    assert "--budget 9 ou plus" in sortie


def test_ligne_de_commande_sous_la_limite_windows(tmp_path):
    """Régression : passées en argument, les instructions dépassaient 32 767 caractères (WinError 206)."""
    from claude_agent_sdk._internal.transport.subprocess_cli import SubprocessCLITransport

    from vba2py.agent import build_system_prompt, ecrire_prompt_systeme

    fichier = ecrire_prompt_systeme(tmp_path)
    assert fichier.read_text(encoding="utf-8") == build_system_prompt()
    options = build_options(tmp_path, None, fichier_prompt=fichier)
    assert "append" not in options.system_prompt
    transport = SubprocessCLITransport(prompt="x", options=options)
    transport._cli_path = "claude"
    commande = transport._build_command()
    assert commande[commande.index("--append-system-prompt-file") + 1] == str(fichier)
    assert len(" ".join(commande)) < 8_000  # très loin de la limite, même si les instructions grossissent
