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
