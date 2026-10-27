"""bibliotheque/dialogues.py : MsgBox et InputBox dans le terminal, et réponses simulées pour les tests."""

import pytest

from vba2py.bibliotheque import dialogues as d


def test_simulation_dans_l_ordre_et_journal():
    with d.simuler(["oui", "Sud", "", "!annuler"]) as journal:
        d.msgbox("Traitement terminé", d.vbInformation, "Info")  # sans choix : ne consomme pas de réponse
        assert d.msgbox("Continuer ?", d.vbYesNo + d.vbQuestion) == d.vbYes
        assert d.inputbox("Region ?", "Filtre", "Nord") == "Sud"
        assert d.inputbox("Region ?", "Filtre", "Nord") == "Nord"  # Entrée seule : valeur par défaut
        assert d.inputbox("Region ?") == ""  # Annuler
    assert [(j.genre, j.texte, j.reponse) for j in journal] == [
        ("msgbox", "Traitement terminé", d.vbOK),
        ("msgbox", "Continuer ?", d.vbYes),
        ("inputbox", "Region ?", "Sud"),
        ("inputbox", "Region ?", "Nord"),
        ("inputbox", "Region ?", ""),
    ]


@pytest.mark.parametrize("reponse, attendu", [("non", d.vbNo), ("N", d.vbNo), ("annuler", d.vbCancel), ("o", d.vbYes)])
def test_boutons_et_synonymes(reponse, attendu):
    with d.simuler([reponse]):
        assert d.msgbox("Lire ?", d.vbYesNoCancel) == attendu


def test_reponse_impossible_ou_manquante():
    with d.simuler(["peut-être"]):
        with pytest.raises(ValueError):
            d.msgbox("Lire ?", d.vbYesNo)
    with d.simuler([]):
        with pytest.raises(d.AucuneReponsePrevue):
            d.inputbox("Ligne ?")


def test_application_inputbox_types():
    with d.simuler(["12,5", "", "vrai", "!annuler"]):
        assert d.application_inputbox("Nombre ?", type_=1) == 12.5
        assert d.application_inputbox("Nombre ?", defaut="3", type_=1) == 3
        assert d.application_inputbox("Booléen ?", type_=4) is True
        assert d.application_inputbox("Nombre ?", type_=1) is False  # Annuler


def test_terminal_redemande_jusqu_a_une_reponse_valide(monkeypatch, capsys):
    saisies = iter(["peut-être", "non", "x", "", "abc", "7"])
    monkeypatch.setattr("builtins.input", lambda invite="": next(saisies))
    assert d.msgbox("Continuer ?", d.vbYesNo, "Mainframe") == d.vbNo
    assert d.inputbox("Texte ?") == "x"
    assert d.inputbox("Texte ?", defaut="defaut") == "defaut"
    assert d.application_inputbox("Ligne ?", type_=1) == 7  # « abc » refusé, redemandé
    sortie = capsys.readouterr().out
    assert "[Mainframe]" in sortie and "Répondre par : oui, non" in sortie and "Valeur incorrecte" in sortie


def test_terminal_fin_de_saisie_vaut_annuler(monkeypatch):
    def fin(invite=""):
        raise EOFError
    monkeypatch.setattr("builtins.input", fin)
    assert d.inputbox("Texte ?", defaut="x") == ""
    assert d.msgbox("Continuer ?", d.vbYesNoCancel) == d.vbCancel
    assert d.application_inputbox("Nombre ?", type_=1) is False
