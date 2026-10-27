"""Cohérence de interface.json avec les scripts et avec le VBA d'origine."""

import json
import re
from pathlib import Path

import pytest

import creation_csv
import macro_carte
import macro_initiale
import simu_mainframe
import test_interface_mvs
import utilisation_csv

RACINE = Path(__file__).resolve().parent.parent
INTERFACE = json.loads((RACINE / "interface.json").read_text(encoding="utf-8"))
MODULES = {
    "creation_csv.py": creation_csv, "macro_carte.py": macro_carte, "macro_initiale.py": macro_initiale,
    "simu_mainframe.py": simu_mainframe, "test_interface_mvs.py": test_interface_mvs,
    "utilisation_csv.py": utilisation_csv,
}


@pytest.mark.parametrize("entree", INTERFACE["points_entree"], ids=lambda e: e["procedure"])
def test_options_declarees_par_argparse(entree, capsys):
    module = MODULES[entree["script"]]
    args = ([entree["sous_commande"]] if entree["sous_commande"] else []) + ["--help"]
    with pytest.raises(SystemExit):
        module.main(args)
    aide = capsys.readouterr().out
    options = {o for o in re.findall(r"--[a-z][a-z0-9-]*", aide) if o != "--help"}
    attendues = {a["nom"] for a in entree["arguments"] if a["nom"].startswith("--")}
    assert options == attendues
    assert (RACINE / entree["fixture"]).exists()
    assert entree["arguments"][0]["nom"] == "classeur"


def test_toutes_les_macros_publiques_sont_classees():
    source = RACINE / "_source_vba"
    if not source.exists():
        pytest.skip("_source_vba absent")
    macros = set()
    for bas in source.glob("*.bas"):
        texte = bas.read_text(encoding="utf-8", errors="replace")
        for m in re.finditer(r"^Public Sub (\w+)\(\)", texte, re.M):
            macros.add(f"{bas.stem}.{m.group(1)}")
    classees = {e["procedure"] for e in INTERFACE["points_entree"]} | {
        e["procedure"] for e in INTERFACE["procedures_internes"]}
    assert macros == classees


def test_boutons_sont_des_points_d_entree():
    boutons = {"MacroInitiale.LancerPipelineComplet", "MacroCarte.LancerPipelineCarteV2",
               "CréationCSV.GenererCSVDuJourV3", "UtilisationCSV.LancerPipelineCSVV3",
               "SimuMainframe.OuvrirMainframe", "TestInterfaceMVS.InteragirMainframe"}
    assert boutons <= {e["procedure"] for e in INTERFACE["points_entree"]}
