"""Contrôle et finalisation du livrable produit par l'agent.

- `interface.json` : contrat de la ligne de commande du script traduit. Il est conservé
  d'une traduction à l'autre pour que les commandes restent identiques.
- `RAPPORT.md` : structure imposée par prompts/rapport_modele.md. Les sections factuelles
  (commandes, tests, extraction) sont écrites ici par le code, pas par l'agent.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from .bibliotheque import LIBRARY_DIRNAME, fichiers_modifies
from .extract import SOURCE_DIRNAME

PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"
INTERFACE_FILENAME = "interface.json"
PREVIOUS_INTERFACE_FILENAME = "interface_precedente.json"
REPORT_FILENAME = "RAPPORT.md"


@dataclass
class Verification:
    erreurs: list[str] = field(default_factory=list)
    avertissements: list[str] = field(default_factory=list)
    resultat_tests: str = ""
    resultat_excel: str = ""

    @property
    def ok(self) -> bool:
        return not self.erreurs


# --------------------------------------------------------------------------- interface


def nom_sous_commande(procedure: str) -> str:
    """Nom de sous-commande déduit du nom VBA : « Demo_EcrireEtLire » → « demo-ecrire-et-lire »."""
    return "-".join(_mots(procedure.split(".")[-1]))


def nom_script(module: str) -> str:
    """Nom du fichier Python d'un module VBA : « CréationCSV » → « creation_csv.py »."""
    return "_".join(_mots(module)) + ".py"


def _mots(nom: str) -> list[str]:
    import unicodedata

    ascii_nom = unicodedata.normalize("NFKD", nom).encode("ascii", "ignore").decode()
    mots = []
    for partie in re.split(r"[^0-9A-Za-z]+", ascii_nom):
        mots += re.split(r"(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", partie)
    return [m.lower() for m in mots if m]


_SUB_LANCABLE = re.compile(r"^[ \t]*(?:Public[ \t]+)?Sub[ \t]+(\w+)[ \t]*\([ \t]*\)", re.MULTILINE | re.IGNORECASE)


def macros_lancables(project_dir: Path) -> list[str]:
    """« Module.Macro » pour chaque Sub publique sans paramètre d'un module standard (.bas).

    Ce sont les macros qu'Excel propose dans Alt+F8 : chacune doit devenir un point d'entrée ou
    être déclarée comme procédure interne dans interface.json.
    """
    resultat = []
    for fichier in sorted((project_dir / SOURCE_DIRNAME).glob("*.bas")):
        code = fichier.read_text(encoding="utf-8")
        nom = re.search(r'^Attribute VB_Name = "([^"]+)"', code, re.MULTILINE)
        module = nom.group(1) if nom else fichier.stem
        if re.search(r"^\s*Option\s+Private\s+Module", code, re.MULTILINE | re.IGNORECASE):
            continue  # module privé : ses macros n'apparaissent pas dans Alt+F8
        resultat += [f"{module}.{m}" for m in _SUB_LANCABLE.findall(code)]
    return resultat


def macros_boutons(project_dir: Path) -> list[str]:
    structure = project_dir / SOURCE_DIRNAME / "structure_classeur.json"
    if not structure.exists():
        return []
    return [b["macro"] for b in json.loads(structure.read_text(encoding="utf-8")).get("boutons", [])]


def _meme_macro(a: str, b: str) -> bool:
    """« Module.Macro » et « Macro » désignent la même macro (comparaison insensible à la casse, comme VBA)."""
    a, b = a.lower(), b.lower()
    return a == b or a.split(".")[-1] == b.split(".")[-1] and ("." not in a or "." not in b)


def charger_interface(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if "points_entree" not in data and "script" in data:  # ancien format : un seul script
        data = {
            "points_entree": [{
                "script": data["script"],
                "sous_commande": None,
                "procedure": data.get("procedure_principale", ""),
                "justification": "contrat converti depuis l'ancien format",
                "fixture": data.get("fixture_principale"),
                "arguments": data.get("arguments", []),
            }],
            "procedures_internes": [],
        }
    if not isinstance(data.get("points_entree"), list) or not data["points_entree"]:
        raise ValueError(f"{path.name} : « points_entree » doit être une liste non vide")
    for point in data["points_entree"]:
        for cle in ("script", "procedure", "arguments"):
            if cle not in point:
                raise ValueError(f"{path.name} : clé « {cle} » manquante dans un point d'entrée")
        point.setdefault("sous_commande", None)
        for arg in point["arguments"]:
            if "nom" not in arg or "remplace" not in arg:
                raise ValueError(f"{path.name} : chaque argument doit avoir « nom » et « remplace »")
    data.setdefault("procedures_internes", [])
    return data


def noms_arguments(point: dict) -> list[str]:
    return [a["nom"] for a in point["arguments"]]


def aide_du_script(project_dir: Path, script: str, sous_commande: str | None = None) -> str | None:
    """Sortie de `python script [sous_commande] --help` (None si la commande échoue)."""
    commande = [sys.executable, script] + ([sous_commande] if sous_commande else []) + ["--help"]
    result = subprocess.run(commande, cwd=project_dir, capture_output=True, text=True, timeout=60)
    return result.stdout if result.returncode == 0 else None


def options_de_l_aide(aide: str) -> set[str]:
    return set(re.findall(r"(?<![\w-])(--[a-z0-9][a-z0-9-]*)", aide)) - {"--help"}


def _cle(point: dict) -> str:
    return f"{point['script']} {point['sous_commande'] or ''}".strip()


def verifier_interface(project_dir: Path, interface: dict, precedente: dict | None) -> list[str]:
    erreurs = []
    points = interface["points_entree"]
    internes = [p.get("procedure", "") for p in interface["procedures_internes"]]

    # --- chaque macro lançable est un point d'entrée ou une procédure interne déclarée
    lancables = macros_lancables(project_dir)
    for macro in lancables:
        declare = [p for p in points if _meme_macro(p["procedure"], macro)] + [i for i in internes if _meme_macro(i, macro)]
        if not declare:
            erreurs.append(f"macro {macro} : ni point d'entrée, ni procédure interne déclarée dans interface.json")
    for procedure in [p["procedure"] for p in points] + internes:
        if lancables and not any(_meme_macro(procedure, m) for m in lancables):
            erreurs.append(f"interface.json : {procedure} n'est pas une macro lançable du VBA (Sub publique sans paramètre)")
    for macro in macros_boutons(project_dir):
        if not any(_meme_macro(p["procedure"], macro) for p in points):
            erreurs.append(f"macro {macro} reliée à un bouton : elle doit être un point d'entrée")

    # --- noms déduits du VBA, sous-commandes cohérentes
    par_script: dict[str, list[dict]] = {}
    for point in points:
        par_script.setdefault(point["script"], []).append(point)
        if "." not in point["procedure"]:
            erreurs.append(f"{point['procedure']} : écrire la procédure sous la forme « Module.Macro »")
            continue
        module = point["procedure"].split(".")[0]
        if point["script"] != nom_script(module):
            erreurs.append(f"{point['procedure']} : le script doit s'appeler {nom_script(module)} (et non {point['script']})")
    for script, groupe in par_script.items():
        if len(groupe) == 1 and groupe[0]["sous_commande"]:
            erreurs.append(f"{script} : un seul point d'entrée, il ne doit pas y avoir de sous-commande")
        if len(groupe) > 1:
            for point in groupe:
                attendu = nom_sous_commande(point["procedure"])
                if point["sous_commande"] != attendu:
                    erreurs.append(f"{point['procedure']} : la sous-commande doit être « {attendu} »")

    # --- arguments et --help
    for point in points:
        script, sc = point["script"], point["sous_commande"]
        if not (project_dir / script).exists():
            erreurs.append(f"interface.json : script introuvable ({script})")
            continue
        fixture = point.get("fixture")
        if fixture and not (project_dir / fixture).exists():
            erreurs.append(f"{_cle(point)} : fixture introuvable ({fixture})")
        noms = noms_arguments(point)
        positionnels = [n for n in noms if not n.startswith("-")]
        if positionnels and positionnels != ["classeur"] or ("classeur" in noms and noms[0] != "classeur"):
            erreurs.append(f"{_cle(point)} : seul « classeur » peut être positionnel, en premier")
        aide = aide_du_script(project_dir, script, sc)
        if aide is None:
            erreurs.append(f"`python {_cle(point)} --help` a échoué")
            continue
        declarees, options = {n for n in noms if n.startswith("--")}, options_de_l_aide(aide)
        if declarees - options:
            erreurs.append(f"{_cle(point)} : options de interface.json absentes du script : {sorted(declarees - options)}")
        if options - declarees:
            erreurs.append(f"{_cle(point)} : options du script absentes de interface.json : {sorted(options - declarees)}")

    # --- stabilité par rapport à la traduction précédente
    if precedente is not None:
        actuels = {p["procedure"].lower(): p for p in points}
        for ancien in precedente["points_entree"]:
            nouveau = actuels.get(ancien["procedure"].lower())
            if nouveau is None:
                erreurs.append(f"point d'entrée {ancien['procedure']} présent dans la traduction précédente : supprimé")
            elif _cle(nouveau) != _cle(ancien) or noms_arguments(nouveau) != noms_arguments(ancien):
                erreurs.append(
                    f"{ancien['procedure']} : la commande a changé depuis la traduction précédente "
                    f"({_cle(ancien)} {noms_arguments(ancien)} → {_cle(nouveau)} {noms_arguments(nouveau)})"
                )
    return erreurs


def commandes(interface: dict) -> str:
    """Section « Commandes » du rapport, générée de façon déterministe depuis interface.json."""
    lignes = [
        "```powershell",
        "pip install -r requirements.txt",
        "python tests/generer_donnees_fictives.py",
        "python -m pytest -q",
        "```",
    ]
    for point in interface["points_entree"]:
        noms = noms_arguments(point)
        appel = f"python {_cle(point)}"
        classeur_fictif = f" {point['fixture']}" if "classeur" in noms and point.get("fixture") else ""
        classeur_reel = ' "C:\\chemin\\vers\\classeur.xlsm"' if "classeur" in noms else ""
        dossier = " --dossier-sortie sortie_test" if "--dossier-sortie" in noms else ""
        dossier_reel = ' --dossier-sortie "C:\\chemin\\vers\\sorties"' if "--dossier-sortie" in noms else ""
        mail_test = " --mode-mail aucun" if "--mode-mail" in noms else ""
        mail_reel = " --mode-mail afficher" if "--mode-mail" in noms else ""
        lignes += [
            "",
            f"### {point['procedure']}",
            "",
            f"Pourquoi une commande : {point.get('justification') or '—'}",
            "",
            "```powershell",
            f"{appel} --help",
            f"{appel}{classeur_fictif}{dossier}{mail_test}        # essai sur données fictives",
            f"{appel}{classeur_reel}{dossier_reel}{mail_reel}        # sur un vrai classeur",
            "```",
        ]
        if point["arguments"]:
            lignes += ["", "| Argument | Remplace dans le VBA | Valeur par défaut |", "|---|---|---|"]
            for arg in point["arguments"]:
                defaut = arg.get("defaut")
                defaut = "—" if defaut in (None, "") else f"`{defaut}`"
                lignes.append(f"| `{arg['nom']}` | {arg['remplace']} | {defaut} |")
    if interface["procedures_internes"]:
        lignes += ["", "### Procédures internes (sans commande)", ""]
        lignes += [f"- {p['procedure']} : {p.get('raison', '')}" for p in interface["procedures_internes"]]
    return "\n".join(lignes)


# --------------------------------------------------------------------------- rapport


def titres_attendus() -> list[str]:
    modele = (PROMPTS_DIR / "rapport_modele.md").read_text(encoding="utf-8")
    return re.findall(r"^## .+$", modele, re.MULTILINE)


def verifier_structure_rapport(rapport: str) -> list[str]:
    titres = re.findall(r"^## .+$", rapport, re.MULTILINE)
    attendus = titres_attendus()
    if titres != attendus:
        manquants = [t for t in attendus if t not in titres]
        en_trop = [t for t in titres if t not in attendus]
        detail = []
        if manquants:
            detail.append(f"manquants : {manquants}")
        if en_trop:
            detail.append(f"non prévus : {en_trop}")
        if not detail:
            detail.append("ordre différent du modèle")
        return ["RAPPORT.md ne suit pas le modèle (" + " ; ".join(detail) + ")"]
    return []


def remplir_section(rapport: str, nom: str, contenu: str) -> str:
    debut, fin = f"<!-- vba2py:debut {nom} -->", f"<!-- vba2py:fin {nom} -->"
    pattern = re.compile(re.escape(debut) + r".*?" + re.escape(fin), re.DOTALL)
    bloc = f"{debut}\n{contenu.strip()}\n{fin}"
    if pattern.search(rapport):
        return pattern.sub(lambda _: bloc, rapport)
    return rapport


def section_extraction(project_dir: Path) -> str:
    extraction = project_dir / SOURCE_DIRNAME / "EXTRACTION.md"
    if not extraction.exists():
        return "Rapport d'extraction absent."
    texte = extraction.read_text(encoding="utf-8")
    return re.sub(r"^# .*\n", "", texte, count=1).strip()


# --------------------------------------------------------------------------- Excel réel

MARQUEUR_EXCEL = "excel_reel"


def verifier_excel_reel(project_dir: Path, verif: Verification) -> None:
    """Les étapes de connecteurs/excel.py doivent être testées sur un vrai Excel.

    Sans Excel (Linux, poste sans Office), ces tests sont ignorés : simple avertissement.
    """
    result = subprocess.run(
        [sys.executable, "-m", "pytest", "-q", "-m", MARQUEUR_EXCEL, "-rs", "tests"],
        cwd=project_dir, capture_output=True, text=True,
    )
    derniere = (result.stdout.strip().splitlines() or [""])[-1]
    if result.returncode == 5:  # aucun test sélectionné
        verif.erreurs.append(
            f"connecteurs/excel.py existe mais aucun test marqué `{MARQUEUR_EXCEL}` ne l'exécute sur un vrai Excel"
        )
        verif.resultat_excel = "❌ aucun test"
    elif result.returncode != 0:
        verif.erreurs.append(f"tests sur Excel réel en échec : {derniere}\n{result.stdout[-3000:]}")
        verif.resultat_excel = f"❌ **{derniere}**"
    elif "passed" in derniere:
        verif.resultat_excel = f"✅ **{derniere}**"
    else:
        verif.avertissements.append(
            "étapes Excel non vérifiées sur un vrai Excel (indisponible sur cette machine) : "
            f"lancer `python -m pytest -m {MARQUEUR_EXCEL}` sous Windows avec Excel"
        )
        verif.resultat_excel = "⚠️ non exécutés (Excel indisponible sur cette machine) — à lancer sous Windows avec Excel"


# --------------------------------------------------------------------------- vérification complète


def verifier_projet(project_dir: Path, interface_precedente: Path | None = None) -> Verification:
    verif = Verification()
    manquants = [f for f in (REPORT_FILENAME, "requirements.txt", "tests", INTERFACE_FILENAME) if not (project_dir / f).exists()]
    if manquants:
        verif.erreurs.append(f"éléments manquants : {', '.join(manquants)}")

    modifies = fichiers_modifies(project_dir)
    if modifies:
        verif.erreurs.append(f"{LIBRARY_DIRNAME}/ modifiée alors qu'elle est fournie par vba2py : {modifies}")

    interface = None
    if (project_dir / INTERFACE_FILENAME).exists():
        try:
            interface = charger_interface(project_dir / INTERFACE_FILENAME)
            precedente = charger_interface(interface_precedente) if interface_precedente and interface_precedente.exists() else None
            verif.erreurs += verifier_interface(project_dir, interface, precedente)
        except (ValueError, json.JSONDecodeError) as exc:
            verif.erreurs.append(f"interface.json invalide : {exc}")

    if (project_dir / "tests").exists():
        result = subprocess.run(
            [sys.executable, "-m", "pytest", "-q", "tests"], cwd=project_dir, capture_output=True, text=True
        )
        derniere = (result.stdout.strip().splitlines() or [""])[-1]
        verif.resultat_tests = derniere
        if result.returncode != 0:
            verif.erreurs.append(f"tests en échec : {derniere}\n{result.stdout[-3000:]}")
        if (project_dir / "connecteurs" / "excel.py").exists():
            verifier_excel_reel(project_dir, verif)

    rapport_path = project_dir / REPORT_FILENAME
    if rapport_path.exists():
        rapport = rapport_path.read_text(encoding="utf-8")
        verif.erreurs += verifier_structure_rapport(rapport)
        statut = "✅" if not any(e.startswith("tests en échec") for e in verif.erreurs) else "❌"
        rapport = remplir_section(
            rapport,
            "tests",
            f"Vérification indépendante par vba2py : `python -m pytest -q tests` → {statut} **{verif.resultat_tests}**",
        )
        if verif.resultat_excel:
            rapport = remplir_section(
                rapport, "tests",
                f"Vérification indépendante par vba2py : `python -m pytest -q tests` → {statut} "
                f"**{verif.resultat_tests}**\n\nTests sur Excel réel (`python -m pytest -m {MARQUEUR_EXCEL}`) : "
                f"{verif.resultat_excel}",
            )
        rapport = remplir_section(rapport, "extraction", section_extraction(project_dir))
        if interface is not None:
            rapport = remplir_section(rapport, "commandes", commandes(interface))
        rapport_path.write_text(rapport, encoding="utf-8")
    return verif
