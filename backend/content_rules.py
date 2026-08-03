from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Mapping


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


# V10 : les reformulations de questions (ex-recouvrements B01/B09 et B06B/B12)
# sont désormais écrites directement dans la feuille 50_Questions de l'Excel
# de paramétrage. Il n'y a plus de dictionnaire codé en dur ici : modifier une
# question dans 50_Questions suffit à changer le comportement de l'outil,
# sans toucher au code. Ce dict est conservé vide pour compatibilité API.
QUESTION_OVERRIDES: Dict[str, Dict[str, Any]] = {}


FALLBACK_GLOSSARY: Dict[str, Dict[str, Any]] = {
    "BIM": {"definition": "Méthode collaborative fondée sur une maquette numérique enrichie de données.", "aliases": ["Building Information Modeling"]},
    "Convention BIM": {"definition": "Document contractuel qui fixe les règles BIM du projet : rôles, logiciels, formats, nommage, échanges, jalons et livrables.", "aliases": ["charte BIM", "protocole BIM"]},
    "CCTP": {"definition": "Cahier des Clauses Techniques Particulières : document qui décrit les prestations et exigences techniques applicables au lot.", "aliases": ["CCTP du lot"]},
    "DPGF": {"definition": "Décomposition du Prix Global et Forfaitaire : détail chiffré des prestations composant l'offre.", "aliases": ["bordereau de prix", "DPGF numérique"]},
    "BIM Manager": {"definition": "Responsable du processus BIM du projet, de la convention, de la coordination et du contrôle des échanges.", "aliases": ["responsable BIM du projet"]},
    "Coordinateur BIM": {"definition": "Référent chargé de coordonner les modèles et de contrôler la qualité BIM d'une discipline ou d'un groupement.", "aliases": ["coordonnateur BIM"]},
    "IFC": {"definition": "Format ouvert et interopérable d'échange de maquettes numériques, indépendant du logiciel auteur.", "aliases": ["format IFC", "export IFC"]},
    "Format natif": {"definition": "Format propriétaire du logiciel auteur, par exemple RVT pour Revit ou PLN pour Archicad.", "aliases": ["natif", "fichier natif"]},
    "CDE": {"definition": "Environnement commun de données utilisé pour déposer, versionner, valider et partager les documents et maquettes du projet.", "aliases": ["plateforme collaborative", "plateforme commune de dépôt", "environnement commun de données"]},
    "DOE numérique": {"definition": "Dossier des ouvrages exécutés remis sous forme numérique avec les plans, données et informations de recollement.", "aliases": ["DOE BIM", "DOE"]},
    "AIM": {"definition": "Asset Information Model : modèle d'information destiné à l'exploitation et à la maintenance de l'ouvrage.", "aliases": ["Asset Information Model", "modèle d'information de l'actif"]},
    "GMAO": {"definition": "Gestion de maintenance assistée par ordinateur : outil qui exploite les données d'actifs pour organiser la maintenance.", "aliases": ["gestion de maintenance assistée par ordinateur"]},
    "LOD": {"definition": "Niveau de développement géométrique et informationnel attendu pour les objets de la maquette.", "aliases": ["niveau de développement"]},
    "LOIN": {"definition": "Level of Information Need : niveau d'information nécessaire pour un usage et une phase donnés.", "aliases": ["niveau d'information nécessaire", "liste des informations à fournir"]},
    "BCF": {"definition": "Format ouvert permettant d'échanger des remarques, réserves et sujets de coordination liés à une maquette BIM.", "aliases": ["BIM Collaboration Format"]},
    "Revit": {"definition": "Logiciel de conception et de production de maquettes BIM édité par Autodesk.", "aliases": ["RVT"]},
    "ArchiCAD": {"definition": "Logiciel de conception et de production de maquettes BIM édité par Graphisoft.", "aliases": ["Archicad", "PLN"]},
    "Allplan": {"definition": "Logiciel de conception et de production BIM utilisé notamment en architecture et en ingénierie.", "aliases": []},
    "Tekla": {"definition": "Logiciel BIM spécialisé dans la modélisation détaillée des structures et assemblages.", "aliases": ["Tekla Structures"]},
    "DWG": {"definition": "Format de dessin CAO couramment utilisé pour les plans 2D et certains échanges de géométrie.", "aliases": ["format DWG"]},
    "BIM 4D": {"definition": "Association de la maquette au planning afin de représenter le phasage dans le temps.", "aliases": ["4D", "phasage 4D"]},
    "BIM 5D": {"definition": "Association de la maquette aux quantités et aux coûts pour contribuer au chiffrage et au suivi économique.", "aliases": ["5D", "chiffrage 5D"]},
    "Profil du lot": {"definition": "Mode de participation du lot au processus BIM : production de maquette ou contribution principalement documentaire.", "aliases": ["lot documentaire", "producteur de maquette"]},
    "Nomenclature BIM": {"definition": "Tableau structuré listant les objets et leurs propriétés, généré depuis une maquette ou renseigné selon un modèle imposé.", "aliases": ["nomenclature", "tableau de paramètres"]},
    "As-built": {"definition": "État réellement exécuté et relevé en fin de chantier, utilisé pour mettre à jour les plans et données de recollement.", "aliases": ["tel que construit", "recollement"]},
}


def normalize_questions_dataframe(df):
    """Applique les libellés canoniques tout en conservant les identifiants."""
    if df is None or getattr(df, "empty", True):
        return df
    out = df.copy()
    if "id_question" not in out.columns:
        return out
    for idx, row in out.iterrows():
        qid = clean(row.get("id_question"))
        override = QUESTION_OVERRIDES.get(qid)
        if not override:
            continue
        for key, value in override.items():
            if key in out.columns:
                out.at[idx, key] = value
    return out


def _split_aliases(value: Any) -> List[str]:
    text = clean(value)
    if not text:
        return []
    return [clean(item) for item in re.split(r"[;,|]", text) if clean(item)]


def glossary_entries_from_dataframe(df) -> List[Dict[str, Any]]:
    entries: Dict[str, Dict[str, Any]] = {}
    if df is not None and not getattr(df, "empty", True):
        for _, row in df.iterrows():
            term = clean(row.get("terme"))
            definition = clean(row.get("definition") or row.get("definition_excel"))
            if not term or not definition:
                continue
            aliases = []
            aliases.extend(_split_aliases(row.get("mots_cles")))
            aliases.extend(_split_aliases(row.get("synonymes")))
            entries[term.casefold()] = {
                "term": term,
                "definition": definition,
                "source": clean(row.get("source")),
                "category": clean(row.get("categorie")),
                "aliases": sorted({a for a in aliases if a and a.casefold() != term.casefold()}, key=len, reverse=True),
            }
    for term, cfg in FALLBACK_GLOSSARY.items():
        key = term.casefold()
        if key not in entries:
            entries[key] = {
                "term": term,
                "definition": clean(cfg.get("definition")),
                "source": "Glossaire de l'outil",
                "category": "",
                "aliases": list(cfg.get("aliases") or []),
            }
        else:
            existing = entries[key]
            existing["aliases"] = sorted(set(existing.get("aliases", [])) | set(cfg.get("aliases") or []), key=len, reverse=True)
    return sorted(entries.values(), key=lambda item: item["term"].casefold())


def glossary_lookup(entries: Iterable[Mapping[str, Any]]) -> Dict[str, Dict[str, Any]]:
    lookup: Dict[str, Dict[str, Any]] = {}
    for raw in entries or []:
        term = clean(raw.get("term") or raw.get("terme"))
        definition = clean(raw.get("definition") or raw.get("definition_excel"))
        if not term or not definition:
            continue
        entry = {
            "term": term,
            "definition": definition,
            "source": clean(raw.get("source") or raw.get("source_excel")),
            "category": clean(raw.get("category") or raw.get("categorie")),
            "aliases": list(raw.get("aliases") or []),
        }
        for alias in [term, *entry["aliases"]]:
            alias = clean(alias)
            if alias:
                lookup.setdefault(alias.casefold(), entry)
    return lookup
