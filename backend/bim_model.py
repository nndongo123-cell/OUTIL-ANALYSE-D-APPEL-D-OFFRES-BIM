from __future__ import annotations

import re
from typing import Any, Dict, List, Mapping, MutableMapping, Tuple

# ═══════════════════════════════════════════════════════════════════════════
# V10 — SOURCE UNIQUE DE VÉRITÉ : Parametrage_Outil_BIM_V1.xlsx
# ═══════════════════════════════════════════════════════════════════════════
# Tout le contenu métier (titres de bloc, lecture entreprise, questions au
# BIM Manager, propositions commerciales, actions internes) est lu depuis
# la feuille 10_Blocs de l'Excel (colonnes titre_bloc, lecture_entreprise,
# question_bim_manager, proposition_commerciale, action_interne_standard,
# action_interne_capacite) et depuis l'axe
# 'statut_public' de 16_Axes_Config (labels/ordre des statuts CONFIRMED /
# CLARIFY / UNDEMONSTRATED / EXCLUDED).
#
# Les dictionnaires _FALLBACK_* ci-dessous ne sont PLUS la source utilisée
# en priorité : ils ne servent que de filet de sécurité si une ligne de
# l'Excel est incomplète ou si un id_bloc inconnu apparaît, afin que l'outil
# ne produise jamais un texte vide. Modifier l'Excel modifie les livrables ;
# modifier ces dictionnaires n'a d'effet que si l'Excel ne renseigne rien.
_FALLBACK_BLOCK_TITLES = {
    "B01": "Maquette numérique à produire",
    "B02": "Données techniques à renseigner",
    "B03": "Formats de fichiers à remettre",
    "B04": "Coordination et contrôle des maquettes",
    "B05": "Positionnement et cohérence géométrique",
    "B06A": "Dossier numérique de fin de chantier",
    "B06B": "Données pour l'exploitation et la maintenance",
    "B07": "Plateforme commune de dépôt",
    "B08": "Organisation du travail BIM",
    "B09": "Organisation des échanges entre intervenants",
    "B10": "Planning lié à la maquette - BIM 4D",
    "B11": "Qualité des objets et modèles numériques",
    "B12": "BIM 5D - quantités et données de coût",
}

_FALLBACK_BLOCK_DEMANDS = {
    "B01": "Produire, mettre à jour et remettre la maquette numérique du lot dans les formats, le niveau d'information et les jalons définis par le projet.",
    "B02": "Renseigner les propriétés techniques demandées pour les ouvrages du lot dans la matrice, la nomenclature ou le support transmis par le responsable BIM du projet.",
    "B03": "Remettre les fichiers dans les formats contractuels attendus et vérifier leur ouverture, leur conformité et leur version avant transmission.",
    "B04": "Participer aux contrôles de coordination, traiter les conflits qui concernent le lot et suivre la clôture des remarques.",
    "B05": "Utiliser le gabarit, l'origine projet et les règles de positionnement communs afin que la maquette du lot se superpose correctement aux autres modèles.",
    "B06A": "Remettre le dossier numérique de fin de chantier attendu pour le lot : plans de recollement, fiches techniques et, lorsque cela est demandé, données ou maquette mises à jour.",
    "B06B": "Fournir les données d'actifs nécessaires à l'exploitation et à la maintenance : identifiants, références fabricant, garanties, durées de vie et consignes d'entretien, selon le modèle AIM ou GMAO du projet.",
    "B07": "Déposer les documents et maquettes du lot sur la plateforme commune du projet en respectant les droits, statuts, versions et circuits de validation.",
    "B08": "Définir l'organisation BIM du lot, les responsabilités, les contrôles internes et les interfaces avec l'équipe projet.",
    "B09": "Recevoir, intégrer, publier et tracer les maquettes échangées avec les autres intervenants selon la fréquence, les versions et les statuts définis par le projet.",
    "B10": "Associer les objets ou zones de la maquette aux tâches du planning lorsqu'une démarche BIM 4D est contractuellement demandée.",
    "B11": "Appliquer les règles de qualité des objets et des modèles : catégories, nommage, propriétés, géométrie, contrôles et corrections avant publication.",
    "B12": "Produire, dans le cadre d'une démarche BIM 5D, des extractions contrôlées de quantités, métrés et données de coût destinées au chiffrage ou à l'économiste.",
}

_FALLBACK_BLOCK_QUESTIONS = {
    "B01": "Quel niveau de développement, quels formats de remise et quels jalons de mise à jour sont attendus pour la maquette du lot ?",
    "B02": "Quels paramètres le lot doit-il renseigner, sur quel support, à quelle phase et selon quel niveau d'information ?",
    "B03": "Quels formats de fichiers, versions et règles de nommage sont exigés pour les livrables du lot ?",
    "B04": "Quel circuit de détection, d'affectation, de correction et de clôture des conflits s'applique au lot ?",
    "B05": "",
    "B06A": "Quel est le contenu exact du DOE numérique attendu pour le lot : plans de recollement, données structurées, maquette as-built et pièces justificatives ?",
    "B06B": "Le lot doit-il remettre des données d'actifs pour un AIM ou une GMAO ? Précisez les attributs, le format et le jalon de remise.",
    "B07": "Quels types de fichiers le lot doit-il déposer sur le CDE, avec quels statuts, quelles validations et quelle fréquence ?",
    "B08": "Quels éléments d'organisation BIM le lot doit-il formaliser dans le PEB ou les documents d'exécution ?",
    "B09": "Quels modèles le lot doit-il recevoir, consulter ou publier, et quelles règles de version, de statut et de fréquence encadrent les échanges ?",
    "B10": "Une prestation BIM 4D est-elle réellement exigée pour le lot ? Si oui, précisez les objets, le planning, les outils et les jalons concernés.",
    "B11": "Quelle grille de contrôle qualité, quels critères IFC et quels livrables d'autocontrôle s'appliquent au modèle du lot ?",
    "B12": "Une prestation BIM 5D est-elle exigée pour le lot ? Si oui, précisez les quantités, données de coût, formats, destinataires et jalons attendus.",
}

_FALLBACK_COMMERCIAL_PROPOSALS = {
    "B01": "L'entreprise assurera la production, la mise à jour et la remise de la maquette numérique de son lot selon le niveau de développement, les formats et les jalons définis pour le projet.",
    "B02": "L'entreprise assurera le renseignement des propriétés techniques demandées pour les ouvrages de son lot dans la matrice, la nomenclature ou le support transmis par le responsable BIM du projet.",
    "B03": "L'entreprise remettra les livrables de son lot dans les formats, versions et règles de nommage définis par le projet, après contrôle de leur ouverture et de leur intégrité.",
    "B04": "L'entreprise participera au processus de coordination, traitera les observations affectées à son lot et assurera la traçabilité de leur correction jusqu'à clôture.",
    "B05": "L'entreprise positionnera son modèle selon le gabarit, le système de coordonnées et le point de base du projet afin d'en garantir la superposition avec les autres modèles.",
    "B06A": "L'entreprise remettra un DOE numérique structuré pour son lot, comprenant les plans de recollement et les pièces techniques prévues, dans les formats et au jalon définis par le marché.",
    "B06B": "L'entreprise transmettra les données d'actifs de ses ouvrages nécessaires à l'exploitation et à la maintenance, selon le modèle AIM ou GMAO, les attributs et le format définis pour le projet.",
    "B07": "L'entreprise assurera les dépôts de son lot sur le CDE du projet en respectant les droits, la codification, les statuts, les circuits de validation et la fréquence prescrits.",
    "B08": "L'entreprise formalisera l'organisation BIM de son lot, les responsabilités, les contrôles internes et les interfaces avec le responsable BIM du projet.",
    "B09": "L'entreprise appliquera le processus d'échange défini pour le projet : réception et consultation des modèles utiles, publication des versions du lot, respect des statuts et traçabilité des transmissions.",
    "B10": "Lorsque cette prestation est requise par écrit, l'entreprise fournira les informations nécessaires au lien entre les ouvrages de son lot et le planning selon les modalités BIM 4D définies pour le projet.",
    "B11": "L'entreprise appliquera les contrôles qualité définis pour les objets et modèles de son lot et remettra les éléments d'autocontrôle prévus avant publication.",
    "B12": "Lorsque cette prestation est requise par écrit, l'entreprise fournira les extractions contrôlées de quantités et les données de coût prévues par la démarche BIM 5D, selon le format et les jalons définis.",
}

_FALLBACK_STATUS_LABELS = {
    "CONFIRMED": "Confirmée pour le lot",
    "CLARIFY": "À confirmer pour le lot",
    "UNDEMONSTRATED": "Non démontrée dans les documents",
    "EXCLUDED": "Non applicable / explicitement exclue",
}

_FALLBACK_STATUS_ORDER = {"CONFIRMED": 0, "CLARIFY": 1, "UNDEMONSTRATED": 2, "EXCLUDED": 3}


def _status_labels_and_order(axes_config: Mapping[str, Any] | None) -> Tuple[Dict[str, str], Dict[str, int]]:
    """Construit STATUS_LABELS / STATUS_ORDER depuis l'axe 'statut_public' de
    16_Axes_Config si présent, sinon retombe sur le filet de sécurité codé
    en dur. C'est l'Excel qui pilote les libellés affichés par défaut."""
    axis = (axes_config or {}).get("statut_public") if axes_config else None
    if not axis:
        return dict(_FALLBACK_STATUS_LABELS), dict(_FALLBACK_STATUS_ORDER)
    labels, order = {}, {}
    for code, cfg in axis.items():
        if code in _FALLBACK_STATUS_ORDER:
            labels[code] = cfg.get("label") or _FALLBACK_STATUS_LABELS[code]
            order[code] = cfg.get("ordre") if isinstance(cfg.get("ordre"), int) else _FALLBACK_STATUS_ORDER[code]
    for code in _FALLBACK_STATUS_ORDER:
        labels.setdefault(code, _FALLBACK_STATUS_LABELS[code])
        order.setdefault(code, _FALLBACK_STATUS_ORDER[code])
    return labels, order


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def block_sort_key(block_id: Any) -> Tuple[int, int, str]:
    text = clean(block_id).upper()
    match = re.fullmatch(r"B(\d+)([A-Z]*)", text)
    if not match:
        return (9999, 99, text)
    number = int(match.group(1))
    suffix = match.group(2)
    suffix_rank = 0 if not suffix else sum((ord(ch) - 64) * (26 ** i) for i, ch in enumerate(reversed(suffix)))
    return (number, suffix_rank, text)


def _public_status(result: Mapping[str, Any]) -> str:
    raw = clean(result.get("statut")).upper()
    applicability = clean(result.get("applicabilite_lot")).upper()
    presence = clean(result.get("presence_contractuelle")).upper()
    if raw == "NON APPLICABLE" or applicability == "NON_APPLICABLE" or result.get("exclusion_explicit"):
        return "EXCLUDED"
    if raw == "CONFIRMÉE" and presence != "ABSENTE":
        return "CONFIRMED"
    if raw in {"PROBABLE", "PARTIELLE"} or (presence != "ABSENTE" and applicability in {"PROBABLE", "A_CONFIRMER"}):
        return "CLARIFY"
    return "UNDEMONSTRATED"


def _capacity(block_id: str, result: Mapping[str, Any], scores: Mapping[str, Mapping[str, Any]]) -> Dict[str, Any]:
    score = scores.get(block_id)
    if score is not None:
        responses = list(score.get("reponses") or [])
        pct_raw = score.get("pct")
        pct = int(round(float(pct_raw))) if isinstance(pct_raw, (int, float)) else None
        if responses or pct is not None:
            pct = max(0, min(100, pct if pct is not None else 0))
            if pct >= 70:
                state, label = "DEMONSTRATED", f"Capacité en place - {pct} %"
            elif pct >= 40:
                state, label = "PARTIAL", f"Capacité partielle - {pct} %"
            else:
                state, label = "NOT_DEMONSTRATED", f"Capacité à développer - {pct} %"
            strengths, gaps = [], []
            for response in responses:
                val = response.get("reponse_val")
                label_txt = clean(response.get("reponse_label") or response.get("question"))
                if not label_txt:
                    continue
                if val == 2:
                    strengths.append(label_txt)
                elif val in (0, 1):
                    gaps.append(label_txt)
            return {
                "state": state,
                "label": label,
                "pct": pct,
                "strength": strengths[0] if strengths else "Aucun acquis suffisamment démontré par les réponses fournies.",
                "gap": gaps[0] if gaps else "Aucun écart majeur identifié dans les réponses fournies.",
                "responses": responses,
            }
    raw_capacity = clean(result.get("capacite_entreprise")).upper()
    if raw_capacity and raw_capacity not in {"NON_EVALUEE", "NON ÉVALUÉE"}:
        return {
            "state": raw_capacity,
            "label": raw_capacity.replace("_", " ").title(),
            "pct": None,
            "strength": "Capacité déclarée sans score détaillé.",
            "gap": "",
            "responses": [],
        }
    return {
        "state": "UNAVAILABLE",
        "label": "Non couverte par l'auto-évaluation",
        "pct": None,
        "strength": "Ce bloc ne possède pas de question active dans le questionnaire actuel.",
        "gap": "",
        "responses": [],
    }


def clean_generated_text(value: Any) -> str:
    text = clean(value)
    replacements = {
        "nous confirmeons": "nous confirmons",
        "Nous confirmeons": "Nous confirmons",
        "la plateforme plateforme commune de dépôt (CDE)": "la plateforme commune de dépôt (CDE)",
        "plateforme plateforme commune de dépôt (CDE)": "plateforme commune de dépôt (CDE)",
        "plateforme commune de dépôt (CDE) (Environnement Commun de Données)": "plateforme commune de dépôt (CDE)",
        "du projet du projet": "du projet",
        "outil de outil de gestion de maintenance (GMAO)": "outil de gestion de maintenance (GMAO)",
        "modèle de données outil de gestion de maintenance (GMAO)": "modèle de données GMAO",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return clean(re.sub(r"\bresponsable BIM du projet \(BIM Manager\) du projet\b", "responsable BIM du projet (BIM Manager)", text, flags=re.I))


def _evidence(result: Mapping[str, Any], lot: str = "") -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for source, text_key, page_key in (("Convention BIM", "conv_pf", "conv_pg"), ("CCTP du lot", "cctp_pf", "cctp_pg")):
        text = clean(result.get(text_key))
        if not text:
            continue
        rows.append({
            "source": source,
            "page": clean(result.get(page_key)) or "Non précisée",
            "lot": clean(lot) or "Lot analysé",
            "qualification": "Extrait ciblé retenu",
            "text": text,
        })
    return rows


def _rejected_evidence(result: Mapping[str, Any]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for item in result.get("preuves_rejetees") or []:
        if isinstance(item, Mapping):
            rows.append({"text": clean(item.get("texte")), "reason": clean(item.get("raison"))})
        else:
            rows.append({"text": clean(item), "reason": "Extrait non retenu."})
    return rows


def _proposal(block_id: str, status_key: str, capacity: Mapping[str, Any], result: Mapping[str, Any], proposals: Mapping[str, Mapping[str, Any]]) -> Tuple[str, str]:
    proposal = proposals.get(block_id) or {}
    source_text = clean_generated_text(proposal.get("texte"))
    bloc_row = result.get("bloc") or {}
    excel_commercial = clean(bloc_row.get("proposition_commerciale"))
    commercial = (
        excel_commercial
        or source_text
        or _FALLBACK_COMMERCIAL_PROPOSALS.get(block_id)
        or "L'entreprise assurera cette prestation conformément aux exigences applicables à son lot."
    )
    contradictions = result.get("contradictions") or []

    title = clean(bloc_row.get("titre_bloc")) or _FALLBACK_BLOCK_TITLES.get(block_id, block_id)
    if status_key == "EXCLUDED":
        return "Aucun engagement à intégrer", f"« {title} » n'est pas incluse dans les exigences applicables au lot, sauf demande écrite modifiant le périmètre contractuel."
    if status_key == "UNDEMONSTRATED":
        return "Proposition à intégrer sous réserve de demande écrite", (
            f"« {title} » n'est pas identifiée comme exigence applicable au lot dans les documents analysés. "
            "Elle pourra être intégrée sur demande écrite après définition du périmètre, des livrables, des jalons et de l'incidence éventuelle."
        )
    if status_key == "CLARIFY":
        return "Proposition à intégrer - périmètre à confirmer", (
            f"L'entreprise réalisera « {title} » selon les modalités qui seront précisées lors de la "
            "mise au point du marché (périmètre exact, format et jalon applicables au lot)."
        )
    if contradictions:
        return "Proposition à valider en interne avant envoi", commercial

    # La réponse externe reste affirmative et orientée résultat. Les écarts de capacité
    # sont traités séparément dans le plan d'actions interne, sans dévaloriser l'offre.
    return "Proposition à intégrer dans l'offre", commercial


def _normalize_question(value: Any) -> str:
    question = clean_generated_text(value)
    if not question:
        return ""
    if "?" not in question:
        question = question.rstrip(" .") + ". Pouvez-vous confirmer ce point et les modalités applicables au lot ?"
    return question


def _origin_from_evidence(evidence: List[Mapping[str, Any]]) -> str:
    if not evidence:
        return "Analyse documentaire - aucune preuve ciblée retenue"
    first = evidence[0]
    page = clean(first.get("page"))
    return f"{clean(first.get('source'))}{f' p.{page}' if page else ''}"


def _gap_response(capacity: Mapping[str, Any]) -> Mapping[str, Any] | None:
    for response in capacity.get("responses") or []:
        if response.get("reponse_val") in (0, 1):
            return response
    return None


_GENERIC_ACTION_STANDARD = "Définir le livrable, le responsable, le contrôle et le jalon associés à cette exigence."
_GENERIC_ACTION_CAPACITE = "Définir un contrôle concret et conserver la preuve de sa réalisation avant engagement."


def _action_template(block_id: str, bloc_row: Mapping[str, Any] | None = None) -> str:
    """Action interne de sécurisation (plan d'actions) — lue en priorité depuis
    la colonne action_interne_standard de 10_Blocs ; filet de sécurité codé en
    dur uniquement si l'Excel est vide."""
    bloc_row = bloc_row or {}
    return clean(bloc_row.get("action_interne_standard")) or _FALLBACK_ACTION_TEMPLATES.get(block_id) or _GENERIC_ACTION_STANDARD


def _capacity_action(block_id: str, response_label: str, bloc_row: Mapping[str, Any] | None = None) -> str:
    """Action interne de test de capacité — lue en priorité depuis la colonne
    action_interne_capacite de 10_Blocs ; filet de sécurité codé en dur
    uniquement si l'Excel est vide."""
    bloc_row = bloc_row or {}
    return clean(bloc_row.get("action_interne_capacite")) or _FALLBACK_CAPACITY_TEMPLATES.get(block_id) or _GENERIC_ACTION_CAPACITE


# Filet de sécurité (id_bloc inconnu de l'Excel, ou colonne vide) — normalement
# jamais utilisé une fois 10_Blocs correctement renseigné.
_FALLBACK_ACTION_TEMPLATES = {
    "B01": "Désigner un responsable de la maquette du lot ; le format, le niveau de détail et les jalons seront confirmés lors de la mise au point du marché.",
    "B02": "Prévoir le renseignement des données demandées (matrice ou nomenclature) ; le support exact sera transmis lors de la mise au point du marché.",
    "B03": "Vérifier en interne que chaque format de fichier demandé peut être produit et ouvert correctement avant remise.",
    "B04": "Désigner un responsable interne du traitement des observations et des délais de correction pour le lot.",
    "B05": "Prévoir le positionnement du modèle selon le gabarit et le point de référence du projet, qui seront transmis lors de la mise au point du marché.",
    "B06A": "Lister en interne les documents du DOE numérique probables pour le lot et désigner un responsable de leur remise.",
    "B06B": "Prévoir la fourniture de données d'actifs structurées (AIM ou GMAO) ; le modèle exact sera transmis lors de la mise au point du marché.",
    "B07": "Prévoir l'utilisation de la plateforme commune du projet ; les accès et modalités de dépôt seront transmis après attribution du marché.",
    "B08": "Désigner un référent BIM interne pour le lot ; son rôle sera précisé avec le responsable BIM du projet lors de la mise au point du marché.",
    "B09": "Prévoir un processus d'échange des maquettes (versions, statuts, fréquence), à finaliser avec le responsable BIM du projet lors de la mise au point du marché.",
    "B10": "Ne pas engager de travail de liaison au planning (BIM 4D) tant que le périmètre n'est pas confirmé lors de la mise au point du marché.",
    "B11": "Prévoir l'application d'une grille de contrôle qualité au modèle du lot ; son détail sera transmis lors de la mise au point du marché.",
    "B12": "Ne pas engager de prestation BIM 5D tant que les quantités et données de coût demandées ne sont pas confirmées lors de la mise au point du marché.",
}

_FALLBACK_CAPACITY_TEMPLATES = {
    "B01": "Vérifier en interne que la maquette peut être exportée dans un format IFC et natif.",
    "B02": "Vérifier en interne que les références, performances et données fabricant demandées sont disponibles.",
    "B03": "Vérifier en interne la capacité à produire les formats de fichiers demandés.",
    "B04": "Vérifier en interne la capacité à traiter et corriger des observations dans des délais courts.",
    "B05": "Vérifier en interne la capacité à positionner un modèle selon un gabarit et un point de référence de projet.",
    "B06A": "Vérifier en interne la capacité à produire les documents probables du DOE numérique.",
    "B06B": "Vérifier en interne que les données d'actifs (identifiants, garanties, entretien) sont disponibles.",
    "B07": "Vérifier en interne l'expérience de dépôt sur une plateforme commune de projet.",
    "B08": "Vérifier en interne la disponibilité d'un référent BIM pour organiser le travail du lot.",
    "B09": "Vérifier en interne la capacité à publier et recevoir des maquettes selon un processus structuré.",
    "B10": "Vérifier en interne la capacité à lier des objets du modèle à un planning, si le périmètre est confirmé.",
    "B11": "Vérifier en interne la capacité à contrôler la qualité des objets et modèles du lot.",
    "B12": "Vérifier en interne la capacité à extraire des quantités et données de coût depuis une maquette.",
}


def _action_items(block_id: str, status_key: str, capacity: Mapping[str, Any], result: Mapping[str, Any], evidence: List[Mapping[str, Any]]) -> List[Dict[str, str]]:
    bloc_row = result.get("bloc") or {}
    if status_key == "EXCLUDED":
        # "exclusion_explicit" n'est vrai que pour une clause textuelle réellement
        # détectée (ex. "exclusions de principe" pour le 4D/5D). Dans tous les
        # autres cas (mot-clé d'exclusion CCTP,
        # nuance mappée en non-applicable), l'origine réelle est différente et
        # ne doit pas prétendre qu'une clause d'exclusion explicite a été lue.
        if result.get("exclusion_explicit"):
            return [{
                "text": "Vérifier, avant remise de l'offre, qu'aucun ordre écrit ou additif ne réintroduit cette prestation.",
                "origin": "Clause d'exclusion explicite détectée dans les documents",
                "kind": "document",
            }]
        if evidence:
            return [{
                "text": "Vérifier, avant remise de l'offre, que le passage ci-dessous exclut bien cette prestation pour le lot, et qu'aucun ordre écrit ou additif ne la réintroduit.",
                "origin": _origin_from_evidence(evidence),
                "kind": "document",
            }]
        return [{
            "text": "Aucune clause d'exclusion explicite n'a été identifiée dans les documents ; conserver une hypothèse chiffrée avec réserve dans l'offre plutôt que d'exclure définitivement cette prestation, sauf confirmation obtenue via le circuit officiel de questions du DCE.",
            "origin": "Absence de preuve contractuelle d'applicabilité - à confirmer",
            "kind": "clarification",
        }]

    items: List[Dict[str, str]] = []
    evidence_origin = _origin_from_evidence(evidence)
    title = clean(bloc_row.get("titre_bloc")) or _FALLBACK_BLOCK_TITLES.get(block_id, block_id)
    if status_key == "UNDEMONSTRATED":
        texte = clean(bloc_row.get("action_interne_non_demontree")) or (
            f"Si le règlement de consultation le permet, poser la question via le circuit officiel de questions du DCE pour savoir si « {title} » s'applique au lot ; sinon, formuler une hypothèse chiffrée avec réserve explicite dans l'offre, à lever lors de la mise au point du marché."
        )
        items.append({"text": texte, "origin": evidence_origin, "kind": "clarification"})
    elif status_key == "CLARIFY":
        texte = clean(bloc_row.get("action_interne_a_confirmer")) or (
            f"Si le règlement de consultation le permet, faire préciser par écrit via le circuit officiel de questions du DCE le périmètre, le format attendu et le jalon de remise pour « {title} » ; sinon, formuler une hypothèse chiffrée avec réserve explicite dans l'offre, à ajuster lors de la mise au point du marché."
        )
        items.append({"text": texte, "origin": evidence_origin, "kind": "clarification"})
    elif capacity.get("state") != "DEMONSTRATED":
        # Si la capacité est déjà pleinement démontrée (>= 70%) pour une
        # exigence confirmée, il n'y a rien de plus à préparer : générer une
        # action ici ferait croire à un travail restant alors que le bloc est
        # déjà conforme.
        items.append({"text": _action_template(block_id, bloc_row), "origin": evidence_origin, "kind": "document"})

    # Ne pas générer d'action de rattrapage de capacité si celle-ci est déjà
    # démontrée dans l'ensemble (>= 70%, même seuil que le libellé "Démontrée -
    # X%") : une réponse isolée à 0/1 parmi plusieurs ne justifie pas une ligne
    # "à faire" quand le bloc est globalement déjà conforme.
    gap = _gap_response(capacity) if capacity.get("state") != "DEMONSTRATED" else None
    if gap:
        qid = clean(gap.get("id_question"))
        label = clean(gap.get("reponse_label"))
        items.append({
            "text": _capacity_action(block_id, label, bloc_row),
            "origin": f"Auto-évaluation {qid} - réponse : {label}" if qid else f"Auto-évaluation - réponse : {label}",
            "kind": "capacity",
        })

    # Dédoublonnage strict par texte.
    unique, seen = [], set()
    for item in items:
        key = clean(item.get("text")).casefold()
        if key and key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def _priority(status_key: str, capacity: Mapping[str, Any], result: Mapping[str, Any]) -> str:
    if status_key == "EXCLUDED":
        return "Aucune action sauf ordre écrit"
    if result.get("contradictions"):
        return "Critique - lever la contradiction"
    if status_key == "CONFIRMED":
        if capacity.get("state") in {"NOT_DEMONSTRATED", "UNAVAILABLE"}:
            return "Critique - sécuriser avant engagement"
        if capacity.get("state") == "PARTIAL":
            return "Haute - à consolider en interne"
        return "À intégrer dans l'offre"
    if status_key == "CLARIFY":
        return "Avant dépôt - réserve ou question DCE à formuler"
    return "À surveiller - preuve ou périmètre à établir"


def _decision(status_key: str, capacity: Mapping[str, Any], result: Mapping[str, Any]) -> str:
    if status_key == "EXCLUDED":
        return "Ne pas chiffrer sauf ordre écrit"
    if status_key == "UNDEMONSTRATED":
        return "Réserve à formuler avant chiffrage"
    if status_key == "CLARIFY":
        return "Clarifier avant engagement"
    if result.get("contradictions"):
        return "Valider l'organisation interne avant envoi"
    if capacity.get("state") in {"NOT_DEMONSTRATED", "PARTIAL", "UNAVAILABLE"}:
        return "Intégrer à l'offre et sécuriser l'exécution"
    return "Intégrer à l'offre"


def ensure_public_model(
    results: MutableMapping[str, MutableMapping[str, Any]],
    scores: Mapping[str, Mapping[str, Any]] | None = None,
    proposals: Mapping[str, Mapping[str, Any]] | None = None,
    lot: str = "",
    axes_config: Mapping[str, Any] | None = None,
) -> MutableMapping[str, MutableMapping[str, Any]]:
    scores = scores or {}
    proposals = proposals or {}
    status_labels, status_order = _status_labels_and_order(axes_config)
    for block_id, result in results.items():
        bloc_row = result.get("bloc") or {}
        status_key = _public_status(result)
        capacity = _capacity(block_id, result, scores)
        response_label, response_text = _proposal(block_id, status_key, capacity, result, proposals)
        title = clean(bloc_row.get("titre_bloc")) or _FALLBACK_BLOCK_TITLES.get(block_id) or block_id
        evidence = _evidence(result, lot=lot)
        action_items = _action_items(block_id, status_key, capacity, result, evidence)
        result.update({
            "public_id": block_id,
            "public_title": title,
            "public_demand": clean(bloc_row.get("lecture_entreprise")) or _FALLBACK_BLOCK_DEMANDS.get(block_id) or clean(result.get("conclusion")),
            "public_status_key": status_key,
            "public_status_label": status_labels[status_key],
            "public_group_order": status_order[status_key],
            "public_capacity_state": capacity["state"],
            "public_capacity_label": capacity["label"],
            "public_capacity_pct": capacity["pct"],
            "public_capacity_strength": capacity["strength"],
            "public_capacity_gap": capacity["gap"],
            "public_priority": _priority(status_key, capacity, result),
            "public_decision": _decision(status_key, capacity, result),
            "public_response_label": response_label,
            "public_response_text": response_text,
            "public_action_items": action_items,
            "public_actions": [item["text"] for item in action_items],
            "public_question": (
                _normalize_question(clean(bloc_row.get("question_bim_manager")) or _FALLBACK_BLOCK_QUESTIONS.get(block_id))
                if status_key != "CONFIRMED" else ""
            ),
            "public_evidence": evidence,
            "public_rejected_evidence": _rejected_evidence(result),
            "public_sort_key": block_sort_key(block_id),
        })
    return results


def sorted_items(results: Mapping[str, Mapping[str, Any]]) -> List[Tuple[str, Mapping[str, Any]]]:
    return sorted(results.items(), key=lambda item: (item[1].get("public_group_order", 99), block_sort_key(item[0])))
