from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Sequence, Tuple


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def block_sort_key(block_id: Any) -> Tuple[int, int, str]:
    """Tri naturel B01, B02, B06A, B06B, B10… (mécanique, sans règle métier)."""
    text = clean(block_id).upper()
    match = re.fullmatch(r"B(\d+)([A-Z]*)", text)
    if not match:
        return (9999, 99, text)
    number = int(match.group(1))
    suffix = match.group(2)
    suffix_rank = 0 if not suffix else sum(
        (ord(ch) - 64) * (26 ** i) for i, ch in enumerate(reversed(suffix))
    )
    return (number, suffix_rank, text)


def clean_generated_text(value: Any) -> str:
    """Nettoyage purement typographique des substitutions automatiques."""
    text = clean(value)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"([,.;:!?])([^\s])", r"\1 \2", text)
    return clean(text)


def _axis(axes_config: Mapping[str, Any] | None, name: str) -> Mapping[str, Any]:
    return (axes_config or {}).get(name) or {}


def _axis_item(axes_config: Mapping[str, Any] | None, axis_name: str, code: str) -> Mapping[str, Any]:
    return _axis(axes_config, axis_name).get(code) or {}


def _axis_label(axes_config: Mapping[str, Any] | None, axis_name: str, code: str) -> str:
    cfg = _axis_item(axes_config, axis_name, code)
    return clean(cfg.get("label")) or clean(code)


def _axis_order(axes_config: Mapping[str, Any] | None, axis_name: str, code: str) -> int:
    cfg = _axis_item(axes_config, axis_name, code)
    try:
        return int(cfg.get("ordre"))
    except (TypeError, ValueError):
        return 9999


def _axis_description(axes_config: Mapping[str, Any] | None, axis_name: str, code: str) -> str:
    return clean(_axis_item(axes_config, axis_name, code).get("description"))


def _axis_color(axes_config: Mapping[str, Any] | None, axis_name: str, code: str) -> str:
    return clean(_axis_item(axes_config, axis_name, code).get("couleur"))


def _axis_field(axes_config: Mapping[str, Any] | None, axis_name: str, code: str, field: str) -> str:
    return clean(_axis_item(axes_config, axis_name, code).get(field))


def capacity_code_for_pct(axes_config: Mapping[str, Any] | None, pct: int | float | None) -> str:
    """Classe un score uniquement à partir de pct_min/pct_max de 16_Axes_Config."""
    if pct is None:
        # La catégorie sans bornes représente le cas non évalué.
        candidates = [
            (code, cfg) for code, cfg in _axis(axes_config, "capacite").items()
            if cfg.get("pct_min") in (None, "") and cfg.get("pct_max") in (None, "")
        ]
        candidates.sort(key=lambda item: _axis_order(axes_config, "capacite", item[0]))
        return candidates[0][0] if candidates else ""
    try:
        value = int(round(float(pct)))
    except (TypeError, ValueError):
        return ""
    matches: List[Tuple[int, str]] = []
    for code, cfg in _axis(axes_config, "capacite").items():
        lo, hi = cfg.get("pct_min"), cfg.get("pct_max")
        if lo in (None, "") or hi in (None, ""):
            continue
        try:
            if float(lo) <= value <= float(hi):
                matches.append((_axis_order(axes_config, "capacite", code), code))
        except (TypeError, ValueError):
            continue
    matches.sort()
    return matches[0][1] if matches else ""


def _best_capacity_code(axes_config: Mapping[str, Any] | None) -> str:
    candidates = []
    for code, cfg in _axis(axes_config, "capacite").items():
        if cfg.get("pct_min") in (None, "") or cfg.get("pct_max") in (None, ""):
            continue
        candidates.append((_axis_order(axes_config, "capacite", code), code))
    candidates.sort()
    return candidates[0][1] if candidates else ""


def _message(messages: Mapping[str, str] | None, key: str, **ctx: Any) -> str:
    tpl = clean((messages or {}).get(key, ""))
    if not tpl:
        return ""
    class Safe(dict):
        def __missing__(self, name):
            return ""
    try:
        return clean(tpl.format_map(Safe({k: clean(v) for k, v in ctx.items()})))
    except Exception:
        return tpl


def _split_expected(value: Any) -> List[str]:
    return [clean(v) for v in str(value or "").split(";") if clean(v)]


def _norm_bool(value: Any) -> str:
    if isinstance(value, bool):
        return "Oui" if value else "Non"
    text = clean(value)
    if text.casefold() in {"true", "yes", "oui", "1"}:
        return "Oui"
    if text.casefold() in {"false", "no", "non", "0", ""}:
        return "Non"
    return text


def _condition_matches(actual: Any, expected: Any) -> bool:
    exprs = _split_expected(expected)
    if not exprs:
        return True
    actual_text = _norm_bool(actual)
    positives = [e for e in exprs if not e.startswith("!")]
    negatives = [e[1:] for e in exprs if e.startswith("!")]
    if negatives and any(actual_text.casefold() == n.casefold() for n in negatives):
        return False
    if positives:
        return any(actual_text.casefold() == p.casefold() for p in positives)
    return True


def choose_rule(rules: Sequence[Mapping[str, Any]] | None, rule_type: str, context: Mapping[str, Any]) -> str:
    """Applique 17_Regles_Restitution par ordre croissant, sans règle de repli Python."""
    candidates = [r for r in (rules or []) if clean(r.get("type_sortie")) == rule_type]
    def order(row: Mapping[str, Any]) -> int:
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999
    candidates.sort(key=order)
    fields = (
        "statut", "presence", "applicabilite", "capacite", "statut_public",
        "contradiction", "exclusion_explicit", "preuve", "capacite_best",
        "ecart_capacite", "interblock_exclusion",
    )
    for row in candidates:
        if all(_condition_matches(context.get(field, ""), row.get(field, "")) for field in fields):
            return clean(row.get("code_sortie"))
    return ""


def _capacity(
    block_id: str,
    result: Mapping[str, Any],
    scores: Mapping[str, Mapping[str, Any]],
    axes_config: Mapping[str, Any] | None,
    messages: Mapping[str, str] | None,
) -> Dict[str, Any]:
    """Restitue la capacité propre de l'entreprise, indépendamment du DCE.

    Le score public d'un bloc provient toujours de l'auto-évaluation globale de
    ce bloc. Les identifiants ``capacity_question_ids`` fournis par les règles
    Excel ne servent plus à recalculer ce score : ils sélectionnent seulement
    les réponses pertinentes pour comparer la capacité déclarée à l'exigence
    précise du marché (écart ciblé / action). Le moteur reste générique et ne
    connaît ni les blocs ni le sens métier des questions.
    """
    score = scores.get(block_id)
    if score is not None:
        responses = list(score.get("reponses") or [])
        pct_raw = score.get("pct")
        pct = int(round(float(pct_raw))) if isinstance(pct_raw, (int, float)) else None
        code = clean(score.get("capacite_code")) or capacity_code_for_pct(axes_config, pct)

        # Les règles Excel peuvent désigner les questions utiles pour vérifier
        # l'adéquation à une exigence précise. Elles peuvent réutiliser une
        # question évaluée dans un autre bloc ; cette sélection ne modifie jamais
        # le score global du bloc affiché à l'utilisateur.
        selected_qid_list = []
        seen_selected_qids = set()
        for raw_qid in (result.get("capacity_question_ids") or []):
            qid = clean(raw_qid).casefold()
            if qid and qid not in seen_selected_qids:
                selected_qid_list.append(qid)
                seen_selected_qids.add(qid)

        requirement_responses: List[Mapping[str, Any]] = []
        missing_requirement_qids: List[str] = []
        if selected_qid_list:
            for wanted_qid in selected_qid_list:
                found = None
                for candidate_score in scores.values():
                    found = next((
                        r for r in (candidate_score.get("reponses") or [])
                        if clean(r.get("id_question")).casefold() == wanted_qid
                    ), None)
                    if found is not None:
                        break
                if found is not None:
                    requirement_responses.append(found)
                else:
                    missing_requirement_qids.append(wanted_qid)

        # Cycle 26 : l'auto-évaluation globale reste la photographie intrinsèque
        # de l'entreprise. En parallèle, lorsqu'Excel désigne les questions
        # nécessaires à une exigence précise, le moteur calcule une couverture
        # ciblée sur ces seules questions avec leurs poids d'origine. Ce second
        # indicateur ne remplace jamais le score global du bloc.
        requirement_coverage_pct = None
        requirement_coverage_evaluated = False
        if selected_qid_list and not missing_requirement_qids and requirement_responses:
            requirement_score = 0
            requirement_max = 0
            for response in requirement_responses:
                try:
                    weight = int(float(response.get("poids", 1) or 1))
                except (TypeError, ValueError):
                    weight = 1
                try:
                    value = int(response.get("reponse_val"))
                except (TypeError, ValueError):
                    continue
                if value not in (0, 1, 2):
                    continue
                requirement_score += value * weight
                requirement_max += 2 * weight
            if requirement_max:
                requirement_coverage_pct = round(requirement_score / requirement_max * 100)
                requirement_coverage_evaluated = True

        requirement_coverage_label = ""
        if selected_qid_list:
            if requirement_coverage_evaluated:
                requirement_coverage_label = f"{requirement_coverage_pct} %"
            else:
                requirement_coverage_label = _message(messages, "capacity_requirement_not_evaluated") or "Non évaluée"

        # Le point à renforcer affiché dans le plan est ciblé sur les questions
        # réellement mobilisées lorsqu'Excel en a désigné. Si une question
        # attendue manque, on ne retombe surtout pas sur les lacunes globales du
        # bloc : la couverture spécifique est alors simplement non évaluée.
        gap_responses = requirement_responses if selected_qid_list else responses

        if code:
            strengths, gaps = [], []
            for response in responses:
                val = response.get("reponse_val")
                label_txt = clean(response.get("reponse_label") or response.get("question"))
                if not label_txt:
                    continue
                if val == 2:
                    strengths.append(label_txt)
            for response in gap_responses:
                val = response.get("reponse_val")
                label_txt = clean(response.get("reponse_label") or response.get("question"))
                if label_txt and val in (0, 1):
                    gaps.append(label_txt)

            label = _axis_label(axes_config, "capacite", code)
            if pct is not None:
                label = f"{label} - {pct} %"
            return {
                "code": code,
                "label": label,
                "pct": pct,
                "strength": strengths[0] if strengths else _message(messages, "capacity_strength_none"),
                "gap": gaps[0] if gaps else _message(messages, "capacity_gap_none"),
                "responses": responses,
                "requirement_responses": requirement_responses,
                "requirement_question_ids": selected_qid_list,
                "requirement_missing_question_ids": missing_requirement_qids,
                "requirement_coverage_pct": requirement_coverage_pct,
                "requirement_coverage_evaluated": requirement_coverage_evaluated,
                "requirement_coverage_label": requirement_coverage_label,
                "gap_responses": gap_responses,
            }
    code = capacity_code_for_pct(axes_config, None)
    return {
        "code": code,
        "label": _axis_label(axes_config, "capacite", code) or _message(messages, "capacity_unavailable_label"),
        "pct": None,
        "strength": _message(messages, "capacity_unavailable_strength"),
        "gap": "",
        "responses": [],
        "requirement_responses": [],
        "requirement_question_ids": [],
        "requirement_missing_question_ids": [],
        "requirement_coverage_pct": None,
        "requirement_coverage_evaluated": False,
        "requirement_coverage_label": "",
        "gap_responses": [],
    }

def _evidence(result: Mapping[str, Any], lot: str, messages: Mapping[str, str] | None) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    specs = [
        ("conv", "source_convention_label", "conv_pf", "conv_pg"),
        ("cctp", "source_cctp_label", "cctp_pf", "cctp_pg"),
    ]
    # Une règle documentaire peut désigner explicitement la source qui doit être
    # affichée comme preuve principale. Le nom du champ et sa valeur sont fournis
    # par Excel via le mécanisme générique champ_resultat / valeur_resultat ;
    # aucune connaissance d'un bloc particulier n'est codée ici.
    primary_source = clean(result.get("primary_evidence_source")).upper()
    if primary_source in {"MATCHED", "SPECIFIC"}:
        primary_source = clean(result.get("specific_source")).upper()
    if primary_source == "CCTP_AND_SPECIFIC":
        specs = [spec for spec in specs if spec[0] == "cctp"]
    elif primary_source == "CONVENTION_AND_SPECIFIC":
        specs = [spec for spec in specs if spec[0] == "conv"]
    if primary_source in {"CCTP", "CONVENTION"}:
        primary_kind = "cctp" if primary_source == "CCTP" else "conv"
        # Lorsqu'une règle documentaire Excel a isolé une clause précise et
        # désigné sa source principale, cette clause remplace les preuves
        # génériques préexistantes de l'autre document. Les preuves
        # complémentaires explicitement ajoutées par une règle ``AJOUTE``
        # restent ensuite conservées via ``specific_evidence_extra``.
        # Le moteur ne connaît ni le bloc ni le document attendu : la source
        # principale est entièrement déterminée par le paramétrage Excel.
        specs = [spec for spec in specs if spec[0] == primary_kind]
    for kind, source_key, text_key, page_key in specs:
        text = clean(result.get(text_key))
        if not text:
            continue
        qualification = _message(messages, "evidence_qualification")
        if result.get("interblock_exclusion"):
            qualification = _message(messages, "evidence_qualification_project_context") or qualification
        rows.append({
            "kind": kind,
            "source": _message(messages, source_key),
            "page": clean(result.get(page_key)) or _message(messages, "evidence_page_unknown"),
            "lot": clean(lot) or _message(messages, "evidence_lot_default"),
            "qualification": qualification,
            "text": text,
        })
    # Lorsqu'une pièce complémentaire a été ajoutée depuis une vigilance, conserver
    # aussi la clause d'origine qui justifie pourquoi cette pièce est pertinente.
    # On obtient ainsi une chaîne de preuve auditable : document principal -> renvoi
    # -> pièce complémentaire. Ce lien provient de l'état d'analyse, pas du frontend.
    supp_origin_excerpt = clean(result.get("supp_origin_excerpt"))
    supp_origin_page = clean(result.get("supp_origin_page"))
    supp_origin_source = clean(result.get("supp_origin_source")).upper()
    if supp_origin_excerpt and supp_origin_page and supp_origin_source in {"CCTP", "CONVENTION"}:
        origin_kind = "cctp" if supp_origin_source == "CCTP" else "conv"
        origin_source = (
            _message(messages, "source_cctp_label") if origin_kind == "cctp"
            else _message(messages, "source_convention_label")
        )
        key = (origin_kind.casefold(), supp_origin_page.casefold(), supp_origin_excerpt.casefold())
        existing = {(clean(r.get("kind")).casefold(), clean(r.get("page")).casefold(), clean(r.get("text")).casefold()) for r in rows}
        if key not in existing:
            rows.append({
                "kind": origin_kind,
                "source": origin_source,
                "page": supp_origin_page,
                "lot": clean(lot) or _message(messages, "evidence_lot_default"),
                "qualification": _message(messages, "evidence_qualification"),
                "text": supp_origin_excerpt,
            })

    supp_text = clean(result.get("supp_pf"))
    if supp_text:
        supp_kind = clean(result.get("supp_kind")) or "supp0"
        supp_source = clean(result.get("supp_source")) or _message(messages, "source_supplement_label") or "Pièce complémentaire"
        supp_page = clean(result.get("supp_pg")) or _message(messages, "evidence_page_unknown")
        key = (supp_kind.casefold(), supp_page.casefold(), supp_text.casefold())
        existing = {(clean(r.get("kind")).casefold(), clean(r.get("page")).casefold(), clean(r.get("text")).casefold()) for r in rows}
        if key not in existing:
            rows.append({
                "kind": supp_kind,
                "source": supp_source,
                "page": supp_page,
                "lot": clean(lot) or _message(messages, "evidence_lot_default"),
                "qualification": _message(messages, "evidence_qualification"),
                "text": supp_text,
            })
    # Des règles documentaires paramétrées peuvent agréger plusieurs clauses
    # distinctes d'un même document (par exemple plusieurs formats/livrables).
    # Elles sont ajoutées comme preuves séparées pour préserver page et traçabilité.
    for extra in result.get("specific_evidence_extra") or []:
        if not isinstance(extra, Mapping):
            continue
        text = clean(extra.get("text"))
        if not text:
            continue
        kind = clean(extra.get("kind")) or "cctp"
        source = clean(extra.get("source")) or (
            _message(messages, "source_cctp_label") if kind == "cctp" else _message(messages, "source_convention_label")
        )
        page = clean(extra.get("page")) or _message(messages, "evidence_page_unknown")
        key = (kind.casefold(), page.casefold(), text.casefold())
        existing = {(clean(r.get("kind")).casefold(), clean(r.get("page")).casefold(), clean(r.get("text")).casefold()) for r in rows}
        if key in existing:
            continue
        rows.append({
            "kind": kind,
            "source": source,
            "page": page,
            "lot": clean(lot) or _message(messages, "evidence_lot_default"),
            "qualification": _message(messages, "evidence_qualification"),
            "text": text,
        })
    # Une exclusion issue d'une règle interbloc doit rester audit-able même si
    # le bloc cible possède par ailleurs une clause de contexte projet. On ajoute
    # donc la preuve du bloc source ayant déclenché la cascade, sans remplacer
    # les éventuelles clauses propres au bloc cible. Le mécanisme est transversal :
    # aucune identité de bloc, de lot ou de projet n'est codée ici.
    if result.get("interblock_exclusion"):
        src_ev = result.get("interblock_source_evidence") or {}
        existing = {
            (clean(r.get("kind")).casefold(), clean(r.get("page")).casefold(), clean(r.get("text")).casefold())
            for r in rows
        }
        for kind, source_key, text_key, page_key in specs:
            text = clean(src_ev.get(text_key))
            if not text:
                continue
            page = clean(src_ev.get(page_key)) or _message(messages, "evidence_page_unknown")
            key = (kind.casefold(), clean(page).casefold(), text.casefold())
            if key in existing:
                continue
            rows.append({
                "kind": kind,
                "source": _message(messages, source_key),
                "page": page,
                "lot": clean(lot) or _message(messages, "evidence_lot_default"),
                "qualification": _message(messages, "evidence_qualification_project_context") or _message(messages, "evidence_qualification"),
                "text": text,
            })
            existing.add(key)
    return rows


def _preferred_demand_evidence(
    result: Mapping[str, Any],
    evidence: Sequence[Mapping[str, Any]],
) -> Mapping[str, Any] | None:
    """Choisit mécaniquement la preuve à restituer dans « Ce qui est demandé ».

    Une règle documentaire Excel peut avoir isolé une source précise sans fournir
    elle-même un texte de demande. Dans ce cas, cette source est prioritaire.
    À défaut, l'ordre normal des preuves retenues est conservé. Aucun bloc, lot,
    nom de document ou vocabulaire métier n'est codé ici.
    """
    rows = [row for row in evidence or [] if isinstance(row, Mapping) and clean(row.get("text"))]
    if not rows:
        return None

    requested = clean(result.get("specific_source")).upper()
    primary = clean(result.get("primary_evidence_source")).upper()
    if primary in {"MATCHED", "SPECIFIC"}:
        primary = requested
    if primary in {"CCTP", "CONVENTION", "COMPLEMENT"}:
        requested = primary

    def _matches_source(row: Mapping[str, Any], source_code: str) -> bool:
        kind = clean(row.get("kind")).upper()
        if source_code == "CCTP":
            return kind == "CCTP"
        if source_code == "CONVENTION":
            return kind in {"CONV", "CONVENTION"}
        if source_code == "COMPLEMENT":
            return kind.startswith("SUPP") or kind == "COMPLEMENT"
        return False

    if requested in {"CCTP", "CONVENTION", "COMPLEMENT"}:
        for row in rows:
            if _matches_source(row, requested):
                return row

    # Une pièce complémentaire réellement retenue pour le bloc correspond à une
    # nouvelle preuve analysée : elle doit pouvoir alimenter directement la demande
    # publique, indépendamment de son nom de fichier ou de son type documentaire.
    supp_kind = clean(result.get("supp_kind")).upper()
    if clean(result.get("supp_pf")) and supp_kind:
        for row in rows:
            if clean(row.get("kind")).upper() == supp_kind:
                return row

    # À niveau de règle égal, une preuve positive du CCTP est plus directement
    # rattachée au lot qu'une clause générale de Convention. Ce choix ne dépend
    # d'aucun id_bloc : il utilise uniquement l'indicateur de preuve produit par
    # le moteur documentaire.
    if clean(result.get("cctp_preuve_positive")).lower() in {"oui", "yes", "true", "1"}:
        for row in rows:
            if _matches_source(row, "CCTP"):
                return row
    return rows[0]


def _compact_evidence_text(value: Any, max_chars: int = 760) -> str:
    """Réduit uniquement la longueur d'affichage sans paraphraser la preuve."""
    text = clean(value)
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars].rstrip()
    # Préférer une fin de phrase raisonnablement proche de la limite.
    last_stop = max(cut.rfind(". "), cut.rfind("; "), cut.rfind(": "))
    if last_stop >= int(max_chars * 0.55):
        cut = cut[: last_stop + 1].rstrip()
    return cut.rstrip(" ,;:") + "…"


def _evidence_backed_demand(
    result: Mapping[str, Any],
    evidence: Sequence[Mapping[str, Any]],
    lot: str,
    status_key: str,
    messages: Mapping[str, str] | None,
) -> str:
    """Construit une demande publique fidèle à la preuve lorsqu'aucun texte
    documentaire spécifique n'est disponible.

    Le texte d'enveloppe reste dans 12_Messages_Moteur. La clause contractuelle
    est reprise sans interprétation métier afin d'éviter de réintroduire un texte
    générique ou une paraphrase non démontrée.
    """
    if status_key not in {"CONFIRMED", "CLARIFY"}:
        return ""
    selected = _preferred_demand_evidence(result, evidence)
    if not selected:
        return ""
    proof = _compact_evidence_text(selected.get("text"))
    if not proof:
        return ""
    ctx = {
        "lot": clean(lot) or _message(messages, "evidence_lot_default"),
        "source": clean(selected.get("source")),
        "page": clean(selected.get("page")) or _message(messages, "evidence_page_unknown"),
        "preuve": proof,
    }
    key = "demand_from_evidence_confirmed" if status_key == "CONFIRMED" else "demand_from_evidence_clarify"
    return _message(messages, key, **ctx) or _message(messages, "demand_from_evidence_default", **ctx)


def _rejected_evidence(result: Mapping[str, Any], messages: Mapping[str, str] | None) -> List[Dict[str, str]]:
    rows = []
    for item in result.get("preuves_rejetees") or []:
        if isinstance(item, Mapping):
            rows.append({"text": clean(item.get("texte")), "reason": clean(item.get("raison"))})
        else:
            rows.append({"text": clean(item), "reason": _message(messages, "rejected_evidence_default")})
    return rows


def _origin_from_evidence(evidence: List[Mapping[str, Any]], messages: Mapping[str, str] | None) -> str:
    if not evidence:
        return _message(messages, "origin_no_evidence")
    first = evidence[0]
    source, page = clean(first.get("source")), clean(first.get("page"))
    return f"{source} p.{page}" if source and page else source


def _gap_response(capacity: Mapping[str, Any]) -> Mapping[str, Any] | None:
    # Si Excel a ciblé des questions nécessaires à l'exigence détectée,
    # l'action de mise en conformité s'appuie sur cet écart ciblé sans
    # recalculer la capacité globale de l'entreprise.
    source = capacity.get("gap_responses") if capacity.get("requirement_question_ids") else capacity.get("responses")
    for response in source or []:
        if response.get("reponse_val") in (0, 1):
            return response
    return None


def _normalize_question(value: Any, messages: Mapping[str, str] | None) -> str:
    question = clean_generated_text(value)
    if not question:
        return ""
    if "?" not in question:
        suffix = _message(messages, "question_suffixe")
        if suffix:
            question = question.rstrip(" .") + ". " + suffix
    return clean(question)


def _proposal(
    mode: str,
    block_id: str,
    capacity: Mapping[str, Any],
    result: Mapping[str, Any],
    proposals: Mapping[str, Mapping[str, Any]],
    messages: Mapping[str, str] | None,
) -> Tuple[str, str]:
    bloc_row = result.get("bloc")
    if bloc_row is None:
        bloc_row = {}
    proposal = proposals.get(block_id) or {}
    commercial = (
        clean(result.get("specific_response"))
        or clean_generated_text(proposal.get("texte"))
        or clean(bloc_row.get("proposition_commerciale"))
        or _message(messages, "response_commercial_fallback")
    )
    title = clean(bloc_row.get("titre_bloc")) or block_id
    key = clean(mode).lower()
    label = _message(messages, f"response_{key}_label", title=title)
    if mode in {"CONFIRMED", "CONTRADICTION"}:
        return label, commercial
    if mode == "CLARIFY" and clean(result.get("specific_response")):
        return label, clean(result.get("specific_response"))
    if mode == "EXCLUDED":
        specific = clean(result.get("interblock_reason")) or clean(result.get("specific_response"))
        if specific:
            return label, specific
    text = _message(messages, f"response_{key}_text", title=title)
    return label, text


def _action_items(
    action_mode: str,
    block_id: str,
    capacity: Mapping[str, Any],
    result: Mapping[str, Any],
    evidence: List[Mapping[str, Any]],
    messages: Mapping[str, str] | None,
    axes_config: Mapping[str, Any] | None = None,
) -> List[Dict[str, Any]]:
    """Construit les actions publiques et leur temporalité.

    La phase ne découle pas du statut ou de la priorité : elle vient du bloc
    (10_Blocs) ou d'une règle documentaire spécifique. Ainsi une obligation
    confirmée peut être intégrée à l'offre tout en n'étant exécutable qu'après
    attribution. Les écarts de capacité et clarifications conservent leur phase
    propre, généralement avant l'engagement.
    """
    bloc_row = result.get("bloc") or {}
    title = clean(bloc_row.get("titre_bloc")) or block_id
    origin = clean(result.get("specific_action_origin")) or _origin_from_evidence(evidence, messages)
    items: List[Dict[str, Any]] = []

    def _phase(code: str) -> Dict[str, Any]:
        code = clean(code)
        return {
            "phase_code": code,
            "phase_label": _axis_label(axes_config, "phase_action", code) if code else "",
            "phase_short": _axis_field(axes_config, "phase_action", code, "label_court") if code else "",
            "phase_note": _axis_field(axes_config, "phase_action", code, "note_restitution") if code else "",
            "phase_order": _axis_order(axes_config, "phase_action", code) if code else 9999,
            "phase_color": _axis_color(axes_config, "phase_action", code) if code else "",
        }

    standard_phase = clean(result.get("specific_action_phase")) or clean(bloc_row.get("phase_action_standard"))
    capacity_phase = clean(bloc_row.get("phase_action_capacite"))
    clarify_phase = clean(bloc_row.get("phase_action_a_confirmer"))

    def _add(text: str, item_origin: str, kind: str, phase_code: str = "") -> None:
        text = clean(text)
        if not text:
            return
        items.append({"text": text, "origin": item_origin, "kind": kind, **_phase(phase_code)})

    if action_mode == "EXCLUDED_EXPLICIT":
        _add(_message(messages, "action_excluded_explicit", title=title), _message(messages, "origin_excluded_explicit"), "document")
    elif action_mode == "EXCLUDED_EVIDENCE":
        _add(_message(messages, "action_excluded_evidence", title=title), origin, "document")
    elif action_mode == "EXCLUDED_NO_EVIDENCE":
        _add(_message(messages, "action_excluded_no_evidence", title=title), _message(messages, "origin_no_evidence"), "clarification", clarify_phase)
    elif action_mode == "UNDEMONSTRATED":
        text = clean(bloc_row.get("action_interne_non_demontree")) or _message(messages, "action_undemonstrated_fallback", title=title)
        _add(text, origin, "clarification", clarify_phase)
    elif action_mode == "CLARIFY":
        text = clean(result.get("specific_action")) or clean(bloc_row.get("action_interne_a_confirmer")) or _message(messages, "action_clarify_fallback", title=title)
        _add(text, origin, "clarification", clarify_phase)
    elif action_mode in {"STANDARD_ONLY", "STANDARD_AND_CAPACITY"}:
        standard = clean(result.get("specific_action")) or clean(bloc_row.get("action_interne_standard")) or _message(messages, "action_generique_standard", title=title)
        _add(standard, origin, "document", standard_phase)

        # Cycle 50 : une action documentaire et un écart de capacité sont deux
        # informations indépendantes. Dès qu'une question effectivement reliée
        # à l'exigence possède une action spécifique Non/Partiel, cette action
        # est restituée, même si une ancienne règle documentaire avait forcé
        # STANDARD_ONLY. Ainsi une réponse Oui ne génère rien, tandis que chaque
        # écart réel produit uniquement l'action de SA question. Le fallback de
        # bloc n'est conservé que pour l'ancien mode STANDARD_AND_CAPACITY et
        # seulement si aucune action question-spécifique n'existe.
        gap_source = capacity.get("gap_responses") if capacity.get("requirement_question_ids") else capacity.get("responses")
        gaps = [g for g in (gap_source or []) if g.get("reponse_val") in (0, 1)]
        fallback_added = False
        for gap in gaps:
            qid, label = clean(gap.get("id_question")), clean(gap.get("reponse_label"))
            origin_key = "origin_autoeval" if qid else "origin_autoeval_sans_id"
            cap_text = clean(gap.get("action_capacite"))
            if cap_text:
                _add(cap_text, _message(messages, origin_key, qid=qid, label=label), "capacity", capacity_phase)
            elif action_mode == "STANDARD_AND_CAPACITY" and not fallback_added:
                fallback = clean(bloc_row.get("action_interne_capacite")) or _message(messages, "action_generique_capacite", title=title)
                _add(fallback, _message(messages, origin_key, qid=qid, label=label), "capacity", capacity_phase)
                fallback_added = True

    unique, seen = [], set()
    for item in items:
        text = clean(item.get("text"))
        key = (text.casefold(), clean(item.get("phase_code")).casefold())
        if text and key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def ensure_public_model(
    results: MutableMapping[str, MutableMapping[str, Any]],
    scores: Mapping[str, Mapping[str, Any]] | None = None,
    proposals: Mapping[str, Mapping[str, Any]] | None = None,
    lot: str = "",
    axes_config: Mapping[str, Any] | None = None,
    restitution_rules: Sequence[Mapping[str, Any]] | None = None,
    messages: Mapping[str, str] | None = None,
) -> MutableMapping[str, MutableMapping[str, Any]]:
    """Construit la restitution publique à partir des seules feuilles Excel.

    - 16_Axes_Config : libellés, ordres et seuils de capacité ;
    - 17_Regles_Restitution : statut public, priorité, décision, mode de réponse,
      mode d'action et affichage des questions ;
    - 10_Blocs / 60_Reponses : textes par bloc ;
    - 12_Messages_Moteur : messages génériques.
    """
    scores, proposals = scores or {}, proposals or {}
    best_capacity = _best_capacity_code(axes_config)

    for block_id, result in results.items():
        bloc_row = result.get("bloc")
        if bloc_row is None:
            bloc_row = {}
        capacity = _capacity(block_id, result, scores, axes_config, messages)
        evidence = _evidence(result, lot=lot, messages=messages)
        selected_for_requirement = bool(capacity.get("requirement_question_ids"))
        gap_source = capacity.get("gap_responses") if selected_for_requirement else capacity.get("responses")
        has_gap = any(
            r.get("reponse_val") in (0, 1)
            for r in (gap_source or [])
        )
        base_context = {
            "statut": clean(result.get("statut")),
            "presence": clean(result.get("presence_contractuelle")),
            "applicabilite": clean(result.get("applicabilite_lot")),
            "capacite": capacity.get("code", ""),
            "contradiction": bool(result.get("contradictions")),
            "exclusion_explicit": bool(result.get("exclusion_explicit")),
            "preuve": bool(evidence),
            "capacite_best": capacity.get("code") == best_capacity if best_capacity else False,
            "ecart_capacite": has_gap,
            "interblock_exclusion": bool(result.get("interblock_exclusion")),
        }
        # Les éventuels forçages proviennent du paramétrage Excel (règles interblocs
        # ou documentaires). Une nuance CCTP peut également avoir un statut public
        # propre défini dans 10_Blocs, sans être assimilée à une preuve positive.
        # Python reste un moteur générique et n'encode aucun id_bloc.
        nuance_only_status = ""
        nuance_detected = clean(result.get("cctp_nuance_detectee")).lower() in {"oui", "yes", "true", "1"}
        positive_cctp = clean(result.get("cctp_preuve_positive")).lower() in {"oui", "yes", "true", "1"}
        if nuance_detected and not positive_cctp:
            nuance_only_status = clean(bloc_row.get("statut_public_si_nuance_seule"))
        status_key = (
            clean(result.get("statut_public_override"))
            or nuance_only_status
            or choose_rule(restitution_rules, "statut_public", base_context)
        )
        context = {**base_context, "statut_public": status_key}
        priority_code = clean(result.get("priorite_override")) or choose_rule(restitution_rules, "priorite", context)
        decision_code = clean(result.get("decision_override")) or choose_rule(restitution_rules, "decision", context)
        response_mode = clean(result.get("response_mode_override")) or choose_rule(restitution_rules, "response_mode", context)
        action_mode = clean(result.get("action_mode_override")) or choose_rule(restitution_rules, "action_mode", context)
        question_mode = clean(result.get("question_mode_override")) or choose_rule(restitution_rules, "question_mode", context)

        response_label, response_text = _proposal(
            response_mode, block_id, capacity, result, proposals, messages
        )
        # Une occurrence contextuelle jugée insuffisante ne doit pas être affichée
        # comme « preuve contractuelle » lorsque le statut public est Non démontré.
        # Elle peut encore alimenter la formulation spécifique du bloc, mais pas
        # la liste publique des preuves de l'exigence.
        display_evidence = [] if status_key == "UNDEMONSTRATED" else [dict(ev) for ev in evidence]
        # DEV02 — qualité de preuve : une clause affichée sous un bloc exclu ou
        # sous une clarification de niveau projet ne doit jamais être présentée
        # comme une obligation directe du lot. Ce traitement est transversal :
        # il ne connaît ni id_bloc ni métier.
        if status_key == "EXCLUDED":
            q = _message(messages, "evidence_qualification_excluded_context")
            if q:
                for ev in display_evidence:
                    ev["qualification"] = q
            # La pièce la plus directement rattachée au lot doit être lue en
            # premier ; la Convention générale reste ensuite comme contexte.
            display_evidence.sort(key=lambda ev: 0 if clean(ev.get("kind")).lower() == "cctp" else 1)
        elif status_key == "CLARIFY":
            positive_cctp = clean(result.get("cctp_preuve_positive")).lower() in {"oui", "yes", "true", "1"}
            if not positive_cctp:
                q = _message(messages, "evidence_qualification_clarify_project")
                if q:
                    for ev in display_evidence:
                        if clean(ev.get("kind")).lower() in {"conv", "convention"}:
                            ev["qualification"] = q
        if status_key == "CONFIRMED" and clean(result.get("cctp_preuve_positive")).lower() in {"oui", "yes", "true", "1"}:
            display_evidence.sort(key=lambda ev: 0 if clean(ev.get("kind")).lower() == "cctp" else 1)
        evidence_demand = _evidence_backed_demand(
            result, display_evidence, lot=lot, status_key=status_key, messages=messages
        )
        action_items = _action_items(
            action_mode, block_id, capacity, result, display_evidence, messages, axes_config
        )
        title = clean(bloc_row.get("titre_bloc")) or block_id
        question = ""
        if question_mode == "SHOW":
            # Une règle interblocs ou documentaire peut fournir une question plus
            # précise. Le moteur reste générique : le texte et sa condition sont
            # entièrement paramétrés dans Excel.
            question = _normalize_question(
                clean(result.get("specific_question")) or clean(bloc_row.get("question_bim_manager")),
                messages,
            )

        result.update({
            "public_id": block_id,
            "public_title": title,
            "public_demand": (
                clean(result.get("interblock_reason"))
                if status_key in {"EXCLUDED", "UNDEMONSTRATED"} and result.get("interblock_reason")
                else (
                    clean(result.get("specific_demand"))
                    or evidence_demand
                    or (
                        clean(_message(messages, "response_undemonstrated_text")).split(".", 1)[0] + "."
                        if status_key == "UNDEMONSTRATED" and not display_evidence
                        else clean(bloc_row.get("lecture_entreprise")) or clean(result.get("conclusion"))
                    )
                )
            ),
            "public_status_key": status_key,
            "public_status_label": _axis_label(axes_config, "statut_public", status_key),
            "public_group_order": _axis_order(axes_config, "statut_public", status_key),
            "public_capacity_state": capacity.get("code", ""),
            "public_capacity_label": capacity.get("label", ""),
            "public_capacity_pct": capacity.get("pct"),
            "public_capacity_strength": capacity.get("strength", ""),
            "public_capacity_gap": capacity.get("gap", ""),
            "public_requirement_coverage_pct": capacity.get("requirement_coverage_pct"),
            "public_requirement_coverage_evaluated": capacity.get("requirement_coverage_evaluated", False),
            "public_requirement_coverage_label": capacity.get("requirement_coverage_label", ""),
            "public_requirement_question_ids": capacity.get("requirement_question_ids", []),
            "public_requirement_missing_question_ids": capacity.get("requirement_missing_question_ids", []),
            "public_priority_code": priority_code,
            "public_priority": _axis_label(axes_config, "priorite", priority_code),
            "public_priority_order": _axis_order(axes_config, "priorite", priority_code),
            "public_priority_description": _axis_description(axes_config, "priorite", priority_code),
            "public_priority_color": _axis_color(axes_config, "priorite", priority_code),
            "public_priority_short": _axis_field(axes_config, "priorite", priority_code, "label_court"),
            "public_priority_note": _axis_field(axes_config, "priorite", priority_code, "note_restitution"),
            "public_decision_code": decision_code,
            "public_decision": _axis_label(axes_config, "decision", decision_code),
            "public_response_label": response_label,
            "public_response_text": response_text,
            "public_action_items": action_items,
            "public_actions": [item["text"] for item in action_items],
            "public_action_phases": [item.get("phase_code", "") for item in action_items if item.get("phase_code")],
            "public_question": question,
            "public_evidence": display_evidence,
            "public_rejected_evidence": _rejected_evidence(result, messages),
            "public_sort_key": block_sort_key(block_id),
        })
    return results


def sorted_items(results: Mapping[str, Mapping[str, Any]]) -> List[Tuple[str, Mapping[str, Any]]]:
    return sorted(
        results.items(),
        key=lambda item: (item[1].get("public_group_order", 9999), block_sort_key(item[0])),
    )
