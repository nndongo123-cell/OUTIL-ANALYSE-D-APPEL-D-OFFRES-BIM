#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard HTML — outil BIM v1
Style : dossier d’ingénierie, titres Calibri Light et contenu Calibri Regular.
"""
import html, json, re
from content_rules import FALLBACK_GLOSSARY, glossary_lookup
from typing import Dict, List

def esc(s) -> str:
    return html.escape(str(s or ""))

def _build_glossaire_lookup(glossaire_complet: List[Dict]) -> Dict[str, Dict]:
    """Indexe termes, synonymes et mots-clés, avec un socle de définitions
    pour les documents, logiciels et formats courants."""
    entries = []
    for g in glossaire_complet or []:
        term = str(g.get("terme") or g.get("term") or "").strip()
        definition = str(g.get("definition_excel") or g.get("definition") or "").strip()
        if not term or not definition:
            continue
        aliases = []
        for field in ("mots_cles", "synonymes"):
            aliases.extend(x.strip() for x in re.split(r"[;,|]", str(g.get(field, ""))) if x.strip())
        entries.append({
            "term": term, "definition": definition,
            "source": g.get("source_excel") or g.get("source") or "",
            "category": g.get("categorie") or "",
            "aliases": aliases,
        })
    for term, cfg in FALLBACK_GLOSSARY.items():
        entries.append({"term": term, "definition": cfg.get("definition", ""), "source": "Glossaire de l'outil", "aliases": cfg.get("aliases", [])})
    raw = glossary_lookup(entries)
    # Adapter la structure attendue par tip().
    return {key: {
        "terme": value.get("term"), "definition": value.get("definition"),
        "source": value.get("source"), "categorie": value.get("category", "")
    } for key, value in raw.items()}

def _make_tip(lookup: Dict[str, Dict]):
    """
    Retourne une fonction tip(terme_affiche, terme_glossaire=None) qui
    enveloppe un terme technique avec une bulle d'info SI une définition
    existe dans le glossaire Excel (40_Glossaire). Sinon, retourne le texte
    brut échappé -- jamais de définition codée en dur dans le template.
    Ajouter/corriger une définition se fait uniquement dans l'Excel.
    """
    def tip(terme_affiche: str, terme_glossaire: str = None) -> str:
        disp = esc(terme_affiche)
        key = (terme_glossaire or terme_affiche).strip().lower()
        g = lookup.get(key)
        if not g:
            return disp
        d = esc(g.get("definition_excel") or g.get("definition") or "")
        if not d:
            return disp
        src = esc(g.get("source_excel") or g.get("source") or "")
        categorie = esc(g.get("categorie", ""))
        niveau = esc(g.get("niveau", ""))
        tag = " · ".join(t for t in [categorie, niveau] if t)
        src_html = f'<div class="tip-src">{src}{f" · {tag}" if tag else ""}</div>' if (src or tag) else ""
        return (f'<span class="tip-w" tabindex="0"><span class="tip">{disp}</span>'
                f'<div class="tip-b">{d}{src_html}</div></span>')
    return tip

def _make_autotip(lookup: Dict[str, Dict], tip):
    """
    Retourne une fonction autotip(texte_html_echappe) qui repère AUTOMATIQUEMENT
    les termes du glossaire à l'intérieur d'un texte libre déjà échappé
    (citations, questions, paragraphes rédigés dans l'Excel...) et les
    enveloppe avec tip(), sans qu'on ait à les lister un par un dans le code.

    Exclut volontairement les entrées "Outil BIM — méthodologie" (Confirmées,
    Bloc critique, etc.) : ce sont des concepts de l'outil, pas du vocabulaire
    BIM à détecter dans un texte contractuel -- les y chercher produirait de
    faux positifs (ex: le mot "confirmée" dans une phrase ordinaire).

    Une seule passe sur le texte (regex à alternance unique) pour éviter
    qu'un terme déjà enveloppé soit ré-enveloppé par un terme suivant.
    """
    termes_domaine = [
        g.get("terme", "") for g in lookup.values()
        if str(g.get("source_excel") or g.get("source") or "").strip().lower()
        != "outil bim — méthodologie"
        and str(g.get("terme", "")).strip()
    ]
    if not termes_domaine:
        return lambda txt: txt
    # Plus long terme d'abord pour que "IFC 2x3" soit détecté avant "IFC" seul
    termes_domaine = sorted(set(termes_domaine), key=len, reverse=True)
    pattern = re.compile(
        r'\b(' + '|'.join(re.escape(t) for t in termes_domaine) + r')\b',
        re.IGNORECASE
    )
    def autotip(texte_html: str) -> str:
        if not texte_html:
            return texte_html
        # Une seule bulle par terme et par bloc de texte : au-delà de la
        # première occurrence, le mot reste en texte brut. Évite qu'un terme
        # générique comme "BIM" soit surligné cinq fois dans le même
        # paragraphe, y compris lorsqu'un terme plus spécifique contenant
        # le même mot (ex. "BIM Manager") est déjà annoté juste à côté.
        seen: set[str] = set()

        def _replace(m: "re.Match[str]") -> str:
            word = m.group(1)
            key = word.strip().lower()
            if key in seen:
                return word
            seen.add(key)
            return tip(word)

        return pattern.sub(_replace, texte_html)
    return autotip


def _render_docs_analyses(meta: dict, tip=None) -> str:
    """
    Génère la liste HTML des documents analysés.
    Pour chaque document (convention, CCTP), affiche :
    nom de fichier · indice · date · émetteur · MO
    — tous détectés depuis la page de garde et les en-têtes du PDF.
    """
    lignes = []
    tip = tip or (lambda value, *_: esc(value))

    # ── Document 1 : Convention BIM ──────────────────────────────────────────
    nom_conv  = esc(meta.get("projet", "Convention BIM"))
    indice    = meta.get("indice_doc", "")
    date_pg   = meta.get("date_doc", "")
    emetteur  = meta.get("emetteur_doc", "")
    mo        = meta.get("maitre_ouvrage_doc", "")
    titre     = meta.get("titre_doc", "")
    num_aff   = meta.get("numero_affaire", "")

    # Ligne de détail : champs non vides séparés par ·
    details_conv = []
    if indice:   details_conv.append(f"Indice {esc(indice)}")
    if date_pg:  details_conv.append(esc(date_pg))
    if emetteur: details_conv.append(esc(emetteur))
    if mo:       details_conv.append(f"MO : {esc(mo)}")
    if num_aff:  details_conv.append(f"N° {esc(num_aff)}")
    if titre and titre.lower() != nom_conv.lower():
        details_conv.insert(0, f"<em>{esc(titre)}</em>")

    detail_str_conv = (" · ".join(details_conv)) if details_conv else "—"

    lignes.append(f'''<div style="margin-bottom:9px;padding-bottom:9px;border-bottom:0.5px solid #eee">
        <div style="font-size:11px;font-weight:600;color:#071C24;margin-bottom:3px">{tip("Convention BIM")}</div>
        <div style="font-size:10px;color:#888;margin-bottom:1px">{esc(nom_conv)}</div>
        <div style="font-size:10px;color:#666">{detail_str_conv}</div>
      </div>''')

    # ── Document 2 : CCTP ────────────────────────────────────────────────────
    n_pgs      = meta.get("n_pgs", 0)
    nom_cctp   = meta.get("cctp_nom", "CCTP du lot")
    ind_cctp   = meta.get("cctp_indice", "")
    date_cctp  = meta.get("cctp_date", "")
    aut_cctp   = meta.get("cctp_auteur", "")
    phase_cctp = meta.get("cctp_phase", "")

    if n_pgs:
        details_cctp = []
        if ind_cctp:   details_cctp.append(f"Indice {esc(ind_cctp)}")
        if date_cctp:  details_cctp.append(esc(date_cctp))
        if phase_cctp: details_cctp.append(esc(phase_cctp))
        if aut_cctp:   details_cctp.append(f"Auteur : {esc(aut_cctp)}")
        details_cctp.append(f"{n_pgs} pages")
        detail_str_cctp = " · ".join(details_cctp)

        lignes.append(f'''<div>
        <div style="font-size:11px;font-weight:600;color:#071C24;margin-bottom:3px">{tip("CCTP du lot", "CCTP")}</div>
        <div style="font-size:10px;color:#888;margin-bottom:1px">{esc(nom_cctp)}</div>
        <div style="font-size:10px;color:#666">{detail_str_cctp}</div>
      </div>''')
    else:
        lignes.append('''<div style="font-size:10.5px;color:#aaa;font-style:italic">CCTP non fourni</div>''')

    return "\n".join(lignes)


def _render_lots(lots: list) -> str:
    """Génère le bloc HTML des lots de l'entreprise."""
    if not lots:
        return ""
    items = ""
    for i, l in enumerate(lots):
        border = "border:1px solid #071C24" if i == 0 else "border:0.5px solid #ccc"
        bg = "" if i == 0 else "background:#FAFAFA;"
        col = "#071C24" if i == 0 else "#555"
        items += (
            f'<div style="{border};padding:5px 12px;{bg}">' +
            '<div style="font-size:9px;color:#888;margin-bottom:1px">Lot</div>' +
            f'<div style="font-size:11px;font-weight:600;color:{col}">' +
            html.escape(l) + '</div></div>'
        )
    return (
        '<div style="margin-top:12px">' +
        '<div class="cap" style="margin-bottom:8px">Lots de l\'entreprise sur ce projet</div>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap">' + items + '</div>' +
        '</div>'
    )


# ============================================================================
# NOTE V10 : quatre anciennes définitions de build_dashboard() (versions de
# travail successives, ~1600 lignes) ont été supprimées ici. Seule la dernière
# définition (ci-dessous) était effectivement exécutée -- en Python, une
# fonction redéfinie plusieurs fois au niveau module écrase silencieusement
# les définitions précédentes. Les anciennes versions n'étaient donc jamais
# appelées, mais contenaient des textes et mappings B01-B12 codés en dur
# (dupliqués depuis 10_Blocs) qui pouvaient induire en erreur quiconque
# modifiait le fichier en pensant agir sur le comportement réel de l'outil.
# ============================================================================

def build_dashboard(
    resultats_lecture: Dict,
    scores_eval: Dict,
    textes_reponse: Dict,
    glossaire: List[Dict],
    meta: Dict,
    conv_text: str = "",
    conv_b64: str = "",
    cctp_b64: str = "",
    pdf_filename: str = "",
    glossaire_complet: List[Dict] = None,
) -> str:
    import math
    from datetime import datetime
    from bim_model import block_sort_key, ensure_public_model

    ensure_public_model(resultats_lecture, scores_eval, textes_reponse, lot=str(meta.get("lot", "")), axes_config=meta.get("axes_config"))
    lookup = _build_glossaire_lookup(glossaire_complet if glossaire_complet is not None else glossaire)
    tip = _make_tip(lookup)
    autotip = _make_autotip(lookup, tip)

    def rich(value) -> str:
        return autotip(esc(value))

    def evidence_html(result: Dict) -> str:
        evidence = result.get("public_evidence") or []
        if not evidence:
            return '<div class="no-proof">Aucune preuve ciblée et recevable n’a été trouvée pour le lot sélectionné.</div>'
        cards = []
        for ev in evidence:
            source = esc(ev.get("source", "Document"))
            page = esc(ev.get("page", "Non précisée"))
            lot = esc(ev.get("lot", "Lot analysé"))
            text = rich(ev.get("text", ""))
            kind = "cctp" if "CCTP" in ev.get("source", "").upper() else "conv"
            page_num = re.search(r"\d+", str(ev.get("page", "")))
            button = ""
            if page_num:
                button = f'<button type="button" class="source-button" data-pdf="{kind}" data-page="{page_num.group(0)}">Ouvrir la source p.{page_num.group(0)}</button>'
            cards.append(f'''<article class="proof-card">
              <div class="proof-meta"><b>{source}</b><span>Page {page}</span><span>Portée : {lot}</span><span>Qualification : extrait ciblé retenu</span></div>
              <blockquote>{text}</blockquote>{button}
            </article>''')
        return "".join(cards)

    def block_html(block_id: str, result: Dict) -> str:
        block = result.get("bloc") or {}
        pct = result.get("public_capacity_pct")
        prep = f"{pct} %" if isinstance(pct, (int, float)) else "Non chiffrée"
        demand = result.get("public_demand") or block.get("lecture_entreprise") or block.get("description_convention") or result.get("conclusion") or "Exigence à analyser."
        question = result.get("public_question") or ""
        action_items = result.get("public_action_items") or [{"text": action, "origin": ""} for action in (result.get("public_actions") or [])]
        rendered_actions = []
        for item in action_items:
            origin = item.get("origin", "") if isinstance(item, dict) else ""
            text = item.get("text", "") if isinstance(item, dict) else str(item)
            origin_html = f'<small>Origine : {rich(origin)}</small>' if origin else ""
            rendered_actions.append(f'<li><span>{rich(text)}</span>{origin_html}</li>')
        actions_html = "".join(rendered_actions)
        question_html = f'<div class="question-box"><b>Question à transmettre</b><p>{rich(question)}</p></div>' if question else ""
        contradictions = result.get("contradictions") or []
        contradictions_html = ""
        if contradictions:
            items = "".join(f"<li>{rich(item.get('message') if isinstance(item, dict) else item)}</li>" for item in contradictions)
            contradictions_html = f'<div class="warning-box"><b>Contradiction détectée</b><ul>{items}</ul></div>'
        status_key = result.get("public_status_key", "UNDEMONSTRATED")
        status_class = {
            "CONFIRMED": "confirmed", "CLARIFY": "clarify",
            "UNDEMONSTRATED": "undemonstrated", "EXCLUDED": "excluded",
        }.get(status_key, "undemonstrated")
        if status_key == "EXCLUDED":
            # Un bloc exclu n'a, par définition, ni capacité à démontrer ni
            # réponse à formuler : une carte à 4 sections donnerait l'impression
            # trompeuse qu'il reste beaucoup à analyser. Une seule section
            # explique pourquoi et rappelle la seule vigilance utile.
            reason = result.get("public_response_text") or result.get("conclusion") or "Non identifiée comme exigence applicable au lot dans les documents analysés."
            first_action = action_items[0] if action_items else None
            action_text = (first_action.get("text") if isinstance(first_action, dict) else str(first_action)) if first_action else "Vérifier, avant remise de l'offre, qu'aucun ordre écrit ou additif ne réintroduit cette prestation."
            action_origin = (first_action.get("origin", "") if isinstance(first_action, dict) else "") if first_action else ""
            body = f'''<div class="req-body">
            <div class="req-meta-strip">
              <div><span>Priorité</span><b>{esc(result.get('public_priority'))}</b></div>
              <div><span>Décision</span><b>{esc(result.get('public_decision'))}</b></div>
              <div><span>Origine</span><b>{esc(action_origin or 'À confirmer')}</b></div>
            </div>
            <section class="excluded-note"><h3>Pourquoi ce bloc est exclu</h3><p>{rich(reason)}</p>
              <p class="excluded-check">{rich(action_text)}</p>
            </section>
          </div>'''
            return f'''<details class="req-card" data-block="{esc(block_id)}" data-status="{status_class}">
          <summary class="req-summary">
            <span class="req-code">{esc(block_id)}</span>
            <span class="req-title">{esc(result.get('public_title', block_id))}</span>
            <span class="status-badge {status_class}">{esc(result.get('public_status_label'))}</span>
            <span class="capacity-badge">—</span>
            <span class="decision-text">{esc(result.get('public_decision'))}</span>
          </summary>
          {body}
        </details>'''
        return f'''<details class="req-card" data-block="{esc(block_id)}" data-status="{status_class}">
          <summary class="req-summary">
            <span class="req-code">{esc(block_id)}</span>
            <span class="req-title">{esc(result.get('public_title', block_id))}</span>
            <span class="status-badge {status_class}">{esc(result.get('public_status_label'))}</span>
            <span class="capacity-badge">{esc(result.get('public_capacity_label'))}</span>
            <span class="decision-text">{esc(result.get('public_decision'))}</span>
          </summary>
          <div class="req-body">
            <div class="req-meta-strip">
              <div><span>Priorité</span><b>{esc(result.get('public_priority'))}</b></div>
              <div><span>Préparation</span><b>{esc(prep)}</b></div>
              <div><span>Décision</span><b>{esc(result.get('public_decision'))}</b></div>
            </div>
            <div class="detail-grid">
              <section><h3>Ce qui est demandé</h3><p>{rich(demand)}</p></section>
              <section><h3>Ce que je suis capable de faire en ce moment</h3><p><b>{esc(result.get('public_capacity_label'))}</b></p><p>{rich(result.get('public_capacity_strength'))}</p>{f'<div class="gap-box"><b>Point à renforcer</b><p>{rich(result.get("public_capacity_gap"))}</p></div>' if result.get('public_capacity_gap') else ''}</section>
              <section><h3>Ce que je dois répondre</h3><span class="response-label">{esc(result.get('public_response_label'))}</span><blockquote>{rich(result.get('public_response_text'))}</blockquote></section>
              <section><h3>Ce que je dois faire pour être conforme</h3><ul class="action-list">{actions_html}</ul>{question_html}</section>
            </div>
            {contradictions_html}
            <section class="proofs"><h3>Preuves contractuelles</h3>{evidence_html(result)}</section>
          </div>
        </details>'''

    group_specs = [
        ("CONFIRMED", "Confirmées pour le lot", "Exigences directement applicables et appuyées par une preuve recevable.", True),
        ("CLARIFY", "À confirmer pour le lot", "Périmètre ou contenu à sécuriser par écrit avant engagement.", True),
        ("UNDEMONSTRATED", "Non démontrées dans les documents", "Aucune preuve ciblée recevable n’établit actuellement l’exigence pour le lot.", True),
        ("EXCLUDED", "Exclues ou non applicables", "Aucun engagement à prendre, sauf ordre écrit ou modification contractuelle.", False),
    ]
    groups_html = []
    for key, label, description, opened in group_specs:
        items = [(bid, res) for bid, res in resultats_lecture.items() if res.get("public_status_key") == key]
        items.sort(key=lambda item: block_sort_key(item[0]))
        if not items:
            continue
        cards = "".join(block_html(bid, res) for bid, res in items)
        groups_html.append(f'''<details class="status-group" data-group="{key}" {'open' if opened else ''}>
          <summary><span><b>{esc(label)}</b><small>{esc(description)}</small></span><strong>{len(items)} bloc(s)</strong></summary>
          <div class="group-tools"><button type="button" data-open-group>Tout ouvrir</button><button type="button" data-close-group>Tout replier</button></div>
          <div class="group-blocks">{cards}</div>
        </details>''')

    confirmed = sum(1 for r in resultats_lecture.values() if r.get("public_status_key") == "CONFIRMED")
    clarify = sum(1 for r in resultats_lecture.values() if r.get("public_status_key") == "CLARIFY")
    undem = sum(1 for r in resultats_lecture.values() if r.get("public_status_key") == "UNDEMONSTRATED")
    excluded = sum(1 for r in resultats_lecture.values() if r.get("public_status_key") == "EXCLUDED")

    nonexcluded = [(bid, r) for bid, r in resultats_lecture.items() if r.get("public_status_key") != "EXCLUDED"]
    evaluated = [(bid, r) for bid, r in resultats_lecture.items() if r.get("public_capacity_pct") is not None]
    eval_pct = round(100 * len(evaluated) / len(nonexcluded)) if nonexcluded else 100
    question_count = sum(len((scores_eval.get(bid) or {}).get("reponses") or []) for bid, _ in evaluated)
    action_count = sum(len(r.get("public_actions") or []) for _, r in nonexcluded)

    # Radar : tous les blocs évalués, triés par numéro croissant.
    radar_items = sorted(evaluated, key=lambda item: block_sort_key(item[0]))
    if len(radar_items) >= 3:
        cx, cy, radius = 220, 195, 142
        n_axes = len(radar_items)
        def point(index: int, ratio: float):
            angle = -math.pi / 2 + (2 * math.pi * index / n_axes)
            return cx + radius * ratio * math.cos(angle), cy + radius * ratio * math.sin(angle)
        grids = []
        for level in (0.2, 0.4, 0.6, 0.8, 1.0):
            pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in (point(i, level) for i in range(n_axes)))
            grids.append(f'<polygon points="{pts}" fill="none" stroke="#D9E3E9" stroke-width="1"/>')
        spokes, labels, dots = [], [], []
        data_points = []
        for i, (bid, result) in enumerate(radar_items):
            value = float(result.get("public_capacity_pct") or 0)
            x, y = point(i, 1.0)
            dx, dy = point(i, 1.17)
            anchor = "middle" if abs(dx-cx) < 10 else ("end" if dx < cx else "start")
            spokes.append(f'<line x1="{cx}" y1="{cy}" x2="{x:.1f}" y2="{y:.1f}" stroke="#D9E3E9" stroke-width="1"/>')
            labels.append(f'<text x="{dx:.1f}" y="{dy:.1f}" text-anchor="{anchor}" dominant-baseline="middle" font-family="Calibri,Arial,sans-serif" font-size="10" font-weight="700" fill="#17232B">{esc(bid)}</text>')
            px, py = point(i, value / 100)
            data_points.append(f"{px:.1f},{py:.1f}")
            dots.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="4" fill="#1F5D8C" stroke="#fff" stroke-width="1.5"><title>{esc(bid)} : {value:g} %</title></circle>')
        radar_svg = f'''<svg viewBox="0 0 440 405" role="img" aria-label="Diagramme radar des capacités déclarées" class="radar-svg">
          {''.join(grids)}{''.join(spokes)}
          <polygon points="{' '.join(data_points)}" fill="rgba(161,0,61,.14)" stroke="#A1003D" stroke-width="2.5"/>
          {''.join(dots)}{''.join(labels)}
          <text x="220" y="397" text-anchor="middle" font-family="Calibri,Arial,sans-serif" font-size="10" fill="#667985">{len(radar_items)} blocs évalués - échelle 0 à 100 %</text>
        </svg>'''
    else:
        radar_svg = '<div class="empty-radar">Le radar sera disponible dès qu’au moins trois blocs auront été évalués.</div>'

    project_rows = [
        ("Projet", meta.get("projet") or "-"), ("Entreprise", meta.get("entreprise") or "-"),
        ("Lot analysé", meta.get("lot") or "-"), ("Édition", meta.get("date") or "-"),
        ("Maître d'ouvrage", meta.get("maitre_ouvrage_doc") or "Non détecté"),
        ("Maîtrise d'œuvre", meta.get("maitrise_oeuvre_doc") or "Non détectée"),
        ("BIM Manager", meta.get("bim_manager_conv") or "Non détecté"),
        ("Coordinateur BIM", meta.get("coordinateur_bim_conv") or "Non détecté"),
    ]
    technical_rows = [
        ("Niveau BIM", meta.get("niveau_bim") or "Non précisé"),
        ("Logiciel", meta.get("logiciel_conv") or "Non précisé"),
        ("Format IFC", meta.get("format_ifc_conv") or "Non précisé"),
        ("LOD / ND", meta.get("nd_conv") or meta.get("lod_conv") or "Non précisé"),
        ("Plateforme CDE", meta.get("plateforme_conv") or "Non précisée"),
        ("Formats livrables", meta.get("formats_livrables") or "Non précisés"),
        ("DOE numérique", "Oui" if meta.get("doe_numerique") else "Non précisé"),
        ("Géoréférencement", "Oui" if meta.get("geo_referencement") else "Non précisé"),
    ]
    def identity_rows(rows):
        return "".join(f'<div class="id-row"><span>{esc(label)}</span><strong>{rich(value)}</strong></div>' for label, value in rows)

    # Le panneau "Glossaire utile" ne doit lister que les termes réellement
    # détectés dans la convention ou le CCTP analysés -- jamais le glossaire
    # de référence complet (qui contient par ex. des noms de plateformes CDE
    # concurrentes non utilisées sur ce projet). `glossaire` est déjà filtré
    # en amont par le moteur ; `glossaire_complet` ne sert qu'aux infobulles
    # inline, pas à ce panneau récapitulatif.
    glossary_rows = []
    for item in glossaire or []:
        term = str(item.get("terme", "")).strip()
        definition = str(item.get("definition_excel") or item.get("definition") or "").strip()
        if term and definition:
            glossary_rows.append((term, definition))
    glossary_rows.sort(key=lambda item: item[0].casefold())
    glossary_html = "".join(f'<div class="gloss-row"><strong>{esc(t)}</strong><span>{esc(d)}</span></div>' for t, d in glossary_rows)

    generated_at = meta.get("generated_at") or datetime.now().strftime("%d/%m/%Y à %H:%M")
    history_html = f'''<div class="history-grid">
      <div><span>Date et heure</span><b>{esc(generated_at)}</b></div>
      <div><span>Lot</span><b>{esc(meta.get('lot') or '-')}</b></div>
      <div><span>Convention</span><b>{esc(meta.get('n_pages_convention') or '?')} page(s)</b></div>
      <div><span>CCTP</span><b>{esc(meta.get('n_pgs') or 0)} page(s)</b></div>
      <div><span>Blocs analysés</span><b>{len(resultats_lecture)}</b></div>
      <div><span>Blocs évalués</span><b>{len(evaluated)}</b></div>
      <div><span>Preuves retenues</span><b>{sum(len(r.get('public_evidence') or []) for r in resultats_lecture.values())}</b></div>
      <div><span>Preuves rejetées</span><b>{sum(len(r.get('public_rejected_evidence') or []) for r in resultats_lecture.values())}</b></div>
    </div>'''

    css = r'''
:root{--navy:#071C24;--blue:#1F5D8C;--slate:#49697D;--pale:#EAF3F8;--pale2:#F7FAFC;--magenta:#A1003D;--magenta-light:#F06A9A;--magenta-pale:#F5C8D8;--text:#17232B;--muted:#667985;--line:#C9D6DE;--green:#176B55;--amber:#A66A12;--white:#fff}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:#F2F5F7;color:var(--text);font:14px/1.5 Calibri,Carlito,"Segoe UI",Arial,sans-serif}button,input{font:inherit}h1,h2,h3,.metric b{font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif}.hero{background:var(--navy);border-bottom:6px solid var(--magenta);padding:30px max(26px,calc((100vw - 1540px)/2));display:flex;align-items:center;gap:26px}.logo{font-size:30px;font-weight:800;color:var(--magenta-light);border-right:1px solid rgba(240,106,154,.5);padding-right:24px}.logo img{max-height:44px;max-width:170px;display:block;object-fit:contain}.hero h1{margin:0;color:var(--magenta-light);font-size:38px;line-height:1.04;font-weight:700}.hero p{margin:9px 0 0;color:var(--magenta-pale);font-size:15px;font-weight:600}.hero nav{margin-left:auto;display:flex;gap:8px;flex-wrap:wrap}.hero a{color:var(--magenta-pale);border:1px solid rgba(240,106,154,.55);text-decoration:none;padding:9px 12px;font-weight:700;font-size:12px}.hero a.primary{background:var(--magenta);color:#fff}.container{max-width:1540px;margin:auto;padding:22px}.overview-grid{display:grid;grid-template-columns:minmax(650px,1.5fr) minmax(390px,.9fr);gap:18px}.panel,.status-group,.progress-panel,.metrics-panel,.support-panel{background:#fff;border:1px solid var(--line);margin-bottom:18px}.panel>summary,.status-group>summary,.support-panel>summary{list-style:none;cursor:pointer}.panel>summary::-webkit-details-marker,.status-group>summary::-webkit-details-marker,.support-panel>summary::-webkit-details-marker,.req-card>summary::-webkit-details-marker{display:none}.panel-head{padding:12px 15px;background:var(--pale2);display:flex;justify-content:space-between;align-items:center}.panel-head h2{margin:0;font-size:18px;color:var(--navy)}.panel-head span{font-size:11px;color:var(--muted)}.panel-head::after,.status-group>summary::after,.support-panel>summary::after,.req-summary::after{content:"+";color:var(--magenta);font-size:20px;margin-left:12px}.panel[open]>.panel-head::after,.status-group[open]>summary::after,.support-panel[open]>summary::after,.req-card[open]>.req-summary::after{content:"−"}.identity-columns{display:grid;grid-template-columns:1fr 1fr}.identity-column:first-child{border-right:1px solid var(--line)}.column-label{padding:8px 12px;color:var(--magenta);font-size:10px;font-weight:800;text-transform:uppercase;border-bottom:1px solid var(--line)}.id-row{display:grid;grid-template-columns:145px 1fr;gap:10px;padding:7px 12px;border-bottom:1px solid #E7EEF2}.id-row span{color:var(--muted);font-size:11px}.id-row strong{font-size:12px;color:var(--navy);overflow-wrap:anywhere}.docs{padding:12px 14px;border-top:1px solid var(--line)}.docs-title{font-size:10px;text-transform:uppercase;color:var(--magenta);font-weight:800;margin-bottom:8px}.radar-wrap{height:440px;padding:10px}.radar-svg{width:100%;height:100%;display:block}.empty-radar{padding:45px;color:var(--muted)}.progress-panel{padding:14px}.progress-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.progress-card{border-left:4px solid var(--blue);background:var(--pale2);padding:12px}.progress-card b{display:block;color:var(--navy);font-size:24px}.progress-card span{font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:800}.track{height:8px;background:#DDE8EE;margin-top:8px}.track i{display:block;height:100%;background:var(--magenta)}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--line)}.metric{background:#fff;padding:15px 18px;border-top:4px solid var(--blue)}.metric.alert{border-color:var(--magenta)}.metric b{display:block;font-size:28px;color:var(--navy)}.metric span{font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:800}.metric small{display:block;margin-top:5px;color:var(--slate)}.requirements-head{display:flex;gap:10px;align-items:center;margin:24px 0 12px}.requirements-head h2{margin:0;color:var(--navy);font-size:22px}.requirements-head input{margin-left:auto;width:320px;border:1px solid var(--line);padding:9px}.status-group>summary{padding:13px 15px;background:var(--pale2);border-left:5px solid var(--magenta);display:flex;align-items:center;gap:10px}.status-group>summary span{display:flex;flex-direction:column}.status-group>summary b{color:var(--navy);font-size:15px}.status-group>summary small{color:var(--muted);font-weight:400}.status-group>summary strong{margin-left:auto;color:var(--slate)}.group-tools{padding:8px 12px;border-top:1px solid var(--line);border-bottom:1px solid var(--line);display:flex;gap:7px}.group-tools button{background:#fff;border:1px solid var(--line);color:var(--navy);padding:6px 9px;cursor:pointer;font-weight:700}.req-card{border-bottom:1px solid var(--line)}.req-summary{list-style:none;cursor:pointer;display:grid;grid-template-columns:70px minmax(230px,1fr) 210px 210px 240px 25px;gap:10px;align-items:center;padding:12px 15px}.req-card[open]>.req-summary{background:var(--pale)}.req-code{color:var(--magenta);font-size:17px;font-weight:800}.req-title{font-weight:700;color:var(--navy)}.status-badge,.capacity-badge{font-size:11px;font-weight:700}.status-badge.confirmed{color:var(--green)}.status-badge.clarify{color:var(--amber)}.status-badge.undemonstrated{color:var(--slate)}.status-badge.excluded{color:var(--muted)}.capacity-badge{color:var(--blue)}.decision-text{font-size:12px;color:var(--slate)}.req-body{padding:0 15px 18px}.req-meta-strip{display:grid;grid-template-columns:repeat(3,1fr);background:var(--pale2);border:1px solid var(--line);margin:0 0 12px}.req-meta-strip div{padding:9px 11px;border-right:1px solid var(--line)}.req-meta-strip div:last-child{border-right:0}.req-meta-strip span{display:block;font-size:9px;text-transform:uppercase;color:var(--muted);font-weight:800}.req-meta-strip b{font-size:12px;color:var(--navy)}.detail-grid{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--line)}.detail-grid section{padding:14px;border-bottom:1px solid var(--line)}.detail-grid section:nth-child(odd){border-right:1px solid var(--line)}.detail-grid h3,.proofs h3{margin:0 0 7px;color:var(--magenta);font-size:11px;text-transform:uppercase;letter-spacing:.04em}.detail-grid p{margin:0 0 7px}.response-label{display:block;color:var(--blue);font-size:10px;font-weight:800;text-transform:uppercase;margin-bottom:5px}.detail-grid blockquote,.proof-card blockquote{margin:0;border-left:3px solid var(--magenta);padding-left:10px;color:#314A59}.gap-box,.question-box,.warning-box{background:var(--pale2);border-left:3px solid var(--blue);padding:9px;margin-top:9px}.gap-box b,.question-box b,.warning-box b{font-size:10px;text-transform:uppercase;color:var(--magenta)}.action-list{margin:0;padding-left:18px}.action-list li{margin-bottom:10px}.action-list li small{display:block;margin-top:3px;color:var(--muted);font-size:10px;font-style:italic}.proofs{margin-top:12px;border:1px solid var(--line);padding:14px}.proof-card{border-top:1px solid var(--line);padding:12px 0}.proof-card:first-of-type{border-top:0}.proof-meta{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}.proof-meta span,.proof-meta b{font-size:10px;padding:3px 6px;background:var(--pale2);color:var(--slate)}.proof-meta b{color:var(--navy)}.source-button{margin-top:8px;border:1px solid var(--blue);background:#fff;color:var(--blue);padding:6px 9px;cursor:pointer;font-weight:700}.no-proof{background:#FFF4E5;color:#7A4C00;padding:10px;border-left:3px solid var(--amber)}.support-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}.support-grid.single{grid-template-columns:1fr}.support-panel>summary{padding:12px 15px;background:var(--pale2);font-weight:800;color:var(--navy);display:flex;align-items:center}.support-body{padding:14px}.history-grid{display:grid;grid-template-columns:repeat(2,1fr);border:1px solid var(--line)}.history-grid div{padding:9px;border-bottom:1px solid var(--line)}.history-grid div:nth-child(odd){border-right:1px solid var(--line)}.history-grid span{display:block;font-size:9px;text-transform:uppercase;color:var(--muted)}.history-grid b{font-size:12px;color:var(--navy)}.gloss-table{border:1px solid var(--line);max-height:520px;overflow:auto}.gloss-row{display:grid;grid-template-columns:115px 1fr;border-bottom:1px solid var(--line)}.gloss-row strong,.gloss-row span{padding:8px}.gloss-row strong{background:var(--pale2);color:var(--blue)}.gloss-row span{font-size:12px;color:var(--slate)}.excluded-note{padding:14px;margin-top:12px}.excluded-note h3{margin:0 0 7px;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.04em}.excluded-check{background:var(--pale2);border-left:3px solid var(--muted);padding:8px 9px;margin-top:9px;color:var(--slate);font-size:13px}.tip-w{position:relative;display:inline-block;outline:none}.tip{border-bottom:1px dotted var(--magenta);cursor:help}.tip-b{display:none;position:absolute;z-index:30;left:0;bottom:calc(100% + 7px);width:300px;background:var(--navy);color:#fff;padding:10px 11px;font-size:12px;line-height:1.4;box-shadow:0 6px 18px rgba(0,0,0,.22)}.tip-w:hover .tip-b,.tip-w:focus .tip-b{display:block}.tip-src{margin-top:6px;padding-top:5px;border-top:1px solid rgba(255,255,255,.25);font-size:10px;color:var(--magenta-pale)}.modal{display:none;position:fixed;inset:0;background:rgba(7,28,36,.84);z-index:50;padding:24px}.modal.open{display:block}.modal-box{background:#fff;height:100%;display:grid;grid-template-rows:46px 1fr}.modal-head{display:flex;align-items:center;padding:0 13px;border-bottom:1px solid var(--line)}.modal-head button{margin-left:auto;background:var(--magenta);color:#fff;border:0;padding:7px 10px}.modal iframe{width:100%;height:100%;border:0}
@media(max-width:1180px){.overview-grid,.support-grid{grid-template-columns:1fr}.req-summary{grid-template-columns:60px 1fr 190px}.capacity-badge,.decision-text{grid-column:2}.req-summary::after{grid-column:3;grid-row:1/3}.radar-wrap{height:400px}}@media(max-width:780px){.hero{flex-wrap:wrap}.logo{width:100%;border-right:0;border-bottom:1px solid rgba(240,106,154,.45);padding-bottom:10px}.hero h1{font-size:29px}.hero nav{margin-left:0}.identity-columns,.progress-grid,.metrics,.detail-grid,.req-meta-strip{grid-template-columns:1fr}.identity-column:first-child{border-right:0}.detail-grid section:nth-child(odd),.req-meta-strip div{border-right:0}.req-summary{grid-template-columns:55px 1fr}.status-badge,.capacity-badge,.decision-text{grid-column:2}.req-summary::after{grid-column:1;grid-row:2/5}.requirements-head{align-items:stretch;flex-direction:column}.requirements-head input{width:100%;margin-left:0}.history-grid{grid-template-columns:1fr}.history-grid div:nth-child(odd){border-right:0}}
@media print{.hero nav,.group-tools,.source-button{display:none!important}.status-group,.req-card,.panel,.support-panel{break-inside:avoid}.container{max-width:none;padding:8px}.hero{padding:12px}.hero h1{font-size:24px}.req-summary{grid-template-columns:55px 1fr 170px 170px 200px 20px}.tip-b{display:none!important}}
'''

    script = f'''
const PDF_CONV={json.dumps(conv_b64)};const PDF_CCTP={json.dumps(cctp_b64)};let blobUrl=null;
function b64Blob(b){{const s=atob(b),u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return new Blob([u],{{type:'application/pdf'}});}}
function openPdf(kind,page){{const b=kind==='conv'?PDF_CONV:PDF_CCTP;if(!b)return alert('PDF source non intégré.');if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(b64Blob(b));document.getElementById('pdfFrame').src=blobUrl+'#page='+page;document.getElementById('modal').classList.add('open');}}
document.querySelectorAll('[data-pdf]').forEach(b=>b.addEventListener('click',()=>openPdf(b.dataset.pdf,Number(b.dataset.page))));
document.getElementById('closeModal').addEventListener('click',()=>{{document.getElementById('modal').classList.remove('open');document.getElementById('pdfFrame').src='';}});
document.querySelectorAll('[data-open-group]').forEach(b=>b.addEventListener('click',()=>b.closest('.status-group').querySelectorAll('.req-card').forEach(x=>x.open=true)));
document.querySelectorAll('[data-close-group]').forEach(b=>b.addEventListener('click',()=>b.closest('.status-group').querySelectorAll('.req-card').forEach(x=>x.open=false)));
document.getElementById('blockSearch').addEventListener('input',e=>{{const q=e.target.value.toLowerCase();document.querySelectorAll('.req-card').forEach(card=>{{card.hidden=!card.innerText.toLowerCase().includes(q)}});document.querySelectorAll('.status-group').forEach(group=>{{group.hidden=![...group.querySelectorAll('.req-card')].some(card=>!card.hidden)}});}});
'''

    return f'''<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Analyse d'appel d'offres BIM - {esc(meta.get('projet',''))}</title><style>{css}</style></head><body>
<header class="hero"><div class="logo">{f'<img src="{esc(meta.get("logo_data_uri"))}" alt="Logo">' if meta.get('logo_data_uri') else esc(meta.get('marque') or meta.get('entreprise') or 'Entreprise')}</div><div><h1>ANALYSE D’APPEL D’OFFRES BIM</h1><p>Lecture contractuelle, auto-évaluation des capacités et préparation d’une réponse maîtrisée pour le lot {esc(meta.get('lot',''))}.</p></div><nav><a href="#identity">Identité du projet</a><a class="primary" href="#requirements">Exigences et preuves</a></nav></header>
<main class="container">
<section id="identity" class="overview-grid">
<details class="panel" open><summary class="panel-head"><h2>Carte d’identité du projet</h2><span>Informations détectées et informations saisies</span></summary><div class="identity-columns"><div class="identity-column"><div class="column-label">Dossier et acteurs</div>{identity_rows(project_rows)}</div><div class="identity-column"><div class="column-label">Paramètres BIM utiles</div>{identity_rows(technical_rows)}</div></div><div class="docs"><div class="docs-title">Documents analysés</div>{_render_docs_analyses(meta, tip=tip)}</div></details>
<details class="panel" open><summary class="panel-head"><h2>Diagramme radar</h2><span>Capacités déclarées, tous blocs évalués</span></summary><div class="radar-wrap">{radar_svg}</div></details>
</section>
<section class="metrics-panel"><div class="metrics"><div class="metric"><b>{confirmed}</b><span>Obligations confirmées</span><small>Concernent directement le lot.</small></div><div class="metric alert"><b>{clarify}</b><span>À confirmer</span><small>Périmètre ou contenu à sécuriser.</small></div><div class="metric"><b>{undem}</b><span>Non démontrées</span><small>Aucune preuve ciblée recevable.</small></div><div class="metric"><b>{excluded}</b><span>Exclues / non applicables</span><small>Aucun engagement à prendre.</small></div></div></section>
<section id="requirements"><div class="requirements-head"><h2>Exigences classées par statut</h2><input id="blockSearch" placeholder="Rechercher un bloc, une preuve ou une action"></div>{''.join(groups_html)}</section>
<section class="support-grid single"><details class="support-panel"><summary>Glossaire utile</summary><div class="support-body"><div class="gloss-table">{glossary_html}</div><p style="color:var(--muted);font-size:11px">Les preuves et leur qualification sont disponibles directement dans chaque bloc d’exigence.</p></div></details></section>
</main><div class="modal" id="modal"><div class="modal-box"><div class="modal-head"><b>Document source</b><button type="button" id="closeModal">Fermer</button></div><iframe id="pdfFrame"></iframe></div></div><script>{script}</script></body></html>'''
