#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outil opérationnel BIM -- V1
Maturité numérique des TPE/PME du BTP face aux appels d'offres BIM
MSBIM 2025-2026 -- N. NDIAYE

3 modules en un seul script :
  1. Lecture des exigences BIM (convention + CCTP)
  2. Auto-évaluation de la maturité numérique (questionnaire interactif)
  3. Aide à la réponse à l'appel d'offres (paragraphes prêts à copier)

Usage :
  python outil_bim_v1.py
    --param   Parametrage_Outil_BIM_V1.xlsx
    --convention  Convention_BIM.pdf
    --cctp        CCTP_Lot_Facade.pdf
    --lot         "Façade"
    --entreprise  "Nom de l'entreprise"
    --out         Rapport_BIM_V1.pdf
"""
from __future__ import annotations
import argparse, base64, html, json, re, sys, textwrap, unicodedata, time
from functools import lru_cache
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import fitz
from datetime import datetime
import base64
try:
    from dashboard_template import build_dashboard
    _HAS_DASHBOARD = True
except ImportError:
    _HAS_DASHBOARD = False

from bim_model import block_sort_key, ensure_public_model, capacity_code_for_pct
from content_rules import glossary_entries_from_dataframe, normalize_questions_dataframe
from autoeval_report import generate_autoevaluation_reports

from reporting_v2 import (
    generate_action_plan_html,
    generate_action_plan_pdf,
    generate_glossary_html,
    generate_glossary_pdf,
)

# ═══════════════════════════════════════════════════════════════════════════════
# PALETTE
# ═══════════════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES TEXTE
# ═══════════════════════════════════════════════════════════════════════════════
def clean(s: object) -> str:
    return re.sub(r"\s+", " ",
                  html.unescape(str(s or "")).replace("\xa0", " ")).strip()

@lru_cache(maxsize=16384)
def _norm_cached_string(raw: str) -> str:
    """Normalisation Unicode mise en cache.

    Les mêmes textes de Convention/CCTP et les mêmes segments sont interrogés
    de très nombreuses fois pendant l'analyse. Le cache évite de refaire à
    chaque mot-clé le nettoyage et la décomposition Unicode du document entier.
    Cette optimisation ne modifie aucune règle métier ni aucun seuil.
    """
    cleaned = re.sub(r"\s+", " ", html.unescape(raw).replace("\xa0", " ")).strip()
    # Canonicalisation typographique générale avant toute comparaison métier.
    # Les PDF mélangent fréquemment apostrophes/dashes Unicode et ASCII ; ces
    # variantes ne doivent jamais nécessiter une règle propre à un DCE.
    translation = str.maketrans({
        "’": "'", "‘": "'", "´": "'", "`": "'", "ʼ": "'",
        "‐": "-", "‑": "-", "‒": "-", "–": "-", "—": "-", "−": "-",
        "œ": "oe", "Œ": "OE",
    })
    cleaned = cleaned.translate(translation)
    n = "".join(c for c in unicodedata.normalize("NFD", cleaned)
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", n.lower()).strip()

def norm(s: object) -> str:
    raw = s if isinstance(s, str) else str(s or "")
    return _norm_cached_string(raw)

def split_kw(s: object) -> List[str]:
    return [clean(x) for x in re.split(r"[;|,]", str(s or "")) if clean(x)]

def contains(text: str, kw: str) -> bool:
    t, m = norm(text), norm(kw)
    if not m: return False
    short_max = p_int(ENGINE_PARAMS_GLOBAL, "texte_mot_court_longueur_max", 0)
    if short_max and len(m) <= short_max:
        return re.search(r"(?<![a-z0-9])" + re.escape(m) + r"(?![a-z0-9])", t) is not None
    return m in t

def trouve_position(text_lower: str, kw: str, depart: int = 0) -> int:
    """
    Comme contains(), mais retourne la position du match (-1 si absent) dans
    `text_lower` (déjà en minuscules, NON normalisé -- pour que la position
    reste directement utilisable sur le texte original). Les mots courts
    (<=4 caractères, ex: 'acc', 'ecc') sont protégés par des limites de mot
    pour ne jamais matcher à l'intérieur d'un autre mot (ex: 'acc' dans
    'accord'). `depart` permet de reprendre la recherche après une position
    donnée (pour ignorer une occurrence déjà vue, ex: dans un sommaire).
    """
    kw_low = kw.lower()
    if not kw_low:
        return -1
    short_max = p_int(ENGINE_PARAMS_GLOBAL, "texte_mot_court_longueur_max", 0)
    if short_max and len(kw_low) <= short_max:
        m = re.search(r"(?<![a-zà-ÿ0-9])" + re.escape(kw_low) + r"(?![a-zà-ÿ0-9])", text_lower[depart:])
        return (m.start() + depart) if m else -1
    return text_lower.find(kw_low, depart)

def any_kw(text: str, kws: List[str]) -> bool:
    return any(contains(text, k) for k in kws)

# ═══════════════════════════════════════════════════════════════════════════════
# FILTRES TEXTUELS PARAMÉTRÉS
# ═══════════════════════════════════════════════════════════════════════════════
# Les motifs métier sont chargés depuis 15_Filtres_Texte. Les listes ci-dessous
# restent vides tant que le classeur de paramétrage n'a pas été chargé.
TEXT_FILTERS: Dict[str, List[Tuple[str, str]]] = {}
ENGINE_PARAMS_GLOBAL: Dict[str, object] = {}

def configure_text_filters(filters_df: pd.DataFrame, params: Dict[str, object]) -> None:
    global TEXT_FILTERS, ENGINE_PARAMS_GLOBAL
    ENGINE_PARAMS_GLOBAL = dict(params or {})
    out: Dict[str, List[Tuple[str, str]]] = {}
    if filters_df is not None and not filters_df.empty:
        for _, row in filters_df.iterrows():
            if clean(row.get("actif", "Oui")).casefold() in {"non", "false", "0"}:
                continue
            typ = clean(row.get("type_filtre", ""))
            motif = str(row.get("motif", "") or "").strip()
            mode = clean(row.get("mode", "contient")).lower()
            if typ and motif:
                out.setdefault(typ, []).append((motif, mode))
    TEXT_FILTERS = out

def _filter_match(text: str, motif: str, mode: str) -> bool:
    if mode == "regex":
        try:
            return re.search(motif, text, flags=re.IGNORECASE) is not None
        except re.error:
            return False
    return contains(text, motif)

def _matches_filter_type(text: str, type_filtre: str) -> bool:
    return any(_filter_match(text, motif, mode) for motif, mode in TEXT_FILTERS.get(type_filtre, []))

def _iter_keyword_occurrences(text_norm: str, keyword_norm: str):
    """Retourne les occurrences d'un mot-clé dans un texte normalisé.

    Le moteur reste générique : les mots courts utilisent les mêmes limites de
    mot que le reste de l'analyse, afin d'éviter par exemple qu'un acronyme soit
    trouvé à l'intérieur d'un autre mot.
    """
    if not keyword_norm:
        return []
    short_max = p_int(ENGINE_PARAMS_GLOBAL, "texte_mot_court_longueur_max", 0)
    if short_max and len(keyword_norm) <= short_max:
        pat = re.compile(r"(?<![a-z0-9])" + re.escape(keyword_norm) + r"(?![a-z0-9])")
        return [(m.start(), m.end()) for m in pat.finditer(text_norm)]
    out = []
    start = 0
    while True:
        pos = text_norm.find(keyword_norm, start)
        if pos < 0:
            break
        out.append((pos, pos + len(keyword_norm)))
        start = pos + max(1, len(keyword_norm))
    return out

def _clause_span(text_norm: str, start: int, end: int) -> Tuple[int, int]:
    """Borne l'analyse de négation à la proposition/phrase locale.

    Les motifs de négation eux-mêmes restent entièrement paramétrés dans
    15_Filtres_Texte (type NEGATION_TEMPLATE). Le code ne connaît ni bloc, ni
    métier, ni formulation contractuelle particulière.
    """
    separators = ".;!?"
    left = 0
    for sep in separators:
        pos = text_norm.rfind(sep, 0, start)
        if pos >= left:
            left = pos + 1
    right_candidates = []
    for sep in separators:
        pos = text_norm.find(sep, end)
        if pos >= 0:
            right_candidates.append(pos)
    right = min(right_candidates) if right_candidates else len(text_norm)
    return left, right

def _occurrence_is_locally_negated(text_norm: str, keyword_norm: str, start: int, end: int) -> bool:
    """Teste si une occurrence est niée par une règle générique Excel.

    Chaque motif NEGATION_TEMPLATE contient le jeton {term}; le moteur le
    remplace par le mot-clé effectivement recherché. La négation est ainsi
    rattachée à l'occurrence qu'elle vise, sans coder de bloc, de métier ou de
    livrable particulier dans Python.
    """
    templates = TEXT_FILTERS.get("NEGATION_TEMPLATE", [])
    if not templates:
        return False
    left, right = _clause_span(text_norm, start, end)
    fragment = text_norm[left:right]
    rel_start = start - left
    escaped_term = re.escape(keyword_norm)
    for motif, mode in templates:
        pattern = str(motif or "").replace("{term}", escaped_term)
        if not pattern:
            continue
        if mode == "regex":
            try:
                for match in re.finditer(pattern, fragment, flags=re.IGNORECASE):
                    # Le motif doit couvrir l'occurrence précise du terme ; une
                    # négation visant un autre objet de la même phrase n'est donc
                    # pas suffisante pour annuler cette preuve.
                    if match.start() <= rel_start < match.end():
                        return True
            except re.error:
                continue
        else:
            # Mode contient conservé pour extensibilité du paramétrage, même si
            # les modèles de négation actuels sont des regex.
            literal = norm(pattern)
            if literal and literal in fragment:
                idx = fragment.find(literal)
                if idx <= rel_start < idx + len(literal):
                    return True
    return False

def keyword_has_non_negated_occurrence(text: str, keyword: str) -> bool:
    """Vrai si au moins une occurrence du mot-clé n'est pas niée localement."""
    t = norm(text)
    k = norm(keyword)
    if not t or not k:
        return False
    occurrences = _iter_keyword_occurrences(t, k)
    if not occurrences:
        return False
    return any(not _occurrence_is_locally_negated(t, k, start, end)
               for start, end in occurrences)

def any_kw_non_negated(text: str, kws: List[str]) -> bool:
    """Équivalent d'any_kw pour les signaux positifs contractuels."""
    return any(keyword_has_non_negated_occurrence(text, k) for k in kws if k)

def _matches_admin_pattern(line: str) -> bool:
    return _matches_filter_type(line, "ADMIN_REGEX")

def detect_running_headers(pages_raw: List[str]) -> set:
    from collections import Counter
    n_pages = len(pages_raw)
    min_pages = p_int(ENGINE_PARAMS_GLOBAL, "entete_repetee_pages_min", 0)
    if min_pages and n_pages < min_pages:
        return set()
    line_min = p_int(ENGINE_PARAMS_GLOBAL, "entete_repetee_ligne_longueur_min", 0)
    line_max = p_int(ENGINE_PARAMS_GLOBAL, "entete_repetee_ligne_longueur_max", 10**9)
    min_ratio = p_float(ENGINE_PARAMS_GLOBAL, "entete_repetee_ratio_min", 0.0)
    min_occ = p_int(ENGINE_PARAMS_GLOBAL, "entete_repetee_occurrences_min", 0)
    counts = Counter()
    for txt in pages_raw:
        seen = set()
        for line in txt.split("\n"):
            line = line.strip().lower()
            if line_min <= len(line) <= line_max:
                seen.add(line)
        for line in seen:
            counts[line] += 1
    threshold = max(min_occ, int(n_pages * min_ratio))
    return {line for line, count in counts.items() if count >= threshold}

def is_hors_bim(txt: str) -> bool:
    return _matches_filter_type(txt, "HORS_BIM")

def has_bim_signal(txt: str) -> bool:
    return _matches_filter_type(txt, "BIM_SIGNAL")

def clean_page_headers(txt: str, running_headers: set = frozenset()) -> str:
    lines = txt.split('\n')
    return '\n'.join(
        line for line in lines
        if line.strip().lower() not in running_headers
        and not _matches_admin_pattern(line)
    )

def clean_proof_text(txt: str) -> str:
    # Nettoyage générique : suppression des lignes reconnues comme administratives.
    return clean_page_headers(txt).strip()

# ═══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT PARAMÉTRAGE
# ═══════════════════════════════════════════════════════════════════════════════
def read_sheet(path: Path, sheet: str, key_col: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    hi = None
    for i, row in raw.iterrows():
        if key_col in [clean(v) for v in row.values]:
            hi = i; break
    if hi is None:
        raise SystemExit(f"Colonne '{key_col}' introuvable dans '{sheet}'")
    headers = [clean(v) for v in raw.iloc[hi].values]
    df = raw.iloc[hi + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")
    df = df.astype(object).where(pd.notna(df), "")
    df = df[[c for c in df.columns if c and c.lower() != "nan"]]
    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.strip().str.lower().ne("non")]
    return df


def load_all(path: Path):
    blocs = read_sheet(path, "10_Blocs", "id_bloc")
    gloss = read_sheet(path, "40_Glossaire", "terme")
    questions = read_sheet(path, "50_Questions", "id_bloc")
    textes = read_sheet(path, "60_Reponses", "id_bloc")

    def _try(sheet, key):
        try:
            return read_sheet(path, sheet, key)
        except SystemExit:
            return pd.DataFrame()

    signaux = _try("30_CCTP_Signaux", "categorie")
    detections = _try("35_Detections_Convention", "cle_ctx")
    metadata_rules = _try("36_Extraction_Metadata", "champ")
    axes_config = _try("16_Axes_Config", "axe")
    coherence = _try("55_Regles_Coherence", "id_regle")
    messages_df = _try("12_Messages_Moteur", "cle")
    params_df = _try("05_Parametres_Moteur", "cle")
    text_filters = _try("15_Filtres_Texte", "id_filtre")
    restitution_rules = _try("17_Regles_Restitution", "id_regle")
    ccap_rules = _try("37_CCAP_Reperage", "id_regle")
    interblock_rules = _try("18_Regles_Interblocs", "id_regle")
    document_rules = _try("19_Regles_Documentaires", "id_regle")
    return (
        blocs, gloss, questions, textes, signaux, detections, metadata_rules,
        axes_config, coherence, messages_df, params_df, text_filters,
        restitution_rules, ccap_rules, interblock_rules, document_rules,
    )


def load_rule_rows(df: pd.DataFrame) -> List[Dict[str, object]]:
    if df is None or df.empty:
        return []
    return [{clean(k): row.get(k, "") for k in df.columns if clean(k)} for _, row in df.iterrows()]

def load_params(params_df: pd.DataFrame) -> Dict[str, object]:
    """
    Lit 05_Parametres_Moteur -> {cle: valeur} avec conversion selon la colonne
    'type' (entier/liste/texte). Aucun seuil/poids codé en dur dans le script.
    """
    out: Dict[str, object] = {}
    if params_df is None or params_df.empty:
        return out
    for _, row in params_df.iterrows():
        cle = clean(row.get("cle", ""))
        if not cle:
            continue
        typ = clean(row.get("type", "texte")).lower()
        val = row.get("valeur", "")
        if typ == "entier":
            try: out[cle] = int(float(val))
            except (TypeError, ValueError): out[cle] = 0
        elif typ in {"decimal", "reel", "float"}:
            try: out[cle] = float(val)
            except (TypeError, ValueError): out[cle] = 0.0
        elif typ == "liste":
            out[cle] = split_kw(val)
        else:
            out[cle] = clean(val)
    return out


def load_messages(messages_df: pd.DataFrame) -> Dict[str, str]:
    """Lit 12_Messages_Moteur -> {cle: texte}. Aucun message codé en dur."""
    if messages_df is None or messages_df.empty:
        return {}
    return {clean(r["cle"]): str(r.get("texte", "")) for _, r in messages_df.iterrows() if clean(r.get("cle", ""))}


def render_message(messages: Dict[str, str], cle: str, **kwargs) -> str:
    """Substitue {titre}/{lot}/{source}... dans un message paramétré."""
    tpl = messages.get(cle, "")
    try:
        return tpl.format(**kwargs)
    except (KeyError, IndexError):
        return tpl


def load_axes_config(axes_df: pd.DataFrame) -> Dict[str, Dict[str, Dict]]:
    """Lit 16_Axes_Config -> {axe: {code: {label, couleur, symbole, ordre}}}."""
    out: Dict[str, Dict[str, Dict]] = {}
    if axes_df is None or axes_df.empty:
        return out
    for _, row in axes_df.iterrows():
        axe = clean(row.get("axe", ""))
        code = clean(row.get("code", ""))
        if not axe or not code:
            continue
        cfg = {
            "label": clean(row.get("label", "")) or code,
            "couleur": clean(row.get("couleur", "")),
            "symbole": clean(row.get("symbole", "")),
            "ordre": int(float(row.get("ordre", 9999) or 9999)),
            "description": clean(row.get("description", "")),
            "label_court": clean(row.get("label_court", "")),
            "note_restitution": clean(row.get("note_restitution", "")),
            "pct_min": None,
            "pct_max": None,
        }
        for key in ("pct_min", "pct_max"):
            value = row.get(key, "")
            if value not in (None, ""):
                try:
                    cfg[key] = float(value)
                except (TypeError, ValueError):
                    cfg[key] = None
        out.setdefault(axe, {})[code] = cfg
    return out


# NOTE V11 : la notion de "profil de lot" (producteur_maquette / documentaire)
# a été supprimée -- elle provoquait des exclusions automatiques de blocs qui
# écrasaient des preuves contractuelles réelles (ex. B04 exclu à tort alors
# que le CCTP décrivait un coordinateur BIM dédié au lot). L'applicabilité
# d'un bloc repose désormais uniquement sur les preuves textuelles trouvées
# dans la convention et le CCTP.


# ═══════════════════════════════════════════════════════════════════════════════
# ENRICHISSEMENT GLOSSAIRE -- extrait du document réel
# ═══════════════════════════════════════════════════════════════════════════════
def filtrer_glossaire_detecte(gloss: pd.DataFrame, texte_combine: str) -> pd.DataFrame:
    """Ne garde, dans le glossaire Excel, que les lignes dont le terme ou l'un
    de ses mots-clés apparaît réellement dans le texte analysé (convention +
    CCTP + questionnaire d'auto-évaluation + textes générés par l'outil).
    Utilisé partout où un glossaire est présenté à l'entreprise -- panneau
    "Glossaire utile", annexe technique, infobulles -- afin de ne jamais
    afficher un terme (ex. le nom d'une plateforme CDE concurrente) qui n'a
    aucun lien avec le projet analysé ni avec ce que l'entreprise va lire."""
    if not isinstance(gloss, pd.DataFrame) or gloss.empty:
        return gloss
    texte_combine = texte_combine or ""
    def _detecte(row) -> bool:
        terme = clean(row.get("terme", ""))
        mots = split_kw(row.get("mots_cles", "")) or [terme]
        return bool(terme) and any(contains(texte_combine, m) for m in mots)
    return gloss[gloss.apply(_detecte, axis=1)]


# Colonnes de 10_Blocs contenant du texte librement rédigé, potentiellement
# affiché à l'entreprise (titre, lecture, actions, questions, propositions...).
_COLONNES_TEXTE_BLOCS = [
    "titre_bloc", "description_convention", "lecture_entreprise",
    "action_confirmee", "action_probable", "action_non_applicable",
    "risque_si_ignore", "interpretation_croisee", "question_bim_manager",
    "proposition_commerciale", "action_interne_standard", "action_interne_capacite",
]
# Colonnes de 50_Questions contenant le texte du questionnaire d'auto-évaluation.
_COLONNES_TEXTE_QUESTIONS = ["question", "indice_0", "indice_1", "indice_2"]


def texte_outil_et_questionnaire(blocs: pd.DataFrame = None, questions: pd.DataFrame = None) -> str:
    """Concatène tous les textes que l'outil affiche lui-même (10_Blocs) et
    tout le texte du questionnaire d'auto-évaluation (50_Questions), pour que
    le glossaire couvre aussi ce vocabulaire -- pas seulement les mots
    trouvés mot pour mot dans les PDF d'entrée."""
    morceaux: List[str] = []
    if isinstance(blocs, pd.DataFrame):
        for col in _COLONNES_TEXTE_BLOCS:
            if col in blocs.columns:
                morceaux.extend(str(v) for v in blocs[col].fillna("").tolist())
    if isinstance(questions, pd.DataFrame):
        for col in _COLONNES_TEXTE_QUESTIONS:
            if col in questions.columns:
                morceaux.extend(str(v) for v in questions[col].fillna("").tolist())
    return " ".join(morceaux)


def enrichir_glossaire(gloss: pd.DataFrame, conv_pages: List[Tuple[int,str]],
                       cctp_pages: List[Tuple[int,str]]) -> List[Dict]:
    """
    Pour chaque terme du glossaire Excel, cherche une phrase de définition
    dans les documents analysés (convention prioritaire, puis CCTP).
    Retourne une liste de dicts enrichis.
    Si aucun extrait trouvé → definition_doc vide, on affiche la définition Excel.
    """
    all_pages = [(pg, txt, "Convention BIM") for pg, txt in conv_pages] + \
                [(pg, txt, "CCTP du lot")    for pg, txt in cctp_pages]

    result = []
    for _, row in gloss.iterrows():
        terme     = clean(row.get("terme", ""))
        def_excel = clean(row.get("definition", ""))
        mots      = split_kw(row.get("mots_cles", "")) or [terme]

        best_phrase, best_page, best_source, best_score = "", "", "", 0

        for pg, txt, src in all_pages:
            phrases = re.split(r'(?<=[.!?])\s+', txt.replace("\n", " "))
            for phrase in phrases:
                phrase = clean(phrase)
                phrase_min = p_int(ENGINE_PARAMS_GLOBAL, "gloss_phrase_longueur_min", 0)
                phrase_max = p_int(ENGINE_PARAMS_GLOBAL, "gloss_phrase_longueur_max", 10**9)
                if len(phrase) < phrase_min or len(phrase) > phrase_max:
                    continue
                sc = sum(1 for m in mots if contains(phrase, m))
                if sc == 0:
                    continue  # le mot-clé doit être présent -- le bonus seul ne doit jamais suffire
                if _matches_filter_type(phrase, "DEFINITION_SIGNAL"):
                    sc += 1
                if sc > best_score:
                    best_score  = sc
                    # Centrer l'extrait sur le mot-clé plutôt que tronquer
                    # depuis le début -- un mot-clé apparaissant tard dans la
                    # phrase ne doit jamais être coupé par la troncature à 250c.
                    pos_kw = next((norm(phrase).find(norm(m)) for m in mots
                                   if norm(m) and norm(m) in norm(phrase)), 0)
                    centre_threshold = p_int(ENGINE_PARAMS_GLOBAL, "gloss_extrait_centrage_seuil", 0)
                    backtrack = p_int(ENGINE_PARAMS_GLOBAL, "gloss_extrait_recul", 0)
                    excerpt_max = p_int(ENGINE_PARAMS_GLOBAL, "gloss_extrait_longueur_max", len(phrase))
                    if centre_threshold and pos_kw > centre_threshold:
                        deb = max(0, pos_kw - backtrack)
                        best_phrase = phrase[deb:deb + excerpt_max]
                    else:
                        best_phrase = phrase[:excerpt_max]
                    best_page   = str(pg)
                    best_source = src

        # Repli générique : si aucune "phrase-définition" de bonne qualité
        # n'a été trouvée (score < 2 -- notamment systématique pour un terme
        # n'ayant qu'un seul mot-clé, ou une phrase de plus de 300 caractères,
        # fréquente dans les conventions BIM), chercher simplement une fenêtre
        # de texte autour de la première occurrence du mot-clé. Une citation
        # imparfaite reste bien plus utile qu'aucune citation du tout.
        definition_score_min = p_int(ENGINE_PARAMS_GLOBAL, "gloss_definition_score_min", 0)
        if best_score < definition_score_min:
            best_phrase, best_page, best_source = "", "", ""
            for pg, txt, src in all_pages:
                low = txt.lower()
                for m in mots:
                    i = trouve_position(low, m)
                    if i == -1:
                        continue
                    fallback_context = p_int(ENGINE_PARAMS_GLOBAL, "gloss_repli_contexte", 0)
                    start = max(0, i - fallback_context)
                    extrait = clean(txt[start:i + len(m) + fallback_context])
                    if re.search(r"\.{4,}", extrait):  # ignorer les lignes de sommaire
                        continue
                    excerpt_max = p_int(ENGINE_PARAMS_GLOBAL, "gloss_extrait_longueur_max", len(extrait))
                    best_phrase, best_page, best_source = extrait[:excerpt_max], str(pg), src
                    break
                if best_page:
                    break

        result.append({
            "terme":            terme,
            "definition_excel": def_excel,
            "definition_doc":   best_phrase if (best_score >= definition_score_min or best_page) else "",
            "page_doc":         best_page,
            "source_doc":       best_source,
            "source_excel":     clean(row.get("source", "")),
            "explication_tpe_pme": clean(row.get("explication_tpe_pme", "")),
            "categorie":        clean(row.get("categorie", "")),
            "niveau":           clean(row.get("niveau", "")),
        })
    return result


def glossary_entries_detected_documents(gloss: pd.DataFrame, conv_pages: List[Tuple[int, str]],
                                         cctp_pages: List[Tuple[int, str]]) -> List[Dict]:
    """Construit le glossaire autonome uniquement depuis les termes réellement
    repérés dans les documents métier analysés (Convention BIM + CCTP).

    La définition reste celle du référentiel Excel ; la source et la page sont
    celles de l'occurrence documentaire retenue par ``enrichir_glossaire``.
    Le CCAP n'est volontairement pas pris en compte ici : il conserve son rôle
    de hiérarchie documentaire et de vigilance, sans alimenter le glossaire BIM.
    """
    enriched = enrichir_glossaire(gloss, conv_pages or [], cctp_pages or [])
    out: List[Dict] = []
    for item in enriched:
        term = clean(item.get("terme", ""))
        definition = clean(item.get("definition_excel", ""))
        page = clean(item.get("page_doc", ""))
        source_doc = clean(item.get("source_doc", ""))
        if not term or not definition or not page or not source_doc:
            continue
        out.append({
            "term": term,
            "definition": definition,
            "practice": clean(item.get("explication_tpe_pme", "")),
            # Compatibilite historique : source reste l'occurrence du DCE.
            "source": f"{source_doc} p.{page}",
            "detected_in": f"{source_doc} p.{page}",
            "reference": clean(item.get("source_excel", "")),
            "source_doc": source_doc,
            "page": page,
            "category": clean(item.get("categorie", "")),
        })
    return sorted(out, key=lambda item: item["term"].casefold())


# ═══════════════════════════════════════════════════════════════════════════════
# LECTURE PDF
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass
class DocData:
    name: str
    pages: List[Tuple[int, str]]
    text: str = field(default="", repr=False)
    # Informations purement techniques sur la composition visuelle des pages.
    # Elles servent uniquement à repérer une page image sans couche texte
    # suffisante ; aucune règle métier BIM n'en dépend.
    page_visual_info: Dict[int, Dict[str, float]] = field(default_factory=dict, repr=False)
    _segments_cache: Optional[List[Tuple[int, str]]] = field(default=None, repr=False)

def _reparer_texte_affichage(value: str) -> str:
    """Répare, uniquement pour la restitution, quelques pertes Unicode typiques
    de l'extraction PDF sans modifier le texte utilisé par le moteur d'analyse.

    Les remplacements sont volontairement contextuels afin de conserver les vrais
    points d'interrogation et les paramètres d'URL.
    """
    txt = str(value or "")
    txt = re.sub(r"\b([dD])\?\?uvre\b", lambda m: m.group(1) + "’œuvre", txt)
    prefixes = r"(?:[cdjlmnst]|qu|jusqu|lorsqu|puisqu|quoiqu|quelqu|aujourd)"
    txt = re.sub(
        rf"\b({prefixes})\?(?=[A-Za-zÀ-ÖØ-öø-ÿ])",
        lambda m: m.group(1) + "’",
        txt,
        flags=re.IGNORECASE,
    )
    return txt

def _reparer_obj_affichage(value):
    """Applique la réparation d'affichage récursivement sans toucher aux objets métier."""
    if isinstance(value, str):
        return _reparer_texte_affichage(value)
    if isinstance(value, dict):
        return {k: _reparer_obj_affichage(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_reparer_obj_affichage(v) for v in value]
    if isinstance(value, tuple):
        return tuple(_reparer_obj_affichage(v) for v in value)
    return value

def lire_pdf(path: Path) -> DocData:
    doc = fitz.open(str(path))
    raw_pages = []
    page_visual_info: Dict[int, Dict[str, float]] = {}
    for page_no, pg in enumerate(doc, start=1):
        txt = pg.get_text("text").replace("\xa0", " ")
        txt = re.sub(r"[ \t]+", " ", txt)
        raw_pages.append(txt)

        # Une photographie / numérisation pleine page est généralement un seul
        # grand objet image. On mesure uniquement la couverture géométrique de
        # la plus grande image de la page. Cette information permet de distinguer
        # une vraie page image d'une page blanche, d'un intertitre ou d'une page
        # contenant simplement un petit logo.
        page_area = max(float(pg.rect.width * pg.rect.height), 1.0)
        max_image_ratio = 0.0
        image_count = 0
        try:
            image_infos = pg.get_image_info()
        except Exception:
            image_infos = []
        for info in image_infos or []:
            bbox = info.get("bbox")
            if not bbox:
                continue
            try:
                rect = fitz.Rect(bbox)
                area = max(0.0, float(rect.width * rect.height))
            except Exception:
                continue
            image_count += 1
            max_image_ratio = max(max_image_ratio, min(1.0, area / page_area))
        page_visual_info[page_no] = {
            "max_image_ratio": max_image_ratio,
            "image_count": float(image_count),
        }

    # Détection générique des en-têtes/pieds de page répétés sur ce document précis
    running_headers = detect_running_headers(raw_pages)
    pages = []
    for i, txt in enumerate(raw_pages, start=1):
        txt = clean_page_headers(txt, running_headers)
        txt = re.sub(r"\n{3,}", "\n\n", txt).strip()
        pages.append((i, txt))
    doc.close()
    return DocData(name=path.name, pages=pages,
                   text=" ".join(t for _, t in pages),
                   page_visual_info=page_visual_info)


def evaluer_lisibilite_document(doc: DocData, params: Optional[Dict[str, object]] = None) -> Dict[str, object]:
    """Évalue si un PDF contient une couche texte réellement exploitable.

    Le contrôle ne réalise pas d'OCR : il mesure le texte extrait par PyMuPDF
    et, page par page, la présence éventuelle d'une grande image couvrant la
    page. Les seuils sont pilotés par 05_Parametres_Moteur.

    États :
      - OK : texte suffisamment présent et aucune page image non lisible repérée ;
      - PARTIEL : au moins une page vraisemblablement scannée/image sans couche
        texte suffisante, ou une proportion importante de pages sans texte ;
      - NON_EXPLOITABLE : texte global trop faible pour analyser le document sans
        risque de conclure à tort à l'absence d'exigence.

    Une page courte ou blanche n'est pas assimilée à une page scannée uniquement
    parce qu'elle contient peu de texte : il faut également qu'une image couvre
    une part importante de la page.
    """
    params = params or ENGINE_PARAMS_GLOBAL or {}
    page_min = max(1, p_int(params, "pdf_lecture_page_chars_min", 20))
    total_min = max(1, p_int(params, "pdf_lecture_total_chars_min", 80))
    ratio_partial = p_float(params, "pdf_lecture_ratio_pages_partiel", 0.5)
    ratio_pages_min = max(1, p_int(params, "pdf_lecture_pages_min_ratio", 3))
    image_ratio_min = p_float(params, "pdf_lecture_image_couverture_min", 0.65)
    image_ratio_min = min(1.0, max(0.05, image_ratio_min))

    counts: List[int] = []
    pages_images_non_lisibles: List[int] = []
    page_details: List[Dict[str, object]] = []
    for page_no, text in (doc.pages or []):
        # Les espaces, ponctuations et traits de mise en page ne doivent pas
        # suffire à faire croire qu'une couche texte exploitable existe.
        count = sum(1 for ch in str(text or "") if ch.isalnum())
        counts.append(count)
        visual = (doc.page_visual_info or {}).get(page_no, {})
        image_ratio = float(visual.get("max_image_ratio", 0.0) or 0.0)
        page_image_non_lisible = count < page_min and image_ratio >= image_ratio_min
        if page_image_non_lisible:
            pages_images_non_lisibles.append(int(page_no))
        page_details.append({
            "page": int(page_no),
            "caracteres": count,
            "max_image_ratio": image_ratio,
            "image_non_lisible": page_image_non_lisible,
        })

    pages_total = len(counts)
    pages_lisibles = sum(1 for count in counts if count >= page_min)
    chars_total = sum(counts)
    ratio = (pages_lisibles / pages_total) if pages_total else 0.0

    if pages_total == 0 or pages_lisibles == 0 or chars_total < total_min:
        statut = "NON_EXPLOITABLE"
    elif pages_images_non_lisibles:
        # Même une seule page image peut contenir une exigence importante.
        # On poursuit l'analyse du reste du document mais on expose précisément
        # l'angle mort au lieu de le masquer derrière un ratio global élevé.
        statut = "PARTIEL"
    elif pages_total >= ratio_pages_min and ratio < ratio_partial:
        statut = "PARTIEL"
    else:
        statut = "OK"

    return {
        "statut": statut,
        "pages_total": pages_total,
        "pages_lisibles": pages_lisibles,
        "ratio_pages_lisibles": ratio,
        "caracteres_extraits": chars_total,
        "pages_images_non_lisibles": pages_images_non_lisibles,
        "pages_images_non_lisibles_count": len(pages_images_non_lisibles),
        "page_details": page_details,
    }

def segments(doc: DocData) -> List[Tuple[int, str]]:
    """
    Découpe le texte en segments exploitables.
    Stratégie :
    1. Reconstituer les paragraphes (lignes séparées par une ligne vide)
    2. Découper sur les numéros de section (2.14, 3.1 etc.)
    3. Filtrer : hors-BIM, titres seuls, trop courts
    """
    if doc._segments_cache is not None:
        return doc._segments_cache
    out = []
    for page, raw in doc.pages:
        txt = re.sub(r"Page\s+\d+\s+sur\s+\d+", " ", raw, flags=re.I)
        # Découper sur : lignes vides, numéros de section, phrases complètes
        parts = re.split(
            r"\n{2,}"                         # ligne vide = nouveau §
            r"|(?=\n\d+\.\d+\s+[A-Z])"   # avant "2.14 TITRE"
            r"|(?=\n\d+\s+[A-Z]{3})"        # avant "3 DESCRIPTION"
            r"|(?<=[.!?])\s{1,3}(?=[A-Z])",   # fin de phrase → majuscule
            txt)
        for part in parts:
            part = clean(part)
            min_len = p_int(ENGINE_PARAMS_GLOBAL, "segment_longueur_min", 0)
            if len(part) < min_len: continue
            if is_hors_bim(part): continue
            # Exclure les lignes de sommaire avec pointillés
            dotted_ratio = p_float(ENGINE_PARAMS_GLOBAL, "segment_ratio_pointilles_max", 1.0)
            if len(part) and part.count('.') > len(part) * dotted_ratio: continue
            # Exclure les titres purs (court + tout en majuscules ou numéro seul)
            title_max = p_int(ENGINE_PARAMS_GLOBAL, "segment_titre_longueur_max", 0)
            if re.match(r"^\d+\.?\d*\s+[A-Z]{3}", part) and title_max and len(part) < title_max: continue
            segment_max = p_int(ENGINE_PARAMS_GLOBAL, "segment_longueur_max", len(part))
            if segment_max and len(part) > segment_max:
                part = part[:max(0, segment_max - len("..."))] + "..."
            out.append((page, part))
    doc._segments_cache = out
    return out

def score_seg(txt: str, kws: List[str], strong: List[str]) -> int:
    params = ENGINE_PARAMS_GLOBAL
    score = 0
    score += p_int(params, "score_poids_mot_cle", 0) * sum(1 for k in kws if keyword_has_non_negated_occurrence(txt, k))
    score += p_int(params, "score_poids_signal_fort", 0) * sum(1 for k in strong if keyword_has_non_negated_occurrence(txt, k))
    if has_bim_signal(txt):
        score += p_int(params, "score_bonus_signal_bim", 0)
    obligation_count = sum(1 for motif, mode in TEXT_FILTERS.get("OBLIGATION", []) if _filter_match(txt, motif, mode))
    score += p_int(params, "score_bonus_formulation_obligation", 0) * obligation_count
    if is_hors_bim(txt):
        score -= p_int(params, "score_penalite_hors_bim", 0)
    return score


def _evidence_matches(txt: str, kws: List[str], strong: List[str]) -> Tuple[List[str], List[str], List[str]]:
    matched = [k for k in kws if k and keyword_has_non_negated_occurrence(txt, k)]
    matched_strong = [k for k in strong if k and keyword_has_non_negated_occurrence(txt, k)]
    generic = {norm(motif) for motif, _ in TEXT_FILTERS.get("EVIDENCE_GENERIC", [])}
    min_len = p_int(ENGINE_PARAMS_GLOBAL, "preuve_longueur_min_terme_specifique", 0)
    acr_min = p_int(ENGINE_PARAMS_GLOBAL, "preuve_acronyme_longueur_min", 0)
    acr_max = p_int(ENGINE_PARAMS_GLOBAL, "preuve_acronyme_longueur_max", 9999)
    specific = []
    for k in matched:
        nk = norm(k)
        raw = str(k).strip()
        is_acronym = raw.isupper() and acr_min <= len(raw) <= acr_max
        is_specific = (" " in nk or len(nk) >= min_len or is_acronym) and nk not in generic
        if is_specific:
            specific.append(k)
    return matched, matched_strong, specific


def _sujets_lot_nommes(texte: object) -> List[str]:
    """Extrait uniquement les sujets explicitement rédigés sous la forme
    « le lot X + verbe ». Les destinataires (« au lot X ») ne sont pas traités
    comme des sujets. Les verbes viennent du classeur de paramétrage ; aucun
    corps d'état n'est connu du moteur.
    """
    # Travailler sur la représentation normalisée (accents, ligatures, tirets,
    # apostrophes et casse neutralisés) afin que la portée d'une clause ne
    # dépende jamais de la qualité typographique/OCR du PDF. Les acteurs
    # extraits servent uniquement à une comparaison normalisée de portée.
    text = norm(texte)
    if not text:
        return []
    raw_verbs = ENGINE_PARAMS_GLOBAL.get("cctp_acteur_lot_verbes", "")
    if isinstance(raw_verbs, (list, tuple, set)):
        verbs = [clean(v) for v in raw_verbs if clean(v)]
    else:
        verbs = split_kw(raw_verbs)
    if not verbs:
        verbs = ["doit", "devra", "réalise", "réalisera", "fournit", "fournira",
                 "produit", "produira", "remet", "remettra", "dépose", "déposera",
                 "développe", "développera", "modélise", "modélisera",
                 "transmet", "transmettra"]
    verb_pat = "|".join(sorted((re.escape(norm(v)) for v in verbs if clean(v)), key=len, reverse=True))
    if not verb_pat:
        return []
    patterns = [
        re.compile(r"(?i)\b(?:le|la)\s+lot\s+(?P<acteur>[^.;:\n]{1,110}?)\s+(?:" + verb_pat + r")\b"),
        re.compile(r"(?im)^\s*lot\s+(?P<acteur>[^.;:\n]{1,110}?)\s+(?:" + verb_pat + r")\b"),
    ]
    actors: List[str] = []
    for pattern in patterns:
        for match in pattern.finditer(text):
            actor = clean(match.group("acteur"))
            if actor and actor.casefold() not in {a.casefold() for a in actors}:
                actors.append(actor)
    return actors


def _clause_compatible_portee_lot(texte: object, lot: object) -> bool:
    """Filtre de sécurité transversal des preuves CCTP.

    - aucune désignation explicite de sujet « lot X » : la clause reste valable
      comme clause générale du CCTP ;
    - un ou plusieurs sujets « lot X » sont nommés : au moins un doit
      correspondre au lot analysé.

    Le mécanisme évite qu'une prescription visant explicitement un autre lot
    confirme l'exigence du lot courant, sans dictionnaire de corps d'état.
    """
    enabled = clean(ENGINE_PARAMS_GLOBAL.get("cctp_filtrer_acteur_lot_explicite", "Oui")).lower()
    if enabled in {"non", "no", "false", "0"}:
        return True
    actors = _sujets_lot_nommes(texte)
    if not actors:
        return True
    return any(_portee_lot_correspond(actor, lot) for actor in actors)


def _texte_document_compatible_lot(doc: Optional[DocData], lot: object) -> str:
    if not doc:
        return ""
    kept = [seg for _page, seg in segments(doc) if _clause_compatible_portee_lot(seg, lot)]
    return "\n".join(kept)


def best_evidence(doc: Optional[DocData], kws: List[str], strong: List[str],
                  used: set, lot: str = "", enforce_lot_scope: bool = False) -> Tuple[str, str, int, str]:
    if not doc:
        return "", "", -999, ""
    best_sc, best_pg, best_txt, best_key = -999, "", "", ""
    min_kw = p_int(ENGINE_PARAMS_GLOBAL, "preuve_min_mots_cles", 0)
    min_specific_no_bim = p_int(ENGINE_PARAMS_GLOBAL, "preuve_min_specificites_sans_signal_bim", 0)
    for page, seg in segments(doc):
        if enforce_lot_scope and lot and not _clause_compatible_portee_lot(seg, lot):
            continue
        dedup_len = p_int(ENGINE_PARAMS_GLOBAL, "preuve_cle_dedoublonnage_longueur", len(seg))
        key = norm(seg[:dedup_len])
        if key in used:
            continue
        matched, matched_strong, specific = _evidence_matches(seg, kws, strong)
        eligible = bool(matched_strong or len(set(map(norm, matched))) >= min_kw or specific)
        if not eligible:
            continue
        sc = score_seg(seg, kws, strong)
        sc += p_int(ENGINE_PARAMS_GLOBAL, "preuve_bonus_terme_specifique", 0) * len(set(map(norm, specific)))
        sc += p_int(ENGINE_PARAMS_GLOBAL, "preuve_bonus_signal_fort", 0) * len(set(map(norm, matched_strong)))
        if not has_bim_signal(seg) and not matched_strong and len(specific) < min_specific_no_bim:
            sc -= p_int(ENGINE_PARAMS_GLOBAL, "preuve_penalite_sans_signal_bim", 0)
        if sc > best_sc:
            best_sc, best_pg, best_txt, best_key = sc, str(page), seg, key
    threshold = p_int(ENGINE_PARAMS_GLOBAL, "seuil_preuve_ciblee", 0)
    if best_sc < threshold or not best_txt:
        return "", "", best_sc, ""
    return best_pg, clean_proof_text(best_txt), best_sc, best_key


def appliquer_gouvernance_preuves(resultats: Dict[str, Dict], params: Dict[str, object]) -> None:
    """Reconstruit l'axe présence uniquement depuis les preuves conservées.

    Aucune exception par bloc, corps d'état ou usage BIM n'est définie ici :
    les règles de détection/exclusion restent dans le classeur de paramétrage.
    """
    code_abs = p_text(params, "presence_code_absente")
    code_conv = p_text(params, "presence_code_convention")
    code_cctp = p_text(params, "presence_code_cctp")
    code_both = p_text(params, "presence_code_convention_cctp")
    for res in resultats.values():
        has_conv = bool(clean(res.get("conv_pf", "")))
        has_cctp = bool(clean(res.get("cctp_pf", "")))
        if has_conv and has_cctp:
            res["presence_contractuelle"] = code_both
        elif has_conv:
            res["presence_contractuelle"] = code_conv
        elif has_cctp:
            res["presence_contractuelle"] = code_cctp
        else:
            res["presence_contractuelle"] = code_abs

def _valeur_regle_correspond(actual: object, expected: object) -> bool:
    """Comparaison générique des conditions déclarées dans les feuilles de règles."""
    exp = [clean(x) for x in str(expected or "").split(";") if clean(x)]
    if not exp:
        return True
    actual_text = clean(actual)
    actual_norm = norm(actual_text)
    positives, negatives = [], []
    for token in exp:
        neg = token.startswith("!")
        raw = token[1:] if neg else token
        raw_upper = clean(raw).upper()
        if raw_upper == "VIDE":
            ok = not actual_text
        elif raw_upper == "NON_VIDE":
            ok = bool(actual_text)
        else:
            ok = actual_norm == norm(raw)
        (negatives if neg else positives).append(ok)
    if any(negatives):
        return False
    return any(positives) if positives else True


def appliquer_regles_interblocs(resultats: Dict[str, Dict], regles: List[Dict[str, object]]) -> None:
    """Applique 18_Regles_Interblocs sans connaissance métier des blocs dans Python."""
    def _ordre(row: Dict[str, object]) -> int:
        # Cycle 41 : l'ordre de résolution d'une preuve peut être distinct de
        # l'ordre historique de la règle. La priorité reste entièrement
        # paramétrée dans Excel (19_Regles_Documentaires). Si la nouvelle
        # colonne est vide, le comportement antérieur basé sur ``ordre`` est
        # conservé à l'identique.
        raw = row.get("priorite_resolution", "")
        if clean(raw):
            try:
                return int(float(raw))
            except (TypeError, ValueError):
                pass
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999
    for rule in sorted(regles or [], key=_ordre):
        if clean(rule.get("actif", "Oui")).lower() == "non":
            continue
        src = resultats.get(clean(rule.get("bloc_source")))
        tgt = resultats.get(clean(rule.get("bloc_cible")))
        if not src or not tgt:
            continue
        champ_src = clean(rule.get("champ_source"))
        champ_cond = clean(rule.get("champ_condition_cible"))
        champ_cond_2 = clean(rule.get("champ_condition_cible_2"))
        if champ_src and not _valeur_regle_correspond(src.get(champ_src, ""), rule.get("valeurs_source", "")):
            continue
        if champ_cond and not _valeur_regle_correspond(tgt.get(champ_cond, ""), rule.get("valeurs_condition_cible", "")):
            continue
        if champ_cond_2 and not _valeur_regle_correspond(tgt.get(champ_cond_2, ""), rule.get("valeurs_condition_cible_2", "")):
            continue
        champ_cible = clean(rule.get("champ_cible"))
        if champ_cible:
            valeur = clean(rule.get("valeur_cible"))
            # Conserver le type du champ cible quand il existe déjà. Cela permet
            # aux règles Excel de modifier aussi des indicateurs booléens sans
            # introduire de sémantique métier dans le moteur.
            if isinstance(tgt.get(champ_cible), bool):
                valeur = valeur.lower() in {"oui", "yes", "true", "1"}
            tgt[champ_cible] = valeur
        if clean(rule.get("marquer_exclusion_interbloc")).lower() in {"oui", "yes", "true", "1"}:
            tgt["interblock_exclusion"] = True
            tgt["interblock_reason"] = clean(rule.get("message"))
            # Conserver la preuve ayant déclenché l'exclusion interbloc afin que
            # les livrables puissent expliquer la décision. Le mécanisme est
            # générique : aucune connaissance de l'identité du bloc source/cible.
            tgt["interblock_source_evidence"] = {
                "conv_pf": clean(src.get("conv_pf")),
                "conv_pg": clean(src.get("conv_pg")),
                "cctp_pf": clean(src.get("cctp_pf")),
                "cctp_pg": clean(src.get("cctp_pg")),
            }
            if tgt["interblock_reason"]:
                tgt["conclusion"] = tgt["interblock_reason"]


def _rule_text_match(text: str, all_terms: object, any_terms: object, excluded_terms: object) -> bool:
    all_kws = split_kw(all_terms)
    any_kws = split_kw(any_terms)
    excluded = split_kw(excluded_terms)
    if not text:
        return False
    # Les termes positifs doivent disposer d'au moins une occurrence non niée
    # localement. Lorsqu'une négation fait déjà partie du terme configuré dans
    # Excel, elle n'est pas extérieure à son occurrence et la règle continue à
    # fonctionner comme paramétrée.
    if all_kws and not all(keyword_has_non_negated_occurrence(text, kw) for kw in all_kws):
        return False
    if any_kws and not any(keyword_has_non_negated_occurrence(text, kw) for kw in any_kws):
        return False
    if excluded and any(contains(text, kw) for kw in excluded):
        return False
    return bool(all_kws or any_kws or not excluded)


def _portee_lot_tokens(value: object) -> set[str]:
    """Retourne des jetons normalisés utilisables pour comparer un libellé de lot
    avec un acteur/périmètre extrait d'une clause. Les mots vides viennent du
    paramétrage moteur ; aucune discipline ni aucun lot n'est codé en dur.
    """
    stopwords = {
        norm(x) for x in (ENGINE_PARAMS_GLOBAL.get("portee_lot_mots_vides") or [])
        if clean(x)
    }
    tokens: set[str] = set()
    acronym_min = max(2, p_int(ENGINE_PARAMS_GLOBAL, "portee_lot_acronyme_longueur_min", 2))
    acronym_max = max(acronym_min, p_int(ENGINE_PARAMS_GLOBAL, "portee_lot_acronyme_longueur_max", 8))
    # On tokenise le libellé original pour conserver l'information de casse.
    # Un acronyme court en capitales (2..N caractères paramétrés) peut être
    # significatif, alors qu'un mot ordinaire de moins de 3 caractères reste
    # ignoré. Aucun acronyme de discipline n'est codé en dur.
    for raw in re.findall(r"[A-Za-zÀ-ÿ0-9]+", clean(value)):
        token = norm(raw)
        is_short_acronym = (
            raw.isupper()
            and any(ch.isalpha() for ch in raw)
            and acronym_min <= len(raw) <= acronym_max
        )
        if not token or token.isdigit() or (len(token) < 3 and not is_short_acronym) or token in stopwords:
            continue
        # Canonicalisation minimale des pluriels fréquents afin de rapprocher
        # des libellés singulier/pluriel sans dictionnaire métier.
        if len(token) > 4 and token.endswith("s"):
            token = token[:-1]
        if token and token not in stopwords:
            tokens.add(token)
    return tokens


def _portee_lot_correspond(scope_text: object, lot: object, min_tokens: object = None) -> bool:
    """Vérifie génériquement qu'une clause visant un acteur correspond au lot analysé.

    Le seuil et les mots vides sont paramétrés dans 05_Parametres_Moteur.
    En cas de périmètre non identifiable, la fonction échoue de manière sûre
    (False) afin de ne jamais exclure un bloc sur une clause visant un autre acteur.
    """
    scope_tokens = _portee_lot_tokens(scope_text)
    lot_tokens = _portee_lot_tokens(lot)
    if not scope_tokens or not lot_tokens:
        return False
    try:
        threshold = int(float(min_tokens)) if clean(min_tokens) else p_int(ENGINE_PARAMS_GLOBAL, "portee_lot_min_tokens_defaut", 1)
    except (TypeError, ValueError):
        threshold = p_int(ENGINE_PARAMS_GLOBAL, "portee_lot_min_tokens_defaut", 1)
    threshold = max(1, threshold)
    return len(scope_tokens & lot_tokens) >= threshold


def appliquer_regles_documentaires(
    resultats: Dict[str, Dict],
    regles: List[Dict[str, object]],
    lot: str,
    ctx_global: Dict[str, object],
    messages: Dict[str, str],
    source_pages: Optional[Dict[str, List[object]]] = None,
) -> None:
    """Applique 19_Regles_Documentaires aux seules preuves effectivement retenues."""
    def _ordre(row: Dict[str, object]) -> int:
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999
    by_block: Dict[str, List[Dict[str, object]]] = {}
    for rule in sorted(regles or [], key=_ordre):
        if clean(rule.get("actif", "Oui")).lower() == "non":
            continue
        bid = clean(rule.get("id_bloc"))
        if bid:
            by_block.setdefault(bid, []).append(rule)
    source_labels = {
        "CCTP": render_message(messages, "source_cctp_label") or "CCTP du lot",
        "CONVENTION": render_message(messages, "source_convention_label") or "Convention BIM",
        "CCAP": "CCAP",
        "COMPLEMENT": render_message(messages, "source_supplement_label") or "Pièce complémentaire",
    }

    def _page_candidates(source_code: str, block_id: str) -> List[Tuple[str, str, str, str, str]]:
        """Normalise les pages sources sans imposer leur nature au moteur.

        Les entrées historiques CCTP/Convention restent des tuples (page, texte).
        Une pièce complémentaire peut être fournie sous forme de dictionnaire avec
        son libellé, son type de preuve et la liste des blocs auxquels l'utilisateur
        l'a ajoutée. Ce dernier champ évite qu'une annexe ajoutée pour lever une
        vigilance B06B ne soit utilisée aveuglément sur tous les blocs.
        """
        rows: List[Tuple[str, str, str, str, str]] = []
        for entry in ((source_pages or {}).get(source_code) or []):
            page = text = label = kind = ""
            related: set[str] = set()
            if isinstance(entry, dict):
                page = clean(entry.get("page"))
                text = str(entry.get("text") or "")
                label = clean(entry.get("source")) or source_labels.get(source_code, source_code)
                kind = clean(entry.get("kind")) or source_code.lower()
                related = {clean(x).upper() for x in (entry.get("related_blocks") or []) if clean(x)}
            elif isinstance(entry, (list, tuple)) and len(entry) >= 2:
                page = clean(entry[0])
                text = str(entry[1] or "")
                label = clean(entry[2]) if len(entry) >= 3 else source_labels.get(source_code, source_code)
                kind = clean(entry[3]) if len(entry) >= 4 else source_code.lower()
            else:
                continue
            if related and clean(block_id).upper() not in related:
                continue
            rows.append((source_code, re.sub(r"[ \t]+", " ", text).strip(), page, label, kind))
        return rows
    def _ordre_doc(row: Dict[str, object]) -> int:
        raw = row.get("priorite_resolution", "")
        if clean(raw):
            try:
                return int(float(raw))
            except (TypeError, ValueError):
                pass
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999

    for bid, res in resultats.items():
        for rule in sorted(by_block.get(bid, []), key=_ordre_doc):
            # Condition interbloc générique (Cycle 46). Le bloc source, le champ
            # observé et les valeurs attendues sont entièrement déclarés dans Excel.
            # Le moteur ne connaît pas le sens métier de la dépendance.
            condition_block_id = clean(rule.get("bloc_condition_source"))
            if condition_block_id:
                condition_result = resultats.get(condition_block_id) or {}
                condition_field = clean(rule.get("champ_condition_bloc_source"))
                if not condition_field or not _valeur_regle_correspond(
                    condition_result.get(condition_field, ""),
                    rule.get("valeurs_condition_bloc_source", ""),
                ):
                    continue
            # Une règle documentaire peut être conditionnée par un champ du résultat
            # courant. Le champ et les valeurs attendues viennent d'Excel ; le moteur
            # reste générique et ne connaît ni le bloc ni le sens métier de la condition.
            result_condition_field = clean(rule.get("champ_condition_resultat"))
            if result_condition_field and not _valeur_regle_correspond(
                res.get(result_condition_field, ""), rule.get("valeurs_condition_resultat", "")
            ):
                continue
            source_req = clean(rule.get("source", "ANY")).upper()
            search_full_document = clean(rule.get("recherche_document_complet")).lower() in {"oui", "yes", "true", "1"}
            candidates: List[Tuple[str, str, str, str, str]] = []
            if source_req in {"ANY", "CCTP"}:
                if search_full_document:
                    # Conserver les retours à la ligne pendant la recherche sur document complet.
                    # Ils constituent une frontière de clause utile pour les regex paramétrées
                    # dans Excel. Seuls les espaces horizontaux sont normalisés ici.
                    candidates.extend(_page_candidates("CCTP", bid))
                else:
                    candidates.append(("CCTP", clean(res.get("cctp_pf")), clean(res.get("cctp_pg")), source_labels["CCTP"], "cctp"))
            if source_req in {"ANY", "CONVENTION"}:
                if search_full_document:
                    candidates.extend(_page_candidates("CONVENTION", bid))
                else:
                    candidates.append(("CONVENTION", clean(res.get("conv_pf")), clean(res.get("conv_pg")), source_labels["CONVENTION"], "convention"))
            # Le CCAP n'entre volontairement PAS dans ANY : seules des règles
            # Excel explicitement déclarées source=CCAP peuvent l'utiliser comme
            # preuve métier. Cela préserve le comportement historique tout en
            # autorisant les clauses prescriptives ciblées.
            if source_req == "CCAP":
                if search_full_document:
                    candidates.extend(_page_candidates("CCAP", bid))
                else:
                    candidates.append(("CCAP", "", "", source_labels["CCAP"], "ccap"))
            # Une pièce ajoutée après la première analyse est une source explicite.
            # Elle n'entre pas implicitement dans les règles ANY : seules les règles
            # Excel qui déclarent source=COMPLEMENT peuvent transformer son contenu
            # en preuve contractuelle. Cela évite qu'une annexe secondaire ne prenne
            # automatiquement le pas sur le CCTP ou la Convention.
            if source_req == "COMPLEMENT":
                if search_full_document:
                    candidates.extend(_page_candidates("COMPLEMENT", bid))
                else:
                    candidates.append((
                        "COMPLEMENT", clean(res.get("supp_pf")), clean(res.get("supp_pg")),
                        clean(res.get("supp_source")) or source_labels["COMPLEMENT"],
                        clean(res.get("supp_kind")) or "supp0",
                    ))
            # Cycle 38 : lors d'une recherche sur document complet, plusieurs pages
            # peuvent satisfaire le préfiltre lexical alors qu'une seule contient la
            # clause structurée attendue par la regex Excel. Il faut donc chercher la
            # première candidate qui satisfait À LA FOIS le préfiltre et la regex,
            # au lieu de s'arrêter sur la première page contenant un mot générique.
            extraction_regex = clean(rule.get("extraction_regex"))
            matched = None
            matched_extraction = None
            for src_code, proof, page, matched_label, matched_kind in candidates:
                if not _rule_text_match(proof, rule.get("contient_tous"), rule.get("contient_un"), rule.get("exclut")):
                    continue
                extraction_match = None
                if extraction_regex:
                    try:
                        extraction_match = re.search(extraction_regex, proof, flags=re.I | re.S)
                    except re.error:
                        extraction_match = None
                    if not extraction_match:
                        continue
                matched = (src_code, proof, page, matched_label, matched_kind)
                matched_extraction = extraction_match
                break
            if not matched:
                continue
            src_code, proof, page, matched_label, matched_kind = matched

            # Une regex facultative, définie dans Excel, peut isoler la clause exacte
            # et exposer ses groupes nommés comme variables de restitution
            # ({obligation}, {jalon}, etc.). Le code reste entièrement agnostique du bloc.
            extracted_ctx = {}
            if extraction_regex and matched_extraction:
                extracted_ctx = {
                    key: clean(value)
                    for key, value in matched_extraction.groupdict().items()
                    if value not in (None, "")
                }
                targeted_proof = clean(
                    extracted_ctx.get("obligation")
                    or extracted_ctx.get("preuve")
                    or matched_extraction.group(0)
                )
                if targeted_proof:
                    proof = targeted_proof

            # Cycle 50 : une clause CCTP dont le sujet est explicitement un autre
            # lot ne peut jamais devenir la preuve du lot analysé, même si une
            # règle métier plus générale reconnaît ses mots-clés. Les clauses
            # générales du CCTP restent, elles, admissibles.
            if src_code == "CCTP" and not _clause_compatible_portee_lot(proof, lot):
                continue

            # Certaines règles Excel portent sur un acteur ou une discipline
            # explicitement nommé(e) dans la clause. Lorsqu'elles activent le
            # contrôle de portée, le moteur compare génériquement le groupe
            # extrait (par défaut « acteur ») avec le lot analysé. Cela évite
            # qu'une exclusion visant un autre lot ne soit appliquée globalement.
            if clean(rule.get("verifier_portee_lot")).lower() in {"oui", "yes", "true", "1"}:
                scope_group = clean(rule.get("groupe_portee_lot")) or "acteur"
                scope_text = extracted_ctx.get(scope_group, "")
                if not _portee_lot_correspond(scope_text, lot, rule.get("portee_lot_min_tokens")):
                    continue

            ctx = {
                **(ctx_global or {}), "lot": lot,
                **extracted_ctx,
                "source": matched_label or source_labels.get(src_code, src_code),
                "page": page or render_message(messages, "evidence_page_unknown"),
                "preuve": proof,
                # Variables génériques utilisables dans les textes Excel pour croiser
                # les deux pièces sans coder le sens métier du bloc dans Python.
                "page_convention": clean(res.get("conv_pg")) or render_message(messages, "evidence_page_unknown"),
                "page_cctp": clean(res.get("cctp_pg")) or render_message(messages, "evidence_page_unknown"),
                "preuve_convention": clean(res.get("conv_pf")),
                "preuve_cctp": clean(res.get("cctp_pf")),
            }
            rule_id = clean(rule.get("id_regle"))
            aggregation = clean(rule.get("agregation", "REMPLACE")).upper() or "REMPLACE"
            demand_template = clean(rule.get("demande"))
            response_template = clean(rule.get("reponse"))
            action_template = clean(rule.get("action_standard"))
            rendered_demand = _sub_ctx(demand_template, ctx, messages)
            rendered_response = _sub_ctx(response_template, ctx, messages)
            rendered_action = _sub_ctx(action_template, ctx, messages)

            def _append_unique_text(existing: object, value: object) -> str:
                current = clean(existing)
                addition = clean(value)
                if not addition:
                    return current
                if not current:
                    return addition
                if addition.casefold() in current.casefold():
                    return current
                return current.rstrip() + " " + addition

            append_to_existing = aggregation == "AJOUTE" and bool(res.get("specific_rule_id"))
            if append_to_existing:
                res["specific_rule_id"] = _append_unique_text(res.get("specific_rule_id"), rule_id)
                res["specific_demand"] = _append_unique_text(res.get("specific_demand"), rendered_demand)
                res["specific_response"] = _append_unique_text(res.get("specific_response"), rendered_response)
                res["specific_action"] = _append_unique_text(res.get("specific_action"), rendered_action)
                res.setdefault("specific_evidence_extra", []).append({
                    "kind": src_code.lower(),
                    "source": matched_label or source_labels.get(src_code, src_code),
                    "page": page,
                    "text": proof,
                })
            else:
                res["specific_rule_id"] = rule_id
                res["specific_demand"] = rendered_demand
                res["specific_response"] = rendered_response
                res["specific_action"] = rendered_action
                res["specific_source"] = src_code
                res["specific_source_label"] = matched_label or source_labels.get(src_code, src_code)
                res["specific_page"] = page

            capacity_qids = split_kw(rule.get("questions_capacite", ""))
            if capacity_qids:
                existing_qids = list(res.get("capacity_question_ids") or [])
                seen_qids = {clean(q).casefold() for q in existing_qids if clean(q)}
                for qid in capacity_qids:
                    if clean(qid).casefold() not in seen_qids:
                        existing_qids.append(clean(qid))
                        seen_qids.add(clean(qid).casefold())
                res["capacity_question_ids"] = existing_qids
            if clean(rule.get("origine_action_specifique")).lower() in {"oui", "yes", "true", "1"}:
                source_label = matched_label or source_labels.get(src_code, src_code)
                res["specific_action_origin"] = f"{source_label} p.{page}" if page else source_label
            # Si la règle a isolé une preuve plus précise, cette clause devient aussi
            # la preuve affichée afin que la décision soit directement auditable.
            if extraction_regex and proof and not append_to_existing:
                if src_code == "CCTP":
                    res["cctp_pf"] = proof
                    res["cctp_pg"] = page
                    res["cctp_extrait"] = proof[:200]
                elif src_code == "CONVENTION":
                    res["conv_pf"] = proof
                    res["conv_pg"] = page
                    res["conv_extrait"] = proof[:200]
                elif src_code == "CCAP":
                    # Le CCAP est conservé comme preuve complémentaire traçable ;
                    # il ne remplace pas les champs CCTP/Convention historiques.
                    res.setdefault("specific_evidence_extra", []).append({
                        "kind": "ccap", "source": matched_label or "CCAP",
                        "page": page, "text": proof,
                    })
                elif src_code == "COMPLEMENT":
                    res["supp_pf"] = proof
                    res["supp_pg"] = page
                    res["supp_source"] = matched_label or source_labels["COMPLEMENT"]
                    res["supp_kind"] = matched_kind or "supp0"
            # Les modes/forçages restent définis dans Excel. Le moteur se contente
            # de les recopier dans le résultat public lorsqu'ils sont renseignés.
            for col, target in (
                ("action_mode_override", "action_mode_override"),
                ("statut_public_override", "statut_public_override"),
                ("priorite_override", "priorite_override"),
                ("question_mode_override", "question_mode_override"),
                ("decision_override", "decision_override"),
                ("response_mode_override", "response_mode_override"),
                # Cycle 33 : la temporalité d'une action reste un paramètre métier
                # du classeur. Le moteur se contente de recopier le code de phase.
                ("phase_action_override", "specific_action_phase"),
            ):
                value = clean(rule.get(col))
                if value:
                    res[target] = value
            # Une règle documentaire peut exposer un ou plusieurs indicateurs génériques
            # au moteur interblocs. Les noms de champs et leurs valeurs viennent entièrement
            # d'Excel ; Python n'embarque aucune connaissance du bloc ni de la condition métier.
            for field_col, value_col in (("champ_resultat", "valeur_resultat"),
                                         ("champ_resultat_2", "valeur_resultat_2")):
                result_field = clean(rule.get(field_col))
                if not result_field:
                    continue
                result_value = clean(rule.get(value_col))
                if isinstance(res.get(result_field), bool):
                    result_value = result_value.lower() in {"oui", "yes", "true", "1"}
                res[result_field] = result_value
            if clean(rule.get("continuer_apres_match")).lower() not in {"oui", "yes", "true", "1"}:
                break

        # Traçabilité finale des valeurs de contexte réellement présentes dans la
        # restitution du bloc. Le contexte global provient de la Convention et
        # porte, lorsqu'elle a été détectée, la page et l'extrait de chaque valeur.
        # On n'ajoute une preuve complémentaire que si la valeur apparaît dans le
        # texte final ET qu'aucune preuve déjà retenue sur cette page ne la trace.
        # Ainsi une page image sans texte ne peut jamais fournir la valeur, et une
        # occurrence lisible située sur une autre page reste correctement sourcée.
        final_text = " ".join(clean(res.get(key)) for key in (
            "specific_demand", "specific_response", "specific_action"
        ) if clean(res.get(key)))
        if final_text:
            for ctx_key, ctx_value in sorted((ctx_global or {}).items()):
                if ctx_key.endswith(("_page", "_extrait", "_source")):
                    continue
                value = clean(ctx_value)
                if not value or not contains(final_text, value):
                    continue
                ctx_page = clean((ctx_global or {}).get(f"{ctx_key}_page"))
                ctx_excerpt = clean((ctx_global or {}).get(f"{ctx_key}_extrait"))
                if not ctx_page or not ctx_excerpt:
                    continue
                ctx_source_code = clean((ctx_global or {}).get(f"{ctx_key}_source")).upper() or "CONVENTION"
                ctx_source_label = source_labels.get(ctx_source_code, ctx_source_code)
                ctx_kind = "cctp" if ctx_source_code == "CCTP" else "convention"
                def _kind_family(value: object) -> str:
                    k = clean(value).casefold()
                    if k in {"conv", "convention"}:
                        return "convention"
                    if k == "cctp":
                        return "cctp"
                    return k
                value_norms = [norm(value)]
                value_without_suffix = clean(re.sub(r"\s*\([^)]{1,24}\)\s*$", "", value))
                if value_without_suffix and norm(value_without_suffix) not in value_norms:
                    value_norms.append(norm(value_without_suffix))

                already_traced = False
                # Preuves principales déjà retenues. Une forme canonique peut
                # comporter un acronyme entre parenthèses (ex. nom + sigle) alors
                # que la clause source ne contient que le nom long : les deux
                # formes sont donc comparées sans connaissance métier spécifique.
                for proof_kind, proof_page, proof_text in (
                    ("conv", clean(res.get("conv_pg")), clean(res.get("conv_pf"))),
                    ("cctp", clean(res.get("cctp_pg")), clean(res.get("cctp_pf"))),
                ):
                    proof_norm = norm(proof_text)
                    if _kind_family(proof_kind) == _kind_family(ctx_kind) and proof_page == ctx_page and any(v and v in proof_norm for v in value_norms):
                        already_traced = True
                        break
                # Preuves complémentaires déjà ajoutées par les règles Excel.
                if not already_traced:
                    for extra in res.get("specific_evidence_extra") or []:
                        if not isinstance(extra, dict):
                            continue
                        if _kind_family(extra.get("kind")) != _kind_family(ctx_kind):
                            continue
                        if clean(extra.get("page")) != ctx_page:
                            continue
                        extra_norm = norm(clean(extra.get("text")))
                        if any(v and v in extra_norm for v in value_norms):
                            already_traced = True
                            break
                if already_traced:
                    continue
                res.setdefault("specific_evidence_extra", []).append({
                    "kind": ctx_kind,
                    "source": ctx_source_label,
                    "page": ctx_page,
                    "text": ctx_excerpt,
                })



# ═══════════════════════════════════════════════════════════════════════════════
# ÉVALUATION STATUT (MODULE LECTURE)
# ═══════════════════════════════════════════════════════════════════════════════
def _kws_depuis_signaux(bid: str, signaux: Optional[pd.DataFrame]) -> Tuple[List[str], List[str]]:
    """
    Lit la feuille 30_CCTP_Signaux et retourne les mots-clés CCTP additionnels
    concernant ce bloc, séparés par poids : (mots_moyens, mots_forts).
    Vient enrichir keywords_cctp_confirme / keywords_preuve_forte de 10_Blocs
    -- catalogue complémentaire, aucun id_bloc codé en dur, tout vient de
    la colonne bloc_concerne (une ligne peut viser plusieurs blocs : "B01,B03").
    """
    extra_conf, extra_fort = [], []
    if signaux is None or signaux.empty:
        return extra_conf, extra_fort
    for _, srow in signaux.iterrows():
        blocs_concernes = split_kw(srow.get("bloc_concerne", ""))
        if bid not in blocs_concernes:
            continue
        kws_sig = split_kw(srow.get("signal_textuel", ""))
        if clean(srow.get("poids", "")).lower() == "fort":
            extra_fort.extend(kws_sig)
        else:
            extra_conf.extend(kws_sig)
    return extra_conf, extra_fort


def charger_logo(logo_path: str | None) -> Dict[str, Any]:
    """Lit un logo entreprise (PNG/JPEG) fourni par l'utilisateur et retourne
    à la fois une data-URI (pour les livrables HTML) et les octets bruts +
    format (pour les en-têtes/pieds de page PDF via reportlab). Retourne un
    dict vide si aucun logo n'est fourni -- les gabarits retombent alors sur
    le sigle texte (--marque / entreprise)."""
    if not logo_path:
        return {}
    p = Path(logo_path)
    if not p.exists() or p.stat().st_size == 0:
        return {}
    data = p.read_bytes()
    suffix = p.suffix.lower()
    mime = "image/png" if suffix == ".png" else "image/jpeg"
    return {
        "logo_data_uri": f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}",
        "logo_bytes": data,
        "logo_format": "PNG" if suffix == ".png" else "JPEG",
    }


def p_text(params: Dict[str, object], cle: str) -> str:
    return str(params.get(cle, "") or "")


def p_int(params: Dict[str, object], cle: str, defaut: int = 0) -> int:
    v = params.get(cle, defaut)
    try: return int(v)
    except (TypeError, ValueError): return defaut

def p_float(params: Dict[str, object], cle: str, defaut: float = 0.0) -> float:
    v = params.get(cle, defaut)
    try: return float(v)
    except (TypeError, ValueError): return defaut




def evaluate(bloc: pd.Series, cctp_text: str, lot: str,
             conv_text: str = "",
             signaux: Optional[pd.DataFrame] = None,
             params: Optional[Dict[str, object]] = None,
             messages: Optional[Dict[str, str]] = None,
             axes_config: Optional[Dict] = None) -> Dict:
    """Évalue présence documentaire et applicabilité avec les seules règles Excel."""
    params = params or {}
    messages = messages or {}
    axes_config = axes_config or {}
    titre = clean(bloc.get("titre_bloc", ""))
    bid = clean(bloc.get("id_bloc", ""))
    kws_conf = split_kw(bloc.get("keywords_cctp_confirme", ""))
    kws_nu = split_kw(bloc.get("keywords_cctp_nuance", ""))
    kws_exclu = split_kw(bloc.get("keywords_cctp_exclu", ""))
    kws_conv = split_kw(bloc.get("keywords_convention", ""))
    kws_fort = split_kw(bloc.get("keywords_preuve_forte", ""))
    extra_conf, extra_fort = _kws_depuis_signaux(bid, signaux)
    kws_conf += extra_conf
    kws_fort += extra_fort

    has_cctp = bool(clean(cctp_text))
    has_conf_cctp = any_kw_non_negated(cctp_text, kws_conf)
    has_nu_cctp = any_kw_non_negated(cctp_text, kws_nu)
    has_exclu_cctp = bool(kws_exclu and any_kw(cctp_text, kws_exclu))
    has_fort_conv = any_kw_non_negated(conv_text, kws_fort)
    has_conf_conv = any_kw_non_negated(conv_text, kws_conf + kws_fort)
    sc_conv = score_seg(conv_text, kws_conv, kws_fort) if conv_text else -999
    seuil_conv = p_int(params, "seuil_presence_convention", 0)
    conv_explicite = has_fort_conv or has_conf_conv or sc_conv >= seuil_conv
    cctp_explicite = has_cctp and (has_conf_cctp or has_nu_cctp or has_exclu_cctp)

    code_abs = p_text(params, "presence_code_absente")
    code_conv = p_text(params, "presence_code_convention")
    code_cctp = p_text(params, "presence_code_cctp")
    code_both = p_text(params, "presence_code_convention_cctp")
    if conv_explicite and cctp_explicite:
        presence = code_both
    elif conv_explicite:
        presence = code_conv
    elif cctp_explicite:
        presence = code_cctp
    else:
        presence = code_abs
    source = render_message(messages, f"source_presence_{presence}")

    raw_nuance = clean(bloc.get("statut_si_nuance_seul", "")).upper()
    valid_app = set((axes_config.get("applicabilite") or {}).keys())
    if presence == code_abs:
        applicabilite = p_text(params, "applicabilite_si_absence")
    elif has_exclu_cctp:
        applicabilite = p_text(params, "applicabilite_si_cctp_exclu")
    elif has_conf_cctp:
        applicabilite = p_text(params, "applicabilite_si_cctp_confirme")
    elif has_nu_cctp:
        applicabilite = raw_nuance if raw_nuance in valid_app else p_text(params, "applicabilite_si_cctp_nuance")
    elif has_cctp:
        applicabilite = p_text(params, "applicabilite_si_silence_cctp")
    else:
        applicabilite = p_text(params, "applicabilite_sans_cctp")

    if has_exclu_cctp:
        kw = next((k for k in kws_exclu if contains(cctp_text, k)), "")
    elif has_conf_cctp:
        kw = next((k for k in kws_conf if keyword_has_non_negated_occurrence(cctp_text, k)), "")
    elif has_nu_cctp:
        kw = next((k for k in kws_nu if keyword_has_non_negated_occurrence(cctp_text, k)), "")
    elif has_conf_conv:
        kw = next((k for k in (kws_fort + kws_conf) if keyword_has_non_negated_occurrence(conv_text, k)), "")
    else:
        kw = ""

    vals = {"titre": titre, "lot": lot, "source": source, "detection": kw}
    if presence == code_abs:
        conclusion = render_message(messages, "conclusion_presence_ABSENTE", **vals)
    else:
        conclusion = render_message(messages, f"conclusion_applicabilite_{applicabilite}", **vals)
    return {
        "presence_contractuelle": presence,
        "applicabilite_lot": applicabilite,
        "capacite_entreprise": capacity_code_for_pct(axes_config, None),
        "conclusion": conclusion,
        "terme_detecte": kw,
        "exclusion_explicit": bool(has_exclu_cctp),
    }


# =============================================================================
# FILTRE GENERIQUE -- QUESTIONS BIM MANAGER
# =============================================================================




def _extraction_metadata_generique(texte: str, regles: Optional[pd.DataFrame],
                                    champs: List[str]) -> Dict[str, str]:
    """
    Moteur générique d'extraction de métadonnées (indice, phase, date,
    émetteur, MOA, MOE, entreprise générale, BIM Manager, Coordinateur BIM...)
    entièrement piloté par la feuille 36_Extraction_Metadata : aucune regex
    de champ n'est codée en dur ici, seulement le moteur qui les applique.

    Pour chaque champ, essaie les règles actives par priorité croissante
    (la plus petite priorité d'abord) jusqu'à trouver un match ; ne garde
    que la première ligne du texte capturé et ignore les valeurs listées
    dans `valeurs_exclues` (ex: éviter de capturer un autre libellé de
    colonne du même tableau).
    """
    res: Dict[str, str] = {c: "" for c in champs}
    if regles is None or regles.empty:
        return res
    r = regles.copy()
    r["_prio"] = r["priorite"].apply(lambda x: int(float(x)) if str(x).strip() else 999)
    r = r.sort_values("_prio")
    for _, row in r.iterrows():
        champ = clean(row.get("champ", ""))
        if champ not in champs or res.get(champ):
            continue
        actif = clean(row.get("actif", "Oui")).lower()
        if actif in ("non", "false", "0"):
            continue
        pattern = str(row.get("pattern", "") or "")
        if not pattern:
            continue
        groupe = int(row.get("groupe_capture", 1) or 1)
        exclusions = {e.lower() for e in split_kw(row.get("valeurs_exclues", ""))}
        try:
            m = re.search(pattern, texte)
        except re.error:
            continue
        if not m:
            continue
        try:
            val = clean(m.group(groupe))
        except (IndexError, Exception):
            val = clean(m.group(0))
        val = val.split("\n")[0].strip()
        if len(val) < 1 or val.lower() in exclusions:
            continue
        value_max = p_int(ENGINE_PARAMS_GLOBAL, f"metadata_{champ}_longueur_max", 0)
        if not value_max:
            value_max = p_int(ENGINE_PARAMS_GLOBAL, "metadata_valeur_longueur_max", len(val))
        acronym_max = p_int(ENGINE_PARAMS_GLOBAL, "metadata_acronyme_longueur_max", 0)
        clipped = val[:value_max] if value_max else val
        res[champ] = clipped.upper() if acronym_max and len(val) <= acronym_max else clipped
    return res


def detecter_points_vigilance_documentaires(conv_text: str, cctp_text: str, ccap_text: str,
                                              documents_analyses: str, params: Dict[str, object],
                                              messages: Dict[str, str],
                                              qualite_documents: Optional[Dict[str, Dict[str, object]]] = None,
                                              conv_pages: Optional[List[Tuple[int, str]]] = None,
                                              cctp_pages: Optional[List[Tuple[int, str]]] = None,
                                              ccap_pages: Optional[List[Tuple[int, str]]] = None,
                                              additional_pages: Optional[List[Tuple[str, List[Tuple[int, str]]]]] = None,
                                              ccap_hierarchy_pages: Optional[set[int]] = None,
                                              document_rules: Optional[List[Dict[str, object]]] = None) -> List[Dict[str, str]]:
    """Détecte des vigilances documentaires sans influer sur B01-B12.

    Cycle 33 : chaque vigilance issue d'un contenu lisible conserve sa provenance
    (document, page PDF et extrait). Les références de pièces sont nettoyées avant
    comparaison afin de ne pas transformer une phrase entière en faux nom de
    document. Le CCAP reste strictement informatif pour B01-B12.
    """
    out: List[Dict[str, str]] = []
    excerpt_max = max(120, p_int(params, "vigilance_extrait_longueur_max", 360))

    def _excerpt(page_text: str, start: int, end: int) -> str:
        text = str(page_text or "")
        if not text:
            return ""
        left = max(0, start - 180)
        right = min(len(text), end + 260)
        snippet = clean(text[left:right])
        if len(snippet) > excerpt_max:
            snippet = snippet[:max(0, excerpt_max - 3)].rstrip() + "..."
        return snippet

    # La qualité de lecture est une vigilance indépendante des blocs métier.
    for role, quality in (qualite_documents or {}).items():
        status = clean((quality or {}).get("statut", "")).upper()
        if not status or status == "OK":
            continue
        role_up = clean(role).upper()
        label = clean((quality or {}).get("label", "")) or role_up.title()
        if status == "NON_EXPLOITABLE" and role_up == "CCTP":
            out.append({
                "type": "pdf_unreadable", "title": render_message(messages, "vigilance_cctp_unreadable_title"),
                "text": render_message(messages, "vigilance_cctp_unreadable_text"),
                "source": label, "page": "", "excerpt": "", "reference": "",
            })
        elif status == "NON_EXPLOITABLE" and role_up == "CCAP":
            out.append({
                "type": "pdf_unreadable", "title": render_message(messages, "vigilance_ccap_unreadable_title"),
                "text": render_message(messages, "vigilance_ccap_unreadable_text"),
                "source": label, "page": "", "excerpt": "", "reference": "",
            })
        elif status == "PARTIEL":
            pages_scan = [int(x) for x in ((quality or {}).get("pages_images_non_lisibles", []) or [])]
            page_field = ""
            if pages_scan:
                max_pages = max(1, p_int(params, "pdf_lecture_pages_signalees_max", 20))
                shown = pages_scan[:max_pages]
                pages_text = ", ".join(str(x) for x in shown)
                if len(pages_scan) > len(shown):
                    pages_text += f" (+{len(pages_scan) - len(shown)} autre(s))"
                page_field = pages_text
                detail = render_message(
                    messages, "vigilance_pdf_partial_scan_text", document=label,
                    pages_non_lisibles=pages_text, nombre_pages_non_lisibles=len(pages_scan),
                    pages_lisibles=(quality or {}).get("pages_lisibles", 0),
                    pages_total=(quality or {}).get("pages_total", 0),
                )
            else:
                detail = render_message(
                    messages, "vigilance_pdf_partial_text", document=label,
                    pages_lisibles=(quality or {}).get("pages_lisibles", 0),
                    pages_total=(quality or {}).get("pages_total", 0),
                )
            out.append({
                "type": "pdf_partial", "title": render_message(messages, "vigilance_pdf_partial_title"),
                "text": detail, "source": label, "page": page_field, "excerpt": "", "reference": "",
            })

    docs_norm = norm(documents_analyses or "")
    pattern = p_text(params, "vigilance_reference_document_pattern")
    bim_words = params.get("vigilance_reference_bim_mots", []) or []
    if isinstance(bim_words, str):
        bim_words = split_kw(bim_words)
    stop_words = params.get("vigilance_reference_stop_words", []) or []
    if isinstance(stop_words, str):
        stop_words = split_kw(stop_words)
    max_len = p_int(params, "vigilance_reference_longueur_max", 160)
    min_tokens = max(1, p_int(params, "vigilance_reference_min_tokens", 2))
    reject_fragments = params.get("vigilance_reference_reject_fragments", []) or []
    if isinstance(reject_fragments, str):
        reject_fragments = split_kw(reject_fragments)
    citation_required_types = params.get("vigilance_reference_types_citation_obligatoire", []) or []
    if isinstance(citation_required_types, str):
        citation_required_types = split_kw(citation_required_types)
    citation_required_types = {norm(x) for x in citation_required_types if clean(x)}
    citation_markers = params.get("vigilance_reference_citation_markers", []) or []
    if isinstance(citation_markers, str):
        citation_markers = split_kw(citation_markers)
    group_ignore_tokens = params.get("vigilance_reference_group_ignore_tokens", []) or []
    if isinstance(group_ignore_tokens, str):
        group_ignore_tokens = split_kw(group_ignore_tokens)
    group_ignore_tokens = {norm(x) for x in group_ignore_tokens if clean(x)}

    generic_tokens = {
        "cahier", "charges", "annexe", "annexes", "guide", "notice", "charte", "protocole",
        "document", "documents", "ses", "son", "sa", "leurs", "leur", "des", "les", "une", "un",
        "avec", "pour", "dans", "sur", "par", "aux", "ainsi", "que", "et",
    }

    # Cycle 42 — taxonomie des références documentaires pilotée par Excel.
    # Les lignes de 19_Regles_Documentaires dont famille_regle vaut
    # VIGILANCE_TAXONOMIE classent une référence (pièce DCE, annexe, guide,
    # norme, URL...) et indiquent si elle peut créer une vigilance. Le moteur
    # ne connaît aucun nom de projet, lot, guide ou document particulier.
    def _taxonomy_order(row: Dict[str, object]) -> int:
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999

    taxonomy_rules = [
        row for row in (document_rules or [])
        if clean(row.get("famille_regle")).upper() == "VIGILANCE_TAXONOMIE"
        and clean(row.get("actif", "Oui")).lower() not in {"non", "false", "0"}
    ]
    taxonomy_rules.sort(key=_taxonomy_order)

    def _taxonomy(reference: str) -> Dict[str, str]:
        nref = norm(reference)
        for row in taxonomy_rules:
            pattern_tax = clean(row.get("vigilance_pattern"))
            if not pattern_tax:
                continue
            try:
                matched = re.search(pattern_tax, nref, flags=re.I)
            except re.error:
                matched = None
            if not matched:
                continue
            return {
                "nature": clean(row.get("vigilance_nature")),
                "family": clean(row.get("vigilance_famille")),
                "treatment": clean(row.get("vigilance_traitement")) or "VIGILANCE",
                "citation_required": clean(row.get("vigilance_citation_obligatoire")),
                "family_label": clean(row.get("vigilance_libelle_famille")),
                "display_group": clean(row.get("vigilance_groupe_affichage")),
                "display_order": clean(row.get("vigilance_ordre_affichage")),
                "identity_note": clean(row.get("vigilance_notes_identite")),
                "rule_id": clean(row.get("id_regle")),
            }
        return {
            "nature": "", "family": "", "treatment": "VIGILANCE",
            "citation_required": "", "family_label": "",
            "display_group": "", "display_order": "",
            "identity_note": "", "rule_id": "",
        }

    def _annexe_number(reference: str) -> str:
        m = re.search(r"(?i)\bannexe(?:s)?\s*(?:n[°o]?\s*)?(\d{1,3})\b", norm(reference))
        return m.group(1) if m else ""

    def _clean_reference(raw: str) -> str:
        ref = clean(raw)
        if not ref:
            return ""
        # Si le titre de l'annexe est directement entre guillemets, le conserver.
        mquoted = re.match(r"(?i)^(annexe(?:s)?)\s*[«\"]\s*([^»\"]{2,120})[»\"]", ref)
        if mquoted:
            ref = clean(f"{mquoted.group(1)} {mquoted.group(2)}")
        else:
            # Un sous-titre/tabulation cité après le nom principal ne crée pas un
            # nouveau document : conserver le nom situé avant « : \"sous-rubrique\" ».
            ref = re.split(r"\s*:\s*[«\"]", ref, maxsplit=1)[0]
            ref = re.split(r"\s+[«\"]", ref, maxsplit=1)[0]
        ref = re.sub(r"\b\d+\s*/\s*\d+\b", " ", ref)
        # Une qualification placée entre parenthèses peut terminer le vrai titre ;
        # si une proposition grammaticale recommence après « ), », elle appartient
        # au contexte et non au nom de la pièce. Ce nettoyage est purement syntaxique.
        ref = re.sub(
            r"(\))\s*,\s+(?=(?:le|la|les|un|une|des|ce|cet|cette|ces)\b).*$",
            r"\1", ref, flags=re.I,
        )
        # Couper au premier verbe/locution de phrase configuré dans Excel.
        cut = len(ref)
        for token in stop_words:
            tok = clean(token)
            if not tok:
                continue
            m = re.search(r"(?i)(?<!\w)" + re.escape(tok) + r"(?!\w)", ref)
            if m and m.start() > 0:
                cut = min(cut, m.start())
        ref = clean(ref[:cut]).strip(" -–—:;,.")
        if max_len:
            ref = ref[:max_len].rstrip()
        if not ref:
            return ""
        if any(contains(ref, fragment) for fragment in reject_fragments if clean(fragment)):
            return ""
        toks = [t for t in re.findall(r"[a-z0-9]+", norm(ref)) if len(t) >= 3]
        significant = [t for t in toks if t not in generic_tokens]
        # Une référence très générique sans contexte BIM ni titre significatif est rejetée.
        has_bim_context = any(contains(ref, w) for w in bim_words) if bim_words else True
        if not has_bim_context:
            return ""
        if len(significant) < min_tokens and not any(t in {"bim", "exe", "aim", "gmao"} for t in significant):
            return ""
        return ref

    def _reference_type(reference: str) -> str:
        nref = norm(reference)
        if nref.startswith("cahier des charges"):
            return "cahier des charges"
        if nref.startswith("suivi de convention"):
            return "suivi de convention"
        parts = re.findall(r"[a-z0-9]+", nref)
        return parts[0] if parts else ""

    def _context_confirms_citation(page_text: str, start: int, end: int, reference: str) -> bool:
        """Évite de confondre un livrable, un titre de section ou une activité
        documentaire avec une pièce externe réellement citée.

        Cycle 42 : la taxonomie Excel peut imposer ce contrôle à une famille
        entière (ex. « Cahier des charges ») sans que Python connaisse le nom
        de la pièce. Les paramètres historiques restent un repli compatible.
        """
        tax = _taxonomy(reference)
        tax_required = clean(tax.get("citation_required")).lower()
        if tax_required in {"oui", "yes", "true", "1"}:
            required = True
        elif tax_required in {"non", "no", "false", "0"}:
            required = False
        else:
            required = _reference_type(reference) in citation_required_types
        if not required:
            return True
        if not citation_markers:
            return False
        text = str(page_text or "")
        # Se limiter à la ligne / puce contenant la référence : un marqueur situé
        # dans la puce précédente ne doit pas transformer un livrable en citation.
        left_candidates = [text.rfind(sep, 0, start) for sep in ("\n", "•", "▪", "\uf0b7")]
        left = max([x for x in left_candidates if x >= 0], default=max(0, start - 120))
        right_candidates = [text.find(sep, end) for sep in ("\n", "•", "▪", "\uf0b7")]
        right_valid = [x for x in right_candidates if x >= 0]
        right = min(right_valid) if right_valid else min(len(text), end + 120)
        context = text[max(0, left):min(len(text), right)]
        return any(contains(context, marker) for marker in citation_markers if clean(marker))

    def _identity(reference: str) -> Tuple[str, set[str]]:
        nref = norm(reference)
        # Le préfixe « annexe » et un éventuel numéro ne distinguent pas le titre
        # d'une pièce : ils restent affichés mais sont ignorés pour le regroupement.
        core = re.sub(r"^(?:annexe(?:s)?\s*)?(?:n[°o]?\s*)?\d{1,3}\s*[-–—:]?\s*", "", nref).strip()
        core = re.sub(r"^annexe(?:s)?\s+", "", core).strip()
        tokens = [t for t in re.findall(r"[a-z0-9]+", core) if len(t) >= 3]
        sig = {t for t in tokens if t not in generic_tokens and t not in group_ignore_tokens}
        return core, sig

    def _same_reference(a: str, b: str) -> bool:
        ca, ta = _identity(a)
        cb, tb = _identity(b)
        taxa, taxb = _taxonomy(a), _taxonomy(b)
        fa, fb = clean(taxa.get("family")), clean(taxb.get("family"))
        # Deux familles explicitement différentes dans Excel ne sont jamais
        # fusionnées uniquement parce qu'elles partagent des mots génériques.
        if fa and fb and fa != fb:
            return False
        if ca and ca == cb:
            return True
        if len(ta) >= 2 and ta == tb:
            return True
        # Deux variantes d'une même annexe numérotée peuvent employer des
        # sous-titres légèrement différents (affectation / codification...).
        # On les réunit seulement si Excel les classe dans la même famille et
        # qu'elles partagent au moins un terme sémantique significatif.
        na, nb = _annexe_number(a), _annexe_number(b)
        if na and nb and na == nb and fa and fa == fb and (ta & tb):
            return True
        # Variante qualifiée d'un même titre (ex. suffixe de rôle/phase ignoré
        # par le paramétrage) : ne fusionner que si la famille documentaire est
        # identique et qu'un intitulé est réellement le préfixe de l'autre.
        if ta and ta == tb and _reference_type(a) == _reference_type(b) and (ca.startswith(cb) or cb.startswith(ca)):
            return True
        common = len(ta & tb)
        union = len(ta | tb)
        return bool(common >= 3 and union and (common / union) >= 0.75)

    def _best_reference(refs: List[str]) -> str:
        vals = [clean(x) for x in refs if clean(x)]
        if not vals:
            return ""
        def score(value: str):
            nv = norm(value)
            starts_annexe = 1 if nv.startswith("annexe") else 0
            return (starts_annexe, len(re.findall(r"[a-z0-9]+", nv)), len(value))
        return min(vals, key=score)

    def _locations(occurrences: List[Dict[str, str]]) -> str:
        grouped: Dict[str, List[str]] = {}
        order: List[str] = []
        for occ in occurrences:
            src = clean(occ.get("source")) or "Document"
            pg = clean(occ.get("page"))
            if src not in grouped:
                grouped[src] = []
                order.append(src)
            if pg and pg not in grouped[src]:
                grouped[src].append(pg)
        parts = []
        for src in order:
            pages = grouped[src]
            parts.append(src + ((" p." + ", ".join(pages)) if pages else ""))
        return " ; ".join(parts)

    def _merge_missing_references(items: List[Dict[str, object]]) -> List[Dict[str, object]]:
        merged: List[Dict[str, object]] = []
        for item in items:
            if clean(item.get("type")) != "missing_reference":
                merged.append(item)
                continue
            occurrence = {
                "source": clean(item.get("source")),
                "page": clean(item.get("page")),
                "excerpt": clean(item.get("excerpt")),
                "reference": clean(item.get("reference")),
            }
            target = next((x for x in merged if clean(x.get("type")) == "missing_reference"
                           and _same_reference(clean(x.get("reference")), clean(item.get("reference")))), None)
            if target is None:
                clone = dict(item)
                clone["occurrences"] = [occurrence]
                merged.append(clone)
                continue
            occs = list(target.get("occurrences") or [])
            occ_key = (norm(occurrence["source"]), occurrence["page"], norm(occurrence["reference"]))
            known = {(norm(o.get("source")), clean(o.get("page")), norm(o.get("reference"))) for o in occs}
            if occ_key not in known:
                occs.append(occurrence)
            target["occurrences"] = occs
            target["reference"] = _best_reference([clean(target.get("reference")), clean(item.get("reference"))])

        for item in merged:
            if clean(item.get("type")) != "missing_reference":
                continue
            occs = list(item.get("occurrences") or [])
            if not occs:
                occs = [{
                    "source": clean(item.get("source")), "page": clean(item.get("page")),
                    "excerpt": clean(item.get("excerpt")), "reference": clean(item.get("reference")),
                }]
                item["occurrences"] = occs
            sources = []
            for occ in occs:
                src = clean(occ.get("source"))
                if src and src not in sources:
                    sources.append(src)
            item["source"] = sources[0] if len(sources) == 1 else " / ".join(sources)
            if len(sources) == 1:
                pages = []
                for occ in occs:
                    pg = clean(occ.get("page"))
                    if pg and pg not in pages:
                        pages.append(pg)
                item["page"] = ", ".join(pages)
            else:
                item["page"] = ""
            item["excerpt"] = clean(occs[0].get("excerpt")) if occs else ""
            tax = _taxonomy(clean(item.get("reference")))
            item["reference_nature"] = clean(tax.get("nature"))
            item["reference_family"] = clean(tax.get("family"))
            item["reference_family_label"] = clean(tax.get("family_label"))
            item["reference_display_group"] = clean(tax.get("display_group"))
            item["reference_display_order"] = clean(tax.get("display_order"))
            item["reference_taxonomy_rule"] = clean(tax.get("rule_id"))
            if len(occs) > 1:
                item["text"] = render_message(
                    messages, "vigilance_missing_doc_text_grouped",
                    reference=clean(item.get("reference")), locations=_locations(occs),
                )
        return merged

    def _refs(pages: Optional[List[Tuple[int, str]]], text: str, source: str):
        if not pattern:
            return
        page_seq = list(pages or []) or [(0, text or "")]
        vus = set()
        for page_no, page_text in page_seq:
            if not page_text:
                continue
            # Les PDF numériques coupent parfois un titre de pièce au milieu d'une
            # ligne (ex. « ... et ses\nannexes »). Pour la seule détection des
            # références documentaires, on ressoude uniquement les retours à la
            # ligne qui ressemblent à un simple habillage typographique entre deux
            # mots en minuscules. La longueur de la chaîne reste inchangée, donc les
            # positions de l'extrait source restent fiables.
            match_text = re.sub(
                r"(?<=[a-zà-ÿ])[ \t]*\n[ \t]*(?=[a-zà-ÿ])",
                lambda m: " " * len(m.group(0)),
                page_text,
            )
            try:
                matches = list(re.finditer(pattern, match_text))
            except re.error:
                return
            for m in matches:
                raw = clean(m.group(1) if m.lastindex else m.group(0))
                ref = _clean_reference(raw)
                if not ref:
                    continue
                tax = _taxonomy(ref)
                if clean(tax.get("treatment")).upper() in {"EXCLURE", "IGNORE", "IGNORER"}:
                    continue
                if not _context_confirms_citation(page_text, m.start(), m.end(), ref):
                    continue
                key = norm(ref)[:120]
                occurrence_key = (str(page_no or ""), key)
                if not key or occurrence_key in vus:
                    continue
                vus.add(occurrence_key)
                tokens = [t for t in re.findall(r"[a-z0-9]+", key) if len(t) >= 3]
                # Les acronymes BIM/EXE/AIM/GMAO sont suffisamment discriminants
                # même avec trois caractères. Les mots génériques de liaison ou de
                # type documentaire ne doivent pas empêcher de reconnaître qu'une
                # pièce ajoutée correspond à la référence signalée auparavant.
                identity_tokens = [
                    t for t in tokens
                    if (t not in generic_tokens and len(t) >= 4) or t in {"bim", "exe", "aim", "gmao"}
                ]
                # Pour reconnaître qu'une pièce ajoutée correspond à une
                # référence précédemment signalée, les deux premiers marqueurs
                # discriminants suffisent (ex. « BIM EXE »). Exiger quatre mots
                # faisait réapparaître le titre de la pièce comme une fausse
                # référence manquante lorsque son en-tête contenait ensuite des
                # mots descriptifs supplémentaires.
                present = bool(identity_tokens) and all(t in docs_norm for t in identity_tokens[:2])
                if present:
                    continue
                page_str = str(page_no) if page_no else "?"
                text_msg = render_message(messages, "vigilance_missing_doc_text", reference=ref, source=source, page=page_str)
                out.append({
                    "type": "missing_reference",
                    "title": render_message(messages, "vigilance_missing_doc_title"),
                    "text": text_msg,
                    "source": source,
                    "page": page_str if page_no else "",
                    "excerpt": _excerpt(page_text, m.start(), m.end()),
                    "reference": ref,
                    "reference_nature": clean(tax.get("nature")),
                    "reference_family": clean(tax.get("family")),
                    "reference_family_label": clean(tax.get("family_label")),
                    "reference_display_group": clean(tax.get("display_group")),
                    "reference_display_order": clean(tax.get("display_order")),
                    "reference_taxonomy_rule": clean(tax.get("rule_id")),
                })

    _refs(conv_pages, conv_text or "", "Convention BIM")
    _refs(cctp_pages, cctp_text or "", "CCTP")
    for source_name, pages in (additional_pages or []):
        _refs(pages, " ".join(text for _, text in (pages or [])), clean(source_name) or "Pièce complémentaire")

    # Cycle 39 : une même pièce citée à plusieurs endroits ne doit produire
    # qu'une seule vigilance. Toutes les occurrences restent tracées.
    out = _merge_missing_references(out)

    # CCAP : information documentaire uniquement. Regrouper les pages réelles et
    # leurs extraits, en excluant les lignes typiques de table des matières.
    ccap_pattern = re.compile(r"(?i)\bBIM\b|maquette\s+num[ée]rique|convention\s+BIM")
    ccap_hits: Dict[int, List[str]] = {}
    ccap_seq = list(ccap_pages or []) or ([(0, ccap_text)] if ccap_text else [])
    for page_no, page_text in ccap_seq:
        if not page_text:
            continue
        if int(page_no or 0) in (ccap_hierarchy_pages or set()):
            substantive_on_hierarchy_page = bool(re.search(
                r"(?i)maquette|\bIFC\b|\bDOE\b|\bCDE\b|plateforme|transmission[^.\n]{0,100}BIM|r[ée]union[^.\n]{0,100}BIM|[ée]l[ée]ments?\s+BIM",
                page_text,
            ))
            if not substantive_on_hierarchy_page:
                continue
        for m in ccap_pattern.finditer(page_text):
            snippet = _excerpt(page_text, m.start(), m.end())
            if not snippet or re.search(r"\.{4,}", snippet):
                continue
            # Ne pas répéter dans les vigilances la simple ligne « Convention BIM »
            # déjà affichée dans le repérage de hiérarchie contractuelle. Une clause
            # BIM substantielle située sur la même page (maquette, IFC, DOE, CDE...)
            # reste en revanche conservée.
            if int(page_no or 0) in (ccap_hierarchy_pages or set()):
                if contains(snippet, "convention BIM") and not any_kw_non_negated(
                    snippet, ["maquette", "IFC", "DOE", "CDE", "plateforme", "données BIM", "donnees BIM"]
                ):
                    continue
            pg = int(page_no or 0)
            bucket = ccap_hits.setdefault(pg, [])
            if snippet.casefold() not in {x.casefold() for x in bucket}:
                bucket.append(snippet)
    if ccap_hits:
        max_pages = max(1, p_int(params, "vigilance_ccap_mentions_max", 10))
        selected_pages = sorted(ccap_hits)[:max_pages]
        page_text = ", ".join(str(p) for p in selected_pages if p)
        excerpt_parts: List[str] = []
        for pg in selected_pages:
            for snippet in ccap_hits.get(pg, [])[:1]:
                excerpt_parts.append((f"p.{pg} — " if pg else "") + snippet)
        excerpt = clean(" | ".join(excerpt_parts))
        if len(excerpt) > excerpt_max:
            excerpt = excerpt[:max(0, excerpt_max - 3)].rstrip() + "..."
        out.append({
            "type": "ccap_info",
            "title": render_message(messages, "vigilance_ccap_title"),
            "text": render_message(messages, "vigilance_ccap_text"),
            "source": "CCAP",
            "page": page_text,
            "excerpt": excerpt,
            "reference": "",
        })
    return out



def verifier_resolution_piece_ajoutee(reference: str, display_name: str,
                                       pages: Optional[List[Tuple[int, str]]],
                                       document_rules: Optional[List[Dict[str, object]]],
                                       params: Dict[str, object]) -> bool:
    """Vérifie qu'un PDF ajouté correspond réellement à la pièce ciblée.

    Cycle 43 : le nom du fichier n'est jamais utilisé seul comme preuve d'identité.
    La stratégie et les seuils sont lus dans les règles VIGILANCE_TAXONOMIE du
    classeur Excel. Le manifeste d'ajout fournit la référence attendue, puis le
    moteur contrôle sa compatibilité avec le titre/contenu des premières pages.
    """
    expected = clean(reference)
    if not expected:
        return False
    nref = norm(expected)
    matching_rule: Dict[str, object] = {}
    rules = [
        row for row in (document_rules or [])
        if clean(row.get("famille_regle")).upper() == "VIGILANCE_TAXONOMIE"
        and clean(row.get("actif", "Oui")).lower() not in {"non", "false", "0"}
    ]
    def _ord(row):
        try:
            return int(float(row.get("ordre", 9999)))
        except (TypeError, ValueError):
            return 9999
    for row in sorted(rules, key=_ord):
        pat = clean(row.get("vigilance_pattern"))
        if not pat:
            continue
        try:
            if re.search(pat, nref, flags=re.I):
                matching_rule = row
                break
        except re.error:
            continue
    if not matching_rule:
        return False
    mode = clean(matching_rule.get("vigilance_resolution_mode")).upper()
    if mode in {"", "NON_APPLICABLE", "EXCLURE", "IGNORE", "IGNORER"}:
        return False
    if mode != "MANIFESTE_CONTENU":
        return False

    pages_max = max(1, p_int(params, "vigilance_resolution_pages_max", 8))
    content_parts = [clean(display_name)]
    for _pg, txt in list(pages or [])[:pages_max]:
        if clean(txt):
            content_parts.append(clean(txt))
    corpus = norm("\n".join(content_parts))
    if not corpus:
        return False

    generic_raw = params.get("vigilance_resolution_mots_generiques", []) or []
    generic = {norm(x) for x in (split_kw(generic_raw) if isinstance(generic_raw, str) else generic_raw) if clean(x)}
    acr_raw = params.get("vigilance_resolution_acronymes_discriminants", []) or []
    acronyms = {norm(x) for x in (split_kw(acr_raw) if isinstance(acr_raw, str) else acr_raw) if clean(x)}
    tokens = [t for t in re.findall(r"[a-z0-9]+", nref) if len(t) >= 3]
    significant = []
    for token in tokens:
        if token in acronyms or token not in generic:
            if token not in significant:
                significant.append(token)
    if not significant:
        return False

    try:
        min_tokens = max(1, int(float(matching_rule.get("vigilance_resolution_min_tokens") or 2)))
    except (TypeError, ValueError):
        min_tokens = 2
    try:
        ratio_min = float(matching_rule.get("vigilance_resolution_ratio_min") or 0.5)
    except (TypeError, ValueError):
        ratio_min = 0.5
    matched = sum(1 for token in significant if re.search(r"(?<![a-z0-9])" + re.escape(token) + r"(?![a-z0-9])", corpus))
    token_match = matched >= min_tokens and (matched / max(1, len(significant))) >= ratio_min

    # Pour une annexe numérotée, le numéro est un marqueur d'identité très fort.
    # Il complète le noyau lexical mais ne suffit jamais seul.
    ann = re.search(r"\bannexe(?:s)?\s*(?:n[°o]?\s*)?(\d{1,3})\b", nref, flags=re.I)
    if ann:
        same_annex = bool(re.search(r"\bannexe(?:s)?\s*(?:n[°o]?\s*)?" + re.escape(ann.group(1)) + r"\b", corpus, flags=re.I))
        return bool(same_annex and token_match)
    return bool(token_match)

def qualifier_vigilances_pour_ajout(points: List[Dict[str, object]], results: Dict[str, Dict],
                                      params: Dict[str, object]) -> List[Dict[str, object]]:
    """Marque uniquement les pièces manquantes reliées à un bloc encore incertain.

    Le moteur ne décide pas qu'une annexe est « importante » par son nom. Il vérifie
    si la référence documentaire est effectivement reprise dans le contenu public
    d'un bloc dont le statut est paramétré comme réévaluable.
    """
    allowed_raw = params.get("vigilance_ajout_statuts", []) or []
    allowed = {clean(x).upper() for x in (split_kw(allowed_raw) if isinstance(allowed_raw, str) else allowed_raw)}
    sources_raw = params.get("vigilance_ajout_sources", []) or []
    allowed_sources = {clean(x).upper() for x in (split_kw(sources_raw) if isinstance(sources_raw, str) else sources_raw)}
    min_tokens = max(1, p_int(params, "vigilance_ajout_min_tokens", 2))
    generic = {"cahier", "charges", "annexe", "annexes", "convention", "document", "documents", "bim"}
    for point in points or []:
        point["upload_recommended"] = False
        point["related_blocks"] = []
        if clean(point.get("type")) != "missing_reference":
            continue
        occurrences = list(point.get("occurrences") or [])
        if not occurrences:
            occurrences = [{
                "source": clean(point.get("source")), "page": clean(point.get("page")),
                "excerpt": clean(point.get("excerpt")), "reference": clean(point.get("reference")),
            }]
        eligible_occurrences = [
            occ for occ in occurrences
            if not allowed_sources or clean(occ.get("source")).upper() in allowed_sources
        ]
        if allowed_sources and not eligible_occurrences:
            continue
        reference = clean(point.get("reference"))
        ref_norm = norm(reference)
        ref_tokens = [t for t in re.findall(r"[a-z0-9]+", ref_norm) if len(t) >= 3 and t not in generic]
        if not reference:
            continue
        related = []
        for bid, res in results.items():
            status = clean(res.get("public_status_key")).upper()
            if allowed and status not in allowed:
                continue
            hay = " ".join(clean(res.get(k)) for k in (
                "public_demand", "public_question", "conclusion", "conv_pf", "cctp_pf", "specific_demand"
            ))
            hay_norm = norm(hay)
            exact = bool(ref_norm and ref_norm in hay_norm)
            common = sum(1 for t in ref_tokens if t in hay_norm)
            strong_overlap = bool(
                len(ref_tokens) >= min_tokens
                and common >= min_tokens
                and (common / max(1, len(ref_tokens))) >= 0.7
            )
            if exact or strong_overlap:
                related.append(clean(bid))
        if related:
            point["upload_recommended"] = True
            point["related_blocks"] = related
            primary = (eligible_occurrences or occurrences)[0]
            point["upload_source"] = clean(primary.get("source"))
            point["upload_page"] = clean(primary.get("page"))
            point["upload_excerpt"] = clean(primary.get("excerpt"))
    return points

def appliquer_securite_cctp_non_lisible(results: Dict[str, Dict], qualite_cctp: Dict[str, object],
                                           params: Dict[str, object], messages: Dict[str, str]) -> None:
    """Empêche un CCTP image/non OCR de produire de faux "non démontré".

    La règle est purement documentaire et générique : si le CCTP a été fourni
    mais ne peut pas être lu, seuls les blocs déjà confirmés par une autre preuve
    ou explicitement exclus conservent leur décision. Les autres passent dans un
    état de clarification dont les codes sont configurés dans Excel.
    """
    if clean((qualite_cctp or {}).get("statut", "")).upper() != "NON_EXPLOITABLE":
        return

    code_applicable = p_text(params, "applicabilite_si_cctp_confirme")
    code_non_applicable = p_text(params, "applicabilite_si_cctp_exclu")
    safe_app = p_text(params, "pdf_cctp_applicabilite_non_lisible")
    safe_status = p_text(params, "pdf_cctp_statut_public_non_lisible")
    safe_decision = p_text(params, "pdf_cctp_decision_non_lisible")
    safe_response = p_text(params, "pdf_cctp_response_mode_non_lisible")
    safe_action = p_text(params, "pdf_cctp_action_mode_non_lisible")
    safe_question = p_text(params, "pdf_cctp_question_mode_non_lisible")
    safe_demand = render_message(messages, "pdf_cctp_unreadable_block_demand")

    for result in (results or {}).values():
        if bool(result.get("exclusion_explicit")) or bool(result.get("interblock_exclusion")):
            continue
        current_app = clean(result.get("applicabilite_lot", ""))
        # Une exigence déjà applicable grâce à une preuve indépendante du CCTP
        # reste confirmée. Une exclusion déjà démontrée reste également inchangée.
        if current_app in {code_applicable, code_non_applicable}:
            continue
        if safe_app:
            result["applicabilite_lot"] = safe_app
        if safe_status:
            result["statut_public_override"] = safe_status
        if safe_decision:
            result["decision_override"] = safe_decision
        if safe_response:
            result["response_mode_override"] = safe_response
        if safe_action:
            result["action_mode_override"] = safe_action
        if safe_question:
            result["question_mode_override"] = safe_question
        if safe_demand:
            result["specific_demand"] = safe_demand


def _extraire_infos_document(pages_textes: list, detections: Optional[pd.DataFrame] = None,
                              metadata_rules: Optional[pd.DataFrame] = None) -> dict:
    """Extrait les métadonnées uniquement avec les règles de 36_Extraction_Metadata."""
    champs = [
        "indice", "phase", "date", "emetteur", "maitre_ouvrage", "maitrise_oeuvre",
        "entreprise_generale", "bim_manager", "coordinateur_bim", "numero_affaire", "operation", "nom_projet", "lot_cctp",
    ]
    res = {"titre": "", **{c: "" for c in champs}}
    if not pages_textes:
        return res
    texte_global = "\n".join(pages_textes)
    res.update(_extraction_metadata_generique(texte_global, metadata_rules, champs))
    res["titre"] = clean(res.get("operation", ""))
    return res


def _localiser_page(pages: Optional[List[Tuple[int, str]]], variants: List[str]) -> Tuple[str, str]:
    """
    Retourne (numéro de page, extrait ~140 caractères) où l'un des `variants`
    apparaît en premier dans `pages` (liste de tuples (num_page, texte)).
    Vide si `pages` est absent ou si rien n'est trouvé -- sert à tracer la
    source d'une détection globale (ex: plateforme CDE) pour affichage/citation.

    Ignore les occurrences situées dans un sommaire/table des matières
    (motif de points de suite "....... 12" typique d'un sommaire) et continue
    la recherche sur l'occurrence suivante -- une citation de sommaire n'a
    aucune valeur probante, contrairement à la vraie section du document.
    """
    if not pages:
        return "", ""
    for num, txt in pages:
        low = txt.lower()
        pos = 0
        while True:
            i = -1
            kw_matched = None
            for kw in variants:
                j = trouve_position(low, kw, pos)
                if j != -1 and (i == -1 or j < i):
                    i, kw_matched = j, kw
            if i == -1:
                break
            context_chars = p_int(ENGINE_PARAMS_GLOBAL, "localisation_contexte_caracteres", 0)
            start = max(0, i - context_chars)
            extrait = txt[start:i + len(kw_matched) + context_chars].strip()
            if re.search(r"\.{4,}", extrait):  # ligne de sommaire (points de suite)
                pos = i + len(kw_matched)
                continue
            return str(num), extrait
    return "", ""


def _plateformes_depuis_glossaire(gloss: Optional[pd.DataFrame]) -> List[Tuple[List[str], str]]:
    """
    Construit la liste des plateformes CDE à détecter depuis 40_Glossaire
    (lignes où categorie == "Plateforme CDE"), au lieu d'une liste codée en
    dur dans le script. Ajouter/renommer une plateforme se fait uniquement
    dans l'Excel -- aucune modification de code nécessaire.
    Retourne [(variantes_lowercase, nom_affichage), ...] dans l'ordre du
    glossaire (le premier terme dont une variante matche le texte l'emporte).
    """
    if gloss is None or gloss.empty or "categorie" not in gloss.columns:
        return []
    out = []
    for _, row in gloss.iterrows():
        if clean(row.get("categorie", "")).lower() != "plateforme cde":
            continue
        terme = clean(row.get("terme", ""))
        variantes = [v.lower() for v in split_kw(row.get("mots_cles", "")) if v]
        if terme and variantes:
            out.append((variantes, terme))
    return out


def _appliquer_detections(ctx: Dict, txt: str, conv_pages, detections: Optional[pd.DataFrame]) -> None:
    """
    Applique les règles génériques de 35_Detections_Convention à ctx, en
    mutant ctx en place. Trois types de règles, aucune liste codée en dur :
      - "valeur"  : le premier mot-clé trouvé (dans l'ordre des lignes Excel)
                    fixe ctx[cle_ctx] à valeur_affichee, avec page/extrait.
      - "booleen" : si un des mots-clés apparaît, ctx[cle_ctx] = True.
      - "liste"   : chaque mot-clé trouvé ajoute valeur_affichee à une liste
                    cumulée (dédoublonnée, ordre conservé), jointe par ", ".
    Ajouter une valeur (nouveau logiciel, nouvelle fréquence...) se fait
    uniquement dans l'Excel.
    """
    if detections is None or detections.empty:
        return

    listes: Dict[str, List[str]] = {}
    for _, row in detections.iterrows():
        cle = clean(row.get("cle_ctx", ""))
        typ = clean(row.get("type", "")).lower()
        valeur_aff = clean(row.get("valeur_affichee", ""))
        mots = split_kw(row.get("mots_cles", ""))
        if not cle or not mots:
            continue

        if typ == "valeur":
            if ctx.get(cle):  # déjà trouvé par une ligne précédente (priorité à l'ordre Excel)
                continue
            for m in mots:
                if contains(txt, m):
                    ctx[cle] = valeur_aff
                    pg, ex = _localiser_page(conv_pages, [m.lower()])
                    ctx[f"{cle}_page"] = pg
                    ctx[f"{cle}_extrait"] = ex
                    break

        elif typ == "booleen":
            if ctx.get(cle):
                continue
            if any(contains(txt, m) for m in mots):
                ctx[cle] = True

        elif typ == "liste":
            if any(contains(txt, m) for m in mots):
                listes.setdefault(cle, [])
                if valeur_aff not in listes[cle]:
                    listes[cle].append(valeur_aff)

    for cle, valeurs in listes.items():
        ctx[cle] = ", ".join(valeurs)



def _fusionner_formats_declares(formats_existants: str, liste_brute: str, format_ifc: str = "") -> str:
    """Ajoute les formats explicitement listés mais absents du référentiel nominatif.

    Le moteur ne devine pas des formats dans la prose : ce repli ne travaille que
    sur une liste déjà capturée par une règle Excel de 36_Extraction_Metadata.
    Les jetons sont dédoublonnés sans connaître de format métier particulier.
    Les variantes IFC restent exposées dans le champ dédié ``format_ifc``.
    """
    existants = [clean(v) for v in str(formats_existants or "").split(",") if clean(v)]
    if not liste_brute:
        return ", ".join(existants)
    texte = re.sub(r"(?i)\b(?:et|and)\b", ";", str(liste_brute))
    candidats = [clean(v).strip(" .") for v in re.split(r"[,;]", texte)]
    vus = {norm(v) for v in existants}
    for brut in candidats:
        if not brut:
            continue
        # Une liste de formats doit contenir des jetons courts (extension, sigle
        # ou forme composée XLS/XLSX), pas une phrase libre.
        if not re.fullmatch(r"\.?[A-Za-z0-9][A-Za-z0-9._+\-/]{0,19}", brut):
            continue
        valeur = brut.lstrip(".").upper()
        if valeur.startswith("IFC") and format_ifc:
            continue
        cle = norm(valeur)
        if cle not in vus:
            existants.append(valeur)
            vus.add(cle)
    return ", ".join(existants)


def _extraire_details_convention(bid: str, conv_pf: str, conv_pg: str,
                                  cctp_pf: str, cctp_pg: str,
                                  conv_text: str, lot: str,
                                  conv_pages: Optional[List[Tuple[int, str]]] = None,
                                  gloss: Optional[pd.DataFrame] = None,
                                  detections: Optional[pd.DataFrame] = None,
                                  metadata_rules: Optional[pd.DataFrame] = None) -> Dict:
    """Construit le contexte à partir des seules tables de paramétrage Excel."""
    ctx: Dict = {
        "plateforme": "", "plateforme_page": "", "plateforme_extrait": "",
        "cout_plateforme": "", "logiciel": "", "logiciel_page": "", "logiciel_extrait": "",
        "format_ifc": "", "format_ifc_page": "", "format_ifc_extrait": "",
        "lod": "", "lod_page": "", "lod_extrait": "", "nd": "", "nd_page": "", "nd_extrait": "",
        "frequence": "", "section_cctp": "", "niveau_bim": "", "niveau_bim_page": "", "niveau_bim_extrait": "",
        "dim_4d": False, "dim_5d": False, "dim_6d": False, "dim_7d": False,
        "bim_manager": "", "bim_manager_page": "", "bim_manager_extrait": "",
        "coordinateur_bim": "", "coordinateur_bim_page": "", "coordinateur_bim_extrait": "",
        "geo_referencement": False, "clash_3d": False, "doe_numerique": False,
        "formats_livrables": "", "contractuel": False, "indice_doc": "", "date_doc": "",
        "emetteur_doc": "", "maitre_ouvrage_doc": "", "numero_affaire": "", "titre_doc": "",
        "conv_extrait": conv_pf[:p_int(ENGINE_PARAMS_GLOBAL, "contexte_bloc_extrait_longueur_max", len(conv_pf))] if conv_pf else "", "conv_page": conv_pg,
        "cctp_extrait": cctp_pf[:p_int(ENGINE_PARAMS_GLOBAL, "contexte_bloc_extrait_longueur_max", len(cctp_pf))] if cctp_pf else "", "cctp_page": cctp_pg,
    }
    txt = conv_text.lower()

    # Les solutions CDE nommées viennent de 40_Glossaire ; les replis génériques,
    # logiciels, formats, LOD/ND/niveaux et autres signaux viennent de 35_Detections_Convention.
    plateforme_depuis_glossaire = False
    for variantes, nom_affichage in _plateformes_depuis_glossaire(gloss):
        matched = next((v for v in variantes if contains(txt, v)), None)
        if matched:
            ctx["plateforme"] = nom_affichage
            ctx["plateforme_page"], ctx["plateforme_extrait"] = _localiser_page(conv_pages, [matched])
            plateforme_depuis_glossaire = True
            break
    _appliquer_detections(ctx, txt, conv_pages, detections)

    # Les extractions regex sont elles aussi configurées dans 36_Extraction_Metadata.
    # Les quatre champs de repli ci-dessous permettent de signaler une valeur
    # explicitement nommée mais absente des listes nominatives de l'Excel.
    extra = _extraction_metadata_generique(
        conv_text, metadata_rules,
        ["cout_plateforme", "bim_manager", "coordinateur_bim",
         "plateforme", "logiciel", "formats_livrables_raw", "format_ifc"]
    )
    for key in ("cout_plateforme", "bim_manager", "coordinateur_bim"):
        value = clean(extra.get(key, ""))
        if value:
            ctx[key] = value
            if key in {"bim_manager", "coordinateur_bim"}:
                pg, excerpt = _localiser_page(conv_pages, [value.lower()])
                ctx[f"{key}_page"] = pg
                ctx[f"{key}_extrait"] = excerpt

    # Une plateforme explicitement nommée prime sur le repli générique
    # « plateforme collaborative du projet », mais pas sur une plateforme déjà
    # identifiée par le glossaire (qui apporte son libellé canonique).
    plateforme_extra = clean(extra.get("plateforme", ""))
    if plateforme_extra and not plateforme_depuis_glossaire:
        ctx["plateforme"] = plateforme_extra
        ctx["plateforme_page"], ctx["plateforme_extrait"] = _localiser_page(
            conv_pages, [plateforme_extra.lower()]
        )

    logiciel_extra = clean(extra.get("logiciel", ""))
    if logiciel_extra:
        ctx["logiciel"] = logiciel_extra
        ctx["logiciel_page"], ctx["logiciel_extrait"] = _localiser_page(
            conv_pages, [logiciel_extra.lower()]
        )

    format_ifc_extra = clean(extra.get("format_ifc", ""))
    if format_ifc_extra:
        format_ifc_actuel = clean(ctx.get("format_ifc", ""))
        compact_extra = re.sub(r"[^a-z0-9]", "", norm(format_ifc_extra))
        compact_actuel = re.sub(r"[^a-z0-9]", "", norm(format_ifc_actuel))
        # Le motif générique ne remplace une valeur déjà connue que s'il est
        # strictement plus précis (ex. IFC 4.3 au lieu du match partiel IFC4).
        plus_precis = bool(compact_actuel and compact_extra.startswith(compact_actuel)
                           and len(compact_extra) > len(compact_actuel))
        if not format_ifc_actuel or plus_precis:
            ctx["format_ifc"] = format_ifc_extra
            ctx["format_ifc_page"], ctx["format_ifc_extrait"] = _localiser_page(
                conv_pages, [format_ifc_extra.lower()]
            )

    ctx["formats_livrables"] = _fusionner_formats_declares(
        ctx.get("formats_livrables", ""),
        clean(extra.get("formats_livrables_raw", "")),
        ctx.get("format_ifc", ""),
    )

    if cctp_pf:
        sec = _extraction_metadata_generique(cctp_pf, metadata_rules, ["section_cctp"])
        if sec.get("section_cctp"):
            ctx["section_cctp"] = sec["section_cctp"]
    return ctx







# ═══════════════════════════════════════════════════════════════════════════════
# QUESTIONNAIRE INTERACTIF CONTEXTUALISÉ
# ═══════════════════════════════════════════════════════════════════════════════


class _SafeCtxDict(dict):
    """Laisse une {variable} inconnue telle quelle plutôt que de lever une erreur."""
    def __missing__(self, key):
        return "{" + key + "}"


def _sub_ctx(txt: str, ctx_full: Dict, messages: Optional[Dict[str, str]] = None) -> str:
    """Substitue les variables des textes Excel ; les valeurs neutres viennent de 12_Messages_Moteur."""
    if not txt or "{" not in txt:
        return txt
    messages = messages or {}
    defaults = {
        "plateforme": render_message(messages, "ctx_default_plateforme"),
        "cout": ctx_full.get("cout_plateforme") or render_message(messages, "ctx_default_cout"),
        "cout_plateforme": ctx_full.get("cout_plateforme") or render_message(messages, "ctx_default_cout"),
        "logiciel": render_message(messages, "ctx_default_logiciel"),
        "fmt_ifc": ctx_full.get("format_ifc") or render_message(messages, "ctx_default_format_ifc"),
        "format_ifc": ctx_full.get("format_ifc") or render_message(messages, "ctx_default_format_ifc"),
        "lod": render_message(messages, "ctx_default_lod"),
        "nd": ctx_full.get("nd") or render_message(messages, "ctx_default_nd"),
        "niveau_bim": render_message(messages, "ctx_default_niveau_bim"),
        "frequence": render_message(messages, "ctx_default_frequence"),
        "ref_conv": (f"p. {ctx_full['conv_page']}" if ctx_full.get("conv_page") else render_message(messages, "ctx_default_ref_convention")),
        "ref_cctp": (f"p. {ctx_full['cctp_page']}" if ctx_full.get("cctp_page") else render_message(messages, "ctx_default_ref_cctp")),
        "ref_plateforme": (f"p. {ctx_full['plateforme_page']}" if ctx_full.get("plateforme_page") else render_message(messages, "ctx_default_ref_plateforme")),
        "lot": "",
    }
    values = {**defaults, **{k: v for k, v in ctx_full.items() if v not in (None, "")}}
    try:
        return txt.format_map(_SafeCtxDict(values))
    except Exception:
        return txt



def poser_questionnaire(blocs: pd.DataFrame, lot: str,
                        resultats_lecture: Dict,
                        questions: pd.DataFrame,
                        axes_config: Dict,
                        messages: Dict[str, str],
                        gloss: pd.DataFrame = None,
                        detections: pd.DataFrame = None,
                        metadata_rules: pd.DataFrame = None,
                        reponses_fournies: Optional[Dict[str, int]] = None) -> Dict:
    """Évalue tous les blocs disposant de questions actives dans 50_Questions."""
    if questions is None or questions.empty:
        return {}
    block_ids = [clean(v) for v in questions["id_bloc"].astype(str).tolist() if clean(v)]
    block_ids = sorted(set(block_ids), key=block_sort_key)
    titles = {clean(b["id_bloc"]): clean(b.get("titre_bloc", "")) for _, b in blocs.iterrows()}
    results = {}
    for bid in block_ids:
        if bid not in resultats_lecture:
            continue
        res = resultats_lecture[bid]
        ctx = _extraire_details_convention(
            bid, res.get("conv_pf", ""), res.get("conv_pg", ""),
            res.get("cctp_pf", ""), res.get("cctp_pg", ""),
            res.get("_conv_text", ""), lot,
            gloss=gloss, detections=detections, metadata_rules=metadata_rules,
        )
        ctx_sub = {**ctx, "lot": lot}
        rows_q = questions[questions["id_bloc"].astype(str).str.strip() == bid]
        score = score_max = 0
        answers = []
        for _, row in rows_q.iterrows():
            if clean(row.get("actif", "Oui")).casefold() in {"non", "false", "0"}:
                continue
            qid = clean(row.get("id_question", ""))
            if not qid:
                continue
            qtxt = _sub_ctx(str(row.get("question", "")), ctx_sub, messages)
            labels = [
                _sub_ctx(str(row.get("indice_0", "")), ctx_sub, messages),
                _sub_ctx(str(row.get("indice_1", "")), ctx_sub, messages),
                _sub_ctx(str(row.get("indice_2", "")), ctx_sub, messages),
            ]
            try:
                weight = int(float(row.get("poids", 1) or 1))
            except (TypeError, ValueError):
                weight = 1
            if reponses_fournies is not None:
                if qid not in reponses_fournies:
                    raise ValueError(f"Réponse manquante pour la question {qid}")
                val = int(reponses_fournies[qid])
                if val not in (0, 1, 2):
                    raise ValueError(f"Réponse invalide pour {qid} : 0, 1 ou 2 attendu")
            else:
                print(qid + ". " + qtxt)
                for idx, label in enumerate(labels):
                    print(f"  {idx} - {label}")
                while True:
                    raw = input("Votre réponse [0/1/2] : ").strip()
                    if raw in {"0", "1", "2"}:
                        val = int(raw)
                        break
            score += val * weight
            score_max += 2 * weight
            action_capacite = ""
            if val == 0:
                action_capacite = _sub_ctx(str(row.get("action_si_non", "")), ctx_sub, messages)
            elif val == 1:
                action_capacite = _sub_ctx(str(row.get("action_si_partiel", "")), ctx_sub, messages)
            answers.append({
                "id_question": qid,
                "question": qtxt,
                "reponse_val": val,
                "reponse_label": labels[val],
                "poids": weight,
                "action_capacite": action_capacite,
            })
        if not answers:
            continue
        pct = round(score / score_max * 100) if score_max else 0
        code = capacity_code_for_pct(axes_config, pct)
        label = clean(((axes_config.get("capacite") or {}).get(code) or {}).get("label")) or code
        results[bid] = {
            "titre": titles.get(bid, bid), "score": score, "max": score_max,
            "pct": pct, "capacite_code": code, "niveau": label, "reponses": answers,
        }
    return results


# ═══════════════════════════════════════════════════════════════════════════════
# DIAGRAMME RADAR
# ═══════════════════════════════════════════════════════════════════════════════
def flatten_answers(scores_eval: Dict) -> Dict[str, int]:
    """{id_question: reponse_val} sur tous les blocs, pour le moteur de cohérence."""
    out: Dict[str, int] = {}
    for bloc in scores_eval.values():
        for r in bloc.get("reponses", []):
            qid = r.get("id_question", "")
            if qid:
                out[qid] = r.get("reponse_val", 0)
    return out




def _regle_reponse_ok(operateur: str, valeur_reelle: int, valeur_attendue: int) -> bool:
    op = clean(operateur)
    return {"=": valeur_reelle == valeur_attendue, "<=": valeur_reelle <= valeur_attendue,
            ">=": valeur_reelle >= valeur_attendue, "<": valeur_reelle < valeur_attendue,
            ">": valeur_reelle > valeur_attendue}.get(op, False)


def detecter_contradictions(coherence: Optional[pd.DataFrame], scores_eval: Dict,
                             resultats_lecture: Optional[Dict] = None) -> List[Dict]:
    """Applique les comparaisons question-vers-question définies dans 55_Regles_Coherence."""
    if coherence is None or coherence.empty:
        return []
    responses = flatten_answers(scores_eval)
    found: List[Dict] = []
    for _, rule in coherence.iterrows():
        if clean(rule.get("actif", "Oui")).casefold() in {"non", "false", "0"}:
            continue
        q1 = clean(rule.get("id_question_1", ""))
        q2 = clean(rule.get("id_question_2", ""))
        if q1 not in responses or q2 not in responses:
            continue
        try:
            v1 = int(float(rule.get("valeur_1", 0)))
            v2 = int(float(rule.get("valeur_2", 0)))
        except (TypeError, ValueError):
            continue
        c1 = _regle_reponse_ok(rule.get("operateur_1", ""), responses[q1], v1)
        c2 = _regle_reponse_ok(rule.get("operateur_2", ""), responses[q2], v2)
        if not (c1 and c2):
            continue
        item = {
            "id_regle": clean(rule.get("id_regle", "")),
            "gravite": clean(rule.get("gravite", "")),
            "penalite": int(float(rule.get("penalite", 0) or 0)),
            "message": clean(rule.get("message", "")),
            "id_question_1": q1,
            "id_question_2": q2,
        }
        found.append(item)
        if resultats_lecture is not None:
            for qid in (q1, q2):
                bid = qid.split("_Q", 1)[0]
                if bid in resultats_lecture:
                    bucket = resultats_lecture[bid].setdefault("contradictions", [])
                    if item not in bucket:
                        bucket.append(item)
    return found





def p_list(params: Dict[str, object], cle: str) -> List[str]:
    v = params.get(cle, [])
    return v if isinstance(v, list) else split_kw(v)



# ═══════════════════════════════════════════════════════════════════════════════
# SÉLECTION TEXTE DE RÉPONSE
# ═══════════════════════════════════════════════════════════════════════════════
def get_ligne_reponse(textes: pd.DataFrame, bid: str, capacite_code: str) -> Optional[pd.Series]:
    """Sélectionne 60_Reponses par code de capacité ; aucun seuil n'est recalculé ici."""
    cands = textes[textes["id_bloc"].astype(str).str.strip() == bid]
    if "capacite_code" in cands.columns:
        exact = cands[cands["capacite_code"].astype(str).str.strip() == clean(capacite_code)]
        if not exact.empty:
            return exact.iloc[0]
    return None



def _evaluer_condition_simple(expr: str, valeurs: Dict[str, int]) -> bool:
    """
    Évalue une condition simple "ID_QUESTION>=N" (ou <=, =, <, >) contre les
    réponses réelles du questionnaire. Vide = toujours vraie (aucune
    condition supplémentaire). Pas d'eval() arbitraire, juste un format
    minimal comparaison-unique, suffisant pour bloquer un texte trop
    affirmatif tant qu'une question précise n'a pas la réponse attendue.
    """
    expr = clean(expr)
    if not expr:
        return True
    m = re.match(r"([A-Za-z0-9_]+)\s*(>=|<=|=|<|>)\s*(-?\d+)", expr)
    if not m:
        return True
    qid, op, val = m.group(1), m.group(2), int(m.group(3))
    reel = valeurs.get(qid)
    if reel is None:
        return False
    return {"=": reel == val, "<=": reel <= val, ">=": reel >= val,
            "<": reel < val, ">": reel > val}.get(op, True)





def preparer_textes_reponse(
    textes: pd.DataFrame,
    scores_eval: Dict,
    resultats_lecture: Dict,
    entreprise: str,
    lot: str,
    messages: Dict[str, str],
    engine_params: Dict[str, object],
) -> Dict[str, Dict]:
    """Prépare les formulations 60_Reponses sans classification codée dans Python."""
    if not scores_eval:
        return {}
    allowed = set(p_list(engine_params, "applicabilites_generation"))
    response_values = flatten_answers(scores_eval)
    output: Dict[str, Dict] = {}
    default_engagement = p_text(engine_params, "niveau_engagement_defaut")
    presence_abs = p_text(engine_params, "presence_code_absente")
    for bid, score in scores_eval.items():
        result = resultats_lecture.get(bid, {})
        if result.get("presence_contractuelle") == presence_abs:
            output[bid] = {"texte": "", "niveau": "", "statut_generation": "", "avertissement": render_message(messages, "paragraphe_presence_absente")}
            continue
        if allowed and result.get("applicabilite_lot") not in allowed:
            output[bid] = {"texte": "", "niveau": "", "statut_generation": "", "avertissement": render_message(messages, "paragraphe_bloque")}
            continue
        if result.get("contradictions"):
            output[bid] = {"texte": "", "niveau": "", "statut_generation": "", "avertissement": render_message(messages, "paragraphe_contradiction")}
            continue
        code = clean(score.get("capacite_code"))
        row = get_ligne_reponse(textes, bid, code)
        if row is None:
            continue
        cond_ok = _evaluer_condition_simple(row.get("conditions_questions", ""), response_values)
        engagement = clean(row.get("niveau_engagement", "")) or default_engagement
        ctx = {"entreprise": entreprise, "lot": lot}
        text = _sub_ctx(clean(row.get("texte_reponse", "")), ctx, messages)
        if cond_ok:
            output[bid] = {"texte": text, "niveau": clean(row.get("niveau_label", "")), "statut_generation": engagement, "conditions_ok": True}
        else:
            fallback = _sub_ctx(clean(row.get("texte_secours", "")), ctx, messages)
            output[bid] = {
                "texte": fallback, "niveau": clean(row.get("niveau_label", "")),
                "statut_generation": engagement if fallback else "",
                "conditions_ok": False, "avertissement": render_message(messages, "paragraphe_bloque"),
            }
    return output


def extraire_priorite_pieces_ccap(path: Optional[Path], rules: List[Dict[str, object]],
                                   params: Dict[str, object], messages: Dict[str, str]) -> Dict[str, object]:
    """Repère un ordre de pièces uniquement avec 37_CCAP_Reperage et 05_Parametres_Moteur."""
    if not path:
        return {"provided": False, "status": "not_provided", "items": [], "message": render_message(messages, "ccap_non_fourni")}
    if not path.exists():
        return {"provided": False, "status": "not_provided", "items": [], "message": render_message(messages, "ccap_non_fourni")}
    doc = fitz.open(str(path))
    configured_max_pages = p_int(params, "ccap_max_pages", 0)
    max_pages = len(doc) if configured_max_pages <= 0 else min(len(doc), configured_max_pages)
    lines = []
    tables = []
    for page_idx in range(max_pages):
        page = doc[page_idx]
        for raw in page.get_text().splitlines():
            text = clean(raw)
            if text:
                lines.append({"page": page_idx + 1, "text": text})
        # Les CCAP présentent parfois la hiérarchie sous forme de tableau :
        # PyMuPDF permet de conserver les colonnes (rang / pièce / observation)
        # que la lecture ligne à ligne mélange. L'interprétation des en-têtes
        # reste entièrement pilotée par 37_CCAP_Reperage.
        try:
            finder = page.find_tables()
            for table in getattr(finder, "tables", []) or []:
                rows = table.extract() or []
                if rows:
                    tables.append({"page": page_idx + 1, "rows": rows})
        except Exception:
            # La détection de tableaux est un enrichissement : le parseur
            # historique ligne à ligne reste le repli sans échec bloquant.
            pass
    doc.close()
    active = [r for r in rules if clean(r.get("actif", "Oui")).casefold() not in {"non", "false", "0"}]
    def by_type(kind):
        rows = [r for r in active if clean(r.get("type_regle", "")).upper() == kind]
        def order(r):
            try: return int(float(r.get("priorite", 9999)))
            except (TypeError, ValueError): return 9999
        return sorted(rows, key=order)
    def matches(text, rule):
        mode = clean(rule.get("mode", "contient")).lower()
        motif = str(rule.get("motif", "") or "")
        if mode == "regex":
            try: return re.search(motif, text, flags=re.IGNORECASE)
            except re.error: return None
        return True if contains(text, motif) else None
    anchors = by_type("ANCRE")
    priority_rules = by_type("MARQUEUR_PRIORITE")
    item_rules = by_type("ITEM")
    stop_rules = by_type("STOP")
    ignore_rules = by_type("IGNORER")
    table_rank_header_rules = by_type("TABLE_RANK_HEADER")
    table_item_header_rules = by_type("TABLE_ITEM_HEADER")
    table_rank_value_rules = by_type("TABLE_RANK_VALUE")
    max_after = p_int(params, "ccap_max_lignes_apres_ancre", 0)
    min_items = p_int(params, "ccap_min_pieces", 0)
    max_items = p_int(params, "ccap_max_pieces", 0)
    require_marker = p_text(params, "ccap_exiger_marqueur_priorite").casefold() in {"oui", "true", "1", "yes"}
    anchor_indices = [i for i, entry in enumerate(lines) if any(matches(entry["text"], r) for r in anchors)]
    if not anchor_indices:
        return {"provided": True, "status": "not_found", "items": [], "message": render_message(messages, "ccap_ordre_non_identifie")}

    marker_seen = False
    first_anchor_page = lines[anchor_indices[0]]["page"]
    for anchor_idx in anchor_indices:
        end_idx = min(len(lines), anchor_idx + 1 + max_after) if max_after else len(lines)
        window = lines[anchor_idx:end_idx]
        marker_found = any(any(matches(entry["text"], r) for r in priority_rules) for entry in window)
        marker_seen = marker_seen or marker_found
        if require_marker and not marker_found:
            continue

        # Priorité aux tableaux structurés lorsque le CCAP contient explicitement
        # des colonnes de rang et de pièce. Les noms de colonnes et le format du
        # rang sont paramétrés dans 37_CCAP_Reperage : aucune discipline, aucun
        # projet et aucun libellé de pièce n'est codé ici.
        if table_rank_header_rules and table_item_header_rules and table_rank_value_rules:
            window_pages = {entry["page"] for entry in window}
            for table_info in tables:
                if table_info.get("page") not in window_pages:
                    continue
                rows = table_info.get("rows") or []
                if len(rows) < 2:
                    continue
                header = [clean(cell) for cell in (rows[0] or [])]
                rank_col = next((idx for idx, cell in enumerate(header) if any(matches(cell, r) for r in table_rank_header_rules)), None)
                item_col = next((idx for idx, cell in enumerate(header) if any(matches(cell, r) for r in table_item_header_rules)), None)
                if rank_col is None or item_col is None or rank_col == item_col:
                    continue
                table_items = []
                for row in rows[1:]:
                    if not row or max(rank_col, item_col) >= len(row):
                        continue
                    rank_text = clean(row[rank_col])
                    item_text = clean(str(row[item_col] or "").replace("\n", " "))
                    rank_match = next((matches(rank_text, r) for r in table_rank_value_rules if matches(rank_text, r)), None)
                    if not rank_match or not item_text:
                        continue
                    try:
                        if hasattr(rank_match, "groups") and rank_match.groups():
                            order_value = int(rank_match.group(1))
                        else:
                            order_value = int(float(rank_text))
                    except (TypeError, ValueError):
                        order_value = len(table_items) + 1
                    table_items.append({"order": order_value, "text": item_text, "page": table_info.get("page")})
                    if max_items and len(table_items) >= max_items:
                        break
                if len(table_items) >= min_items:
                    page = table_items[0]["page"] if table_items else lines[anchor_idx]["page"]
                    return {
                        "provided": True, "status": "identified", "page": page, "items": table_items,
                        "message": render_message(messages, "ccap_ordre_identifie", page=page),
                        "source_label": render_message(messages, "ccap_source_label", page=page),
                        "extraction_mode": "table",
                    }

        items = []
        started = False
        current_item = None
        for entry in window[1:]:
            text = entry["text"]
            # Les lignes de pagination/en-tête explicitement paramétrées sont
            # ignorées sans interrompre une liste qui se poursuit sur la page suivante.
            if any(matches(text, r) for r in ignore_rules):
                continue
            if started and any(matches(text, r) for r in stop_rules):
                break
            captured = None
            for rule in item_rules:
                m = matches(text, rule)
                if not m:
                    continue
                if hasattr(m, "groups") and m.groups():
                    captured = clean(m.group(1))
                else:
                    captured = text
                break
            if captured:
                started = True
                current_item = {"order": len(items) + 1, "text": captured, "page": entry["page"]}
                items.append(current_item)
                if max_items and len(items) >= max_items:
                    break
                continue
            # Mécanique générique de reconstruction des lignes coupées par la mise
            # en page du PDF : après le début d'un item, une ligne qui n'est ni
            # un nouvel item ni un marqueur d'arrêt prolonge l'item courant.
            if started and current_item:
                current_item["text"] = clean(f'{current_item["text"]} {text}')
        if len(items) < min_items:
            continue

        page = items[0]["page"] if items else lines[anchor_idx]["page"]
        return {
            "provided": True, "status": "identified", "page": page, "items": items,
            "message": render_message(messages, "ccap_ordre_identifie", page=page),
            "source_label": render_message(messages, "ccap_source_label", page=page),
        }

    if require_marker and not marker_seen:
        return {"provided": True, "status": "section_without_priority", "items": [], "page": first_anchor_page, "message": render_message(messages, "ccap_rubrique_sans_priorite")}
    return {"provided": True, "status": "not_found", "items": [], "page": first_anchor_page, "message": render_message(messages, "ccap_ordre_non_identifie")}

# ═══════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION DU RAPPORT COMPLET
# ═══════════════════════════════════════════════════════════════════════════════
def generate_rapport(
    out: Path,
    conv: DocData,
    cctp: Optional[DocData],
    blocs: pd.DataFrame,
    gloss: pd.DataFrame,
    textes: pd.DataFrame,
    resultats_lecture: Dict,
    scores_eval: Dict,
    lot: str,
    entreprise: str,
    conv_text: str,
    marque: str = None,
    logo_path: str = None,
    adresse: str = None,
    email: str = None,
    analyste: str = None,
    nom_projet: str = None,
    documents: str = None,
    detections: pd.DataFrame = None,
    metadata_rules: pd.DataFrame = None,
    questions: pd.DataFrame = None,
    textes_reponse: Dict[str, Dict] = None,
    conv_infos: Dict[str, str] = None,
    cctp_infos: Dict[str, str] = None,
    contradictions: List[Dict] = None,
    axes_config: Dict = None,
    restitution_rules: List[Dict[str, object]] = None,
    messages: Dict[str, str] = None,
    ccap_reperage: Dict[str, object] = None,
    ccap_text: str = "",
    points_vigilance: List[Dict[str, str]] = None,
):
    """Génère plan d'actions et glossaire à partir du modèle public déjà calculé."""
    textes_reponse = textes_reponse or {}
    conv_infos = conv_infos or {}
    cctp_infos = cctp_infos or {}
    contradictions = contradictions or []
    messages = messages or {}
    ccap_reperage = ccap_reperage or {}
    points_vigilance = points_vigilance or []

    ctx = {
        **_extraire_details_convention(
            "GLOBAL", "", "", "", "", conv_text, lot,
            conv_pages=conv.pages, gloss=gloss, detections=detections, metadata_rules=metadata_rules,
        ),
        "lot": lot,
    }
    free_cols = [
        "description_convention", "lecture_entreprise", "risque_si_ignore",
        "type_preuve_cctp_attendue", "interpretation_croisee",
        "action_confirmee", "action_probable", "action_non_applicable",
        "question_bim_manager", "proposition_commerciale", "action_interne_standard",
        "action_interne_capacite", "action_interne_a_confirmer", "action_interne_non_demontree",
    ]
    report_results: Dict[str, Dict] = {}
    for bid, res in resultats_lecture.items():
        copy = dict(res)
        block = dict(res.get("bloc", {}))
        for col in free_cols:
            block[col] = _sub_ctx(str(block.get(col, "")), ctx, messages)
        block["question_bim_manager"] = clean(block.get("question_bim_manager", ""))
        copy["bloc"] = block
        report_results[bid] = copy

    convention_detail = conv_infos.get("nom_fichier") or conv.name
    if conv_infos.get("indice"):
        convention_detail += f" - indice {conv_infos['indice']}"
    convention_detail += f" - {len(conv.pages)} pages"
    cctp_detail = ""
    if cctp:
        cctp_detail = cctp_infos.get("nom_fichier") or cctp.name
        if cctp_infos.get("indice"):
            cctp_detail += f" - indice {cctp_infos['indice']}"
        cctp_detail += f" - {len(cctp.pages)} pages"

    docs_resume = clean(documents)
    # Deux périmètres de glossaire sont volontairement distingués :
    # - glossary_entries : vocabulaire nécessaire aux infobulles des livrables,
    #   qui peut inclure des termes utilisés par l'outil lui-même ;
    # - glossary_export_entries : glossaire autonome, limité aux termes réellement
    #   repérés dans les documents techniques analysés (Convention + CCTP).
    # Cela évite qu'un nom de plateforme ou un terme uniquement présent dans le
    # questionnaire soit présenté comme « détecté dans ce dossier ».
    glossary_tool_text = f"{conv_text or ''} {cctp.text if cctp else ''} {texte_outil_et_questionnaire(blocs, questions)}"
    cctp_pages = cctp.pages if cctp else []
    glossary_export_entries = glossary_entries_detected_documents(gloss, conv.pages, cctp_pages)
    meta = {
        "projet": conv.name.replace(".pdf", ""),
        "nom_projet": clean(nom_projet) if nom_projet else clean(conv_infos.get("operation", "")),
        "documents_analyses": docs_resume,
        "documents_resume": docs_resume,
        "convention_detail": convention_detail,
        "cctp_detail": cctp_detail,
        "entreprise": entreprise,
        "lot": lot,
        "marque": marque or entreprise,
        "adresse": clean(adresse), "email": clean(email), "analyste": clean(analyste),
        "date": datetime.now().strftime("%d/%m/%Y"),
        "contradictions": contradictions,
        "glossary_entries": glossary_entries_from_dataframe(filtrer_glossaire_detecte(gloss, glossary_tool_text)),
        "glossary_export_entries": glossary_export_entries,
        "axes_config": axes_config or {},
        "restitution_rules": restitution_rules or [],
        "messages": messages,
        "engine_params": ENGINE_PARAMS_GLOBAL,
        "ccap_reperage": ccap_reperage,
        "points_vigilance": points_vigilance,
        **charger_logo(logo_path),
    }
    # Réparation strictement visuelle : le moteur a déjà terminé ses décisions.
    report_results = _reparer_obj_affichage(report_results)
    textes_reponse = _reparer_obj_affichage(textes_reponse)
    meta = _reparer_obj_affichage(meta)

    base_path = out if out.suffix.lower() == ".pdf" else out.with_suffix(".pdf")
    generate_action_plan_pdf(base_path.with_name("Plan_Actions_Offre_BIM.pdf"), report_results, scores_eval, textes_reponse, meta)
    generate_action_plan_html(base_path.with_name("Plan_Actions_Offre_BIM.html"), report_results, scores_eval, textes_reponse, meta)
    generate_glossary_html(base_path.with_name("Glossaire_BIM.html"), meta)
    generate_glossary_pdf(base_path.with_name("Glossaire_BIM.pdf"), meta)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(description="Outil BIM paramétré")
    ap.add_argument("--param", default="Parametrage_Outil_BIM_V1.xlsx")
    ap.add_argument("--convention", default=None)
    ap.add_argument("--cctp", default=None)
    ap.add_argument("--ccap", default=None)
    ap.add_argument("--lot", default="")
    ap.add_argument("--entreprise", default="")
    ap.add_argument("--marque", default="")
    ap.add_argument("--logo", default=None)
    ap.add_argument("--adresse", default="")
    ap.add_argument("--email", default="")
    ap.add_argument("--analyste", default="")
    ap.add_argument("--nom-projet", dest="nom_projet", default="")
    ap.add_argument("--documents", default="")
    ap.add_argument("--complements-json", default="", dest="complements_json")
    ap.add_argument("--out", default="Rapport_BIM.html")
    ap.add_argument("--sans_eval", action="store_true")
    ap.add_argument("--web-only", action="store_true")
    ap.add_argument("--reponses-json", default="", dest="reponses_json")
    ap.add_argument("--autoeval-only", action="store_true", dest="autoeval_only")
    ap.add_argument("--format", default="both", choices=["pdf", "html", "both"])
    args = ap.parse_args()

    param = Path(args.param)
    if not param.exists():
        raise SystemExit(f"Paramétrage introuvable : {param}")
    (blocs, gloss, questions, textes, signaux, detections, metadata_rules,
     axes_df, coherence_rules, messages_df, params_df, text_filters_df,
     restitution_df, ccap_rules_df, interblock_rules_df, document_rules_df) = load_all(param)
    questions = normalize_questions_dataframe(questions)
    engine_params = load_params(params_df)
    messages = load_messages(messages_df)
    axes_config = load_axes_config(axes_df)
    restitution_rules = load_rule_rows(restitution_df)
    ccap_rules = load_rule_rows(ccap_rules_df)
    interblock_rules = load_rule_rows(interblock_rules_df)
    document_rules = load_rule_rows(document_rules_df)
    configure_text_filters(text_filters_df, engine_params)

    supplied_answers = None
    if args.reponses_json:
        answer_path = Path(args.reponses_json)
        if not answer_path.exists():
            raise SystemExit(f"Fichier de réponses introuvable : {answer_path}")
        supplied_answers = json.loads(answer_path.read_text(encoding="utf-8"))
        if not isinstance(supplied_answers, dict):
            raise SystemExit("Le JSON de réponses doit être un objet {id_question: 0|1|2}.")

    if args.autoeval_only:
        if supplied_answers is None:
            raise SystemExit("--autoeval-only nécessite --reponses-json.")
        dummy = {}
        for _, block in blocs.iterrows():
            bid = clean(block.get("id_bloc", ""))
            if bid:
                dummy[bid] = {"bloc": block, "conv_pf": "", "conv_pg": "", "cctp_pf": "", "cctp_pg": "", "_conv_text": "", "contradictions": []}
        scores = poser_questionnaire(
            blocs, args.lot, dummy, questions, axes_config, messages,
            gloss=gloss, detections=detections, metadata_rules=metadata_rules,
            reponses_fournies=supplied_answers,
        )
        contradictions = detecter_contradictions(coherence_rules, scores, None)
        out = Path(args.out)
        out_html = out if out.suffix.lower() == ".html" else out.with_suffix(".html")
        out_pdf = out_html.with_suffix(".pdf")
        generate_autoevaluation_reports(
            out_html, out_pdf, scores, axes_config, messages,
            {"entreprise": args.entreprise, "lot": args.lot, "adresse": args.adresse,
             "email": args.email, "analyste": args.analyste, "date": datetime.now().strftime("%d/%m/%Y"),
             "contradictions": contradictions, "engine_params": engine_params,
             "marque": args.marque or args.entreprise, **charger_logo(args.logo)},
        )
        print(f"Auto-évaluation HTML générée : {out_html}")
        print(f"Auto-évaluation PDF générée : {out_pdf}")
        return

    if not args.convention:
        raise SystemExit("La convention BIM est requise pour l'analyse classique.")
    conv_path = Path(args.convention)
    if not conv_path.exists():
        raise SystemExit(f"Convention introuvable : {conv_path}")
    conv = lire_pdf(conv_path)
    qualite_conv = evaluer_lisibilite_document(conv, engine_params)
    qualite_conv["label"] = conv_path.name
    if qualite_conv.get("statut") == "NON_EXPLOITABLE":
        raise SystemExit(render_message(messages, "pdf_convention_non_exploitable_error"))
    metadata_pages = p_int(engine_params, "fenetre_metadata_pages", len(conv.pages))
    raw_conv_doc = fitz.open(str(conv_path))
    raw_conv_pages = [raw_conv_doc[i].get_text() for i in range(min(metadata_pages, len(raw_conv_doc)))]
    conv_infos = _extraire_infos_document(raw_conv_pages, detections, metadata_rules)
    conv_infos["nom_fichier"] = conv_path.stem
    raw_conv_doc.close()

    cctp = None
    cctp_infos = {}
    qualite_cctp: Dict[str, object] = {}
    if args.cctp:
        cctp_path = Path(args.cctp)
        if cctp_path.exists():
            cctp = lire_pdf(cctp_path)
            qualite_cctp = evaluer_lisibilite_document(cctp, engine_params)
            qualite_cctp["label"] = cctp_path.name
            raw_doc = fitz.open(str(cctp_path))
            raw_pages = [raw_doc[i].get_text() for i in range(min(metadata_pages, len(raw_doc)))]
            cctp_infos = _extraire_infos_document(raw_pages, detections, metadata_rules)
            cctp_infos.update({"n_pages": len(raw_doc), "nom_fichier": cctp_path.stem})
            raw_doc.close()
    cctp_text = cctp.text if cctp else ""
    ccap_path = Path(args.ccap) if args.ccap else None
    ccap_reperage = extraire_priorite_pieces_ccap(ccap_path, ccap_rules, engine_params, messages)
    # Le CCAP reste hors sélection documentaire implicite. Sa lisibilité est
    # contrôlée et seules les règles Excel déclarées source=CCAP peuvent utiliser
    # une clause prescriptive ciblée pour enrichir un bloc.
    qualite_ccap: Dict[str, object] = {}
    ccap_doc = lire_pdf(ccap_path) if ccap_path and ccap_path.exists() else None
    if ccap_doc:
        qualite_ccap = evaluer_lisibilite_document(ccap_doc, engine_params)
        qualite_ccap["label"] = ccap_path.name
    ccap_text = ccap_doc.text if ccap_doc else ""

    supplement_docs: List[Dict[str, object]] = []
    if args.complements_json:
        manifest_path = Path(args.complements_json)
        if manifest_path.exists():
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            for idx, item in enumerate(manifest.get("items", []) or []):
                if not isinstance(item, dict):
                    continue
                stored_name = clean(item.get("stored_name"))
                supp_path = manifest_path.parent / stored_name
                if not stored_name or not supp_path.exists():
                    continue
                supp_doc = lire_pdf(supp_path)
                supp_quality = evaluer_lisibilite_document(supp_doc, engine_params)
                supp_quality["label"] = clean(item.get("display_name")) or supp_path.name
                if supp_quality.get("statut") == "NON_EXPLOITABLE":
                    raise SystemExit(render_message(messages, "supplement_unreadable_error"))
                supplement_docs.append({
                    "index": idx,
                    "path": supp_path,
                    "doc": supp_doc,
                    "quality": supp_quality,
                    "display_name": clean(item.get("display_name")) or supp_path.name,
                    "reference": clean(item.get("reference")),
                    "origin_source": clean(item.get("origin_source")).upper(),
                    "origin_page": clean(item.get("origin_page")),
                    "origin_excerpt": clean(item.get("origin_excerpt")),
                    "related_blocks": {clean(x).upper() for x in (item.get("related_blocks") or []) if clean(x)},
                    "resolved_reference": clean(item.get("reference")) if verifier_resolution_piece_ajoutee(
                        clean(item.get("reference")), clean(item.get("display_name")) or supp_path.name,
                        supp_doc.pages, document_rules, engine_params,
                    ) else "",
                })

    print(f"Analyse pour : {args.entreprise} - {args.lot}")
    print("[1/3] Analyse des exigences BIM")
    threshold = p_int(engine_params, "seuil_preuve_ciblee", 0)
    raw: Dict[str, Dict] = {}
    evidence_started = time.perf_counter()
    for _, block in blocs.iterrows():
        bid = clean(block.get("id_bloc", ""))
        kws = split_kw(block.get("keywords_convention", ""))
        strong = split_kw(block.get("keywords_preuve_forte", ""))
        extra_conf, extra_fort = _kws_depuis_signaux(bid, signaux)
        cctp_confirm = split_kw(block.get("keywords_cctp_confirme", ""))
        cctp_nuance = split_kw(block.get("keywords_cctp_nuance", ""))
        cctp_kws = cctp_confirm + cctp_nuance + extra_conf + extra_fort
        cctp_strong = cctp_confirm + extra_fort
        cctp_exclu = split_kw(block.get("keywords_cctp_exclu", ""))
        cctp_positive_signal = any_kw_non_negated(cctp_text, cctp_confirm + extra_conf + extra_fort)
        cctp_nuance_signal = any_kw_non_negated(cctp_text, cctp_nuance)
        conv_pg, conv_pf, conv_sc, conv_key = best_evidence(conv, kws, strong, set())
        cctp_pg, cctp_pf, cctp_sc, cctp_key = best_evidence(
            cctp, cctp_kws, cctp_strong, set(), lot=args.lot, enforce_lot_scope=True
        )
        cctp_ex_pg, cctp_ex_pf, cctp_ex_sc, cctp_ex_key = best_evidence(
            cctp, cctp_exclu, cctp_exclu, set(), lot=args.lot, enforce_lot_scope=True
        ) if cctp_exclu else ("", "", -999, "")
        supp_best = {"score": -999, "page": "", "text": "", "key": "", "source": "", "kind": "", "origin": ""}
        effective_conv_parts = [conv.text]
        # Cycle 50 : pour l'applicabilité du lot, les clauses explicitement
        # attribuées à un autre lot sont retirées du corpus CCTP utilisé par
        # l'évaluation générique. Une clause sans sujet de lot nommé reste active.
        effective_cctp_parts = [_texte_document_compatible_lot(cctp, args.lot)]
        for supp in supplement_docs:
            scoped = supp.get("related_blocks") or set()
            if scoped and bid.upper() not in scoped:
                continue
            supp_doc = supp.get("doc")
            origin = clean(supp.get("origin_source")).upper()
            if origin.startswith("CCTP"):
                effective_cctp_parts.append(_texte_document_compatible_lot(supp_doc, args.lot))
                skws, sstrong = cctp_kws, cctp_strong
            else:
                effective_conv_parts.append(supp_doc.text)
                skws, sstrong = kws, strong
            spg, spf, ssc, skey = best_evidence(
                supp_doc, skws, sstrong, set(), lot=args.lot,
                enforce_lot_scope=origin.startswith("CCTP")
            )
            if ssc > supp_best["score"]:
                supp_best = {
                    "score": ssc, "page": spg, "text": spf, "key": skey,
                    "source": clean(supp.get("display_name")),
                    "kind": f"supp{int(supp.get('index', 0))}", "origin": origin,
                    "origin_page": clean(supp.get("origin_page")),
                    "origin_excerpt": clean(supp.get("origin_excerpt")),
                }
        effective_conv_text = "\n".join(x for x in effective_conv_parts if clean(x))
        effective_cctp_text = "\n".join(x for x in effective_cctp_parts if clean(x))
        cctp_positive_signal = any_kw_non_negated(effective_cctp_text, cctp_confirm + extra_conf + extra_fort)
        cctp_nuance_signal = any_kw_non_negated(effective_cctp_text, cctp_nuance)
        raw[bid] = {"block": block, "conv_pg": conv_pg, "conv_pf": conv_pf, "conv_sc": conv_sc, "conv_key": conv_key,
                    "cctp_pg": cctp_pg, "cctp_pf": cctp_pf, "cctp_sc": cctp_sc, "cctp_key": cctp_key,
                    "cctp_ex_pg": cctp_ex_pg, "cctp_ex_pf": cctp_ex_pf, "cctp_ex_sc": cctp_ex_sc, "cctp_ex_key": cctp_ex_key,
                    "cctp_positive_signal": cctp_positive_signal, "cctp_nuance_signal": cctp_nuance_signal,
                    "effective_conv_text": effective_conv_text, "effective_cctp_text": effective_cctp_text,
                    "supp_best": supp_best}
    print(f"Pré-calcul des preuves terminé en {time.perf_counter() - evidence_started:.2f}s")
    reutiliser_preuve = p_text(engine_params, "preuve_reutilisation_interblocs").strip().lower() in {"oui", "yes", "true", "1"}
    excerpt_winner: Dict[str, Tuple[int, str]] = {}
    if not reutiliser_preuve:
        for bid, item in raw.items():
            key, score = item.get("cctp_key", ""), item.get("cctp_sc", -999)
            if key and score >= threshold and (key not in excerpt_winner or score > excerpt_winner[key][0]):
                excerpt_winner[key] = (score, bid)

    results: Dict[str, Dict] = {}
    for bid in sorted(raw, key=block_sort_key):
        item = raw[bid]; block = item["block"]
        conv_pf = item["conv_pf"] if item["conv_sc"] >= threshold else ""
        conv_pg = item["conv_pg"] if conv_pf else ""
        block_started = time.perf_counter()
        axes = evaluate(block, item.get("effective_cctp_text", cctp_text), args.lot, item.get("effective_conv_text", conv.text), signaux=signaux, params=engine_params, messages=messages, axes_config=axes_config)
        block_elapsed = time.perf_counter() - block_started
        code_non_app = p_text(engine_params, "applicabilite_si_cctp_exclu")
        has_positive_cctp = bool(item.get("cctp_positive_signal"))
        has_nuance_cctp = bool(item.get("cctp_nuance_signal"))
        use_exclusion_evidence = axes.get("applicabilite_lot") == code_non_app and bool(item.get("cctp_ex_pf")) and item.get("cctp_ex_sc", -999) >= threshold
        if use_exclusion_evidence:
            cctp_pf = item.get("cctp_ex_pf", "")
            cctp_pg = item.get("cctp_ex_pg", "")
            # Ce champ exprime l'existence d'une preuve positive propre au bloc,
            # indépendamment de la preuve finalement affichée. Les règles Excel
            # peuvent ainsi gérer un conflit positif/exclusion sans logique métier codée ici.
            cctp_positive = "Oui" if has_positive_cctp else "Non"
        else:
            # Une preuve de nuance est conservée comme contexte documentaire sans
            # être assimilée à une preuve positive. Le statut reste calculé par les
            # règles Excel (statut_si_nuance_seul / applicabilité).
            cctp_keep = has_positive_cctp or has_nuance_cctp
            if cctp_keep and not reutiliser_preuve:
                cctp_keep = excerpt_winner.get(item.get("cctp_key", ""), (None, None))[1] == bid
            cctp_pf = item["cctp_pf"] if cctp_keep else ""
            cctp_pg = item["cctp_pg"] if cctp_keep else ""
            cctp_positive = "Oui" if (has_positive_cctp and cctp_pf) else "Non"
        preuve_policy = clean(block.get("politique_preuve_convention", "TOUJOURS")).upper()
        if cctp_pf and preuve_policy == "MASQUER_SI_PREUVE_CCTP":
            conv_pf = ""
            conv_pg = ""
            if not clean(item.get("supp_best", {}).get("origin", "")).startswith("CCTP"):
                item["supp_best"] = {"score": -999, "page": "", "text": "", "key": "", "source": "", "kind": "", "origin": ""}
        results[bid] = {
            "bloc": block, "conclusion": axes.get("conclusion", ""),
            "presence_contractuelle": axes.get("presence_contractuelle", ""),
            "applicabilite_lot": axes.get("applicabilite_lot", ""),
            "capacite_entreprise": axes.get("capacite_entreprise", ""),
            "exclusion_explicit": axes.get("exclusion_explicit", False), "contradictions": [],
            "conv_pg": conv_pg, "conv_pf": conv_pf, "cctp_pg": cctp_pg, "cctp_pf": cctp_pf,
            "supp_pg": item.get("supp_best", {}).get("page", "") if item.get("supp_best", {}).get("score", -999) >= threshold else "",
            "supp_pf": item.get("supp_best", {}).get("text", "") if item.get("supp_best", {}).get("score", -999) >= threshold else "",
            "supp_source": item.get("supp_best", {}).get("source", ""),
            "supp_kind": item.get("supp_best", {}).get("kind", ""),
            "supp_origin_source": item.get("supp_best", {}).get("origin", ""),
            "supp_origin_page": item.get("supp_best", {}).get("origin_page", ""),
            "supp_origin_excerpt": item.get("supp_best", {}).get("origin_excerpt", ""),
            "cctp_preuve_positive": "Oui" if (cctp_positive == "Oui" or (item.get("supp_best", {}).get("score", -999) >= threshold and clean(item.get("supp_best", {}).get("origin", "")).startswith("CCTP"))) else "Non",
            # Indicateur documentaire générique : une nuance a été repérée dans
            # le CCTP sans être assimilée à une preuve positive. Le paramétrage
            # du bloc peut ensuite décider du statut public correspondant.
            "cctp_nuance_detectee": "Oui" if has_nuance_cctp else "Non",
            "_cctp_positive_pg": item.get("cctp_pg", "") if item.get("cctp_sc", -999) >= threshold else "",
            "_cctp_positive_pf": item.get("cctp_pf", "") if item.get("cctp_sc", -999) >= threshold else "",
            # DEV réel 01 : conserver aussi la preuve d'exclusion explicite. Une règle
            # documentaire plus générale peut affiner une preuve positive, mais elle ne
            # doit jamais faire disparaître une exclusion CCTP explicite visant le lot.
            # Le mécanisme reste générique : aucun bloc, lot ou projet n'est codé ici.
            "_cctp_exclusion_pg": item.get("cctp_ex_pg", "") if item.get("cctp_ex_sc", -999) >= threshold else "",
            "_cctp_exclusion_pf": item.get("cctp_ex_pf", "") if item.get("cctp_ex_sc", -999) >= threshold else "",
            "conv_extrait": conv_pf[:200] if conv_pf else "", "cctp_extrait": cctp_pf[:200] if cctp_pf else "",
            "_conv_text": conv.text,
        }
        print(f"{bid} -> {results[bid]['presence_contractuelle']} | {results[bid]['applicabilite_lot']} | {block_elapsed:.2f}s")
    # Cycle 37 : la qualification documentaire doit précéder les dépendances
    # interblocs et l'auto-évaluation. Ainsi, lorsqu'une preuve ciblée issue du
    # paramétrage Excel confirme ou exclut une exigence pour le lot, les blocs
    # dépendants sont évalués une seule fois à partir de cet état final.
    # Le moteur reste générique : aucune identité de bloc, de lot ou de projet
    # n'est codée dans cet ordonnancement.
    appliquer_gouvernance_preuves(results, engine_params)

    # Substitution générique des variables de contexte dans les textes de bloc
    # avant l'application des règles documentaires. Les textes métier restent
    # exclusivement dans le classeur de paramétrage.
    public_ctx = {
        **_extraire_details_convention(
            "GLOBAL", "", "", "", "", conv.text, args.lot,
            conv_pages=conv.pages, gloss=gloss, detections=detections, metadata_rules=metadata_rules,
        ),
        "lot": args.lot,
    }
    public_text_cols = [
        "description_convention", "lecture_entreprise", "risque_si_ignore",
        "type_preuve_cctp_attendue", "interpretation_croisee",
        "action_confirmee", "action_probable", "action_non_applicable",
        "question_bim_manager", "proposition_commerciale", "action_interne_standard",
        "action_interne_capacite", "action_interne_a_confirmer", "action_interne_non_demontree",
    ]
    for _res in results.values():
        _block = dict(_res.get("bloc", {}))
        for _col in public_text_cols:
            _block[_col] = _sub_ctx(str(_block.get(_col, "")), {**public_ctx, "lot": args.lot}, messages)
        _res["bloc"] = _block

    appliquer_regles_documentaires(
        results, document_rules, args.lot, {**public_ctx, "lot": args.lot}, messages,
        source_pages={
            "CONVENTION": conv.pages,
            "CCTP": (cctp.pages if cctp else []),
            "CCAP": (ccap_doc.pages if ccap_doc else []),
            "COMPLEMENT": [
                {
                    "page": page_no, "text": page_text,
                    "source": clean(s.get("display_name")) or render_message(messages, "source_supplement_label"),
                    "kind": f"supp{int(s.get('index', 0))}",
                    "related_blocks": list(s.get("related_blocks") or []),
                }
                for s in supplement_docs
                for page_no, page_text in (s.get("doc").pages or [])
            ],
        },
    )
    # Une règle documentaire peut remplacer une preuve grossière par une clause
    # ciblée trouvée ailleurs dans le document. Recalculer alors la présence
    # contractuelle avant d'appliquer les dépendances.
    appliquer_gouvernance_preuves(results, engine_params)

    # Les dépendances interblocs sont appliquées après la qualification
    # documentaire finale. Cela évite de conserver un forçage obsolète lorsque
    # la preuve ciblée a entre-temps fait évoluer le bloc source.
    appliquer_regles_interblocs(results, interblock_rules)

    # Si une règle Excel a levé une exclusion au profit d'une clarification et
    # qu'une preuve positive propre au bloc existe, afficher cette preuve plutôt
    # que l'extrait d'exclusion général. Le choix métier reste porté par Excel.
    for _res in results.values():
        if (
            clean(_res.get("cctp_preuve_positive")).lower() in {"oui", "yes", "true", "1"}
            and not bool(_res.get("exclusion_explicit"))
            and clean(_res.get("_cctp_positive_pf"))
        ):
            # Ne pas écraser une preuve plus précise déjà isolée par une règle
            # documentaire (specific_page/specific_rule_id).
            if not clean(_res.get("specific_rule_id")):
                _res["cctp_pf"] = _res.get("_cctp_positive_pf", "")
                _res["cctp_pg"] = _res.get("_cctp_positive_pg", "")
                _res["cctp_extrait"] = clean(_res.get("cctp_pf"))[:200]
    appliquer_gouvernance_preuves(results, engine_params)

    # DEV réel 01 — garde-fou transversal de hiérarchie documentaire.
    # Lorsqu'une exclusion CCTP explicite a été détectée pour le lot, elle reste
    # prioritaire après les règles documentaires/interblocs. Cela protège les cas
    # contradictoires « exclusion explicite + liste positive » sans introduire de
    # sémantique propre à B01 (ou à un autre bloc) dans le moteur.
    code_non_app = p_text(engine_params, "applicabilite_si_cctp_exclu")
    for _res in results.values():
        if not bool(_res.get("exclusion_explicit")):
            continue
        _res["applicabilite_lot"] = code_non_app
        _res["statut_public_override"] = ""
        _res["priorite_override"] = ""
        _res["interblock_exclusion"] = False
        _res["interblock_reason"] = ""
        if clean(_res.get("_cctp_exclusion_pf")):
            _res["cctp_pf"] = _res.get("_cctp_exclusion_pf", "")
            _res["cctp_pg"] = _res.get("_cctp_exclusion_pg", "")
            _res["cctp_extrait"] = clean(_res.get("cctp_pf"))[:200]
    appliquer_gouvernance_preuves(results, engine_params)
    appliquer_regles_interblocs(results, interblock_rules)
    appliquer_gouvernance_preuves(results, engine_params)
    appliquer_securite_cctp_non_lisible(results, qualite_cctp, engine_params, messages)

    print("[2/3] Auto-évaluation")
    scores = {}
    contradictions = []
    if not args.sans_eval:
        scores = poser_questionnaire(
            blocs, args.lot, results, questions, axes_config, messages,
            gloss=gloss, detections=detections, metadata_rules=metadata_rules,
            reponses_fournies=supplied_answers,
        )
        for bid, score in scores.items():
            if bid in results:
                results[bid]["capacite_entreprise"] = clean(score.get("capacite_code", ""))
        contradictions = detecter_contradictions(coherence_rules, scores, results)

    proposals = preparer_textes_reponse(textes, scores, results, args.entreprise, args.lot, messages, engine_params)
    ensure_public_model(results, scores, proposals, lot=args.lot, axes_config=axes_config, restitution_rules=restitution_rules, messages=messages)

    docs = clean(args.documents)
    if not docs:
        names = [conv_path.name]
        if cctp: names.append(Path(args.cctp).name)
        if ccap_path and ccap_path.exists(): names.append(ccap_path.name)
        docs = " | ".join(names)
    # Les noms des pièces ajoutées font partie du corpus de la nouvelle révision.
    # Ils sont ajoutés à l'index documentaire pour que la vigilance qui a motivé
    # l'ajout disparaisse lorsqu'elle correspond effectivement au fichier fourni.
    if supplement_docs:
        supp_names = [clean(s.get("display_name")) for s in supplement_docs if clean(s.get("display_name"))]
        existing_norm = norm(docs)
        for supp_name in supp_names:
            if norm(supp_name) not in existing_norm:
                docs = (docs + " | " + supp_name).strip(" |")
                existing_norm = norm(docs)
        # Cycle 43 : une référence ciblée n'est ajoutée à l'index documentaire
        # qu'après vérification de compatibilité entre le manifeste et le contenu
        # du PDF. Le nom opaque du fichier n'est donc ni nécessaire ni suffisant.
        for supp in supplement_docs:
            resolved_ref = clean(supp.get("resolved_reference"))
            if resolved_ref and norm(resolved_ref) not in existing_norm:
                docs = (docs + " | " + resolved_ref).strip(" |")
                existing_norm = norm(docs)

    quality_map = {
        "CONVENTION": qualite_conv,
        "CCTP": qualite_cctp,
        "CCAP": qualite_ccap,
    }
    for supp in supplement_docs:
        quality_map[f"COMPLEMENT_{supp.get('index', 0)}"] = supp.get("quality") or {}
    points_vigilance = detecter_points_vigilance_documentaires(
        conv.text, cctp.text if cctp else "", ccap_text, docs, engine_params, messages,
        qualite_documents=quality_map,
        conv_pages=conv.pages,
        cctp_pages=(cctp.pages if cctp else []),
        ccap_pages=(ccap_doc.pages if ccap_doc else []),
        additional_pages=[(clean(s.get("display_name")), s.get("doc").pages) for s in supplement_docs],
        ccap_hierarchy_pages={int(ccap_reperage.get("page"))} if ccap_reperage.get("page") else set(),
        document_rules=document_rules,
    )
    points_vigilance = qualifier_vigilances_pour_ajout(points_vigilance, results, engine_params)

    print("[3/3] Génération des livrables")
    out_path = Path(args.out)
    generate_rapport(
        out_path.with_suffix(".pdf"), conv, cctp, blocs, gloss, textes, results, scores,
        args.lot, args.entreprise, conv.text, marque=args.marque, logo_path=args.logo,
        adresse=args.adresse, email=args.email, analyste=args.analyste, nom_projet=args.nom_projet,
        documents=docs, detections=detections, metadata_rules=metadata_rules, questions=questions,
        textes_reponse=proposals, conv_infos=conv_infos, cctp_infos=cctp_infos,
        contradictions=contradictions, axes_config=axes_config, restitution_rules=restitution_rules,
        messages=messages, ccap_reperage=ccap_reperage, ccap_text=ccap_text,
        points_vigilance=points_vigilance,
    )

    if args.format in ("html", "both") and _HAS_DASHBOARD:
        html_path = out_path if out_path.suffix.lower() == ".html" else out_path.with_suffix(".html")
        all_text = conv.text + ((" " + cctp.text) if cctp else "") + " " + " ".join(s.get("doc").text for s in supplement_docs) + " " + texte_outil_et_questionnaire(blocs, questions)
        cctp_pages = cctp.pages if cctp else []
        full_glossary = enrichir_glossaire(gloss, conv.pages, cctp_pages)
        # Le panneau "Glossaire utile" du Dashboard suit le même périmètre que
        # le glossaire autonome : uniquement des termes réellement repérés dans
        # la Convention ou le CCTP. Les termes de l'outil restent disponibles
        # dans ``glossaire_complet`` pour les infobulles, sans être présentés
        # comme détectés dans le dossier.
        gloss_list = [g for g in full_glossary if clean(g.get("page_doc", "")) and clean(g.get("source_doc", ""))]
        global_ctx = _extraire_details_convention("GLOBAL", "", "", "", "", conv.text, args.lot, conv_pages=conv.pages, gloss=gloss, detections=detections, metadata_rules=metadata_rules)
        meta = {
            "projet": conv.name.replace(".pdf", ""), "nom_projet": args.nom_projet or conv_infos.get("operation", ""),
            "entreprise": args.entreprise, "lot": args.lot, "marque": args.marque or args.entreprise,
            "adresse": args.adresse, "email": args.email, "analyste": args.analyste,
            "date": datetime.now().strftime("%d/%m/%Y"), "documents_analyses": docs,
            "axes_config": axes_config, "restitution_rules": restitution_rules, "messages": messages,
            "engine_params": engine_params, "contradictions": contradictions, "ccap_reperage": ccap_reperage,
            "points_vigilance": points_vigilance,
            "supplementary_documents": [clean(s.get("display_name")) for s in supplement_docs],
            "n_pages_convention": len(conv.pages), "n_pgs": len(cctp.pages) if cctp else 0,
            "indice_doc": conv_infos.get("indice", ""), "date_doc": conv_infos.get("date", ""),
            "emetteur_doc": conv_infos.get("emetteur", ""), "maitre_ouvrage_doc": conv_infos.get("maitre_ouvrage", ""),
            "maitrise_oeuvre_doc": conv_infos.get("maitrise_oeuvre", ""),
            "cctp_nom": cctp_infos.get("nom_fichier", ""), "cctp_indice": cctp_infos.get("indice", ""),
            "cctp_date": cctp_infos.get("date", ""), "cctp_auteur": cctp_infos.get("emetteur", ""), "cctp_phase": cctp_infos.get("phase", ""),
            "niveau_bim": global_ctx.get("niveau_bim", ""), "bim_manager_conv": global_ctx.get("bim_manager", ""),
            "coordinateur_bim_conv": global_ctx.get("coordinateur_bim", ""), "geo_referencement": global_ctx.get("geo_referencement", False),
            "doe_numerique": global_ctx.get("doe_numerique", False), "formats_livrables": global_ctx.get("formats_livrables", ""),
            "plateforme_conv": global_ctx.get("plateforme", ""), "plateforme_page": global_ctx.get("plateforme_page", ""),
            "logiciel_conv": global_ctx.get("logiciel", ""),
            "format_ifc_conv": global_ctx.get("format_ifc", ""),
            "lod_conv": global_ctx.get("lod", ""), "lod_page": global_ctx.get("lod_page", ""),
            "nd_conv": global_ctx.get("nd", ""), "nd_page": global_ctx.get("nd_page", ""),
            "glossary_entries": glossary_entries_from_dataframe(filtrer_glossaire_detecte(gloss, all_text)),
            **charger_logo(args.logo),
        }
        conv_b64 = base64.b64encode(conv_path.read_bytes()).decode("ascii")
        cctp_b64 = base64.b64encode(Path(args.cctp).read_bytes()).decode("ascii") if cctp else ""
        ccap_b64 = base64.b64encode(ccap_path.read_bytes()).decode("ascii") if ccap_path and ccap_path.exists() else ""
        supplement_b64 = {
            f"supp{int(s.get('index', 0))}": {
                "name": clean(s.get("display_name")),
                "b64": base64.b64encode(Path(s.get("path")).read_bytes()).decode("ascii"),
            }
            for s in supplement_docs
        }
        # Même principe pour le dashboard : on répare uniquement les chaînes
        # affichées, après calcul des statuts et des règles métier.
        results_display = _reparer_obj_affichage(results)
        proposals_display = _reparer_obj_affichage(proposals)
        meta_display = _reparer_obj_affichage(meta)
        html_text = build_dashboard(
            results_display, scores, proposals_display, gloss_list, meta_display, conv_text=_reparer_texte_affichage(conv.text),
            conv_b64=conv_b64, cctp_b64=cctp_b64, ccap_b64=ccap_b64, supplement_b64=supplement_b64,
            pdf_filename="", glossaire_complet=full_glossary,
        )
        html_path.write_text(html_text, encoding="utf-8")
        print(f"Dashboard HTML généré : {html_path}")


if __name__ == "__main__":
    main()
