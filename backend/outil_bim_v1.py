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
import argparse, base64, html, io, json, re, sys, textwrap, unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import fitz
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image, HRFlowable, KeepTogether
)
from datetime import datetime
import base64
try:
    from dashboard_template import build_dashboard
    _HAS_DASHBOARD = True
except ImportError:
    _HAS_DASHBOARD = False

from bim_model import block_sort_key, ensure_public_model
from content_rules import glossary_entries_from_dataframe, normalize_questions_dataframe

from reporting_v2 import (
    generate_action_plan_html,
    generate_action_plan_pdf,
    generate_glossary_html,
    generate_glossary_pdf,
)

# ═══════════════════════════════════════════════════════════════════════════════
# PALETTE
# ═══════════════════════════════════════════════════════════════════════════════
C_BLUE    = colors.HexColor("#1F4E78")
C_BLUE_L  = colors.HexColor("#D9EAF7")
C_RED_L   = colors.HexColor("#FCE4D6")
C_RED_DK  = colors.HexColor("#7D0000")
C_ORA_L   = colors.HexColor("#FFF2CC")
C_ORA_DK  = colors.HexColor("#7D4E00")
C_GRN_L   = colors.HexColor("#E2EFDA")
C_GRN_DK  = colors.HexColor("#276221")
C_GRY_L   = colors.HexColor("#F3F6FA")
C_GRY_MID = colors.HexColor("#555555")
C_WHITE   = colors.white

STATUT_META = {
    "CONFIRMÉE":     (C_RED_L,  C_RED_DK,  "CONFIRMÉE"),
    "PROBABLE":      (C_ORA_L,  C_ORA_DK,  "PROBABLE"),
    "PARTIELLE":     (C_ORA_L,  C_ORA_DK,  "PARTIELLE"),
    "NON DÉMONTRÉE": (C_GRN_L,  C_GRN_DK,  "NON DÉMONTRÉE"),
    "NON APPLICABLE":(C_GRY_L,  C_GRY_MID, "NON APPLICABLE"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES TEXTE
# ═══════════════════════════════════════════════════════════════════════════════
def clean(s: object) -> str:
    return re.sub(r"\s+", " ",
                  html.unescape(str(s or "")).replace("\xa0", " ")).strip()

def norm(s: object) -> str:
    n = "".join(c for c in unicodedata.normalize("NFD", clean(s))
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", n.lower()).strip()

def split_kw(s: object) -> List[str]:
    return [clean(x) for x in re.split(r"[;|,]", str(s or "")) if clean(x)]

def contains(text: str, kw: str) -> bool:
    t, m = norm(text), norm(kw)
    if not m: return False
    if len(m) <= 4:
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
    if len(kw_low) <= 4:
        m = re.search(r"(?<![a-zà-ÿ0-9])" + re.escape(kw_low) + r"(?![a-zà-ÿ0-9])", text_lower[depart:])
        return (m.start() + depart) if m else -1
    return text_lower.find(kw_low, depart)

def any_kw(text: str, kws: List[str]) -> bool:
    return any(contains(text, k) for k in kws)

# ═══════════════════════════════════════════════════════════════════════════════
# FILTRES BIM
# ═══════════════════════════════════════════════════════════════════════════════
# Motifs administratifs GÉNÉRIQUES (regex, aucun nom de projet/adresse en dur).
# Sert à repérer des lignes d'en-tête typiques indépendamment du projet analysé.
ADMIN_PATTERNS = [
    r"\bsiret\b", r"\bnaf\s", r"tva intracommunautaire", r"\S+@\S+\.\S+",
    r"sarl au capital", r"ordre national des architectes",
    r"\bpage\s+\d+\s+sur\s+\d+",
    r"\b0[1-9](?:[\s.]\d{2}){4}\b",           # téléphone FR
    r"\b\d{1,4}\s+(?:rue|avenue|boulevard|impasse|allée)\s",  # adresse
]

def _matches_admin_pattern(line: str) -> bool:
    n = line.lower()
    return any(re.search(p, n) for p in ADMIN_PATTERNS)

def detect_running_headers(pages_raw: List[str], min_ratio: float = 0.3) -> set:
    """
    Détecte automatiquement les lignes qui se répètent sur plusieurs pages
    d'un même document (en-têtes/pieds de page, nom de projet répété,
    adresse d'agence en pied de page...). Générique : fonctionne sur
    n'importe quelle convention/CCTP, sans aucun texte de projet codé en dur.
    """
    from collections import Counter
    n_pages = len(pages_raw)
    if n_pages < 3:
        return set()
    counts = Counter()
    for txt in pages_raw:
        seen = set()
        for l in txt.split("\n"):
            l = l.strip().lower()
            if 8 <= len(l) <= 120:
                seen.add(l)
        for l in seen:
            counts[l] += 1
    threshold = max(3, int(n_pages * min_ratio))
    return {l for l, c in counts.items() if c >= threshold}

HORS_BIM = [
    "fsc", "pefc", "fdes", "acermi", "cov etiquette", "biosource", "biosourcé",
    "gestion des dechets", "gestion des déchets", "soged", "chantier vert",
    "bon de livraison", "label e+c", "e+c-", "siret", "tva intracommunautaire",
    "contact@", "naf ", "sarl au capital", "ordre national des architectes",
    "plan de gestion des dechets", "valorisation dechet",
    # Légendes d'images et titres de schémas dans les CCTP
    "exemple de données réinjectées", "exemple de donnees reinjectees",
    "synthèse du processus de renseignement",
    "synthese du processus de renseignement",
    "exemple de nomenclature excel renseignée",
    "exemple de nomenclature excel renseignee",
    # Titres de chapitres et sections
    "description des ouvrages",
    "prototype", "echafaudages et protections",
    "bardage bois", "bardage en polycarbonate",
    "obligations administratives", "responsabilite de l",
    "contenu des prix", "caractere global et forfaitaire",
    "generalites", "consistance des travaux",
    "documents et materiaux a soumettre", "dessins d execution",
    "contraintes concernant le site", "etude d execution",
    "performances acoustiques", "essais", "tolerances",
    "travaux a la charge des autres", "mesures obligatoires portant",
    "les mesures obligatoires portant",
    # Sections réglementaires et administratives du CCTP (hors BIM)
    "prescriptions environnementales", "nature de la reglementation",
    "liaisons entre les corps d", "contenu des prix",
    "demarches et autorisations", "responsabilite de l entrepreneur",
    "obligations administratives", "normes", "avis techniques",
    "regles de securite", "protections", "moyens de levage",
    "autocontrole", "provenance et agrement", "variantes exigees",
    "repliement du lot", "ouvrages complementaires",
    # Sommaires et tables des matières (pointillés)
    "............", "...........",
    # Annexes de codification (pas des règles contractuelles)
    "codification des plans generaux dans l arborescence",
    "codification des plans techniques par lot",
    "codification et organisation des vues",
    "codification des modelisation",
]

BIM_SIGNAL = [
    "maquette", "ifc", "bim", "rvt", "revit", "cde", "ged", "kroqi",
    "bim manager", "doe numerique", "doe numérique", "natif", "parametre",
    "paramètre", "nomenclature", "livrable bim", "convention bim",
    "charte bim", "loin", "lod", "point de base", "georeferencement",
    "géoréférencement", "gabarit", "plateforme collaborative", "data drop",
    "as-built", "aim", "gmao", "bep", "format natif", "export ifc",
    "coordination bim", "clash", "revue maquette", "tableau excel",
    "renseignement de la maquette", "processus bim", "maquette numerique",
    "maquette numérique", "doe numerique bim", "cahier des charges bim",
    "proprietes demandees", "propriétés demandées",
]

def is_hors_bim(txt: str) -> bool:
    n = norm(txt)
    return any(h in n for h in HORS_BIM)

def has_bim_signal(txt: str) -> bool:
    return any(contains(txt, b) for b in BIM_SIGNAL)

def clean_page_headers(txt: str, running_headers: set = frozenset()) -> str:
    lines = txt.split('\n')
    return '\n'.join(
        line for line in lines
        if line.strip().lower() not in running_headers
        and not _matches_admin_pattern(line)
    )

def clean_proof_text(txt: str) -> str:
    txt = txt.strip()
    txt = re.sub(r"^\d+\s+rue\s+[A-Za-zéèê]+\s+\d{5}\s+[A-Za-zéè\-]+\s*", "", txt).strip()
    txt = re.sub(r"^[a-z]\.\s+", "", txt).strip()
    # Supprimer en-tête de page CCTP (toutes variantes)
    txt = re.sub(
        r"DCE[\s\S]{0,400}?\d{2}/\d{2}/\d{4}\s*",
        "", txt, count=1, flags=re.IGNORECASE).strip()
    return txt

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
    df = df.dropna(how="all").fillna("")
    df = df[[c for c in df.columns if c and c.lower() != "nan"]]
    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.strip().str.lower().ne("non")]
    return df


# =============================================================================
# LECTURE FEUILLE 70_ENTREPRISE
# =============================================================================
def load_entreprise(path: Path) -> Dict:
    """Lit la feuille 70_Entreprise et retourne un dict clé/valeur."""
    try:
        import openpyxl as _ox
        wb = _ox.load_workbook(str(path), data_only=True)
        if "70_Entreprise" not in wb.sheetnames:
            return {}
        ws = wb["70_Entreprise"]
        data = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and not str(row[0]).startswith("#"):
                # Lire colonne B (index 1) en priorité, sinon colonne C (index 2)
                val = row[1] if (row[1] is not None and str(row[1]).strip()) \
                      else (row[2] if len(row) > 2 and row[2] is not None else None)
                data[str(row[0]).strip()] = str(val).strip() if val is not None else ""
        return data
    except Exception:
        return {}

def get_lots_entreprise(ent: Dict) -> List[str]:
    """Retourne la liste des lots renseignés (lot_1 à lot_5)."""
    lots = []
    for k in ["lot_1","lot_2","lot_3","lot_4","lot_5"]:
        v = ent.get(k,"").strip()
        if v:
            lots.append(v)
    return lots

def ecrire_historique(path: Path, convention: str, lot: str, score: int):
    """Écrit l'analyse dans la section E de 70_Entreprise."""
    try:
        import openpyxl as _ox
        from datetime import datetime as _dt
        wb = _ox.load_workbook(str(path))
        if "70_Entreprise" not in wb.sheetnames:
            return
        ws = wb["70_Entreprise"]
        # Trouver les cellules ao_X_convention vides
        for i in range(1, 6):
            for row in ws.iter_rows():
                if row[0].value == f"ao_{i}_convention" and not row[1].value:
                    row[1].value = convention
                    # Trouver les lignes suivantes
                    for row2 in ws.iter_rows():
                        if row2[0].value == f"ao_{i}_lot":    row2[1].value = lot
                        if row2[0].value == f"ao_{i}_date":   row2[1].value = _dt.now().strftime("%d/%m/%Y")
                        if row2[0].value == f"ao_{i}_score":  row2[1].value = str(score)
                    wb.save(str(path))
                    return
    except Exception:
        pass

def load_all(path: Path):
    blocs    = read_sheet(path, "10_Blocs",       "id_bloc")
    gloss    = read_sheet(path, "40_Glossaire",   "terme")
    questions= read_sheet(path, "50_Questions",   "id_bloc")
    textes   = read_sheet(path, "60_Reponses",    "id_bloc")
    def _try(sheet, key):
        try:
            return read_sheet(path, sheet, key)
        except SystemExit:
            return pd.DataFrame()
    signaux       = _try("30_CCTP_Signaux", "categorie")
    detections    = _try("35_Detections_Convention", "cle_ctx")
    metadata_rules= _try("36_Extraction_Metadata", "champ")
    axes_config   = _try("16_Axes_Config", "axe")
    coherence     = _try("55_Regles_Coherence", "id_regle")
    messages_df   = _try("12_Messages_Moteur", "cle")
    params_df     = _try("05_Parametres_Moteur", "cle")
    return (blocs, gloss, questions, textes, signaux, detections,
            metadata_rules, axes_config, coherence, messages_df, params_df)


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
        out.setdefault(axe, {})[code] = {
            "label": clean(row.get("label", "")) or code,
            "couleur": clean(row.get("couleur", "")) or "#666666",
            "symbole": clean(row.get("symbole", "")),
            "ordre": int(row.get("ordre", 99) or 99),
        }
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
                if len(phrase) < 30 or len(phrase) > 300:
                    continue
                sc = sum(1 for m in mots if contains(phrase, m))
                if sc == 0:
                    continue  # le mot-clé doit être présent -- le bonus seul ne doit jamais suffire
                if any(w in norm(phrase) for w in
                       ["est ", "designe ", "signifie ", "correspond ", "permet ",
                        "defini", "represente", "constitue", "comprend"]):
                    sc += 1
                if sc > best_score:
                    best_score  = sc
                    # Centrer l'extrait sur le mot-clé plutôt que tronquer
                    # depuis le début -- un mot-clé apparaissant tard dans la
                    # phrase ne doit jamais être coupé par la troncature à 250c.
                    pos_kw = next((norm(phrase).find(norm(m)) for m in mots
                                   if norm(m) and norm(m) in norm(phrase)), 0)
                    if pos_kw > 150:
                        deb = max(0, pos_kw - 100)
                        best_phrase = phrase[deb:deb + 250]
                    else:
                        best_phrase = phrase[:250]
                    best_page   = str(pg)
                    best_source = src

        # Repli générique : si aucune "phrase-définition" de bonne qualité
        # n'a été trouvée (score < 2 -- notamment systématique pour un terme
        # n'ayant qu'un seul mot-clé, ou une phrase de plus de 300 caractères,
        # fréquente dans les conventions BIM), chercher simplement une fenêtre
        # de texte autour de la première occurrence du mot-clé. Une citation
        # imparfaite reste bien plus utile qu'aucune citation du tout.
        if best_score < 2:
            best_phrase, best_page, best_source = "", "", ""
            for pg, txt, src in all_pages:
                low = txt.lower()
                for m in mots:
                    i = trouve_position(low, m)
                    if i == -1:
                        continue
                    start = max(0, i - 80)
                    extrait = clean(txt[start:i + len(m) + 80])
                    if re.search(r"\.{4,}", extrait):  # ignorer les lignes de sommaire
                        continue
                    best_phrase, best_page, best_source = extrait[:250], str(pg), src
                    break
                if best_page:
                    break

        result.append({
            "terme":            terme,
            "definition_excel": def_excel,
            "definition_doc":   best_phrase if (best_score >= 2 or best_page) else "",
            "page_doc":         best_page,
            "source_doc":       best_source,
            "source_excel":     clean(row.get("source", "")),
            "categorie":        clean(row.get("categorie", "")),
            "niveau":           clean(row.get("niveau", "")),
        })
    return result

# ═══════════════════════════════════════════════════════════════════════════════
# LECTURE PDF
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass
class DocData:
    name: str
    pages: List[Tuple[int, str]]
    text: str = field(default="", repr=False)

def lire_pdf(path: Path) -> DocData:
    doc = fitz.open(str(path))
    raw_pages = []
    for pg in doc:
        txt = pg.get_text("text").replace("\xa0", " ")
        txt = re.sub(r"[ \t]+", " ", txt)
        raw_pages.append(txt)
    # Détection générique des en-têtes/pieds de page répétés sur ce document précis
    running_headers = detect_running_headers(raw_pages)
    pages = []
    for i, txt in enumerate(raw_pages, start=1):
        txt = clean_page_headers(txt, running_headers)
        txt = re.sub(r"\n{3,}", "\n\n", txt).strip()
        pages.append((i, txt))
    return DocData(name=path.name, pages=pages,
                   text=" ".join(t for _, t in pages))

def segments(doc: DocData) -> List[Tuple[int, str]]:
    """
    Découpe le texte en segments exploitables.
    Stratégie :
    1. Reconstituer les paragraphes (lignes séparées par une ligne vide)
    2. Découper sur les numéros de section (2.14, 3.1 etc.)
    3. Filtrer : hors-BIM, titres seuls, trop courts
    """
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
            if len(part) < 40: continue
            if is_hors_bim(part): continue
            # Exclure les lignes de sommaire avec pointillés
            if part.count('.') > len(part) * 0.3: continue
            # Exclure les titres purs (court + tout en majuscules ou numéro seul)
            if re.match(r"^\d+\.?\d*\s+[A-Z]{3}", part) and len(part) < 80: continue
            if len(part) > 500: part = part[:497] + "..."
            out.append((page, part))
    return out

def score_seg(txt: str, kws: List[str], strong: List[str]) -> int:
    sc = 0
    for k in kws:
        if contains(txt, k): sc += 3
    for k in strong:
        if contains(txt, k): sc += 7
    if has_bim_signal(txt): sc += 10  # Signal BIM structurant = prioritaire
    for w in ["doit","devra","doivent","obligatoire","est attendu",
              "à fournir","livrable","conforme","déposer"]:
        if contains(txt, w): sc += 2
    if is_hors_bim(txt): sc -= 60
    return sc

# Termes trop génériques pour justifier, seuls, qu’un extrait prouve un bloc.
# Cette liste est transversale (aucun id_bloc) : elle empêche par exemple qu’un
# passage générique sur une « maquette DOE » soit retenu pour un bloc portant
# sur les paramètres, la codification ou les nomenclatures.
_EVIDENCE_GENERIC_TERMS = {
    "bim", "ifc", "maquette", "maquette numerique", "modele", "modele numerique",
    "information", "informations", "donnee", "donnees", "livrable", "livrables",
    "description", "type", "famille", "fabricant", "mise a jour", "format natif",
}


def _evidence_matches(txt: str, kws: List[str], strong: List[str]) -> Tuple[List[str], List[str], List[str]]:
    """Retourne (mots ordinaires, signaux forts, mots ordinaires spécifiques)."""
    matched = [k for k in kws if k and contains(txt, k)]
    matched_strong = [k for k in strong if k and contains(txt, k)]
    specific = []
    for k in matched:
        nk = norm(k)
        # Une expression multi-mots, un terme suffisamment long ou un acronyme
        # métier est considéré comme ciblé, sauf s’il appartient au socle BIM
        # très générique ci-dessus.
        is_acronym = str(k).strip().isupper() and 3 <= len(str(k).strip()) <= 8
        is_specific = (" " in nk or len(nk) >= 7 or is_acronym) and nk not in _EVIDENCE_GENERIC_TERMS
        if is_specific:
            specific.append(k)
    return matched, matched_strong, specific


def best_evidence(doc: Optional[DocData], kws: List[str], strong: List[str],
                  used: set) -> Tuple[str, str, int, str]:
    """
    Sélectionne un extrait réellement rattaché au bloc.

    Un simple signal BIM générique ne suffit plus. Le passage doit contenir :
      - au moins un signal fort du bloc ; ou
      - au moins deux mots-clés du bloc ; ou
      - au moins un mot-clé suffisamment spécifique.

    Cette règle évite les preuves hors sujet et conserve les clauses métier
    explicites même lorsqu’elles ne contiennent pas le mot « BIM » (par ex.
    « le titulaire devra renseigner ce tableau Excel »).
    """
    if not doc:
        return "", "", -999, ""
    best_sc, best_pg, best_txt, best_key = -999, "", "", ""
    for page, seg in segments(doc):
        key = norm(seg[:160])
        if key in used:
            continue

        matched, matched_strong, specific = _evidence_matches(seg, kws, strong)
        eligible = bool(matched_strong or len(set(map(norm, matched))) >= 2 or specific)
        if not eligible:
            continue

        sc = score_seg(seg, kws, strong)
        # Bonus de ciblage : à score proche, privilégier l’extrait qui contient
        # le vocabulaire le plus propre au bloc plutôt qu’un paragraphe BIM large.
        sc += 4 * len(set(map(norm, specific)))
        sc += 3 * len(set(map(norm, matched_strong)))

        # L’absence d’un signal BIM générique ne doit pas éliminer une clause
        # explicite et ciblée. On ne pénalise que les candidats sans signal fort
        # et faiblement spécifiques.
        if not has_bim_signal(seg) and not matched_strong and len(specific) < 2:
            sc -= 4

        if sc > best_sc:
            best_sc, best_pg, best_txt, best_key = sc, str(page), seg, key

    if best_sc < 8 or not best_txt:
        return "", "", best_sc, ""
    return best_pg, clean_proof_text(best_txt), best_sc, best_key


def appliquer_gouvernance_preuves(resultats: Dict[str, Dict], conv_text: str, cctp_text: str, lot: str) -> None:
    """Durcit la cohérence statut/preuve/lot après sélection des extraits.

    Principes : une preuve d'un autre lot est rejetée, une exclusion explicite
    prime sur une occurrence positive et aucun statut actif n'est conservé sans
    extrait affichable. Les statuts techniques historiques restent compatibles
    avec le reste du moteur, tandis que les libellés publics sont gérés par les
    générateurs de livrables.
    """
    lot_n = norm(lot or "")
    all_text = norm((conv_text or "") + " " + (cctp_text or ""))

    def autre_lot(txt: str) -> bool:
        n = norm(txt or "")
        if not n:
            return False
        # Cas explicites fréquents : une clause limitée à un autre corps d'état.
        autres = ("cvc", "plomberie", "electricite", "structure", "charpente", "lot 03")
        facade = any(x in lot_n for x in ("facade", "bardage", "couverture", "etancheite"))
        return facade and any(x in n for x in autres) and not any(x in n for x in ("facade", "bardage", "couverture", "etancheite", "tous les lots"))

    def rejet(res: Dict, champ: str) -> None:
        res[champ] = ""
        res[champ.replace("_pf", "_pg")] = ""
        res[champ.replace("_pf", "_extrait")] = ""

    for bid, res in resultats.items():
        for champ in ("conv_pf", "cctp_pf"):
            if autre_lot(res.get(champ, "")):
                res.setdefault("preuves_rejetees", []).append({
                    "texte": res.get(champ, ""),
                    "raison": "Extrait limité à un autre lot ; non recevable pour le lot analysé.",
                })
                rejet(res, champ)

        # Une clause de géoréférencement ne prouve pas la qualité de modélisation B11.
        if bid == "B11":
            for champ in ("conv_pf", "cctp_pf"):
                n = norm(res.get(champ, ""))
                if n and ("georeferencement" in n or "origine projet" in n) and not any(k in n for k in ("controle qualite", "autocontrole", "ifcspace", "categorie ifc", "audit", "regles de modelisation")):
                    res.setdefault("preuves_rejetees", []).append({"texte": res.get(champ, ""), "raison": "Clause de géoréférencement sans lien direct avec la qualité de modélisation."})
                    rejet(res, champ)

    # Exclusions explicites : elles priment sur la simple présence des termes 4D/5D.
    exclusions_4d = "exclusions de principe" in all_text and "phasage 4d" in all_text
    exclusions_5d = "exclusions de principe" in all_text and ("chiffrage 5d" in all_text or "5d contractuel" in all_text)
    for bid, excluded, reason in (
        ("B10", exclusions_4d, "Phasage 4D explicitement exclu de principe, sauf ordre écrit contraire."),
        ("B12", exclusions_5d, "Chiffrage 5D explicitement exclu de principe, sauf ordre écrit contraire."),
    ):
        if excluded and bid in resultats:
            r=resultats[bid]
            r.update({"statut":"NON APPLICABLE", "certitude":"Fort", "applicabilite_lot":"NON_APPLICABLE", "conclusion":reason, "exclusion_explicit":True})
            r["conv_pf"] = r["cctp_pf"] = ""
            r["conv_pg"] = r["cctp_pg"] = ""

    # B09 : une clause CCTP nommant le profil BIM du lot et son niveau est directe.
    if "B09" in resultats:
        n=norm(cctp_text or "")
        if "profil bim du lot" in n and "niveau bim" in n and any(k in n for k in ("facade", "bardage", "enveloppe")):
            r=resultats["B09"]
            r.update({"statut":"CONFIRMÉE", "certitude":"Fort", "presence_contractuelle":"CCTP", "applicabilite_lot":"APPLICABLE", "conclusion":"Obligation de collaboration BIM confirmée pour le lot analysé par le CCTP."})

    # B07 : une plateforme commune de données est quasiment toujours nécessaire
    # dès qu'un projet a des exigences BIM confirmées ailleurs -- elle ne doit
    # pas dépendre de sa propre preuve indépendante dans le CCTP. Seule une
    # exclusion explicite (déjà gérée par keywords_cctp_exclu, via le statut
    # NON_APPLICABLE) doit l'écarter.
    if "B07" in resultats:
        r7 = resultats["B07"]
        deja_exclu = r7.get("applicabilite_lot") == "NON_APPLICABLE"
        deja_preuve = bool(r7.get("conv_pf") or r7.get("cctp_pf"))
        if not deja_exclu and not deja_preuve:
            autres_confirmes = sum(
                1 for bid2, r2 in resultats.items()
                if bid2 != "B07" and r2.get("statut") in ("CONFIRMÉE", "PROBABLE")
            )
            if autres_confirmes >= 2:
                r7.update({
                    "statut": "PROBABLE", "certitude": "Moyen",
                    "presence_contractuelle": "PRESUMEE",
                    "applicabilite_lot": "PROBABLE",
                    "conclusion": (
                        "Aucune plateforme n'est nommée explicitement, mais le projet comporte "
                        "plusieurs exigences BIM confirmées : un environnement commun de données "
                        "est présumé nécessaire, sauf indication contraire des documents."
                    ),
                })
                r7["b07_defaut_applicable"] = True

    # Recalcul final : statut actif impossible sans preuve affichable, hors B09 confirmé
    # par la clause globale du CCTP et hors exclusions explicites.
    for bid, r in resultats.items():
        has_proof = bool(r.get("conv_pf") or r.get("cctp_pf"))
        if r.get("exclusion_explicit"):
            continue
        if bid == "B09" and r.get("statut") == "CONFIRMÉE":
            continue
        if bid == "B07" and r.get("b07_defaut_applicable"):
            continue
        if r.get("statut") in ("CONFIRMÉE", "PROBABLE", "PARTIELLE") and not has_proof:
            r.update({"statut":"NON DÉMONTRÉE", "certitude":"Faible", "presence_contractuelle":"ABSENTE", "applicabilite_lot":"A_CONFIRMER", "conclusion":"Aucune preuve suffisamment ciblée et applicable au lot n'a été retenue dans les documents analysés."})

        # Source publique reconstruite uniquement depuis les preuves conservées.
        if r.get("conv_pf") and r.get("cctp_pf"):
            r["presence_contractuelle"]="CONVENTION_CCTP"
        elif r.get("cctp_pf"):
            r["presence_contractuelle"]="CCTP"
        elif r.get("conv_pf"):
            r["presence_contractuelle"]="CONVENTION"
        elif r.get("statut") != "NON APPLICABLE":
            r["presence_contractuelle"]="ABSENTE"

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


def statut_legacy(presence: str, applicabilite: str) -> Tuple[str, str]:
    """
    Dérive un statut/certitude "historique" (CONFIRMÉE/PROBABLE/NON DÉMONTRÉE/
    NON APPLICABLE) à partir des 3 axes, pour le tri, le regroupement en
    sections et l'éligibilité au questionnaire -- sans perdre l'information
    des 3 axes eux-mêmes, toujours disponibles séparément dans le résultat.
    """
    if applicabilite == "NON_APPLICABLE":
        return "NON APPLICABLE", "Faible"
    if presence == "ABSENTE":
        return "NON DÉMONTRÉE", "Faible"
    if applicabilite == "APPLICABLE":
        return "CONFIRMÉE", "Fort"
    if applicabilite == "PROBABLE":
        return "PROBABLE", "Moyen"
    return "PROBABLE", "Moyen"  # A_CONFIRMER


def evaluate(bloc: pd.Series, cctp_text: str, lot: str,
             conv_text: str = "",
             signaux: Optional[pd.DataFrame] = None,
             params: Optional[Dict[str, object]] = None,
             messages: Optional[Dict[str, str]] = None) -> Dict:
    """
    Évalue un bloc BIM selon 3 axes INDÉPENDANTS, plutôt qu'un statut unique
    qui mélangeait présence contractuelle / applicabilité au lot / capacité
    de l'entreprise :
      - presence_contractuelle : ABSENTE / CONVENTION / CCTP / CONVENTION_CCTP
      - applicabilite_lot      : APPLICABLE / PROBABLE / A_CONFIRMER / NON_APPLICABLE
      - capacite_entreprise    : NON_EVALUEE (mise à jour ensuite par le questionnaire)
    Tous les seuils/messages/codes viennent de l'Excel (05_Parametres_Moteur,
    12_Messages_Moteur, 16_Axes_Config) -- aucun id_bloc codé en dur.
    """
    params = params or {}
    messages = messages or {}
    titre        = clean(bloc.get("titre_bloc", ""))
    bid_courant  = clean(bloc.get("id_bloc", ""))
    kws_conf     = split_kw(bloc.get("keywords_cctp_confirme", ""))
    kws_nu       = split_kw(bloc.get("keywords_cctp_nuance", ""))
    kws_exclu    = split_kw(bloc.get("keywords_cctp_exclu", ""))
    kws_conv     = split_kw(bloc.get("keywords_convention", ""))
    kws_fort     = split_kw(bloc.get("keywords_preuve_forte", ""))

    _extra_conf, _extra_fort = _kws_depuis_signaux(bid_courant, signaux)
    kws_conf = kws_conf + _extra_conf
    kws_fort = kws_fort + _extra_fort

    has_cctp      = bool(cctp_text.strip())
    has_conf_cctp = any_kw(cctp_text, kws_conf)
    has_nu_cctp   = any_kw(cctp_text, kws_nu)
    has_exclu_cctp= bool(kws_exclu and any_kw(cctp_text, kws_exclu))
    has_fort_conv = any_kw(conv_text, kws_fort)
    has_conf_conv = any_kw(conv_text, kws_conf + kws_fort)
    sc_conv = score_seg(conv_text, kws_conv, kws_fort) if conv_text else -999

    # ═══ AXE 1 : PRÉSENCE CONTRACTUELLE ═══════════════════════════════════
    # Existe-t-il une clause explicite (convention et/ou CCTP), indépendamment
    # de savoir si elle s'applique à CE lot précis ?
    seuil_conv = p_int(params, "seuil_presence_convention", 2)
    conv_explicite = has_fort_conv or has_conf_conv or sc_conv >= seuil_conv
    cctp_explicite = has_cctp and (has_conf_cctp or has_nu_cctp or has_exclu_cctp)
    if conv_explicite and cctp_explicite:
        presence, source = "CONVENTION_CCTP", "Convention + CCTP"
    elif conv_explicite:
        presence, source = "CONVENTION", "Convention"
    elif cctp_explicite:
        presence, source = "CCTP", "CCTP"
    else:
        presence, source = "ABSENTE", "Aucune source explicite"

    # ═══ AXE 2 : APPLICABILITÉ AU LOT ══════════════════════════════════════
    # Cette exigence, si elle existe, concerne-t-elle CE lot précis ?
    # Piloté par les défauts globaux de 05_Parametres_Moteur, jamais un id_bloc.
    _raw_nuance = clean(bloc.get("statut_si_nuance_seul", "")).upper()
    if presence == "ABSENTE":
        applicabilite = p_text(params, "applicabilite_si_absence") or "A_CONFIRMER"
    elif has_exclu_cctp:
        applicabilite = p_text(params, "applicabilite_si_cctp_exclu") or "NON_APPLICABLE"
    elif has_conf_cctp and not has_nu_cctp:
        applicabilite = p_text(params, "applicabilite_si_cctp_confirme") or "APPLICABLE"
    elif has_nu_cctp:
        if _raw_nuance == "NON_APPLICABLE":
            applicabilite = "NON_APPLICABLE"
        elif _raw_nuance in ("NON_DEMONSTREE", "NON_DEMONTREE"):
            applicabilite = "A_CONFIRMER"
        else:
            applicabilite = p_text(params, "applicabilite_si_cctp_nuance") or "PROBABLE"
    elif has_cctp:
        applicabilite = p_text(params, "applicabilite_si_silence_cctp") or "A_CONFIRMER"
    else:
        applicabilite = p_text(params, "applicabilite_sans_cctp") or "A_CONFIRMER"

    # ═══ Citation (mot-clé détecté, pour le message) ══════════════════════
    if has_exclu_cctp:
        kw = next((k for k in kws_exclu if contains(cctp_text, k)), "")
    elif has_conf_cctp:
        kw = next((k for k in kws_conf if contains(cctp_text, k)), "")
    elif has_nu_cctp:
        kw = next((k for k in kws_nu if contains(cctp_text, k)), "")
    elif has_conf_conv:
        kw = next((k for k in (kws_fort + kws_conf) if contains(conv_text, k)), "")
    else:
        kw = ""

    # ═══ Message de conclusion (paramétré, 12_Messages_Moteur) ════════════
    vals = {"titre": titre, "lot": lot, "source": source, "detection": kw}
    if presence == "ABSENTE":
        conclusion = render_message(messages, "conclusion_absente", **vals)
    elif applicabilite == "APPLICABLE":
        conclusion = render_message(messages, "conclusion_applicable", **vals)
    elif applicabilite == "PROBABLE":
        conclusion = render_message(messages, "conclusion_probable", **vals)
    elif applicabilite == "NON_APPLICABLE":
        conclusion = render_message(messages, "conclusion_non_applicable", **vals)
    else:
        conclusion = render_message(messages, "conclusion_a_confirmer", **vals)

    return {
        "presence_contractuelle": presence,
        "applicabilite_lot": applicabilite,
        "capacite_entreprise": "NON_EVALUEE",
        "conclusion": conclusion,
        "terme_detecte": kw,
    }

# ═══════════════════════════════════════════════════════════════════════════════
# STYLES PDF — typographie identique au dashboard HTML
# Correspondance px → pt : 1px ≈ 0.75pt (72pt/96px)
# Dashboard body  13px → 9.75pt   | leading 1.6 → leading*1.6
# ═══════════════════════════════════════════════════════════════════════════════
# Correspondance exacte dashboard → PDF
# body 13px/1.6            → Sm      9.5pt leading 15
# titre 20px/400           → Title   15pt leading 20
# KPI val 28px/700         → KPIVal  21pt bold
# KPI label 10px/700 upper → Cap     7.5pt bold
# section cap 8.5px/700    → SecCap  6.5pt bold
# titre exig 12px/600      → H3      9pt bold
# corps carte 11.5px       → Body    8.5pt leading 13.5
# citation 10.5px italic   → Cit     7.8pt italic
# tableau th 8px/700 upper → TH      6pt bold
# tableau td 11px          → TD      8pt
# bid bloc 10px/700        → BID     7.5pt bold
# statut/diff 9.5px        → Sub     7pt
# pied 9.5px               → Foot    7pt italic
# ═══════════════════════════════════════════════════════════════════════════════
def build_styles():
    S = getSampleStyleSheet()
    BLK  = colors.HexColor("#111111")
    GRY  = colors.HexColor("#888888")
    GRY6 = colors.HexColor("#666666")
    GRY4 = colors.HexColor("#444444")
    CIT  = colors.HexColor("#666666")

    # (name, parent, fontSize, textColor, fontName, leading, spaceAfter, leftIndent)
    defs = [
        # Titre principal — 20px/400 → 15pt  (préfixe B_ pour éviter conflits ReportLab)
        ("B_Title",   "Normal", 15,   BLK,  "Helvetica",        23, 4,  0),
        # Sous-titre entête — 8.5px/color:#999
        ("B_Sub",     "Normal",  6.5, colors.HexColor("#999999"), "Helvetica", 10, 2, 0),
        # Section cap — 8.5px upper/700 → 6.5pt
        ("B_Cap",     "Normal",  6.5, GRY,  "Helvetica-Bold",   10, 6,  0),
        # KPI valeur — 28px/700 → 21pt
        ("B_KPIVal",  "Normal", 21,   BLK,  "Helvetica-Bold",   24, 2,  0),
        # Badges vigilance/maturité de l'en-tête -- plus petits que les 4 KPI
        # principaux (cohérent avec le dashboard HTML : 17px vs 28px)
        ("B_HdrBadge","Normal", 13,   BLK,  "Helvetica-Bold",   15, 2,  0),
        # KPI label — 10px/700 upper → 7.5pt
        ("B_KPILbl",  "Normal",  7.5, colors.HexColor("#aaaaaa"), "Helvetica", 11, 0, 0),
        # Titre exigence (carte) — 12px/600 → 9pt bold
        ("B_H3",      "Normal",  9,   BLK,  "Helvetica-Bold",   13, 2,  0),
        # Corps carte — 11.5px → 8.5pt
        ("B_Body",    "Normal",  8.5, GRY4, "Helvetica",        13, 2,  0),
        # Corps neutre — 11px td → 8pt
        ("B_TD",      "Normal",  8,   BLK,  "Helvetica",        12, 0,  0),
        # Gras neutre
        ("B_TDBold",  "Normal",  8,   BLK,  "Helvetica-Bold",   12, 0,  0),
        # BID bloc — 10px/700 → 7.5pt bold
        ("B_BID",     "Normal",  7.5, GRY,  "Helvetica-Bold",   11, 0,  0),
        # Statut sous — 9.5px → 7pt
        ("B_SubGry",  "Normal",  7,   colors.HexColor("#aaaaaa"), "Helvetica", 10, 0, 0),
        # Citation — 10.5px italic → 7.8pt italic
        ("B_Cit",     "Normal",  7.8, CIT,  "Helvetica-Oblique", 12, 0,  0),
        # Référence citation — 8px
        ("B_Ref",     "Normal",  6,   GRY,  "Helvetica-Oblique",  9, 2,  0),
        # Badge statut — 7pt bold
        ("B_Badge",   "Normal",  7,   BLK,  "Helvetica-Bold",   10, 0,  0),
        # Pied de page — 9.5px → 7pt italic
        ("B_Foot",    "Normal",  7,   GRY,  "Helvetica-Oblique", 10, 0,  0),
        # Paragraphe réponse — 11.5px italic → 8.5pt
        ("B_Rep",     "Normal",  8.5, GRY4, "Helvetica-Oblique", 13, 2,  8),
        # Question BIM manager — 11.5px → 8pt
        ("B_QBM",     "Normal",  8,   colors.HexColor("#555555"), "Helvetica", 12, 0, 0),
        # Risque si ignoré — texte d'alerte
        ("B_Risque",  "Normal",  8,   colors.HexColor("#7D0000"), "Helvetica", 12, 0, 0),
    ]
    for name, parent, fs, color, font, leading, sa, li in defs:
        kw = dict(parent=S[parent], fontSize=fs, textColor=color,
                  fontName=font, leading=leading, spaceAfter=sa, leftIndent=li)
        S.add(ParagraphStyle(name=name, **kw))
    return S

def p(txt, sty): return Paragraph(html.escape(clean(str(txt))), sty)

def hr_thick():
    return HRFlowable(width="100%", thickness=2, color=colors.HexColor("#111111"),
                      spaceAfter=0, spaceBefore=0)

def hr_thin():
    return HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#111111"),
                      spaceAfter=0, spaceBefore=0)

def hr_light():
    return HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc"),
                      spaceAfter=4, spaceBefore=4)

# TableStyle de base — grille 0.5px #ccc, padding dashboard
TS = [
    ("GRID",         (0,0),(-1,-1), 0.5, colors.HexColor("#cccccc")),
    ("VALIGN",       (0,0),(-1,-1), "TOP"),
    ("LEFTPADDING",  (0,0),(-1,-1), 7),
    ("RIGHTPADDING", (0,0),(-1,-1), 7),
    ("TOPPADDING",   (0,0),(-1,-1), 5),
    ("BOTTOMPADDING",(0,0),(-1,-1), 5),
]

# TableStyle pour en-tête th — fond #f5f5f5
TS_TH = [
    ("BACKGROUND",   (0,0),(-1,0), colors.HexColor("#f5f5f5")),
    ("FONTNAME",     (0,0),(-1,0), "Helvetica-Bold"),
    ("FONTSIZE",     (0,0),(-1,0), 6),
    ("TEXTCOLOR",    (0,0),(-1,0), colors.HexColor("#666666")),
]

def mktbl(data, widths, hdr=True, ex=None):
    t = Table(data, colWidths=widths, repeatRows=1 if hdr else 0)
    ts = list(TS)
    if hdr: ts += TS_TH
    if ex:  ts += ex
    t.setStyle(TableStyle(ts)); return t


# =============================================================================
# FILTRE GENERIQUE -- QUESTIONS BIM MANAGER
# =============================================================================
_STOP_MOTS = {
    "le","la","les","un","une","des","du","de","est","il","y","a","au",
    "pour","sur","quel","quelle","quels","quelles","qui","quoi","comment",
    "combien","dans","avec","par","ce","cette","ces","votre","vous","nous",
    "leur","leurs","son","sa","ses","si","ou","et","en","sont","ont",
    "peut","doit","faut","sera","seront","devez","devra","pouvez",
    "aussi","plus","meme","bien","tout","tous","toute","toutes","dont",
}

def filtrer_question_bm(question: str, conv_pf: str, cctp_pf: str,
                         conv_text: str, lot: str) -> str:
    """
    Generique -- filtre les sous-QUESTIONS (phrases finissant par "?") dont
    la reponse est deja dans les preuves documentaires (conv_pf, cctp_pf,
    conv_text). Les phrases DECLARATIVES (finissant par ".") ne sont jamais
    filtrees : ce sont des elements de contexte injectes via {variables}
    (ex: "La plateforme detectee est {plateforme}.") qui recoupent forcement
    le vocabulaire des documents -- les filtrer sur ce critere supprimerait
    a tort le contexte utile, meme quand la question qui suit reste pertinente.
    Remplace {lot} par le vrai nom du lot.
    Retourne les phrases restantes, ou vide si tout est couvert.
    """
    if not question:
        return ""

    # Injecter le vrai lot
    question = question.replace("{lot}", lot)

    # Corpus documentaire
    doc = (str(conv_pf) + " " + str(cctp_pf) + " " +
           str(conv_text)[:4000]).lower()

    # Decoupe en phrases sur ". " ou "? " -- mais jamais sur une abreviation
    # du type "p." (page) suivie d'un nombre, pour ne pas casser les citations
    # de page ("p.20 de la convention.") au milieu d'une phrase.
    parties = [p.strip() for p in
               re.split(r"(?<=[.?])\s+(?!\d)", question) if p.strip()]
    if not parties:
        parties = [question]

    gardees = []
    for sq in parties:
        if not sq.endswith("?"):
            # Phrase déclarative (contexte) : toujours conservée
            gardees.append(sq)
            continue
        sq_propre = sq.rstrip("?").strip()
        # Mots informatifs (longueur > 3, hors stop words)
        mots = []
        for m in sq_propre.split():
            m_clean = m.lower().strip("?,.:;()!'\"")
            if len(m_clean) > 3 and m_clean not in _STOP_MOTS:
                mots.append(m_clean)
        if not mots:
            gardees.append(sq)
            continue
        # Ratio de couverture documentaire
        n_couverts = sum(1 for m in mots if m in doc)
        taux = n_couverts / len(mots)
        # < 60% couverts -> question pertinente
        if taux < 0.35:
            gardees.append(sq if sq.endswith("?") else sq + "?")

    return " ".join(gardees).strip()

def _mots_cles_pour(detections: Optional[pd.DataFrame], cle_ctx: str) -> List[str]:
    """
    Retourne la liste des mots-clés (colonne mots_cles) de toutes les lignes
    de 35_Detections_Convention portant ce cle_ctx, quel que soit le type.
    Sert de lookup générique pour des listes utilitaires internes (ex:
    libellés de colonnes de tableau à exclure), sans jamais coder ces
    valeurs en dur dans le script.
    """
    if detections is None or detections.empty:
        return []
    out: List[str] = []
    for _, row in detections.iterrows():
        if clean(row.get("cle_ctx", "")) == cle_ctx:
            out.extend(m.lower() for m in split_kw(row.get("mots_cles", "")))
    return out


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
        res[champ] = val[:60].upper() if len(val) <= 3 else val[:60]
    return res


def _extraire_infos_document(pages_textes: list, detections: Optional[pd.DataFrame] = None,
                              metadata_rules: Optional[pd.DataFrame] = None) -> dict:
    """
    Extraction générique depuis n'importe quel PDF de convention/CCTP français.
    Fonctionne avec ou sans page de garde structurée.
    Les motifs de reconnaissance (indice/phase/date/émetteur/MOA/MOE/entreprise
    générale/BIM Manager/Coordinateur BIM) viennent de 36_Extraction_Metadata --
    seule la stratégie de repli pour le titre (heuristique générique : ligne la
    plus longue et significative de la page 1) reste un algorithme en code,
    faute de pouvoir exprimer "la ligne la plus longue" comme une regex simple.
    """
    champs = ["indice", "phase", "date", "emetteur", "maitre_ouvrage", "maitrise_oeuvre",
              "entreprise_generale", "bim_manager", "coordinateur_bim", "numero_affaire"]
    res = {"titre": "", **{c: "" for c in champs}}
    if not pages_textes:
        return res

    texte_global = "\n".join(pages_textes)
    pg1 = pages_textes[0]
    lignes_pg1 = [l.strip() for l in pg1.split("\n") if l.strip()]

    extraits = _extraction_metadata_generique(texte_global, metadata_rules, champs)
    res.update(extraits)

    # ── Structure page de garde (repli générique) : émetteur = première ligne
    # significative après un label "Bureau d'études"/"Maîtrise d'œuvre", si le
    # moteur générique n'a rien trouvé pour l'émetteur.
    if not res["emetteur"]:
        for i, l in enumerate(lignes_pg1):
            if re.match(r"bureau\s+d['\u2019]?\s*[eé]tudes?|"
                        r"ma[iî]tr(?:e|ise)\s+d['\u2019]?\s*(?:oeuvre|œuvre)", l, re.I):
                for s in lignes_pg1[i+1:i+8]:
                    s = s.strip("_").strip("-").strip()
                    # Exclure adresses, tirets, codes postaux, MO déjà trouvé
                    if (len(s) > 3 and
                        not re.match(r"^\d+\s+(?:impasse|rue|avenue|bd)", s, re.I) and
                        not re.match(r"^[\-_\s]+$", s) and
                        not re.match(r"^\d{4,5}\s", s) and
                        "cedex" not in s.lower() and
                        (not res["maitre_ouvrage"] or
                         s.lower() != res["maitre_ouvrage"].lower()) and
                        not re.match(r"ma[iî]tre?\s+d", s, re.I)):
                        res["emetteur"] = s[:60]
                        break
                break
    # MO depuis structure
    if not res["maitre_ouvrage"]:
        for i, l in enumerate(lignes_pg1):
            if re.match(r"ma[iî]tre?\s+d['\u2019]?\s*ouvrage", l, re.I):
                for s in lignes_pg1[i+1:i+5]:
                    s = s.strip()
                    if len(s) > 2 and not re.match(r"^[\-_]+$", s):
                        res["maitre_ouvrage"] = s[:60]; break
                break

    # ── 5. En-têtes pages 2-4 — numéro d'affaire + indice uniquement ────
    # NE PAS chercher l'émetteur ici : les en-têtes répétés contiennent
    # le nom du client/projet, pas l'émetteur → source de faux positifs.
    # Si la page 1 n'a pas d'émetteur lisible (logo image), rester vide.
    for pg in pages_textes[1:4]:
        entete = pg[:300]
        if not res["numero_affaire"]:
            m_n = re.search(r"N[°o\.] *:? *([\.\.\d\-]{4,15})", entete)
            if m_n: res["numero_affaire"] = m_n.group(1)
        # Indice — pattern strict "Indice : B" uniquement (pas "DCE-B" free-text)
        if not res["indice"]:
            m_i = re.search(r"Indice\s*:?\s*([A-E])\b", entete, re.IGNORECASE)
            if m_i: res["indice"] = m_i.group(1).upper()


    # ── 6. Titre ─────────────────────────────────────────────────────
    # Mots typiques de titres de chapitres BIM qui ne sont PAS des titres de document
    _mots_chapitre = (
        "codification", "nomenclature", "classification", "processus",
        "exigences", "livrables", "réunion", "compte rendu",
        "organisation", "responsabilités", "protocole", "procédure",
        "sommaire", "table des", "annexe", "glossaire", "définition"
    )
    def _ok(l):
        ll = l.lower()
        return (len(l) > 12 and
                not re.match(r"^[\d\s\-\./,_°]+$", l) and
                not re.match(r"^\d+\s+(?:impasse|rue|avenue|boulevard)", l, re.I) and
                "cedex" not in ll and "@" not in l and
                not re.match(
                    r"^(?:phase|indice|date|objet|version|rédacteur|relecture|"
                    r"bureau|maitre|ma[iî]tre|émetteur|établi|rédigé|auteur|"
                    r"n°|ref\.|réf\.|page)", l, re.I) and
                not re.search(
                    r"(?:ESQ|APS|APD|PRO|DCE|EXE|DOE)\s+[A-E]\s+\d{1,2}\s+\w+\s+\d{4}", l, re.I) and
                # Exclure titres de chapitres typiques BIM
                not any(mot in ll for mot in _mots_chapitre))
    cands = [l for l in lignes_pg1[:30] if _ok(l)]
    if cands:
        res["titre"] = max(cands, key=len)[:80]

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
            start = max(0, i - 60)
            extrait = txt[start:i + len(kw_matched) + 60].strip()
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


def _extraire_details_convention(bid: str, conv_pf: str, conv_pg: str,
                                  cctp_pf: str, cctp_pg: str,
                                  conv_text: str, lot: str,
                                  conv_pages: Optional[List[Tuple[int, str]]] = None,
                                  gloss: Optional[pd.DataFrame] = None,
                                  detections: Optional[pd.DataFrame] = None,
                                  metadata_rules: Optional[pd.DataFrame] = None) -> Dict:
    """Extrait les détails contextuels détectés dans les documents pour un bloc."""
    ctx: Dict = {
        "plateforme": "",
        "plateforme_page": "",
        "plateforme_extrait": "",
        "cout_plateforme": "",
        "logiciel": "",
        "logiciel_page": "", "logiciel_extrait": "",
        "format_ifc": "",
        "format_ifc_page": "", "format_ifc_extrait": "",
        "lod": "",
        "lod_page": "", "lod_extrait": "",
        "nd": "",
        "nd_page": "", "nd_extrait": "",
        "frequence": "",
        "section_cctp": "",
        "niveau_bim": "",
        "niveau_bim_page": "", "niveau_bim_extrait": "",
        "dim_4d": False,
        "dim_5d": False,
        "dim_6d": False,
        "dim_7d": False,
        "bim_manager": "",
        "bim_manager_page": "", "bim_manager_extrait": "",
        "coordinateur_bim": "",
        "coordinateur_bim_page": "", "coordinateur_bim_extrait": "",
        "geo_referencement": False,
        "clash_3d": False,
        "doe_numerique": False,
        "formats_livrables": "",
        "contractuel": False,
        "indice_doc": "",
        "date_doc": "",
        "emetteur_doc": "",
        "maitre_ouvrage_doc": "",
        "numero_affaire": "",
        "titre_doc": "",
        "conv_extrait": conv_pf[:200] if conv_pf else "",
        "conv_page": conv_pg,
        "cctp_extrait": cctp_pf[:200] if cctp_pf else "",
        "cctp_page": cctp_pg,
    }
    txt = conv_text.lower()
    # Plateforme CDE — liste construite dynamiquement depuis 40_Glossaire
    # (lignes categorie == "Plateforme CDE"), plus un repli générique
    # ("ECC", "environnement commun de données"...) si aucune marque précise
    # n'est nommée. Ajouter une plateforme = une ligne dans l'Excel, jamais
    # de modification de code.
    for variantes, nom_affichage in _plateformes_depuis_glossaire(gloss):
        matched = next((v for v in variantes if contains(txt, v)), None)
        if matched:
            ctx["plateforme"] = nom_affichage
            ctx["plateforme_page"], ctx["plateforme_extrait"] = _localiser_page(conv_pages, [matched])
            break
    if not ctx["plateforme"]:
        for p in ["ecc", "environnement commun de données", "environnement de collaboration"]:
            if contains(txt, p):
                # "CDE" / "ECC" sont des noms GÉNÉRIQUES (le type de plateforme), pas le nom
                # d'une solution précise (KROQI, ACC, ProjectWise...). Ne jamais les afficher
                # comme si c'était le nom de la plateforme détectée -- cela produit des phrases
                # absurdes ("la plateforme CDE détectée est CDE"). On l'indique explicitement
                # comme non nommée.
                nom = "un CDE (Environnement Commun de Données), sans nom de solution précisé"
                ctx["plateforme"] = nom
                ctx["plateforme_page"], ctx["plateforme_extrait"] = _localiser_page(conv_pages, [p])
                break
    if not ctx["plateforme"] and any(contains(txt, k) for k in
            ["plateforme d'échange", "plateforme collaborative", "une plateforme",
             "espace dédié pour chaque intervenant", "plateforme de collaboration"]):
        ctx["plateforme"] = "la plateforme collaborative du projet"
        ctx["plateforme_page"], ctx["plateforme_extrait"] = _localiser_page(conv_pages,
            ["plateforme d'échange", "plateforme collaborative", "plateforme de collaboration"])
    # Coût plateforme
    m = re.search(r"(\d{2,4})\s*€\s*(?:par|/)\s*corps", txt)
    if m: ctx["cout_plateforme"] = m.group(1) + " €"
    # Logiciel BIM, fréquence, géoréférencement, clash 3D, DOE numérique,
    # dimensions 4D-7D, caractère contractuel, formats livrables -- tout
    # piloté depuis 35_Detections_Convention (Excel), aucune liste en dur.
    _appliquer_detections(ctx, txt, conv_pages, detections)
    # Format IFC
    m2 = re.search(r"ifc\s*2x3|ifc\s*4|ifc4|ifc2x3", txt)
    if m2:
        ctx["format_ifc"] = m2.group(0).upper().replace(" ","")
        ctx["format_ifc_page"], ctx["format_ifc_extrait"] = _localiser_page(conv_pages, [m2.group(0)])
    # LOD / ND
    m3 = re.search(r"lod\s*(\d{3})", txt)
    if m3:
        ctx["lod"] = "LOD " + m3.group(1)
        ctx["lod_page"], ctx["lod_extrait"] = _localiser_page(conv_pages, [m3.group(0)])
    mnd = re.search(r"nd\s*([3-5])", txt)
    if mnd:
        ctx["nd"] = "ND" + mnd.group(1)
        ctx["nd_page"], ctx["nd_extrait"] = _localiser_page(conv_pages, [mnd.group(0)])
    # Niveau BIM — regex générique, insensible à l'ordre des mots (« Niveau BIM 3 »
    # ou « BIM Niveau 3 ») et à la casse
    m_niv = re.search(r"(?:bim\s+)?niveau\s*(?:bim\s*)?([1-3])(?:\s+bim)?", txt, re.IGNORECASE)
    if m_niv:
        ctx["niveau_bim"] = "NIVEAU " + m_niv.group(1)
        ctx["niveau_bim_page"], ctx["niveau_bim_extrait"] = _localiser_page(conv_pages, [m_niv.group(0).lower()])
    # BIM Manager / Coordinateur BIM -- délégué au même moteur générique que
    # _extraire_infos_document() (règles dans 36_Extraction_Metadata), pour
    # ne jamais dupliquer les mêmes motifs à deux endroits du code.
    _bm_cb = _extraction_metadata_generique(conv_text, metadata_rules, ["bim_manager", "coordinateur_bim"])
    if _bm_cb.get("bim_manager"):
        ctx["bim_manager"] = _bm_cb["bim_manager"]
        ctx["bim_manager_page"], ctx["bim_manager_extrait"] = _localiser_page(conv_pages, [_bm_cb["bim_manager"].lower()])
    if _bm_cb.get("coordinateur_bim"):
        ctx["coordinateur_bim"] = _bm_cb["coordinateur_bim"]
        ctx["coordinateur_bim_page"], ctx["coordinateur_bim_extrait"] = _localiser_page(conv_pages, [_bm_cb["coordinateur_bim"].lower()])
    # Section CCTP
    m4 = re.search(r"§\s*(\d+[\.\d]*)", cctp_pf or "")
    if m4: ctx["section_cctp"] = "§" + m4.group(1)

    # Note : les métadonnées de page de garde (indice, date, émetteur) sont extraites
    # depuis les pages brutes fitz dans main() via _conv_infos / _cctp_infos.
    # Cette fonction gère uniquement les détections dynamiques (plateforme, LOD, etc.)
    return ctx


def _questions_programmatiques(bid: str, titre: str, statut: str,
                                ctx: Dict, lot: str) -> List[Dict]:
    """
    Génère 2 questions génériques contextualisées à partir des éléments
    réellement détectés dans les documents (ctx), SANS AUCUNE dépendance à
    un identifiant de bloc précis (bid). Fonctionne pour n'importe quel bloc
    défini dans 10_Blocs -- y compris un bloc ajouté ou renommé ultérieurement.

    Utilisée uniquement en dernier recours : si 50_Questions (Excel) ne
    contient aucune question pour ce bloc, et si le mode API n'est pas
    utilisé ou a échoué. La source de vérité pour des questions précises
    et métier reste la feuille 50_Questions -- c'est là qu'il faut enrichir
    le questionnaire, pas dans le code.
    """
    ref = f"p. {ctx['conv_page']}" if ctx.get("conv_page") else "dans la convention"
    extrait = clean(ctx.get("conv_extrait") or ctx.get("cctp_extrait") or "")
    detail = (f" La documentation précise : « {extrait[:140]}"
              f"{'…' if len(extrait) > 140 else ''} » ({ref})." if extrait else "")

    # Éléments techniques génériques détectés (aucun lien avec un bloc précis) :
    # on les mentionne s'ils existent, sans jamais présumer lesquels s'appliquent.
    elements = []
    if ctx.get("plateforme"):  elements.append(f"la plateforme {ctx['plateforme']}")
    if ctx.get("logiciel"):    elements.append(f"le logiciel {ctx['logiciel']}")
    if ctx.get("format_ifc"):  elements.append(f"le format {ctx['format_ifc']}")
    if ctx.get("lod"):         elements.append(f"le {ctx['lod']}")
    contexte_tech = f" Cela peut impliquer notamment {', '.join(elements)}." if elements else ""

    qs: List[Dict] = [
        {
            "question": (
                f"L'exigence « {titre} » a été détectée ({statut.lower()}) "
                f"pour le lot {lot}.{detail}{contexte_tech} "
                f"Disposez-vous des compétences, outils ou ressources internes "
                f"pour y répondre ?"
            ),
            "indice_0": "Non. Nous n'avons pas les ressources ou l'expérience pour cette exigence.",
            "indice_1": "Partiellement. Nous pouvons répondre avec un accompagnement extérieur ou une montée en compétence.",
            "indice_2": "Oui. Nous maîtrisons cette exigence et pouvons y répondre sans aide extérieure.",
            "poids": 2,
        },
        {
            "question": (
                f"Pouvez-vous désigner dès maintenant une personne responsable "
                f"du suivi de cette exigence pour le lot {lot} ?"
            ),
            "indice_0": "Non. Personne n'est identifié pour cette tâche.",
            "indice_1": "Partiellement. Une personne pourrait le faire mais n'est pas encore formée.",
            "indice_2": "Oui. Un responsable est identifié et opérationnel immédiatement.",
            "poids": 1,
        },
    ]
    return qs



def _questions_via_api(bid: str, titre: str, statut: str, ctx: Dict,
                        lot: str, api_key: str) -> Optional[List[Dict]]:
    """
    Génère des questions via l'API Anthropic (mode enrichi optionnel).
    Retourne None si l'appel échoue → fallback programmatique.
    """
    try:
        import urllib.request, json as _json
        prompt = (
            f"Tu es un expert BIM qui aide des TPE/PME du BTP à répondre à des appels d'offres.\n\n"
            f"Bloc d'exigence : {bid} -- {titre}\n"
            f"Statut détecté : {statut}\n"
            f"Lot concerné : {lot}\n"
            f"Extrait convention BIM (p.{ctx['conv_page']}) : {ctx['conv_extrait']}\n"
            f"Extrait CCTP (p.{ctx['cctp_page']}) : {ctx['cctp_extrait']}\n"
            f"Plateforme CDE détectée : {ctx['plateforme'] or 'non détectée'}\n"
            f"Logiciel BIM : {ctx['logiciel'] or 'non détecté'}\n"
            f"Format IFC : {ctx['format_ifc'] or 'non détecté'}\n\n"
            f"Génère exactement 2 questions d'auto-évaluation contextualisées "
            f"(basées sur les extraits réels ci-dessus, pas des questions génériques). "
            f"Chaque question doit avoir 3 niveaux de réponse (0=Non, 1=Partiel, 2=Oui) "
            f"et un poids (1 ou 2).\n"
            f"Réponds UNIQUEMENT en JSON valide, sans texte autour, au format :\n"
            f'[{{"question":"...","indice_0":"...","indice_1":"...","indice_2":"...","poids":2}}]'
        )
        body = _json.dumps({
            "model": "claude-sonnet-4-6",
            "max_tokens": 800,
            "messages": [{"role": "user", "content": prompt}]
        }).encode()
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=body,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = _json.loads(resp.read())
        raw = data["content"][0]["text"].strip()
        # Nettoyer les éventuels backticks markdown
        raw = re.sub(r"^```json\s*|^```\s*|```$", "", raw, flags=re.MULTILINE).strip()
        parsed = _json.loads(raw)
        if isinstance(parsed, list) and parsed:
            return parsed
    except Exception as e:
        print(f"  ⚠ API Anthropic indisponible ({e}) -- questions programmatiques utilisées.")
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# QUESTIONNAIRE INTERACTIF CONTEXTUALISÉ
# ═══════════════════════════════════════════════════════════════════════════════
REPS = {"0":0,"n":0,"non":0,"jamais":0,
        "1":1,"p":1,"partiel":1,"partiellement":1,
        "2":2,"o":2,"oui":2,"maitrise":2,"maîtrisé":2}


class _SafeCtxDict(dict):
    """Laisse une {variable} inconnue telle quelle plutôt que de lever une erreur."""
    def __missing__(self, key):
        return "{" + key + "}"


def _sub_ctx(txt: str, ctx_full: Dict) -> str:
    """
    Remplace les {variable} (ex: {plateforme}, {cout}, {logiciel}, {lot}...)
    dans un texte lu depuis l'Excel (50_Questions) par les valeurs réellement
    détectées dans les documents. Une valeur absente/vide du contexte est
    remplacée par une formulation neutre plutôt que de casser l'affichage.
    Aucune valeur métier n'est codée en dur ici -- tout vient de ctx_full,
    construit depuis _extraire_details_convention() + le lot courant.
    """
    if not txt or "{" not in txt:
        return txt
    defauts = {
        "plateforme":    "la plateforme CDE du projet",
        "cout":          ctx_full.get("cout_plateforme") or "un coût à préciser",
        "cout_plateforme": ctx_full.get("cout_plateforme") or "un coût à préciser",
        "logiciel":      "un logiciel BIM",
        "fmt_ifc":       ctx_full.get("format_ifc") or "IFC",
        "format_ifc":    ctx_full.get("format_ifc") or "IFC",
        "lod":           "le LOD défini dans la convention",
        "nd":            ctx_full.get("nd") or "le niveau de développement défini",
        "niveau_bim":    "Niveau 2",
        "frequence":     "régulièrement",
        "ref_conv":      f"p. {ctx_full['conv_page']}" if ctx_full.get("conv_page") else "dans la convention",
        "ref_cctp":      f"p. {ctx_full['cctp_page']}" if ctx_full.get("cctp_page") else "dans le CCTP",
        "ref_plateforme": f"p. {ctx_full['plateforme_page']} de la convention" if ctx_full.get("plateforme_page") else "dans la convention",
        "lot":           "",
    }
    valeurs = {**defauts, **{k: v for k, v in ctx_full.items() if v}}
    try:
        return txt.format_map(_SafeCtxDict(valeurs))
    except Exception:
        return txt


def poser_questionnaire(blocs: pd.DataFrame, lot: str,
                        resultats_lecture: Dict,
                        questions: pd.DataFrame = None,
                        api_key: str = "",
                        gloss: pd.DataFrame = None,
                        detections: pd.DataFrame = None,
                        metadata_rules: pd.DataFrame = None,
                        reponses_fournies: Optional[Dict[str, int]] = None) -> Dict:
    """
    Génère et pose des questions contextualisées à partir des extraits
    réellement trouvés dans les documents (convention + CCTP).
    Les questions sont lues depuis la feuille 50_Questions de l'Excel (questions DataFrame).
    Toute ligne ajoutée dans 50_Questions sera automatiquement posée.
    Mode API optionnel si api_key fournie — utilisé si 50_Questions vide pour un bloc.
    """
    # La capacité de l'entreprise est indépendante du statut contractuel.
    # On évalue donc tous les blocs non exclus qui possèdent des questions
    # actives dans 50_Questions, y compris les blocs dont la preuve
    # contractuelle n'est pas encore démontrée.
    question_block_ids = set()
    if questions is not None and not questions.empty:
        question_block_ids = {
            clean(value) for value in questions["id_bloc"].astype(str).tolist()
            if clean(value)
        }
    # La capacité de l'entreprise reste mesurée même lorsqu'une exigence est
    # contractuellement exclue : cela permet au radar de représenter l'ensemble
    # du questionnaire sans confondre applicabilité et capacité.
    blocs_applicables = {
        bid for bid in resultats_lecture
        if (not question_block_ids or bid in question_block_ids)
    }
    titres = {clean(b["id_bloc"]): clean(b["titre_bloc"])
              for _, b in blocs.iterrows()}

    if not blocs_applicables:
        print("\nAucun bloc non exclu avec questions actives. Questionnaire ignoré.")
        return {}

    resultats = {}
    print()
    print("=" * 70)
    print(f"  AUTO-ÉVALUATION BIM -- Lot : {lot}")
    print("  Questions générées à partir de la convention et du CCTP analysés.")
    print("  Répondez : 0 (Non)  /  1 (Partiel)  /  2 (Oui)")
    if api_key:
        print("  Mode enrichi : questions générées par l'API Anthropic.")
    print("=" * 70)

    # Ordre naturel croissant : B01, B02, ..., B06A, B06B, B10, etc.
    blocs_ordonnes = sorted(blocs_applicables, key=block_sort_key)

    for num, bid in enumerate(blocs_ordonnes, 1):
        res   = resultats_lecture[bid]
        titre = titres.get(bid, bid)

        # Extraire le contexte documentaire réel
        ctx = _extraire_details_convention(
            bid,
            res.get("conv_pf", ""), res.get("conv_pg", ""),
            res.get("cctp_pf", ""), res.get("cctp_pg", ""),
            res.get("_conv_text", ""),
            lot, gloss=gloss, detections=detections, metadata_rules=metadata_rules
        )

        # ── Générer les questions depuis 50_Questions (Excel) ───────────────────
        # Toute ligne ajoutée dans la feuille 50_Questions est lue automatiquement.
        # Les {variables} (ex: {plateforme}, {cout}) sont substituées avec le
        # contexte réellement détecté dans les documents pour ce bloc.
        # Si aucune question Excel pour ce bloc → fallback API ou programmatique.
        qs_excel = []
        ctx_sub = {**ctx, "lot": lot}
        if questions is not None and not questions.empty:
            rows_q = questions[
                questions["id_bloc"].astype(str).str.strip() == bid
            ]
            # Lire chaque ligne active de la feuille 50_Questions
            for _, qrow in rows_q.iterrows():
                actif = str(qrow.get("actif", "oui")).strip().lower()
                if actif in ("non", "false", "0"):
                    continue
                qs_excel.append({
                    "id_question": clean(qrow.get("id_question", f"{bid}_Q{len(qs_excel)+1}")),
                    "question":  _sub_ctx(str(qrow.get("question",  "")), ctx_sub),
                    "indice_0":  _sub_ctx(str(qrow.get("indice_0",  "Non.")), ctx_sub),
                    "indice_1":  _sub_ctx(str(qrow.get("indice_1",  "Partiellement.")), ctx_sub),
                    "indice_2":  _sub_ctx(str(qrow.get("indice_2",  "Oui.")), ctx_sub),
                    "poids":     int(float(qrow.get("poids", 1) or 1)),
                })

        if qs_excel:
            # Questions lues depuis l'Excel — source de vérité
            qs = qs_excel
        elif api_key:
            # Pas de questions Excel pour ce bloc → essayer l'API
            qs = _questions_via_api(bid, titre, res["statut"], ctx, lot, api_key)
            if qs is None:
                qs = _questions_programmatiques(bid, titre, res["statut"], ctx, lot)
        else:
            # Fallback programmatique (questions génériques contextualisées)
            qs = _questions_programmatiques(bid, titre, res["statut"], ctx, lot)

        print(f"\n{'─'*70}")
        print(f"  Bloc {num}/{len(blocs_ordonnes)} -- {titre}")
        print(f"  Statut : {res['statut']}")
        if ctx["conv_extrait"]:
            extrait_court = ctx["conv_extrait"][:120].replace("\n", " ")
            print(f"  Ref. convention p.{ctx['conv_page']} : \u00ab {extrait_court}... \u00bb")
        if ctx["cctp_extrait"]:
            extrait_court = ctx["cctp_extrait"][:120].replace("\n", " ")
            print(f"  Ref. CCTP p.{ctx['cctp_page']} : \u00ab {extrait_court}... \u00bb")
        print(f"{'─'*70}")

        score, score_max, reps_bloc = 0, 0, []
        for qi, q in enumerate(qs, 1):
            qtxt  = str(q.get("question", ""))
            ind0  = str(q.get("indice_0", "Non."))
            ind1  = str(q.get("indice_1", "Partiellement."))
            ind2  = str(q.get("indice_2", "Oui."))
            poids = int(q.get("poids", 1) or 1)

            print(f"\n  Q{qi}. {textwrap.fill(qtxt, 66, subsequent_indent='      ')}")
            print(f"     0 -- {textwrap.fill(ind0, 60, subsequent_indent='         ')}")
            print(f"     1 -- {textwrap.fill(ind1, 60, subsequent_indent='         ')}")
            print(f"     2 -- {textwrap.fill(ind2, 60, subsequent_indent='         ')}")

            qid = clean(q.get("id_question", f"{bid}_Q{qi}"))
            if reponses_fournies is not None:
                if qid not in reponses_fournies:
                    raise ValueError(f"Réponse manquante pour la question {qid}")
                try:
                    val = int(reponses_fournies[qid])
                except (TypeError, ValueError) as exc:
                    raise ValueError(f"Réponse invalide pour {qid} : 0, 1 ou 2 attendu") from exc
                if val not in (0, 1, 2):
                    raise ValueError(f"Réponse invalide pour {qid} : 0, 1 ou 2 attendu")
                print(f"     Réponse web : {val}")
            else:
                while True:
                    rep = input("     Votre réponse [0/1/2] : ").strip().lower()
                    if rep in REPS: val = REPS[rep]; break
                    print("     → Entrez 0, 1 ou 2.")

            score     += val * poids
            score_max += 2 * poids
            reps_bloc.append({
                "id_question": qid,
                "question": qtxt,
                "reponse_val": val,
                "reponse_label": [ind0, ind1, ind2][val],
                "poids": poids,
            })

        pct    = round(score / score_max * 100) if score_max else 0
        niveau = "Confirmé" if pct >= 70 else ("Intermédiaire" if pct >= 40 else "Débutant")
        resultats[bid] = {
            "titre": titre, "score": score,
            "max": score_max, "pct": pct,
            "niveau": niveau, "reponses": reps_bloc,
        }
        print(f"\n  ✓ {bid} : {score}/{score_max} ({pct}%) -- {niveau}")

    return resultats

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


def _regle_entreprise_ok(operateur: str, valeur_reelle: object, valeur_attendue: str) -> bool:
    op = clean(operateur).lower()
    reel = clean(valeur_reelle)
    if op == "non_vide":
        return bool(reel)
    if op == "equals":
        return norm(reel) == norm(valeur_attendue)
    if op == "contains_any":
        return any(norm(v) in norm(reel) for v in valeur_attendue.split("|") if clean(v))
    return False


def _regle_reponse_ok(operateur: str, valeur_reelle: int, valeur_attendue: int) -> bool:
    op = clean(operateur)
    return {"=": valeur_reelle == valeur_attendue, "<=": valeur_reelle <= valeur_attendue,
            ">=": valeur_reelle >= valeur_attendue, "<": valeur_reelle < valeur_attendue,
            ">": valeur_reelle > valeur_attendue}.get(op, False)


def detecter_contradictions(coherence: Optional[pd.DataFrame], entreprise: Dict[str, str],
                             scores_eval: Dict, resultats_lecture: Dict) -> List[Dict]:
    """
    Applique 55_Regles_Coherence : compare un champ de 70_Entreprise à une
    réponse précise du questionnaire (pas juste la moyenne du bloc). Une
    contradiction trouvée est à la fois retournée globalement et rattachée
    au bloc concerné (resultats_lecture[bid]["contradictions"]), pour bloquer
    la génération du texte à copier sur ce bloc précis. Aucune règle codée
    en dur -- tout vient de l'Excel.
    """
    if coherence is None or coherence.empty:
        return []
    reponses = flatten_answers(scores_eval)
    trouvees: List[Dict] = []
    for _, regle in coherence.iterrows():
        actif = clean(regle.get("actif", "Oui")).lower()
        if actif in ("non", "false", "0"):
            continue
        qid = clean(regle.get("id_question", ""))
        if qid not in reponses:
            continue
        champ = clean(regle.get("champ_entreprise", ""))
        cond_e = _regle_entreprise_ok(regle.get("operateur_entreprise", ""),
                                       entreprise.get(champ, ""), clean(regle.get("valeurs_entreprise", "")))
        cond_r = _regle_reponse_ok(regle.get("operateur_reponse", ""),
                                    reponses[qid], int(float(regle.get("valeur_reponse", 0) or 0)))
        if cond_e and cond_r:
            item = {
                "id_regle": clean(regle.get("id_regle", "")),
                "gravite": clean(regle.get("gravite", "")),
                "penalite": int(float(regle.get("penalite", 0) or 0)),
                "message": clean(regle.get("message", "")),
                "id_question": qid,
            }
            trouvees.append(item)
            bid = qid.split("_Q", 1)[0]
            if bid in resultats_lecture:
                resultats_lecture[bid].setdefault("contradictions", []).append(item)
    return trouvees


def calculer_indice_fiabilite(resultats_lecture: Dict, textes_reponse: Dict,
                               contradictions: List[Dict], params: Dict[str, object]) -> Tuple[int, Dict[str, int]]:
    """
    Indice de fiabilité = moyenne pondérée de 4 composantes mesurables
    (poids dans 05_Parametres_Moteur), plutôt qu'un simple ratio de
    complétude qui ne mesure ni l'exactitude ni la cohérence :
      - exactitude_contractuelle : % de blocs pertinents avec une citation réelle
      - coherence     : 100 - somme des pénalités de contradiction
      - preuves       : % de textes générés dont les conditions sont satisfaites
      - couverture    : % de blocs applicables ayant un texte généré
    """
    pertinents = [r for r in resultats_lecture.values() if r.get("presence_contractuelle") != "ABSENTE"]
    if pertinents:
        preuve_ok = sum(1 for r in pertinents if r.get("conv_pf") or r.get("cctp_pf"))
        exactitude = round(preuve_ok / len(pertinents) * 100)
    else:
        exactitude = 100

    coherence = max(0, 100 - sum(c.get("penalite", 0) for c in contradictions))

    tentatives = list(textes_reponse.values())
    if tentatives:
        ok = sum(1 for v in tentatives if v.get("conditions_ok") and v.get("texte"))
        preuves = round(ok / len(tentatives) * 100)
    else:
        preuves = 100 if not pertinents else 0

    applicables_gen = set(p_list(params, "applicabilites_generation")) or {"APPLICABLE", "PROBABLE", "A_CONFIRMER"}
    ids_applicables = [bid for bid, r in resultats_lecture.items()
                        if r.get("presence_contractuelle") != "ABSENTE"
                        and r.get("applicabilite_lot") in applicables_gen]
    if ids_applicables:
        couverts = sum(1 for bid in ids_applicables if textes_reponse.get(bid, {}).get("texte"))
        couverture = round(couverts / len(ids_applicables) * 100)
    else:
        couverture = 100

    composantes = {"exactitude_contractuelle": exactitude, "coherence": coherence,
                   "preuves": preuves, "couverture": couverture}
    poids = {
        "exactitude_contractuelle": p_int(params, "poids_qualite_contractuelle", 30),
        "coherence": p_int(params, "poids_qualite_coherence", 30),
        "preuves": p_int(params, "poids_qualite_preuves", 20),
        "couverture": p_int(params, "poids_qualite_couverture", 20),
    }
    total_poids = sum(poids.values()) or 100
    indice = round(sum(composantes[k] * poids[k] for k in composantes) / total_poids)
    return indice, composantes


def p_list(params: Dict[str, object], cle: str) -> List[str]:
    v = params.get(cle, [])
    return v if isinstance(v, list) else split_kw(v)


def generate_radar(scores: Dict, lot: str) -> io.BytesIO:
    blocs = list(scores.keys())
    # Labels depuis l'Excel uniquement — troncature simple, aucun replace codé en dur
    labels = []
    for bid in blocs:
        t = scores[bid]["titre"].split("(")[0].split("—")[0].strip()[:22]
        labels.append(f"{bid}\n{t}")

    vals = [scores[b]["pct"] / 100 for b in blocs]
    N = len(blocs)
    angles = [n / N * 2 * np.pi for n in range(N)] + [0]
    vals_plot = vals + vals[:1]

    fig, ax = plt.subplots(figsize=(6.5,6.5), subplot_kw=dict(polar=True))
    ax.set_facecolor("#F8F9FA"); fig.patch.set_facecolor("white")
    ax.set_rlabel_position(30)
    ax.set_yticks([.25,.5,.75,1]); ax.set_yticklabels(["25%","50%","75%","100%"],size=7,color="#888")
    ax.set_ylim(0,1)
    ax.fill(angles,[1]*len(angles),color="#FCE4D6",alpha=0.25)
    ax.fill(angles,[0.7]*len(angles),color="#FFF2CC",alpha=0.35)
    ax.fill(angles,[0.4]*len(angles),color="#E2EFDA",alpha=0.45)
    ax.plot(angles, vals_plot,'o-',linewidth=2,color="#1F4E78",markersize=5)
    ax.fill(angles, vals_plot,alpha=0.25,color="#2E75B6")
    ax.set_xticks(angles[:-1]); ax.set_xticklabels(labels,size=7.5,color="#1F4E78",fontweight="bold")
    patches = [mpatches.Patch(color="#E2EFDA",label="≥ 40% -- Intermédiaire"),
               mpatches.Patch(color="#FFF2CC",label="≥ 70% -- Confirmé")]
    ax.legend(handles=patches,loc="upper right",bbox_to_anchor=(1.35,1.1),fontsize=7)
    ax.set_title(f"Maturité BIM -- Lot {lot}",size=11,color="#1F4E78",fontweight="bold",pad=20)
    buf = io.BytesIO()
    plt.tight_layout(); plt.savefig(buf,format="png",dpi=150,bbox_inches="tight"); plt.close()
    buf.seek(0); return buf

# ═══════════════════════════════════════════════════════════════════════════════
# SÉLECTION TEXTE DE RÉPONSE
# ═══════════════════════════════════════════════════════════════════════════════
def get_ligne_reponse(textes: pd.DataFrame, bid: str, pct: int) -> Optional[pd.Series]:
    """Retourne la ligne 60_Reponses complète (toutes colonnes) pour ce bloc/pct."""
    cands = textes[textes["id_bloc"].astype(str).str.strip() == bid]
    for _, r in cands.iterrows():
        smin = int(r.get("pct_min", r.get("score_min", 0)) or 0)
        smax = int(r.get("pct_max", r.get("score_max", 100)) or 100)
        if smin <= pct <= smax:
            return r
    return cands.iloc[-1] if not cands.empty else None


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


def get_texte(textes: pd.DataFrame, bid: str, pct: int) -> Tuple[str,str]:
    cands = textes[textes["id_bloc"].astype(str).str.strip() == bid]
    for _, r in cands.iterrows():
        smin = int(r.get("pct_min", r.get("score_min", 0)) or 0)
        smax = int(r.get("pct_max", r.get("score_max", 100)) or 100)
        if smin <= pct <= smax:
            return clean(r["texte_reponse"]), clean(r["niveau_label"])
    if not cands.empty:
        last = cands.iloc[-1]
        return clean(last["texte_reponse"]), clean(last.get("niveau_label",""))
    return "", ""


def preparer_textes_reponse(
    textes: pd.DataFrame,
    scores_eval: Dict,
    resultats_lecture: Dict,
    entreprise: str,
    lot: str,
    messages: Dict[str, str],
    engine_params: Dict[str, object],
) -> Dict[str, Dict]:
    """Construit une seule fois les engagements proposés pour tous les livrables.

    Les mêmes règles de prudence sont ainsi utilisées dans le dashboard, le
    rapport de synthèse, le plan d'actions et l'annexe technique.
    """
    if not scores_eval:
        return {}
    applicabilites = set(p_list(engine_params, "applicabilites_generation")) or {
        "APPLICABLE", "PROBABLE", "A_CONFIRMER"
    }
    reponses_brutes = flatten_answers(scores_eval)
    resultat: Dict[str, Dict] = {}
    nom = entreprise or "L'entreprise"
    lot_val = lot or "ce lot"

    def nettoie(t: str) -> str:
        return (t.replace("[Nom de l'entreprise]", nom).replace("[nom de l'entreprise]", nom)
                 .replace("{entreprise}", nom).replace("[lot]", lot_val).replace("{lot}", lot_val)
                 .replace("du projet du projet", "du projet")
                 .replace("la plateforme la plateforme", "la plateforme")
                 .replace("au workflow la plateforme", "au workflow de la plateforme")
                 .replace("aux règles la plateforme", "aux règles de la plateforme"))

    for bid, sc in scores_eval.items():
        res_bid = resultats_lecture.get(bid, {})
        if res_bid.get("presence_contractuelle") == "ABSENTE":
            resultat[bid] = {
                "texte": "", "niveau": "", "statut_generation": "bloqué",
                "avertissement": render_message(messages, "paragraphe_presence_absente"),
            }
            continue
        if res_bid.get("applicabilite_lot") not in applicabilites:
            resultat[bid] = {
                "texte": "", "niveau": "", "statut_generation": "bloqué",
                "avertissement": render_message(messages, "paragraphe_bloque"),
            }
            continue
        if res_bid.get("contradictions"):
            resultat[bid] = {
                "texte": "", "niveau": "", "statut_generation": "bloqué",
                "avertissement": render_message(messages, "paragraphe_contradiction"),
            }
            continue

        ligne = get_ligne_reponse(textes, bid, sc.get("pct", 0))
        if ligne is None:
            continue
        cond_ok = _evaluer_condition_simple(ligne.get("conditions_questions", ""), reponses_brutes)
        niveau_eng = clean(ligne.get("niveau_engagement", "")) or (
            "ferme" if sc.get("pct", 0) >= 70 else "conditionnel"
        )
        if cond_ok:
            resultat[bid] = {
                "texte": nettoie(clean(ligne.get("texte_reponse", ""))),
                "niveau": clean(ligne.get("niveau_label", "")),
                "statut_generation": niveau_eng,
                "conditions_ok": True,
            }
        else:
            secours = clean(ligne.get("texte_secours", ""))
            if secours:
                resultat[bid] = {
                    "texte": nettoie(secours),
                    "niveau": clean(ligne.get("niveau_label", "")),
                    "statut_generation": "conditionnel",
                    "conditions_ok": False,
                    "avertissement": render_message(messages, "paragraphe_bloque"),
                }
            else:
                resultat[bid] = {
                    "texte": "", "niveau": "", "statut_generation": "bloqué",
                    "avertissement": render_message(messages, "paragraphe_bloque"),
                }
    return resultat

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
    detections: pd.DataFrame = None,
    metadata_rules: pd.DataFrame = None,
    questions: pd.DataFrame = None,
    textes_reponse: Dict[str, Dict] = None,
    entreprise_data: Dict[str, str] = None,
    conv_infos: Dict[str, str] = None,
    cctp_infos: Dict[str, str] = None,
    contradictions: List[str] = None,
    web_only: bool = False,
    axes_config: Dict = None,
):
    """Génère les 3 livrables, toujours dans les deux formats disponibles :

    1. Plan_Actions_Offre_BIM.pdf + .html : actions, questions et textes d'offre.
    2. Glossaire_BIM.pdf + .html : termes BIM détectés (documents, questionnaire, outil).
    3. Dashboard (Rapport_BIM.html) : généré séparément par build_dashboard(), HTML uniquement.
    """
    textes_reponse = textes_reponse or {}
    entreprise_data = entreprise_data or {}
    conv_infos = conv_infos or {}
    cctp_infos = cctp_infos or {}
    contradictions = contradictions or []

    # Les champs libres de l'Excel peuvent contenir des variables de contexte.
    # On les substitue sur une copie afin de conserver les données moteur intactes.
    ctx = {
        **_extraire_details_convention(
            "GLOBAL", "", "", "", "", conv_text, lot,
            conv_pages=conv.pages, gloss=gloss,
            detections=detections, metadata_rules=metadata_rules,
        ),
        "lot": lot,
    }
    colonnes_libres = [
        "description_convention", "lecture_entreprise", "risque_si_ignore",
        "type_preuve_cctp_attendue", "interpretation_croisee",
        "action_confirmee", "action_probable", "action_non_applicable",
        "question_bim_manager",
        "proposition_commerciale", "action_interne_standard",
        "action_interne_capacite",
    ]
    report_results: Dict[str, Dict] = {}
    for bid, res in resultats_lecture.items():
        copie = dict(res)
        bloc = dict(res.get("bloc", {}))
        for col in colonnes_libres:
            bloc[col] = _sub_ctx(str(bloc.get(col, "")), ctx)
        bloc["question_bim_manager"] = filtrer_question_bm(
            bloc.get("question_bim_manager", ""),
            res.get("conv_pf", ""), res.get("cctp_pf", ""),
            conv.text, lot,
        )
        copie["bloc"] = bloc
        report_results[bid] = copie

    cctp_detail = "Non fourni"
    if cctp:
        cctp_detail = cctp_infos.get("nom_fichier") or cctp.name
        if cctp_infos.get("indice"):
            cctp_detail += f" - indice {cctp_infos['indice']}"
        cctp_detail += f" - {len(cctp.pages)} pages"
    convention_detail = conv_infos.get("nom_fichier") or conv.name
    if conv_infos.get("indice"):
        convention_detail += f" - indice {conv_infos['indice']}"
    convention_detail += f" - {len(conv.pages)} pages"

    documents_resume = f"Convention : {conv.name}"
    if cctp:
        documents_resume += f" | CCTP : {cctp.name}"
    else:
        documents_resume += " | CCTP non fourni"

    _texte_glossaire = (
        f"{conv_text or ''} {cctp.text if cctp else ''} "
        f"{texte_outil_et_questionnaire(blocs, questions)}"
    )

    meta = {
        "projet": conv.name.replace(".pdf", ""),
        "entreprise": entreprise,
        "lot": lot,
        "marque": marque or entreprise or "Entreprise",
        "date": datetime.now().strftime("%d/%m/%Y"),
        "documents_resume": documents_resume,
        "convention_detail": convention_detail,
        "cctp_detail": cctp_detail,
        "contradictions": contradictions,
        "referent_bim": entreprise_data.get("referent_bim_nom", ""),
        "glossary_entries": glossary_entries_from_dataframe(
            filtrer_glossaire_detecte(gloss, _texte_glossaire)
        ),
        "axes_config": axes_config or {},
        **charger_logo(logo_path),
    }

    # V14 : chaque livrable qui a un PDF (Plan d'Actions, Glossaire) est
    # désormais généré dans les deux formats, y compris via le web -- pour
    # que l'interface propose systématiquement "Télécharger PDF" / "Ouvrir HTML".
    base_path = out if out.suffix.lower() == ".pdf" else out.with_suffix(".pdf")
    action_pdf = base_path.with_name("Plan_Actions_Offre_BIM.pdf")
    action_html = base_path.with_name("Plan_Actions_Offre_BIM.html")
    glossary_html_path = base_path.with_name("Glossaire_BIM.html")
    glossary_pdf_path = base_path.with_name("Glossaire_BIM.pdf")

    generate_action_plan_pdf(action_pdf, report_results, scores_eval, textes_reponse, meta)
    generate_action_plan_html(action_html, report_results, scores_eval, textes_reponse, meta)
    generate_glossary_html(glossary_html_path, meta)
    generate_glossary_pdf(glossary_pdf_path, meta)

    print(f"✓ Plan d'actions PDF généré : {action_pdf}")
    print(f"✓ Plan d'actions HTML généré : {action_html}")
    print(f"✓ Glossaire HTML généré : {glossary_html_path}")
    print(f"✓ Glossaire PDF généré : {glossary_pdf_path}")

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(
        description="Outil opérationnel BIM -- Lecture + Auto-éval + Réponse AO"
    )
    ap.add_argument("--param",        default="Parametrage_Outil_BIM_V1.xlsx")
    ap.add_argument("--convention",   required=True)
    ap.add_argument("--cctp",         default=None)
    ap.add_argument("--lot",          default=None)
    ap.add_argument("--entreprise",   default=None)
    ap.add_argument("--marque",       default=None,
                    help="Sigle affiché en en-tête des livrables (par défaut : --entreprise).")
    ap.add_argument("--logo",         default=None,
                    help="Chemin d'une image PNG/JPEG à insérer en en-tête des livrables, à la place du sigle texte.")
    ap.add_argument("--out",          default="Rapport_BIM_Complet.pdf")
    ap.add_argument("--sans_eval",    action="store_true",
                    help="Génère uniquement la partie lecture, sans questionnaire")
    ap.add_argument("--web-only", action="store_true",
                    help="Génère uniquement le dashboard HTML et le plan d'action HTML")
    ap.add_argument("--api-key",      default="",
                    dest="api_key",
                    help="Clé API Anthropic pour questions enrichies (optionnel). "
                         "Peut aussi être définie via la variable d'environnement ANTHROPIC_API_KEY.")
    ap.add_argument("--reponses-json", default="", dest="reponses_json",
                    help="Fichier JSON {id_question: 0|1|2} pour un lancement web/non interactif.")
    ap.add_argument("--format",       default="both",
                    choices=["pdf","html","both"],
                    help="Format de sortie : pdf, html (dashboard interactif), both (défaut)")
    args = ap.parse_args()

    reponses_web = None
    if args.reponses_json:
        rp = Path(args.reponses_json)
        if not rp.exists():
            raise SystemExit(f"Fichier de réponses introuvable : {rp}")
        try:
            reponses_web = json.loads(rp.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SystemExit(f"JSON de réponses invalide : {exc}")
        if not isinstance(reponses_web, dict):
            raise SystemExit("Le JSON de réponses doit être un objet {id_question: 0|1|2}.")

    param = Path(args.param)
    if not param.exists(): raise SystemExit(f"Paramétrage introuvable : {param}")
    conv_path = Path(args.convention)
    if not conv_path.exists(): raise SystemExit(f"Convention introuvable : {conv_path}")

    # ── Lecture feuille 70_Entreprise ─────────────────────────────────
    ent = load_entreprise(param)
    lots_ent = get_lots_entreprise(ent)

    # Nom entreprise : priorité CLI > feuille 70_Entreprise (clé nom_entreprise)
    if not args.entreprise:
        args.entreprise = ent.get("nom_entreprise", "").strip() or "[Nom de l'entreprise]"
    print(f"  → Entreprise : {args.entreprise}")

    # Lot : priorité CLI > feuille 70_Entreprise (clés lot_1…lot_5)
    if not args.lot:
        if lots_ent:
            args.lot = lots_ent[0]
            print(f"  → Lot lu depuis 70_Entreprise : {args.lot}")
        else:
            args.lot = "Lot"
    # Normaliser le nom du lot
    import re as _re
    lot_clean = args.lot.strip()
    # Supprimer préfixes redondants : "le lot ", "lot " (sauf "Lot 3 — Façade")
    lot_clean = _re.sub(r"^le\s+lot\s+", "", lot_clean, flags=_re.IGNORECASE).strip()
    # Garder "Lot 3 — Façade" tel quel, supprimer "lot " seul
    if not _re.match(r"^lot\s+\d", lot_clean, _re.IGNORECASE):
        lot_clean = _re.sub(r"^lot\s+", "", lot_clean, flags=_re.IGNORECASE).strip()
    # Si après nettoyage le lot est vide ou juste un article → garder l'original
    args.lot = lot_clean if len(lot_clean) > 2 else args.lot.strip()
    # Si encore vide après nettoyage → fallback silencieux
    if not args.lot or args.lot.lower() in ("le lot", "lot", "le"):
        args.lot = "Lot"
    print(f"\n→ Analyse pour : {args.entreprise} · {args.lot}")

    # Chargement
    (blocs, gloss, questions, textes, signaux, detections,
     metadata_rules, axes_config_df, coherence_rules, messages_df, params_df) = load_all(param)
    questions = normalize_questions_dataframe(questions)
    engine_params = load_params(params_df)
    messages = load_messages(messages_df)
    axes_config = load_axes_config(axes_config_df)
    conv = lire_pdf(conv_path)
    # Extraction métadonnées convention depuis pages BRUTES fitz (avant nettoyage)
    _conv_doc_raw = fitz.open(str(conv_path))
    _conv_pgs_raw = [_conv_doc_raw[i].get_text() for i in range(min(10, len(_conv_doc_raw)))]
    _conv_infos   = _extraire_infos_document(_conv_pgs_raw, detections, metadata_rules)
    _conv_doc_raw.close()
    cctp = None
    if args.cctp:
        cp = Path(args.cctp)
        if not cp.exists():
            print(f"  ⚠ CCTP introuvable : {cp} -- poursuite de l'analyse avec la convention seule.")
        else:
            cctp = lire_pdf(cp)
            # Extraction page de garde + en-têtes du CCTP (même logique que convention)
            _cctp_doc  = fitz.open(str(cp))
            _cctp_pgs  = [_cctp_doc[i].get_text() for i in range(min(10, len(_cctp_doc)))]
            _cctp_infos = _extraire_infos_document(_cctp_pgs, detections, metadata_rules)
            _cctp_infos["n_pages"] = len(_cctp_doc)
            _cctp_infos["nom_fichier"] = cp.stem
    if not cctp:
        _cctp_infos = {}
    cctp_text = cctp.text if cctp else ""

    # ── MODULE 1 : LECTURE ────────────────────────────────────────────────────
    print(f"\n[1/3] Analyse des exigences BIM -- Convention : {conv.name}")
    if not cctp:
        print("  ⚠ Aucun CCTP fourni -- analyse basée sur la convention uniquement.")
        print("  Les statuts PROBABLE/NON DÉMONTRÉE seront plus fréquents sans CCTP de lot.")
    ORD = {"CONFIRMÉE":0,"PROBABLE":1,"PARTIELLE":2,"NON DÉMONTRÉE":3,"NON APPLICABLE":4}
    resultats_lecture = {}

    # Ordre depuis l'Excel 10_Blocs uniquement — jamais codé en dur
    blocs_sorted = blocs.copy()

    # ── PASSE 1 : calculer scores et extraits bruts pour tous les blocs ──────
    # On ne décide PAS encore quelles citations afficher.
    _raw: Dict = {}
    for _, b in blocs_sorted.iterrows():
        bid   = clean(b["id_bloc"])
        kws   = split_kw(b.get("keywords_convention",""))
        strong= split_kw(b.get("keywords_preuve_forte",""))
        kws_c = (split_kw(b.get("keywords_cctp_confirme","")) +
                 split_kw(b.get("keywords_cctp_nuance","")))
        conv_pg, conv_pf, conv_sc, conv_key = best_evidence(conv, kws, strong, set())
        cctp_pg, cctp_pf, cctp_sc, cctp_key = best_evidence(
            cctp, kws+kws_c, split_kw(b.get("keywords_cctp_confirme","")), set())
        _raw[bid] = {
            "b": b,
            "conv_pg": conv_pg, "conv_pf": conv_pf, "conv_sc": conv_sc, "conv_key": conv_key,
            "cctp_pg": cctp_pg, "cctp_pf": cctp_pf, "cctp_sc": cctp_sc, "cctp_key": cctp_key,
        }

    # ── PASSE 2 : dédoublonnage des citations par EXTRAIT ───────────────────
    # Une même page peut contenir plusieurs exigences distinctes. L’ancien
    # arbitrage « un seul bloc gagnant par page » supprimait donc parfois la
    # preuve directe d’un bloc confirmé. On ne dédoublonne désormais que les
    # extraits réellement identiques (même clé normalisée).
    _cctp_excerpt_winner: Dict[str, tuple] = {}  # clé extrait → (score, bid)
    for bid, r in _raw.items():
        key, sc = r.get("cctp_key", ""), r["cctp_sc"]
        if key and sc >= 8:
            if key not in _cctp_excerpt_winner or sc > _cctp_excerpt_winner[key][0]:
                _cctp_excerpt_winner[key] = (sc, bid)

    for _, b in blocs_sorted.iterrows():
        bid = clean(b["id_bloc"])
        r   = _raw[bid]
        # Convention : score < 8 → passage trop générique
        conv_pg = r["conv_pg"] if r["conv_sc"] >= 8 else ""
        conv_pf = r["conv_pf"] if r["conv_sc"] >= 8 else ""
        # CCTP : conserver tout extrait ciblé ; ne retirer que les doublons
        # stricts déjà attribués à un bloc mieux scoré.
        cctp_pg_raw, cctp_sc_raw, cctp_key = r["cctp_pg"], r["cctp_sc"], r.get("cctp_key", "")
        if (cctp_pg_raw and cctp_key and cctp_sc_raw >= 8 and
                _cctp_excerpt_winner.get(cctp_key, (0, ""))[1] == bid):
            cctp_pg, cctp_pf = cctp_pg_raw, r["cctp_pf"]
        else:
            cctp_pg, cctp_pf = "", ""

        axes = evaluate(b, cctp_text, args.lot, conv.text,
                         signaux=signaux, params=engine_params, messages=messages)
        statut, cert = statut_legacy(axes["presence_contractuelle"], axes["applicabilite_lot"])
        concl = axes["conclusion"]
        resultats_lecture[bid] = {
            "bloc": b, "statut": statut, "certitude": cert, "conclusion": concl,
            "presence_contractuelle": axes["presence_contractuelle"],
            "applicabilite_lot": axes["applicabilite_lot"],
            "capacite_entreprise": axes["capacite_entreprise"],
            "contradictions": [],
            "conv_pg": conv_pg, "conv_pf": conv_pf,
            "cctp_pg": cctp_pg, "cctp_pf": cctp_pf,
            "conv_extrait": conv_pf[:200] if conv_pf else "",
            "cctp_extrait": cctp_pf[:200] if cctp_pf else "",
        }
        print(f"  {bid} → {axes['presence_contractuelle']} | {axes['applicabilite_lot']}")

    # Gouvernance finale : cohérence stricte entre le lot, les preuves et le statut.
    appliquer_gouvernance_preuves(resultats_lecture, conv.text, cctp_text, args.lot)

    # Trier par priorité
    resultats_lecture = dict(sorted(
        resultats_lecture.items(),
        key=lambda x: ORD.get(x[1]["statut"],5)
    ))

    # ── MODULE 2 : AUTO-ÉVALUATION ────────────────────────────────────────────
    scores_eval = {}
    n_pages_cctp = len(cctp.pages) if cctp else 0  # initialisé ici pour éviter NameError
    if not args.sans_eval:
        print(f"\n[2/3] Auto-évaluation de la maturité numérique")
        # Résoudre la clé API : argument CLI > variable d'environnement
        import os
        api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        # Injecter conv_text dans les résultats pour la génération contextuelle
        for bid in resultats_lecture:
            resultats_lecture[bid]["_conv_text"] = conv.text
        scores_eval = poser_questionnaire(
            blocs, args.lot, resultats_lecture,
            questions=questions,   # feuille 50_Questions lue depuis l'Excel
            api_key=api_key, gloss=gloss, detections=detections, metadata_rules=metadata_rules,
            reponses_fournies=reponses_web)
        # Axe 3 : capacité entreprise, déduite de la maturité du questionnaire
        for bid, sc in scores_eval.items():
            if bid in resultats_lecture:
                pct = sc.get("pct", 0)
                resultats_lecture[bid]["capacite_entreprise"] = (
                    "DEMONTREE" if pct >= 70 else ("PARTIELLE" if pct >= 40 else "NON_DEMONTREE"))
        # Détection de contradictions fiche entreprise / réponses (55_Regles_Coherence)
        contradictions_globales = detecter_contradictions(coherence_rules, ent, scores_eval, resultats_lecture)
        if contradictions_globales:
            print(f"  ⚠ {len(contradictions_globales)} contradiction(s) détectée(s) entre la fiche entreprise et le questionnaire.")
    else:
        print(f"\n[2/3] Auto-évaluation ignorée (--sans_eval)")
        contradictions_globales = []

    # ── MODULE 3 : RAPPORT ────────────────────────────────────────────────────
    print(f"\n[3/3] Génération du rapport...")
    out_path = Path(args.out)

    # Les mêmes propositions prudentes alimentent tous les livrables.
    textes_reponse = preparer_textes_reponse(
        textes, scores_eval, resultats_lecture, args.entreprise, args.lot,
        messages, engine_params,
    )

    if args.web_only:
        generate_rapport(
            out_path.with_suffix(".pdf"), conv, cctp, blocs, gloss, textes,
            resultats_lecture, scores_eval,
            args.lot, args.entreprise, conv.text,
            marque=args.marque, logo_path=args.logo,
            detections=detections, metadata_rules=metadata_rules, questions=questions,
            textes_reponse=textes_reponse, entreprise_data=ent,
            conv_infos=_conv_infos, cctp_infos=_cctp_infos,
            contradictions=contradictions_globales, web_only=True,
            axes_config=axes_config,
        )

    # PDF
    if not args.web_only and args.format in ("pdf", "both"):
        pdf_path = out_path if out_path.suffix == ".pdf" else out_path.with_suffix(".pdf")
        generate_rapport(
            pdf_path, conv, cctp, blocs, gloss, textes,
            resultats_lecture, scores_eval,
            args.lot, args.entreprise, conv.text,
            marque=args.marque, logo_path=args.logo,
            detections=detections, metadata_rules=metadata_rules, questions=questions,
            textes_reponse=textes_reponse, entreprise_data=ent,
            conv_infos=_conv_infos, cctp_infos=_cctp_infos,
            contradictions=contradictions_globales,
            axes_config=axes_config,
        )

    # HTML Dashboard
    if args.format in ("html", "both"):
        if not _HAS_DASHBOARD:
            print("  ⚠ Module dashboard_template.py introuvable -- HTML non généré.")
        else:
            html_path = out_path if out_path.suffix == ".html" else out_path.with_suffix(".html")
            # Les propositions ont déjà été construites avant la génération PDF.
            # On réutilise exactement le même dictionnaire pour le dashboard.
            indice_fiabilite, composantes_qualite = calculer_indice_fiabilite(
                resultats_lecture, textes_reponse, contradictions_globales, engine_params)
            # Glossaire enrichi : extrait du document réel si trouvé, sinon définition Excel
            # Le texte de référence pour "un terme est-il pertinent ?" couvre la
            # convention, le CCTP, le questionnaire d'auto-évaluation (50_Questions)
            # et les textes que l'outil génère lui-même (10_Blocs) -- pas seulement
            # une citation mot pour mot des PDF d'entrée.
            all_txt = (
                conv.text + (" " + cctp.text if cctp else "")
                + " " + texte_outil_et_questionnaire(blocs, questions)
            )
            cctp_pages_list = cctp.pages if cctp else []
            gloss_enrichi = enrichir_glossaire(gloss, conv.pages, cctp_pages_list)
            # Filtrer : garder seulement les termes présents dans les documents
            gloss_list = [
                g for g in gloss_enrichi
                if any(contains(all_txt, m)
                       for m in (split_kw(
                           next((row.get("mots_cles","") for _, row in gloss.iterrows()
                                 if clean(row.get("terme","")) == g["terme"]), "")
                       ) or [g["terme"]]))
            ]
            # ── Contexte global convention (détections dynamiques) ───────────
            _ctx_glob = _extraire_details_convention(
                "GLOBAL", "", "", "", "", conv.text, args.lot, conv_pages=conv.pages,
                gloss=gloss, detections=detections, metadata_rules=metadata_rules
            )

            # ── KPI thèse ────────────────────────────────────────────────────
            # Axe 2 : Bloc critique (confirmé avec maturité la plus faible)
            blocs_conf = {bid: sc for bid, sc in scores_eval.items()
                          if resultats_lecture.get(bid,{}).get("statut") == "CONFIRMÉE"}
            bloc_critique = min(blocs_conf, key=lambda b: blocs_conf[b]["pct"]) if blocs_conf else ""
            bloc_critique_pct = blocs_conf[bloc_critique]["pct"] if bloc_critique else 0
            bloc_critique_titre = (blocs_conf[bloc_critique]["titre"].split("(")[0].strip()[:35]
                                   if bloc_critique else "")

            # Axe 1 : Densité BIM CCTP (signaux / pages)
            n_pages_cctp = len(cctp.pages) if cctp else 0
            n_signaux_cctp = 0
            if cctp:
                for _, txt in cctp.pages:
                    if has_bim_signal(txt):
                        n_signaux_cctp += 1
            if not cctp:
                densite_cctp = "Non fourni"
            else:
                densite_cctp = "Élevée" if n_signaux_cctp >= 5 else ("Moyenne" if n_signaux_cctp >= 2 else "Faible")

            # Axe 3 : Score réponse BIM (maturité moy × complétude paragraphes)
            n_blocs_app = sum(1 for r in resultats_lecture.values()
                              if r["statut"] in ("CONFIRMÉE","PROBABLE","PARTIELLE"))
            n_para = len(textes_reponse)
            completude = (n_para / n_blocs_app) if n_blocs_app > 0 else 0
            mat_moy_kpi = (round(sum(s["pct"] for s in scores_eval.values()) / len(scores_eval))
                           if scores_eval else 0)
            score_reponse = round(mat_moy_kpi * completude) if scores_eval else 0

            # Axe 4 : Barrière dominante
            # Sémantique = blocs doc/données (B02, B03, B06A, B06B) faibles
            # Technologique = blocs maquette (B01, B04, B05) faibles
            # Organisationnelle = blocs processus (B07, B04) faibles
            blocs_faibles = [bid for bid, sc in scores_eval.items() if sc["pct"] < 50
                             and resultats_lecture.get(bid,{}).get("statut")
                             in ("CONFIRMÉE","PROBABLE","PARTIELLE")]
            # Catégorisation lue depuis la colonne categorie_barriere de 10_Blocs (Excel)
            # Un nouveau bloc ajouté dans l'Excel sera classé automatiquement
            _cat_blocs = {clean(b["id_bloc"]): clean(b.get("categorie_barriere",""))
                          for _, b in blocs.iterrows() if b.get("categorie_barriere","")}
            sem  = sum(1 for b in blocs_faibles if _cat_blocs.get(b,"") == "semantique")
            tech = sum(1 for b in blocs_faibles if _cat_blocs.get(b,"") == "technologique")
            org  = sum(1 for b in blocs_faibles if _cat_blocs.get(b,"") == "organisationnelle")
            if sem >= tech and sem >= org:
                barriere = "Sémantique"
                barriere_desc = "Difficulté à lire et interpréter les documents contractuels BIM"
            elif tech >= sem and tech >= org:
                barriere = "Technologique"
                barriere_desc = "Absence d'outils ou de compétences BIM pour produire les livrables"
            else:
                barriere = "Organisationnelle"
                barriere_desc = "Absence de processus et de référent BIM dans l'organisation"
            barriere_blocs = blocs_faibles[:3]

            meta = {
                "projet": conv.name.replace(".pdf",""),
                "lot": args.lot,
                "entreprise": args.entreprise,
                "marque": args.marque or args.entreprise or "Entreprise",
                "date": datetime.now().strftime("%d/%m/%Y"),
                "glossary_entries": glossary_entries_from_dataframe(filtrer_glossaire_detecte(gloss, all_txt)),
                # Modèle à 3 axes / cohérence / indice de fiabilité
                "axes_config": axes_config,
                "contradictions": contradictions_globales,
                **charger_logo(args.logo),
                "indice_fiabilite": indice_fiabilite,
                "indice_fiabilite_nom": "Indice de fiabilité de la réponse BIM",
                "composantes_qualite": composantes_qualite,
                "composantes_labels": {
                    "exactitude_contractuelle": "Exactitude contractuelle",
                    "coherence": "Cohérence entreprise / réponses",
                    "preuves": "Preuves et conditions",
                    "couverture": "Couverture utile",
                },
                "qualite_explication": ("L'indice combine exactitude contractuelle, cohérence, preuves et "
                                        "couverture. Il ne remplace pas une revue humaine avant engagement contractuel."),
                # KPI thèse
                "bloc_critique": bloc_critique,
                "bloc_critique_pct": bloc_critique_pct,
                "bloc_critique_titre": bloc_critique_titre,
                "n_signaux_cctp": n_signaux_cctp,
                "n_pages_cctp": n_pages_cctp,
                "densite_cctp": densite_cctp,
                "score_reponse": score_reponse,
                "n_para": n_para,
                "n_blocs_app": n_blocs_app,
                "barriere": barriere,
                "barriere_desc": barriere_desc,
                "barriere_blocs": barriere_blocs,
                # Détections dynamiques convention (plateforme, LOD, etc.)
                "niveau_bim": _ctx_glob.get("niveau_bim", ""),
                "dim_4d": _ctx_glob.get("dim_4d", False),
                "dim_5d": _ctx_glob.get("dim_5d", False),
                "dim_6d": _ctx_glob.get("dim_6d", False),
                "dim_7d": _ctx_glob.get("dim_7d", False),
                "bim_manager_conv": _ctx_glob.get("bim_manager", ""),
                "coordinateur_bim_conv": _ctx_glob.get("coordinateur_bim", ""),
                "geo_referencement": _ctx_glob.get("geo_referencement", False),
                "clash_3d": _ctx_glob.get("clash_3d", False),
                "doe_numerique": _ctx_glob.get("doe_numerique", False),
                "formats_livrables": _ctx_glob.get("formats_livrables", ""),
                "contractuel": _ctx_glob.get("contractuel", False),
                "plateforme_conv": _ctx_glob.get("plateforme", ""),
                "plateforme_page_conv": _ctx_glob.get("plateforme_page", ""),
                "plateforme_extrait_conv": _ctx_glob.get("plateforme_extrait", ""),
                "logiciel_page_conv": _ctx_glob.get("logiciel_page", ""),
                "logiciel_extrait_conv": _ctx_glob.get("logiciel_extrait", ""),
                "format_ifc_page_conv": _ctx_glob.get("format_ifc_page", ""),
                "format_ifc_extrait_conv": _ctx_glob.get("format_ifc_extrait", ""),
                "lod_page_conv": _ctx_glob.get("lod_page", ""),
                "lod_extrait_conv": _ctx_glob.get("lod_extrait", ""),
                "nd_page_conv": _ctx_glob.get("nd_page", ""),
                "nd_extrait_conv": _ctx_glob.get("nd_extrait", ""),
                "niveau_bim_page_conv": _ctx_glob.get("niveau_bim_page", ""),
                "niveau_bim_extrait_conv": _ctx_glob.get("niveau_bim_extrait", ""),
                "bim_manager_page_conv": _ctx_glob.get("bim_manager_page", ""),
                "bim_manager_extrait_conv": _ctx_glob.get("bim_manager_extrait", ""),
                "coordinateur_bim_page_conv": _ctx_glob.get("coordinateur_bim_page", ""),
                "coordinateur_bim_extrait_conv": _ctx_glob.get("coordinateur_bim_extrait", ""),
                "cout_plateforme_conv": _ctx_glob.get("cout_plateforme", ""),
                "logiciel_conv": _ctx_glob.get("logiciel", ""),
                "format_ifc_conv": _ctx_glob.get("format_ifc", ""),
                "lod_conv": _ctx_glob.get("lod", ""),
                "nd_conv": _ctx_glob.get("nd", ""),
                "frequence_conv": _ctx_glob.get("frequence", ""),
                # Métadonnées convention depuis pages BRUTES fitz (indice, date, émetteur)
                "indice_doc":  _conv_infos.get("indice", ""),
                "date_doc":    _conv_infos.get("date", ""),
                "emetteur_doc": _conv_infos.get("emetteur", ""),
                "phase_doc":   _conv_infos.get("phase", ""),
                "maitrise_oeuvre_doc": _conv_infos.get("maitrise_oeuvre", ""),
                "n_pages_convention": len(conv.pages) if conv else 0,
                # Infos entreprise
                # Infos entreprise — toutes les clés de 70_Entreprise
                "ent_adresse":    ent.get("adresse","").strip(),
                "ent_cp":         ent.get("code_postal","").strip(),
                "ent_ville":      ent.get("ville","").strip(),
                "ent_tel":        ent.get("telephone","").strip(),
                "ent_email":      ent.get("email_contact","").strip(),
                "ent_dirigeant":  ent.get("dirigeant","").strip(),
                "ent_siret":      ent.get("siret","").strip(),
                "ent_referent":   ent.get("referent_bim_nom","").strip(),
                "ent_ref_email":  ent.get("referent_bim_email","").strip(),
                "ent_ref_tel":    ent.get("referent_bim_tel","").strip(),
                "ent_logiciels":  ent.get("logiciels_bim","").strip(),
                "ent_experience": ent.get("experience_bim","").strip(),
                "ent_plateforme": ent.get("plateforme_interne","").strip(),
                "lots_entreprise": lots_ent,
                # Métadonnées CCTP — extraites depuis page de garde + en-têtes
                "n_pgs":       _cctp_infos.get("n_pages", n_pages_cctp),
                "cctp_nom":    _cctp_infos.get("nom_fichier", ""),
                "cctp_indice": _cctp_infos.get("indice", ""),
                "cctp_date":   _cctp_infos.get("date", ""),
                "cctp_auteur": _cctp_infos.get("emetteur", ""),
                "cctp_phase":  _cctp_infos.get("phase", ""),
                "cctp_mo":     _cctp_infos.get("maitre_ouvrage", ""),
                # Métadonnées convention — extraites depuis page de garde
                "maitre_ouvrage_doc": _conv_infos.get("maitre_ouvrage", ""),
            }
            # Encoder les PDF source en base64 pour la visionneuse intégrée
            conv_b64 = base64.b64encode(conv_path.read_bytes()).decode("ascii")
            cctp_b64 = ""
            if cctp:
                cctp_b64 = base64.b64encode(Path(args.cctp).read_bytes()).decode("ascii")

            # Substituer les {variables} détectées ({plateforme}, {cout}, {lot}...)
            # dans TOUTES les colonnes de texte libre de 10_Blocs -- pas seulement
            # question_bim_manager -- puis filtrer les questions BIM Manager déjà
            # couvertes par les documents. Logique générique, rien dans le template :
            # ajouter une {variable} dans l'Excel suffit, aucun code à modifier.
            _ctx_glob_sub = {**_ctx_glob, "lot": args.lot}
            _COLONNES_LIBRES = [
                "description_convention", "lecture_entreprise", "risque_si_ignore",
                "type_preuve_cctp_attendue", "interpretation_croisee",
                "action_confirmee", "action_probable", "action_non_applicable",
                "question_bim_manager",
                "proposition_commerciale", "action_interne_standard",
                "action_interne_capacite",
            ]
            for bid_f, res_f in resultats_lecture.items():
                res_f["bloc"] = dict(res_f["bloc"])
                for col in _COLONNES_LIBRES:
                    res_f["bloc"][col] = _sub_ctx(str(res_f["bloc"].get(col, "")), _ctx_glob_sub)
                q_filt = filtrer_question_bm(
                    res_f["bloc"]["question_bim_manager"],
                    res_f.get("conv_pf",""), res_f.get("cctp_pf",""),
                    conv.text, args.lot
                )
                res_f["bloc"]["question_bim_manager"] = q_filt

            # Nom du fichier PDF associé (même dossier, même base de nom)
            pdf_filename = "" if args.web_only else out_path.with_suffix(".pdf").name
            html_str = build_dashboard(
                resultats_lecture, scores_eval, textes_reponse, gloss_list, meta,
                conv_text=conv.text,
                conv_b64=conv_b64, cctp_b64=cctp_b64,
                pdf_filename=pdf_filename,
                glossaire_complet=gloss_enrichi,
            )
            html_path.write_text(html_str, encoding="utf-8")
            print(f"✓ Dashboard HTML généré : {html_path}")
            # Écrire l'analyse dans l'historique 70_Entreprise
            ecrire_historique(param, conv.name, args.lot, meta.get("score_reponse", 0))

if __name__ == "__main__":
    main()
