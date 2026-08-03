from __future__ import annotations

import html
import io
import json
import re
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from content_rules import glossary_lookup

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# Profils graphiques réels, appliqués au même moteur de génération.
STYLE_OPTION = os.environ.get("BIM_STYLE_OPTION", "1")
STYLE_PROFILES = {
    "1": {"name":"Bleu institutionnel", "navy":"#082A3A", "navy2":"#154A63", "slate":"#41677A", "text":"#163444", "muted":"#668496", "line":"#C9D9E2", "soft":"#EEF5F8", "soft2":"#F8FBFC", "red":"#9A0038", "redbg":"#FBEFF4", "amber":"#A66A12", "amberbg":"#FFF7E8", "green":"#176B55", "greenbg":"#EDF7F1", "header_h":24, "footer_h":13, "header_bg":"dark", "table_bg":"navy"},
    "2": {"name":"Bleu premium", "navy":"#061F33", "navy2":"#123F62", "slate":"#365E7A", "text":"#132E43", "muted":"#607D91", "line":"#C7D7E3", "soft":"#ECF3F8", "soft2":"#F8FBFD", "red":"#A4003C", "redbg":"#FBEAF1", "amber":"#A56A13", "amberbg":"#FFF5E3", "green":"#176B55", "greenbg":"#EDF7F1", "header_h":27, "footer_h":14, "header_bg":"dark", "table_bg":"navy"},
    "3": {"name":"Bleu technique clair", "navy":"#174C66", "navy2":"#39758F", "slate":"#587F91", "text":"#234351", "muted":"#708D9B", "line":"#CBDDE5", "soft":"#F1F7FA", "soft2":"#FFFFFF", "red":"#8F1740", "redbg":"#FAF0F4", "amber":"#9B6B1C", "amberbg":"#FFF8E9", "green":"#2B6D5B", "greenbg":"#EFF7F3", "header_h":17, "footer_h":10, "header_bg":"light", "table_bg":"navy"},
    "4": {"name":"Bleu exécutif sombre", "navy":"#031B2B", "navy2":"#083C5A", "slate":"#2D5872", "text":"#102D40", "muted":"#58758A", "line":"#BED2DE", "soft":"#EAF2F6", "soft2":"#F5F9FB", "red":"#B40042", "redbg":"#FCEAF1", "amber":"#B16D08", "amberbg":"#FFF2DF", "green":"#11705A", "greenbg":"#EAF7F2", "header_h":32, "footer_h":16, "header_bg":"dark", "table_bg":"navy"},
    "5": {"name":"Bleu ardoise élégant", "navy":"#19394C", "navy2":"#355E73", "slate":"#527287", "text":"#203C4C", "muted":"#6E8796", "line":"#CDD9E0", "soft":"#EFF4F6", "soft2":"#FAFCFD", "red":"#8F1D46", "redbg":"#F7EEF1", "amber":"#996B2A", "amberbg":"#FAF3E8", "green":"#426A5A", "greenbg":"#EEF4F1", "header_h":23, "footer_h":12, "header_bg":"graphite", "table_bg":"graphite"},
}
PROFILE = STYLE_PROFILES.get(STYLE_OPTION, STYLE_PROFILES["1"])
NAVY = colors.HexColor(PROFILE["navy"])
NAVY_2 = colors.HexColor(PROFILE["navy2"])
SLATE = colors.HexColor(PROFILE["slate"])
TEXT = colors.HexColor(PROFILE["text"])
MUTED = colors.HexColor(PROFILE["muted"])
LINE = colors.HexColor(PROFILE["line"])
SOFT = colors.HexColor(PROFILE["soft"])
SOFT_2 = colors.HexColor(PROFILE["soft2"])
RED = colors.HexColor(PROFILE["red"])
RED_BG = colors.HexColor(PROFILE["redbg"])
AMBER = colors.HexColor(PROFILE["amber"])
AMBER_BG = colors.HexColor(PROFILE["amberbg"])
GREEN = colors.HexColor(PROFILE["green"])
GREEN_BG = colors.HexColor(PROFILE["greenbg"])
WHITE = colors.white

# Filet de sécurité uniquement (voir bim_model._FALLBACK_BLOCK_TITLES) : le
# titre affiché vient en priorité de la colonne titre_bloc de 10_Blocs.
_FALLBACK_PLAIN_TITLES = {
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
    "B10": "Planning lié à la maquette",
    "B11": "Qualité des objets et modèles numériques",
    "B12": "Quantités et données de coût",
}

TECH_REPLACEMENTS = [
    (r"\blivrables\s+outil\s+de\s+gestion\s+de\s+maintenance\b", "livrables destinés à l'outil de gestion de maintenance"),
    (r"\bBIM Manager\b(?!\))", "responsable BIM du projet (BIM Manager)"),
    (r"\bdes\s+données\s+pour\s+la\s+GMAO/AIM\b", "des données pour l'exploitation et la maintenance"),
    (r"\bla\s+GMAO\b", "l'outil de gestion de maintenance (GMAO)"),
    (r"\bl['’]AIM\b", "le modèle d'information pour l'exploitation (AIM)"),
    (r"\bGMAO/AIM\b", "exploitation et maintenance"),
    (r"\bCDE\b", "plateforme commune de dépôt (CDE)"),
    (r"\bLOIN\b", "liste des informations à fournir (LOIN)"),
    (r"\bBIM\s*4D\b", "lien entre la maquette et le planning (BIM 4D)"),
    (r"\bBIM\s*5D\b", "lien entre la maquette, les quantités et les coûts (BIM 5D)"),
    (r"\bAIM\b(?!\))", "modèle d'information pour l'exploitation (AIM)"),
    (r"\bGMAO\b(?!\))", "outil de gestion de maintenance (GMAO)"),
]


def _clean(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text


def _pdf_safe(value: Any) -> str:
    text = _clean(value)
    replacements = {
        "—": "-", "–": "-", "→": "->", "⚡": "", "✓": "OK",
        "●": "", "◐": "", "◌": "", "○": "", "“": '"', "”": '"',
        "’": "'", "…": "...", "×": "x", " ": " ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def _plain_language(value: Any) -> str:
    text = _pdf_safe(value)
    for pattern, replacement in TECH_REPLACEMENTS:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _esc(value: Any) -> str:
    return html.escape(_pdf_safe(value), quote=False)


def _para(value: Any, style: ParagraphStyle) -> Paragraph:
    return Paragraph(_esc(value).replace("\n", "<br/>"), style)


def _rich_para(value: Any, style: ParagraphStyle) -> Paragraph:
    """Accept a tiny, controlled subset already generated by this module."""
    return Paragraph(str(value), style)


def _register_fonts() -> Tuple[str, str, str, str]:
    """Enregistre Calibri lorsqu'elle est disponible sur Windows.

    Sur Linux/Railway, Carlito est utilisé comme substitut métrique libre.
    Aucun fichier de police n'est embarqué dans le projet.
    """
    candidates = [
        (
            Path("C:/Windows/Fonts/calibri.ttf"),
            Path("C:/Windows/Fonts/calibrib.ttf"),
            Path("C:/Windows/Fonts/calibrii.ttf"),
            Path("C:/Windows/Fonts/calibril.ttf"),
        ),
        (
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"),
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"),
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Italic.ttf"),
            Path("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"),
        ),
        (
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ),
        (
            Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
            Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
            Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf"),
            Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
        ),
    ]
    for regular, bold, italic, light in candidates:
        if regular.exists() and bold.exists() and italic.exists() and light.exists():
            try:
                pdfmetrics.registerFont(TTFont("BIMSans", str(regular)))
                pdfmetrics.registerFont(TTFont("BIMSans-Bold", str(bold)))
                pdfmetrics.registerFont(TTFont("BIMSans-Italic", str(italic)))
                pdfmetrics.registerFont(TTFont("BIMSans-Light", str(light)))
                pdfmetrics.registerFontFamily(
                    "BIMSans", normal="BIMSans", bold="BIMSans-Bold",
                    italic="BIMSans-Italic", boldItalic="BIMSans-Bold",
                )
                return "BIMSans", "BIMSans-Bold", "BIMSans-Italic", "BIMSans-Light"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica"


FONT, FONT_BOLD, FONT_ITALIC, FONT_LIGHT = _register_fonts()


def _styles() -> Dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "cover_kicker": ParagraphStyle(
            "cover_kicker", parent=base["Normal"], fontName=FONT_BOLD,
            fontSize=8, leading=10, textColor=MUTED, spaceAfter=8,
            uppercase=True, letterSpacing=0.8,
        ),
        "cover_title": ParagraphStyle(
            "cover_title", parent=base["Title"], fontName=FONT_LIGHT,
            fontSize=24, leading=28, textColor=NAVY, spaceAfter=12,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub", parent=base["Normal"], fontName=FONT,
            fontSize=11, leading=16, textColor=SLATE,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontName=FONT_LIGHT,
            fontSize=16, leading=20, textColor=NAVY, spaceBefore=4, spaceAfter=10,
        ),
        "h2": ParagraphStyle(
            "h2", parent=base["Heading2"], fontName=FONT_LIGHT,
            fontSize=12.5, leading=16, textColor=NAVY, spaceBefore=5, spaceAfter=6,
        ),
        "h3": ParagraphStyle(
            "h3", parent=base["Heading3"], fontName=FONT_BOLD,
            fontSize=10.5, leading=13, textColor=TEXT, spaceBefore=3, spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body", parent=base["BodyText"], fontName=FONT,
            fontSize=9.2, leading=13.2, textColor=TEXT, spaceAfter=5,
        ),
        "body_small": ParagraphStyle(
            "body_small", parent=base["BodyText"], fontName=FONT,
            fontSize=8.2, leading=11.5, textColor=SLATE, spaceAfter=3,
        ),
        "label": ParagraphStyle(
            "label", parent=base["Normal"], fontName=FONT_BOLD,
            fontSize=7, leading=9, textColor=MUTED, uppercase=True,
            letterSpacing=0.45,
        ),
        "metric": ParagraphStyle(
            "metric", parent=base["Normal"], fontName=FONT_LIGHT,
            fontSize=20, leading=22, textColor=NAVY, alignment=TA_CENTER,
        ),
        "metric_label": ParagraphStyle(
            "metric_label", parent=base["Normal"], fontName=FONT_BOLD,
            fontSize=7.5, leading=9.5, textColor=MUTED, alignment=TA_CENTER,
        ),
        "quote": ParagraphStyle(
            "quote", parent=base["BodyText"], fontName=FONT_ITALIC,
            fontSize=8.7, leading=12.5, textColor=SLATE, leftIndent=6,
            borderColor=LINE, borderWidth=0, borderPadding=0,
        ),
        "table": ParagraphStyle(
            "table", parent=base["BodyText"], fontName=FONT,
            fontSize=7.8, leading=10.5, textColor=TEXT,
        ),
        "table_bold": ParagraphStyle(
            "table_bold", parent=base["BodyText"], fontName=FONT_BOLD,
            fontSize=7.8, leading=10.5, textColor=TEXT,
        ),
        "tiny": ParagraphStyle(
            "tiny", parent=base["Normal"], fontName=FONT,
            fontSize=6.8, leading=9, textColor=MUTED,
        ),
        "footer": ParagraphStyle(
            "footer", parent=base["Normal"], fontName=FONT,
            fontSize=6.8, leading=8, textColor=MUTED,
        ),
        "right": ParagraphStyle(
            "right", parent=base["Normal"], fontName=FONT,
            fontSize=7.2, leading=9, textColor=MUTED, alignment=TA_RIGHT,
        ),
        "center": ParagraphStyle(
            "center", parent=base["Normal"], fontName=FONT,
            fontSize=8, leading=11, textColor=SLATE, alignment=TA_CENTER,
        ),
    }


ST = _styles()


def _status_key(result: Dict[str, Any]) -> str:
    return _clean(result.get("statut", "")).upper()


def _status_info(result: Dict[str, Any]) -> Tuple[str, colors.Color, colors.Color]:
    status = _status_key(result)
    applicability = _clean(result.get("applicabilite_lot", "")).upper()
    if status == "CONFIRMÉE" and applicability == "APPLICABLE":
        return "Confirmée pour le lot", NAVY, SOFT
    if status in {"PROBABLE", "PARTIELLE"} or applicability in {"PROBABLE", "A_CONFIRMER"}:
        return "À confirmer pour le lot", AMBER, AMBER_BG
    if status == "NON DÉMONTRÉE":
        return "Non démontrée dans les documents", SLATE, SOFT
    if status == "NON APPLICABLE":
        return "Non applicable / explicitement exclue", MUTED, SOFT
    return status.title() or "A examiner", SLATE, SOFT


def _capacity_info(block_id: str, scores: Dict[str, Dict[str, Any]]) -> Tuple[str, Optional[int], colors.Color, colors.Color]:
    score = scores.get(block_id) or {}
    pct = score.get("pct")
    if not isinstance(pct, int):
        return "Non évaluée", None, MUTED, SOFT
    if pct >= 70:
        return "Démontrée", pct, GREEN, GREEN_BG
    if pct >= 40:
        return "Partielle", pct, AMBER, AMBER_BG
    return "Non démontrée", pct, RED, RED_BG


def _priority(result: Dict[str, Any], block_id: str, scores: Dict[str, Dict[str, Any]]) -> Tuple[int, str]:
    status = _status_key(result)
    pct = (scores.get(block_id) or {}).get("pct")
    if status == "CONFIRMÉE" and isinstance(pct, int) and pct < 40:
        return 0, "Critique - sécuriser avant engagement"
    if status == "CONFIRMÉE":
        return 1, "Immédiate"
    if status in {"PROBABLE", "PARTIELLE"}:
        return 2, "Avant dépôt"
    if status == "NON DÉMONTRÉE":
        return 3, "À surveiller"
    return 9, "Sans action"


def _plain_title(block_id: str, result: Dict[str, Any]) -> str:
    excel_title = _clean((result.get("bloc") or {}).get("titre_bloc"))
    if excel_title:
        return _plain_language(excel_title)
    return _FALLBACK_PLAIN_TITLES.get(block_id, block_id)


def _original_title(result: Dict[str, Any]) -> str:
    return _pdf_safe((result.get("bloc") or {}).get("titre_bloc", ""))


def _relevant_items(results: Dict[str, Dict[str, Any]], scores: Dict[str, Dict[str, Any]]) -> List[Tuple[str, Dict[str, Any]]]:
    items = [(bid, res) for bid, res in results.items() if _status_key(res) != "NON APPLICABLE"]
    return sorted(items, key=lambda item: (_priority(item[1], item[0], scores)[0], (scores.get(item[0]) or {}).get("pct", 101), item[0]))


def _actions_for(block_id: str, result: Dict[str, Any], scores: Dict[str, Dict[str, Any]]) -> List[str]:
    block = result.get("bloc") or {}
    status = _status_key(result)
    actions: List[str] = []
    action_text = _clean(block.get("action_confirmee" if status == "CONFIRMÉE" else "action_probable", ""))
    if action_text:
        actions.append(_plain_language(action_text))
    question = _clean(block.get("question_bim_manager", ""))
    if question:
        actions.append("Obtenir une réponse écrite du BIM Manager sur le point suivant : " + _plain_language(question))
    pct = (scores.get(block_id) or {}).get("pct")
    if isinstance(pct, int):
        if pct < 40:
            actions.append("Ne pas prendre un engagement ferme tant que l'écart de capacité n'est pas traité.")
        elif pct < 70:
            actions.append("Désigner un responsable et sécuriser le point partiellement maîtrisé avant de confirmer l'engagement.")
        else:
            actions.append("Conserver une preuve de la capacité annoncée pour la mise au point du marché.")
    if not actions:
        actions.append("Faire valider cette exigence par le référent BIM avant de finaliser l'offre.")
    # Dedupe while keeping order.
    seen = set()
    unique = []
    for action in actions:
        key = action.lower()
        if key not in seen:
            seen.add(key)
            unique.append(action)
    return unique


def _proposal_for(block_id: str, proposals: Dict[str, Dict[str, Any]]) -> Tuple[str, str]:
    proposal = proposals.get(block_id) or {}
    text = _clean(proposal.get("texte", ""))
    if text:
        label = "Formulation prudente - à confirmer" if proposal.get("conditions_ok") is False else "Proposition à intégrer dans l'offre"
        return label, _plain_language(text)
    warning = _clean(proposal.get("avertissement", ""))
    return "Engagement à ne pas formuler en l'état", _plain_language(warning)


def _situation_summary(block_id: str, scores: Dict[str, Dict[str, Any]]) -> Tuple[str, str]:
    score = scores.get(block_id) or {}
    responses = score.get("reponses") or []
    if not responses:
        return "Capacité non évaluée.", "Réaliser l'auto-évaluation avant tout engagement ferme."
    strengths = []
    gaps = []
    for response in responses:
        value = response.get("reponse_val")
        label = _clean(response.get("reponse_label", ""))
        question = _clean(response.get("question", ""))
        item = _plain_language(label or question)
        if value == 2 and item:
            strengths.append(item)
        elif value in (0, 1) and item:
            gaps.append(item)
    strength = strengths[0] if strengths else "Aucun point fort suffisamment démontré dans les réponses fournies."
    gap = gaps[0] if gaps else "Aucun écart majeur identifié dans les réponses fournies."
    return strength, gap


def _readiness(results: Dict[str, Dict[str, Any]], scores: Dict[str, Dict[str, Any]], contradictions: Sequence[str] | None = None) -> Tuple[str, str, colors.Color, colors.Color]:
    relevant = _relevant_items(results, scores)
    if not scores:
        return "À ÉVALUER AVANT DÉPÔT", "Les exigences ont été analysées, mais la capacité de l'entreprise n'a pas été évaluée.", AMBER, AMBER_BG
    confirmed_low = [bid for bid, res in relevant if _status_key(res) == "CONFIRMÉE" and (scores.get(bid) or {}).get("pct", 100) < 40]
    if confirmed_low or contradictions:
        return "RISQUE IMPORTANT", "Une obligation confirmée n'est pas suffisamment maîtrisée ou une incohérence doit être levée avant engagement.", RED, RED_BG
    at_risk = [bid for bid, _ in relevant if (scores.get(bid) or {}).get("pct", 100) < 70]
    if at_risk:
        return "À SÉCURISER AVANT LE DÉPÔT", "Plusieurs points sont partiellement maîtrisés ou doivent être clarifiés avant de finaliser l'offre.", AMBER, AMBER_BG
    return "PRÊT À RÉPONDRE", "Les capacités déclarées couvrent les exigences identifiées. Conserver les preuves et confirmer les derniers points contractuels.", GREEN, GREEN_BG


def _doc_factory(path: Path, title: str):
    header_h = float(PROFILE["header_h"]) * mm
    footer_h = float(PROFILE["footer_h"]) * mm
    top_margin = header_h + (7 * mm if PROFILE["header_bg"] != "light" else 10 * mm)
    bottom_margin = footer_h + 7 * mm
    doc = BaseDocTemplate(
        str(path), pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=top_margin, bottomMargin=bottom_margin,
        title=title, author="BNN - Outil d'analyse BIM",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")

    def on_page(canvas, document):
        canvas.saveState()
        page = canvas.getPageNumber()
        mode = PROFILE["header_bg"]
        header_color = RED if mode == "red" else (SLATE if mode == "graphite" else (WHITE if mode == "light" else NAVY))
        header_text = NAVY if mode == "light" else WHITE
        sub_text = SLATE if mode == "light" else colors.HexColor("#DCE5E8")
        canvas.setFillColor(header_color)
        canvas.rect(0, A4[1] - header_h, A4[0], header_h, stroke=0, fill=1)
        canvas.setFillColor(RED)
        accent_y = A4[1] - header_h - (0.7 * mm if mode != "light" else 0.3 * mm)
        canvas.rect(0, accent_y, A4[0], 0.7 * mm if mode != "light" else 0.35 * mm, stroke=0, fill=1)
        canvas.setFillColor(header_text)
        canvas.setFont(FONT_BOLD, 16 if header_h >= 22*mm else 13)
        canvas.drawString(document.leftMargin, A4[1] - header_h/2 - 2*mm, "BNN")
        canvas.setFont(FONT_BOLD, 8.2)
        canvas.drawString(document.leftMargin + 25 * mm, A4[1] - header_h/2, _pdf_safe(title)[:78])
        canvas.setFont(FONT, 6.5)
        canvas.setFillColor(sub_text)
        canvas.drawString(document.leftMargin + 25 * mm, A4[1] - header_h/2 - 4.5*mm, f"{PROFILE['name']} - Lecture contractuelle - auto-évaluation - aide à la réponse")
        footer_color = NAVY if mode != "light" else WHITE
        footer_text = WHITE if mode != "light" else SLATE
        canvas.setFillColor(footer_color)
        canvas.rect(0, 0, A4[0], footer_h, stroke=0, fill=1)
        canvas.setFillColor(RED)
        canvas.rect(0, footer_h, A4[0], 0.5 * mm, stroke=0, fill=1)
        canvas.setFont(FONT_BOLD, 7)
        canvas.setFillColor(footer_text)
        canvas.drawString(document.leftMargin, max(3.5*mm, footer_h/2-1*mm), "BNN")
        canvas.setFont(FONT, 6.1)
        canvas.drawString(document.leftMargin + 13 * mm, max(3.5*mm, footer_h/2-1*mm), "Analyse automatisée - vérification humaine recommandée avant engagement contractuel")
        canvas.drawRightString(A4[0] - document.rightMargin, max(3.5*mm, footer_h/2-1*mm), f"Page {page}")
        canvas.restoreState()

    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=on_page)])
    return doc


def _section_title(text: str) -> List[Any]:
    return [Spacer(1, 2), _para(text, ST["h1"]), HRFlowable(width="100%", thickness=0.8, color=NAVY, spaceAfter=8)]


def _badge(text: str, fg: colors.Color, bg: colors.Color, width: float = 43 * mm) -> Table:
    tbl = Table([[_para(text, ParagraphStyle("badge", parent=ST["table_bold"], textColor=fg, alignment=TA_CENTER))]], colWidths=[width])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("BOX", (0, 0), (-1, -1), 0.6, fg),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return tbl


def _metric_card(value: str, label: str, accent: colors.Color) -> Table:
    tbl = Table([
        [_para(value, ParagraphStyle("m", parent=ST["metric"], textColor=accent))],
        [_para(label, ST["metric_label"])],
    ], colWidths=[52 * mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("LINEABOVE", (0, 0), (-1, 0), 2.2, accent),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
        ("TOPPADDING", (0, 1), (-1, 1), 0),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
    ]))
    return tbl


def _bullet_list(items: Sequence[str], style: ParagraphStyle | None = None) -> List[Any]:
    result: List[Any] = []
    target = style or ST["body"]
    for item in items:
        result.append(_rich_para(f"<b>-</b>&nbsp;&nbsp;{_esc(item)}", target))
    return result


def _cover_story(meta: Dict[str, Any], title: str, subtitle: str, readiness: Tuple[str, str, colors.Color, colors.Color], counts: Tuple[int, int, int]) -> List[Any]:
    ready_label, ready_desc, ready_fg, ready_bg = readiness
    confirmed, clarify, actions = counts
    story: List[Any] = [
        Spacer(1, 2 * mm),
        _para("NOTE D'INGÉNIERIE - ANALYSE D'APPEL D'OFFRES", ParagraphStyle("cover_kicker_bnn", parent=ST["cover_kicker"], textColor=RED)),
        _para(title, ST["cover_title"]),
        _para(subtitle, ST["cover_sub"]),
        Spacer(1, 10 * mm),
    ]
    info = Table([
        [_para("PROJET", ST["label"]), _para(meta.get("projet", "-"), ST["body"])],
        [_para("ENTREPRISE", ST["label"]), _para(meta.get("entreprise", "-"), ST["body"])],
        [_para("LOT", ST["label"]), _para(meta.get("lot", "-"), ST["body"])],
        [_para("DOCUMENTS", ST["label"]), _para(meta.get("documents_resume", "-"), ST["body_small"])],
        [_para("ÉDITION", ST["label"]), _para(meta.get("date", datetime.now().strftime("%d/%m/%Y")), ST["body"])],
    ], colWidths=[34 * mm, 128 * mm])
    info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), SOFT),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.extend([info, Spacer(1, 12 * mm)])
    ready_tbl = Table([
        [_para("NIVEAU DE PRÉPARATION", ST["label"])],
        [_para(ready_label, ParagraphStyle("ready", parent=ST["h1"], textColor=ready_fg, alignment=TA_CENTER, fontSize=18, leading=22))],
        [_para(ready_desc, ParagraphStyle("ready_desc", parent=ST["body_small"], alignment=TA_CENTER, textColor=SLATE))],
    ], colWidths=[162 * mm])
    ready_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ready_bg),
        ("BOX", (0, 0), (-1, -1), 0.9, ready_fg),
        ("TOPPADDING", (0, 0), (-1, 0), 7),
        ("TOPPADDING", (0, 1), (-1, 1), 5),
        ("BOTTOMPADDING", (0, 2), (-1, 2), 9),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.extend([ready_tbl, Spacer(1, 10 * mm)])
    metrics = Table([[
        _metric_card(str(confirmed), "OBLIGATIONS CONFIRMÉES", RED),
        _metric_card(str(clarify), "POINTS À CLARIFIER", AMBER),
        _metric_card(str(actions), "ACTIONS PRIORITAIRES", NAVY_2),
    ]], colWidths=[54 * mm] * 3)
    metrics.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    story.extend([metrics, Spacer(1, 14 * mm), _para("Document de synthèse destiné à la direction, au chargé d'affaires et au rédacteur du mémoire technique.", ST["center"]), PageBreak()])
    return story


def _summary_sentence(confirmed: int, clarify: int, no_eval: bool) -> str:
    if confirmed and clarify:
        text = f"Le lot est concerne par {confirmed} obligation(s) clairement identifiee(s) et {clarify} point(s) dont le perimetre doit encore etre confirme."
    elif confirmed:
        text = f"Le lot est concerne par {confirmed} obligation(s) clairement identifiee(s) dans les documents analyses."
    elif clarify:
        text = f"Aucune obligation directe n'est confirmee a ce stade, mais {clarify} point(s) doivent etre clarifies avant de prendre un engagement."
    else:
        text = "Aucune exigence operationnelle significative n'a ete identifiee pour ce lot dans les documents analyses."
    if no_eval:
        text += " La capacite de l'entreprise n'a pas encore ete evaluee."
    return text


def generate_action_plan_pdf(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    relevant = _relevant_items(results, scores)
    doc = _doc_factory(out, "Plan d'actions avant le dépôt de l'offre")
    story: List[Any] = [
        _para("PLAN D'ACTIONS AVANT LE DÉPÔT DE L'OFFRE", ST["cover_title"]),
        _para(f"{meta.get('entreprise', '-')} - Lot {meta.get('lot', '-')}", ST["cover_sub"]),
        Spacer(1, 5),
        _para("Ce document ne contient que les actions, les questions à transmettre et les formulations proposées pour l'offre.", ST["body_small"]),
        HRFlowable(width="100%", thickness=1, color=NAVY, spaceBefore=6, spaceAfter=10),
    ]

    action_rows: List[List[Any]] = [[
        _para("FAIT", ST["label"]), _para("PRIORITÉ", ST["label"]),
        _para("ACTION", ST["label"]), _para("BLOC", ST["label"]),
        _para("RESPONSABLE", ST["label"]), _para("ÉCHÉANCE", ST["label"]),
    ]]
    for bid, result in relevant:
        _, priority_label = _priority(result, bid, scores)
        for action in _actions_for(bid, result, scores):
            action_rows.append([
                _para("[ ]", ST["table_bold"]), _para(priority_label, ST["table_bold"]),
                _para(action, ST["table"]), _para(bid, ST["table_bold"]),
                _para("", ST["table"]), _para("Avant dépôt", ST["table"]),
            ])
    action_table = Table(action_rows, colWidths=[10 * mm, 24 * mm, 72 * mm, 13 * mm, 24 * mm, 20 * mm], repeatRows=1)
    action_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([action_table, PageBreak()])
    story += _section_title("Questions à transmettre et propositions d'offre")

    for bid, result in relevant:
        block = result.get("bloc") or {}
        question = _plain_language(block.get("question_bim_manager", ""))
        label, proposal = _proposal_for(bid, proposals)
        status_label, status_fg, status_bg = _status_info(result)
        cap_label, pct, cap_fg, cap_bg = _capacity_info(bid, scores)
        header = Table([[
            _rich_para(f"<b>{_esc(bid)} - {_esc(_plain_title(bid, result))}</b>", ST["h2"]),
            _badge(status_label, status_fg, status_bg, 47 * mm),
        ]], colWidths=[114 * mm, 48 * mm])
        header.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), SOFT_2),
            ("BOX", (0, 0), (-1, -1), 0.6, LINE),
            ("LINEBEFORE", (0, 0), (0, 0), 2.2, status_fg),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        content: List[Any] = [header, Spacer(1, 5), _rich_para(f"<b>Preparation :</b> {_esc(cap_label)}{f' - {pct}%' if pct is not None else ''}", ST["body_small"])]
        if question:
            content.extend([_para("Question à transmettre", ST["h3"]), _para(question, ST["body_small"])])
        if proposal:
            content.extend([_para(label, ST["h3"]), _para(proposal, ST["quote"])])
        content.append(Spacer(1, 8))
        story.append(KeepTogether(content))
    doc.build(story)


def generate_action_plan_html(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    relevant = _relevant_items(results, scores)
    storage_key = re.sub(r"[^a-z0-9]+", "-", f"bim-actions-{meta.get('projet','')}-{meta.get('entreprise','')}-{meta.get('lot','')}".lower()).strip("-")
    cards = []
    all_proposals = []
    action_index = 0
    for bid, result in relevant:
        status_label, _, _ = _status_info(result)
        cap_label, pct, _, _ = _capacity_info(bid, scores)
        actions_html = []
        for action in _actions_for(bid, result, scores):
            action_index += 1
            aid = f"a{action_index}"
            actions_html.append(
                f'<label class="action"><input type="checkbox" data-action-id="{aid}"><span>{html.escape(action)}</span></label>'
            )
        question = _plain_language((result.get("bloc") or {}).get("question_bim_manager", ""))
        label, proposal = _proposal_for(bid, proposals)
        if proposal:
            all_proposals.append(f"{bid} - {_plain_title(bid, result)}\n{proposal}")
        cards.append(f"""
        <article class="card">
          <header>
            <div><span class="bid">{html.escape(bid)}</span><h2>{html.escape(_plain_title(bid, result))}</h2></div>
            <div class="badges"><span>{html.escape(status_label)}</span><span>Preparation : {html.escape(cap_label)}{f' - {pct}%' if pct is not None else ''}</span></div>
          </header>
          <section><h3>Actions avant le depot</h3>{''.join(actions_html)}</section>
          {f'<section><h3>Question a transmettre</h3><p>{html.escape(question)}</p></section>' if question else ''}
          {f'<section class="proposal"><div class="proposal-head"><h3>{html.escape(label)}</h3><button type="button" class="copy-one">Copier</button></div><blockquote>{html.escape(proposal)}</blockquote></section>' if proposal else ''}
        </article>
        """)
    html_doc = f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Plan d'actions avant depot</title>
<style>
:root{{--navy:{PROFILE["navy"]};--magenta:{PROFILE["red"]};--text:{PROFILE["text"]};--muted:{PROFILE["muted"]};--line:{PROFILE["line"]};--soft:{PROFILE["soft"]};--amber:{PROFILE["amber"]}}}
*{{box-sizing:border-box}}body{{margin:0;background:#eef1f4;color:var(--text);font:15px/1.55 Calibri,Carlito,"Segoe UI",Arial,sans-serif;border-top:5px solid var(--magenta)}}
main{{max-width:980px;margin:28px auto;padding:{34 if STYLE_OPTION in ["2","4"] else 28}px {38 if STYLE_OPTION in ["2","4"] else 34}px;background:#fff;box-shadow:0 8px 30px rgba(20,40,60,.08)}}
.top{{background:{PROFILE["red"] if PROFILE["header_bg"]=="red" else ("#FFFFFF" if PROFILE["header_bg"]=="light" else PROFILE["navy"])};border-bottom:{6 if STYLE_OPTION in ["2","4"] else 4}px solid var(--magenta);padding:{30 if STYLE_OPTION=="4" else (16 if STYLE_OPTION=="3" else 22)}px 24px;display:flex;justify-content:space-between;gap:20px;align-items:flex-start}}
h1{{font-size:{31 if STYLE_OPTION in ["2","4"] else 28}px;line-height:1.15;color:{PROFILE["navy"] if PROFILE["header_bg"]=="light" else "#fff"};margin:0 0 8px}}.top .bid{{color:{PROFILE["red"] if PROFILE["header_bg"]=="light" else "#f2bfd1"}}}.meta{{color:{PROFILE["slate"] if PROFILE["header_bg"]=="light" else "#dce5e8"}}}
.toolbar{{display:flex;gap:8px;flex-wrap:wrap;margin:22px 0;padding:12px;background:var(--soft);border:1px solid var(--line);position:sticky;top:0;z-index:5}}
button{{font:inherit;padding:8px 13px;background:#fff;border:1px solid var(--navy);color:var(--navy);cursor:pointer}}button.primary{{background:var(--magenta);border-color:var(--magenta);color:#fff}}
.progress{{margin-left:auto;align-self:center;font-weight:700;color:var(--navy)}}
.card{{border:1px solid var(--line);border-left:5px solid var(--magenta);margin:0 0 16px;background:#fff}}
.card>header{{display:flex;justify-content:space-between;gap:18px;padding:14px 16px;background:#fafbfc;border-bottom:1px solid var(--line)}}
.card h2{{font-size:18px;margin:2px 0 0;color:var(--navy)}}.bid{{font-size:11px;color:var(--muted);font-weight:bold}}
.badges{{display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end}}.badges span{{border:1px solid var(--line);padding:4px 8px;font-size:12px;background:#fff}}
section{{padding:12px 16px;border-bottom:1px solid #edf0f2}}section:last-child{{border-bottom:0}}h3{{margin:0 0 8px;font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}}
.action{{display:grid;grid-template-columns:24px 1fr;gap:7px;align-items:start;padding:7px 0;border-bottom:1px dotted #dfe3e7}}.action:last-child{{border-bottom:0}}.action input{{width:17px;height:17px;margin-top:3px}}
.action input:checked+span{{text-decoration:line-through;color:#8b949e}}blockquote{{margin:0;padding:12px 14px;border-left:3px solid var(--magenta);background:var(--soft)}}.proposal-head{{display:flex;justify-content:space-between;align-items:center}}
@media print{{body{{background:#fff}}main{{box-shadow:none;margin:0;max-width:none;padding:12mm}}.toolbar,button{{display:none!important}}.card{{break-inside:avoid}}}}
</style></head><body><main>
<div class="top"><div><div class="bid">PLAN OPERATIONNEL</div><h1>Actions avant le dépôt de l'offre</h1><div class="meta">{html.escape(str(meta.get('entreprise','-')))} - Lot {html.escape(str(meta.get('lot','-')))} - {html.escape(str(meta.get('date','-')))}</div></div></div>
<div class="toolbar"><button id="copyAll">Copier toutes les propositions</button><button id="reset">Réinitialiser les cases</button><span class="progress" id="progress">0 / {action_index} actions réalisées</span></div>
{''.join(cards)}
<script>
const KEY={json.dumps(storage_key)};const boxes=[...document.querySelectorAll('[data-action-id]')];
function save(){{const data={{}};boxes.forEach(b=>data[b.dataset.actionId]=b.checked);localStorage.setItem(KEY,JSON.stringify(data));update();}}
function load(){{try{{const data=JSON.parse(localStorage.getItem(KEY)||'{{}}');boxes.forEach(b=>b.checked=!!data[b.dataset.actionId]);}}catch(e){{}}update();}}
function update(){{const done=boxes.filter(b=>b.checked).length;document.getElementById('progress').textContent=`${{done}} / ${{boxes.length}} actions réalisées`;}}
boxes.forEach(b=>b.addEventListener('change',save));
document.getElementById('reset').onclick=()=>{{if(confirm('Réinitialiser toutes les cases ?')){{boxes.forEach(b=>b.checked=false);save();}}}};
document.querySelectorAll('.copy-one').forEach(btn=>btn.onclick=async()=>{{const text=btn.closest('.proposal').querySelector('blockquote').innerText;await navigator.clipboard.writeText(text);const old=btn.textContent;btn.textContent='Copie';setTimeout(()=>btn.textContent=old,1200);}});
document.getElementById('copyAll').onclick=async()=>{{await navigator.clipboard.writeText({json.dumps('\n\n'.join(all_proposals), ensure_ascii=False)});const b=document.getElementById('copyAll');const old=b.textContent;b.textContent='Propositions copiées';setTimeout(()=>b.textContent=old,1200);}};load();
</script></main></body></html>"""
    out.write_text(html_doc, encoding="utf-8")

# ============================================================================
# PROTOTYPE V4 - DOSSIER D'INGENIERIE EXECUTIF
# Overrides des sorties principales. L'annexe technique reste celle de V2.
# ============================================================================
from reportlab.lib.pagesizes import landscape as _landscape

V4_NAVY = colors.HexColor('#082032')
V4_BLUE = colors.HexColor('#1F5D8C')
V4_SLATE = colors.HexColor('#49697D')
V4_PALE = colors.HexColor('#EAF3F8')
V4_PALE_2 = colors.HexColor('#F6F9FB')
V4_MAGENTA = colors.HexColor('#A1003D')
V4_TEXT = colors.HexColor('#17232B')
V4_MUTED = colors.HexColor('#667985')
V4_LINE = colors.HexColor('#C9D6DE')
V4_WHITE = colors.white
V4_AMBER = colors.HexColor('#A66A12')
V4_AMBER_BG = colors.HexColor('#FFF7E8')
V4_GREEN = colors.HexColor('#176B55')
V4_GREEN_BG = colors.HexColor('#EDF7F1')
V4_RED_BG = colors.HexColor('#FBEFF4')


def _v4_styles() -> Dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        'kicker': ParagraphStyle('v4_kicker', parent=base['Normal'], fontName=FONT_BOLD, fontSize=7.4, leading=9,
                                 textColor=V4_MAGENTA, uppercase=True, letterSpacing=.65, spaceAfter=5),
        'cover_title': ParagraphStyle('v4_cover_title', parent=base['Title'], fontName=FONT_LIGHT, fontSize=25,
                                     leading=29, textColor=V4_NAVY, spaceAfter=8),
        'cover_sub': ParagraphStyle('v4_cover_sub', parent=base['Normal'], fontName=FONT, fontSize=10.5,
                                   leading=14.5, textColor=V4_SLATE),
        'h1': ParagraphStyle('v4_h1', parent=base['Heading1'], fontName=FONT_LIGHT, fontSize=15.5, leading=19,
                             textColor=V4_NAVY, spaceBefore=2, spaceAfter=7),
        'h2': ParagraphStyle('v4_h2', parent=base['Heading2'], fontName=FONT_LIGHT, fontSize=11.3, leading=14,
                             textColor=V4_NAVY, spaceBefore=3, spaceAfter=4),
        'h3': ParagraphStyle('v4_h3', parent=base['Heading3'], fontName=FONT_BOLD, fontSize=8.1, leading=10,
                             textColor=V4_SLATE, uppercase=True, letterSpacing=.35, spaceBefore=2, spaceAfter=3),
        'body': ParagraphStyle('v4_body', parent=base['BodyText'], fontName=FONT, fontSize=8.9, leading=12.6,
                               textColor=V4_TEXT, spaceAfter=4),
        'small': ParagraphStyle('v4_small', parent=base['BodyText'], fontName=FONT, fontSize=7.8, leading=10.7,
                                textColor=V4_SLATE, spaceAfter=2.5),
        'tiny': ParagraphStyle('v4_tiny', parent=base['BodyText'], fontName=FONT, fontSize=6.8, leading=8.8,
                               textColor=V4_MUTED),
        'label': ParagraphStyle('v4_label', parent=base['Normal'], fontName=FONT_BOLD, fontSize=6.7, leading=8.2,
                                textColor=V4_MUTED, uppercase=True, letterSpacing=.42),
        'table': ParagraphStyle('v4_table', parent=base['BodyText'], fontName=FONT, fontSize=7.6, leading=10.2,
                                textColor=V4_TEXT),
        'table_bold': ParagraphStyle('v4_table_bold', parent=base['BodyText'], fontName=FONT_BOLD, fontSize=7.6,
                                     leading=10.2, textColor=V4_TEXT),
        'quote': ParagraphStyle('v4_quote', parent=base['BodyText'], fontName=FONT, fontSize=8.2, leading=11.5,
                                textColor=V4_TEXT, leftIndent=0),
        'decision': ParagraphStyle('v4_decision', parent=base['Heading1'], fontName=FONT_LIGHT, fontSize=18,
                                   leading=22, textColor=V4_NAVY, alignment=TA_LEFT),
        'metric': ParagraphStyle('v4_metric', parent=base['Normal'], fontName=FONT_LIGHT, fontSize=19, leading=21,
                                 textColor=V4_NAVY, alignment=TA_LEFT),
        'metric_label': ParagraphStyle('v4_metric_label', parent=base['Normal'], fontName=FONT_BOLD, fontSize=6.8,
                                       leading=8.4, textColor=V4_MUTED, uppercase=True, letterSpacing=.25),
    }

V4 = _v4_styles()


def _v4_status(result: Dict[str, Any]) -> Tuple[str, colors.Color, colors.Color]:
    key = _status_key(result)
    if key == 'CONFIRMÉE':
        return 'Confirmée pour le lot', V4_GREEN, V4_GREEN_BG
    if key in {'PROBABLE', 'PARTIELLE'}:
        return 'À confirmer pour le lot', V4_AMBER, V4_AMBER_BG
    if key == 'NON DÉMONTRÉE':
        return 'Non démontrée', V4_SLATE, V4_PALE
    if key == 'NON APPLICABLE':
        return 'Non applicable / exclue', V4_MUTED, colors.HexColor('#F0F3F5')
    return key or 'À vérifier', V4_SLATE, V4_PALE


def _v4_doc(path: Path, title: str, pagesize=A4, landscape_mode: bool = False, marque: str = "Entreprise", logo_bytes: bytes = None):
    page_w, page_h = pagesize
    doc = BaseDocTemplate(
        str(path), pagesize=pagesize,
        leftMargin=16*mm, rightMargin=16*mm, topMargin=18*mm, bottomMargin=16*mm,
        title=title, author=f'{_pdf_safe(marque)} - Outil analyse BIM',
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='v4')
    logo_reader = None
    if logo_bytes:
        try:
            logo_reader = ImageReader(io.BytesIO(logo_bytes))
        except Exception:
            logo_reader = None

    def on_page(canvas, document):
        canvas.saveState()
        page = canvas.getPageNumber()
        # Ligne éditoriale fine, pas de gros bandeau.
        canvas.setStrokeColor(V4_MAGENTA)
        canvas.setLineWidth(.8)
        canvas.line(document.leftMargin, page_h-11*mm, page_w-document.rightMargin, page_h-11*mm)
        label_x = document.leftMargin
        if logo_reader is not None:
            logo_h = 7 * mm
            try:
                iw, ih = logo_reader.getSize()
                logo_w = logo_h * (iw / ih) if ih else logo_h
            except Exception:
                logo_w = logo_h
            canvas.drawImage(logo_reader, document.leftMargin, page_h - 9.6*mm, width=logo_w, height=logo_h,
                              preserveAspectRatio=True, mask='auto')
            label_x = document.leftMargin + logo_w + 3*mm
        else:
            canvas.setFont(FONT_BOLD, 8)
            canvas.setFillColor(V4_NAVY)
            marque_text = _pdf_safe(marque)[:24]
            canvas.drawString(document.leftMargin, page_h-8.2*mm, marque_text)
            marque_w = canvas.stringWidth(marque_text, FONT_BOLD, 8)
            label_x = document.leftMargin + marque_w + 6*mm
        canvas.setFont(FONT, 6.3)
        canvas.setFillColor(V4_MUTED)
        canvas.drawString(label_x, page_h-8.2*mm, _pdf_safe(title)[:90])
        canvas.setStrokeColor(V4_LINE)
        canvas.setLineWidth(.45)
        canvas.line(document.leftMargin, 10*mm, page_w-document.rightMargin, 10*mm)
        canvas.setFont(FONT, 6.1)
        canvas.setFillColor(V4_MUTED)
        canvas.drawString(document.leftMargin, 6.6*mm, 'Analyse automatisée - vérification humaine recommandée avant engagement contractuel')
        canvas.drawRightString(page_w-document.rightMargin, 6.6*mm, f'Page {page}')
        canvas.restoreState()

    doc.addPageTemplates([PageTemplate(id='v4_main', frames=[frame], onPage=on_page)])
    return doc


def _v4_section(text: str) -> List[Any]:
    return [Spacer(1, 2), _para(text, V4['h1']), HRFlowable(width='100%', thickness=.55, color=V4_LINE, spaceAfter=6)]


def _v4_pill(text: str, fg: colors.Color, bg: colors.Color, width=44*mm) -> Table:
    pstyle = ParagraphStyle('v4_pill', parent=V4['table_bold'], textColor=fg, alignment=TA_CENTER, fontSize=6.8, leading=8)
    t = Table([[_para(text, pstyle)]], colWidths=[width])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),bg), ('BOX',(0,0),(-1,-1),.45,fg),
        ('LEFTPADDING',(0,0),(-1,-1),4), ('RIGHTPADDING',(0,0),(-1,-1),4),
        ('TOPPADDING',(0,0),(-1,-1),3), ('BOTTOMPADDING',(0,0),(-1,-1),3),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    return t


def _v4_metric(value: str, label: str, accent: colors.Color) -> Table:
    t = Table([
        [_para(value, ParagraphStyle('v4_mv', parent=V4['metric'], textColor=accent))],
        [_para(label, V4['metric_label'])],
    ], colWidths=[50*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),V4_WHITE), ('BOX',(0,0),(-1,-1),.45,V4_LINE),
        ('LINEBEFORE',(0,0),(0,-1),2.2,accent),
        ('LEFTPADDING',(0,0),(-1,-1),8), ('RIGHTPADDING',(0,0),(-1,-1),6),
        ('TOPPADDING',(0,0),(-1,0),7), ('BOTTOMPADDING',(0,0),(-1,0),1),
        ('TOPPADDING',(0,1),(-1,1),0), ('BOTTOMPADDING',(0,1),(-1,1),6),
    ]))
    return t


def _v4_source(result: Dict[str, Any]) -> str:
    parts = []
    if result.get('cctp_pg'):
        parts.append(f"CCTP p.{result.get('cctp_pg')}")
    if result.get('conv_pg'):
        parts.append(f"Convention p.{result.get('conv_pg')}")
    return ' - '.join(parts) if parts else 'Aucune preuve ciblée retenue'


def _v4_decision_text(confirmed: int, clarify: int, excluded: int, scores: Dict[str, Any], contradictions: Sequence[Any]) -> Tuple[str,str]:
    if contradictions:
        return ('Décision à suspendre', 'Une contradiction documentaire doit être levée avant tout engagement ferme dans l’offre.')
    if confirmed and clarify:
        return ('Réponse possible sous réserves', f'{confirmed} obligation(s) sont confirmée(s) et {clarify} point(s) doivent être sécurisés avant le dépôt.')
    if confirmed:
        return ('Réponse possible', f'{confirmed} obligation(s) sont confirmée(s). Les engagements proposés doivent être adaptés aux capacités réelles de l’entreprise.')
    if clarify:
        return ('Clarification requise', f'Aucune obligation directe n’est confirmée, mais {clarify} point(s) nécessitent une réponse écrite.')
    return ('Aucun engagement BIM majeur identifié', 'Les documents analysés ne démontrent pas d’obligation opérationnelle significative pour le lot.')


def generate_action_plan_pdf(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    out.parent.mkdir(parents=True,exist_ok=True)
    relevant=_relevant_items(results,scores)
    pagesize=_landscape(A4)
    doc=_v4_doc(out,"Plan d'actions avant le dépôt de l'offre",pagesize=pagesize,landscape_mode=True)
    story=[_para("PLAN D'ACTIONS AVANT LE DÉPÔT DE L'OFFRE",V4['cover_title']),_para(f"{meta.get('entreprise','-')} - Lot {meta.get('lot','-')}",V4['cover_sub']),Spacer(1,4),_para("Document opérationnel : actions, responsabilités et échéances. Les questions et formulations sont regroupées dans la seconde partie.",V4['small']),Spacer(1,8)]
    rows=[[_para('FAIT',V4['label']),_para('PRIORITÉ',V4['label']),_para('ACTION',V4['label']),_para('BLOC',V4['label']),_para('RESPONSABLE',V4['label']),_para('ÉCHÉANCE',V4['label'])]]
    for bid,r in relevant:
        _,prio=_priority(r,bid,scores)
        for a in _actions_for(bid,r,scores):
            rows.append([_para('□',V4['table_bold']),_para(prio,V4['table_bold']),_para(a,V4['table']),_para(bid,V4['table_bold']),_para('',V4['table']),_para('Avant dépôt',V4['table'])])
    t=Table(rows,colWidths=[12*mm,25*mm,142*mm,16*mm,42*mm,27*mm],repeatRows=1)
    styles=[('BACKGROUND',(0,0),(-1,0),V4_NAVY),('TEXTCOLOR',(0,0),(-1,0),V4_WHITE),('BOX',(0,0),(-1,-1),.45,V4_LINE),('INNERGRID',(0,0),(-1,-1),.3,V4_LINE),('VALIGN',(0,0),(-1,-1),'TOP'),('ALIGN',(0,1),(0,-1),'CENTER'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]
    for i in range(1,len(rows)):
        if i%2==0: styles.append(('BACKGROUND',(0,i),(-1,i),V4_PALE_2))
    t.setStyle(TableStyle(styles)); story += [t,PageBreak(),_para("QUESTIONS ET FORMULATIONS POUR L'OFFRE",V4['h1']),Spacer(1,5)]
    for bid,r in relevant:
        q=_plain_language((r.get('bloc') or {}).get('question_bim_manager',''))
        label,prop=_proposal_for(bid,proposals)
        if not q and not prop: continue
        left=[_rich_para(f'<b>{_esc(bid)} - {_esc(_plain_title(bid,r))}</b>',V4['h2'])]
        if q: left += [_para('QUESTION À TRANSMETTRE',V4['h3']),_para(q,V4['small'])]
        right=[]
        if prop: right += [_para('FORMULATION POUR L’OFFRE',V4['h3']),_para(prop,V4['small'])]
        if right:
            box=Table([[left,right]],colWidths=[129*mm,135*mm])
            box_style=[('BACKGROUND',(0,0),(0,0),V4_PALE_2),('LINEAFTER',(0,0),(0,0),.35,V4_LINE)]
        else:
            box=Table([[left]],colWidths=[264*mm])
            box_style=[('BACKGROUND',(0,0),(0,0),V4_PALE_2)]
        box.setStyle(TableStyle(box_style+[('BOX',(0,0),(-1,-1),.45,V4_LINE),('LINEBEFORE',(0,0),(0,0),2,V4_MAGENTA),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        story += [box,Spacer(1,6)]
    doc.build(story)


def generate_action_plan_html(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    relevant=_relevant_items(results,scores)
    storage_key=re.sub(r'[^a-z0-9]+','-',f"bim-actions-v4-{meta.get('projet','')}-{meta.get('entreprise','')}-{meta.get('lot','')}".lower()).strip('-')
    rows=[]; idx=0; props=[]
    for bid,r in relevant:
        sl,_,_=_v4_status(r); _,prio=_priority(r,bid,scores); q=_plain_language((r.get('bloc') or {}).get('question_bim_manager','')); label,prop=_proposal_for(bid,proposals)
        for a in _actions_for(bid,r,scores):
            idx+=1; aid=f'a{idx}'
            rows.append(f'''<article class="action-row" data-status="{html.escape(sl)}" data-block="{html.escape(bid)}">
              <label class="check"><input type="checkbox" data-action-id="{aid}"><span></span></label>
              <div class="block"><b>{html.escape(bid)}</b><small>{html.escape(_plain_title(bid,r))}</small></div>
              <div class="main"><strong>{html.escape(a)}</strong>{f'<details><summary>Question à transmettre</summary><p>{html.escape(q)}</p></details>' if q else ''}{f'<details><summary>Formulation pour l’offre</summary><blockquote>{html.escape(prop)}</blockquote><button class="copy-one" type="button">Copier la formulation</button></details>' if prop else ''}</div>
              <div class="priority">{html.escape(prio)}</div>
              <label class="editable"><span>Responsable</span><input data-field="owner" placeholder="À désigner"></label>
              <label class="editable"><span>Échéance</span><input data-field="date" value="Avant dépôt"></label>
            </article>''')
        if prop: props.append(f"{bid} - {_plain_title(bid,r)}\n{prop}")
    doc=f'''<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Plan d'actions BIM</title>
<style>
:root{{--navy:#082032;--blue:#1F5D8C;--slate:#49697D;--pale:#EAF3F8;--pale2:#F6F9FB;--magenta:#A1003D;--text:#17232B;--muted:#667985;--line:#C9D6DE;--white:#fff}}
*{{box-sizing:border-box}}body{{margin:0;background:#f2f5f7;color:var(--text);font:14px/1.45 Calibri,Carlito,"Segoe UI",Arial,sans-serif}}button,input{{font:inherit}}.shell{{min-height:100vh;display:grid;grid-template-columns:245px 1fr}}
aside{{background:var(--navy);color:#fff;padding:28px 22px;position:sticky;top:0;height:100vh}}.brand{{font-size:26px;font-weight:800;border-bottom:2px solid var(--magenta);padding-bottom:16px;margin-bottom:24px}}aside h1{{font-size:18px;line-height:1.25;margin:0 0 8px}}aside p{{color:#cad6dc;font-size:12px}}.progress{{margin:28px 0}}.progress b{{font-size:30px;display:block}}.bar{{height:5px;background:#38505e;margin-top:8px}}.bar i{{height:100%;display:block;background:var(--magenta);width:0}}.aside-actions button{{display:block;width:100%;margin:8px 0;padding:10px;border:1px solid #8ba0ab;background:transparent;color:#fff;cursor:pointer}}.aside-actions .primary{{background:var(--magenta);border-color:var(--magenta)}}
main{{padding:28px 34px;max-width:1500px;width:100%}}.top{{display:flex;justify-content:space-between;gap:24px;align-items:end;border-bottom:1px solid var(--line);padding-bottom:16px}}.top h2{{font-size:27px;line-height:1.15;color:var(--navy);margin:0 0 6px}}.meta{{color:var(--muted)}}.summary{{display:grid;grid-template-columns:repeat(3,minmax(160px,1fr));gap:12px;margin:18px 0}}.metric{{background:#fff;border:1px solid var(--line);border-left:4px solid var(--blue);padding:14px}}.metric b{{font-size:24px;color:var(--navy);display:block}}.metric span{{font-size:11px;text-transform:uppercase;color:var(--muted)}}
.filters{{display:flex;gap:8px;align-items:center;margin:14px 0}}.filters button{{border:1px solid var(--line);background:#fff;padding:8px 12px;color:var(--navy);cursor:pointer}}.filters button.active{{background:var(--navy);color:#fff}}.filters input{{margin-left:auto;width:280px;padding:9px;border:1px solid var(--line)}}
.list{{display:grid;gap:8px}}.action-row{{display:grid;grid-template-columns:32px 110px minmax(360px,1fr) 105px 145px 130px;gap:12px;align-items:start;background:#fff;border:1px solid var(--line);padding:13px 14px}}.check input{{display:none}}.check span{{display:block;width:19px;height:19px;border:1.5px solid var(--slate);margin-top:3px;cursor:pointer}}.check input:checked+span{{background:var(--magenta);border-color:var(--magenta);box-shadow:inset 0 0 0 4px #fff}}.block b{{display:inline-block;background:var(--navy);color:#fff;padding:4px 7px}}.block small{{display:block;color:var(--muted);margin-top:5px}}.main strong{{font-weight:600}}details{{margin-top:7px;border-top:1px dotted var(--line);padding-top:5px}}summary{{color:var(--blue);cursor:pointer;font-weight:600;font-size:12px}}details p,blockquote{{margin:7px 0;color:#3e505a}}blockquote{{border-left:3px solid var(--magenta);padding-left:10px}}.copy-one{{border:1px solid var(--magenta);background:#fff;color:var(--magenta);padding:5px 8px;cursor:pointer}}.priority{{color:var(--magenta);font-weight:700;font-size:12px;padding-top:4px}}.editable span{{display:block;font-size:10px;text-transform:uppercase;color:var(--muted);margin-bottom:3px}}.editable input{{width:100%;border:0;border-bottom:1px solid var(--line);padding:4px 0;color:var(--text)}}.action-row.done{{opacity:.55}}.action-row.done .main>strong{{text-decoration:line-through}}
@media(max-width:1100px){{.shell{{grid-template-columns:1fr}}aside{{height:auto;position:static}}.action-row{{grid-template-columns:30px 90px 1fr}}.priority,.editable{{grid-column:auto}}}}@media print{{body{{background:#fff}}aside,.filters,.copy-one{{display:none!important}}.shell{{display:block}}main{{padding:0}}.action-row{{break-inside:avoid;grid-template-columns:25px 90px 1fr 90px 110px 95px}}}}
h1,h2,h3,.brand,.metric b{{font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-weight:300}}
</style></head><body><div class="shell"><aside><div class="brand">BNN</div><h1>Plan d'actions avant le dépôt</h1><p>{html.escape(str(meta.get('entreprise','-')))} - Lot {html.escape(str(meta.get('lot','-')))}</p><div class="progress"><b id="progressText">0 / {idx}</b><span>actions réalisées</span><div class="bar"><i id="progressBar"></i></div></div><div class="aside-actions"><button id="copyAll">Copier les formulations</button><button id="reset">Réinitialiser</button></div></aside><main><header class="top"><div><h2>Actions à engager</h2><div class="meta">Responsables et échéances modifiables - mémorisation locale</div></div><div class="meta">Édition {html.escape(str(meta.get('date','-')))}</div></header><section class="summary"><div class="metric"><b>{idx}</b><span>Actions identifiées</span></div><div class="metric"><b>{len(relevant)}</b><span>Blocs concernés</span></div><div class="metric"><b>{sum(1 for _,r in relevant if _status_key(r)=='CONFIRMÉE')}</b><span>Obligations confirmées</span></div></section><div class="filters"><button class="active" data-filter="all">Toutes</button><button data-filter="Immédiate">Prioritaires</button><button data-filter="Avant dépôt">Avant dépôt</button><input id="search" placeholder="Rechercher une action ou un bloc"></div><section class="list">{''.join(rows)}</section></main></div>
<script>const KEY={json.dumps(storage_key)};const rows=[...document.querySelectorAll('.action-row')];const checks=[...document.querySelectorAll('[data-action-id]')];
function state(){{try{{return JSON.parse(localStorage.getItem(KEY)||'{{}}')}}catch(e){{return {{}}}}}}function save(){{const d=state();checks.forEach(c=>d[c.dataset.actionId]=c.checked);document.querySelectorAll('[data-field]').forEach((x,i)=>d['f'+i]=x.value);localStorage.setItem(KEY,JSON.stringify(d));update()}}function load(){{const d=state();checks.forEach(c=>c.checked=!!d[c.dataset.actionId]);document.querySelectorAll('[data-field]').forEach((x,i)=>{{if(d['f'+i]!==undefined)x.value=d['f'+i]}});update()}}function update(){{const n=checks.filter(c=>c.checked).length;document.getElementById('progressText').textContent=n+' / '+checks.length;document.getElementById('progressBar').style.width=(checks.length?100*n/checks.length:0)+'%';rows.forEach(r=>r.classList.toggle('done',r.querySelector('[data-action-id]').checked))}}checks.forEach(c=>c.onchange=save);document.querySelectorAll('[data-field]').forEach(x=>x.oninput=save);document.getElementById('reset').onclick=()=>{{if(confirm('Réinitialiser les données ?')){{localStorage.removeItem(KEY);load()}}}};document.getElementById('copyAll').onclick=async()=>{{await navigator.clipboard.writeText({json.dumps(chr(10).join(props),ensure_ascii=False)});}};
document.querySelectorAll('.copy-one').forEach(b=>b.onclick=async()=>navigator.clipboard.writeText(b.closest('details').querySelector('blockquote').innerText));document.querySelectorAll('[data-filter]').forEach(b=>b.onclick=()=>{{document.querySelectorAll('[data-filter]').forEach(x=>x.classList.remove('active'));b.classList.add('active');const f=b.dataset.filter;rows.forEach(r=>r.hidden=f!=='all'&&!r.querySelector('.priority').innerText.includes(f))}});document.getElementById('search').oninput=e=>{{const q=e.target.value.toLowerCase();rows.forEach(r=>r.hidden=!r.innerText.toLowerCase().includes(q))}};load();</script></body></html>'''
    out.write_text(doc,encoding='utf-8')


# ============================================================================
# Version consolidée - plan d'action en quatre colonnes métier
# Ces définitions finales remplacent les versions précédentes.
# ============================================================================
def _v5_action_block_data(
    block_id: str,
    result: Dict[str, Any],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    block = result.get("bloc") or {}
    status_label, _, _ = _v4_status(result)
    _, priority = _priority(result, block_id, scores)
    capacity_label, pct, _, _ = _capacity_info(block_id, scores)
    strength, gap = _situation_summary(block_id, scores)
    demand = _plain_language(
        result.get("public_demand")
        or block.get("lecture_entreprise")
        or block.get("description_convention")
        or result.get("conclusion")
        or "L'attente exacte n'est pas suffisamment explicite dans les extraits retenus."
    )
    proposal_label, proposal = _proposal_for(block_id, proposals)
    if not proposal:
        proposal = "Aucune formulation ferme ne doit être intégrée tant que le périmètre et la capacité réelle ne sont pas confirmés."
    actions = _actions_for(block_id, result, scores)
    question = _plain_language(block.get("question_bim_manager", ""))
    return {
        "id": block_id,
        "title": _plain_title(block_id, result),
        "status": status_label,
        "priority": priority,
        "capacity_label": capacity_label,
        "pct": pct,
        "strength": strength,
        "gap": gap,
        "demand": demand,
        "proposal_label": proposal_label,
        "proposal": proposal,
        "actions": actions,
        "question": question,
    }


def generate_action_plan_pdf(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    relevant = _relevant_items(results, scores)
    blocks = [_v5_action_block_data(bid, result, scores, proposals) for bid, result in relevant]
    pagesize = _landscape(A4)
    doc = _v4_doc(out, "Plan d'actions - réponse et mise en conformité", pagesize=pagesize, landscape_mode=True)
    story: List[Any] = [
        _para("PLAN D'ACTIONS AVANT LE DÉPÔT DE L'OFFRE", V4["cover_title"]),
        _para(f"{meta.get('entreprise','-')} - Lot {meta.get('lot','-')}", V4["cover_sub"]),
        Spacer(1, 4),
        _para(
            "Lecture par exigence : demande contractuelle, capacité actuelle, réponse proposée et actions nécessaires pour rendre cette réponse conforme.",
            V4["small"],
        ),
        Spacer(1, 8),
    ]

    metrics = Table([[
        _v4_metric(str(len(blocks)), "BLOCS À TRAITER", V4_BLUE),
        _v4_metric(str(sum(len(b["actions"]) for b in blocks)), "ACTIONS IDENTIFIÉES", V4_MAGENTA),
        _v4_metric(str(sum(1 for b in blocks if b["pct"] is not None)), "CAPACITÉS ÉVALUÉES", V4_GREEN),
    ]], colWidths=[88*mm] * 3)
    metrics.setStyle(TableStyle([
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
    ]))
    story += [metrics, Spacer(1, 10)]

    col_widths = [53*mm, 58*mm, 68*mm, 85*mm]
    header_style = ParagraphStyle(
        "v5_col_header", parent=V4["label"], textColor=V4_WHITE,
        fontSize=7.1, leading=8.5, alignment=TA_LEFT,
    )
    cell_style = ParagraphStyle(
        "v5_cell", parent=V4["table"], fontSize=7.1, leading=9.1,
        textColor=V4_TEXT,
    )
    cell_bold = ParagraphStyle(
        "v5_cell_bold", parent=cell_style, fontName=FONT_BOLD,
    )
    note_style = ParagraphStyle(
        "v5_note", parent=V4["tiny"], fontSize=6.3, leading=7.6,
        textColor=V4_MUTED,
    )

    for index, block in enumerate(blocks):
        prep = f"{block['capacity_label']} - {block['pct']} %" if block["pct"] is not None else block["capacity_label"]
        block_head = Table([[
            _rich_para(f"<b>{_esc(block['id'])} - {_esc(block['title'])}</b>", V4["h2"]),
            _para(block["status"], V4["table_bold"]),
            _para(block["priority"], V4["table_bold"]),
            _para(prep, V4["table_bold"]),
        ]], colWidths=[120*mm, 55*mm, 52*mm, 37*mm])
        block_head.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (0,0), V4_PALE_2),
            ("BACKGROUND", (1,0), (-1,0), V4_PALE),
            ("BOX", (0,0), (-1,-1), .5, V4_LINE),
            ("LINEBEFORE", (0,0), (0,0), 2.6, V4_MAGENTA),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))

        capacity_parts: List[Any] = [
            _para(prep, cell_bold),
            Spacer(1, 2),
            _para(block["strength"], cell_style),
            Spacer(1, 3),
            _para("POINT À RENFORCER", note_style),
            _para(block["gap"], cell_style),
        ]
        response_parts: List[Any] = [
            _para(block["proposal_label"].upper(), note_style),
            _para(block["proposal"], cell_style),
        ]
        action_parts: List[Any] = []
        for action in block["actions"]:
            action_parts.append(_rich_para(f"□&nbsp;&nbsp;{_esc(action)}", cell_style))
            action_parts.append(Spacer(1, 2))
        if block["question"]:
            action_parts += [
                Spacer(1, 2),
                _para("QUESTION À TRANSMETTRE", note_style),
                _para(block["question"], cell_style),
            ]
        action_parts += [
            Spacer(1, 4),
            _para("Responsable : ____________________", note_style),
            _para("Échéance : avant le dépôt", note_style),
        ]

        table = Table([
            [
                _para("CE QUI EST DEMANDÉ", header_style),
                _para("CE QUE JE SUIS CAPABLE DE FAIRE EN CE MOMENT", header_style),
                _para("CE QUE JE DOIS RÉPONDRE", header_style),
                _para("CE QUE JE DOIS FAIRE POUR ÊTRE CONFORME", header_style),
            ],
            [
                [_para(block["demand"], cell_style)],
                capacity_parts,
                response_parts,
                action_parts,
            ],
        ], colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), V4_NAVY),
            ("TEXTCOLOR", (0,0), (-1,0), V4_WHITE),
            ("BACKGROUND", (0,1), (-1,1), V4_WHITE),
            ("BOX", (0,0), (-1,-1), .5, V4_LINE),
            ("INNERGRID", (0,0), (-1,-1), .35, V4_LINE),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,0), 5),
            ("BOTTOMPADDING", (0,0), (-1,0), 5),
            ("TOPPADDING", (0,1), (-1,1), 7),
            ("BOTTOMPADDING", (0,1), (-1,1), 7),
        ]))
        story += [KeepTogether([block_head, table]), Spacer(1, 8)]

    story += [
        Spacer(1, 2),
        _para(
            "Les formulations proposées doivent être relues avant dépôt. Une capacité déclarée ne constitue pas, à elle seule, une preuve contractuelle ou opérationnelle.",
            V4["small"],
        ),
    ]
    doc.build(story)


def generate_action_plan_html(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    relevant = _relevant_items(results, scores)
    blocks = [_v5_action_block_data(bid, result, scores, proposals) for bid, result in relevant]
    glossary = glossary_lookup(meta.get("glossary_entries") or [])
    aliases = sorted(glossary.keys(), key=len, reverse=True)
    alias_pattern = re.compile(r"(?<![\w])(" + "|".join(re.escape(a) for a in aliases) + r")(?![\w])", re.IGNORECASE) if aliases else None

    def tip_text(value: Any) -> str:
        raw = _clean(value)
        if not raw:
            return ""
        escaped = html.escape(raw)
        if not alias_pattern:
            return escaped
        def repl(match):
            visible = match.group(0)
            entry = glossary.get(visible.casefold())
            if not entry:
                return html.escape(visible)
            definition = html.escape(str(entry.get("definition", "")))
            source = html.escape(str(entry.get("source", "")))
            source_html = f'<small>{source}</small>' if source else ''
            return f'<span class="term-tip" tabindex="0"><span>{html.escape(visible)}</span><span class="term-bubble">{definition}{source_html}</span></span>'
        return alias_pattern.sub(repl, escaped)

    storage_key = re.sub(
        r"[^a-z0-9]+", "-",
        f"bim-actions-columns-{meta.get('projet','')}-{meta.get('entreprise','')}-{meta.get('lot','')}".lower(),
    ).strip("-")

    body_rows: List[str] = []
    all_responses: List[str] = []
    action_count = 0
    for block in blocks:
        prep = f"{block['capacity_label']} - {block['pct']} %" if block["pct"] is not None else block["capacity_label"]
        action_html = []
        for action in block["actions"]:
            action_count += 1
            action_id = f"action-{action_count}"
            action_html.append(
                f'''<label class="task"><input type="checkbox" data-action-id="{action_id}"><span>{html.escape(action)}</span></label>'''
            )
        question = (
            f'''<div class="question"><b>Question à transmettre</b><p>{html.escape(block['question'])}</p></div>'''
            if block["question"] else ""
        )
        body_rows.append(f'''
        <tbody class="action-block" data-status="{html.escape(block['status'])}" data-priority="{html.escape(block['priority'])}" data-block="{html.escape(block['id'])}">
          <tr class="block-heading"><th colspan="4"><div><strong>{html.escape(block['id'])} - {html.escape(block['title'])}</strong><span>{html.escape(block['status'])}</span><span>{html.escape(block['priority'])}</span><span>{html.escape(prep)}</span></div></th></tr>
          <tr class="content-row">
            <td data-label="Ce qui est demandé"><p>{html.escape(block['demand'])}</p></td>
            <td data-label="Ce que je suis capable de faire en ce moment"><p class="capacity"><b>{html.escape(prep)}</b></p><p>{html.escape(block['strength'])}</p><div class="gap"><b>Point à renforcer</b><p>{html.escape(block['gap'])}</p></div></td>
            <td data-label="Ce que je dois répondre"><span class="response-label">{html.escape(block['proposal_label'])}</span><blockquote>{html.escape(block['proposal'])}</blockquote><button class="copy-response" type="button">Copier la réponse</button></td>
            <td data-label="Ce que je dois faire pour être conforme"><div class="tasks">{''.join(action_html)}</div>{question}<div class="tracking"><label>Responsable<input data-field="owner" placeholder="À désigner"></label><label>Échéance<input data-field="date" value="Avant dépôt"></label></div></td>
          </tr>
        </tbody>''')
        all_responses.append(f"{block['id']} - {block['title']}\n{block['proposal']}")

    confirmed = sum(1 for bid, result in relevant if _status_key(result) == "CONFIRMÉE")
    doc = f'''<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Plan d'actions BIM - {html.escape(str(meta.get('projet','')))}</title>
<style>
:root{{--navy:#071C24;--blue:#1F5D8C;--slate:#49697D;--pale:#EAF3F8;--pale2:#F7FAFC;--magenta:#A1003D;--magenta-light:#F06A9A;--text:#17232B;--muted:#667985;--line:#C9D6DE;--green:#176B55;--amber:#A66A12;--white:#fff}}
*{{box-sizing:border-box}}body{{margin:0;background:#f2f5f7;color:var(--text);font:14px/1.48 Calibri,Carlito,"Segoe UI",Arial,sans-serif}}button,input{{font:inherit}}h1,h2,.metric b{{font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif}}.hero{{background:var(--navy);border-bottom:5px solid var(--magenta);padding:27px max(25px,calc((100vw - 1500px)/2));display:flex;align-items:center;gap:22px}}.brand{{font-size:29px;font-weight:800;color:var(--magenta-light);border-right:1px solid rgba(240,106,154,.45);padding-right:22px}}.hero h1{{margin:0;color:var(--magenta-light);font-size:31px;line-height:1;font-weight:700}}.hero p{{margin:7px 0 0;color:#F5C8D8;font-weight:600}}.hero-actions{{margin-left:auto;display:flex;gap:8px}}.hero-actions button{{border:1px solid rgba(240,106,154,.55);background:transparent;color:#F5C8D8;padding:9px 12px;cursor:pointer;font-weight:700}}.hero-actions .primary{{background:var(--magenta);border-color:var(--magenta);color:#fff}}.container{{max-width:1500px;margin:0 auto;padding:22px}}.summary{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px}}.metric{{background:#fff;border:1px solid var(--line);border-left:4px solid var(--blue);padding:13px 16px}}.metric b{{display:block;font-size:27px;color:var(--navy);font-weight:700}}.metric span{{font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:700}}.progress-box{{background:#fff;border:1px solid var(--line);padding:12px 15px;margin-bottom:14px;display:grid;grid-template-columns:180px 1fr;gap:15px;align-items:center}}.progress-box strong{{color:var(--navy)}}.progress{{height:8px;background:#E5EDF1}}.progress i{{display:block;height:100%;background:var(--magenta);width:0;transition:width .25s}}.toolbar{{display:flex;gap:8px;align-items:center;background:#fff;border:1px solid var(--line);padding:10px 12px;margin-bottom:12px;flex-wrap:wrap}}.toolbar button{{border:1px solid var(--line);background:#fff;color:var(--navy);padding:7px 10px;cursor:pointer;font-weight:700}}.toolbar button.active{{background:var(--navy);color:#fff}}.toolbar input{{margin-left:auto;width:280px;padding:8px;border:1px solid var(--line)}}.table-wrap{{overflow-x:auto;background:#fff;border:1px solid var(--line)}}table{{width:100%;border-collapse:collapse;table-layout:fixed}}col.demand{{width:22%}}col.capacity{{width:22%}}col.response{{width:26%}}col.compliance{{width:30%}}thead th{{position:sticky;top:0;z-index:2;background:var(--navy);color:#fff;text-align:left;padding:10px;font-size:11px;text-transform:uppercase;letter-spacing:.04em;border-right:1px solid #39505A}}thead th:last-child{{border-right:0}}.block-heading th{{background:var(--pale2);padding:9px 10px;border-top:3px solid var(--magenta);border-bottom:1px solid var(--line)}}.block-heading div{{display:flex;align-items:center;gap:8px;flex-wrap:wrap}}.block-heading strong{{color:var(--navy);font-size:15px;margin-right:auto}}.block-heading span{{border:1px solid var(--line);background:#fff;color:var(--slate);font-size:11px;padding:3px 7px;font-weight:700}}.content-row td{{vertical-align:top;padding:13px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);overflow-wrap:anywhere}}.content-row td:last-child{{border-right:0}}.content-row p{{margin:0 0 8px}}.capacity b{{color:var(--blue)}}.gap,.question{{background:var(--pale2);border-left:3px solid var(--blue);padding:8px 9px;margin-top:10px}}.gap b,.question b{{font-size:10px;text-transform:uppercase;color:var(--magenta)}}.gap p,.question p{{margin:4px 0 0}}.response-label{{display:block;color:var(--magenta);font-size:10px;text-transform:uppercase;font-weight:700;margin-bottom:5px}}blockquote{{margin:0 0 9px;border-left:3px solid var(--magenta);padding-left:10px;color:#314A59}}.copy-response{{border:1px solid var(--magenta);color:var(--magenta);background:#fff;padding:6px 9px;cursor:pointer;font-weight:700}}.tasks{{display:grid;gap:7px}}.task{{display:grid;grid-template-columns:18px 1fr;gap:7px;align-items:start;cursor:pointer}}.task input{{margin:3px 0 0;accent-color:var(--magenta)}}.task input:checked+span{{text-decoration:line-through;color:var(--muted)}}.tracking{{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:12px;padding-top:9px;border-top:1px solid var(--line)}}.tracking label{{font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:700}}.tracking input{{width:100%;border:0;border-bottom:1px solid var(--line);padding:5px 0;color:var(--text);text-transform:none;font-weight:400}}.action-block.done{{opacity:.62}}
@media(max-width:950px){{.hero{{flex-wrap:wrap}}.hero-actions{{margin-left:0}}.summary{{grid-template-columns:1fr}}.progress-box{{grid-template-columns:1fr}}.toolbar input{{margin-left:0;width:100%}}.table-wrap{{overflow:visible;border:0;background:transparent}}table,thead,tbody,tr,th,td{{display:block}}thead{{display:none}}.action-block{{display:block;margin-bottom:14px;border:1px solid var(--line);background:#fff}}.block-heading th{{display:block}}.content-row td{{border-right:0;padding:12px}}.content-row td::before{{content:attr(data-label);display:block;font-size:10px;text-transform:uppercase;color:var(--magenta);font-weight:700;margin-bottom:7px}}}}
@media print{{body{{background:#fff;font-size:10px}}.hero{{background:#fff;border-bottom:2px solid var(--magenta);padding:8px 0}}.brand,.hero h1{{color:var(--navy)}}.hero p{{color:var(--muted)}}.hero-actions,.toolbar,.copy-response{{display:none!important}}.container{{max-width:none;padding:8px 0}}.summary{{grid-template-columns:repeat(3,1fr)}}.table-wrap{{overflow:visible}}thead th{{position:static}}.action-block{{break-inside:avoid}}}}
</style></head><body>
<header class="hero"><div class="brand">BNN</div><div><h1>PLAN D'ACTIONS AVANT LE DÉPÔT</h1><p>{html.escape(str(meta.get('entreprise','-')))} - Lot {html.escape(str(meta.get('lot','-')))} - Lecture en quatre colonnes métier</p></div><div class="hero-actions"><button id="copyAll" type="button">Copier toutes les réponses</button><button id="reset" type="button">Réinitialiser</button></div></header>
<main class="container"><section class="summary"><div class="metric"><b>{len(blocks)}</b><span>Blocs à traiter</span></div><div class="metric"><b>{action_count}</b><span>Actions identifiées</span></div><div class="metric"><b>{confirmed}</b><span>Obligations confirmées</span></div></section><section class="progress-box"><strong id="progressText">0 / {action_count} actions réalisées</strong><div class="progress"><i id="progressBar"></i></div></section><div class="toolbar"><button class="active" data-filter="all" type="button">Tous</button><button data-filter="Immédiate" type="button">Immédiates</button><button data-filter="Avant dépôt" type="button">Avant dépôt</button><button data-filter="À surveiller" type="button">À surveiller</button><input id="search" placeholder="Rechercher un bloc, une demande ou une action"></div><div class="table-wrap"><table><colgroup><col class="demand"><col class="capacity"><col class="response"><col class="compliance"></colgroup><thead><tr><th>Ce qui est demandé</th><th>Ce que je suis capable de faire en ce moment</th><th>Ce que je dois répondre</th><th>Ce que je dois faire pour être conforme</th></tr></thead>{''.join(body_rows)}</table></div></main>
<script>
const KEY={json.dumps(storage_key)};const blocks=[...document.querySelectorAll('.action-block')];const checks=[...document.querySelectorAll('[data-action-id]')];const fields=[...document.querySelectorAll('[data-field]')];
function readState(){{try{{return JSON.parse(localStorage.getItem(KEY)||'{{}}')}}catch(e){{return {{}}}}}}function saveState(){{const state=readState();checks.forEach(c=>state[c.dataset.actionId]=c.checked);fields.forEach((f,i)=>state['field-'+i]=f.value);localStorage.setItem(KEY,JSON.stringify(state));update();}}function loadState(){{const state=readState();checks.forEach(c=>c.checked=!!state[c.dataset.actionId]);fields.forEach((f,i)=>{{if(state['field-'+i]!==undefined)f.value=state['field-'+i]}});update();}}function update(){{const done=checks.filter(c=>c.checked).length;document.getElementById('progressText').textContent=done+' / '+checks.length+' actions réalisées';document.getElementById('progressBar').style.width=(checks.length?100*done/checks.length:0)+'%';blocks.forEach(b=>{{const local=[...b.querySelectorAll('[data-action-id]')];b.classList.toggle('done',local.length>0&&local.every(c=>c.checked));}});}}
checks.forEach(c=>c.addEventListener('change',saveState));fields.forEach(f=>f.addEventListener('input',saveState));document.getElementById('reset').addEventListener('click',()=>{{if(confirm('Réinitialiser les actions, responsables et échéances ?')){{localStorage.removeItem(KEY);loadState();}}}});document.getElementById('copyAll').addEventListener('click',async()=>{{await navigator.clipboard.writeText({json.dumps(chr(10)+chr(10).join(all_responses), ensure_ascii=False)});}});document.querySelectorAll('.copy-response').forEach(button=>button.addEventListener('click',async()=>{{await navigator.clipboard.writeText(button.closest('td').querySelector('blockquote').innerText);button.textContent='Copié';setTimeout(()=>button.textContent='Copier la réponse',1200);}}));document.querySelectorAll('[data-filter]').forEach(button=>button.addEventListener('click',()=>{{document.querySelectorAll('[data-filter]').forEach(x=>x.classList.remove('active'));button.classList.add('active');const filter=button.dataset.filter;blocks.forEach(block=>block.hidden=filter!=='all'&&!block.dataset.priority.includes(filter));}}));document.getElementById('search').addEventListener('input',event=>{{const query=event.target.value.toLowerCase();blocks.forEach(block=>block.hidden=!block.innerText.toLowerCase().includes(query));}});loadState();
</script></body></html>'''
    out.write_text(doc, encoding="utf-8")

# ============================================================================
# Version V6 - 3 colonnes métier + ligne de conformité pleine largeur
# Ces définitions finales remplacent la version V5.
# ============================================================================
def generate_action_plan_pdf(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    """Plan d'action A4 paysage : 3 colonnes + conformité sur toute la largeur."""
    out.parent.mkdir(parents=True, exist_ok=True)
    relevant = _relevant_items(results, scores)
    blocks = [_v5_action_block_data(bid, result, scores, proposals) for bid, result in relevant]
    pagesize = _landscape(A4)
    doc = _v4_doc(out, "Plan d'actions - réponse et mise en conformité", pagesize=pagesize, landscape_mode=True, marque=meta.get("marque") or meta.get("entreprise") or "Entreprise", logo_bytes=meta.get("logo_bytes"))
    story: List[Any] = [
        _para("PLAN D'ACTIONS AVANT LE DÉPÔT DE L'OFFRE", V4["cover_title"]),
        _para(f"{meta.get('entreprise','-')} - Lot {meta.get('lot','-')}", V4["cover_sub"]),
        Spacer(1, 4),
        _para(
            "Pour chaque exigence : demande contractuelle, capacité actuelle et réponse proposée. "
            "Les actions de mise en conformité sont regroupées juste en dessous sur toute la largeur.",
            V4["small"],
        ),
        Spacer(1, 8),
    ]

    metrics = Table([[
        _v4_metric(str(len(blocks)), "BLOCS À TRAITER", V4_BLUE),
        _v4_metric(str(sum(len(b["actions"]) for b in blocks)), "ACTIONS IDENTIFIÉES", V4_MAGENTA),
        _v4_metric(str(sum(1 for b in blocks if b["pct"] is not None)), "CAPACITÉS ÉVALUÉES", V4_GREEN),
    ]], colWidths=[88*mm] * 3)
    metrics.setStyle(TableStyle([
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
    ]))
    story += [metrics, Spacer(1, 10)]

    # Checklist récapitulative par niveau de priorité (même logique que la
    # version HTML) : aucun niveau n'empêche à proprement parler le dépôt de
    # l'offre, ce sont des degrés d'importance à traiter.
    _ORDRE_PRIORITE = ["Critique", "Haute", "Avant dépôt", "À surveiller"]
    _NOTE_PRIORITE = {
        "Critique": "à traiter avant de finaliser l'offre",
        "Haute": "à consolider en interne, ne bloque pas le dépôt",
        "Avant dépôt": "réserve ou question à formuler dans l'offre",
        "À surveiller": "sans urgence immédiate",
    }
    _groupes_pdf: Dict[str, List[str]] = {p: [] for p in _ORDRE_PRIORITE}
    for block in blocks:
        prio = next((p for p in _ORDRE_PRIORITE if block["priority"].startswith(p)), "À surveiller")
        for item in block.get("action_items") or []:
            texte = item.get("text", "") if isinstance(item, dict) else str(item)
            if texte:
                _groupes_pdf[prio].append(f"{block['id']} — {texte}")
    _checklist_title_style = ParagraphStyle(
        "checklist_group_title", parent=V4["h2"], fontSize=9.5, leading=11.5, textColor=V4_MAGENTA,
    )
    _checklist_item_style = ParagraphStyle(
        "checklist_item", parent=V4["table"], fontSize=7.6, leading=9.6,
    )
    if any(_groupes_pdf.values()):
        story.append(_para("CHECKLIST DES ACTIONS PAR NIVEAU DE PRIORITÉ", V4["h2"]))
        story.append(Spacer(1, 3))
        for prio in _ORDRE_PRIORITE:
            if not _groupes_pdf[prio]:
                continue
            story.append(_para(f"{prio} — {_NOTE_PRIORITE[prio]}", _checklist_title_style))
            for line in _groupes_pdf[prio]:
                story.append(_rich_para(f"□&nbsp;&nbsp;{_esc(line)}", _checklist_item_style))
            story.append(Spacer(1, 4))
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width="100%", thickness=0.6, color=V4_LINE, spaceAfter=8))

    # Colonne d'identité du bloc à gauche, puis les trois colonnes métier.
    col_widths = [42*mm, 63*mm, 70*mm, 89*mm]
    header_style = ParagraphStyle(
        "v6_col_header", parent=V4["label"], textColor=V4_WHITE,
        fontSize=7.3, leading=8.7, alignment=TA_LEFT,
    )
    cell_style = ParagraphStyle(
        "v6_cell", parent=V4["table"], fontSize=7.25, leading=9.25,
        textColor=V4_TEXT,
    )
    cell_bold = ParagraphStyle(
        "v6_cell_bold", parent=cell_style, fontName=FONT_BOLD,
    )
    note_style = ParagraphStyle(
        "v6_note", parent=V4["tiny"], fontSize=6.4, leading=7.7,
        textColor=V4_MUTED,
    )
    compliance_title = ParagraphStyle(
        "v6_compliance_title", parent=V4["label"], fontName=FONT_BOLD,
        fontSize=7.2, leading=8.6, textColor=V4_MAGENTA,
    )

    for block in blocks:
        prep = f"{block['capacity_label']} - {block['pct']} %" if block["pct"] is not None else block["capacity_label"]
        meta_parts: List[Any] = [
            _rich_para(f"<b>{_esc(block['id'])}</b>", V4["h2"]),
            _para(block["title"], V4["table_bold"]),
            Spacer(1, 7),
            _para("STATUT", note_style),
            _para(block["status"], cell_bold),
            Spacer(1, 6),
            _para("PRIORITÉ", note_style),
            _para(block["priority"], cell_bold),
            Spacer(1, 6),
            _para("CAPACITÉ ACTUELLE", note_style),
            _para(prep, cell_bold),
        ]

        capacity_parts: List[Any] = [
            _para(prep, cell_bold), Spacer(1, 2),
            _para(block["strength"], cell_style), Spacer(1, 3),
            _para("POINT À RENFORCER", note_style),
            _para(block["gap"], cell_style),
        ]
        response_parts: List[Any] = [
            _para(block["proposal_label"].upper(), note_style),
            _para(block["proposal"], cell_style),
        ]

        compliance_parts: List[Any] = [
            _para("CE QUE JE DOIS FAIRE POUR ÊTRE CONFORME", compliance_title),
            Spacer(1, 4),
        ]
        for action in block["actions"]:
            compliance_parts.append(_rich_para(f"□&nbsp;&nbsp;{_esc(action)}", cell_style))
            compliance_parts.append(Spacer(1, 2))
        if block["question"]:
            compliance_parts += [
                Spacer(1, 2),
                _para("QUESTION À TRANSMETTRE", note_style),
                _para(block["question"], cell_style),
            ]
        if block["actions"] or block["question"]:
            compliance_parts += [
                Spacer(1, 5),
                HRFlowable(width="100%", thickness=.35, color=V4_LINE, spaceBefore=0, spaceAfter=5),
                _para("Responsable : ______________________________", note_style),
                Spacer(1, 3),
                _para("Échéance : avant le dépôt", note_style),
            ]
        else:
            compliance_parts.append(
                _para("Aucune action supplémentaire : exigence déjà confirmée et capacité démontrée.", cell_style)
            )

        table = Table([
            [
                meta_parts,
                _para("CE QUI EST DEMANDÉ", header_style),
                _para("CE QUE JE SUIS CAPABLE DE FAIRE EN CE MOMENT", header_style),
                _para("CE QUE JE DOIS RÉPONDRE", header_style),
            ],
            [
                "",
                [_para(block["demand"], cell_style)],
                capacity_parts,
                response_parts,
            ],
            ["", compliance_parts, "", ""],
        ], colWidths=col_widths)
        table.setStyle(TableStyle([
            ("SPAN", (0,0), (0,2)),
            ("SPAN", (1,2), (3,2)),
            ("BACKGROUND", (0,0), (0,2), V4_PALE),
            ("BACKGROUND", (1,0), (3,0), V4_NAVY),
            ("TEXTCOLOR", (1,0), (3,0), V4_WHITE),
            ("BACKGROUND", (1,1), (3,1), V4_WHITE),
            ("BACKGROUND", (1,2), (3,2), V4_PALE_2),
            ("BOX", (0,0), (-1,-1), .5, V4_LINE),
            ("LINEBEFORE", (0,0), (0,2), 2.8, V4_MAGENTA),
            ("INNERGRID", (1,0), (3,1), .35, V4_LINE),
            ("LINEAFTER", (0,0), (0,2), .7, V4_BLUE),
            ("LINEABOVE", (1,2), (3,2), 1.1, V4_BLUE),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 7),
            ("RIGHTPADDING", (0,0), (-1,-1), 7),
            ("TOPPADDING", (0,0), (-1,0), 6),
            ("BOTTOMPADDING", (0,0), (-1,0), 6),
            ("TOPPADDING", (0,1), (-1,1), 7),
            ("BOTTOMPADDING", (0,1), (-1,1), 7),
            ("TOPPADDING", (0,2), (-1,2), 7),
            ("BOTTOMPADDING", (0,2), (-1,2), 7),
        ]))
        story += [KeepTogether([table]), Spacer(1, 8)]

    story += [
        Spacer(1, 2),
        _para(
            "Les formulations proposées doivent être relues avant dépôt. Une capacité déclarée ne constitue pas, à elle seule, une preuve contractuelle ou opérationnelle.",
            V4["small"],
        ),
    ]

    doc.build(story)


def generate_action_plan_html(
    out: Path,
    results: Dict[str, Dict[str, Any]],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
    meta: Dict[str, Any],
) -> None:
    """Plan HTML : identité du bloc à gauche, 3 colonnes métier et conformité pleine largeur."""
    relevant = _relevant_items(results, scores)
    blocks = [_v5_action_block_data(bid, result, scores, proposals) for bid, result in relevant]
    glossary = glossary_lookup(meta.get("glossary_entries") or [])
    aliases = sorted(glossary.keys(), key=len, reverse=True)
    alias_pattern = re.compile(
        r"(?<![\w])(" + "|".join(re.escape(alias) for alias in aliases) + r")(?![\w])",
        re.IGNORECASE,
    ) if aliases else None

    def tip_text(value: Any) -> str:
        raw = _clean(value)
        if not raw:
            return ""
        escaped = html.escape(raw)
        if not alias_pattern:
            return escaped
        seen: set[str] = set()

        def repl(match: re.Match[str]) -> str:
            visible = match.group(0)
            key = visible.casefold()
            entry = glossary.get(key)
            if not entry:
                return html.escape(visible)
            entry_key = str(entry.get("term") or key).casefold()
            if entry_key in seen:
                return html.escape(visible)
            seen.add(entry_key)
            definition = html.escape(str(entry.get("definition", "")))
            source = html.escape(str(entry.get("source", "")))
            source_html = f'<small>{source}</small>' if source else ""
            return (
                f'<span class="term-tip" tabindex="0"><span>{html.escape(visible)}</span>'
                f'<span class="term-bubble">{definition}{source_html}</span></span>'
            )

        return alias_pattern.sub(repl, escaped)

    storage_key = re.sub(
        r"[^a-z0-9]+", "-",
        f"bim-actions-3cols-{meta.get('projet','')}-{meta.get('entreprise','')}-{meta.get('lot','')}".lower(),
    ).strip("-")

    body_rows: List[str] = []
    all_responses: List[str] = []
    _all_actions: List[Tuple[str, str, str, str]] = []  # (action_id, block_id, priority, text)
    action_count = 0
    for block in blocks:
        prep = f"{block['capacity_label']} - {block['pct']} %" if block["pct"] is not None else block["capacity_label"]
        action_html: List[str] = []
        for item in block.get("action_items") or []:
            action_count += 1
            action_id = f"action-{action_count}"
            action = item.get("text", "") if isinstance(item, dict) else str(item)
            origin = item.get("origin", "") if isinstance(item, dict) else ""
            origin_html = f'<small class="action-origin">Origine : {tip_text(origin)}</small>' if origin else ""
            action_html.append(
                f'<label class="task"><input type="checkbox" data-action-id="{action_id}"><span>{tip_text(action)}{origin_html}</span></label>'
            )
            if action:
                _all_actions.append((action_id, block["id"], block["priority"], action))
        question = (
            f'<div class="question"><b>Question à transmettre</b><p>{tip_text(block["question"])}</p></div>'
            if block["question"] else ""
        )
        gap_html = (
            f'<div class="gap"><b>Point à renforcer</b><p>{tip_text(block.get("gap", ""))}</p></div>'
            if block.get("gap") else ""
        )
        has_compliance_work = bool(action_html) or bool(block["question"])
        if has_compliance_work:
            tasks_block = (
                f'<div class="tasks">{"".join(action_html)}</div>{question}'
                '<div class="tracking"><label>Responsable<input data-field="owner" placeholder="À désigner"></label>'
                '<label>Échéance<input data-field="date" value="Avant dépôt"></label></div>'
            )
        else:
            tasks_block = '<p class="compliance-ok">Aucune action supplémentaire : exigence déjà confirmée et capacité démontrée.</p>'
        body_rows.append(
            f'<tbody class="action-block" data-status="{html.escape(block["status"])}" '
            f'data-priority="{html.escape(block["priority"])}" data-block="{html.escape(block["id"])}">'
            f'<tr class="content-row">'
            f'<td class="block-meta" rowspan="2"><strong class="block-code">{html.escape(block["id"])}</strong>'
            f'<div class="block-title">{tip_text(block["title"])}</div>'
            f'<div class="meta-item"><span>Statut</span><b>{html.escape(block["status"])}</b></div>'
            f'<div class="meta-item"><span>Priorité</span><b>{html.escape(block["priority"])}</b></div>'
            f'<div class="meta-item"><span>Capacité actuelle</span><b>{html.escape(prep)}</b></div></td>'
            f'<td data-label="Ce qui est demandé"><p>{tip_text(block["demand"])}</p></td>'
            f'<td data-label="Ce que je suis capable de faire en ce moment"><p class="capacity"><b>{tip_text(prep)}</b></p>'
            f'<p>{tip_text(block["strength"])}</p>{gap_html}</td>'
            f'<td data-label="Ce que je dois répondre"><span class="response-label">{html.escape(block["proposal_label"])}</span>'
            f'<blockquote>{tip_text(block["proposal"])}</blockquote><button class="copy-response" type="button">Copier la réponse</button></td></tr>'
            f'<tr class="compliance-row"><td colspan="3"><div class="compliance-title">Ce que je dois faire pour être conforme</div>'
            f'<div class="compliance-layout">{tasks_block}</div></td></tr></tbody>'
        )
        all_responses.append(f"{block['id']} - {block['title']}\n{block['proposal']}")

    confirmed = sum(1 for _, result in relevant if _status_key(result) == "CONFIRMÉE")

    # Checklist récapitulative : toutes les actions "à faire", regroupées par
    # niveau de priorité (Critique / Haute / Avant dépôt / À surveiller). Aucun
    # de ces niveaux n'empêche à proprement parler de déposer une offre : ce
    # sont des degrés d'importance à traiter, pas des blocages de dépôt.
    # Les cases à cocher réutilisent le même data-action-id que le tableau
    # principal : cocher une ligne ici coche aussi la ligne correspondante
    # dans le bloc détaillé, et inversement.
    _ORDRE_PRIORITE = ["Critique", "Haute", "Avant dépôt", "À surveiller"]
    _groupes: Dict[str, List[str]] = {p: [] for p in _ORDRE_PRIORITE}
    for action_id, block_id, priority, texte in _all_actions:
        prio = next((p for p in _ORDRE_PRIORITE if priority.startswith(p)), "À surveiller")
        _groupes[prio].append(
            f'<label class="task checklist-task"><input type="checkbox" data-action-id="{action_id}">'
            f'<span><b>{html.escape(block_id)}</b> — {tip_text(texte)}</span></label>'
        )
    _sections_checklist = []
    _CLASSES = {"Critique": "critique", "Haute": "haute", "Avant dépôt": "avant", "À surveiller": "surveiller"}
    _NOTE = {
        "Critique": "à traiter avant de finaliser l'offre",
        "Haute": "à consolider en interne, ne bloque pas le dépôt",
        "Avant dépôt": "réserve ou question à formuler dans l'offre",
        "À surveiller": "sans urgence immédiate",
    }
    for prio in _ORDRE_PRIORITE:
        if _groupes[prio]:
            _sections_checklist.append(
                f'<div class="checklist-group checklist-{_CLASSES[prio]}">'
                f'<h3>{html.escape(prio)} <small>({html.escape(_NOTE[prio])})</small></h3>{"".join(_groupes[prio])}</div>'
            )
    checklist_html = (
        '<details class="checklist" open><summary>Checklist des actions par niveau de priorité</summary>'
        f'<div class="checklist-stack">{"".join(_sections_checklist)}</div></details>'
    ) if _sections_checklist else ""

    css = r'''
:root{--navy:#071C24;--blue:#1F5D8C;--slate:#49697D;--pale:#EAF3F8;--pale2:#F7FAFC;--magenta:#A1003D;--magenta-light:#9C5B45;--text:#17232B;--muted:#667985;--line:#C9D6DE;--green:#176B55;--amber:#A66A12;--white:#fff}
*{box-sizing:border-box}body{margin:0;background:#f2f5f7;color:var(--text);font:14px/1.48 Calibri,Carlito,"Segoe UI",Arial,sans-serif}button,input{font:inherit}h1,h2,.metric b{font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif}.hero{background:var(--navy);border-bottom:5px solid var(--magenta);padding:27px max(25px,calc((100vw - 1500px)/2));display:flex;align-items:center;gap:22px}.brand{font-size:29px;font-weight:800;color:var(--magenta-light);border-right:1px solid rgba(240,106,154,.45);padding-right:22px}.brand img{max-height:40px;max-width:160px;display:block;object-fit:contain}.hero h1{margin:0;color:var(--magenta-light);font-size:31px;line-height:1;font-weight:700}.hero p{margin:7px 0 0;color:#F5C8D8;font-weight:600}.hero-actions{margin-left:auto;display:flex;gap:8px}.hero-actions button{border:1px solid rgba(240,106,154,.55);background:transparent;color:#F5C8D8;padding:9px 12px;cursor:pointer;font-weight:700}.hero-actions .primary{background:var(--magenta);border-color:var(--magenta);color:#fff}.container{max-width:1500px;margin:0 auto;padding:22px}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}.metric{position:relative;background:#fff;border:1px solid var(--line);border-radius:6px;padding:16px 18px 16px 20px;box-shadow:0 1px 3px rgba(7,28,36,.06);overflow:hidden;transition:transform .15s,box-shadow .15s}.metric:hover{transform:translateY(-2px);box-shadow:0 6px 14px rgba(7,28,36,.1)}.metric::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;border-radius:6px 0 0 6px}.metric b{display:block;font-size:30px;color:var(--navy);font-weight:800;line-height:1}.metric span{display:block;margin-top:5px;font-size:10.5px;text-transform:uppercase;color:var(--muted);font-weight:700;letter-spacing:.03em}.metric-blocks::before{background:var(--blue)}.metric-actions::before{background:var(--magenta-light)}.metric-confirmed::before{background:#3E8A5C}.metric-confirmed b{color:#3E8A5C}.progress-box{background:#fff;border:1px solid var(--line);border-radius:6px;padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:16px;box-shadow:0 1px 3px rgba(7,28,36,.06)}.progress-box strong{color:var(--navy);font-size:13px;white-space:nowrap}.progress{flex:1;height:10px;background:#E5EDF1;border-radius:6px;overflow:hidden}.progress i{display:block;height:100%;background:linear-gradient(90deg,var(--magenta-light),var(--magenta));width:0;border-radius:6px;transition:width .3s ease}.toolbar{display:flex;gap:8px;align-items:center;background:#fff;border:1px solid var(--line);padding:10px 12px;margin-bottom:12px;flex-wrap:wrap}.toolbar button{border:1px solid var(--line);background:#fff;color:var(--navy);padding:7px 10px;cursor:pointer;font-weight:700}.toolbar button.active{background:var(--navy);color:#fff}.toolbar input{margin-left:auto;width:280px;padding:8px;border:1px solid var(--line)}.table-wrap{overflow:visible;background:#fff;border:1px solid var(--line)}table{width:100%;border-collapse:collapse;table-layout:fixed}col.meta{width:18%}col.demand{width:24%}col.capacity{width:27%}col.response{width:31%}.sticky-column-head{position:sticky;top:0;z-index:60;display:grid;grid-template-columns:18% 24% 27% 31%;background:var(--navy);color:#fff;border-bottom:3px solid var(--magenta);box-shadow:0 5px 12px rgba(7,28,36,.22)}.sticky-column-head span{padding:11px 10px;font-size:11px;text-transform:uppercase;letter-spacing:.04em;font-weight:800;border-right:1px solid #39505A}.sticky-column-head span:last-child{border-right:0}thead{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)}.content-row td{vertical-align:top;padding:14px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);border-top:3px solid var(--magenta);overflow-wrap:anywhere}.content-row td:last-child{border-right:0}.block-meta{background:var(--pale);border-right:1px solid var(--blue)!important}.block-code{display:block;color:var(--magenta);font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:23px;line-height:1}.block-title{margin:5px 0 18px;color:var(--navy);font-weight:700;font-size:15px}.meta-item{padding:10px 0;border-top:1px solid var(--line)}.meta-item span{display:block;margin-bottom:4px;color:var(--muted);font-size:10px;text-transform:uppercase;font-weight:700}.meta-item b{display:block;color:var(--navy);font-size:12px;line-height:1.35}.content-row p{margin:0 0 8px}.capacity b{color:var(--blue)}.gap,.question{background:var(--pale2);border-left:3px solid var(--blue);padding:8px 9px;margin-top:10px}.gap b,.question b{font-size:10px;text-transform:uppercase;color:var(--magenta)}.gap p,.question p{margin:4px 0 0}.response-label{display:block;color:var(--magenta);font-size:10px;text-transform:uppercase;font-weight:700;margin-bottom:5px}blockquote{margin:0 0 9px;border-left:3px solid var(--magenta);padding-left:10px;color:#314A59}.copy-response{border:1px solid var(--magenta);color:var(--magenta);background:#fff;padding:6px 9px;cursor:pointer;font-weight:700}.compliance-row td{padding:13px 14px 14px;background:var(--pale2);border-bottom:1px solid var(--line);border-top:2px solid var(--blue)}.compliance-title{color:var(--magenta);font-size:11px;text-transform:uppercase;font-weight:800;letter-spacing:.045em;margin-bottom:9px}.compliance-layout{display:block}.tasks{display:block}.task{display:grid;grid-template-columns:18px 1fr;gap:7px;align-items:start;cursor:pointer;margin-bottom:8px}.task input{margin:3px 0 0;accent-color:var(--magenta)}.task input:checked+span{text-decoration:line-through;color:var(--muted)}.action-origin{display:block;margin-top:4px;color:var(--muted);font-size:10px;font-style:italic;text-decoration:none!important}.term-tip{position:relative;display:inline-block;outline:none}.term-tip>span:first-child{border-bottom:1px dotted var(--magenta);cursor:help}.term-bubble{display:none;position:absolute;left:0;bottom:calc(100% + 7px);z-index:100;width:300px;background:var(--navy);color:#fff;padding:10px 11px;font-size:12px;line-height:1.4;text-transform:none;font-weight:400;box-shadow:0 6px 18px rgba(0,0,0,.25)}.term-bubble small{display:block;margin-top:6px;padding-top:5px;border-top:1px solid rgba(255,255,255,.25);color:#F5C8D8}.term-tip:hover .term-bubble,.term-tip:focus .term-bubble{display:block}.tracking{display:block;margin-top:12px;padding-top:10px;border-top:1px solid var(--line)}.tracking label{display:block;font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:700;margin-top:8px}.tracking input{width:100%;border:0;border-bottom:1px solid var(--line);padding:5px 0;color:var(--text);text-transform:none;font-weight:400}.action-block.done{opacity:.62}.compliance-ok{margin:0;color:#3E8A5C;font-weight:700}.no-results{background:#fff;border:1px solid var(--line);border-left:4px solid var(--muted);padding:16px;margin-top:12px;color:var(--muted);font-weight:700;text-align:center}
.checklist{background:#fff;border:1px solid var(--line);padding:16px 18px;margin-bottom:14px;box-shadow:0 1px 3px rgba(7,28,36,.05)}.checklist summary{cursor:pointer;font-weight:800;color:var(--navy);text-transform:uppercase;font-size:12.5px;letter-spacing:.04em;display:flex;align-items:center;gap:8px}.checklist summary::before{content:"☰";color:var(--magenta-light);font-size:14px}.checklist-stack{display:flex;flex-direction:column;gap:14px;margin-top:14px}.checklist-group{background:var(--pale2);border-left:4px solid var(--line);border-radius:0 4px 4px 0;padding:12px 14px}.checklist-group h3{margin:0 0 10px;font-size:11.5px;text-transform:uppercase;letter-spacing:.03em;color:var(--navy);font-weight:800;display:flex;align-items:baseline;gap:6px}.checklist-group h3 small{color:var(--muted);text-transform:none;font-weight:400;font-size:11px}.checklist-task{display:grid;grid-template-columns:18px 1fr;gap:9px;align-items:start;padding:6px 4px;border-radius:3px;cursor:pointer;transition:background .12s}.checklist-task:hover{background:rgba(255,255,255,.75)}.checklist-task+.checklist-task{border-top:1px solid rgba(201,214,222,.6)}.checklist-task input{margin-top:2px;width:15px;height:15px;accent-color:var(--magenta-light)}.checklist-task span{line-height:1.45}.checklist-task span b{color:var(--navy);font-weight:800}.checklist-task input:checked+span{text-decoration:line-through;color:var(--muted)}.checklist-critique{border-left-color:#8A3B2E}.checklist-critique h3{color:#8A3B2E}.checklist-haute{border-left-color:var(--magenta-light)}.checklist-haute h3{color:var(--magenta-light)}.checklist-avant{border-left-color:var(--blue)}.checklist-avant h3{color:var(--blue)}.checklist-surveiller{border-left-color:var(--muted)}.checklist-surveiller h3{color:var(--muted)}
@media print{.checklist{break-inside:avoid}}
@media(max-width:950px){.hero{flex-wrap:wrap}.sticky-column-head{display:none}.hero-actions{margin-left:0}.summary{grid-template-columns:1fr}.progress-box{flex-direction:column;align-items:stretch}.toolbar input{margin-left:0;width:100%}.table-wrap{overflow:visible;border:0;background:transparent}table,thead,tbody,tr,th,td{display:block}thead{display:none}.action-block{display:block;margin-bottom:14px;border:1px solid var(--line);background:#fff}.block-meta{display:block;border-right:0!important}.content-row td{border-right:0;padding:12px}.content-row td::before{content:attr(data-label);display:block;font-size:10px;text-transform:uppercase;color:var(--magenta);font-weight:700;margin-bottom:7px}.block-meta::before{display:none!important}.compliance-row td{display:block}}
@media print{body{background:#fff;font-size:10px}.hero{background:#fff;border-bottom:2px solid var(--magenta);padding:8px 0}.brand,.hero h1{color:var(--navy)}.hero p{color:var(--muted)}.hero-actions,.toolbar,.copy-response{display:none!important}.container{max-width:none;padding:8px 0}.summary{grid-template-columns:repeat(3,1fr)}.table-wrap{overflow:visible}thead th{position:static}.action-block{break-inside:avoid}}
'''
    script = f'''
const KEY={json.dumps(storage_key)};const blocks=[...document.querySelectorAll('.action-block')];const checks=[...document.querySelectorAll('[data-action-id]')];const fields=[...document.querySelectorAll('[data-field]')];
function readState(){{try{{return JSON.parse(localStorage.getItem(KEY)||'{{}}')}}catch(e){{return {{}}}}}}
function saveState(){{const state=readState();checks.forEach(c=>state[c.dataset.actionId]=c.checked);fields.forEach((f,i)=>state['field-'+i]=f.value);localStorage.setItem(KEY,JSON.stringify(state));update();}}
function loadState(){{const state=readState();checks.forEach(c=>c.checked=!!state[c.dataset.actionId]);fields.forEach((f,i)=>{{if(state['field-'+i]!==undefined)f.value=state['field-'+i]}});update();}}
function update(){{const ids=[...new Set(checks.map(c=>c.dataset.actionId))];const doneIds=ids.filter(id=>checks.some(c=>c.dataset.actionId===id&&c.checked));document.getElementById('progressText').textContent=doneIds.length+' / '+ids.length+' actions réalisées';document.getElementById('progressBar').style.width=(ids.length?100*doneIds.length/ids.length:0)+'%';blocks.forEach(b=>{{const local=[...b.querySelectorAll('[data-action-id]')];b.classList.toggle('done',local.length>0&&local.every(c=>c.checked));}});}}
function syncAction(id,checked){{checks.forEach(c=>{{if(c.dataset.actionId===id)c.checked=checked;}});}}
function copyText(text){{if(navigator.clipboard&&window.isSecureContext){{return navigator.clipboard.writeText(text);}}const ta=document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.focus();ta.select();let ok=false;try{{ok=document.execCommand('copy');}}catch(e){{ok=false;}}document.body.removeChild(ta);return ok?Promise.resolve():Promise.reject(new Error('copy-unavailable'));}}
checks.forEach(c=>c.addEventListener('change',()=>{{syncAction(c.dataset.actionId,c.checked);saveState();}}));fields.forEach(f=>f.addEventListener('input',saveState));document.getElementById('reset').addEventListener('click',()=>{{if(confirm('Réinitialiser les actions, responsables et échéances ?')){{localStorage.removeItem(KEY);loadState();}}}});document.getElementById('copyAll').addEventListener('click',()=>{{copyText({json.dumps(chr(10)+chr(10).join(all_responses), ensure_ascii=False)}).then(()=>{{const b=document.getElementById('copyAll');const t=b.textContent;b.textContent='Copié';setTimeout(()=>b.textContent=t,1200);}}).catch(()=>alert('Copie automatique indisponible dans ce navigateur. Sélectionnez et copiez le texte manuellement.'));}});document.querySelectorAll('.copy-response').forEach(button=>button.addEventListener('click',()=>{{copyText(button.closest('td').querySelector('blockquote').innerText).then(()=>{{button.textContent='Copié';setTimeout(()=>button.textContent='Copier la réponse',1200);}}).catch(()=>alert('Copie automatique indisponible dans ce navigateur. Sélectionnez et copiez le texte manuellement.'));}}));function updateEmptyState(){{const visible=blocks.some(b=>!b.hidden);const noRes=document.getElementById('noResults');if(noRes)noRes.hidden=visible;}}document.querySelectorAll('[data-filter]').forEach(button=>button.addEventListener('click',()=>{{document.querySelectorAll('[data-filter]').forEach(x=>x.classList.remove('active'));button.classList.add('active');const filter=button.dataset.filter;blocks.forEach(block=>block.hidden=filter!=='all'&&!block.dataset.priority.includes(filter));updateEmptyState();}}));document.getElementById('search').addEventListener('input',event=>{{const query=event.target.value.toLowerCase();blocks.forEach(block=>block.hidden=!block.innerText.toLowerCase().includes(query));updateEmptyState();}});loadState();updateEmptyState();
'''
    if meta.get("logo_data_uri"):
        brand_html = f'<img src="{html.escape(str(meta.get("logo_data_uri")))}" alt="Logo">'
    else:
        brand_html = html.escape(str(meta.get("marque") or meta.get("entreprise") or "Entreprise"))

    doc = (
        '<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
        f'<title>Plan d\'actions BIM - {html.escape(str(meta.get("projet", "")))}</title><style>{css}</style></head><body>'
        f'<header class="hero"><div class="brand">{brand_html}</div><div><h1>PLAN D\'ACTIONS AVANT LE DÉPÔT</h1>'
        f'<p>{html.escape(str(meta.get("entreprise", "-")))} - Lot {html.escape(str(meta.get("lot", "-")))} - Demande, capacité et réponse par exigence, avec suivi de mise en conformité</p></div>'
        '<div class="hero-actions"><button id="copyAll" type="button">Copier toutes les réponses</button><button id="reset" type="button">Réinitialiser</button></div></header>'
        f'<main class="container"><section class="summary"><div class="metric metric-blocks"><b>{len(blocks)}</b><span>Blocs à traiter</span></div><div class="metric metric-actions"><b>{action_count}</b><span>Actions identifiées</span></div><div class="metric metric-confirmed"><b>{confirmed}</b><span>Obligations confirmées</span></div></section>'
        f'<section class="progress-box"><strong id="progressText">0 / {action_count} actions réalisées</strong><div class="progress"><i id="progressBar"></i></div></section>'
        f'{checklist_html}'
        '<div class="toolbar"><button class="active" data-filter="all" type="button">Tous</button><button data-filter="Critique" type="button">Critique</button><button data-filter="Haute" type="button">Haute</button><button data-filter="Avant dépôt" type="button">Avant dépôt</button><button data-filter="À surveiller" type="button">À surveiller</button><input id="search" placeholder="Rechercher un bloc, une demande ou une action"></div>'
        f'<div class="table-wrap"><div class="sticky-column-head"><span>Bloc</span><span>Ce qui est demandé</span><span>Ce que je suis capable de faire en ce moment</span><span>Ce que je dois répondre</span></div><table><colgroup><col class="meta"><col class="demand"><col class="capacity"><col class="response"></colgroup><thead><tr><th>Bloc</th><th>Ce qui est demandé</th><th>Ce que je suis capable de faire en ce moment</th><th>Ce que je dois répondre</th></tr></thead>{"".join(body_rows)}</table></div>'
        '<p id="noResults" class="no-results" hidden>Aucun bloc ne correspond à ce filtre pour ce lot.</p>'
        '</main>'
        f'<script>{script}</script></body></html>'
    )
    out.write_text(doc, encoding="utf-8")

# ============================================================================
# Version V7 - modèle public unique partagé avec le dashboard.
# Les fonctions ci-dessous remplacent les helpers précédents au moment de
# l'exécution et garantissent les mêmes statuts, capacités, décisions, réponses
# et actions dans tous les livrables.
# ============================================================================
from bim_model import block_sort_key as _model_block_sort_key
from bim_model import ensure_public_model as _ensure_public_model


def _status_key(result: Dict[str, Any]) -> str:
    public = result.get("public_status_key")
    return {
        "CONFIRMED": "CONFIRMÉE",
        "CLARIFY": "PROBABLE",
        "UNDEMONSTRATED": "NON DÉMONTRÉE",
        "EXCLUDED": "NON APPLICABLE",
    }.get(public, _clean(result.get("statut", "")).upper())


def _status_info(result: Dict[str, Any]) -> Tuple[str, colors.Color, colors.Color]:
    key = result.get("public_status_key")
    if key == "CONFIRMED":
        return result.get("public_status_label", "Confirmée pour le lot"), GREEN, GREEN_BG
    if key == "CLARIFY":
        return result.get("public_status_label", "À confirmer pour le lot"), AMBER, AMBER_BG
    if key == "UNDEMONSTRATED":
        return result.get("public_status_label", "Non démontrée dans les documents"), SLATE, SOFT
    if key == "EXCLUDED":
        return result.get("public_status_label", "Non applicable / explicitement exclue"), MUTED, SOFT
    return _status_info.__wrapped__(result) if hasattr(_status_info, "__wrapped__") else (_clean(result.get("statut")) or "À examiner", SLATE, SOFT)


def _v4_status(result: Dict[str, Any]) -> Tuple[str, colors.Color, colors.Color]:
    key = result.get("public_status_key")
    if key == "CONFIRMED":
        return result.get("public_status_label", "Confirmée pour le lot"), V4_GREEN, V4_GREEN_BG
    if key == "CLARIFY":
        return result.get("public_status_label", "À confirmer pour le lot"), V4_AMBER, V4_AMBER_BG
    if key == "UNDEMONSTRATED":
        return result.get("public_status_label", "Non démontrée dans les documents"), V4_SLATE, V4_PALE
    if key == "EXCLUDED":
        return result.get("public_status_label", "Non applicable / explicitement exclue"), V4_MUTED, colors.HexColor("#F0F3F5")
    return _clean(result.get("statut")) or "À vérifier", V4_SLATE, V4_PALE


def _capacity_info(block_id: str, scores: Dict[str, Dict[str, Any]], result: Dict[str, Any] | None = None) -> Tuple[str, Optional[int], colors.Color, colors.Color]:
    result = result or {}
    label = result.get("public_capacity_label")
    pct = result.get("public_capacity_pct")
    state = result.get("public_capacity_state")
    if label:
        if state == "DEMONSTRATED":
            return label.replace(f" - {pct} %", ""), pct, GREEN, GREEN_BG
        if state == "PARTIAL":
            return label.replace(f" - {pct} %", ""), pct, AMBER, AMBER_BG
        if state == "NOT_DEMONSTRATED":
            return label.replace(f" - {pct} %", ""), pct, RED, RED_BG
        return label, None, MUTED, SOFT
    score = scores.get(block_id) or {}
    raw = score.get("pct")
    if isinstance(raw, (int, float)):
        pct = int(round(raw))
        if pct >= 70:
            return "Capacité en place", pct, GREEN, GREEN_BG
        if pct >= 40:
            return "Capacité partielle", pct, AMBER, AMBER_BG
        return "Capacité à développer", pct, RED, RED_BG
    return "Évaluation non disponible pour ce bloc", None, MUTED, SOFT


def _priority(result: Dict[str, Any], block_id: str, scores: Dict[str, Dict[str, Any]]) -> Tuple[int, str]:
    label = result.get("public_priority")
    group = result.get("public_group_order", 99)
    if label:
        return int(group), label
    return 99, "À examiner"


def _plain_title(block_id: str, result: Dict[str, Any]) -> str:
    excel_title = _clean((result.get("bloc") or {}).get("titre_bloc"))
    return _plain_language(result.get("public_title") or excel_title or _FALLBACK_PLAIN_TITLES.get(block_id) or block_id)


def _relevant_items(results: Dict[str, Dict[str, Any]], scores: Dict[str, Dict[str, Any]]) -> List[Tuple[str, Dict[str, Any]]]:
    items = [(bid, res) for bid, res in results.items() if res.get("public_status_key") != "EXCLUDED"]
    return sorted(items, key=lambda item: (item[1].get("public_group_order", 99), _model_block_sort_key(item[0])))


def _actions_for(block_id: str, result: Dict[str, Any], scores: Dict[str, Dict[str, Any]]) -> List[str]:
    actions = result.get("public_actions") or []
    return [_plain_language(action) for action in actions]


def _proposal_for(block_id: str, proposals: Dict[str, Dict[str, Any]], result: Dict[str, Any] | None = None) -> Tuple[str, str]:
    if result and result.get("public_response_text"):
        return result.get("public_response_label", "Réponse proposée"), _plain_language(result.get("public_response_text"))
    proposal = proposals.get(block_id) or {}
    return _plain_language(proposal.get("niveau") or "Réponse proposée"), _plain_language(proposal.get("texte") or proposal.get("avertissement") or "")


def _situation_summary(block_id: str, scores: Dict[str, Dict[str, Any]], result: Dict[str, Any] | None = None) -> Tuple[str, str]:
    if result and result.get("public_capacity_strength"):
        return _plain_language(result.get("public_capacity_strength")), _plain_language(result.get("public_capacity_gap"))
    return "Évaluation non disponible pour ce bloc.", "Compléter le paramétrage de l'auto-évaluation."


def _v4_source(result: Dict[str, Any]) -> str:
    evidence = result.get("public_evidence") or []
    if evidence:
        return " - ".join(f"{item.get('source')} p.{item.get('page')}" for item in evidence)
    return "Aucune preuve ciblée et recevable retenue"


def _v5_action_block_data(
    block_id: str,
    result: Dict[str, Any],
    scores: Dict[str, Dict[str, Any]],
    proposals: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    block = result.get("bloc") or {}
    status_label, _, _ = _v4_status(result)
    _, priority = _priority(result, block_id, scores)
    capacity_label, pct, _, _ = _capacity_info(block_id, scores, result)
    strength, gap = _situation_summary(block_id, scores, result)
    demand = _plain_language(
        result.get("public_demand")
        or block.get("lecture_entreprise")
        or block.get("description_convention")
        or result.get("conclusion")
        or "L'attente exacte n'est pas suffisamment explicite dans les extraits retenus."
    )
    proposal_label, proposal = _proposal_for(block_id, proposals, result)
    return {
        "id": block_id,
        "title": _plain_title(block_id, result),
        "status": status_label,
        "priority": priority,
        "capacity_label": capacity_label,
        "pct": pct,
        "strength": strength,
        "gap": gap,
        "demand": demand,
        "proposal_label": proposal_label,
        "proposal": proposal or "Aucune formulation ferme ne doit être intégrée en l'état.",
        "actions": _actions_for(block_id, result, scores),
        "action_items": result.get("public_action_items") or [{"text": action, "origin": ""} for action in _actions_for(block_id, result, scores)],
        "question": _plain_language(result.get("public_question") or ""),
    }


# Enveloppes : enrichissement unique avant chaque génération.
_generate_action_plan_pdf_v6 = generate_action_plan_pdf
_generate_action_plan_html_v6 = generate_action_plan_html


def generate_action_plan_pdf(out, results, scores, proposals, meta):
    _ensure_public_model(results, scores, proposals, lot=str(meta.get("lot", "")), axes_config=meta.get("axes_config"))
    return _generate_action_plan_pdf_v6(out, results, scores, proposals, meta)


def generate_action_plan_html(out, results, scores, proposals, meta):
    _ensure_public_model(results, scores, proposals, lot=str(meta.get("lot", "")), axes_config=meta.get("axes_config"))
    return _generate_action_plan_html_v6(out, results, scores, proposals, meta)


def _situation_summary(block_id: str, scores: Dict[str, Dict[str, Any]], result: Dict[str, Any] | None = None) -> Tuple[str, str]:
    if result and result.get("public_capacity_strength"):
        return _plain_language(result.get("public_capacity_strength")), _plain_language(result.get("public_capacity_gap"))
    score = scores.get(block_id) or {}
    responses = score.get("reponses") or []
    if not responses:
        return "Évaluation non disponible pour ce bloc.", "Compléter le paramétrage de l'auto-évaluation avant de conclure sur la capacité."
    strengths, gaps = [], []
    for response in responses:
        value = response.get("reponse_val")
        label = _clean(response.get("reponse_label") or response.get("question"))
        if value == 2 and label:
            strengths.append(label)
        elif value in (0, 1) and label:
            gaps.append(label)
    return (
        _plain_language(strengths[0]) if strengths else "Aucun point fort suffisamment démontré dans les réponses fournies.",
        _plain_language(gaps[0]) if gaps else "Aucun écart majeur identifié dans les réponses fournies.",
    )


def _canonical_proposals(results: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {
        bid: {
            "texte": result.get("public_response_text", ""),
            "niveau": result.get("public_response_label", ""),
            "conditions_ok": result.get("public_response_label", "").startswith("Proposition à intégrer"),
            "avertissement": result.get("public_response_text", ""),
        }
        for bid, result in results.items()
    }


def generate_action_plan_pdf(out, results, scores, proposals, meta):
    _ensure_public_model(results, scores, proposals, lot=str(meta.get("lot", "")), axes_config=meta.get("axes_config"))
    return _generate_action_plan_pdf_v6(out, results, scores, _canonical_proposals(results), meta)


def generate_action_plan_html(out, results, scores, proposals, meta):
    _ensure_public_model(results, scores, proposals, lot=str(meta.get("lot", "")), axes_config=meta.get("axes_config"))
    return _generate_action_plan_html_v6(out, results, scores, _canonical_proposals(results), meta)

# Nettoyage typographique final après expansion des termes techniques.
_plain_language_v7_base = _plain_language

def _plain_language(value: Any) -> str:
    text = _plain_language_v7_base(value)
    replacements = {
        "plateforme plateforme commune de dépôt (CDE)": "plateforme commune de dépôt (CDE)",
        "plateformes plateforme commune de dépôt (CDE)": "plateformes communes de dépôt (CDE)",
        "la plateforme la plateforme commune de dépôt (CDE)": "la plateforme commune de dépôt (CDE)",
        "(exploitation et maintenance)": "(GMAO / AIM)",
        "nous confirmeons": "nous confirmons",
        "outil de outil de gestion de maintenance (GMAO)": "outil de gestion de maintenance (GMAO)",
        "outil de gestion de maintenance (outil de gestion de maintenance (GMAO))": "outil de gestion de maintenance (GMAO)",
        "modèle de données outil de gestion de maintenance (GMAO)": "modèle de données GMAO",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return re.sub(r"\\s+", " ", text).strip()


def _plain_title(block_id: str, result: Dict[str, Any]) -> str:
    return _pdf_safe(result.get("public_title") or _clean((result.get("bloc") or {}).get("titre_bloc")) or _FALLBACK_PLAIN_TITLES.get(block_id) or block_id)


def generate_glossary_html(out: Path, meta: Dict[str, Any]) -> None:
    """3e livrable : glossaire autonome des termes BIM détectés dans les
    documents analysés, le questionnaire d'auto-évaluation et les textes de
    l'outil (même liste, déjà filtrée en amont, que les sections glossaire du
    Dashboard et du Plan d'Actions -- jamais le référentiel complet)."""
    if meta.get("logo_data_uri"):
        brand_html = f'<img src="{html.escape(str(meta.get("logo_data_uri")))}" alt="Logo">'
    else:
        brand_html = html.escape(str(meta.get("marque") or meta.get("entreprise") or "Entreprise"))
    entries = sorted((meta.get("glossary_entries") or []), key=lambda e: str(e.get("term", "")).casefold())
    rows = "".join(
        f'<div class="gloss-row"><strong>{html.escape(str(e.get("term","")))}</strong>'
        f'<span>{html.escape(str(e.get("definition","")))}'
        + (f'<small>{html.escape(str(e.get("source")))}</small>' if e.get("source") else "")
        + '</span></div>'
        for e in entries if e.get("term") and e.get("definition")
    )
    doc = (
        '<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
        f'<title>Glossaire BIM - {html.escape(str(meta.get("lot","")))}</title><style>'
        ':root{--navy:#0B2733;--magenta:#A1003D;--magenta-light:#9C5B45;--blue:#1F5D8C;--line:#D7E1E6;--muted:#5B7280;--pale2:#EEF3F5;--slate:#33495A}'
        '*{box-sizing:border-box}body{margin:0;background:#f2f5f7;color:#1A2B33;font:14px/1.5 Calibri,Carlito,"Segoe UI",Arial,sans-serif}'
        '.hero{background:var(--navy);border-bottom:5px solid var(--magenta);padding:27px max(25px,calc((100vw - 1100px)/2));display:flex;align-items:center;gap:22px}'
        '.brand{font-size:29px;font-weight:800;color:var(--magenta-light);border-right:1px solid rgba(156,91,69,.45);padding-right:22px}'
        '.brand img{max-height:40px;max-width:160px;display:block;object-fit:contain}'
        '.hero h1{margin:0;color:var(--magenta-light);font-size:29px;line-height:1;font-weight:700}'
        '.hero p{margin:7px 0 0;color:#E7D3CB;font-weight:600}'
        '.container{max-width:1100px;margin:0 auto;padding:22px}'
        '.count-bar{background:#fff;border:1px solid var(--line);border-radius:6px;padding:12px 16px;margin-bottom:14px;color:var(--muted);font-weight:700}'
        '.count-bar b{color:var(--navy)}'
        '#search{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:6px;margin-bottom:14px;font:inherit}'
        '.gloss-table{border:1px solid var(--line);border-radius:6px;overflow:hidden;background:#fff}'
        '.gloss-row{display:grid;grid-template-columns:200px 1fr;border-bottom:1px solid var(--line)}'
        '.gloss-row:last-child{border-bottom:0}'
        '.gloss-row strong{padding:12px;background:var(--pale2);color:var(--blue);font-size:13px}'
        '.gloss-row span{padding:12px;font-size:13px;color:var(--slate)}'
        '.gloss-row small{display:block;margin-top:5px;color:var(--muted);font-style:italic}'
        '@media(max-width:700px){.gloss-row{grid-template-columns:1fr}.gloss-row strong{border-bottom:1px solid var(--line)}}'
        '</style></head><body>'
        f'<header class="hero"><div class="brand">{brand_html}</div><div><h1>GLOSSAIRE DES TERMES BIM</h1>'
        f'<p>{html.escape(str(meta.get("entreprise","-")))} - Lot {html.escape(str(meta.get("lot","-")))} - Termes détectés dans les documents, le questionnaire et l\'outil</p></div></header>'
        f'<div class="container"><div class="count-bar"><b>{len(entries)}</b> terme(s) technique(s) BIM détecté(s) pour ce dossier.</div>'
        '<input id="search" placeholder="Rechercher un terme ou une définition…">'
        f'<div class="gloss-table" id="glossTable">{rows}</div></div>'
        '<script>document.getElementById("search").addEventListener("input",e=>{const q=e.target.value.toLowerCase();'
        'document.querySelectorAll("#glossTable .gloss-row").forEach(r=>{r.hidden=!r.innerText.toLowerCase().includes(q);});});</script>'
        '</body></html>'
    )
    Path(out).write_text(doc, encoding="utf-8")


def generate_glossary_pdf(out: Path, meta: Dict[str, Any]) -> None:
    """Version PDF du glossaire autonome (même contenu que generate_glossary_html)."""
    entries = sorted((meta.get("glossary_entries") or []), key=lambda e: str(e.get("term", "")).casefold())
    doc = _v4_doc(
        out, "Glossaire des termes BIM détectés dans ce dossier", pagesize=A4, landscape_mode=False,
        marque=meta.get("marque") or meta.get("entreprise") or "Entreprise",
        logo_bytes=meta.get("logo_bytes"),
    )
    story: List[Any] = [
        _para("GLOSSAIRE DES TERMES BIM", V4["h1"]),
        _para(
            f'{_esc(str(meta.get("entreprise","-")))} - Lot {_esc(str(meta.get("lot","-")))} - '
            f'{len(entries)} terme(s) détecté(s) dans les documents, le questionnaire et l\'outil.',
            V4["small"],
        ),
        Spacer(1, 8),
    ]
    _gloss_term_style = ParagraphStyle("gloss_term_std", parent=V4["table"], fontName=FONT_BOLD, fontSize=8.2, leading=10.2, textColor=V4_MAGENTA)
    _gloss_def_style = ParagraphStyle("gloss_def_std", parent=V4["table"], fontSize=8.2, leading=10.4)
    for entry in entries:
        term = str(entry.get("term", ""))
        definition = str(entry.get("definition", ""))
        if not term or not definition:
            continue
        story.append(_rich_para(_esc(term), _gloss_term_style))
        story.append(_rich_para(_esc(definition), _gloss_def_style))
        story.append(Spacer(1, 4))
    doc.build(story)

