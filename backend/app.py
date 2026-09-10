from __future__ import annotations

import json
import os
import re
import hashlib
import secrets
import shutil
import subprocess
import tempfile
import threading
import sys
import time
import uuid
import unicodedata
from pathlib import Path
from typing import Any

from bim_model import block_sort_key
from content_rules import glossary_entries_from_dataframe, normalize_questions_dataframe
from outil_bim_v1 import (
    _sub_ctx, lire_pdf, _extraire_infos_document,
    evaluer_lisibilite_document, load_params, load_messages,
    configure_text_filters, _portee_lot_correspond, _portee_lot_tokens,
)

import pandas as pd
import fitz
from flask import Flask, jsonify, request, send_file, url_for, redirect
from flask_cors import CORS
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
TOOL_PATH = BASE_DIR / "outil_bim_v1.py"
PARAM_TEMPLATE = BASE_DIR / "Parametrage_Outil_BIM_V1.xlsx"
JOBS_DIR = BASE_DIR / "runtime" / "jobs"
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "80"))
JOB_TTL_HOURS = int(os.environ.get("JOB_TTL_HOURS", "168"))
ANALYSIS_TIMEOUT_SECONDS = int(os.environ.get("ANALYSIS_TIMEOUT_SECONDS", "600"))

JOBS_DIR.mkdir(parents=True, exist_ok=True)
STATUS_LOCK = threading.Lock()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.json.ensure_ascii = False
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

_origins = [x.strip() for x in os.environ.get(
    "ALLOWED_ORIGINS",
    "http://127.0.0.1:5173,http://localhost:5173"
).split(",") if x.strip()]
CORS(app, resources={r"/api/*": {"origins": _origins, "allow_headers": ["Content-Type", "X-BIM-Client-Key"]}})


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _norm_simple(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def _param_tokens(params: dict[str, object], key: str) -> set[str]:
    raw = params.get(key, []) or []
    if isinstance(raw, str):
        values = [x.strip() for x in raw.split(";")]
    else:
        values = [str(x).strip() for x in raw]
    return {_norm_simple(x) for x in values if _norm_simple(x)}


def _significant_tokens(value: Any, params: dict[str, object]) -> set[str]:
    min_len = max(2, int(float(params.get("precheck_projet_token_longueur_min", 4) or 4)))
    stop = _param_tokens(params, "precheck_projet_mots_vides")
    return {
        token for token in _norm_simple(value).split()
        if len(token) >= min_len and token not in stop and not token.isdigit()
    }


def _project_name_match(saisi: str, detecte: str, params: dict[str, object]) -> bool:
    ns = _norm_simple(saisi)
    nd = _norm_simple(detecte)
    if not ns or not nd:
        return False
    if ns in nd or nd in ns:
        return True
    a = _significant_tokens(saisi, params)
    b = _significant_tokens(detecte, params)
    if not a or not b:
        return False
    common = len(a & b)
    ratio = common / max(1, min(len(a), len(b)))
    threshold = float(params.get("precheck_projet_nom_overlap_min", 0.5) or 0.5)
    return common >= 1 and ratio >= threshold


def _raw_first_pages(path: Path, max_pages: int) -> list[str]:
    pages: list[str] = []
    with fitz.open(str(path)) as pdf:
        for idx in range(min(len(pdf), max(1, max_pages))):
            pages.append(pdf[idx].get_text() or "")
    return pages


def _project_label(infos: dict[str, object]) -> str:
    heading = clean(infos.get("nom_projet"))
    operation = clean(infos.get("operation") or infos.get("titre"))
    if heading and operation:
        if _norm_simple(heading) in _norm_simple(operation) or _norm_simple(operation) in _norm_simple(heading):
            return operation if len(operation) >= len(heading) else heading
        return f"{heading} : {operation}"
    return heading or operation


def _document_project_similarity(ref_pages: list[str], other_pages: list[str], params: dict[str, object]) -> dict[str, object]:
    min_common = max(1, int(float(params.get("precheck_projet_similarity_common_tokens_min", 2) or 2)))
    min_overlap = float(params.get("precheck_projet_similarity_overlap_min", 0.30) or 0.30)
    best = {"overlap": 0.0, "jaccard": 0.0, "common": 0, "ref_page": 0, "other_page": 0}
    usable = False
    for i, ref_text in enumerate(ref_pages, start=1):
        a = _significant_tokens(ref_text, params)
        if not a:
            continue
        for j, other_text in enumerate(other_pages, start=1):
            b = _significant_tokens(other_text, params)
            if not b:
                continue
            usable = True
            common = len(a & b)
            overlap = common / max(1, min(len(a), len(b)))
            jaccard = common / max(1, len(a | b))
            candidate = (overlap, common, jaccard)
            current = (float(best["overlap"]), int(best["common"]), float(best["jaccard"]))
            if candidate > current:
                best = {"overlap": overlap, "jaccard": jaccard, "common": common, "ref_page": i, "other_page": j}
    best["usable"] = usable
    best["match"] = bool(usable and int(best["common"]) >= min_common and float(best["overlap"]) >= min_overlap)
    best["score_pct"] = round(float(best["overlap"]) * 100)
    return best


DOCUMENT_ROLE_LABELS = {
    "CONVENTION_BIM": "Convention BIM",
    "CCTP": "CCTP",
    "CCAP": "CCAP",
}
EXPECTED_ROLE_BY_KEY = {
    "convention": "CONVENTION_BIM",
    "cctp": "CCTP",
    "ccap": "CCAP",
}


def _role_scores(page_texts: list[str], filename: str = "") -> dict[str, int]:
    """Classe le type documentaire à partir du titre/nom, sans règle projet-spécifique.

    Les mentions situées sur la première page et dans le nom de fichier ont un
    poids supérieur aux références documentaires trouvées ensuite dans un
    sommaire ou une liste de pièces associées.
    """
    scores = {"CONVENTION_BIM": 0, "CCTP": 0, "CCAP": 0}
    fn = _norm_simple(Path(filename or "").stem)
    first = _norm_simple(page_texts[0] if page_texts else "")
    second = _norm_simple(page_texts[1] if len(page_texts) > 1 else "")

    def add(role: str, cond: bool, weight: int) -> None:
        if cond:
            scores[role] += weight

    # Nom du fichier : indice fort, mais jamais suffisant à lui seul face à un
    # titre de première page contradictoire.
    add("CONVENTION_BIM", bool(re.search(r"(?:^| )convention(?: |.* )bim(?: |$)|(?:^| )bim(?: |.* )convention(?: |$)", fn)), 7)
    add("CCTP", bool(re.search(r"(?:^| )cctp(?: |$)", fn)), 7)
    add("CCAP", bool(re.search(r"(?:^| )ccap(?: |$)", fn)), 7)

    def scan(text: str, strong: int, weak: int) -> None:
        add("CONVENTION_BIM", "convention bim" in text, strong)
        add("CONVENTION_BIM", "convention" in text and "building information" in text, weak)
        add("CCTP", bool(re.search(r"(?:^| )cctp(?: |$)", text)), strong)
        add("CCTP", "cahier des clauses techniques particulier" in text, strong)
        add("CCAP", bool(re.search(r"(?:^| )ccap(?: |$)", text)), strong)
        add("CCAP", "cahier des clauses administratives particulier" in text, strong)

    scan(first, 11, 5)
    scan(second, 3, 2)
    return scores


def _document_role_check(key: str, page_texts: list[str], filename: str) -> dict[str, object]:
    expected = EXPECTED_ROLE_BY_KEY[key]
    scores = _role_scores(page_texts, filename)
    ranked = sorted(scores.items(), key=lambda kv: (kv[1], kv[0]), reverse=True)
    best_role, best_score = ranked[0]
    second_score = ranked[1][1]
    if best_score < 6 or best_score - second_score < 2:
        detected = "UNKNOWN"
    else:
        detected = best_role

    if detected == expected:
        coherence = "MATCH"
        detail = f"Type documentaire reconnu : {DOCUMENT_ROLE_LABELS[expected]}."
    elif detected == "UNKNOWN":
        coherence = "UNKNOWN"
        detail = f"Le document n'a pas pu être identifié de façon fiable comme {DOCUMENT_ROLE_LABELS[expected]}."
    else:
        coherence = "MISMATCH"
        detail = (
            f"Document chargé dans l'emplacement {DOCUMENT_ROLE_LABELS[expected]}, "
            f"mais il ressemble à un document de type {DOCUMENT_ROLE_LABELS.get(detected, detected)}."
        )
    return {
        "expected_role": expected,
        "expected_label": DOCUMENT_ROLE_LABELS[expected],
        "detected_role": detected,
        "detected_label": DOCUMENT_ROLE_LABELS.get(detected, "Type non identifié"),
        "coherence": coherence,
        "detail": detail,
        "scores": scores,
    }


def _lot_match_from_pages(lot: str, page_texts: list[str], min_tokens: object = None) -> dict[str, object]:
    lot_tokens = _portee_lot_tokens(lot)
    text_tokens = _portee_lot_tokens("\n".join(page_texts or []))
    if not lot_tokens or not text_tokens:
        return {"coherence": "UNKNOWN", "common": [], "lot_tokens": sorted(lot_tokens)}
    try:
        threshold = max(1, int(float(min_tokens or 1)))
    except (TypeError, ValueError):
        threshold = 1
    common = sorted(lot_tokens & text_tokens)
    return {
        "coherence": "MATCH" if len(common) >= threshold else "MISMATCH",
        "common": common,
        "lot_tokens": sorted(lot_tokens),
    }


def _sha256_path(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _precheck_fingerprint(projet: str, lot: str, docs: dict[str, dict[str, object] | None]) -> str:
    payload = {
        "projet": _norm_simple(projet),
        "lot": _norm_simple(lot),
        "documents": {
            key: {
                "sha256": clean((doc or {}).get("sha256")),
                "filename": clean((doc or {}).get("filename")),
            }
            for key, doc in sorted(docs.items())
        },
    }
    raw = json.dumps(payload, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _read_precheck_document(
    key: str,
    path: Path | None,
    filename: str,
    params: dict[str, object],
    metadata_rules: pd.DataFrame,
    max_pages: int,
) -> dict[str, object] | None:
    if path is None or not path.exists():
        return None
    if path.stat().st_size == 0 or path.read_bytes()[:5] != b"%PDF-":
        raise ValueError(f"{key.upper()} PDF invalide.")
    doc = lire_pdf(path)
    quality = evaluer_lisibilite_document(doc, params)
    page_texts = _raw_first_pages(path, max_pages)
    infos = _extraire_infos_document(page_texts, None, metadata_rules) if page_texts else {}
    role = _document_role_check(key, page_texts, filename)
    return {
        "key": key,
        "filename": clean(filename),
        "quality": clean(quality.get("statut", "OK")),
        "reading": {
            "statut": clean(quality.get("statut", "OK")),
            "pages_total": int(quality.get("pages_total", 0) or 0),
            "pages_lisibles": int(quality.get("pages_lisibles", 0) or 0),
            "pages_images_non_lisibles": [int(x) for x in (quality.get("pages_images_non_lisibles", []) or [])],
            "pages_images_non_lisibles_count": int(quality.get("pages_images_non_lisibles_count", 0) or 0),
            "caracteres_extraits": int(quality.get("caracteres_extraits", 0) or 0),
        },
        "pages": page_texts,
        "projet_detecte": _project_label(infos),
        "projet_court": clean(infos.get("nom_projet")),
        "operation": clean(infos.get("operation")),
        "lot_detecte": clean(infos.get("lot_cctp")) if key == "cctp" else "",
        "role": role,
        "sha256": _sha256_path(path),
    }


def _evaluate_precheck_paths(
    convention_path: Path,
    cctp_path: Path | None,
    ccap_path: Path | None,
    filenames: dict[str, str],
    lot: str,
    projet_saisi: str,
) -> dict[str, object]:
    params, pdf_messages = pdf_reading_config()
    metadata_rules = read_sheet("36_Extraction_Metadata", "champ")
    max_pages = max(1, int(float(params.get("precheck_projet_pages_max", 2) or 2)))
    docs = {
        "convention": _read_precheck_document("convention", convention_path, filenames.get("convention", ""), params, metadata_rules, max_pages),
        "cctp": _read_precheck_document("cctp", cctp_path, filenames.get("cctp", ""), params, metadata_rules, max_pages) if cctp_path else None,
        "ccap": _read_precheck_document("ccap", ccap_path, filenames.get("ccap", ""), params, metadata_rules, max_pages) if ccap_path else None,
    }
    convention = docs["convention"] or {}

    # 1) Type documentaire de chaque emplacement.
    role_checks = []
    blocking_errors: list[str] = []
    role_states = []
    for key, label in (("convention", "Convention BIM"), ("cctp", "CCTP"), ("ccap", "CCAP")):
        current = docs.get(key)
        if not current:
            continue
        role = dict(current.get("role") or {})
        role_states.append(clean(role.get("coherence")))
        role_checks.append({
            "document": label,
            "filename": clean(current.get("filename")),
            "coherence": clean(role.get("coherence")) or "UNKNOWN",
            "expected_label": clean(role.get("expected_label")),
            "detected_label": clean(role.get("detected_label")),
            "detail": clean(role.get("detail")),
        })
        if role.get("coherence") == "MISMATCH":
            blocking_errors.append(clean(role.get("detail")))
    if any(state == "MISMATCH" for state in role_states):
        types_coherence = "MISMATCH"
    elif any(state == "UNKNOWN" for state in role_states):
        types_coherence = "UNKNOWN"
    else:
        types_coherence = "MATCH"

    # 1 bis) Lisibilité PDF : expose les pages scannées avant analyse.
    reading_checks = []
    reading_states = []
    for key, label in (("convention", "Convention BIM"), ("cctp", "CCTP"), ("ccap", "CCAP")):
        current = docs.get(key)
        if not current:
            continue
        reading = dict(current.get("reading") or {})
        statut = clean(reading.get("statut") or current.get("quality") or "OK").upper()
        pages_total = int(reading.get("pages_total", 0) or 0)
        pages_lisibles = int(reading.get("pages_lisibles", 0) or 0)
        scan_pages = [int(x) for x in (reading.get("pages_images_non_lisibles") or [])]
        if statut == "OK":
            coherence = "MATCH"
            detail = f"Lecture automatique complète : {pages_lisibles}/{pages_total} page(s) exploitable(s)." if pages_total else "Lecture automatique exploitable."
        elif statut == "PARTIEL":
            coherence = "UNKNOWN"
            if scan_pages:
                shown = scan_pages[:20]
                suffix = f" (+{len(scan_pages)-len(shown)} autre(s))" if len(scan_pages) > len(shown) else ""
                detail = f"Lecture partielle : {pages_lisibles}/{pages_total} page(s) exploitables ; page(s) image sans texte exploitable : {', '.join(map(str, shown))}{suffix}."
            else:
                detail = f"Lecture partielle : {pages_lisibles}/{pages_total} page(s) exploitables. Certaines exigences peuvent ne pas être détectées automatiquement."
        else:
            coherence = "MISMATCH" if key == "convention" else "UNKNOWN"
            if key == "convention":
                detail = pdf_messages.get(
                    "pdf_convention_non_exploitable_error",
                    "La Convention BIM n'est pas suffisamment exploitable automatiquement.",
                )
                blocking_errors.append(detail)
            elif key == "cctp":
                detail = "CCTP non exploitable automatiquement : son contenu et le lot ne peuvent pas être vérifiés de façon fiable. L'absence de preuve ne doit pas être interprétée comme une absence d'exigence."
            else:
                detail = "CCAP non exploitable automatiquement : l'ordre de priorité contractuel et les mentions administratives ne peuvent pas être vérifiés de façon fiable."
        reading_states.append(coherence)
        reading_checks.append({
            "document": label,
            "filename": clean(current.get("filename")),
            "statut": statut,
            "coherence": coherence,
            "pages_total": pages_total,
            "pages_lisibles": pages_lisibles,
            "pages_images_non_lisibles": scan_pages,
            "detail": detail,
        })
    reading_coherence = "MATCH" if reading_states and all(x == "MATCH" for x in reading_states) else ("UNKNOWN" if reading_states else "MATCH")

    # 2) Projet saisi vs Convention.
    projet_detecte = clean(convention.get("projet_detecte"))
    if not projet_detecte:
        projet_coherence = "UNKNOWN"
        projet_warning = "Le nom du projet n'a pas été détecté automatiquement dans les premières pages de la Convention BIM. Vérifiez et confirmez le nom saisi avant de poursuivre."
    elif _project_name_match(projet_saisi, projet_detecte, params):
        projet_coherence = "MATCH"
        projet_warning = ""
    else:
        projet_coherence = "MISMATCH"
        projet_warning = "Le nom du projet saisi ne semble pas correspondre au nom détecté dans la Convention BIM. Vérifiez le dossier avant de poursuivre."

    # 3) Lot saisi vs CCTP. Le texte des premières pages sert de repli si la
    # règle de métadonnées ne renvoie pas un libellé de lot exploitable.
    cctp = docs.get("cctp")
    lot_detecte = clean((cctp or {}).get("lot_detecte"))
    lot_match_detail = ""
    if not cctp:
        lot_coherence = "NO_CCTP"
        lot_warning = "Aucun CCTP n'est chargé : la cohérence entre le lot saisi et le CCTP ne peut pas être vérifiée."
    elif clean((cctp.get("role") or {}).get("coherence")) == "MISMATCH":
        lot_coherence = "MISMATCH"
        lot_warning = "Le fichier placé comme CCTP n'a pas été reconnu comme un CCTP."
    elif clean(cctp.get("quality")).upper() == "NON_EXPLOITABLE":
        lot_coherence = "UNKNOWN"
        lot_warning = "Le CCTP n'est pas suffisamment lisible automatiquement pour vérifier son lot."
    elif lot_detecte:
        lot_match = _portee_lot_correspond(lot_detecte, lot, params.get("cctp_lot_coherence_min_tokens", 1))
        lot_coherence = "MATCH" if lot_match else "MISMATCH"
        lot_warning = "" if lot_match else f"Le lot saisi « {lot} » ne correspond pas au lot détecté dans le CCTP : « {lot_detecte} »."
    else:
        fallback_lot = _lot_match_from_pages(lot, list(cctp.get("pages") or []), params.get("cctp_lot_coherence_min_tokens", 1))
        lot_coherence = clean(fallback_lot.get("coherence")) or "UNKNOWN"
        common = fallback_lot.get("common") or []
        lot_match_detail = ", ".join(common)
        if lot_coherence == "MATCH":
            lot_warning = ""
        elif lot_coherence == "MISMATCH":
            lot_warning = f"Le lot saisi « {lot} » n'a pas été retrouvé dans les premières pages du CCTP chargé. Vérifiez que le CCTP correspond bien au lot analysé."
        else:
            lot_warning = "Le libellé du lot n'a pas pu être vérifié automatiquement dans le CCTP ; vérifiez le fichier avant de poursuivre."

    # 4) Projet de la Convention vs autres documents.
    project_document_checks = []
    for key, label in (("cctp", "CCTP"), ("ccap", "CCAP")):
        other = docs.get(key)
        if not other:
            continue
        if clean(convention.get("quality")).upper() == "NON_EXPLOITABLE" or clean(other.get("quality")).upper() == "NON_EXPLOITABLE":
            project_document_checks.append({
                "document": label, "filename": clean(other.get("filename")),
                "coherence": "UNKNOWN", "score_pct": None,
                "detail": "Comparaison automatique impossible : au moins un des deux PDF n'est pas suffisamment lisible.",
            })
            continue
        sim = _document_project_similarity(list(convention.get("pages") or []), list(other.get("pages") or []), params)
        coherence = "MATCH" if sim.get("match") else ("MISMATCH" if sim.get("usable") else "UNKNOWN")
        other_project = clean(other.get("projet_detecte"))
        if projet_detecte and other_project:
            coherence = "MATCH" if _project_name_match(projet_detecte, other_project, params) else "MISMATCH"
        if coherence == "MATCH":
            detail = f"Ressemblance de projet trouvée entre les premières pages (score {sim.get('score_pct', 0)} %)."
        elif coherence == "MISMATCH":
            detail = f"Le {label} semble appartenir à un autre projet que la Convention BIM (meilleur score {sim.get('score_pct', 0)} % sur {max_pages} page(s))."
        else:
            detail = "La cohérence de projet n'a pas pu être vérifiée automatiquement."
        project_document_checks.append({
            "document": label, "filename": clean(other.get("filename")),
            "coherence": coherence, "score_pct": sim.get("score_pct"),
            "ref_page": sim.get("ref_page"), "document_page": sim.get("other_page"),
            "projet_detecte": other_project, "detail": detail,
        })

    if any(x.get("coherence") == "MISMATCH" for x in project_document_checks):
        documents_coherence = "MISMATCH"
        documents_warning = "Au moins un document chargé semble appartenir à un autre projet. Vérifiez ou déchargez le fichier concerné avant l'analyse."
    elif any(x.get("coherence") == "UNKNOWN" for x in project_document_checks):
        documents_coherence = "UNKNOWN"
        documents_warning = "La cohérence de tous les documents n'a pas pu être démontrée automatiquement. Vérifiez les fichiers avant l'analyse."
    else:
        documents_coherence = "MATCH"
        documents_warning = ""

    alert_states = [projet_coherence, lot_coherence, documents_coherence, types_coherence, reading_coherence]
    has_alert = any(state not in {"MATCH", ""} for state in alert_states)
    result = {
        "coherence": lot_coherence,
        "lot_saisi": lot,
        "lot_detecte": lot_detecte,
        "lot_match_detail": lot_match_detail,
        "warning": lot_warning,
        "projet_saisi": projet_saisi,
        "projet_detecte": projet_detecte,
        "projet_coherence": projet_coherence,
        "projet_warning": projet_warning,
        "types_coherence": types_coherence,
        "role_checks": role_checks,
        "reading_coherence": reading_coherence,
        "reading_checks": reading_checks,
        "documents_coherence": documents_coherence,
        "documents_warning": documents_warning,
        "document_checks": project_document_checks,
        "pages_comparees": max_pages,
        "blocking_errors": blocking_errors,
        "has_blocking_error": bool(blocking_errors),
        "has_alert": has_alert,
    }
    result["fingerprint"] = _precheck_fingerprint(projet_saisi, lot, docs)
    return result


CLIENT_KEY_HEADER = "X-BIM-Client-Key"
CLIENT_KEY_RE = re.compile(r"^[A-Za-z0-9_-]{24,160}$")
ACCESS_TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{24,200}$")


def _client_key() -> str:
    value = clean(request.headers.get(CLIENT_KEY_HEADER, ""))
    return value if CLIENT_KEY_RE.fullmatch(value) else ""


def _owner_hash(client_key: str) -> str:
    return hashlib.sha256(client_key.encode("utf-8")).hexdigest() if client_key else ""


def _new_access_token() -> str:
    return secrets.token_urlsafe(32)


def _protect_path(path: Path, directory: bool = False) -> None:
    try:
        path.chmod(0o700 if directory else 0o600)
    except OSError:
        pass


def require_client_key():
    key = _client_key()
    if not key:
        return None, (jsonify({"error": "Session locale non identifiée. Actualisez l’outil puis relancez l’opération."}), 401)
    return key, None


def _access_from_request() -> str:
    value = clean(request.args.get("access", ""))
    return value if ACCESS_TOKEN_RE.fullmatch(value) else ""


def authorize_job(job_id: str, allow_access_token: bool = True) -> tuple[Path | None, dict[str, Any], tuple | None]:
    try:
        job_dir = job_file(job_id, "job.json").parent
    except FileNotFoundError:
        return None, {}, (jsonify({"error": "Analyse introuvable."}), 404)
    metadata = read_json_file(job_dir / "job.json")
    if not metadata:
        return None, {}, (jsonify({"error": "Analyse introuvable."}), 404)
    client_key = _client_key()
    if client_key and clean(metadata.get("owner_hash")) == _owner_hash(client_key):
        return job_dir, metadata, None
    if allow_access_token:
        access = _access_from_request()
        expected = clean(metadata.get("access_token"))
        if access and expected and secrets.compare_digest(access, expected):
            return job_dir, metadata, None
    return None, {}, (jsonify({"error": "Accès refusé à cette analyse."}), 403)


def _job_url(endpoint: str, job_id: str, metadata: dict[str, Any], **values: Any) -> str:
    token = clean(metadata.get("access_token"))
    if token:
        values["access"] = token
    return url_for(endpoint, job_id=job_id, _external=True, **values)


@app.after_request
def security_headers(response):
    path = request.path or ""
    if path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store, private"
        response.headers["Pragma"] = "no-cache"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["X-Content-Type-Options"] = "nosniff"
    return response

_protect_path(JOBS_DIR, directory=True)

def read_sheet(sheet: str, key_col: str) -> pd.DataFrame:
    raw = pd.read_excel(PARAM_TEMPLATE, sheet_name=sheet, header=None)
    header_index = None
    for i, row in raw.iterrows():
        if key_col in [clean(v) for v in row.values]:
            header_index = i
            break
    if header_index is None:
        raise RuntimeError(f"Colonne {key_col!r} introuvable dans {sheet!r}")
    headers = [clean(v) for v in raw.iloc[header_index].values]
    df = raw.iloc[header_index + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").fillna("")
    df = df[[c for c in df.columns if c and c.lower() != "nan"]]
    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.strip().str.lower().ne("non")]
    return df


def pdf_reading_config() -> tuple[dict[str, object], dict[str, str]]:
    """Charge les seuils/messages et initialise les filtres de lecture PDF.

    `lire_pdf()` s'appuie sur les paramètres globaux du moteur pour distinguer
    les véritables en-têtes répétés du contenu métier. Dans l'application web,
    cette initialisation doit avoir lieu avant toute lecture rapide (détection
    projet, contrôle d'une Convention ou ajout d'une pièce complémentaire).
    """
    params = load_params(read_sheet("05_Parametres_Moteur", "cle"))
    messages = load_messages(read_sheet("12_Messages_Moteur", "cle"))
    configure_text_filters(read_sheet("15_Filtres_Texte", "id_filtre"), params)
    return params, messages


def message_map() -> dict[str, str]:
    df = read_sheet("12_Messages_Moteur", "cle")
    return {clean(row.get("cle", "")): clean(row.get("message", "")) for _, row in df.iterrows() if clean(row.get("cle", ""))}


def questionnaire_payload() -> dict[str, Any]:
    blocks = read_sheet("10_Blocs", "id_bloc")
    questions = normalize_questions_dataframe(read_sheet("50_Questions", "id_bloc"))
    glossary = read_sheet("40_Glossaire", "terme")
    messages = message_map()
    titles = {clean(r["id_bloc"]): clean(r.get("titre_bloc", "")) for _, r in blocks.iterrows()}
    groups: list[dict[str, Any]] = []
    grouped = list(questions.groupby(questions["id_bloc"].astype(str).str.strip(), sort=False))
    grouped.sort(key=lambda item: block_sort_key(item[0]))
    for block_id, group in grouped:
        items = []
        for _, row in group.iterrows():
            # Le questionnaire est affiché avant toute analyse de documents :
            # aucune valeur détectée n'existe encore pour {plateforme}/{cout}/...
            # On applique donc _sub_ctx avec un contexte vide, pour que les
            # formulations neutres par défaut s'affichent au lieu des accolades
            # brutes ("{plateforme}"). Les mêmes textes seront re-substitués
            # avec les vraies valeurs détectées une fois l'analyse lancée.
            items.append({
                "id": clean(row.get("id_question", "")),
                "question": _sub_ctx(clean(row.get("question", "")), {}, messages),
                "options": [
                    {"value": 0, "label": _sub_ctx(clean(row.get("indice_0", "")), {}, messages)},
                    {"value": 1, "label": _sub_ctx(clean(row.get("indice_1", "")), {}, messages)},
                    {"value": 2, "label": _sub_ctx(clean(row.get("indice_2", "")), {}, messages)},
                ],
                "weight": int(float(row.get("poids", 1) or 1)),
            })
        groups.append({"block_id": block_id, "title": titles.get(block_id, block_id), "questions": items})
    return {
        "groups": groups,
        "count": sum(len(g["questions"]) for g in groups),
        "glossary": glossary_entries_from_dataframe(glossary),
    }


def cleanup_old_jobs() -> None:
    cutoff = time.time() - JOB_TTL_HOURS * 3600
    for path in JOBS_DIR.iterdir():
        try:
            if path.is_dir() and path.stat().st_mtime < cutoff:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            continue


def save_pdf(file_storage, destination: Path, required: bool = False) -> bool:
    if not file_storage or not file_storage.filename:
        if required:
            raise ValueError("La convention BIM est obligatoire.")
        return False
    filename = secure_filename(file_storage.filename)
    if not filename.lower().endswith(".pdf"):
        raise ValueError(f"{filename or 'Le fichier'} doit être un PDF.")
    file_storage.save(destination)
    _protect_path(destination)
    if destination.stat().st_size == 0:
        raise ValueError(f"Le fichier {filename} est vide.")
    if destination.read_bytes()[:5] != b"%PDF-":
        raise ValueError(f"Le fichier {filename} n'est pas un PDF valide.")
    return True


def save_logo(file_storage, destination: Path) -> bool:
    """Enregistre un logo entreprise (PNG/JPEG) optionnel, inséré ensuite dans
    les en-têtes des livrables à la place du sigle texte codé en dur."""
    if not file_storage or not file_storage.filename:
        return False
    filename = secure_filename(file_storage.filename)
    if not filename.lower().endswith((".png", ".jpg", ".jpeg")):
        raise ValueError(f"{filename or 'Le logo'} doit être une image PNG ou JPEG.")
    file_storage.save(destination)
    _protect_path(destination)
    if destination.stat().st_size == 0:
        destination.unlink(missing_ok=True)
        return False
    return True


def job_file(job_id: str, name: str) -> Path:
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise FileNotFoundError
    path = (JOBS_DIR / job_id / name).resolve()
    if JOBS_DIR.resolve() not in path.parents:
        raise FileNotFoundError
    return path


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/api/questions")
def questions_api():
    try:
        return jsonify(questionnaire_payload())
    except Exception as exc:
        return jsonify({"error": f"Impossible de lire le questionnaire : {exc}"}), 500


@app.post("/api/detect-project")
def detect_project_api():
    """Lecture rapide de la Convention BIM seule, pour pré-remplir le nom du
    projet dès qu'elle est sélectionnée -- l'utilisateur reste libre de le
    corriger avant de lancer l'analyse complète (le champ est obligatoire
    mais jamais figé)."""
    upload = request.files.get("convention")
    if not upload or not upload.filename:
        return jsonify({"error": "Aucune convention fournie."}), 400
    tmp_path = Path(tempfile.gettempdir()) / f"detect_{uuid.uuid4().hex}.pdf"
    try:
        upload.save(tmp_path)
        if tmp_path.stat().st_size == 0 or tmp_path.read_bytes()[:5] != b"%PDF-":
            return jsonify({"error": "Fichier PDF invalide."}), 400
        pdf_params, pdf_messages = pdf_reading_config()
        doc = lire_pdf(tmp_path)
        quality = evaluer_lisibilite_document(doc, pdf_params)
        if quality.get("statut") == "NON_EXPLOITABLE":
            return jsonify({"error": pdf_messages.get("pdf_convention_non_exploitable_error", "Document PDF insuffisamment exploitable automatiquement.")}), 400
        metadata_rules = read_sheet("36_Extraction_Metadata", "champ")
        max_pages = max(1, int(float(pdf_params.get("precheck_projet_pages_max", 2) or 2)))
        pages_texte = _raw_first_pages(tmp_path, max_pages)
        infos = _extraire_infos_document(pages_texte, None, metadata_rules)
        role_check = _document_role_check("convention", pages_texte, upload.filename or "")
        payload = {
            "projet": _project_label(infos),
            "projet_court": clean(infos.get("nom_projet")),
            "operation": clean(infos.get("operation")),
            "lecture_pdf": quality.get("statut", "OK"),
            "document_role_coherence": role_check.get("coherence"),
            "document_role_detected": role_check.get("detected_label"),
            "document_role_warning": "" if role_check.get("coherence") == "MATCH" else role_check.get("detail"),
        }
        if quality.get("statut") == "PARTIEL":
            payload["warning"] = pdf_messages.get("vigilance_pdf_partial_title", "Lecture documentaire partielle")
            scan_pages = [int(x) for x in (quality.get("pages_images_non_lisibles", []) or [])]
            if scan_pages:
                max_pages = int(float(pdf_params.get("pdf_lecture_pages_signalees_max", 20) or 20))
                shown = scan_pages[:max(1, max_pages)]
                pages_text = ", ".join(str(x) for x in shown)
                if len(scan_pages) > len(shown):
                    pages_text += f" (+{len(scan_pages) - len(shown)} autre(s))"
                template = pdf_messages.get("vigilance_pdf_partial_scan_text", "")
                if template:
                    try:
                        payload["warning_detail"] = template.format(
                            document=upload.filename or "Convention BIM",
                            pages_non_lisibles=pages_text,
                            nombre_pages_non_lisibles=len(scan_pages),
                            pages_lisibles=quality.get("pages_lisibles", 0),
                            pages_total=quality.get("pages_total", 0),
                        )
                    except Exception:
                        payload["warning_detail"] = payload["warning"]
        return jsonify(payload)
    except Exception as exc:
        return jsonify({"error": f"Lecture impossible : {exc}"}), 500
    finally:
        tmp_path.unlink(missing_ok=True)


@app.post("/api/precheck")
def precheck_api():
    """Contrôle préalable obligatoire avant analyse depuis l'interface web.

    DEV07 : contrôle le type des pièces, le projet, le lot et la cohérence entre
    documents. Il renvoie aussi une empreinte du dossier pour empêcher qu'une
    autre combinaison de fichiers soit analysée après validation.
    """
    client_key, auth_error = require_client_key()
    if auth_error:
        return auth_error

    lot = clean(request.form.get("lot"))
    projet_saisi = clean(request.form.get("projet"))
    uploads = {
        "convention": request.files.get("convention"),
        "cctp": request.files.get("cctp"),
        "ccap": request.files.get("ccap"),
    }
    if not uploads["convention"] or not getattr(uploads["convention"], "filename", ""):
        return jsonify({"error": "La Convention BIM est obligatoire pour la vérification avant analyse."}), 400
    if not lot:
        return jsonify({"error": "Le lot à analyser est obligatoire."}), 400
    if not projet_saisi:
        return jsonify({"error": "Le nom du projet est obligatoire."}), 400

    temp_paths: list[Path] = []
    try:
        paths: dict[str, Path | None] = {"convention": None, "cctp": None, "ccap": None}
        filenames: dict[str, str] = {"convention": "", "cctp": "", "ccap": ""}
        for key, upload in uploads.items():
            if not upload or not getattr(upload, "filename", ""):
                continue
            tmp_path = Path(tempfile.gettempdir()) / f"precheck_{key}_{uuid.uuid4().hex}.pdf"
            temp_paths.append(tmp_path)
            upload.save(tmp_path)
            paths[key] = tmp_path
            filenames[key] = clean(getattr(upload, "filename", ""))
        result = _evaluate_precheck_paths(
            paths["convention"], paths["cctp"], paths["ccap"],
            filenames, lot, projet_saisi,
        )
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Vérification documentaire impossible : {exc}"}), 500
    finally:
        for tmp_path in temp_paths:
            tmp_path.unlink(missing_ok=True)


@app.post("/api/analyse")
def analyse_api():
    client_key, auth_error = require_client_key()
    if auth_error:
        return auth_error
    cleanup_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    _protect_path(job_dir, directory=True)

    try:
        convention_upload = request.files.get("convention")
        cctp_upload = request.files.get("cctp")
        ccap_upload = request.files.get("ccap")
        convention_name = clean(getattr(convention_upload, "filename", "")) or "Convention BIM.pdf"
        cctp_name = clean(getattr(cctp_upload, "filename", "")) if cctp_upload else ""
        ccap_name = clean(getattr(ccap_upload, "filename", "")) if ccap_upload else ""

        convention_path = job_dir / "Convention_BIM.pdf"
        cctp_path = job_dir / "CCTP_Lot.pdf"
        ccap_path = job_dir / "CCAP.pdf"
        save_pdf(convention_upload, convention_path, required=True)
        has_cctp = save_pdf(cctp_upload, cctp_path, required=False)
        has_ccap = save_pdf(ccap_upload, ccap_path, required=False)

        # Protection immédiate : la Convention est la pièce d'entrée obligatoire.
        # Si elle n'a pas de couche texte exploitable, ne pas lancer un traitement
        # qui pourrait produire de faux "non démontré". Les seuils et le message
        # viennent du classeur de paramétrage.
        pdf_params, pdf_messages = pdf_reading_config()
        convention_quality = evaluer_lisibilite_document(lire_pdf(convention_path), pdf_params)
        if convention_quality.get("statut") == "NON_EXPLOITABLE":
            raise ValueError(pdf_messages.get(
                "pdf_convention_non_exploitable_error",
                "La Convention BIM n'est pas suffisamment exploitable automatiquement.",
            ))

        entreprise = clean(request.form.get("entreprise")) or "[Nom de l'entreprise]"
        nom_projet = clean(request.form.get("projet"))
        if not nom_projet:
            raise ValueError("Le nom du projet est obligatoire.")
        lot = clean(request.form.get("lot")) or "Lot"

        # DEV07 - le backend refait le précontrôle sur les fichiers réellement
        # reçus. Une interface web ne peut donc plus lancer silencieusement une
        # analyse avec un dossier différent de celui qui a été validé.
        precheck = _evaluate_precheck_paths(
            convention_path, cctp_path if has_cctp else None, ccap_path if has_ccap else None,
            {"convention": convention_name, "cctp": cctp_name if has_cctp else "", "ccap": ccap_name if has_ccap else ""},
            lot, nom_projet,
        )
        browser_request = bool(clean(request.headers.get("Origin")))
        ack = str(request.form.get("precheck_ack", "false")).lower() in {"1", "true", "yes", "oui"}
        override = str(request.form.get("precheck_override", "false")).lower() in {"1", "true", "yes", "oui"}
        supplied_fingerprint = clean(request.form.get("precheck_fingerprint"))
        if precheck.get("has_blocking_error"):
            raise ValueError("Contrôle documentaire bloquant : " + " ".join(precheck.get("blocking_errors") or []))
        if browser_request:
            if not ack:
                raise ValueError("La validation préalable du projet, du lot et des documents est obligatoire avant l'analyse.")
            if not supplied_fingerprint or supplied_fingerprint != clean(precheck.get("fingerprint")):
                raise ValueError("Le dossier a changé depuis sa validation préalable. Relancez le contrôle avant l'analyse.")
            if precheck.get("has_alert") and not override:
                raise ValueError("Une incohérence ou une incertitude documentaire doit être confirmée explicitement avant l'analyse.")

        marque = entreprise  # "Entreprise" sert directement de sigle en en-tête (champ dédié retiré)
        adresse = clean(request.form.get("adresse"))
        email = clean(request.form.get("email"))
        analyste = clean(request.form.get("analyste"))
        sans_eval = str(request.form.get("sans_eval", "false")).lower() in {"1", "true", "yes", "oui"}

        logo_upload = request.files.get("logo")
        logo_ext = Path(secure_filename(getattr(logo_upload, "filename", "") or "")).suffix.lower()
        if logo_ext not in {".png", ".jpg", ".jpeg"}:
            logo_ext = ".png"
        logo_path = job_dir / f"logo{logo_ext}"
        has_logo = save_logo(logo_upload, logo_path)

        answers: dict[str, int] = {}
        question_total = questionnaire_payload()["count"]
        if not sans_eval:
            try:
                raw_answers = json.loads(request.form.get("answers", "{}"))
            except json.JSONDecodeError as exc:
                raise ValueError("Les réponses au questionnaire ne sont pas valides.") from exc
            if not isinstance(raw_answers, dict):
                raise ValueError("Les réponses doivent être un objet JSON.")
            expected_ids = {
                q["id"] for group in questionnaire_payload()["groups"] for q in group["questions"]
            }
            missing = sorted(expected_ids - set(raw_answers))
            if missing:
                raise ValueError(f"Réponses manquantes : {', '.join(missing[:6])}" + ("…" if len(missing) > 6 else ""))
            for qid, value in raw_answers.items():
                try:
                    value = int(value)
                except (TypeError, ValueError) as exc:
                    raise ValueError(f"Réponse invalide pour {qid}.") from exc
                if value not in (0, 1, 2):
                    raise ValueError(f"Réponse invalide pour {qid} : 0, 1 ou 2 attendu.")
                answers[str(qid)] = value

        answers_path = job_dir / "reponses.json"
        answers_path.write_text(json.dumps(answers, ensure_ascii=False, indent=2), encoding="utf-8")
        _protect_path(answers_path)
        now = int(time.time())
        metadata = {
            "job_id": job_id,
            "project": Path(convention_name).stem,
            "nom_projet": nom_projet,
            "entreprise": entreprise,
            "lot": lot,
            "marque": marque,
            "adresse": adresse,
            "email": email,
            "analyste": analyste,
            "sans_eval": sans_eval,
            "question_total": question_total,
            "question_answered": 0 if sans_eval else len(answers),
            "documents": [name for name in [convention_name, cctp_name if has_cctp else "", ccap_name if has_ccap else ""] if name],
            "job_type": "analysis",
            "revision": 1,
            "complements": [],
            "created_at": now,
            "status": "queued",
            "owner_hash": _owner_hash(client_key),
            "access_token": _new_access_token(),
        }
        write_json_atomic(job_dir / "job.json", metadata)
        update_job_status(
            job_dir,
            status="queued",
            progress=5,
            stage="Analyse enregistrée",
            detail="Les fichiers sont prêts pour le traitement.",
            started_at=now,
        )

        options = {
            "entreprise": entreprise,
            "nom_projet": nom_projet,
            "lot": lot,
            "marque": marque,
            "adresse": adresse,
            "email": email,
            "analyste": analyste,
            "logo_path": str(logo_path) if has_logo else "",
            "sans_eval": sans_eval,
            "has_cctp": has_cctp,
            "has_ccap": has_ccap,
            "documents": [name for name in [convention_name, cctp_name if has_cctp else "", ccap_name if has_ccap else ""] if name],
        }
        worker = threading.Thread(
            target=run_analysis_job,
            args=(job_id, options),
            daemon=True,
            name=f"bim-analysis-{job_id[:8]}",
        )
        worker.start()

        return jsonify({
            "job_id": job_id,
            "status_url": _job_url("job_status", job_id, metadata),
        }), 202
    except ValueError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(exc)}), 500


def _analysis_options_from_metadata(metadata: dict[str, Any], job_dir: Path) -> dict[str, Any]:
    logo_path = ""
    for candidate in sorted(job_dir.glob("logo.*")):
        if candidate.is_file():
            logo_path = str(candidate)
            break
    return {
        "entreprise": clean(metadata.get("entreprise")) or "[Nom de l'entreprise]",
        "nom_projet": clean(metadata.get("nom_projet")),
        "lot": clean(metadata.get("lot")) or "Lot",
        "marque": clean(metadata.get("marque")) or clean(metadata.get("entreprise")),
        "adresse": clean(metadata.get("adresse")),
        "email": clean(metadata.get("email")),
        "analyste": clean(metadata.get("analyste")),
        "logo_path": logo_path,
        "sans_eval": bool(metadata.get("sans_eval")),
        "has_cctp": (job_dir / "CCTP_Lot.pdf").exists(),
        "has_ccap": (job_dir / "CCAP.pdf").exists(),
        "documents": list(metadata.get("documents") or []),
    }


def _copy_analysis_inputs(parent_dir: Path, child_dir: Path) -> None:
    for name in ("Convention_BIM.pdf", "CCTP_Lot.pdf", "CCAP.pdf", "reponses.json"):
        src = parent_dir / name
        if src.exists():
            shutil.copy2(src, child_dir / name)
            _protect_path(child_dir / name)
    for src in parent_dir.glob("logo.*"):
        if src.is_file():
            shutil.copy2(src, child_dir / src.name)
            _protect_path(child_dir / src.name)
    if (parent_dir / "complements.json").exists():
        shutil.copy2(parent_dir / "complements.json", child_dir / "complements.json")
        for item in read_json_file(parent_dir / "complements.json", {"items": []}).get("items", []):
            stored = clean(item.get("stored_name")) if isinstance(item, dict) else ""
            if stored and (parent_dir / stored).exists():
                shutil.copy2(parent_dir / stored, child_dir / stored)
                _protect_path(child_dir / stored)


@app.post("/api/jobs/<job_id>/add-document")
def add_document_to_analysis(job_id: str):
    """Crée une nouvelle révision cohérente de l'analyse à partir du Dashboard.

    L'ancienne analyse n'est jamais modifiée en place : les documents, réponses et
    paramètres sont clonés dans un nouveau job. Le Dashboard et le Plan de la
    nouvelle révision sont donc générés ensemble à partir du même corpus.
    """
    parent_dir, parent_meta, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert parent_dir is not None
    parent_status = read_json_file(parent_dir / "status.json")
    if parent_meta.get("job_type") != "analysis":
        return jsonify({"error": "L'ajout de pièce est réservé aux analyses documentaires."}), 400
    if clean(parent_meta.get("superseded_by")):
        current_id = clean(parent_meta.get("superseded_by"))
        return jsonify({
            "error": "Cette analyse a déjà été remplacée par une version plus récente.",
            "result_url": _job_url("result_job", current_id, read_json_file(JOBS_DIR / current_id / "job.json")),
        }), 409
    if clean(parent_meta.get("pending_revision")):
        pending_id = clean(parent_meta.get("pending_revision"))
        return jsonify({
            "error": "Une mise à jour de cette analyse est déjà en cours.",
            "status_url": _job_url("job_status", pending_id, read_json_file(JOBS_DIR / pending_id / "job.json")),
        }), 409
    if parent_status.get("status") != "completed":
        return jsonify({"error": "L'analyse doit être terminée avant d'ajouter une pièce."}), 409

    upload = request.files.get("document")
    if not upload or not upload.filename:
        return jsonify({"error": "Aucune pièce complémentaire fournie."}), 400

    temp_path = Path(tempfile.gettempdir()) / f"bim_supp_{uuid.uuid4().hex}.pdf"
    try:
        save_pdf(upload, temp_path, required=True)
        pdf_params, pdf_messages = pdf_reading_config()
        quality = evaluer_lisibilite_document(lire_pdf(temp_path), pdf_params)
        if quality.get("statut") == "NON_EXPLOITABLE":
            return jsonify({
                "error": pdf_messages.get(
                    "supplement_unreadable_error",
                    "La pièce complémentaire n'est pas suffisamment exploitable automatiquement.",
                )
            }), 400

        new_job_id = uuid.uuid4().hex
        child_dir = JOBS_DIR / new_job_id
        child_dir.mkdir(parents=True, exist_ok=False)
        _protect_path(child_dir, directory=True)
        _copy_analysis_inputs(parent_dir, child_dir)

        existing_manifest = read_json_file(child_dir / "complements.json", {"items": []})
        items = list(existing_manifest.get("items") or [])
        index = len(items) + 1
        stored_name = f"Complement_{index:02d}.pdf"
        shutil.copy2(temp_path, child_dir / stored_name)
        _protect_path(child_dir / stored_name)

        original_name = clean(getattr(upload, "filename", "")) or stored_name
        reference = clean(request.form.get("reference"))
        origin_source = clean(request.form.get("source"))
        origin_page = clean(request.form.get("page"))
        origin_excerpt = clean(request.form.get("excerpt"))
        related_blocks = [clean(x) for x in str(request.form.get("blocks", "")).split(",") if clean(x)]
        items.append({
            "stored_name": stored_name,
            "display_name": original_name,
            "reference": reference,
            "origin_source": origin_source,
            "origin_page": origin_page,
            "origin_excerpt": origin_excerpt,
            "related_blocks": related_blocks,
            "added_at": int(time.time()),
            "reading_status": clean(quality.get("statut")) or "OK",
        })
        write_json_atomic(child_dir / "complements.json", {"items": items})

        now = int(time.time())
        child_meta = dict(parent_meta)
        for key in ("status", "completed_at", "duration_seconds", "error", "superseded_by", "pending_revision"):
            child_meta.pop(key, None)
        child_meta.update({
            "job_id": new_job_id,
            "parent_job_id": job_id,
            "revision": int(parent_meta.get("revision") or 1) + 1,
            "created_at": now,
            "status": "queued",
            "complements": items,
            "documents": list(parent_meta.get("documents") or []) + [original_name],
            "trigger_reference": reference,
            "trigger_blocks": related_blocks,
        })
        write_json_atomic(child_dir / "job.json", child_meta)
        update_job_status(
            child_dir,
            status="queued",
            progress=5,
            stage="Pièce complémentaire ajoutée",
            detail="Une nouvelle version cohérente de l'analyse va être générée.",
            started_at=now,
        )
        parent_meta["pending_revision"] = new_job_id
        write_json_atomic(parent_dir / "job.json", parent_meta)

        worker = threading.Thread(
            target=run_analysis_job,
            args=(new_job_id, _analysis_options_from_metadata(child_meta, child_dir)),
            daemon=True,
            name=f"bim-reanalysis-{new_job_id[:8]}",
        )
        worker.start()
        return jsonify({
            "job_id": new_job_id,
            "revision": child_meta["revision"],
            "status_url": _job_url("job_status", new_job_id, child_meta),
        }), 202
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        temp_path.unlink(missing_ok=True)


@app.post("/api/autoevaluation")
def autoevaluation_api():
    client_key, auth_error = require_client_key()
    if auth_error:
        return auth_error
    cleanup_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    _protect_path(job_dir, directory=True)
    try:
        entreprise = clean(request.form.get("entreprise")) or "[Nom de l'entreprise]"
        lot = clean(request.form.get("lot")) or "Activité / lot non renseigné"
        adresse = clean(request.form.get("adresse"))
        email = clean(request.form.get("email"))
        analyste = clean(request.form.get("analyste"))
        logo_upload = request.files.get("logo")
        logo_ext = Path(secure_filename(getattr(logo_upload, "filename", "") or "")).suffix.lower()
        if logo_ext not in {".png", ".jpg", ".jpeg"}:
            logo_ext = ".png"
        logo_path = job_dir / f"logo{logo_ext}"
        has_logo = save_logo(logo_upload, logo_path)
        try:
            raw_answers = json.loads(request.form.get("answers", "{}"))
        except json.JSONDecodeError as exc:
            raise ValueError("Les réponses au questionnaire ne sont pas valides.") from exc
        if not isinstance(raw_answers, dict):
            raise ValueError("Les réponses doivent être un objet JSON.")
        qpayload = questionnaire_payload()
        expected_ids = {q["id"] for group in qpayload["groups"] for q in group["questions"]}
        missing = sorted(expected_ids - set(raw_answers))
        if missing:
            raise ValueError(f"Réponses manquantes : {', '.join(missing[:6])}" + ("…" if len(missing) > 6 else ""))
        answers: dict[str, int] = {}
        for qid in expected_ids:
            try:
                value = int(raw_answers[qid])
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Réponse invalide pour {qid}.") from exc
            if value not in (0, 1, 2):
                raise ValueError(f"Réponse invalide pour {qid} : 0, 1 ou 2 attendu.")
            answers[qid] = value
        (job_dir / "reponses.json").write_text(json.dumps(answers, ensure_ascii=False, indent=2), encoding="utf-8")
        _protect_path(job_dir / "reponses.json")
        now = int(time.time())
        messages = message_map()
        metadata = {
            "job_id": job_id, "job_type": "autoevaluation", "project": clean(messages.get("autoeval_job_project_label")),
            "entreprise": entreprise, "lot": lot, "adresse": adresse, "email": email, "analyste": analyste,
            "question_total": len(expected_ids), "question_answered": len(answers), "documents": [],
            "created_at": now, "status": "queued",
            "owner_hash": _owner_hash(client_key), "access_token": _new_access_token(),
        }
        write_json_atomic(job_dir / "job.json", metadata)
        update_job_status(job_dir, status="queued", progress=5, stage="Auto-évaluation enregistrée", detail="Les réponses sont prêtes pour le traitement.", started_at=now)
        options = {"entreprise": entreprise, "lot": lot, "adresse": adresse, "email": email, "analyste": analyste,
                   "logo": str(logo_path) if has_logo else ""}
        threading.Thread(target=run_autoevaluation_job, args=(job_id, options), daemon=True, name=f"bim-autoeval-{job_id[:8]}").start()
        return jsonify({"job_id": job_id, "status_url": _job_url("job_status", job_id, metadata)}), 202
    except ValueError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(exc)}), 500


def write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    """Écrit un JSON de façon atomique (fichier .tmp puis renommage). Sur
    Windows, ce renommage peut échouer temporairement (WinError 5/32) si un
    antivirus ou un client de synchronisation (OneDrive, Dropbox...) a le
    fichier ouvert au même instant -- notamment quand le dossier de travail
    est lui-même synchronisé (ex. chemin sous "OneDrive\\..."). On retente
    quelques fois avant d'abandonner, plutôt que de faire échouer l'analyse."""
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    last_error: Exception | None = None
    for attempt in range(8):
        try:
            temp.replace(path)
            _protect_path(path)
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            time.sleep(0.15 * (attempt + 1))
    # Dernier recours : écriture directe (non atomique, mais évite de perdre
    # la progression si le renommage reste bloqué par un tiers).
    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        _protect_path(path)
        temp.unlink(missing_ok=True)
    except OSError:
        if last_error:
            raise last_error


def read_json_file(path: Path, default: dict[str, Any] | None = None) -> dict[str, Any]:
    """Lit un JSON en tolérant les verrouillages/écritures transitoires.

    Sous Windows avec OneDrive/antivirus, un très court verrouillage peut arriver
    au moment précis où le navigateur interroge status.json. Plusieurs lectures
    rapprochées évitent de transformer cet incident transitoire en faux 404.
    """
    for attempt in range(6):
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else (default or {})
        except (OSError, json.JSONDecodeError):
            if attempt < 5:
                time.sleep(0.04 * (attempt + 1))
    return default or {}


def update_job_status(job_dir: Path, **changes: Any) -> dict[str, Any]:
    with STATUS_LOCK:
        path = job_dir / "status.json"
        data = read_json_file(path)
        data.update(changes)
        data["updated_at"] = int(time.time())
        write_json_atomic(path, data)
        return data


def _progress_from_line(line: str, current: int) -> tuple[int, str, str]:
    text = clean(line)
    if "Analyse pour" in text:
        return max(current, 24), "Préparation du dossier", text
    if "[1/3]" in text:
        return max(current, 34), "Analyse des exigences BIM", "Lecture de la convention et du CCTP par bloc."
    if re.match(r"^B\d+[A-Z]?\s*[→-]", text):
        return min(max(current + 1, 36), 54), "Analyse des exigences BIM", text
    if "[2/3]" in text:
        return max(current, 58), "Croisement avec l’auto-évaluation", "Calcul des capacités et détection des contradictions."
    if "contradiction" in text.lower():
        return max(current, 66), "Contrôle de cohérence", text
    if "[3/3]" in text:
        return max(current, 72), "Préparation des livrables HTML", "Construction du dashboard et du plan d’action."
    if "Plan d'actions HTML généré" in text or "Plan d’actions HTML généré" in text:
        return max(current, 85), "Plan d’action généré", "Le plan d’action interactif est prêt."
    if "Dashboard HTML généré" in text:
        return max(current, 94), "Dashboard généré", "Vérification finale des deux livrables."
    return current, "", ""


def run_analysis_job(job_id: str, options: dict[str, Any]) -> None:
    job_dir = JOBS_DIR / job_id
    started = time.time()
    process: subprocess.Popen[str] | None = None
    timer: threading.Timer | None = None
    try:
        update_job_status(
            job_dir,
            status="running",
            progress=12,
            stage="Préparation technique",
            detail="Création du fichier de paramètres et contrôle des entrées.",
        )
        param_path = job_dir / "Parametrage_Outil_BIM_V1.xlsx"
        shutil.copy2(PARAM_TEMPLATE, param_path)
        output_html = job_dir / "Rapport_BIM.html"
        cmd = [
            sys.executable, "-u", str(TOOL_PATH),
            "--param", str(param_path),
            "--convention", str(job_dir / "Convention_BIM.pdf"),
            "--lot", options["lot"],
            "--entreprise", options["entreprise"],
            "--marque", options["marque"],
            "--out", str(output_html),
            "--format", "html",
            "--web-only",
        ]
        if options.get("has_cctp"):
            cmd.extend(["--cctp", str(job_dir / "CCTP_Lot.pdf")])
        if options.get("has_ccap"):
            cmd.extend(["--ccap", str(job_dir / "CCAP.pdf")])
        if options.get("logo_path"):
            cmd.extend(["--logo", options["logo_path"]])
        if options.get("adresse"):
            cmd.extend(["--adresse", options["adresse"]])
        if options.get("email"):
            cmd.extend(["--email", options["email"]])
        if options.get("analyste"):
            cmd.extend(["--analyste", options["analyste"]])
        if options.get("nom_projet"):
            cmd.extend(["--nom-projet", options["nom_projet"]])
        if options.get("documents"):
            cmd.extend(["--documents", ", ".join(options["documents"])])
        complements_manifest = job_dir / "complements.json"
        if complements_manifest.exists():
            cmd.extend(["--complements-json", str(complements_manifest)])
        if options.get("sans_eval"):
            cmd.append("--sans_eval")
        else:
            cmd.extend(["--reponses-json", str(job_dir / "reponses.json")])

        env = os.environ.copy()
        env.setdefault("MPLCONFIGDIR", str(job_dir / ".matplotlib"))
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUNBUFFERED"] = "1"

        update_job_status(
            job_dir,
            progress=20,
            stage="Démarrage du moteur d’analyse",
            detail="Le traitement des documents commence.",
        )
        log_path = job_dir / "analyse.log"
        with log_path.open("w", encoding="utf-8") as log_file:
            process = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            timed_out = {"value": False}

            def stop_process() -> None:
                timed_out["value"] = True
                if process and process.poll() is None:
                    process.kill()

            timer = threading.Timer(ANALYSIS_TIMEOUT_SECONDS, stop_process)
            timer.daemon = True
            timer.start()
            current = 20
            assert process.stdout is not None
            for line in process.stdout:
                log_file.write(line)
                log_file.flush()
                progress_value, stage, detail = _progress_from_line(line, current)
                if progress_value != current:
                    current = progress_value
                    update_job_status(job_dir, progress=current, stage=stage, detail=detail)
            return_code = process.wait()
            timer.cancel()
            timer = None
            if timed_out["value"]:
                raise TimeoutError("L’analyse a dépassé la durée maximale autorisée.")
            if return_code != 0:
                tail = "\n".join(log_path.read_text(encoding="utf-8", errors="replace").splitlines()[-25:])
                raise RuntimeError(f"Le moteur d’analyse s’est arrêté.\n{tail}")

        dashboard_html = job_dir / "Rapport_BIM.html"
        action_html = job_dir / "Plan_Actions_Offre_BIM.html"
        missing = [label for label, path in {
            "dashboard HTML": dashboard_html,
            "plan d’action HTML": action_html,
        }.items() if not path.exists()]
        if missing:
            raise RuntimeError("Livrables non générés : " + ", ".join(missing))

        completed_at = int(time.time())
        duration = max(1, round(time.time() - started))
        metadata = read_json_file(job_dir / "job.json")
        metadata.update({
            "status": "completed",
            "completed_at": completed_at,
            "duration_seconds": duration,
            "deliverables": ["dashboard", "actions", "glossary"],
        })
        write_json_atomic(job_dir / "job.json", metadata)
        for produced in job_dir.iterdir():
            if produced.is_file():
                _protect_path(produced)
        parent_job_id = clean(metadata.get("parent_job_id"))
        if parent_job_id:
            try:
                parent_dir = job_file(parent_job_id, "job.json").parent
                parent_meta = read_json_file(parent_dir / "job.json")
                if parent_meta:
                    parent_meta["superseded_by"] = job_id
                    parent_meta.pop("pending_revision", None)
                    write_json_atomic(parent_dir / "job.json", parent_meta)
            except FileNotFoundError:
                pass
        update_job_status(
            job_dir,
            status="completed",
            progress=100,
            stage="Analyse terminée",
            detail="Le dashboard et le plan d’action sont disponibles.",
            completed_at=completed_at,
            duration_seconds=duration,
        )
    except Exception as exc:
        if timer:
            timer.cancel()
        if process and process.poll() is None:
            process.kill()
        completed_at = int(time.time())
        metadata = read_json_file(job_dir / "job.json")
        metadata.update({
            "status": "error",
            "completed_at": completed_at,
            "duration_seconds": max(1, round(time.time() - started)),
            "error": str(exc),
        })
        write_json_atomic(job_dir / "job.json", metadata)
        parent_job_id = clean(metadata.get("parent_job_id"))
        if parent_job_id:
            try:
                parent_dir = job_file(parent_job_id, "job.json").parent
                parent_meta = read_json_file(parent_dir / "job.json")
                if clean(parent_meta.get("pending_revision")) == job_id:
                    parent_meta.pop("pending_revision", None)
                    write_json_atomic(parent_dir / "job.json", parent_meta)
            except FileNotFoundError:
                pass
        update_job_status(
            job_dir,
            status="error",
            progress=100,
            stage="Analyse interrompue",
            detail=str(exc),
            error=str(exc),
            completed_at=completed_at,
            duration_seconds=metadata["duration_seconds"],
        )


def run_autoevaluation_job(job_id: str, options: dict[str, Any]) -> None:
    job_dir = JOBS_DIR / job_id
    started = time.time()
    try:
        update_job_status(job_dir, status="running", progress=18, stage="Calcul de l’auto-évaluation", detail="Calcul des capacités déclarées par bloc.")
        param_path = job_dir / "Parametrage_Outil_BIM_V1.xlsx"
        shutil.copy2(PARAM_TEMPLATE, param_path)
        output_html = job_dir / "Auto_Evaluation_BIM.html"
        cmd = [sys.executable, "-u", str(TOOL_PATH), "--param", str(param_path), "--autoeval-only",
               "--reponses-json", str(job_dir / "reponses.json"), "--out", str(output_html),
               "--entreprise", options.get("entreprise", ""), "--lot", options.get("lot", "")]
        for flag, key in (("--adresse", "adresse"), ("--email", "email"), ("--analyste", "analyste"), ("--logo", "logo")):
            if options.get(key): cmd.extend([flag, options[key]])
        env = os.environ.copy(); env["PYTHONUTF8"] = "1"; env["PYTHONIOENCODING"] = "utf-8"
        log_path = job_dir / "autoevaluation.log"
        with log_path.open("w", encoding="utf-8") as log_file:
            result = subprocess.run(cmd, cwd=str(BASE_DIR), env=env, stdout=log_file, stderr=subprocess.STDOUT, text=True, timeout=ANALYSIS_TIMEOUT_SECONDS)
        if result.returncode != 0:
            tail = "\n".join(log_path.read_text(encoding="utf-8", errors="replace").splitlines()[-25:])
            raise RuntimeError(f"Le moteur d’auto-évaluation s’est arrêté.\n{tail}")
        out_pdf = job_dir / "Auto_Evaluation_BIM.pdf"
        if not output_html.exists() or not out_pdf.exists():
            raise RuntimeError("Le rapport d’auto-évaluation HTML/PDF n’a pas été généré.")
        completed_at = int(time.time()); duration = max(1, round(time.time() - started))
        metadata = read_json_file(job_dir / "job.json")
        metadata.update({"status": "completed", "completed_at": completed_at, "duration_seconds": duration, "deliverables": ["autoevaluation"]})
        write_json_atomic(job_dir / "job.json", metadata)
        update_job_status(job_dir, status="completed", progress=100, stage="Auto-évaluation terminée", detail="Le rapport HTML et PDF est disponible.", completed_at=completed_at, duration_seconds=duration)
    except Exception as exc:
        completed_at = int(time.time())
        metadata = read_json_file(job_dir / "job.json")
        metadata.update({"status": "error", "completed_at": completed_at, "duration_seconds": max(1, round(time.time() - started)), "error": str(exc)})
        write_json_atomic(job_dir / "job.json", metadata)
        update_job_status(job_dir, status="error", progress=100, stage="Auto-évaluation interrompue", detail=str(exc), error=str(exc), completed_at=completed_at, duration_seconds=metadata["duration_seconds"])


def _public_job_payload(job_id: str, status: dict[str, Any], metadata: dict[str, Any]) -> dict[str, Any]:
    # Ne jamais renvoyer job.json brut : adresse, email, analyste, jetons et
    # empreinte propriétaire restent exclusivement côté serveur.
    allowed = (
        "job_type", "project", "nom_projet", "entreprise", "lot", "documents",
        "revision", "created_at", "completed_at", "duration_seconds",
        "question_total", "question_answered",
    )
    payload = {key: metadata.get(key) for key in allowed if key in metadata}
    for key in ("status", "progress", "stage", "detail", "error", "updated_at"):
        if key in status:
            payload[key] = status.get(key)
    payload["job_id"] = job_id
    payload["status_url"] = _job_url("job_status", job_id, metadata)
    if status.get("status") == "completed":
        if metadata.get("job_type") == "autoevaluation":
            payload["autoevaluation_result_url"] = _job_url("autoevaluation_result_job", job_id, metadata)
            payload["autoevaluation_html_url"] = _job_url("download_job", job_id, metadata, kind="autoeval_html")
            payload["autoevaluation_pdf_url"] = _job_url("download_job", job_id, metadata, kind="autoeval_pdf")
        else:
            payload["result_url"] = _job_url("result_job", job_id, metadata)
            payload["actions_result_url"] = _job_url("actions_result_job", job_id, metadata)
            payload["glossary_result_url"] = _job_url("glossary_result_job", job_id, metadata)
            payload["result_download_url"] = _job_url("download_job", job_id, metadata, kind="html")
            payload["actions_download_url"] = _job_url("download_job", job_id, metadata, kind="actions_html")
            payload["actions_pdf_url"] = _job_url("download_job", job_id, metadata, kind="actions_pdf")
            payload["glossary_download_url"] = _job_url("download_job", job_id, metadata, kind="glossary_html")
            payload["glossary_pdf_url"] = _job_url("download_job", job_id, metadata, kind="glossary_pdf")
    return payload


@app.get("/api/jobs/<job_id>/status")
def job_status(job_id: str):
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    status = read_json_file(job_dir / "status.json")
    if not status:
        return jsonify({"error": "État de l’analyse temporairement indisponible.", "transient": True}), 503
    return jsonify(_public_job_payload(job_id, status, metadata))


def _owned_revision_chain(job_id: str, owner_hash: str) -> set[str]:
    pending = [job_id]
    found: set[str] = set()
    while pending:
        current = pending.pop()
        if current in found or not re.fullmatch(r"[a-f0-9]{32}", current):
            continue
        meta = read_json_file(JOBS_DIR / current / "job.json")
        if not meta or clean(meta.get("owner_hash")) != owner_hash:
            continue
        found.add(current)
        for key in ("parent_job_id", "superseded_by", "pending_revision"):
            linked = clean(meta.get(key))
            if linked and linked not in found:
                pending.append(linked)
    # Sécurité supplémentaire : retrouver d'éventuelles révisions enfants dont
    # le parent a été supprimé/masqué sans que le pointeur inverse soit complet.
    changed = True
    while changed:
        changed = False
        for d in JOBS_DIR.iterdir():
            if not d.is_dir() or d.name in found or not re.fullmatch(r"[a-f0-9]{32}", d.name):
                continue
            meta = read_json_file(d / "job.json")
            if clean(meta.get("owner_hash")) != owner_hash:
                continue
            if clean(meta.get("parent_job_id")) in found:
                found.add(d.name); changed = True
    return found


@app.get("/api/history")
def history_api():
    client_key, auth_error = require_client_key()
    if auth_error:
        return auth_error
    owner = _owner_hash(client_key)
    cleanup_old_jobs()
    items: list[dict[str, Any]] = []
    for job_dir in JOBS_DIR.iterdir():
        if not job_dir.is_dir() or not re.fullmatch(r"[a-f0-9]{32}", job_dir.name):
            continue
        metadata = read_json_file(job_dir / "job.json")
        if not metadata or clean(metadata.get("owner_hash")) != owner:
            continue
        status = read_json_file(job_dir / "status.json")
        if clean(metadata.get("superseded_by")):
            continue
        items.append(_public_job_payload(job_dir.name, status, metadata))
    items.sort(key=lambda item: int(item.get("created_at", 0) or 0), reverse=True)
    return jsonify({"items": items[:50]})


@app.delete("/api/history/<job_id>")
def delete_history_job(job_id: str):
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=False)
    if auth_error:
        return auth_error
    assert job_dir is not None
    owner = clean(metadata.get("owner_hash"))
    chain = _owned_revision_chain(job_id, owner)
    for linked_id in chain:
        shutil.rmtree(JOBS_DIR / linked_id, ignore_errors=True)
    return jsonify({"deleted": True, "deleted_revisions": len(chain)})


@app.get("/api/jobs/<job_id>/result")
def result_job(job_id: str):
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    path = job_dir / "Rapport_BIM.html"
    newer = clean(metadata.get("superseded_by"))
    if newer:
        return redirect(_job_url("result_job", newer, read_json_file(JOBS_DIR / newer / "job.json")))
    if not path.exists():
        return jsonify({"error": "Dashboard HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/actions")
def actions_result_job(job_id: str):
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    path = job_dir / "Plan_Actions_Offre_BIM.html"
    newer = clean(metadata.get("superseded_by"))
    if newer:
        return redirect(_job_url("actions_result_job", newer, read_json_file(JOBS_DIR / newer / "job.json")))
    if not path.exists():
        return jsonify({"error": "Plan d’actions HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/glossary")
def glossary_result_job(job_id: str):
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    path = job_dir / "Glossaire_BIM.html"
    newer = clean(metadata.get("superseded_by"))
    if newer:
        return redirect(_job_url("glossary_result_job", newer, read_json_file(JOBS_DIR / newer / "job.json")))
    if not path.exists():
        return jsonify({"error": "Glossaire HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/autoevaluation")
def autoevaluation_result_job(job_id: str):
    job_dir, _metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    path = job_dir / "Auto_Evaluation_BIM.html"
    if not path.exists():
        return jsonify({"error": "Rapport d’auto-évaluation HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/download/<kind>")
def download_job(job_id: str, kind: str):
    # Conservé uniquement pour l’export des deux HTML et le diagnostic du log.
    names = {
        "html": ("Rapport_BIM.html", "text/html"),
        "actions_html": ("Plan_Actions_Offre_BIM.html", "text/html"),
        "actions_pdf": ("Plan_Actions_Offre_BIM.pdf", "application/pdf"),
        "glossary_html": ("Glossaire_BIM.html", "text/html"),
        "glossary_pdf": ("Glossaire_BIM.pdf", "application/pdf"),
        "autoeval_html": ("Auto_Evaluation_BIM.html", "text/html"),
        "autoeval_pdf": ("Auto_Evaluation_BIM.pdf", "application/pdf"),
    }
    if kind not in names:
        return jsonify({"error": "Ce livrable n’est plus généré."}), 404
    name, mimetype = names[kind]
    job_dir, metadata, auth_error = authorize_job(job_id, allow_access_token=True)
    if auth_error:
        return auth_error
    assert job_dir is not None
    path = job_dir / name
    if not path.exists():
        return jsonify({"error": "Fichier introuvable."}), 404
    newer = clean(metadata.get("superseded_by"))
    if newer and metadata.get("job_type") == "analysis":
        return redirect(_job_url("download_job", newer, read_json_file(JOBS_DIR / newer / "job.json"), kind=kind))
    project = secure_filename(clean(metadata.get("project")) or "Projet")
    lot = secure_filename(clean(metadata.get("lot")) or "Lot")
    download_names = {
        "html": f"Dashboard_BIM_{project}_{lot}.html",
        "actions_html": f"Plan_Actions_BIM_{project}_{lot}.html",
        "actions_pdf": f"Plan_Actions_BIM_{project}_{lot}.pdf",
        "glossary_html": f"Glossaire_BIM_{project}_{lot}.html",
        "glossary_pdf": f"Glossaire_BIM_{project}_{lot}.pdf",
        "autoeval_html": f"Auto_Evaluation_BIM_{project}_{lot}.html",
        "autoeval_pdf": f"Auto_Evaluation_BIM_{project}_{lot}.pdf",
    }
    return send_file(path, mimetype=mimetype, as_attachment=True, download_name=download_names[kind])


@app.errorhandler(413)
def file_too_large(_):
    return jsonify({"error": f"Fichiers trop volumineux. Limite totale : {MAX_UPLOAD_MB} Mo."}), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
