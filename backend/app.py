from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import threading
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from bim_model import block_sort_key
from content_rules import glossary_entries_from_dataframe, normalize_questions_dataframe

import pandas as pd
from flask import Flask, jsonify, request, send_file, url_for
from flask_cors import CORS
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
TOOL_PATH = BASE_DIR / "outil_bim_v1.py"
PARAM_TEMPLATE = BASE_DIR / "Parametrage_Outil_BIM_V1.xlsx"
JOBS_DIR = BASE_DIR / "runtime" / "jobs"
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "80"))
JOB_TTL_HOURS = int(os.environ.get("JOB_TTL_HOURS", "720"))
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
CORS(app, resources={r"/api/*": {"origins": _origins}})


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


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


def questionnaire_payload() -> dict[str, Any]:
    blocks = read_sheet("10_Blocs", "id_bloc")
    questions = normalize_questions_dataframe(read_sheet("50_Questions", "id_bloc"))
    glossary = read_sheet("40_Glossaire", "terme")
    titles = {clean(r["id_bloc"]): clean(r.get("titre_bloc", "")) for _, r in blocks.iterrows()}
    groups: list[dict[str, Any]] = []
    grouped = list(questions.groupby(questions["id_bloc"].astype(str).str.strip(), sort=False))
    grouped.sort(key=lambda item: block_sort_key(item[0]))
    for block_id, group in grouped:
        items = []
        for _, row in group.iterrows():
            items.append({
                "id": clean(row.get("id_question", "")),
                "question": clean(row.get("question", "")),
                "options": [
                    {"value": 0, "label": clean(row.get("indice_0", "Non"))},
                    {"value": 1, "label": clean(row.get("indice_1", "Partiellement"))},
                    {"value": 2, "label": clean(row.get("indice_2", "Oui"))},
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


@app.post("/api/analyse")
def analyse_api():
    cleanup_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=False)

    try:
        convention_upload = request.files.get("convention")
        cctp_upload = request.files.get("cctp")
        convention_name = clean(getattr(convention_upload, "filename", "")) or "Convention BIM.pdf"
        cctp_name = clean(getattr(cctp_upload, "filename", "")) if cctp_upload else ""

        convention_path = job_dir / "Convention_BIM.pdf"
        cctp_path = job_dir / "CCTP_Lot.pdf"
        save_pdf(convention_upload, convention_path, required=True)
        has_cctp = save_pdf(cctp_upload, cctp_path, required=False)

        entreprise = clean(request.form.get("entreprise")) or "[Nom de l'entreprise]"
        lot = clean(request.form.get("lot")) or "Lot"
        marque = clean(request.form.get("marque")) or entreprise
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
        now = int(time.time())
        metadata = {
            "job_id": job_id,
            "project": Path(convention_name).stem,
            "entreprise": entreprise,
            "lot": lot,
            "marque": marque,
            "sans_eval": sans_eval,
            "question_total": question_total,
            "question_answered": 0 if sans_eval else len(answers),
            "documents": [name for name in [convention_name, cctp_name if has_cctp else ""] if name],
            "created_at": now,
            "status": "queued",
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
            "lot": lot,
            "marque": marque,
            "logo_path": str(logo_path) if has_logo else "",
            "sans_eval": sans_eval,
            "has_cctp": has_cctp,
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
            "status_url": url_for("job_status", job_id=job_id, _external=True),
        }), 202
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
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            time.sleep(0.15 * (attempt + 1))
    # Dernier recours : écriture directe (non atomique, mais évite de perdre
    # la progression si le renommage reste bloqué par un tiers).
    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.unlink(missing_ok=True)
    except OSError:
        if last_error:
            raise last_error


def read_json_file(path: Path, default: dict[str, Any] | None = None) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else (default or {})
    except (OSError, json.JSONDecodeError):
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
        if options.get("logo_path"):
            cmd.extend(["--logo", options["logo_path"]])
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


def _public_job_payload(job_id: str, status: dict[str, Any], metadata: dict[str, Any]) -> dict[str, Any]:
    payload = {**metadata, **status, "job_id": job_id}
    payload["status_url"] = url_for("job_status", job_id=job_id, _external=True)
    if status.get("status") == "completed":
        payload["result_url"] = url_for("result_job", job_id=job_id, _external=True)
        payload["actions_result_url"] = url_for("actions_result_job", job_id=job_id, _external=True)
        payload["glossary_result_url"] = url_for("glossary_result_job", job_id=job_id, _external=True)
        payload["result_download_url"] = url_for("download_job", job_id=job_id, kind="html", _external=True)
        payload["actions_download_url"] = url_for("download_job", job_id=job_id, kind="actions_html", _external=True)
        payload["actions_pdf_url"] = url_for("download_job", job_id=job_id, kind="actions_pdf", _external=True)
        payload["glossary_download_url"] = url_for("download_job", job_id=job_id, kind="glossary_html", _external=True)
        payload["glossary_pdf_url"] = url_for("download_job", job_id=job_id, kind="glossary_pdf", _external=True)
    return payload


@app.get("/api/jobs/<job_id>/status")
def job_status(job_id: str):
    try:
        job_dir = job_file(job_id, "status.json").parent
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    status = read_json_file(job_dir / "status.json")
    metadata = read_json_file(job_dir / "job.json")
    if not status:
        return jsonify({"error": "État de l’analyse introuvable."}), 404
    return jsonify(_public_job_payload(job_id, status, metadata))


@app.get("/api/history")
def history_api():
    cleanup_old_jobs()
    items: list[dict[str, Any]] = []
    for job_dir in JOBS_DIR.iterdir():
        if not job_dir.is_dir() or not re.fullmatch(r"[a-f0-9]{32}", job_dir.name):
            continue
        metadata = read_json_file(job_dir / "job.json")
        status = read_json_file(job_dir / "status.json")
        if not metadata:
            continue
        items.append(_public_job_payload(job_dir.name, status, metadata))
    items.sort(key=lambda item: int(item.get("created_at", 0)), reverse=True)
    return jsonify({"items": items[:50]})


@app.delete("/api/history/<job_id>")
def delete_history_job(job_id: str):
    try:
        path = job_file(job_id, "job.json").parent
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    shutil.rmtree(path, ignore_errors=True)
    return jsonify({"deleted": True})


@app.get("/api/jobs/<job_id>/result")
def result_job(job_id: str):
    try:
        path = job_file(job_id, "Rapport_BIM.html")
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    if not path.exists():
        return jsonify({"error": "Dashboard HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/actions")
def actions_result_job(job_id: str):
    try:
        path = job_file(job_id, "Plan_Actions_Offre_BIM.html")
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    if not path.exists():
        return jsonify({"error": "Plan d’actions HTML introuvable."}), 404
    return send_file(path, mimetype="text/html")


@app.get("/api/jobs/<job_id>/glossary")
def glossary_result_job(job_id: str):
    try:
        path = job_file(job_id, "Glossaire_BIM.html")
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    if not path.exists():
        return jsonify({"error": "Glossaire HTML introuvable."}), 404
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
        "log": ("analyse.log", "text/plain"),
    }
    if kind not in names:
        return jsonify({"error": "Ce livrable n’est plus généré."}), 404
    name, mimetype = names[kind]
    try:
        path = job_file(job_id, name)
    except FileNotFoundError:
        return jsonify({"error": "Analyse introuvable."}), 404
    if not path.exists():
        return jsonify({"error": "Fichier introuvable."}), 404
    metadata = read_json_file(path.parent / "job.json")
    project = secure_filename(clean(metadata.get("project")) or "Projet")
    lot = secure_filename(clean(metadata.get("lot")) or "Lot")
    download_names = {
        "html": f"Dashboard_BIM_{project}_{lot}.html",
        "actions_html": f"Plan_Actions_BIM_{project}_{lot}.html",
        "actions_pdf": f"Plan_Actions_BIM_{project}_{lot}.pdf",
        "glossary_html": f"Glossaire_BIM_{project}_{lot}.html",
        "glossary_pdf": f"Glossaire_BIM_{project}_{lot}.pdf",
        "log": f"Journal_Analyse_BIM_{project}_{lot}.txt",
    }
    return send_file(path, mimetype=mimetype, as_attachment=True, download_name=download_names[kind])


@app.errorhandler(413)
def file_too_large(_):
    return jsonify({"error": f"Fichiers trop volumineux. Limite totale : {MAX_UPLOAD_MB} Mo."}), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
