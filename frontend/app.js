const API = (window.BIM_API_URL || "http://127.0.0.1:8000").replace(/\/$/, "");
const form = document.querySelector("#analysis-form");
const questionnaire = document.querySelector("#questionnaire");
const questionnaireStatus = document.querySelector("#questionnaire-status");
const questionnairePanel = document.querySelector("#questionnaire-panel");
const sansEval = document.querySelector("#sans-eval");
const submitButton = document.querySelector("#submit-button");
const clearMemoryButton = document.querySelector("#clear-memory-button");
const progress = document.querySelector("#progress");
const historyList = document.querySelector("#history-list");
const historyStatus = document.querySelector("#history-status");
const refreshHistoryButton = document.querySelector("#refresh-history");
const conventionInput = form.querySelector('input[name="convention"]');
const cctpInput = form.querySelector('input[name="cctp"]');

const FORM_STORAGE_KEY = "bim_web_local_form_v1";
const FILE_DB_NAME = "bim_web_local_files_v1";
const FILE_STORE_NAME = "uploaded_files";

let questionIds = [];
let questionsLoading = null;
let savedState = readSavedState();

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]));
}


let glossaryByAlias = new Map();
let glossaryAliases = [];

function normalizeTerm(value) {
  return String(value || "").trim().toLocaleLowerCase("fr-FR");
}

function tooltipMarkup(display, entry) {
  const source = entry?.source ? `<small>${escapeHtml(entry.source)}</small>` : "";
  return `<span class="term-tip" tabindex="0"><span>${escapeHtml(display)}</span><span class="term-bubble">${escapeHtml(entry?.definition || "")}${source}</span></span>`;
}

function setGlossary(entries) {
  glossaryByAlias = new Map();
  (Array.isArray(entries) ? entries : []).forEach(entry => {
    [entry.term, ...(entry.aliases || [])].forEach(alias => {
      const key = normalizeTerm(alias);
      if (key && !glossaryByAlias.has(key)) glossaryByAlias.set(key, entry);
    });
  });
  glossaryAliases = [...glossaryByAlias.keys()].sort((a, b) => b.length - a.length);
  hydrateStaticTooltips();
}

function tooltipify(value) {
  const raw = String(value ?? "");
  if (!raw || !glossaryAliases.length) return escapeHtml(raw);
  const escapedAliases = glossaryAliases.map(alias => alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const pattern = new RegExp(`(^|[^\\p{L}\\p{N}_])(${escapedAliases.join("|")})(?=$|[^\\p{L}\\p{N}_])`, "giu");
  let result = "";
  let last = 0;
  raw.replace(pattern, (full, prefix, term, offset) => {
    const start = offset + prefix.length;
    result += escapeHtml(raw.slice(last, start));
    const entry = glossaryByAlias.get(normalizeTerm(term));
    result += entry ? tooltipMarkup(term, entry) : escapeHtml(term);
    last = start + term.length;
    return full;
  });
  return result + escapeHtml(raw.slice(last));
}

function hydrateStaticTooltips() {
  document.querySelectorAll("[data-term]").forEach(node => {
    const entry = glossaryByAlias.get(normalizeTerm(node.dataset.term));
    if (!entry || node.dataset.tooltipReady === "1") return;
    const display = node.textContent;
    node.innerHTML = tooltipMarkup(display, entry);
    node.dataset.tooltipReady = "1";
  });
}

function sleep(ms) {
  return new Promise(resolve => window.setTimeout(resolve, ms));
}

async function readJsonResponse(response) {
  const text = await response.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch (_error) {
    throw new Error(`Réponse inattendue du backend (HTTP ${response.status}).`);
  }
}

function readSavedState() {
  try {
    const raw = window.localStorage.getItem(FORM_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_error) {
    return {};
  }
}

function collectFormState() {
  const answers = {};
  form.querySelectorAll('input[type="radio"]:checked').forEach(input => {
    const id = input.name.startsWith("q_") ? input.name.slice(2) : "";
    if (id) answers[id] = Number(input.value);
  });

  return {
    entreprise: String(form.elements.entreprise?.value || ""),
    lot: String(form.elements.lot?.value || ""),
    sans_eval: Boolean(sansEval.checked),
    answers,
    updated_at: Date.now(),
  };
}

function saveFormState() {
  savedState = collectFormState();
  try {
    window.localStorage.setItem(FORM_STORAGE_KEY, JSON.stringify(savedState));
  } catch (_error) {
    // La persistance est un confort : une erreur de stockage ne doit pas bloquer l'analyse.
  }
}

function restoreBaseFields() {
  if (typeof savedState.entreprise === "string") form.elements.entreprise.value = savedState.entreprise;
  if (typeof savedState.lot === "string") form.elements.lot.value = savedState.lot;
  sansEval.checked = Boolean(savedState.sans_eval);
  applySansEvalState();
}

function restoreQuestionAnswers() {
  const answers = savedState.answers && typeof savedState.answers === "object" ? savedState.answers : {};
  Object.entries(answers).forEach(([id, value]) => {
    const selector = `input[name="q_${CSS.escape(id)}"][value="${CSS.escape(String(value))}"]`;
    const input = form.querySelector(selector);
    if (input) input.checked = true;
  });
}

function applySansEvalState() {
  questionnairePanel.style.opacity = sansEval.checked ? ".45" : "1";
  questionnairePanel.querySelectorAll("input").forEach(input => {
    input.disabled = sansEval.checked;
  });
}

function openFileDb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error("IndexedDB indisponible"));
      return;
    }
    const request = window.indexedDB.open(FILE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(FILE_STORE_NAME)) {
        db.createObjectStore(FILE_STORE_NAME, {keyPath: "field"});
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Impossible d’ouvrir le stockage local"));
  });
}

async function fileDbOperation(mode, action) {
  const db = await openFileDb();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction(FILE_STORE_NAME, mode);
      const store = transaction.objectStore(FILE_STORE_NAME);
      let result;
      try {
        result = action(store);
      } catch (error) {
        reject(error);
        return;
      }
      transaction.oncomplete = () => resolve(result?.result);
      transaction.onerror = () => reject(transaction.error || result?.error || new Error("Erreur de stockage local"));
      transaction.onabort = () => reject(transaction.error || new Error("Stockage local interrompu"));
    });
  } finally {
    db.close();
  }
}

async function storeFile(field, file) {
  if (!file) {
    await fileDbOperation("readwrite", store => store.delete(field));
    return;
  }
  await fileDbOperation("readwrite", store => store.put({
    field,
    name: file.name,
    type: file.type || "application/pdf",
    lastModified: file.lastModified || Date.now(),
    blob: file,
  }));
}

async function readStoredFile(field) {
  const db = await openFileDb();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction(FILE_STORE_NAME, "readonly");
      const request = transaction.objectStore(FILE_STORE_NAME).get(field);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Lecture du fichier mémorisé impossible"));
    });
  } finally {
    db.close();
  }
}

async function clearStoredFiles() {
  try {
    await fileDbOperation("readwrite", store => store.clear());
  } catch (_error) {
    // Rien à faire si le navigateur ne permet pas le stockage.
  }
}

function updateFileStatus(input, file, restored = false) {
  const status = document.querySelector(`#${input.name}-file-status`);
  if (!status) return;
  if (!file) {
    status.textContent = input.name === "convention" ? "PDF obligatoire" : "PDF facultatif, mais recommandé";
    status.classList.remove("stored-file");
    return;
  }
  const prefix = restored ? "Fichier restauré" : "Fichier mémorisé";
  status.textContent = `${prefix} : ${file.name}`;
  status.classList.add("stored-file");
}

async function rememberSelectedFile(input) {
  const file = input.files?.[0] || null;
  updateFileStatus(input, file, false);
  try {
    await storeFile(input.name, file);
  } catch (_error) {
    const status = document.querySelector(`#${input.name}-file-status`);
    if (status && file) status.textContent = `Fichier sélectionné : ${file.name} — conservation après actualisation indisponible`;
  }
}

async function restoreStoredFile(input) {
  try {
    const record = await readStoredFile(input.name);
    if (!record?.blob || !record?.name) return;

    const file = new File([record.blob], record.name, {
      type: record.type || "application/pdf",
      lastModified: record.lastModified || Date.now(),
    });
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    updateFileStatus(input, file, true);
  } catch (_error) {
    // Certains navigateurs interdisent de restaurer un input fichier. Le formulaire reste utilisable.
  }
}

async function loadQuestions({retries = 2, delayMs = 800} = {}) {
  if (questionsLoading) return questionsLoading;

  questionsLoading = (async () => {
    let lastError = null;
    questionnaireStatus.classList.remove("error");
    questionnaireStatus.textContent = "Connexion au backend et chargement du questionnaire…";

    for (let attempt = 0; attempt <= retries; attempt += 1) {
      try {
        const response = await fetch(`${API}/api/questions`, {cache: "no-store"});
        const data = await readJsonResponse(response);
        if (!response.ok) throw new Error(data.error || `Questionnaire indisponible (HTTP ${response.status})`);
        if (!Array.isArray(data.groups)) throw new Error("Format du questionnaire invalide.");
        setGlossary(data.glossary || []);

        questionIds = [];
        questionnaire.innerHTML = data.groups.map((group) => {
          const questionsHtml = group.questions.map(q => {
            questionIds.push(q.id);
            return `<div class="question">
              <p class="question-title"><span class="question-id">${escapeHtml(q.id)}</span>${tooltipify(q.question)}</p>
              <div class="options">${q.options.map(opt => `<label class="option">
                <input type="radio" name="q_${escapeHtml(q.id)}" value="${opt.value}">
                <span>${tooltipify(opt.label)}</span>
              </label>`).join("")}</div>
            </div>`;
          }).join("");
          return `<details class="question-group">
            <summary><span>${escapeHtml(group.block_id)} — ${tooltipify(group.title)}</span><span>${group.questions.length} question(s)</span></summary>
            <div class="group-body">${questionsHtml}</div>
          </details>`;
        }).join("");

        restoreQuestionAnswers();
        applySansEvalState();
        questionnaireStatus.classList.remove("error");
        questionnaireStatus.textContent = `${data.count} questions chargées. Les réponses sont mémorisées automatiquement dans ce navigateur.`;
        return true;
      } catch (error) {
        lastError = error;
        if (attempt < retries) await sleep(delayMs);
      }
    }

    questionIds = [];
    questionnaire.innerHTML = "";
    questionnaireStatus.classList.add("error");
    questionnaireStatus.innerHTML = `Connexion au backend impossible : ${escapeHtml(lastError?.message || "erreur inconnue")}. ` +
      `Vérifiez <code>${escapeHtml(API)}/health</code>, puis actualisez la page.`;
    return false;
  })();

  try {
    return await questionsLoading;
  } finally {
    questionsLoading = null;
  }
}

sansEval.addEventListener("change", () => {
  applySansEvalState();
  saveFormState();
});

form.addEventListener("input", event => {
  if (event.target instanceof HTMLInputElement && event.target.type === "file") return;
  saveFormState();
});

form.addEventListener("change", event => {
  if (event.target instanceof HTMLInputElement && event.target.type === "file") return;
  saveFormState();
});

conventionInput.addEventListener("change", () => rememberSelectedFile(conventionInput));
cctpInput.addEventListener("change", () => rememberSelectedFile(cctpInput));

function showProgress(html, isError = false) {
  const wasHidden = progress.classList.contains("hidden");
  progress.classList.remove("hidden", "error");
  if (isError) progress.classList.add("error");
  progress.innerHTML = html;
  if (wasHidden) progress.scrollIntoView({behavior: "smooth", block: "start"});
}



const ANALYSIS_STAGES = [
  [12, "Fichiers reçus"],
  [20, "Préparation technique"],
  [34, "Lecture des exigences BIM"],
  [58, "Croisement avec l’auto-évaluation"],
  [72, "Préparation des livrables HTML"],
  [94, "Vérification finale"],
  [100, "Livrables disponibles"],
];

function formatDuration(seconds) {
  const total = Number(seconds || 0);
  if (!total) return "—";
  const minutes = Math.floor(total / 60);
  const rest = total % 60;
  return minutes ? `${minutes} min ${String(rest).padStart(2, "0")} s` : `${rest} s`;
}

function formatDateTime(timestamp) {
  if (!timestamp) return "Date non disponible";
  return new Intl.DateTimeFormat("fr-FR", {
    dateStyle: "short", timeStyle: "short"
  }).format(new Date(Number(timestamp) * 1000));
}

function renderLiveProgress(status) {
  const value = Math.max(0, Math.min(100, Number(status.progress || 0)));
  const stages = ANALYSIS_STAGES.map(([threshold, label]) => {
    const done = value >= threshold;
    const current = !done && value < 100 && threshold === ANALYSIS_STAGES.find(([level]) => level > value)?.[0];
    return `<li class="${done ? "done" : current ? "current" : ""}"><span>${done ? "✓" : current ? "→" : "○"}</span>${escapeHtml(label)}</li>`;
  }).join("");
  showProgress(`<div class="live-progress">
    <div class="live-progress-head"><div><p class="eyebrow">Analyse en cours</p><h2>${escapeHtml(status.stage || "Traitement")}</h2></div><strong>${value} %</strong></div>
    <p>${escapeHtml(status.detail || "Le traitement se poursuit.")}</p>
    <div class="analysis-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${value}"><i style="width:${value}%"></i></div>
    <ol class="stage-list">${stages}</ol>
  </div>`);
}

async function pollAnalysis(statusUrl) {
  while (true) {
    const response = await fetch(statusUrl, {cache: "no-store"});
    const status = await readJsonResponse(response);
    if (!response.ok) throw new Error(status.error || `Suivi impossible (HTTP ${response.status})`);
    if (status.status === "completed") return status;
    if (status.status === "error") throw new Error(status.error || status.detail || "L’analyse a échoué.");
    renderLiveProgress(status);
    await sleep(650);
  }
}

function renderCompletedAnalysis(result) {
  showProgress(`<h2>Analyse terminée</h2>
    <p>Deux espaces complémentaires sont prêts. Toute l’analyse contractuelle et les preuves sont centralisées dans le dashboard.</p>
    <div class="deliverables">
      <div class="deliverable primary-deliverable">
        <h3>Dashboard d’analyse</h3>
        <p>Identité du projet, radar, statuts, capacités, décisions, réponses, actions, preuves et glossaire.</p>
        <div class="actions"><a class="primary" href="${escapeHtml(result.result_url)}" target="_blank" rel="noopener">Ouvrir HTML</a><a href="${escapeHtml(result.result_download_url)}" download>Télécharger HTML</a></div>
      </div>
      <div class="deliverable">
        <h3>Plan d’action</h3>
        <p>Réponses à intégrer, actions de conformité, responsables, échéances et suivi des tâches.</p>
        <div class="actions"><a class="primary" href="${escapeHtml(result.actions_pdf_url)}" download>Télécharger PDF</a><a href="${escapeHtml(result.actions_result_url)}" target="_blank" rel="noopener">Ouvrir HTML</a></div>
      </div>
      <div class="deliverable">
        <h3>Glossaire BIM</h3>
        <p>Termes techniques détectés dans les documents, le questionnaire et l’outil, avec leur définition.</p>
        <div class="actions"><a class="primary" href="${escapeHtml(result.glossary_pdf_url)}" download>Télécharger PDF</a><a href="${escapeHtml(result.glossary_result_url)}" target="_blank" rel="noopener">Ouvrir HTML</a></div>
      </div>
    </div>
    <p class="completion-meta">Durée : <strong>${escapeHtml(formatDuration(result.duration_seconds))}</strong></p>`);
}

function historyStatusLabel(status) {
  return ({completed: "Terminée", running: "En cours", queued: "En attente", error: "En erreur"})[status] || "Inconnue";
}

function renderHistory(items) {
  if (!items.length) {
    historyStatus.classList.remove("hidden", "error");
    historyStatus.textContent = "Aucune analyse enregistrée pour le moment.";
    historyList.innerHTML = "";
    return;
  }
  historyStatus.classList.add("hidden");
  historyList.innerHTML = items.map(item => {
    const completed = item.status === "completed";
    const links = completed ? `<div class="history-actions">
      <a href="${escapeHtml(item.result_url)}" target="_blank" rel="noopener">Dashboard - Ouvrir HTML</a>
      <a href="${escapeHtml(item.actions_pdf_url)}" download>Plan d’action - Télécharger PDF</a>
      <a href="${escapeHtml(item.actions_result_url)}" target="_blank" rel="noopener">Plan d’action - Ouvrir HTML</a>
      <a href="${escapeHtml(item.glossary_pdf_url)}" download>Glossaire - Télécharger PDF</a>
      <a href="${escapeHtml(item.glossary_result_url)}" target="_blank" rel="noopener">Glossaire - Ouvrir HTML</a>
    </div>` : "";
    const documents = Array.isArray(item.documents) ? item.documents.join(" · ") : "Documents non renseignés";
    return `<article class="history-item" data-job-id="${escapeHtml(item.job_id)}">
      <div class="history-main">
        <div class="history-topline"><strong>${escapeHtml(item.project || "Analyse BIM")}</strong><span class="history-badge ${escapeHtml(item.status || "")}">${escapeHtml(historyStatusLabel(item.status))}</span></div>
        <p>${escapeHtml(item.entreprise || "Entreprise non renseignée")} · Lot ${escapeHtml(item.lot || "—")}</p>
        <small>${escapeHtml(formatDateTime(item.created_at))} · ${escapeHtml(documents)}</small>
        ${item.status === "running" || item.status === "queued" ? `<div class="mini-track"><i style="width:${Number(item.progress || 0)}%"></i></div><small>${escapeHtml(item.stage || "Traitement")} — ${Number(item.progress || 0)} %</small>` : ""}
        ${item.status === "error" ? `<p class="history-error">${escapeHtml(item.error || item.detail || "Analyse interrompue")}</p>` : ""}
      </div>
      <div class="history-side"><span>${escapeHtml(formatDuration(item.duration_seconds))}</span>${links}${item.status === "running" || item.status === "queued" ? "" : `<button type="button" class="delete-history" data-delete-job="${escapeHtml(item.job_id)}">Supprimer</button>`}</div>
    </article>`;
  }).join("");
}

async function loadHistory() {
  historyStatus.classList.remove("hidden", "error");
  historyStatus.textContent = "Chargement de l’historique…";
  try {
    const response = await fetch(`${API}/api/history`, {cache: "no-store"});
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || `Historique indisponible (HTTP ${response.status})`);
    renderHistory(Array.isArray(data.items) ? data.items : []);
  } catch (error) {
    historyStatus.classList.remove("hidden");
    historyStatus.classList.add("error");
    historyStatus.textContent = error.message;
  }
}

refreshHistoryButton.addEventListener("click", loadHistory);
historyList.addEventListener("click", async event => {
  const button = event.target.closest("[data-delete-job]");
  if (!button) return;
  if (!window.confirm("Supprimer cette analyse de l’historique local ?")) return;
  button.disabled = true;
  try {
    const response = await fetch(`${API}/api/history/${encodeURIComponent(button.dataset.deleteJob)}`, {method: "DELETE"});
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "Suppression impossible.");
    await loadHistory();
  } catch (error) {
    button.disabled = false;
    window.alert(error.message);
  }
});

clearMemoryButton.addEventListener("click", async () => {
  const confirmed = window.confirm("Effacer les réponses, les informations saisies et les PDF mémorisés dans ce navigateur ?");
  if (!confirmed) return;

  window.localStorage.removeItem(FORM_STORAGE_KEY);
  savedState = {};
  await clearStoredFiles();
  form.reset();
  questionnaire.querySelectorAll('input[type="radio"]').forEach(input => { input.checked = false; });
  updateFileStatus(conventionInput, null);
  updateFileStatus(cctpInput, null);
  applySansEvalState();
  progress.classList.add("hidden");
  progress.innerHTML = "";
  saveFormState();
});

form.addEventListener("submit", async event => {
  event.preventDefault();
  saveFormState();
  const data = new FormData(form);
  const convention = data.get("convention");
  const lot = String(data.get("lot") || "").trim();

  if (!convention || !convention.name) {
    showProgress("<h2>Convention manquante</h2><p>Ajoutez la convention BIM au format PDF.</p>", true);
    return;
  }
  if (!lot) {
    showProgress("<h2>Lot manquant</h2><p>Indiquez le lot à analyser.</p>", true);
    return;
  }

  const answers = {};
  if (!sansEval.checked) {
    if (!questionIds.length) {
      showProgress("<h2>Vérification du questionnaire</h2><p>Nouvelle tentative de connexion au backend…</p>");
      const loaded = await loadQuestions({retries: 1, delayMs: 500});
      if (!loaded || !questionIds.length) {
        showProgress(`<h2>Questionnaire indisponible</h2>
          <p>Le questionnaire n’a pas été chargé. Vérifiez <code>${escapeHtml(API)}/health</code>, puis actualisez la page.</p>
          <p>Pour tester uniquement la génération, cochez <strong>Test rapide sans auto-évaluation</strong>.</p>`, true);
        return;
      }
    }

    const missing = [];
    for (const id of questionIds) {
      const selected = form.querySelector(`input[name="q_${CSS.escape(id)}"]:checked`);
      if (!selected) missing.push(id);
      else answers[id] = Number(selected.value);
    }
    if (missing.length) {
      showProgress(`<h2>Questionnaire incomplet</h2><p>${missing.length} réponse(s) manquante(s). Première question non renseignée : <strong>${escapeHtml(missing[0])}</strong>.</p>`, true);
      return;
    }
  }

  data.set("sans_eval", sansEval.checked ? "true" : "false");
  data.set("answers", JSON.stringify(answers));
  submitButton.disabled = true;
  submitButton.textContent = "Analyse en cours…";
  renderLiveProgress({progress: 2, stage: "Envoi des documents", detail: "Transmission des fichiers au backend local."});

  try {
    const response = await fetch(`${API}/api/analyse`, {method: "POST", body: data});
    const queued = await readJsonResponse(response);
    if (!response.ok) throw new Error(queued.error || `Échec du lancement (HTTP ${response.status})`);
    saveFormState();
    renderLiveProgress({progress: 5, stage: "Analyse enregistrée", detail: "Les fichiers ont été reçus par le backend."});
    const result = await pollAnalysis(queued.status_url);
    renderCompletedAnalysis(result);
    await loadHistory();
  } catch (error) {
    const networkHint = error instanceof TypeError
      ? `<p>Le navigateur n’a pas réussi à joindre <code>${escapeHtml(API)}</code>. Vérifiez que le backend est toujours ouvert.</p>`
      : "";
    showProgress(`<h2>Analyse interrompue</h2><p>${escapeHtml(error.message)}</p>${networkHint}<p>Consultez le terminal du backend et, si un identifiant de traitement a été créé, le fichier <code>runtime/jobs/&lt;job&gt;/analyse.log</code>.</p>`, true);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "Analyser les documents";
  }
});

async function initialize() {
  restoreBaseFields();
  await Promise.all([
    restoreStoredFile(conventionInput),
    restoreStoredFile(cctpInput),
  ]);
  await Promise.all([loadQuestions(), loadHistory()]);
}

initialize();
