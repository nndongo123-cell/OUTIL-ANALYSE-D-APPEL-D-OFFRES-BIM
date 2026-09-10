const API = (window.BIM_API_URL || "http://127.0.0.1:8000").replace(/\/$/, "");
const CLIENT_KEY_STORAGE = "bim_client_key_v1";

function getClientKey() {
  let key = "";
  try { key = String(window.localStorage.getItem(CLIENT_KEY_STORAGE) || ""); } catch (_error) {}
  if (/^[A-Za-z0-9_-]{24,160}$/.test(key)) return key;
  const bytes = new Uint8Array(32);
  window.crypto.getRandomValues(bytes);
  key = Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  try { window.localStorage.setItem(CLIENT_KEY_STORAGE, key); } catch (_error) {}
  return key;
}

const CLIENT_KEY = getClientKey();

function apiFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-BIM-Client-Key", CLIENT_KEY);
  return fetch(url, {...options, headers});
}

const form = document.querySelector("#analysis-form");
const questionnaire = document.querySelector("#questionnaire");
const questionnaireStatus = document.querySelector("#questionnaire-status");
const questionnairePanel = document.querySelector("#questionnaire-panel");
const sansEval = document.querySelector("#sans-eval");
const submitButton = document.querySelector("#submit-button");
const autoevalButton = document.querySelector("#autoeval-button");
const clearMemoryButton = document.querySelector("#clear-memory-button");
const precheckButton = document.querySelector("#precheck-button");
const livePrecheck = document.querySelector("#live-precheck");
const livePrecheckText = document.querySelector("#live-precheck-text");
const progress = document.querySelector("#progress");
const historyList = document.querySelector("#history-list");
const historyStatus = document.querySelector("#history-status");
const refreshHistoryButton = document.querySelector("#refresh-history");
const conventionInput = form.querySelector('input[name="convention"]');
const cctpInput = form.querySelector('input[name="cctp"]');
const ccapInput = form.querySelector('input[name="ccap"]');

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
    projet: String(form.elements.projet?.value || ""),
    entreprise: String(form.elements.entreprise?.value || ""),
    lot: String(form.elements.lot?.value || ""),
    adresse: String(form.elements.adresse?.value || ""),
    email: String(form.elements.email?.value || ""),
    analyste: String(form.elements.analyste?.value || ""),
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
  if (typeof savedState.projet === "string") form.elements.projet.value = savedState.projet;
  if (typeof savedState.entreprise === "string") form.elements.entreprise.value = savedState.entreprise;
  if (typeof savedState.lot === "string") form.elements.lot.value = savedState.lot;
  if (typeof savedState.adresse === "string") form.elements.adresse.value = savedState.adresse;
  if (typeof savedState.email === "string") form.elements.email.value = savedState.email;
  if (typeof savedState.analyste === "string") form.elements.analyste.value = savedState.analyste;
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
  const clearButton = document.querySelector(`[data-clear-file="${CSS.escape(input.name)}"]`);
  if (clearButton) {
    clearButton.disabled = !file;
    clearButton.setAttribute("aria-disabled", String(!file));
  }
  if (!status) return;
  if (!file) {
    if (input.name === "convention") status.textContent = "PDF obligatoire";
    else if (input.name === "ccap") status.textContent = "PDF facultatif — repérage documentaire si une hiérarchie explicite est trouvée";
    else status.textContent = "PDF facultatif, mais recommandé";
    status.classList.remove("stored-file");
    return;
  }
  const prefix = restored ? "Fichier restauré" : "Fichier mémorisé";
  status.textContent = `${prefix} : ${file.name}`;
  status.classList.add("stored-file");
}


let livePrecheckTimer = null;
let lastLivePrecheck = null;

function setLivePrecheck(state, text) {
  if (!livePrecheck || !livePrecheckText) return;
  livePrecheck.classList.remove("pending", "ok", "warning", "danger", "checking");
  livePrecheck.classList.add(state || "pending");
  livePrecheckText.textContent = text;
}

function invalidateLivePrecheck(text = "Dossier modifié — un nouveau contrôle est nécessaire avant analyse.") {
  lastLivePrecheck = null;
  setLivePrecheck("pending", text);
}

function buildPrecheckFormData() {
  const convention = conventionInput.files?.[0] || null;
  const cctp = cctpInput.files?.[0] || null;
  const ccap = ccapInput.files?.[0] || null;
  const projet = String(form.elements.projet?.value || "").trim();
  const lot = String(form.elements.lot?.value || "").trim();
  if (!convention || !convention.name || !projet || !lot) return null;
  const fd = new FormData();
  fd.set("convention", convention, convention.name);
  if (cctp?.name) fd.set("cctp", cctp, cctp.name);
  if (ccap?.name) fd.set("ccap", ccap, ccap.name);
  fd.set("projet", projet);
  fd.set("lot", lot);
  return fd;
}

function summarizePrecheck(check) {
  const blocking = Boolean(check?.has_blocking_error || (check?.blocking_errors || []).length);
  if (blocking) return {state: "danger", text: `Blocage documentaire : ${(check.blocking_errors || []).join(" ")}`};
  if (check?.has_alert) {
    const parts = [];
    if (check?.types_coherence && check.types_coherence !== "MATCH") parts.push("type de pièce à vérifier");
    if (check?.projet_coherence && check.projet_coherence !== "MATCH") parts.push("projet à vérifier");
    if (check?.coherence && check.coherence !== "MATCH") parts.push("lot/CCTP à vérifier");
    if (check?.documents_coherence && check.documents_coherence !== "MATCH") parts.push("cohérence entre documents à vérifier");
    if (check?.reading_coherence && check.reading_coherence !== "MATCH") parts.push("lisibilité PDF / pages scannées à vérifier");
    return {state: "warning", text: `Alerte avant analyse : ${parts.join(" ; ") || "vérification manuelle requise"}. Cliquez sur « Vérifier le dossier ». `};
  }
  return {state: "ok", text: "Contrôle automatique cohérent : type des pièces, projet, lot et dossier. Une validation restera obligatoire au lancement."};
}

async function runLivePrecheck({openModal = false, requireReview = false} = {}) {
  const fd = buildPrecheckFormData();
  if (!fd) {
    setLivePrecheck("pending", "Chargez une Convention BIM puis renseignez le nom du projet et le lot pour lancer le contrôle préalable.");
    return null;
  }
  setLivePrecheck("checking", "Contrôle du type des pièces, du projet, du lot et de la cohérence du dossier…");
  try {
    const response = await apiFetch(`${API}/api/precheck`, {method: "POST", body: fd});
    const check = await readJsonResponse(response);
    if (!response.ok) throw new Error(check.error || `Vérification impossible (HTTP ${response.status})`);
    lastLivePrecheck = check;
    const summary = summarizePrecheck(check);
    setLivePrecheck(summary.state, summary.text);
    if (openModal) {
      const decision = await showPreAnalysisModal(check, "Auto-évaluation : contrôle documentaire uniquement.", {requireReview});
      return decision ? {...check, user_override: Boolean(decision.override)} : null;
    }
    return check;
  } catch (error) {
    lastLivePrecheck = null;
    setLivePrecheck("danger", `Contrôle préalable indisponible : ${error.message}. L'analyse sera bloquée tant que le contrôle n'aura pas abouti.`);
    return null;
  }
}

function scheduleLivePrecheck(delay = 450) {
  invalidateLivePrecheck();
  if (livePrecheckTimer) window.clearTimeout(livePrecheckTimer);
  livePrecheckTimer = window.setTimeout(() => { runLivePrecheck().catch(() => {}); }, delay);
}

async function unloadDocument(input) {
  if (!input) return;
  input.value = "";
  updateFileStatus(input, null, false);
  try { await storeFile(input.name, null); } catch (_error) {}
  if (input.name === "convention") {
    projetField.value = "";
    projetField.dataset.autofilled = "false";
    if (projetStatus) projetStatus.textContent = "Convention déchargée — le nom du projet sera redétecté à la prochaine Convention BIM.";
  }
  saveFormState();
  scheduleLivePrecheck(250);
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
    questionnaireStatus.textContent = "Connexion au service d’analyse et chargement du questionnaire…";

    for (let attempt = 0; attempt <= retries; attempt += 1) {
      try {
        const response = await apiFetch(`${API}/api/questions`, {cache: "no-store"});
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
    questionnaireStatus.innerHTML = `Connexion au service d’analyse impossible : ${escapeHtml(lastError?.message || "erreur inconnue")}. ` +
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

conventionInput.addEventListener("change", async () => { await rememberSelectedFile(conventionInput); invalidateLivePrecheck("Nouvelle Convention chargée — contrôle en attente de redétection du projet."); });
cctpInput.addEventListener("change", async () => { await rememberSelectedFile(cctpInput); scheduleLivePrecheck(250); });
ccapInput.addEventListener("change", async () => { await rememberSelectedFile(ccapInput); scheduleLivePrecheck(250); });

document.querySelectorAll("[data-clear-file]").forEach(button => {
  button.addEventListener("click", async event => {
    event.preventDefault();
    event.stopPropagation();
    const field = String(button.dataset.clearFile || "");
    const input = form.querySelector(`input[type="file"][name="${CSS.escape(field)}"]`);
    await unloadDocument(input);
  });
});

const projetField = form.elements.projet;
const projetStatus = document.getElementById("projet-status");

async function detectProjectFromConvention(file, {resetFirst = true} = {}) {
  if (!file) return;
  if (resetFirst) {
    // Une nouvelle Convention invalide toujours le nom issu du dossier précédent.
    projetField.value = "";
    projetField.dataset.autofilled = "false";
  }
  projetStatus.textContent = "Détection du nom du projet en cours…";
  try {
    const fd = new FormData();
    fd.append("convention", file);
    const resp = await apiFetch(`${API}/api/detect-project`, {method: "POST", body: fd});
    const data = await readJsonResponse(resp);
    if (!resp.ok) {
      projetStatus.textContent = data.error || "Le PDF ne peut pas être lu automatiquement. Saisissez le nom du projet : une validation sera demandée avant analyse.";
      saveFormState();
      return;
    }
    const roleAlert = data.document_role_coherence && data.document_role_coherence !== "MATCH"
      ? ` ALERTE TYPE DE DOCUMENT : ${data.document_role_warning || "le PDF chargé n'est pas reconnu comme une Convention BIM."}`
      : "";
    if (data.projet) {
      projetField.value = data.projet;
      projetField.dataset.autofilled = "true";
      projetStatus.textContent = (data.warning_detail
        ? `Nom redétecté depuis la Convention BIM. Attention : ${data.warning_detail}`
        : "Nom redétecté depuis la Convention BIM — vérifiez-le avant de lancer l’analyse.") + roleAlert;
      if (data.document_role_coherence !== "MATCH") setLivePrecheck(data.document_role_coherence === "MISMATCH" ? "danger" : "warning", data.document_role_warning || "Le PDF chargé comme Convention BIM doit être vérifié.");
    } else {
      projetField.value = "";
      projetField.dataset.autofilled = "false";
      projetStatus.textContent = data.warning_detail
        ? `Nom du projet non détecté. Attention : ${data.warning_detail} Saisissez-le manuellement ; une alerte sera affichée avant analyse.`
        : "Nom du projet non détecté automatiquement — saisissez-le manuellement ; une alerte sera affichée avant analyse.";
    }
  } catch (_err) {
    projetField.value = "";
    projetField.dataset.autofilled = "false";
    projetStatus.textContent = "Détection automatique indisponible — saisissez le nom du projet ; une validation sera demandée avant analyse.";
  }
  saveFormState();
  scheduleLivePrecheck(250);
}

conventionInput.addEventListener("change", async () => {
  const file = conventionInput.files && conventionInput.files[0];
  if (!file) {
    projetField.value = "";
    projetField.dataset.autofilled = "false";
    projetStatus.textContent = "Aucune Convention BIM chargée.";
    saveFormState();
    return;
  }
  await detectProjectFromConvention(file, {resetFirst: true});
});
projetField.addEventListener("input", () => {
  projetField.dataset.autofilled = "false";
  saveFormState();
  scheduleLivePrecheck(650);
});
form.elements.lot?.addEventListener("input", () => scheduleLivePrecheck(650));
precheckButton?.addEventListener("click", async () => {
  precheckButton.disabled = true;
  try { await runLivePrecheck({openModal: true, requireReview: false}); }
  finally { precheckButton.disabled = false; }
});

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
  let currentIndex = 0;
  for (let i = 0; i < ANALYSIS_STAGES.length; i += 1) {
    if (value >= ANALYSIS_STAGES[i][0]) currentIndex = i;
  }
  const stages = ANALYSIS_STAGES.map(([threshold, label], index) => {
    const current = value < 100 && index === currentIndex;
    const done = value >= 100 || index < currentIndex;
    return `<li class="${done ? "done" : current ? "current" : ""}"><span>${done ? "✓" : current ? "→" : "○"}</span>${escapeHtml(label)}</li>`;
  }).join("");
  showProgress(`<div class="live-progress">
    <div class="live-progress-head"><div><p class="eyebrow">Analyse en cours</p><h2>${escapeHtml(status.stage || "Traitement")}</h2></div><strong>${value} %</strong></div>
    <p>${escapeHtml(status.detail || "Le traitement se poursuit.")}</p>
    <div class="analysis-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${value}"><i style="width:${value}%"></i></div>
    <ol class="stage-list">${stages}</ol>
  </div>`);
}

async function pollAnalysis(statusUrl, mode = "analysis") {
  let consecutiveFollowupErrors = 0;
  while (true) {
    try {
      const response = await apiFetch(statusUrl, {cache: "no-store"});
      const status = await readJsonResponse(response);
      if (!response.ok) {
        const transient = response.status === 503 || (response.status === 404 && consecutiveFollowupErrors < 2);
        if (transient && consecutiveFollowupErrors < 6) {
          consecutiveFollowupErrors += 1;
          await sleep(450);
          continue;
        }
        throw new Error(status.error || `Suivi impossible (HTTP ${response.status})`);
      }
      consecutiveFollowupErrors = 0;
      if (status.status === "completed") return status;
      if (status.status === "error") throw new Error(status.error || status.detail || "L’analyse a échoué.");
      if (mode === "autoevaluation") {
        showProgress(`<div class="live-progress"><div class="live-progress-head"><div><p class="eyebrow">Auto-évaluation en cours</p><h2>${escapeHtml(status.stage || "Traitement")}</h2></div><strong>${Math.max(0, Math.min(100, Number(status.progress || 0)))} %</strong></div><p>${escapeHtml(status.detail || "Le traitement se poursuit.")}</p><div class="analysis-track"><i style="width:${Math.max(0, Math.min(100, Number(status.progress || 0)))}%"></i></div></div>`);
      } else {
        renderLiveProgress(status);
      }
      await sleep(650);
    } catch (error) {
      if (consecutiveFollowupErrors < 6 && (error instanceof TypeError || /HTTP 503|HTTP 404|fetch/i.test(String(error?.message || "")))) {
        consecutiveFollowupErrors += 1;
        await sleep(450);
        continue;
      }
      throw error;
    }
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

function renderCompletedAutoevaluation(result) {
  showProgress(`<h2>Auto-évaluation terminée</h2>
    <p>Le rapport autonome présente les capacités déclarées par bloc, indépendamment d’un appel d’offres.</p>
    <div class="deliverables">
      <div class="deliverable primary-deliverable">
        <h3>Rapport d’auto-évaluation</h3>
        <p>Résultats par bloc, radar, capacités déclarées et points à renforcer.</p>
        <div class="actions"><a class="primary" href="${escapeHtml(result.autoevaluation_result_url)}" target="_blank" rel="noopener">Ouvrir HTML</a><a href="${escapeHtml(result.autoevaluation_pdf_url)}" download>Télécharger PDF</a><a href="${escapeHtml(result.autoevaluation_html_url)}" download>Télécharger HTML</a></div>
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
    historyStatus.textContent = "Aucun traitement enregistré pour le moment.";
    historyList.innerHTML = "";
    return;
  }
  historyStatus.classList.add("hidden");
  historyList.innerHTML = items.map(item => {
    const completed = item.status === "completed";
    const isAuto = item.job_type === "autoevaluation";
    const links = completed ? (isAuto ? `<div class="history-actions">
      <a href="${escapeHtml(item.autoevaluation_result_url)}" target="_blank" rel="noopener">Auto-évaluation - Ouvrir HTML</a>
      <a href="${escapeHtml(item.autoevaluation_pdf_url)}" download>Auto-évaluation - Télécharger PDF</a>
    </div>` : `<div class="history-actions">
      <a href="${escapeHtml(item.result_url)}" target="_blank" rel="noopener">Dashboard - Ouvrir HTML</a>
      <a href="${escapeHtml(item.actions_pdf_url)}" download>Plan d’action - Télécharger PDF</a>
      <a href="${escapeHtml(item.actions_result_url)}" target="_blank" rel="noopener">Plan d’action - Ouvrir HTML</a>
      <a href="${escapeHtml(item.glossary_pdf_url)}" download>Glossaire - Télécharger PDF</a>
      <a href="${escapeHtml(item.glossary_result_url)}" target="_blank" rel="noopener">Glossaire - Ouvrir HTML</a>
    </div>`) : "";
    const documents = Array.isArray(item.documents) && item.documents.length ? item.documents.join(" · ") : (isAuto ? "Sans document" : "Documents non renseignés");
    const title = isAuto ? "Auto-évaluation BIM" : (item.project || "Analyse BIM");
    return `<article class="history-item" data-job-id="${escapeHtml(item.job_id)}">
      <div class="history-main">
        <div class="history-topline"><strong>${escapeHtml(title)}</strong><span class="history-badge ${escapeHtml(item.status || "")}">${escapeHtml(historyStatusLabel(item.status))}</span></div>
        <p>${escapeHtml(item.entreprise || "Entreprise non renseignée")} · Lot ${escapeHtml(item.lot || "—")}</p>
        <small>${escapeHtml(formatDateTime(item.created_at))} · ${escapeHtml(documents)}</small>
        ${item.status === "running" || item.status === "queued" ? `<div class="mini-track"><i style="width:${Number(item.progress || 0)}%"></i></div><small>${escapeHtml(item.stage || "Traitement")} — ${Number(item.progress || 0)} %</small>` : ""}
        ${item.status === "error" ? `<p class="history-error">${escapeHtml(item.error || item.detail || "Traitement interrompu")}</p>` : ""}
      </div>
      <div class="history-side"><span>${escapeHtml(formatDuration(item.duration_seconds))}</span>${links}${item.status === "running" || item.status === "queued" ? "" : `<button type="button" class="delete-history" data-delete-job="${escapeHtml(item.job_id)}">Supprimer</button>`}</div>
    </article>`;
  }).join("");
}

async function loadHistory() {
  historyStatus.classList.remove("hidden", "error");
  historyStatus.textContent = "Chargement de l’historique…";
  try {
    const response = await apiFetch(`${API}/api/history`, {cache: "no-store"});
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
    const response = await apiFetch(`${API}/api/history/${encodeURIComponent(button.dataset.deleteJob)}`, {method: "DELETE"});
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
  updateFileStatus(ccapInput, null);
  applySansEvalState();
  progress.classList.add("hidden");
  progress.innerHTML = "";
  saveFormState();
});

function collectCompleteAnswers() {
  if (!questionIds.length) return {answers: {}, missing: ["questionnaire"]};
  const answers = {}; const missing = [];
  for (const id of questionIds) {
    const selected = form.querySelector(`input[name="q_${CSS.escape(id)}"]:checked`);
    if (!selected) missing.push(id); else answers[id] = Number(selected.value);
  }
  return {answers, missing};
}

autoevalButton.addEventListener("click", async () => {
  saveFormState();
  if (!questionIds.length) {
    const loaded = await loadQuestions({retries: 1, delayMs: 500});
    if (!loaded || !questionIds.length) { showProgress("<h2>Questionnaire indisponible</h2><p>Impossible de lancer l’auto-évaluation seule sans le questionnaire.</p>", true); return; }
  }
  const {answers, missing} = collectCompleteAnswers();
  if (missing.length) {
    showProgress(`<h2>Questionnaire incomplet</h2><p>${missing.length} réponse(s) manquante(s). Première question non renseignée : <strong>${escapeHtml(missing[0])}</strong>.</p>`, true); return;
  }
  const data = new FormData();
  ["entreprise", "lot", "adresse", "email", "analyste"].forEach(name => data.set(name, String(form.elements[name]?.value || "")));
  const logoFile = form.elements.logo?.files?.[0];
  if (logoFile) data.set("logo", logoFile, logoFile.name);
  data.set("answers", JSON.stringify(answers));
  autoevalButton.disabled = true; submitButton.disabled = true; autoevalButton.textContent = "Auto-évaluation en cours…";
  showProgress(`<div class="live-progress"><div class="live-progress-head"><div><p class="eyebrow">Auto-évaluation en cours</p><h2>Envoi des réponses</h2></div><strong>2 %</strong></div><p>Aucun document n’est transmis pour ce parcours.</p><div class="analysis-track"><i style="width:2%"></i></div></div>`);
  try {
    const response = await apiFetch(`${API}/api/autoevaluation`, {method: "POST", body: data});
    const queued = await readJsonResponse(response);
    if (!response.ok) throw new Error(queued.error || `Échec du lancement (HTTP ${response.status})`);
    const result = await pollAnalysis(queued.status_url, "autoevaluation");
    renderCompletedAutoevaluation(result); await loadHistory();
  } catch (error) {
    showProgress(`<h2>Auto-évaluation interrompue</h2><p>${escapeHtml(error.message)}</p>`, true);
  } finally {
    autoevalButton.disabled = false; submitButton.disabled = false; autoevalButton.textContent = "Auto-évaluation seule";
  }
});

function precheckStateLabel(value) {
  if (value === "MATCH") return {text: "Cohérent", cls: "ok"};
  if (value === "MISMATCH") return {text: "Alerte", cls: "warning"};
  if (value === "NO_CCTP") return {text: "Non vérifié", cls: "neutral"};
  return {text: "À vérifier", cls: "neutral"};
}

function showPreAnalysisModal(check, evaluationText, options = {}) {
  return new Promise(resolve => {
    const requireReview = options.requireReview !== false;
    const lotState = precheckStateLabel(check?.coherence);
    const projectState = precheckStateLabel(check?.projet_coherence);
    const docsState = precheckStateLabel(check?.documents_coherence);
    const typesState = precheckStateLabel(check?.types_coherence);
    const readingState = precheckStateLabel(check?.reading_coherence);
    const documentChecks = Array.isArray(check?.document_checks) ? check.document_checks : [];
    const roleChecks = Array.isArray(check?.role_checks) ? check.role_checks : [];
    const readingChecks = Array.isArray(check?.reading_checks) ? check.reading_checks : [];
    const blockingErrors = Array.isArray(check?.blocking_errors) ? check.blocking_errors : [];
    const hasBlocking = Boolean(check?.has_blocking_error || blockingErrors.length);
    const hasAlert = Boolean(check?.has_alert || [check?.coherence, check?.projet_coherence, check?.documents_coherence, check?.types_coherence, check?.reading_coherence].some(v => v && v !== "MATCH"));
    const overlay = document.createElement("div");
    overlay.className = "precheck-overlay";
    overlay.setAttribute("role", "presentation");

    const alerts = [];
    blockingErrors.forEach(text => alerts.push({type: "danger", text: `BLOCAGE : ${text}`}));
    if (check?.projet_warning) alerts.push({type: check?.projet_coherence === "MISMATCH" ? "danger" : "", text: check.projet_warning});
    if (check?.warning) alerts.push({type: check?.coherence === "MISMATCH" ? "danger" : "", text: check.warning});
    if (check?.documents_warning) alerts.push({type: check?.documents_coherence === "MISMATCH" ? "danger" : "", text: check.documents_warning});
    roleChecks.filter(x => x.coherence !== "MATCH").forEach(x => alerts.push({type: x.coherence === "MISMATCH" ? "danger" : "", text: x.detail}));
    readingChecks.filter(x => x.coherence !== "MATCH").forEach(x => alerts.push({type: x.coherence === "MISMATCH" ? "danger" : "", text: `${x.document || "Document"} — ${x.detail || "Lecture PDF à vérifier."}`}));

    const roleHtml = roleChecks.length ? `
      <div class="precheck-docs">
        <div class="precheck-docs-title">Type réel des documents chargés</div>
        ${roleChecks.map(item => {
          const state = precheckStateLabel(item.coherence);
          return `<div class="precheck-doc-row">
            <strong>${escapeHtml(item.document || "Document")}</strong>
            <small>${escapeHtml(item.filename || "—")}<br>Attendu : ${escapeHtml(item.expected_label || "—")} · Détecté : ${escapeHtml(item.detected_label || "—")}<br>${escapeHtml(item.detail || "")}</small>
            <span class="precheck-status ${state.cls}">${escapeHtml(state.text)}</span>
          </div>`;
        }).join("")}
      </div>` : "";

    const readingHtml = readingChecks.length ? `
      <div class="precheck-docs">
        <div class="precheck-docs-title">Lisibilité PDF — pages scannées / images sans couche texte</div>
        ${readingChecks.map(item => {
          const state = precheckStateLabel(item.coherence);
          const pages = Array.isArray(item.pages_images_non_lisibles) && item.pages_images_non_lisibles.length
            ? ` · Pages non analysables : ${item.pages_images_non_lisibles.join(", ")}` : "";
          return `<div class="precheck-doc-row">
            <strong>${escapeHtml(item.document || "Document")}</strong>
            <small>${escapeHtml(item.filename || "—")}<br>Lecture : ${escapeHtml(item.statut || "—")} · ${escapeHtml(item.pages_lisibles ?? "—")}/${escapeHtml(item.pages_total ?? "—")} page(s) exploitable(s)${escapeHtml(pages)}<br>${escapeHtml(item.detail || "")}</small>
            <span class="precheck-status ${state.cls}">${escapeHtml(state.text)}</span>
          </div>`;
        }).join("")}
      </div>` : "";

    const docsHtml = documentChecks.length ? `
      <div class="precheck-docs">
        <div class="precheck-docs-title">Même projet — comparaison des ${escapeHtml(check?.pages_comparees || 2)} premières pages</div>
        ${documentChecks.map(item => {
          const state = precheckStateLabel(item.coherence);
          return `<div class="precheck-doc-row">
            <strong>${escapeHtml(item.document || "Document")}</strong>
            <small>${escapeHtml(item.filename || "—")}<br>${escapeHtml(item.detail || "")}</small>
            <span class="precheck-status ${state.cls}">${escapeHtml(state.text)}</span>
          </div>`;
        }).join("")}
      </div>` : "";

    overlay.innerHTML = `
      <section class="precheck-dialog" role="dialog" aria-modal="true" aria-labelledby="precheckTitle">
        <div class="precheck-head">
          <p class="precheck-kicker">Avant de lancer l'analyse</p>
          <h2 id="precheckTitle">Validation obligatoire du projet, du lot et des documents</h2>
        </div>
        <div class="precheck-body">
          <div class="precheck-summary">
            <div><span>Projet saisi</span><strong>${escapeHtml(check?.projet_saisi || "—")}</strong></div>
            <div><span>Projet détecté dans la convention</span><strong>${escapeHtml(check?.projet_detecte || "Non détecté")}</strong></div>
            <div><span>Cohérence projet</span><strong class="precheck-status ${projectState.cls}">${escapeHtml(projectState.text)}</strong></div>
            <div><span>Lot analysé</span><strong>${escapeHtml(check?.lot_saisi || "—")}</strong></div>
            <div><span>Lot détecté dans le CCTP</span><strong>${escapeHtml(check?.lot_detecte || "Vérification par contenu")}</strong></div>
            <div><span>Cohérence lot / CCTP</span><strong class="precheck-status ${lotState.cls}">${escapeHtml(lotState.text)}</strong></div>
            <div><span>Type des pièces</span><strong class="precheck-status ${typesState.cls}">${escapeHtml(typesState.text)}</strong></div>
            <div><span>Lisibilité PDF / pages scannées</span><strong class="precheck-status ${readingState.cls}">${escapeHtml(readingState.text)}</strong></div>
            <div><span>Cohérence du projet entre pièces</span><strong class="precheck-status ${docsState.cls}">${escapeHtml(docsState.text)}</strong></div>
            <div><span>Auto-évaluation</span><strong>${escapeHtml(evaluationText.replace(/^Auto-évaluation\s*:\s*/i, ""))}</strong></div>
          </div>
          ${roleHtml}
          ${readingHtml}
          ${docsHtml}
          ${alerts.length ? `<div class="precheck-alerts">${alerts.map(a => `<div class="precheck-warning ${a.type}">${escapeHtml(a.text)}</div>`).join("")}</div>` : ""}
          ${hasBlocking ? `<div class="precheck-blocking-message">Corrigez le document bloquant (type de pièce ou Convention BIM non exploitable) avant de poursuivre. L'analyse est bloquée.</div>` : `
          <label class="precheck-check">
            <input type="checkbox" id="precheckScopeReview">
            <span>Je valide le <strong>nom du projet</strong>, le <strong>lot analysé</strong> et les <strong>documents chargés</strong> affichés ci-dessus.</span>
          </label>
          ${requireReview ? `
            <label class="precheck-check">
              <input type="checkbox" id="precheckReview">
              <span>J'ai vérifié que les réponses de l'auto-évaluation sont complètes et à jour pour cette entreprise et cet appel d'offres.</span>
            </label>` : ""}
          ${hasAlert ? `
            <label class="precheck-check precheck-check-warning">
              <input type="checkbox" id="precheckOverride">
              <span>Je confirme explicitement avoir vérifié l'incohérence ou l'incertitude signalée et vouloir poursuivre malgré cette alerte.</span>
            </label>` : ""}`}
        </div>
        <div class="precheck-footer">
          <div class="precheck-confirmations" aria-label="Validations obligatoires"></div>
          <div class="precheck-actions">
            <button type="button" class="precheck-cancel secondary-button">Retour aux documents</button>
            <button type="button" class="precheck-launch">${hasBlocking ? "Analyse bloquée" : "Valider et lancer l'analyse"}</button>
          </div>
        </div>
      </section>`;

    // DEV08 : déplacer les validations dans un pied fixe et verrouiller le scroll de la page arrière.
    const confirmations = overlay.querySelector(".precheck-confirmations");
    overlay.querySelectorAll(".precheck-body > .precheck-check, .precheck-body > .precheck-blocking-message").forEach(node => confirmations?.appendChild(node));
    const previousBodyOverflow = document.body.style.overflow;
    const previousBodyPaddingRight = document.body.style.paddingRight;
    const scrollbarGap = Math.max(0, window.innerWidth - document.documentElement.clientWidth);
    if (scrollbarGap) document.body.style.paddingRight = `${scrollbarGap}px`;
    document.body.style.overflow = "hidden";
    document.body.classList.add("precheck-open");
    document.body.appendChild(overlay);
    const launch = overlay.querySelector(".precheck-launch");
    const cancel = overlay.querySelector(".precheck-cancel");
    const scopeReview = overlay.querySelector("#precheckScopeReview");
    const review = overlay.querySelector("#precheckReview");
    const override = overlay.querySelector("#precheckOverride");

    const update = () => {
      if (hasBlocking) { launch.disabled = true; return; }
      const scopeOk = Boolean(scopeReview?.checked);
      const reviewOk = !requireReview || Boolean(review?.checked);
      const alertOk = !hasAlert || Boolean(override?.checked);
      launch.disabled = !(scopeOk && reviewOk && alertOk);
    };
    update();
    scopeReview?.addEventListener("change", update);
    review?.addEventListener("change", update);
    override?.addEventListener("change", update);

    let closed = false;
    const finish = value => {
      if (closed) return;
      closed = true;
      document.removeEventListener("keydown", onKey);
      overlay.remove();
      document.body.classList.remove("precheck-open");
      document.body.style.overflow = previousBodyOverflow;
      document.body.style.paddingRight = previousBodyPaddingRight;
      resolve(value);
    };
    const onKey = event => { if (event.key === "Escape") finish(null); };
    document.addEventListener("keydown", onKey);
    cancel.addEventListener("click", () => finish(null));
    launch.addEventListener("click", () => { if (!launch.disabled) finish({accepted: true, override: Boolean(override?.checked)}); });
    overlay.addEventListener("click", event => { if (event.target === overlay) finish(null); });
    (scopeReview || review || override || cancel).focus();
  });
}

async function confirmPreAnalysisCoherence(data, answersCount) {
  const convention = data.get("convention");
  const cctp = data.get("cctp");
  const ccap = data.get("ccap");
  const evaluationText = sansEval.checked
    ? "Auto-évaluation : test rapide sans auto-évaluation."
    : `Auto-évaluation : ${answersCount}/${questionIds.length} réponses renseignées.`;
  const requireReview = !sansEval.checked;

  if (!convention || !convention.name) {
    setLivePrecheck("danger", "Convention BIM absente : analyse impossible.");
    return null;
  }

  const checkData = new FormData();
  checkData.set("convention", convention, convention.name);
  if (cctp && cctp.name) checkData.set("cctp", cctp, cctp.name);
  if (ccap && ccap.name) checkData.set("ccap", ccap, ccap.name);
  checkData.set("projet", String(data.get("projet") || ""));
  checkData.set("lot", String(data.get("lot") || ""));
  try {
    const response = await apiFetch(`${API}/api/precheck`, {method: "POST", body: checkData});
    const check = await readJsonResponse(response);
    if (!response.ok) throw new Error(check.error || `Vérification impossible (HTTP ${response.status})`);
    lastLivePrecheck = check;
    const summary = summarizePrecheck(check);
    setLivePrecheck(summary.state, summary.text);
    const decision = await showPreAnalysisModal(check, evaluationText, {requireReview});
    if (!decision) return null;
    return {...check, user_override: Boolean(decision.override)};
  } catch (error) {
    setLivePrecheck("danger", `Vérification préalable impossible : ${error.message}. Analyse bloquée.`);
    showProgress(`<h2>Contrôle préalable impossible</h2><p>${escapeHtml(error.message)}</p><p>L'analyse n'est pas lancée. Corrigez le dossier ou rétablissez le service de vérification.</p>`, true);
    return null;
  }
}

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
      showProgress("<h2>Vérification du questionnaire</h2><p>Nouvelle tentative de connexion au service d’analyse…</p>");
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
  const precheckValidated = await confirmPreAnalysisCoherence(data, Object.keys(answers).length);
  if (!precheckValidated) {
    showProgress("<h2>Analyse non lancée</h2><p>Vérifiez le type des pièces, le nom du projet, le lot, la cohérence des documents et les réponses d’auto-évaluation, puis relancez l’analyse.</p>");
    return;
  }
  data.set("precheck_fingerprint", String(precheckValidated.fingerprint || ""));
  data.set("precheck_ack", "true");
  data.set("precheck_override", precheckValidated.user_override ? "true" : "false");
  submitButton.disabled = true;
  autoevalButton.disabled = true;
  submitButton.textContent = "Analyse en cours…";
  renderLiveProgress({progress: 2, stage: "Envoi des documents", detail: "Transmission sécurisée des fichiers au service d’analyse."});

  try {
    const response = await apiFetch(`${API}/api/analyse`, {method: "POST", body: data});
    const queued = await readJsonResponse(response);
    if (!response.ok) throw new Error(queued.error || `Échec du lancement (HTTP ${response.status})`);
    saveFormState();
    renderLiveProgress({progress: 5, stage: "Analyse enregistrée", detail: "Les fichiers ont été reçus par le service d’analyse."});
    const result = await pollAnalysis(queued.status_url);
    renderCompletedAnalysis(result);
    await loadHistory();
  } catch (error) {
    const networkHint = error instanceof TypeError
      ? `<p>Le navigateur n’a pas réussi à joindre <code>${escapeHtml(API)}</code>. Vérifiez que le service d’analyse est accessible.</p>`
      : "";
    showProgress(`<h2>Analyse interrompue</h2><p>${escapeHtml(error.message)}</p>${networkHint}<p>Veuillez réessayer. Si le problème persiste, contactez l’administrateur de l’outil.</p>`, true);
  } finally {
    submitButton.disabled = false;
    autoevalButton.disabled = false;
    submitButton.textContent = "Analyser les documents";
  }
});

async function initialize() {
  restoreBaseFields();
  await Promise.all([
    restoreStoredFile(conventionInput),
    restoreStoredFile(cctpInput),
    restoreStoredFile(ccapInput),
  ]);
  const restoredConvention = conventionInput.files && conventionInput.files[0];
  if (restoredConvention) await detectProjectFromConvention(restoredConvention, {resetFirst: true});
  await Promise.all([loadQuestions(), loadHistory()]);
}

initialize();
