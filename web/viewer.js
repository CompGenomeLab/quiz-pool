const state = {
  examSets: [],
  examStorePath: "",
  selectedExamSetId: "",
  selectedExamSet: null,
  selectedVariantId: "",
  isDeletingExamSet: false,
  isDownloadingPrintableZip: false,
  statusIsError: false,
  statusMessage: "Loading exam sets...",
  validationErrors: [],
};

const elements = {
  examSetList: document.querySelector("#exam-set-list"),
  examSetSelect: document.querySelector("#exam-set-select"),
  examStorePath: document.querySelector("#exam-store-path"),
  errorList: document.querySelector("#viewer-error-list"),
  errorPanel: document.querySelector("#viewer-errors"),
  deleteExam: document.querySelector("#viewer-delete-exam"),
  printResults: document.querySelector("#viewer-print-results"),
  results: document.querySelector("#viewer-results"),
  variantList: document.querySelector("#viewer-variant-list"),
  variantMenu: document.querySelector("#viewer-variant-menu"),
  variantSelect: document.querySelector("#viewer-variant-select"),
  viewerExamName: document.querySelector("#viewer-exam-name"),
  viewerExamSetId: document.querySelector("#viewer-exam-set-id"),
  viewerGeneratedAt: document.querySelector("#viewer-generated-at"),
  viewerHeading: document.querySelector("#viewer-heading"),
  viewerQuestionCount: document.querySelector("#viewer-question-count"),
  viewerQuestionPoolBody: document.querySelector("#viewer-question-pool-body"),
  viewerQuizTitle: document.querySelector("#viewer-quiz-title"),
  viewerStatus: document.querySelector("#viewer-status"),
  viewerVariantCount: document.querySelector("#viewer-variant-count"),
};

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  elements.viewerStatus.textContent = message;
  elements.viewerStatus.style.color = isError ? "var(--danger-strong)" : "var(--muted)";
}

function truncate(text, maxLength = 120) {
  if (text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, maxLength - 1)}…`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderErrors() {
  if (state.validationErrors.length === 0) {
    elements.errorPanel.classList.add("hidden");
    elements.errorList.replaceChildren();
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const error of state.validationErrors) {
    const item = document.createElement("li");
    item.textContent = `${error.path}: ${error.message}`;
    fragment.append(item);
  }
  elements.errorList.replaceChildren(fragment);
  elements.errorPanel.classList.remove("hidden");
}

function examSetLabel(summary) {
  const examName = summary.printSettings?.examName || summary.quiz?.title || "Untitled Exam";
  const when = summary.generatedAt ? new Date(summary.generatedAt).toLocaleString() : "Unknown time";
  return `${examName} · ${when}`;
}

function renderExamSetOptions() {
  const fragment = document.createDocumentFragment();
  for (const summary of state.examSets) {
    const option = document.createElement("option");
    option.value = summary.examSetId;
    option.textContent = examSetLabel(summary);
    option.selected = summary.examSetId === state.selectedExamSetId;
    fragment.append(option);
  }
  elements.examSetSelect.replaceChildren(fragment);
}

function renderExamSetList() {
  const fragment = document.createDocumentFragment();
  for (const summary of state.examSets) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "question-tile";
    if (summary.examSetId === state.selectedExamSetId) {
      button.classList.add("is-selected");
    }

    const id = document.createElement("span");
    id.className = "question-tile__id";
    id.textContent = summary.examSetId;

    const text = document.createElement("span");
    text.className = "question-tile__text";
    text.textContent = examSetLabel(summary);

    button.append(id, text);
    button.addEventListener("click", async () => {
      elements.examSetSelect.value = summary.examSetId;
      await loadExamSet(summary.examSetId);
    });
    fragment.append(button);
  }
  elements.examSetList.replaceChildren(fragment);
}

function renderQuestionPool(questionPool) {
  const fragment = document.createDocumentFragment();
  for (const question of questionPool) {
    const row = document.createElement("tr");
    for (const value of [
      question.sourceQuestionId,
      truncate(question.question),
      String(question.points ?? 1),
      question.choices.map((choice) => `${choice.key}. ${choice.text}`).join(" | "),
      question.sourceCorrectAnswers.join(", "),
    ]) {
      const cell = document.createElement("td");
      cell.textContent = value || "—";
      row.append(cell);
    }
    fragment.append(row);
  }
  elements.viewerQuestionPoolBody.replaceChildren(fragment);
}

function renderVariants(variants, printSettings) {
  const selectedVariant = variants.find((variant) => variant.variantId === state.selectedVariantId) ?? variants[0] ?? null;
  const fragment = document.createDocumentFragment();
  if (!selectedVariant) {
    elements.variantList.replaceChildren();
    return;
  }
  {
    const variant = selectedVariant;
    const card = document.createElement("article");
    card.className = "variant-card";

    const header = document.createElement("div");
    header.className = "variant-card__header";

    const headingBlock = document.createElement("div");
    const eyebrow = document.createElement("p");
    eyebrow.className = "eyebrow";
    eyebrow.textContent = "Saved Variant";
    const title = document.createElement("h3");
    title.textContent = `${printSettings.examName || "Generated Exam"} · ${variant.printableFileName || variant.variantId}`;
    const meta = document.createElement("p");
    meta.className = "variant-card__meta";
    meta.textContent = `${variant.variantId} · ${variant.questions.length} question(s) · ${variant.questions.reduce((total, question) => total + Number(question.points ?? 1), 0)} pts`;
    headingBlock.append(eyebrow, title, meta);

    const qr = document.createElement("aside");
    qr.className = "variant-card__qr";
    const qrImage = document.createElement("img");
    qrImage.className = "variant-card__qr-image";
    qrImage.alt = "Variant tracking QR code";
    qrImage.loading = "lazy";
    qrImage.src = `/api/exams/variant-qr/${encodeURIComponent(variant.variantId)}.svg`;
    qr.append(qrImage);

    header.append(headingBlock, qr);

    const questionList = document.createElement("div");
    questionList.className = "question-preview-list";

    for (const question of variant.questions) {
      const section = document.createElement("section");
      section.className = "question-preview";

      const questionHead = document.createElement("div");
      questionHead.className = "question-preview__head";
      questionHead.textContent = `Question ${question.position} · ${question.sourceQuestionId} · ${question.points ?? 1} pt`;

      const prompt = document.createElement("p");
      prompt.className = "question-preview__title";
      prompt.textContent = question.question;

      const choices = document.createElement("ul");
      choices.className = "choice-list";
      for (const choice of question.displayChoices) {
        const item = document.createElement("li");
        const pill = document.createElement("span");
        pill.className = "choice-pill";
        pill.textContent = `${choice.key}. ${choice.text}`;
        item.append(pill);
        choices.append(item);
      }

      const correct = document.createElement("p");
      correct.className = "helper-copy";
      correct.textContent = `Displayed correct: ${question.displayCorrectAnswers.join(", ") || "—"}`;

      section.append(questionHead, prompt, choices, correct);
      questionList.append(section);
    }

    card.append(header, questionList);
    fragment.append(card);
  }
  elements.variantList.replaceChildren(fragment);
}

function renderVariantOptions(variants) {
  const fragment = document.createDocumentFragment();
  for (const variant of variants) {
    const option = document.createElement("option");
    option.value = variant.variantId;
    option.textContent = variant.printableFileName || variant.variantId;
    option.selected = variant.variantId === state.selectedVariantId;
    fragment.append(option);
  }
  elements.variantSelect.replaceChildren(fragment);
}

function renderVariantMenu(variants) {
  const fragment = document.createDocumentFragment();
  for (const variant of variants) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "question-tile";
    if (variant.variantId === state.selectedVariantId) {
      button.classList.add("is-selected");
    }

    const id = document.createElement("span");
    id.className = "question-tile__id";
    id.textContent = variant.printableFileName || "Variant";

    const text = document.createElement("span");
    text.className = "question-tile__text";
    text.textContent = `${variant.variantId} · ${variant.questions.length} question(s)`;

    button.append(id, text);
    button.addEventListener("click", () => {
      state.selectedVariantId = variant.variantId;
      renderSelectedExamSet();
    });
    fragment.append(button);
  }
  elements.variantMenu.replaceChildren(fragment);
}

function renderSelectedExamSet() {
  const hasExamSet = Boolean(state.selectedExamSet);
  elements.results.classList.toggle("hidden", !hasExamSet);
  elements.printResults.disabled = !hasExamSet || state.isDownloadingPrintableZip || state.isDeletingExamSet;
  elements.deleteExam.disabled = !hasExamSet || state.isDownloadingPrintableZip || state.isDeletingExamSet;
  if (!hasExamSet) {
    elements.viewerQuestionPoolBody.replaceChildren();
    elements.variantMenu.replaceChildren();
    elements.variantSelect.replaceChildren();
    elements.variantList.replaceChildren();
    return;
  }

  const { summary, examSet } = state.selectedExamSet;
  elements.viewerHeading.textContent = summary.printSettings.examName || summary.quiz.title || "Saved Exam Set";
  elements.viewerExamSetId.textContent = examSet.examSetId;
  elements.viewerGeneratedAt.textContent = new Date(examSet.generatedAt).toLocaleString();
  elements.viewerExamName.textContent = summary.printSettings.examName || "—";
  elements.viewerVariantCount.textContent = String(summary.variantCount);
  elements.viewerQuestionCount.textContent = String(summary.selectedQuestionCount);
  elements.viewerQuizTitle.textContent = summary.quiz.title || "—";

  renderQuestionPool(examSet.questionPool);
  renderVariantOptions(examSet.variants);
  renderVariantMenu(examSet.variants);
  renderVariants(examSet.variants, examSet.printSettings);
}

async function downloadPrintableZip() {
  if (!state.selectedExamSet) {
    return;
  }

  const examSetId = state.selectedExamSet.examSet.examSetId;
  state.isDownloadingPrintableZip = true;
  renderSelectedExamSet();
  setStatus("Preparing printable ZIP...");

  const response = await fetch(`/api/exams/export/${encodeURIComponent(examSetId)}.zip`);
  if (!response.ok) {
    let message = `Printable export failed (${response.status})`;
    try {
      const payload = await response.json();
      message = payload.errors?.[0]?.message ?? message;
    } catch {
      // Keep the fallback message if the response is not JSON.
    }
    state.isDownloadingPrintableZip = false;
    renderSelectedExamSet();
    setStatus(message, true);
    return;
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `exam-set-${examSetId}-printables.zip`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
  state.isDownloadingPrintableZip = false;
  renderSelectedExamSet();
  setStatus(`Downloaded printable ZIP for exam set ${examSetId}.`);
}

async function deleteSelectedExamSet() {
  if (!state.selectedExamSet) {
    return;
  }

  const examSetId = state.selectedExamSet.examSet.examSetId;
  const examName = state.selectedExamSet.summary.printSettings.examName || state.selectedExamSet.summary.quiz.title || examSetId;
  const confirmed = window.confirm(`Delete exam set "${examName}" (${examSetId}) from the exam store?`);
  if (!confirmed) {
    return;
  }

  state.isDeletingExamSet = true;
  renderSelectedExamSet();
  setStatus(`Deleting exam set ${examSetId}...`);

  try {
    const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}`, {
      method: "DELETE",
    });
    let payload = null;
    try {
      payload = await response.json();
    } catch {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.errors?.[0]?.message ?? `Could not delete exam set (${response.status})`;
      setStatus(message, true);
      return;
    }

    state.selectedExamSet = null;
    state.selectedExamSetId = "";
    state.selectedVariantId = "";
    await loadExamSets();
    setStatus(`Deleted exam set ${examSetId}.`);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    state.isDeletingExamSet = false;
    renderSelectedExamSet();
  }
}

async function loadExamSets() {
  setStatus("Loading exam sets...");
  const response = await fetch("/api/exams");
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not load exam sets (${response.status})`);
  }

  state.examStorePath = payload.examStorePath;
  state.examSets = Array.isArray(payload.examSets) ? payload.examSets : [];
  elements.examStorePath.textContent = state.examStorePath;
  if (state.examSets.length === 0) {
    state.selectedExamSetId = "";
    state.selectedExamSet = null;
    renderExamSetOptions();
    renderExamSetList();
    renderSelectedExamSet();
    setStatus("No generated exam sets were found.");
    return;
  }

  state.selectedExamSetId = state.examSets[0].examSetId;
  renderExamSetOptions();
  renderExamSetList();
  await loadExamSet(state.selectedExamSetId);
}

async function loadExamSet(examSetId) {
  state.selectedExamSetId = examSetId;
  renderExamSetOptions();
  renderExamSetList();
  setStatus("Loading exam set...");

  const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}`);
  const payload = await response.json();
  if (!response.ok) {
    state.selectedExamSet = null;
    state.validationErrors = payload.errors ?? [{ path: "<unknown>", message: "Could not load exam set" }];
    renderErrors();
    renderSelectedExamSet();
    setStatus("Exam set load failed.", true);
    return;
  }

  state.selectedExamSet = payload;
  state.selectedVariantId = payload.examSet.variants[0]?.variantId ?? "";
  state.validationErrors = [];
  renderErrors();
  renderSelectedExamSet();
  setStatus(`Loaded exam set ${examSetId}.`);
}

function wireEvents() {
  elements.examSetSelect.addEventListener("change", async (event) => {
    await loadExamSet(event.target.value);
  });
  elements.deleteExam.addEventListener("click", () => {
    void deleteSelectedExamSet();
  });
  elements.printResults.addEventListener("click", () => {
    void downloadPrintableZip();
  });
  elements.variantSelect.addEventListener("change", (event) => {
    state.selectedVariantId = event.target.value;
    renderSelectedExamSet();
  });
}

wireEvents();
loadExamSets().catch((error) => {
  console.error(error);
  state.validationErrors = [{ path: "<viewer>", message: error.message }];
  renderErrors();
  setStatus(error.message, true);
});
