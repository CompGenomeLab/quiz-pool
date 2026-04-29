import { hasRichTextMarkup, renderRichTextIntoElement, stripRichTextMarkup } from "./rich-text.js";

const state = {
  examSets: [],
  examStorePath: "",
  isQuestionPoolCollapsed: true,
  isVariantsCollapsed: true,
  selectedExamSetId: "",
  selectedExamSet: null,
  selectedVariantId: "",
  isDeletingExamSet: false,
  isDownloadingPrintableZip: false,
  isUpdatingPrintSettings: false,
  statusIsError: false,
  statusMessage: "Loading exam sets...",
  validationErrors: [],
};

const DEFAULT_EXAM_RULES = [
  "Fill bubbles fully. Complete all ID columns with leading zeros (e.g., 00012345).",
  "Read every question carefully and select all correct answers for each question.",
  "Mark answers clearly and keep your paper neat for printing, photocopying, and scanning.",
  "Do not communicate with other students or use unauthorized materials during the exam.",
  "Remain seated until instructed to stop and submit your paper.",
];
const DEFAULT_OMR_INSTRUCTIONS = DEFAULT_EXAM_RULES[0];

function defaultExamRules(omrInstructions = DEFAULT_OMR_INSTRUCTIONS) {
  return [
    omrInstructions,
    ...DEFAULT_EXAM_RULES.slice(1),
  ];
}

const elements = {
  examSetList: document.querySelector("#exam-set-list"),
  examSetSelect: document.querySelector("#exam-set-select"),
  examStorePath: document.querySelector("#exam-store-path"),
  errorList: document.querySelector("#viewer-error-list"),
  errorPanel: document.querySelector("#viewer-errors"),
  deleteExam: document.querySelector("#viewer-delete-exam"),
  institutionName: document.querySelector("#viewer-institution-name"),
  courseName: document.querySelector("#viewer-course-name"),
  examDate: document.querySelector("#viewer-exam-date"),
  printExamName: document.querySelector("#viewer-print-exam-name"),
  omrInstructions: document.querySelector("#viewer-omr-instructions"),
  examRules: document.querySelector("#viewer-exam-rules"),
  examRulesPreview: document.querySelector("#viewer-exam-rules-preview"),
  printResults: document.querySelector("#viewer-print-results"),
  updatePrintSettings: document.querySelector("#viewer-update-print-settings"),
  startTime: document.querySelector("#viewer-start-time"),
  totalTimeMinutes: document.querySelector("#viewer-total-time-minutes"),
  instructor: document.querySelector("#viewer-instructor"),
  allowedMaterials: document.querySelector("#viewer-allowed-materials"),
  results: document.querySelector("#viewer-results"),
  toggleQuestionPool: document.querySelector("#toggle-question-pool"),
  toggleVariants: document.querySelector("#toggle-variants"),
  variantList: document.querySelector("#viewer-variant-list"),
  variantMenu: document.querySelector("#viewer-variant-menu"),
  variantsSection: document.querySelector("#viewer-variants-section"),
  variantSelect: document.querySelector("#viewer-variant-select"),
  viewerExamName: document.querySelector("#viewer-exam-name"),
  viewerExamSetId: document.querySelector("#viewer-exam-set-id"),
  viewerGeneratedAt: document.querySelector("#viewer-generated-at"),
  viewerHeading: document.querySelector("#viewer-heading"),
  viewerQuestionCount: document.querySelector("#viewer-question-count"),
  viewerQuestionPoolBody: document.querySelector("#viewer-question-pool-body"),
  viewerQuestionPoolSection: document.querySelector("#viewer-question-pool-section"),
  viewerQuizTitle: document.querySelector("#viewer-quiz-title"),
  viewerStatus: document.querySelector("#viewer-status"),
  viewerVariantCount: document.querySelector("#viewer-variant-count"),
};

const printSettingsElements = [
  elements.institutionName,
  elements.printExamName,
  elements.courseName,
  elements.examDate,
  elements.startTime,
  elements.totalTimeMinutes,
  elements.instructor,
  elements.allowedMaterials,
  elements.omrInstructions,
  elements.examRules,
];

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

function normalizeTextValue(value, fallback = "") {
  return typeof value === "string" && value.trim() !== "" ? value.trim() : fallback;
}

function parseExamRules(value) {
  return value
    .split(/\r?\n/u)
    .map((line) => line.trim())
    .filter(Boolean);
}

function normalizeExamRules(value) {
  if (Array.isArray(value)) {
    const rules = value.map((rule) => normalizeTextValue(rule)).filter(Boolean);
    if (rules.length > 0) {
      return rules;
    }
  }
  if (typeof value === "string" && value.trim() !== "") {
    const rules = parseExamRules(value);
    if (rules.length > 0) {
      return rules;
    }
  }
  return defaultExamRules();
}

function renderExamRulesPreview() {
  const rules = parseExamRules(elements.examRules.value);
  const hasMathRule = rules.some((rule) => hasRichTextMarkup(rule));
  elements.examRulesPreview.classList.toggle("hidden", !hasMathRule);
  elements.examRulesPreview.replaceChildren();
  if (!hasMathRule) {
    return;
  }

  const list = document.createElement("ol");
  for (const rule of rules) {
    const item = document.createElement("li");
    renderRichTextIntoElement(item, rule);
    list.append(item);
  }
  elements.examRulesPreview.append(list);
}

function printableMetadata() {
  const totalTimeRaw = elements.totalTimeMinutes.value.trim();
  const parsedTotalTime = totalTimeRaw === "" ? null : Number.parseInt(totalTimeRaw, 10);
  return {
    institutionName: elements.institutionName.value.trim(),
    examName: elements.printExamName.value.trim(),
    courseName: elements.courseName.value.trim(),
    examDate: elements.examDate.value.trim(),
    startTime: elements.startTime.value.trim(),
    totalTimeMinutes: Number.isFinite(parsedTotalTime) ? parsedTotalTime : null,
    instructor: elements.instructor.value.trim(),
    allowedMaterials: elements.allowedMaterials.value.trim(),
    omrInstructions: elements.omrInstructions.value.trim(),
    examRules: parseExamRules(elements.examRules.value),
  };
}

function normalizeTotalTimeForComparison(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  const text = String(value).trim();
  if (!text) {
    return "";
  }
  const numeric = Number.parseInt(text, 10);
  return Number.isFinite(numeric) ? String(numeric) : text;
}

function normalizeRulesForComparison(value, useDefaultRules = false) {
  if (useDefaultRules) {
    return normalizeExamRules(value);
  }
  if (Array.isArray(value)) {
    return value.map((rule) => normalizeTextValue(rule)).filter(Boolean);
  }
  if (typeof value === "string") {
    return parseExamRules(value);
  }
  return [];
}

function normalizePrintSettingsForComparison(printSettings = {}, useDefaultRules = false) {
  return {
    institutionName: normalizeTextValue(printSettings.institutionName, ""),
    examName: normalizeTextValue(printSettings.examName, ""),
    courseName: normalizeTextValue(printSettings.courseName, ""),
    examDate: normalizeTextValue(printSettings.examDate, ""),
    startTime: normalizeTextValue(printSettings.startTime, ""),
    totalTimeMinutes: normalizeTotalTimeForComparison(printSettings.totalTimeMinutes),
    instructor: normalizeTextValue(printSettings.instructor, ""),
    allowedMaterials: normalizeTextValue(printSettings.allowedMaterials, ""),
    omrInstructions: normalizeTextValue(printSettings.omrInstructions, ""),
    examRules: normalizeRulesForComparison(printSettings.examRules, useDefaultRules),
  };
}

function printSettingsHaveUnsavedChanges() {
  if (!state.selectedExamSet) {
    return false;
  }
  const current = normalizePrintSettingsForComparison(printableMetadata(), false);
  const saved = normalizePrintSettingsForComparison(
    state.selectedExamSet.examSet.printSettings,
    true,
  );
  return JSON.stringify(current) !== JSON.stringify(saved);
}

function populatePrintSettingsForm(printSettings = {}) {
  elements.institutionName.value = normalizeTextValue(printSettings.institutionName, "");
  elements.printExamName.value = normalizeTextValue(printSettings.examName, "");
  elements.courseName.value = normalizeTextValue(printSettings.courseName, "");
  elements.examDate.value = normalizeTextValue(printSettings.examDate, "");
  elements.startTime.value = normalizeTextValue(printSettings.startTime, "");
  elements.totalTimeMinutes.value = String(printSettings.totalTimeMinutes ?? "").trim();
  elements.instructor.value = normalizeTextValue(printSettings.instructor, "");
  elements.allowedMaterials.value = normalizeTextValue(printSettings.allowedMaterials, "");
  elements.omrInstructions.value = normalizeTextValue(printSettings.omrInstructions, DEFAULT_OMR_INSTRUCTIONS);
  elements.examRules.value = normalizeExamRules(printSettings.examRules).join("\n");
  renderExamRulesPreview();
}

function setPrintSettingsFormDisabled(disabled) {
  for (const element of printSettingsElements) {
    element.disabled = disabled;
  }
}

function assetUrl(assetId) {
  return `/api/assets/${encodeURIComponent(assetId)}`;
}

function questionImageAssetIds(question = {}) {
  return Array.isArray(question.imageAssetIds)
    ? question.imageAssetIds.filter((assetId) => typeof assetId === "string" && assetId.trim() !== "")
    : [];
}

function createQuestionImagePreviews(question) {
  const imageAssetIds = questionImageAssetIds(question);
  if (imageAssetIds.length === 0) {
    return null;
  }
  const wrap = document.createElement("div");
  wrap.className = "question-image-list";
  imageAssetIds.forEach((assetId, index) => {
    const image = document.createElement("img");
    image.className = "question-image-preview";
    image.src = assetUrl(assetId);
    image.alt = `Question image ${index + 1}`;
    image.loading = "lazy";
    wrap.append(image);
  });
  return wrap;
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
    const idCell = document.createElement("td");
    idCell.textContent = question.sourceQuestionId || "—";

    const questionCell = document.createElement("td");
    questionCell.className = "cell-copy";
    renderRichTextIntoElement(questionCell, question.question || "—");
    const poolImages = createQuestionImagePreviews(question);
    if (poolImages) {
      questionCell.append(poolImages);
    }

    const pointsCell = document.createElement("td");
    pointsCell.textContent = String(question.points ?? 1);

    const choicesCell = document.createElement("td");
    choicesCell.textContent = question.choices.map((choice) => `${choice.key}. ${stripRichTextMarkup(choice.text)}`).join(" | ") || "—";

    const correctCell = document.createElement("td");
    correctCell.textContent = question.sourceCorrectAnswers.join(", ") || "—";

    row.append(idCell, questionCell, pointsCell, choicesCell, correctCell);
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
      renderRichTextIntoElement(prompt, question.question);
      const imagePreviews = createQuestionImagePreviews(question);

      const choices = document.createElement("ul");
      choices.className = "choice-list";
      for (const choice of question.displayChoices) {
        const item = document.createElement("li");
        const pill = document.createElement("span");
        pill.className = "choice-pill";
        const key = document.createElement("span");
        key.className = "choice-pill__key";
        key.textContent = `${choice.key}.`;
        const text = document.createElement("span");
        text.className = "choice-pill__text";
        renderRichTextIntoElement(text, choice.text);
        pill.append(key, text);
        item.append(pill);
        choices.append(item);
      }

      const correct = document.createElement("p");
      correct.className = "helper-copy";
      correct.textContent = `Displayed correct: ${question.displayCorrectAnswers.join(", ") || "—"}`;

      section.append(questionHead, prompt);
      if (imagePreviews) {
        section.append(imagePreviews);
      }
      section.append(choices, correct);
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

function renderSectionToggles() {
  elements.viewerQuestionPoolSection.classList.toggle("hidden", state.isQuestionPoolCollapsed);
  elements.variantsSection.classList.toggle("hidden", state.isVariantsCollapsed);
  elements.toggleQuestionPool.textContent = state.isQuestionPoolCollapsed ? "Expand" : "Collapse";
  elements.toggleVariants.textContent = state.isVariantsCollapsed ? "Expand" : "Collapse";
  elements.toggleQuestionPool.setAttribute("aria-expanded", String(!state.isQuestionPoolCollapsed));
  elements.toggleVariants.setAttribute("aria-expanded", String(!state.isVariantsCollapsed));
}

function renderSelectedExamSet() {
  const hasExamSet = Boolean(state.selectedExamSet);
  elements.results.classList.toggle("hidden", !hasExamSet);
  const isBusy = state.isDownloadingPrintableZip || state.isDeletingExamSet || state.isUpdatingPrintSettings;
  elements.printResults.disabled = !hasExamSet || isBusy;
  elements.deleteExam.disabled = !hasExamSet || isBusy;
  elements.updatePrintSettings.disabled = !hasExamSet || isBusy;
  setPrintSettingsFormDisabled(!hasExamSet || isBusy);
  if (!hasExamSet) {
    elements.viewerQuestionPoolBody.replaceChildren();
    elements.variantMenu.replaceChildren();
    elements.variantSelect.replaceChildren();
    elements.variantList.replaceChildren();
    populatePrintSettingsForm();
    renderSectionToggles();
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
  renderSectionToggles();
}

async function downloadPrintableZip() {
  if (!state.selectedExamSet) {
    return;
  }
  if (printSettingsHaveUnsavedChanges()) {
    setStatus("Printable metadata has unsaved changes. Apply Metadata Updates before downloading printables.", true);
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

async function updateSelectedPrintSettings() {
  if (!state.selectedExamSet) {
    return;
  }

  const examSetId = state.selectedExamSet.examSet.examSetId;
  state.isUpdatingPrintSettings = true;
  renderSelectedExamSet();
  setStatus("Updating printable metadata...");

  try {
    const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}/print-settings`, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(printableMetadata()),
    });
    const payload = await response.json();
    if (!response.ok) {
      state.validationErrors = payload.errors ?? [{ path: "<metadata>", message: "Could not update printable metadata" }];
      renderErrors();
      setStatus("Metadata update failed. Review the validation messages.", true);
      return;
    }

    state.selectedExamSet.summary = payload.summary;
    state.selectedExamSet.examSet.printSettings = payload.printSettings;
    state.examSets = state.examSets.map((summary) => (
      summary.examSetId === examSetId ? payload.summary : summary
    ));
    state.validationErrors = [];
    renderErrors();
    populatePrintSettingsForm(payload.printSettings);
    renderExamSetOptions();
    renderExamSetList();
    renderSelectedExamSet();
    setStatus(`Updated printable metadata for exam set ${examSetId}.`);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    state.isUpdatingPrintSettings = false;
    renderSelectedExamSet();
  }
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
  populatePrintSettingsForm(payload.examSet.printSettings);
  renderSelectedExamSet();
  setStatus(`Loaded exam set ${examSetId}.`);
}

function wireEvents() {
  elements.examSetSelect.addEventListener("change", async (event) => {
    await loadExamSet(event.target.value);
  });
  elements.toggleQuestionPool.addEventListener("click", () => {
    state.isQuestionPoolCollapsed = !state.isQuestionPoolCollapsed;
    renderSectionToggles();
  });
  elements.toggleVariants.addEventListener("click", () => {
    state.isVariantsCollapsed = !state.isVariantsCollapsed;
    renderSectionToggles();
  });
  elements.deleteExam.addEventListener("click", () => {
    void deleteSelectedExamSet();
  });
  elements.printResults.addEventListener("click", () => {
    void downloadPrintableZip();
  });
  elements.updatePrintSettings.addEventListener("click", () => {
    void updateSelectedPrintSettings();
  });
  for (const element of printSettingsElements) {
    element.addEventListener("input", () => {
      if (element === elements.examRules) {
        renderExamRulesPreview();
      }
      if (state.selectedExamSet) {
        setStatus("Printable metadata changed. Apply updates before downloading printables.");
      }
    });
  }
  elements.omrInstructions.addEventListener("input", () => {
    const currentRules = parseExamRules(elements.examRules.value);
    const omrInstructions = elements.omrInstructions.value.trim() || DEFAULT_OMR_INSTRUCTIONS;
    if (
      currentRules.length === 0
      || JSON.stringify(currentRules) === JSON.stringify(defaultExamRules())
      || JSON.stringify(currentRules) === JSON.stringify(defaultExamRules(currentRules[0] || DEFAULT_OMR_INSTRUCTIONS))
    ) {
      elements.examRules.value = defaultExamRules(omrInstructions).join("\n");
      renderExamRulesPreview();
    }
  });
  elements.variantSelect.addEventListener("change", (event) => {
    state.selectedVariantId = event.target.value;
    renderSelectedExamSet();
  });
}

wireEvents();
renderSectionToggles();
loadExamSets().catch((error) => {
  console.error(error);
  state.validationErrors = [{ path: "<viewer>", message: error.message }];
  renderErrors();
  setStatus(error.message, true);
});
