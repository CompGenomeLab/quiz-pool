import { renderRichTextHtml, renderRichTextIntoElement, stripRichTextMarkup } from "./rich-text.js";

const state = {
  activePoolQuestionId: "",
  dbPath: "",
  examStorePath: "",
  generatedRun: null,
  quiz: null,
  statusSortDirection: "default",
  statusIsError: false,
  statusMessage: "Loading quiz data...",
  validationErrors: [],
  selection: {
    questionCount: 1,
    variantCount: 1,
    sources: [],
    difficulties: [],
    learningObjectiveIds: [],
    overrides: {},
  },
};

const DEFAULT_EXAM_RULES = [
  "Fill bubbles fully. Complete all ID columns with leading zeros (e.g., 00012345).",
  "Read every question carefully and select all correct answers for each question.",
  "Mark answers clearly and keep your paper neat for printing, photocopying, and scanning.",
  "Do not communicate with other students or use unauthorized materials during the exam.",
  "Remain seated until instructed to stop and submit your paper.",
];
const DEFAULT_OMR_INSTRUCTIONS = DEFAULT_EXAM_RULES[0];
const MAX_QUESTIONS_PER_EXAM = 100;

function defaultExamRules(omrInstructions = DEFAULT_OMR_INSTRUCTIONS) {
  return [
    omrInstructions,
    ...DEFAULT_EXAM_RULES.slice(1),
  ];
}

const elements = {
  availableCount: document.querySelector("#available-count"),
  sourceFilters: document.querySelector("#source-filters"),
  courseName: document.querySelector("#course-name"),
  dbPath: document.querySelector("#db-path"),
  difficultyFilters: document.querySelector("#difficulty-filters"),
  examDate: document.querySelector("#exam-date"),
  examName: document.querySelector("#exam-name"),
  omrInstructions: document.querySelector("#omr-instructions"),
  examRules: document.querySelector("#exam-rules"),
  examStorePath: document.querySelector("#exam-store-path"),
  errorList: document.querySelector("#generator-error-list"),
  errorPanel: document.querySelector("#generator-errors"),
  excludedCount: document.querySelector("#excluded-count"),
  exportJson: document.querySelector("#export-json"),
  filteredCount: document.querySelector("#filtered-count"),
  generateExams: document.querySelector("#generate-exams"),
  generatorStatus: document.querySelector("#generator-status"),
  includedCount: document.querySelector("#included-count"),
  institutionName: document.querySelector("#institution-name"),
  instructor: document.querySelector("#instructor"),
  objectiveFilters: document.querySelector("#objective-filters"),
  allowedMaterials: document.querySelector("#allowed-materials"),
  poolTableBody: document.querySelector("#pool-table-body"),
  poolQuestionBackdrop: document.querySelector("#pool-question-backdrop"),
  poolQuestionDetail: document.querySelector("#pool-question-detail"),
  poolQuestionEditLink: document.querySelector("#pool-question-edit-link"),
  poolQuestionModal: document.querySelector("#pool-question-modal"),
  poolQuestionTitle: document.querySelector("#pool-question-title"),
  printResults: document.querySelector("#print-results"),
  questionCount: document.querySelector("#question-count"),
  resetOverrides: document.querySelector("#reset-overrides"),
  resultCourseName: document.querySelector("#result-course-name"),
  resultExamDate: document.querySelector("#result-exam-date"),
  resultExamName: document.querySelector("#result-exam-name"),
  resultExamSetId: document.querySelector("#result-exam-set-id"),
  resultGeneratedAt: document.querySelector("#result-generated-at"),
  resultHeading: document.querySelector("#result-heading"),
  resultSelectedCount: document.querySelector("#result-selected-count"),
  resultVariantCount: document.querySelector("#result-variant-count"),
  results: document.querySelector("#generation-results"),
  closePoolQuestion: document.querySelector("#close-pool-question"),
  sortStatus: document.querySelector("#sort-status"),
  startTime: document.querySelector("#start-time"),
  teacherSummaryBody: document.querySelector("#teacher-summary-body"),
  totalTimeMinutes: document.querySelector("#total-time-minutes"),
  variantCount: document.querySelector("#variant-count"),
  variantPreviews: document.querySelector("#variant-previews"),
};

function printableMetadata() {
  const totalTimeRaw = elements.totalTimeMinutes.value.trim();
  const parsedTotalTime = totalTimeRaw === "" ? null : Number.parseInt(totalTimeRaw, 10);
  return {
    institutionName: elements.institutionName.value.trim(),
    examName: elements.examName.value.trim(),
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

function normalizeTextValue(value, fallback = "") {
  return typeof value === "string" && value.trim() !== "" ? value.trim() : fallback;
}

function runPrintSettings(run) {
  const settings = run?.printSettings ?? {};
  return {
    institutionName: normalizeTextValue(settings.institutionName, "Institution Name"),
    examName: normalizeTextValue(settings.examName, run?.quiz?.title || "Generated Exam"),
    courseName: normalizeTextValue(settings.courseName, "—"),
    examDate: normalizeTextValue(settings.examDate, "—"),
    startTime: normalizeTextValue(settings.startTime, "—"),
    totalTimeMinutes: String(settings.totalTimeMinutes ?? "").trim() || "—",
    instructor: normalizeTextValue(settings.instructor, ""),
    allowedMaterials: normalizeTextValue(settings.allowedMaterials, ""),
    omrInstructions: normalizeTextValue(settings.omrInstructions, DEFAULT_OMR_INSTRUCTIONS),
    examRules: normalizeExamRules(settings.examRules),
  };
}

function variantPageCount(variant) {
  const totalPages = variant?.printLayout?.totalPages;
  return Number.isInteger(totalPages) && totalPages > 0 ? totalPages : 1;
}

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  elements.generatorStatus.textContent = message;
  elements.generatorStatus.style.color = isError ? "var(--danger-strong)" : "var(--muted)";
}

function dedupe(items) {
  return [...new Set(items)];
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

function renderQuestionImageHtml(question) {
  const images = questionImageAssetIds(question);
  if (images.length === 0) {
    return "";
  }
  return `
    <div class="question-image-list">
      ${images.map((assetId, index) => `
        <img class="question-image-preview" src="${assetUrl(assetId)}" alt="Question image ${index + 1}" loading="lazy" />
      `).join("")}
    </div>
  `;
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

function locationText(value) {
  if (typeof value === "string") {
    return value.trim();
  }
  if (typeof value === "number") {
    return String(value);
  }
  return "";
}

function locationSourceLabel(location = {}) {
  return (
    locationText(location.chapter)
    || locationText(location.source)
    || locationText(location.url)
    || locationText(location.reference)
  );
}

function locationSourceDisplay(location = {}) {
  return locationText(location.source) || locationText(location.url) || "—";
}

function locationLocator(location = {}) {
  const chapter = locationText(location.chapter);
  const section = locationText(location.section);
  const page = locationText(location.page);
  return [
    chapter,
    section,
    page ? `Page ${page}` : "",
  ].filter(Boolean).join(" · ") || "—";
}

function questionSources(question) {
  return dedupe(
    ((question.locations ?? question.bookLocations) ?? [])
      .map((location) => locationSourceLabel(location))
      .filter(Boolean),
  );
}

function objectiveLabel(objectiveId) {
  return state.quiz.learningObjectives.find((objective) => objective.id === objectiveId)?.label ?? objectiveId;
}

function filterOptions() {
  const sources = dedupe(state.quiz.questions.flatMap(questionSources)).sort((left, right) =>
    left.localeCompare(right),
  );
  const difficulties = dedupe(state.quiz.questions.map((question) => question.difficulty)).sort((left, right) => left - right);
  const learningObjectives = state.quiz.learningObjectives.map((objective) => ({
    id: objective.id,
    label: objective.label,
  }));

  return { sources, difficulties, learningObjectives };
}

function overrideMode(questionId) {
  return state.selection.overrides[questionId] ?? "auto";
}

function selectedOverrideIds(mode) {
  return Object.entries(state.selection.overrides)
    .filter(([, value]) => value === mode)
    .map(([questionId]) => questionId);
}

function matchesFilters(question) {
  if (state.selection.sources.length > 0) {
    const sources = questionSources(question);
    if (!sources.some((source) => state.selection.sources.includes(source))) {
      return false;
    }
  }

  if (state.selection.difficulties.length > 0) {
    if (!state.selection.difficulties.includes(question.difficulty)) {
      return false;
    }
  }

  if (state.selection.learningObjectiveIds.length > 0) {
    if (!question.learningObjectiveIds.some((objectiveId) => state.selection.learningObjectiveIds.includes(objectiveId))) {
      return false;
    }
  }

  return true;
}

function filteredQuestions() {
  return state.quiz.questions.filter(matchesFilters);
}

function availableQuestions() {
  const filteredIds = new Set(filteredQuestions().map((question) => question.id));
  const includeIds = new Set(selectedOverrideIds("include"));
  const excludeIds = new Set(selectedOverrideIds("exclude"));

  return state.quiz.questions.filter((question) => {
    if (excludeIds.has(question.id)) {
      return false;
    }
    return filteredIds.has(question.id) || includeIds.has(question.id);
  });
}

function rowStatus(question) {
  const override = overrideMode(question.id);
  if (override === "exclude") {
    return { label: "Excluded", tone: "exclude" };
  }
  if (override === "include") {
    return { label: "Forced In", tone: "include" };
  }
  if (matchesFilters(question)) {
    return { label: "Eligible", tone: "eligible" };
  }
  return { label: "Filtered Out", tone: "filtered" };
}

function statusSortRank(tone) {
  if (state.statusSortDirection === "default") {
    return {
      eligible: 0,
      include: 1,
      filtered: 2,
      exclude: 3,
    }[tone] ?? 99;
  }
  return {
    exclude: 0,
    filtered: 1,
    include: 2,
    eligible: 3,
  }[tone] ?? 99;
}

function updateStatusSortButton() {
  const label = state.statusSortDirection === "default" ? "Status ↓" : "Status ↑";
  const ariaLabel = state.statusSortDirection === "default"
    ? "Sort by status category, eligible first"
    : "Sort by status category, excluded first";
  elements.sortStatus.textContent = label;
  elements.sortStatus.setAttribute("aria-label", ariaLabel);
}

function updateSummary() {
  elements.filteredCount.textContent = String(filteredQuestions().length);
  elements.availableCount.textContent = String(availableQuestions().length);
  elements.includedCount.textContent = String(selectedOverrideIds("include").length);
  elements.excludedCount.textContent = String(selectedOverrideIds("exclude").length);
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

function createFilterChip(labelText, checked, onChange) {
  const label = document.createElement("label");
  label.className = "filter-chip";

  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  input.addEventListener("change", onChange);

  const text = document.createElement("span");
  text.textContent = labelText;

  label.append(input, text);
  return label;
}

function renderFilterGroups() {
  const options = filterOptions();

  const sourceFragment = document.createDocumentFragment();
  for (const source of options.sources) {
    sourceFragment.append(
      createFilterChip(source, state.selection.sources.includes(source), (event) => {
        state.selection.sources = event.target.checked
          ? [...state.selection.sources, source]
          : state.selection.sources.filter((item) => item !== source);
        state.selection.sources = dedupe(state.selection.sources);
        renderPoolState();
        scheduleDraftSave();
      }),
    );
  }
  elements.sourceFilters.replaceChildren(sourceFragment);

  const difficultyFragment = document.createDocumentFragment();
  for (const difficulty of options.difficulties) {
    difficultyFragment.append(
      createFilterChip(`Difficulty ${difficulty}`, state.selection.difficulties.includes(difficulty), (event) => {
        state.selection.difficulties = event.target.checked
          ? [...state.selection.difficulties, difficulty]
          : state.selection.difficulties.filter((item) => item !== difficulty);
        state.selection.difficulties = dedupe(state.selection.difficulties).sort((left, right) => left - right);
        renderPoolState();
        scheduleDraftSave();
      }),
    );
  }
  elements.difficultyFilters.replaceChildren(difficultyFragment);

  const objectiveFragment = document.createDocumentFragment();
  for (const objective of options.learningObjectives) {
    objectiveFragment.append(
      createFilterChip(
        `${objective.id} · ${objective.label}`,
        state.selection.learningObjectiveIds.includes(objective.id),
        (event) => {
          state.selection.learningObjectiveIds = event.target.checked
            ? [...state.selection.learningObjectiveIds, objective.id]
            : state.selection.learningObjectiveIds.filter((item) => item !== objective.id);
          state.selection.learningObjectiveIds = dedupe(state.selection.learningObjectiveIds);
          renderPoolState();
          scheduleDraftSave();
        },
      ),
    );
  }
  elements.objectiveFilters.replaceChildren(objectiveFragment);
}

function renderPoolTable() {
  const fragment = document.createDocumentFragment();
  updateStatusSortButton();
  const sortedQuestions = [...state.quiz.questions].sort((left, right) => {
    const leftStatus = rowStatus(left);
    const rightStatus = rowStatus(right);
    const rankDelta = statusSortRank(leftStatus.tone) - statusSortRank(rightStatus.tone);
    if (rankDelta !== 0) {
      return rankDelta;
    }
    return String(left.id).localeCompare(String(right.id));
  });

  for (const question of sortedQuestions) {
    const row = document.createElement("tr");
    row.className = "pool-table__row";
    const status = rowStatus(question);

    const questionId = document.createElement("td");
    questionId.className = "cell-mono";
    questionId.textContent = question.id;

    const prompt = document.createElement("td");
    const promptCopy = document.createElement("p");
    promptCopy.className = "cell-copy";
    promptCopy.textContent = truncate(stripRichTextMarkup(question.question));
    prompt.append(promptCopy);

    const sources = document.createElement("td");
    sources.className = "cell-copy";
    sources.textContent = questionSources(question).join(", ") || "—";

    const difficulty = document.createElement("td");
    difficulty.textContent = String(question.difficulty);

    const points = document.createElement("td");
    points.textContent = String(question.points ?? 1);

    const objectives = document.createElement("td");
    objectives.className = "cell-copy";
    objectives.textContent = question.learningObjectiveIds
      .map((objectiveId) => `${objectiveId} · ${objectiveLabel(objectiveId)}`)
      .join(", ");

    const shuffleChoices = document.createElement("td");
    shuffleChoices.textContent = question.shuffleChoices ? "Yes" : "No";

    const statusCell = document.createElement("td");
    const badge = document.createElement("span");
    badge.className = `status-badge status-badge--${status.tone}`;
    badge.textContent = status.label;
    statusCell.append(badge);

    const overrideCell = document.createElement("td");
    const select = document.createElement("select");
    select.className = "inline-select";
    for (const [value, labelText] of [
      ["auto", "Auto"],
      ["include", "Force Include"],
      ["exclude", "Force Exclude"],
    ]) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = labelText;
      option.selected = overrideMode(question.id) === value;
      select.append(option);
    }
    select.addEventListener("change", (event) => {
      const nextValue = event.target.value;
      if (nextValue === "auto") {
        delete state.selection.overrides[question.id];
      } else {
        state.selection.overrides[question.id] = nextValue;
      }
      renderPoolState();
      scheduleDraftSave();
    });
    select.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    overrideCell.append(select);

    row.addEventListener("click", () => {
      state.activePoolQuestionId = question.id;
      renderPoolQuestionModal();
    });

    row.append(questionId, prompt, sources, difficulty, points, objectives, shuffleChoices, statusCell, overrideCell);
    fragment.append(row);
  }

  elements.poolTableBody.replaceChildren(fragment);
}

function getActivePoolQuestion() {
  return state.quiz?.questions?.find((question) => question.id === state.activePoolQuestionId) ?? null;
}

function closePoolQuestionModal() {
  state.activePoolQuestionId = "";
  renderPoolQuestionModal();
}

function renderPoolQuestionModal() {
  const question = getActivePoolQuestion();
  const isOpen = Boolean(question);
  elements.poolQuestionModal.classList.toggle("is-open", isOpen);
  elements.poolQuestionModal.setAttribute("aria-hidden", String(!isOpen));
  document.body.style.overflow = isOpen ? "hidden" : "";
  if (!question) {
    elements.poolQuestionTitle.textContent = "Selected Question";
    elements.poolQuestionDetail.replaceChildren();
    elements.poolQuestionEditLink.href = "/index.html";
    return;
  }

  const status = rowStatus(question);
  elements.poolQuestionTitle.textContent = `${question.id} · ${question.points ?? 1} pt`;
  elements.poolQuestionEditLink.href = `/index.html?questionId=${encodeURIComponent(question.id)}`;

  const sources = questionSources(question).join(", ") || "—";
  const objectives = question.learningObjectiveIds
    .map((objectiveId) => `${objectiveId} · ${objectiveLabel(objectiveId)}`)
    .join(", ") || "—";
  const choicesMarkup = question.choices.map((choice) => {
    const isCorrect = question.correctAnswers.includes(choice.key);
    return `
      <li>
        <span class="choice-pill">
          <span class="choice-pill__key">${escapeHtml(choice.key)}.</span>
          <span class="choice-pill__text">${renderRichTextHtml(choice.text)}</span>
          ${isCorrect ? '<span class="choice-pill__meta">Correct</span>' : ""}
        </span>
      </li>
    `;
  }).join("");
  const referencesMarkup = ((question.locations ?? question.bookLocations) ?? []).map((location) => `
    <tr>
      <td>${escapeHtml(locationSourceDisplay(location))}</td>
      <td>${escapeHtml(locationLocator(location))}</td>
      <td class="cell-copy">${escapeHtml(locationText(location.url) || "—")}</td>
      <td class="cell-copy">${escapeHtml(locationText(location.reference) || "—")}</td>
    </tr>
  `).join("");

  elements.poolQuestionDetail.innerHTML = `
    <div class="question-detail-sheet__meta">
      <div class="metric"><span class="metric__label">Status</span><span class="metric__value"><span class="status-badge status-badge--${status.tone}">${escapeHtml(status.label)}</span></span></div>
      <div class="metric"><span class="metric__label">Difficulty</span><span class="metric__value">${question.difficulty}</span></div>
      <div class="metric"><span class="metric__label">Points</span><span class="metric__value">${question.points ?? 1}</span></div>
      <div class="metric"><span class="metric__label">Shuffle Choices</span><span class="metric__value">${question.shuffleChoices ? "Yes" : "No"}</span></div>
    </div>
    <section class="question-detail-sheet__block">
      <h3>Prompt</h3>
      <div class="question-detail-sheet__copy">${renderRichTextHtml(question.question)}</div>
      ${renderQuestionImageHtml(question)}
    </section>
    <section class="question-detail-sheet__block">
      <h3>Coverage</h3>
      <p class="question-detail-sheet__copy"><strong>Sources:</strong> ${escapeHtml(sources)}</p>
      <p class="question-detail-sheet__copy"><strong>Learning Objectives:</strong> ${escapeHtml(objectives)}</p>
    </section>
    <section class="question-detail-sheet__block">
      <h3>Choices</h3>
      <ul class="choice-list">${choicesMarkup}</ul>
    </section>
    <section class="question-detail-sheet__block">
      <h3>Explanation</h3>
      <div class="question-detail-sheet__copy">${renderRichTextHtml(question.explanation || "—")}</div>
    </section>
    <section class="question-detail-sheet__block">
      <h3>References</h3>
      <div class="table-wrap">
        <table class="pool-table">
          <thead>
            <tr>
              <th>Source</th>
              <th>Locator</th>
              <th>URL</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody>${referencesMarkup || '<tr><td colspan="4">No references listed.</td></tr>'}</tbody>
        </table>
      </div>
    </section>
  `;
}

function renderPoolState() {
  renderFilterGroups();
  renderPoolTable();
  updateSummary();
}

function renderGeneratedRun() {
  const run = state.generatedRun;
  const hasRun = Boolean(run);
  elements.results.classList.toggle("hidden", !hasRun);
  elements.exportJson.disabled = !hasRun;
  elements.printResults.disabled = !hasRun;

  if (!hasRun) {
    elements.teacherSummaryBody.replaceChildren();
    elements.variantPreviews.replaceChildren();
    elements.resultExamName.textContent = "";
    elements.resultCourseName.textContent = "";
    elements.resultExamDate.textContent = "";
    return;
  }

  const printSettings = runPrintSettings(run);
  elements.resultHeading.textContent = `${printSettings.examName} Variants`;
  elements.resultExamSetId.textContent = run.examSetId;
  elements.resultGeneratedAt.textContent = new Date(run.generatedAt).toLocaleString();
  elements.resultSelectedCount.textContent = `${run.selection.selectedQuestionIds.length} selected`;
  elements.resultVariantCount.textContent = `${run.variants.length} generated`;
  elements.resultExamName.textContent = printSettings.examName;
  elements.resultCourseName.textContent = printSettings.courseName;
  elements.resultExamDate.textContent = printSettings.examDate;

  const teacherFragment = document.createDocumentFragment();
  for (const variant of run.variants) {
    for (const question of variant.questions) {
      const row = document.createElement("tr");
      for (const value of [
        variant.variantId,
        String(question.position),
        question.sourceQuestionId,
        String(question.points ?? 1),
        question.displayCorrectAnswers.join(", "),
        question.sourceCorrectAnswers.join(", "),
      ]) {
        const cell = document.createElement("td");
        cell.textContent = value;
        row.append(cell);
      }
      teacherFragment.append(row);
    }
  }
  elements.teacherSummaryBody.replaceChildren(teacherFragment);

  const variantFragment = document.createDocumentFragment();
  for (const variant of run.variants) {
    const card = document.createElement("article");
    card.className = "variant-card";

    const header = document.createElement("div");
    header.className = "variant-card__header";

    const headingBlock = document.createElement("div");
    const eyebrow = document.createElement("p");
    eyebrow.className = "eyebrow";
    eyebrow.textContent = "Student Variant";
    const title = document.createElement("h3");
    title.textContent = printSettings.examName;
    const meta = document.createElement("p");
    meta.className = "variant-card__meta";
    const metaParts = [];
    metaParts.push(printSettings.institutionName);
    if (printSettings.courseName !== "—") metaParts.push(printSettings.courseName);
    if (printSettings.examDate !== "—") metaParts.push(printSettings.examDate);
    if (printSettings.startTime !== "—") metaParts.push(printSettings.startTime);
    if (printSettings.totalTimeMinutes !== "—") metaParts.push(`${printSettings.totalTimeMinutes} min`);
    metaParts.push(`${variant.questions.reduce((total, question) => total + Number(question.points ?? 1), 0)} pts`);
    meta.textContent = metaParts.join(" · ") || "Generated from the current quiz pool.";
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

    const cover = document.createElement("section");
    cover.className = "variant-card__cover";

    const studentInfo = document.createElement("div");
    studentInfo.className = "variant-card__block";
    studentInfo.innerHTML = `
      <h4>Student Information</h4>
      <div class="variant-card__line-grid">
        <div class="variant-card__line-field variant-card__line-field--wide"><span>Name</span><i></i></div>
        <div class="variant-card__line-field"><span>ID</span><i></i></div>
        <div class="variant-card__line-field"><span>Class / Section</span><i></i></div>
        <div class="variant-card__line-field variant-card__line-field--wide"><span>Signature</span><i></i></div>
      </div>
    `;

    const examInfo = document.createElement("div");
    examInfo.className = "variant-card__block";
    const instructorFact = printSettings.instructor
      ? `<div class="variant-card__fact"><span>Instructor</span><strong>${escapeHtml(printSettings.instructor)}</strong></div>`
      : "";
    const materialsFact = printSettings.allowedMaterials
      ? `<div class="variant-card__fact"><span>Materials</span><strong>${escapeHtml(printSettings.allowedMaterials)}</strong></div>`
      : "";
    examInfo.innerHTML = `
      <h4>Exam Information</h4>
      <div class="variant-card__fact-grid">
        <div class="variant-card__fact"><span>Exam Name</span><strong>${escapeHtml(printSettings.examName)}</strong></div>
        <div class="variant-card__fact"><span>Course / Subject</span><strong>${escapeHtml(printSettings.courseName)}</strong></div>
        <div class="variant-card__fact"><span>Exam Date</span><strong>${escapeHtml(printSettings.examDate)}</strong></div>
        <div class="variant-card__fact"><span>Start Time</span><strong>${escapeHtml(printSettings.startTime)}</strong></div>
        <div class="variant-card__fact"><span>Total Time in Minutes</span><strong>${escapeHtml(printSettings.totalTimeMinutes)}</strong></div>
        ${instructorFact}
        ${materialsFact}
        <div class="variant-card__fact"><span>Number of Questions</span><strong>${variant.questions.length}</strong></div>
        <div class="variant-card__fact"><span>Total Points</span><strong>${variant.questions.reduce((total, question) => total + Number(question.points ?? 1), 0)}</strong></div>
        <div class="variant-card__fact"><span>Number of Pages</span><strong>${variantPageCount(variant)}</strong></div>
      </div>
    `;

    const rules = document.createElement("div");
    rules.className = "variant-card__block";
    const rulesMarkup = printSettings.examRules
      .map((rule) => `<li>${escapeHtml(rule)}</li>`)
      .join("");
    rules.innerHTML = `
      <h4>Exam Rules</h4>
      <ol class="variant-card__rules">${rulesMarkup}</ol>
      <p class="variant-card__instruction">Review each question and mark all correct answers. Questions begin on page 2.</p>
    `;

    cover.append(studentInfo, examInfo, rules);

    const questionList = document.createElement("div");
    questionList.className = "question-preview-list";

    for (const question of variant.questions) {
      const section = document.createElement("section");
      section.className = "question-preview";

      const questionHead = document.createElement("div");
      questionHead.className = "question-preview__head";
      questionHead.textContent = `Question ${question.position} · ${question.points ?? 1} pt`;

      const title = document.createElement("p");
      title.className = "question-preview__title";
      renderRichTextIntoElement(title, question.question);
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

      section.append(questionHead, title);
      if (imagePreviews) {
        section.append(imagePreviews);
      }
      section.append(choices);
      questionList.append(section);
    }

    card.append(header, cover, questionList);
    variantFragment.append(card);
  }
  elements.variantPreviews.replaceChildren(variantFragment);
}

function currentDraft() {
  return {
    selection: {
      questionCount: state.selection.questionCount,
      variantCount: state.selection.variantCount,
      sources: [...state.selection.sources],
      difficulties: [...state.selection.difficulties],
      learningObjectiveIds: [...state.selection.learningObjectiveIds],
      overrides: { ...state.selection.overrides },
    },
    statusSortDirection: state.statusSortDirection,
    printableMetadata: printableMetadata(),
    lastGeneratedExamSetId: state.generatedRun?.examSetId ?? "",
  };
}

let draftSaveTimer = null;

function scheduleDraftSave() {
  if (!state.quiz) {
    return;
  }
  window.clearTimeout(draftSaveTimer);
  draftSaveTimer = window.setTimeout(async () => {
    try {
      await fetch("/api/generator-draft", {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(currentDraft()),
      });
    } catch (error) {
      console.warn("Could not save generator draft", error);
    }
  }, 350);
}

async function loadGeneratorDraft() {
  const response = await fetch("/api/generator-draft");
  if (!response.ok) {
    return null;
  }
  const payload = await response.json();
  return payload.draft && typeof payload.draft === "object" ? payload.draft : null;
}

function applyGeneratorDraft(draft) {
  if (!draft || !state.quiz) {
    return;
  }
  const selection = draft.selection && typeof draft.selection === "object" ? draft.selection : {};
  const options = filterOptions();
  const sourceSet = new Set(options.sources);
  const difficultySet = new Set(options.difficulties);
  const objectiveSet = new Set(options.learningObjectives.map((objective) => objective.id));
  const questionIdSet = new Set(state.quiz.questions.map((question) => question.id));

  const questionCount = Number.parseInt(selection.questionCount, 10);
  const variantCount = Number.parseInt(selection.variantCount, 10);
  state.selection.questionCount = Math.min(
    MAX_QUESTIONS_PER_EXAM,
    Math.max(1, Number.isFinite(questionCount) ? questionCount : state.selection.questionCount),
  );
  state.selection.variantCount = Math.max(1, Number.isFinite(variantCount) ? variantCount : state.selection.variantCount);
  state.selection.sources = Array.isArray(selection.sources)
    ? dedupe(selection.sources.filter((source) => sourceSet.has(source)))
    : [];
  state.selection.difficulties = Array.isArray(selection.difficulties)
    ? dedupe(selection.difficulties.filter((difficulty) => difficultySet.has(difficulty))).sort((left, right) => left - right)
    : [];
  state.selection.learningObjectiveIds = Array.isArray(selection.learningObjectiveIds)
    ? dedupe(selection.learningObjectiveIds.filter((objectiveId) => objectiveSet.has(objectiveId)))
    : [];
  state.selection.overrides = {};
  if (selection.overrides && typeof selection.overrides === "object") {
    for (const [questionId, mode] of Object.entries(selection.overrides)) {
      if (questionIdSet.has(questionId) && (mode === "include" || mode === "exclude")) {
        state.selection.overrides[questionId] = mode;
      }
    }
  }

  const metadata = draft.printableMetadata && typeof draft.printableMetadata === "object" ? draft.printableMetadata : {};
  elements.institutionName.value = typeof metadata.institutionName === "string" ? metadata.institutionName : elements.institutionName.value;
  elements.examName.value = typeof metadata.examName === "string" ? metadata.examName : elements.examName.value;
  elements.courseName.value = typeof metadata.courseName === "string" ? metadata.courseName : elements.courseName.value;
  elements.examDate.value = typeof metadata.examDate === "string" ? metadata.examDate : elements.examDate.value;
  elements.startTime.value = typeof metadata.startTime === "string" ? metadata.startTime : elements.startTime.value;
  elements.totalTimeMinutes.value = Number.isInteger(metadata.totalTimeMinutes) ? String(metadata.totalTimeMinutes) : elements.totalTimeMinutes.value;
  elements.instructor.value = typeof metadata.instructor === "string" ? metadata.instructor : elements.instructor.value;
  elements.allowedMaterials.value = typeof metadata.allowedMaterials === "string" ? metadata.allowedMaterials : elements.allowedMaterials.value;
  elements.omrInstructions.value = typeof metadata.omrInstructions === "string" ? metadata.omrInstructions : elements.omrInstructions.value;
  elements.examRules.value = Array.isArray(metadata.examRules) ? metadata.examRules.join("\n") : elements.examRules.value;
  state.statusSortDirection = draft.statusSortDirection === "reverse" ? "reverse" : "default";
}

async function restoreGeneratedRunFromDraft(draft) {
  const examSetId = typeof draft?.lastGeneratedExamSetId === "string" ? draft.lastGeneratedExamSetId.trim() : "";
  if (!examSetId) {
    return;
  }
  try {
    const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}`);
    if (!response.ok) {
      return;
    }
    const payload = await response.json();
    state.generatedRun = payload.examSet;
  } catch (error) {
    console.warn("Could not restore generated exam set", error);
  }
}

async function loadQuiz() {
  setStatus("Loading quiz data...");
  const response = await fetch("/api/quiz");
  if (!response.ok) {
    throw new Error(`Could not load quiz (${response.status})`);
  }
  const payload = await response.json();
  state.quiz = payload.quiz;
  state.dbPath = payload.dbPath;
  state.examStorePath = payload.examStorePath;
  state.selection.questionCount = Math.max(1, Math.min(10, state.quiz.questions.length));
  state.selection.questionCount = Math.min(state.selection.questionCount, MAX_QUESTIONS_PER_EXAM);
  state.selection.variantCount = 1;
  if (!elements.examName.value.trim()) {
    elements.examName.value = state.quiz.title ?? "";
  }
  if (!elements.omrInstructions.value.trim()) {
    elements.omrInstructions.value = DEFAULT_OMR_INSTRUCTIONS;
  }
  if (!elements.examRules.value.trim()) {
    elements.examRules.value = defaultExamRules(
      elements.omrInstructions.value.trim() || DEFAULT_OMR_INSTRUCTIONS,
    ).join("\n");
  }
  const draft = await loadGeneratorDraft();
  applyGeneratorDraft(draft);
  await restoreGeneratedRunFromDraft(draft);
  elements.dbPath.textContent = payload.projectPath ?? state.dbPath;
  elements.examStorePath.textContent = payload.projectPath ?? state.examStorePath;
  elements.questionCount.value = String(state.selection.questionCount);
  elements.variantCount.value = String(state.selection.variantCount);
  renderPoolState();
  renderGeneratedRun();
  setStatus("Quiz pool loaded.");
}

async function generateExams() {
  state.validationErrors = [];
  renderErrors();
  setStatus("Generating variants...");

  if (Number.parseInt(elements.questionCount.value, 10) > MAX_QUESTIONS_PER_EXAM) {
    state.validationErrors = [
      {
        path: "questionCount",
        message: `Questions per exam cannot be greater than ${MAX_QUESTIONS_PER_EXAM}.`,
      },
    ];
    renderErrors();
    setStatus(`Questions per exam cannot be greater than ${MAX_QUESTIONS_PER_EXAM}.`, true);
    return;
  }

  const payload = {
    questionCount: Number.parseInt(elements.questionCount.value, 10),
    variantCount: Number.parseInt(elements.variantCount.value, 10),
    sources: [...state.selection.sources],
    difficulties: [...state.selection.difficulties],
    learningObjectiveIds: [...state.selection.learningObjectiveIds],
    includeQuestionIds: selectedOverrideIds("include"),
    excludeQuestionIds: selectedOverrideIds("exclude"),
    ...printableMetadata(),
  };

  const response = await fetch("/api/exams/generate", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const result = await response.json();
  if (!response.ok) {
    state.validationErrors = result.errors ?? [{ path: "<unknown>", message: "Generation failed" }];
    renderErrors();
    setStatus("Generation failed. Review the validation messages.", true);
    return;
  }

  state.generatedRun = result;
  state.validationErrors = [];
  renderErrors();
  renderGeneratedRun();
  scheduleDraftSave();
  setStatus(`Generated ${result.variants.length} variant(s) for exam set ${result.examSetId}.`);
}

function exportRun() {
  if (!state.generatedRun) {
    return;
  }
  const blob = new Blob([JSON.stringify(state.generatedRun, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = `exam-set-${state.generatedRun.examSetId}.json`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

async function downloadPrintableZip() {
  if (!state.generatedRun) {
    return;
  }

  setStatus("Preparing printable ZIP...");
  const response = await fetch(`/api/exams/export/${encodeURIComponent(state.generatedRun.examSetId)}.zip`);
  if (!response.ok) {
    let message = `Printable export failed (${response.status})`;
    try {
      const payload = await response.json();
      message = payload.errors?.[0]?.message ?? message;
    } catch {
      // Keep the fallback message if the response is not JSON.
    }
    setStatus(message, true);
    return;
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `exam-set-${state.generatedRun.examSetId}-printables.zip`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
  setStatus(`Downloaded printable ZIP for exam set ${state.generatedRun.examSetId}.`);
}

function wireEvents() {
  for (const element of [
    elements.institutionName,
    elements.examName,
    elements.courseName,
    elements.examDate,
    elements.startTime,
    elements.totalTimeMinutes,
    elements.instructor,
    elements.allowedMaterials,
    elements.examRules,
  ]) {
    element.addEventListener("input", () => {
      scheduleDraftSave();
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
    }
    scheduleDraftSave();
  });

  elements.questionCount.addEventListener("input", (event) => {
    const nextValue = Math.max(1, Number.parseInt(event.target.value || "1", 10));
    state.selection.questionCount = Math.min(MAX_QUESTIONS_PER_EXAM, nextValue);
    if (nextValue > MAX_QUESTIONS_PER_EXAM) {
      state.validationErrors = [
        {
          path: "questionCount",
          message: `Questions per exam cannot be greater than ${MAX_QUESTIONS_PER_EXAM}.`,
        },
      ];
      renderErrors();
      setStatus(`Questions per exam cannot be greater than ${MAX_QUESTIONS_PER_EXAM}.`, true);
    } else if (state.validationErrors.some((error) => error.path === "questionCount")) {
      state.validationErrors = state.validationErrors.filter((error) => error.path !== "questionCount");
      renderErrors();
      setStatus("Quiz pool loaded.");
    }
    event.target.value = String(state.selection.questionCount);
    scheduleDraftSave();
  });

  elements.variantCount.addEventListener("input", (event) => {
    state.selection.variantCount = Math.max(1, Number.parseInt(event.target.value || "1", 10));
    scheduleDraftSave();
  });

  elements.generateExams.addEventListener("click", async () => {
    await generateExams();
  });

  elements.exportJson.addEventListener("click", () => {
    exportRun();
  });

  elements.printResults.addEventListener("click", () => {
    void downloadPrintableZip();
  });

  elements.resetOverrides.addEventListener("click", () => {
    state.selection.overrides = {};
    renderPoolState();
    scheduleDraftSave();
    setStatus("Question overrides reset.");
  });

  elements.sortStatus.addEventListener("click", () => {
    state.statusSortDirection = state.statusSortDirection === "default" ? "reverse" : "default";
    renderPoolTable();
    scheduleDraftSave();
  });

  elements.closePoolQuestion.addEventListener("click", () => {
    closePoolQuestionModal();
  });

  elements.poolQuestionBackdrop.addEventListener("click", () => {
    closePoolQuestionModal();
  });

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && state.activePoolQuestionId) {
      closePoolQuestionModal();
    }
  });
}

wireEvents();
loadQuiz().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
