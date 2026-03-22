const state = {
  dbPath: "",
  examStorePath: "",
  generatedRun: null,
  quiz: null,
  statusIsError: false,
  statusMessage: "Loading quiz data...",
  validationErrors: [],
  selection: {
    questionCount: 1,
    variantCount: 1,
    chapters: [],
    difficulties: [],
    learningObjectiveIds: [],
    overrides: {},
  },
};

const DEFAULT_EXAM_RULES = [
  "Complete the student information block before the exam begins.",
  "Read every question carefully and select all correct answers for each question.",
  "Mark answers clearly and keep your paper neat for printing, photocopying, and scanning.",
  "Do not communicate with other students or use unauthorized materials during the exam.",
  "Remain seated until instructed to stop and submit your paper.",
];

const elements = {
  availableCount: document.querySelector("#available-count"),
  chapterFilters: document.querySelector("#chapter-filters"),
  courseName: document.querySelector("#course-name"),
  dbPath: document.querySelector("#db-path"),
  difficultyFilters: document.querySelector("#difficulty-filters"),
  examDate: document.querySelector("#exam-date"),
  examName: document.querySelector("#exam-name"),
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
  objectiveFilters: document.querySelector("#objective-filters"),
  poolTableBody: document.querySelector("#pool-table-body"),
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
  return [...DEFAULT_EXAM_RULES];
}

function questionChapters(question) {
  return dedupe(
    (question.bookLocations ?? [])
      .map((location) => location.chapter)
      .filter((chapter) => typeof chapter === "string" && chapter.trim() !== "")
      .map((chapter) => chapter.trim()),
  );
}

function objectiveLabel(objectiveId) {
  return state.quiz.learningObjectives.find((objective) => objective.id === objectiveId)?.label ?? objectiveId;
}

function filterOptions() {
  const chapters = dedupe(state.quiz.questions.flatMap(questionChapters)).sort((left, right) =>
    left.localeCompare(right),
  );
  const difficulties = dedupe(state.quiz.questions.map((question) => question.difficulty)).sort((left, right) => left - right);
  const learningObjectives = state.quiz.learningObjectives.map((objective) => ({
    id: objective.id,
    label: objective.label,
  }));

  return { chapters, difficulties, learningObjectives };
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
  if (state.selection.chapters.length > 0) {
    const chapters = questionChapters(question);
    if (!chapters.some((chapter) => state.selection.chapters.includes(chapter))) {
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

  const chapterFragment = document.createDocumentFragment();
  for (const chapter of options.chapters) {
    chapterFragment.append(
      createFilterChip(chapter, state.selection.chapters.includes(chapter), (event) => {
        state.selection.chapters = event.target.checked
          ? [...state.selection.chapters, chapter]
          : state.selection.chapters.filter((item) => item !== chapter);
        state.selection.chapters = dedupe(state.selection.chapters);
        renderPoolState();
      }),
    );
  }
  elements.chapterFilters.replaceChildren(chapterFragment);

  const difficultyFragment = document.createDocumentFragment();
  for (const difficulty of options.difficulties) {
    difficultyFragment.append(
      createFilterChip(`Difficulty ${difficulty}`, state.selection.difficulties.includes(difficulty), (event) => {
        state.selection.difficulties = event.target.checked
          ? [...state.selection.difficulties, difficulty]
          : state.selection.difficulties.filter((item) => item !== difficulty);
        state.selection.difficulties = dedupe(state.selection.difficulties).sort((left, right) => left - right);
        renderPoolState();
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
        },
      ),
    );
  }
  elements.objectiveFilters.replaceChildren(objectiveFragment);
}

function renderPoolTable() {
  const fragment = document.createDocumentFragment();

  for (const question of state.quiz.questions) {
    const row = document.createElement("tr");
    const status = rowStatus(question);

    const questionId = document.createElement("td");
    questionId.className = "cell-mono";
    questionId.textContent = question.id;

    const prompt = document.createElement("td");
    const promptCopy = document.createElement("p");
    promptCopy.className = "cell-copy";
    promptCopy.textContent = truncate(question.question);
    prompt.append(promptCopy);

    const chapters = document.createElement("td");
    chapters.className = "cell-copy";
    chapters.textContent = questionChapters(question).join(", ") || "—";

    const difficulty = document.createElement("td");
    difficulty.textContent = String(question.difficulty);

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
    });
    overrideCell.append(select);

    row.append(questionId, prompt, chapters, difficulty, objectives, shuffleChoices, statusCell, overrideCell);
    fragment.append(row);
  }

  elements.poolTableBody.replaceChildren(fragment);
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
        <div class="variant-card__line-field variant-card__line-field--wide"><span>Student Name</span><i></i></div>
        <div class="variant-card__line-field"><span>Student ID</span><i></i></div>
        <div class="variant-card__line-field"><span>Class / Section</span><i></i></div>
        <div class="variant-card__line-field variant-card__line-field--wide"><span>Signature</span><i></i></div>
      </div>
    `;

    const examInfo = document.createElement("div");
    examInfo.className = "variant-card__block";
    examInfo.innerHTML = `
      <h4>Exam Information</h4>
      <div class="variant-card__fact-grid">
        <div class="variant-card__fact"><span>Exam Name</span><strong>${escapeHtml(printSettings.examName)}</strong></div>
        <div class="variant-card__fact"><span>Course / Subject</span><strong>${escapeHtml(printSettings.courseName)}</strong></div>
        <div class="variant-card__fact"><span>Exam Date</span><strong>${escapeHtml(printSettings.examDate)}</strong></div>
        <div class="variant-card__fact"><span>Start Time</span><strong>${escapeHtml(printSettings.startTime)}</strong></div>
        <div class="variant-card__fact"><span>Total Time in Minutes</span><strong>${escapeHtml(printSettings.totalTimeMinutes)}</strong></div>
        <div class="variant-card__fact"><span>Number of Questions</span><strong>${variant.questions.length}</strong></div>
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
      questionHead.textContent = `Question ${question.position}`;

      const title = document.createElement("p");
      title.className = "question-preview__title";
      title.textContent = question.question;

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

      section.append(questionHead, title, choices);
      questionList.append(section);
    }

    card.append(header, cover, questionList);
    variantFragment.append(card);
  }
  elements.variantPreviews.replaceChildren(variantFragment);
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
  state.selection.variantCount = 1;
  if (!elements.examName.value.trim()) {
    elements.examName.value = state.quiz.title ?? "";
  }
  if (!elements.examRules.value.trim()) {
    elements.examRules.value = DEFAULT_EXAM_RULES.join("\n");
  }
  elements.dbPath.textContent = state.dbPath;
  elements.examStorePath.textContent = state.examStorePath;
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

  const payload = {
    questionCount: Number.parseInt(elements.questionCount.value, 10),
    variantCount: Number.parseInt(elements.variantCount.value, 10),
    chapters: [...state.selection.chapters],
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
  elements.questionCount.addEventListener("input", (event) => {
    state.selection.questionCount = Math.max(1, Number.parseInt(event.target.value || "1", 10));
  });

  elements.variantCount.addEventListener("input", (event) => {
    state.selection.variantCount = Math.max(1, Number.parseInt(event.target.value || "1", 10));
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
    setStatus("Question overrides reset.");
  });
}

wireEvents();
loadQuiz().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
