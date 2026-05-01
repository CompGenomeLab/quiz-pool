import { renderRichTextHtml, stripRichTextMarkup } from "./rich-text.js";

const state = {
  activePoolQuestionId: "",
  dbPath: "",
  examStorePath: "",
  generatedRun: null,
  quiz: null,
  sourceSearch: "",
  statusSortDirection: "default",
  statusIsError: false,
  statusMessage: "Loading quiz data...",
  validationErrors: [],
  selection: {
    questionCount: 1,
    variantCount: 1,
    generationSeed: "",
    sources: [],
    difficulties: [],
    learningObjectiveIds: [],
    overrides: {},
  },
};

const MAX_QUESTIONS_PER_EXAM = 100;

const elements = {
  availableCount: document.querySelector("#available-count"),
  clearSourceFilters: document.querySelector("#clear-source-filters"),
  dbPath: document.querySelector("#db-path"),
  difficultyFilters: document.querySelector("#difficulty-filters"),
  examStorePath: document.querySelector("#exam-store-path"),
  errorList: document.querySelector("#generator-error-list"),
  errorPanel: document.querySelector("#generator-errors"),
  excludedCount: document.querySelector("#excluded-count"),
  filteredCount: document.querySelector("#filtered-count"),
  generateExams: document.querySelector("#generate-exams"),
  generatorStatus: document.querySelector("#generator-status"),
  generationSeed: document.querySelector("#generation-seed"),
  includedCount: document.querySelector("#included-count"),
  objectiveFilters: document.querySelector("#objective-filters"),
  poolTableBody: document.querySelector("#pool-table-body"),
  poolQuestionBackdrop: document.querySelector("#pool-question-backdrop"),
  poolQuestionDetail: document.querySelector("#pool-question-detail"),
  poolQuestionEditLink: document.querySelector("#pool-question-edit-link"),
  poolQuestionModal: document.querySelector("#pool-question-modal"),
  poolQuestionTitle: document.querySelector("#pool-question-title"),
  questionCount: document.querySelector("#question-count"),
  resetOverrides: document.querySelector("#reset-overrides"),
  resultExamSetId: document.querySelector("#result-exam-set-id"),
  resultGeneratedAt: document.querySelector("#result-generated-at"),
  resultGenerationSeed: document.querySelector("#result-generation-seed"),
  resultHeading: document.querySelector("#result-heading"),
  resultMessage: document.querySelector("#result-message"),
  resultSelectedCount: document.querySelector("#result-selected-count"),
  resultVariantCount: document.querySelector("#result-variant-count"),
  resultViewerLink: document.querySelector("#result-viewer-link"),
  results: document.querySelector("#generation-results"),
  selectVisibleSources: document.querySelector("#select-visible-sources"),
  sourceFilterCount: document.querySelector("#source-filter-count"),
  sourceFilterSearch: document.querySelector("#source-filter-search"),
  sourceFilterSummary: document.querySelector("#source-filter-summary"),
  sourceFilters: document.querySelector("#source-filters"),
  closePoolQuestion: document.querySelector("#close-pool-question"),
  sortStatus: document.querySelector("#sort-status"),
  variantCount: document.querySelector("#variant-count"),
};

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  elements.generatorStatus.textContent = message;
  elements.generatorStatus.style.color = isError ? "var(--danger-strong)" : "var(--muted)";
}

function dedupe(items) {
  return [...new Set(items)];
}

function pluralize(count, singular, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

function formatPoints(value) {
  const points = Number(value);
  const normalized = Number.isFinite(points) ? points : 0;
  return Number.isInteger(normalized) ? String(normalized) : normalized.toFixed(1).replace(/\.0$/u, "");
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

function questionLocations(question = {}) {
  if (Array.isArray(question.locations)) {
    return question.locations.filter((location) => location && typeof location === "object");
  }
  if (Array.isArray(question.bookLocations)) {
    return question.bookLocations.filter((location) => location && typeof location === "object");
  }
  return [];
}

function questionSources(question) {
  return dedupe(
    questionLocations(question)
      .map((location) => locationSourceLabel(location))
      .filter(Boolean),
  );
}

function questionSourceDetails(question) {
  return questionLocations(question)
    .map((location) => {
      const label = locationSourceLabel(location);
      if (!label) {
        return null;
      }
      return {
        label,
        source: locationSourceDisplay(location),
        locator: locationLocator(location),
        url: locationText(location.url),
        reference: locationText(location.reference),
      };
    })
    .filter(Boolean);
}

function sourceDetailText(detail) {
  return [
    detail.source !== "—" && detail.source !== detail.label ? detail.source : "",
    detail.locator !== "—" && detail.locator !== detail.label ? detail.locator : "",
    detail.url && detail.url !== detail.label ? detail.url : "",
    detail.reference && detail.reference !== detail.label ? detail.reference : "",
  ].filter(Boolean).join(" · ");
}

function difficultySummary(difficulties) {
  if (difficulties.length === 0) {
    return "No difficulty";
  }
  if (difficulties.length <= 3) {
    return `Difficulty ${difficulties.join(", ")}`;
  }
  return `Difficulty ${difficulties[0]}-${difficulties[difficulties.length - 1]}`;
}

function sourceQuestionIdPreview(questionIds) {
  const visibleIds = questionIds.slice(0, 5).join(", ");
  const hiddenCount = questionIds.length - 5;
  return hiddenCount > 0 ? `${visibleIds}, +${hiddenCount}` : visibleIds;
}

function sourceFilterOptions() {
  const optionsByLabel = new Map();
  for (const question of state.quiz.questions) {
    const details = questionSourceDetails(question);
    const labels = dedupe(details.map((detail) => detail.label));
    for (const label of labels) {
      if (!optionsByLabel.has(label)) {
        optionsByLabel.set(label, {
          label,
          questionIds: [],
          questionCount: 0,
          pointTotal: 0,
          difficulties: new Set(),
          detailLines: new Set(),
        });
      }
      const option = optionsByLabel.get(label);
      option.questionIds.push(question.id);
      option.questionCount += 1;
      option.pointTotal += Number(question.points ?? 1) || 0;
      if (question.difficulty !== undefined && question.difficulty !== null && question.difficulty !== "") {
        option.difficulties.add(question.difficulty);
      }
      for (const detail of details.filter((item) => item.label === label)) {
        const detailLine = sourceDetailText(detail);
        if (detailLine) {
          option.detailLines.add(detailLine);
        }
      }
    }
  }

  return [...optionsByLabel.values()]
    .map((option) => {
      const questionIds = [...option.questionIds].sort((left, right) => String(left).localeCompare(String(right)));
      const difficulties = [...option.difficulties].sort((left, right) => Number(left) - Number(right));
      const detailLines = [...option.detailLines].sort((left, right) => left.localeCompare(right));
      const detailPreview = detailLines.length > 0
        ? `${detailLines.slice(0, 2).join(" | ")}${detailLines.length > 2 ? ` | +${detailLines.length - 2}` : ""}`
        : "Reference label only";
      const meta = [
        pluralize(option.questionCount, "question"),
        `${formatPoints(option.pointTotal)} pt`,
        difficultySummary(difficulties),
        sourceQuestionIdPreview(questionIds),
      ].filter(Boolean);
      return {
        ...option,
        questionIds,
        difficulties,
        detailLines,
        detailPreview,
        metaText: meta.join(" · "),
        searchText: [option.label, ...detailLines, ...questionIds].join(" ").toLocaleLowerCase(),
      };
    })
    .sort((left, right) => left.label.localeCompare(right.label));
}

function visibleSourceFilterOptions(sourceOptions) {
  const query = state.sourceSearch.trim().toLocaleLowerCase();
  if (!query) {
    return sourceOptions;
  }
  return sourceOptions.filter((option) => option.searchText.includes(query));
}

function objectiveLabel(objectiveId) {
  return state.quiz.learningObjectives.find((objective) => objective.id === objectiveId)?.label ?? objectiveId;
}

function filterOptions() {
  const sourceOptions = sourceFilterOptions();
  const sources = sourceOptions.map((option) => option.label);
  const difficulties = dedupe(state.quiz.questions.map((question) => question.difficulty)).sort((left, right) => left - right);
  const learningObjectives = state.quiz.learningObjectives.map((objective) => ({
    id: objective.id,
    label: objective.label,
  }));

  return { sources, sourceOptions, difficulties, learningObjectives };
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

function setSourceSelected(source, selected) {
  state.selection.sources = selected
    ? dedupe([...state.selection.sources, source])
    : state.selection.sources.filter((item) => item !== source);
}

function createSourceOption(option) {
  const checked = state.selection.sources.includes(option.label);
  const label = document.createElement("label");
  label.className = `source-option${checked ? " is-selected" : ""}`;

  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  input.addEventListener("change", (event) => {
    setSourceSelected(option.label, event.target.checked);
    renderPoolState();
    scheduleDraftSave();
  });

  const body = document.createElement("span");
  body.className = "source-option__body";

  const top = document.createElement("span");
  top.className = "source-option__top";

  const title = document.createElement("strong");
  title.textContent = option.label;

  const count = document.createElement("span");
  count.className = "source-option__count";
  count.textContent = pluralize(option.questionCount, "question");

  const meta = document.createElement("span");
  meta.className = "source-option__meta";
  meta.textContent = option.metaText;

  const detail = document.createElement("span");
  detail.className = "source-option__detail";
  detail.textContent = option.detailPreview;

  top.append(title, count);
  body.append(top, meta, detail);
  label.append(input, body);
  return label;
}

function createSelectedSourceToken(source) {
  const button = document.createElement("button");
  button.className = "source-filter-token";
  button.type = "button";
  button.setAttribute("aria-label", `Remove ${source} source filter`);

  const text = document.createElement("span");
  text.textContent = source;

  const remove = document.createElement("span");
  remove.className = "source-filter-token__remove";
  remove.textContent = "x";

  button.append(text, remove);
  button.addEventListener("click", () => {
    setSourceSelected(source, false);
    renderPoolState();
    scheduleDraftSave();
  });
  return button;
}

function createSourceFilterEmpty(message) {
  const empty = document.createElement("div");
  empty.className = "source-filter-empty";
  empty.textContent = message;
  return empty;
}

function renderSourceFilterGroup(sourceOptions) {
  const visibleOptions = visibleSourceFilterOptions(sourceOptions);
  const selectedSet = new Set(state.selection.sources);
  const selectedSources = state.selection.sources.filter((source) =>
    sourceOptions.some((option) => option.label === source),
  );
  const selectedCount = selectedSources.length;
  const shownText = state.sourceSearch
    ? `${visibleOptions.length}/${sourceOptions.length} shown`
    : pluralize(sourceOptions.length, "source");

  elements.sourceFilterCount.textContent = selectedCount > 0
    ? `${selectedCount} selected · ${shownText}`
    : shownText;
  elements.sourceFilterSearch.value = state.sourceSearch;
  elements.sourceFilterSearch.disabled = sourceOptions.length === 0;
  elements.selectVisibleSources.disabled = visibleOptions.length === 0
    || visibleOptions.every((option) => selectedSet.has(option.label));
  elements.clearSourceFilters.disabled = selectedCount === 0;

  if (selectedSources.length === 0) {
    elements.sourceFilterSummary.classList.add("hidden");
    elements.sourceFilterSummary.replaceChildren();
  } else {
    const summaryFragment = document.createDocumentFragment();
    for (const source of selectedSources) {
      summaryFragment.append(createSelectedSourceToken(source));
    }
    elements.sourceFilterSummary.replaceChildren(summaryFragment);
    elements.sourceFilterSummary.classList.remove("hidden");
  }

  const sourceFragment = document.createDocumentFragment();
  if (sourceOptions.length === 0) {
    sourceFragment.append(createSourceFilterEmpty("No sources in this pool."));
  } else if (visibleOptions.length === 0) {
    sourceFragment.append(createSourceFilterEmpty("No sources match your search."));
  } else {
    for (const option of visibleOptions) {
      sourceFragment.append(createSourceOption(option));
    }
  }
  elements.sourceFilters.replaceChildren(sourceFragment);
}

function renderFilterGroups() {
  const options = filterOptions();
  renderSourceFilterGroup(options.sourceOptions);

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

function generatedExamViewerUrl(examSetId) {
  return `/viewer.html?examSetId=${encodeURIComponent(examSetId)}`;
}

function renderGeneratedRun() {
  const run = state.generatedRun;
  const hasRun = Boolean(run);
  elements.results.classList.toggle("hidden", !hasRun);

  if (!hasRun) {
    elements.resultHeading.textContent = "Exam Set Saved";
    elements.resultMessage.textContent = "";
    elements.resultViewerLink.href = "/viewer.html";
    elements.resultExamSetId.textContent = "";
    elements.resultGeneratedAt.textContent = "";
    elements.resultGenerationSeed.textContent = "";
    elements.resultSelectedCount.textContent = "";
    elements.resultVariantCount.textContent = "";
    return;
  }

  const selectedQuestionIds = Array.isArray(run.selection?.selectedQuestionIds)
    ? run.selection.selectedQuestionIds
    : [];
  const selectedCount = selectedQuestionIds.length || run.variants?.[0]?.questions?.length || 0;
  const variantCount = Array.isArray(run.variants) ? run.variants.length : 0;
  const viewerUrl = generatedExamViewerUrl(run.examSetId);

  elements.resultHeading.textContent = "Generation Successful";
  elements.resultMessage.textContent = `Exam set ${run.examSetId} was saved. Open it in Exam Viewer to review details, update printable metadata, and export the print ZIP.`;
  elements.resultViewerLink.href = viewerUrl;
  elements.resultExamSetId.textContent = run.examSetId;
  elements.resultGeneratedAt.textContent = new Date(run.generatedAt).toLocaleString();
  elements.resultGenerationSeed.textContent = run.generationSeed || run.selection?.generationSeed || "—";
  elements.resultSelectedCount.textContent = `${selectedCount} selected`;
  elements.resultVariantCount.textContent = `${variantCount} generated`;
}

function currentDraft() {
  return {
    selection: {
      questionCount: state.selection.questionCount,
      variantCount: state.selection.variantCount,
      generationSeed: state.selection.generationSeed,
      sources: [...state.selection.sources],
      difficulties: [...state.selection.difficulties],
      learningObjectiveIds: [...state.selection.learningObjectiveIds],
      overrides: { ...state.selection.overrides },
    },
    statusSortDirection: state.statusSortDirection,
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
  state.selection.generationSeed = typeof selection.generationSeed === "string"
    ? selection.generationSeed.slice(0, 128)
    : "";
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
  const draft = await loadGeneratorDraft();
  applyGeneratorDraft(draft);
  await restoreGeneratedRunFromDraft(draft);
  elements.dbPath.textContent = payload.projectPath ?? state.dbPath;
  elements.examStorePath.textContent = payload.projectPath ?? state.examStorePath;
  elements.questionCount.value = String(state.selection.questionCount);
  elements.variantCount.value = String(state.selection.variantCount);
  elements.generationSeed.value = state.selection.generationSeed;
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
    generationSeed: elements.generationSeed.value.trim(),
    sources: [...state.selection.sources],
    difficulties: [...state.selection.difficulties],
    learningObjectiveIds: [...state.selection.learningObjectiveIds],
    includeQuestionIds: selectedOverrideIds("include"),
    excludeQuestionIds: selectedOverrideIds("exclude"),
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
  state.selection.generationSeed = result.generationSeed || result.selection?.generationSeed || payload.generationSeed;
  elements.generationSeed.value = state.selection.generationSeed;
  state.validationErrors = [];
  renderErrors();
  renderGeneratedRun();
  scheduleDraftSave();
  setStatus(`Generated exam set ${result.examSetId}. Open it in Exam Viewer for details and export.`);
}

function wireEvents() {
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

  elements.generationSeed.addEventListener("input", (event) => {
    state.selection.generationSeed = event.target.value.slice(0, 128);
    scheduleDraftSave();
  });

  elements.sourceFilterSearch.addEventListener("input", (event) => {
    state.sourceSearch = event.target.value;
    if (state.quiz) {
      renderFilterGroups();
    }
  });

  elements.selectVisibleSources.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    const visibleSources = visibleSourceFilterOptions(filterOptions().sourceOptions).map((option) => option.label);
    state.selection.sources = dedupe([...state.selection.sources, ...visibleSources]);
    renderPoolState();
    scheduleDraftSave();
  });

  elements.clearSourceFilters.addEventListener("click", () => {
    state.selection.sources = [];
    renderPoolState();
    scheduleDraftSave();
  });

  elements.generateExams.addEventListener("click", async () => {
    await generateExams();
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
