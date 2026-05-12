import { renderRichTextHtml, stripRichTextMarkup } from "./rich-text.js";
import { buildQuestionSearchText, fuzzyQueryScore } from "./question-search.js";

const state = {
  activePoolQuestionId: "",
  dbPath: "",
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
    excludeFromExamSetId: "",
  },
  previousExams: [],
  derivedExcludeQuestionIds: new Set(),
  questionHashes: {},
  poolSearch: "",
};

const MAX_QUESTIONS_PER_EXAM = 100;

const elements = {
  availableCount: document.querySelector("#available-count"),
  clearSourceFilters: document.querySelector("#clear-source-filters"),
  dbPath: document.querySelector("#db-path"),
  difficultyFilters: document.querySelector("#difficulty-filters"),
  errorList: document.querySelector("#generator-error-list"),
  errorPanel: document.querySelector("#generator-errors"),
  excludeFromExam: document.querySelector("#exclude-from-exam"),
  restoreFromExam: document.querySelector("#restore-from-exam"),
  restoreDiffModal: document.querySelector("#restore-diff-modal"),
  restoreDiffBackdrop: document.querySelector("#restore-diff-backdrop"),
  restoreDiffSummary: document.querySelector("#restore-diff-summary"),
  restoreDiffAddedBody: document.querySelector("#restore-diff-added-body"),
  restoreDiffAddedCount: document.querySelector("#restore-diff-added-count"),
  restoreDiffRemovedBody: document.querySelector("#restore-diff-removed-body"),
  restoreDiffRemovedCount: document.querySelector("#restore-diff-removed-count"),
  restoreDiffModifiedBody: document.querySelector("#restore-diff-modified-body"),
  restoreDiffModifiedCount: document.querySelector("#restore-diff-modified-count"),
  restoreDiffAddedBlock: document.querySelector("#restore-diff-added-block"),
  restoreDiffAddedHeading: document.querySelector("#restore-diff-added-heading"),
  restoreDiffAddedDescription: document.querySelector("#restore-diff-added-description"),
  restoreDiffRemovedBlock: document.querySelector("#restore-diff-removed-block"),
  restoreDiffModifiedBlock: document.querySelector("#restore-diff-modified-block"),
  restoreDiffDetectionBlock: document.querySelector("#restore-diff-detection-block"),
  restoreDiffDetection: document.querySelector("#restore-diff-detection"),
  restoreDiffPreviewWhen: document.querySelector("#restore-diff-preview-when"),
  restoreDiffPreviewSeed: document.querySelector("#restore-diff-preview-seed"),
  restoreDiffPreviewCount: document.querySelector("#restore-diff-preview-count"),
  restoreDiffPreviewVariants: document.querySelector("#restore-diff-preview-variants"),
  restoreDiffPreviewSources: document.querySelector("#restore-diff-preview-sources"),
  restoreDiffPreviewDifficulties: document.querySelector("#restore-diff-preview-difficulties"),
  restoreDiffPreviewObjectives: document.querySelector("#restore-diff-preview-objectives"),
  restoreDiffPreviewPool: document.querySelector("#restore-diff-preview-pool"),
  restoreDiffPreviewEligible: document.querySelector("#restore-diff-preview-eligible"),
  restoreDiffApply: document.querySelector("#restore-diff-apply"),
  restoreDiffCancel: document.querySelector("#restore-diff-cancel"),
  excludedCount: document.querySelector("#excluded-count"),
  filteredCount: document.querySelector("#filtered-count"),
  generateExams: document.querySelector("#generate-exams"),
  generatorStatus: document.querySelector("#generator-status"),
  generationSeed: document.querySelector("#generation-seed"),
  includedCount: document.querySelector("#included-count"),
  objectiveFilters: document.querySelector("#objective-filters"),
  poolSearch: document.querySelector("#pool-search"),
  poolSearchSummary: document.querySelector("#pool-search-summary"),
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
  const examExcludeIds = state.derivedExcludeQuestionIds;

  return state.quiz.questions.filter((question) => {
    if (excludeIds.has(question.id)) {
      return false;
    }
    if (includeIds.has(question.id)) {
      return true;
    }
    if (examExcludeIds.has(question.id)) {
      return false;
    }
    return filteredIds.has(question.id);
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
  if (state.derivedExcludeQuestionIds.has(question.id)) {
    return { label: "Excluded (Prev Exam)", tone: "exclude" };
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
  const manualExcludes = new Set(selectedOverrideIds("exclude"));
  const manualIncludes = new Set(selectedOverrideIds("include"));
  let examExcludeCount = 0;
  for (const id of state.derivedExcludeQuestionIds) {
    if (!manualExcludes.has(id) && !manualIncludes.has(id)) {
      examExcludeCount += 1;
    }
  }
  elements.excludedCount.textContent = String(manualExcludes.size + examExcludeCount);
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
  const query = state.poolSearch.trim();
  const totalQuestions = state.quiz.questions.length;
  const scored = state.quiz.questions.map((question) => {
    const haystack = buildQuestionSearchText(question, { objectiveLabel });
    const score = query ? fuzzyQueryScore(haystack, query) : 1;
    return { question, score };
  });
  const filtered = query ? scored.filter((entry) => entry.score > 0) : scored;
  const scoreById = new Map(filtered.map((entry) => [entry.question.id, entry.score]));
  const sortedQuestions = filtered
    .map((entry) => entry.question)
    .sort((left, right) => {
      if (query) {
        const scoreDelta = (scoreById.get(right.id) ?? 0) - (scoreById.get(left.id) ?? 0);
        if (scoreDelta !== 0) {
          return scoreDelta;
        }
      }
      const leftStatus = rowStatus(left);
      const rightStatus = rowStatus(right);
      const rankDelta = statusSortRank(leftStatus.tone) - statusSortRank(rightStatus.tone);
      if (rankDelta !== 0) {
        return rankDelta;
      }
      return String(left.id).localeCompare(String(right.id));
    });

  if (elements.poolSearchSummary) {
    if (!query) {
      elements.poolSearchSummary.textContent = `${totalQuestions} question${totalQuestions === 1 ? "" : "s"}`;
    } else {
      elements.poolSearchSummary.textContent = `${sortedQuestions.length} of ${totalQuestions} match`;
    }
  }

  if (sortedQuestions.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 9;
    cell.className = "cell-copy";
    cell.textContent = query ? "No questions match this search." : "No questions in this pool.";
    row.append(cell);
    fragment.append(row);
    elements.poolTableBody.replaceChildren(fragment);
    return;
  }

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

function examSetOptionLabel(summary) {
  const examName = summary.printSettings?.examName || summary.quiz?.title || "Untitled Exam";
  const when = summary.generatedAt ? new Date(summary.generatedAt).toLocaleString() : "Unknown time";
  const count = Number.isFinite(summary.selectedQuestionCount) ? summary.selectedQuestionCount : 0;
  return `${examName} · ${when} · ${count}q`;
}

function renderPreviousExamSelect() {
  const excludeFragment = document.createDocumentFragment();
  const excludeNone = document.createElement("option");
  excludeNone.value = "";
  excludeNone.textContent = "None";
  excludeFragment.append(excludeNone);

  const restoreFragment = document.createDocumentFragment();
  const restoreNone = document.createElement("option");
  restoreNone.value = "";
  restoreNone.textContent = "Select to restore...";
  restoreFragment.append(restoreNone);

  let hasSelectedMatch = false;
  for (const summary of state.previousExams) {
    const label = `${summary.examSetId} — ${examSetOptionLabel(summary)}`;

    const excludeOption = document.createElement("option");
    excludeOption.value = summary.examSetId;
    excludeOption.textContent = label;
    if (summary.examSetId === state.selection.excludeFromExamSetId) {
      excludeOption.selected = true;
      hasSelectedMatch = true;
    }
    excludeFragment.append(excludeOption);

    const restoreOption = document.createElement("option");
    restoreOption.value = summary.examSetId;
    restoreOption.textContent = label;
    restoreFragment.append(restoreOption);
  }
  elements.excludeFromExam.replaceChildren(excludeFragment);
  if (!hasSelectedMatch) {
    elements.excludeFromExam.value = "";
  }
  elements.restoreFromExam.replaceChildren(restoreFragment);
  elements.restoreFromExam.value = "";
}

let pendingRestore = null;

function questionMatchesSavedFilters(question, savedSelection) {
  const sources = Array.isArray(savedSelection.sources) ? savedSelection.sources : [];
  if (sources.length > 0) {
    const qs = questionSources(question);
    if (!qs.some((source) => sources.includes(source))) {
      return false;
    }
  }
  const difficulties = Array.isArray(savedSelection.difficulties) ? savedSelection.difficulties : [];
  if (difficulties.length > 0) {
    if (!difficulties.includes(question.difficulty)) {
      return false;
    }
  }
  const objectiveIds = Array.isArray(savedSelection.learningObjectiveIds) ? savedSelection.learningObjectiveIds : [];
  if (objectiveIds.length > 0) {
    const qObjectives = Array.isArray(question.learningObjectiveIds) ? question.learningObjectiveIds : [];
    if (!qObjectives.some((id) => objectiveIds.includes(id))) {
      return false;
    }
  }
  return true;
}

function normalizeCurrentForDiff(question) {
  if (!question) return null;
  return {
    id: question.id,
    question: question.question ?? "",
    choices: (question.choices || []).map((choice) => ({
      key: choice?.key ?? "",
      text: choice?.text ?? "",
    })),
    correctAnswers: (question.correctAnswers || []).map(String),
    explanation: question.explanation ?? "",
    difficulty: question.difficulty,
    points: question.points,
    shuffleChoices: Boolean(question.shuffleChoices),
    learningObjectiveIds: (question.learningObjectiveIds || []).map(String),
    imageAssetIds: (question.imageAssetIds || []).map(String),
    locations: Array.isArray(question.locations)
      ? question.locations
      : Array.isArray(question.bookLocations) ? question.bookLocations : [],
  };
}

const DIFF_FIELDS = [
  "question",
  "choices",
  "correctAnswers",
  "explanation",
  "difficulty",
  "points",
  "shuffleChoices",
  "learningObjectiveIds",
  "imageAssetIds",
  "locations",
];

function fieldEquals(a, b) {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

function computeChangedFields(vintage, current) {
  const changed = [];
  for (const field of DIFF_FIELDS) {
    if (!fieldEquals(vintage[field], current[field])) {
      changed.push(field);
    }
  }
  return changed;
}

function computeRestoreDiff(selection) {
  const currentQuestions = state.quiz.questions;
  const currentById = new Map(currentQuestions.map((question) => [question.id, question]));
  const currentIds = new Set(currentById.keys());

  const hasSnapshot = Array.isArray(selection.poolQuestionIds) && selection.poolQuestionIds.length > 0;
  const vintageHashes = selection.poolQuestionHashes && typeof selection.poolQuestionHashes === "object"
    ? selection.poolQuestionHashes
    : null;
  const hashesAvailable = Boolean(vintageHashes);
  const vintageSnapshot = Array.isArray(selection.poolSnapshot) ? selection.poolSnapshot : null;
  const snapshotAvailable = Boolean(vintageSnapshot);
  const vintageSnapshotById = snapshotAvailable
    ? new Map(vintageSnapshot.map((snap) => [snap.id, snap]))
    : new Map();

  let addedIds = [];
  let removedIds = [];
  let addedScope = "exam";

  if (hasSnapshot) {
    const vintageIds = new Set(selection.poolQuestionIds);
    addedIds = [...currentIds].filter((id) => !vintageIds.has(id)).sort();
    removedIds = [...vintageIds].filter((id) => !currentIds.has(id)).sort();
    addedScope = "pool";
  } else {
    const referenced = new Set();
    for (const key of ["availableQuestionIds", "filteredQuestionIds", "selectedQuestionIds", "includeQuestionIds", "excludeQuestionIds"]) {
      const list = Array.isArray(selection[key]) ? selection[key] : [];
      for (const id of list) {
        referenced.add(id);
      }
    }
    removedIds = [...referenced].filter((id) => !currentIds.has(id)).sort();

    const savedFiltered = new Set(Array.isArray(selection.filteredQuestionIds) ? selection.filteredQuestionIds : []);
    const currentInScope = currentQuestions.filter((question) => questionMatchesSavedFilters(question, selection));
    addedIds = currentInScope
      .map((question) => question.id)
      .filter((id) => !savedFiltered.has(id))
      .sort();
  }

  const addedRows = addedIds.map((id) => ({
    id,
    question: currentById.get(id) || null,
    decision: "allow",
  }));

  const removedRows = removedIds.map((id) => ({
    id,
    vintageSnapshot: vintageSnapshotById.get(id) || null,
  }));

  const modifiedRows = [];
  const currentHashById = new Map(
    Object.entries(state.questionHashes || {}).filter(([, hash]) => typeof hash === "string" && hash !== ""),
  );
  if (hashesAvailable) {
    for (const [id, hash] of Object.entries(vintageHashes)) {
      if (!currentIds.has(id)) continue;
      const currentHash = currentHashById.get(id);
      if (!currentHash || currentHash === hash) continue;

      const currentQuestion = currentById.get(id);
      const vintage = vintageSnapshotById.get(id) || null;
      const currentNormalized = normalizeCurrentForDiff(currentQuestion);
      const changedFields = vintage && currentNormalized
        ? computeChangedFields(vintage, currentNormalized)
        : [];
      modifiedRows.push({
        id,
        currentQuestion,
        currentNormalized,
        vintageSnapshot: vintage,
        changedFields,
        decision: "keep",
      });
    }
    modifiedRows.sort((left, right) => String(left.id).localeCompare(String(right.id)));
  }

  const currentScopeCount = currentQuestions.filter((question) => questionMatchesSavedFilters(question, selection)).length;

  return {
    vintageSource: hasSnapshot ? "snapshot" : "fallback",
    hashesAvailable,
    snapshotAvailable,
    added: addedIds,
    removed: removedIds,
    modified: modifiedRows.map((row) => row.id),
    addedRows,
    removedRows,
    modifiedRows,
    addedScope,
    currentScopeCount,
    currentPoolCount: currentQuestions.length,
  };
}

function closeRestoreDiffModal() {
  pendingRestore = null;
  elements.restoreDiffModal.classList.remove("is-open");
  elements.restoreDiffModal.setAttribute("aria-hidden", "true");
  document.body.style.overflow = "";
}

function describeFilterList(values, fallback) {
  if (!Array.isArray(values) || values.length === 0) {
    return fallback;
  }
  if (values.length <= 6) {
    return values.join(", ");
  }
  return `${values.slice(0, 6).join(", ")} (+${values.length - 6} more)`;
}

function describeObjectives(objectiveIds) {
  if (!Array.isArray(objectiveIds) || objectiveIds.length === 0) {
    return "Any objective";
  }
  const labels = objectiveIds.map((id) => {
    const label = objectiveLabel(id);
    return label && label !== id ? `${id} · ${label}` : id;
  });
  return describeFilterList(labels, "Any objective");
}

const DIFF_FIELD_LABELS = {
  question: "Prompt",
  choices: "Choices",
  correctAnswers: "Correct answers",
  explanation: "Explanation",
  difficulty: "Difficulty",
  points: "Points",
  shuffleChoices: "Shuffle choices",
  learningObjectiveIds: "Objectives",
  imageAssetIds: "Images",
  locations: "Sources",
};

function diffValueToText(field, value) {
  if (value === null || value === undefined) return "—";
  if (field === "choices" && Array.isArray(value)) {
    if (value.length === 0) return "—";
    return value.map((choice) => `${choice.key ?? ""}: ${stripRichTextMarkup(choice.text ?? "")}`).join(" | ");
  }
  if (field === "locations" && Array.isArray(value)) {
    if (value.length === 0) return "—";
    return value
      .map((location) => {
        const chapter = location?.chapter || location?.source || "";
        const section = location?.section || "";
        const page = location?.page ? `Page ${location.page}` : "";
        const url = location?.url || "";
        return [chapter, section, page, url].filter(Boolean).join(" · ") || "—";
      })
      .join(" | ");
  }
  if (field === "correctAnswers" && Array.isArray(value)) {
    return value.length ? value.join(", ") : "—";
  }
  if (field === "learningObjectiveIds" && Array.isArray(value)) {
    return value.length ? value.join(", ") : "—";
  }
  if (field === "imageAssetIds" && Array.isArray(value)) {
    return value.length ? value.join(", ") : "—";
  }
  if (field === "shuffleChoices") {
    return value ? "Yes" : "No";
  }
  if (field === "question" || field === "explanation") {
    return stripRichTextMarkup(String(value)) || "—";
  }
  if (Array.isArray(value)) return value.length ? JSON.stringify(value) : "—";
  return String(value);
}

function shortSources(question) {
  if (!question) return "—";
  return questionSources(question).join(", ") || "—";
}

function shortVintageSources(snapshot) {
  if (!snapshot || !Array.isArray(snapshot.locations) || snapshot.locations.length === 0) return "—";
  return snapshot.locations
    .map((location) => location?.chapter || location?.source || "")
    .filter(Boolean)
    .join(", ") || "—";
}

function shortObjectives(objectiveIds) {
  if (!Array.isArray(objectiveIds) || objectiveIds.length === 0) return "—";
  return objectiveIds
    .map((id) => {
      const label = objectiveLabel(id);
      return label && label !== id ? `${id} · ${truncate(label, 28)}` : id;
    })
    .join(", ");
}

function createPerRowSelect(decision, optionPairs, onChange) {
  const select = document.createElement("select");
  select.className = "inline-select";
  for (const [value, label] of optionPairs) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    if (value === decision) option.selected = true;
    select.append(option);
  }
  select.addEventListener("change", (event) => onChange(event.target.value));
  return select;
}

function createAddedRow(row, onDecisionChange) {
  const tr = document.createElement("tr");
  const idCell = document.createElement("td");
  idCell.className = "cell-mono";
  idCell.textContent = row.id;

  const promptCell = document.createElement("td");
  promptCell.className = "cell-copy";
  promptCell.textContent = truncate(stripRichTextMarkup(row.question?.question ?? ""), 140) || "—";

  const difficultyCell = document.createElement("td");
  difficultyCell.textContent = row.question?.difficulty ?? "—";

  const pointsCell = document.createElement("td");
  pointsCell.textContent = row.question?.points ?? "—";

  const sourcesCell = document.createElement("td");
  sourcesCell.className = "cell-copy";
  sourcesCell.textContent = shortSources(row.question);

  const objectivesCell = document.createElement("td");
  objectivesCell.className = "cell-copy";
  objectivesCell.textContent = shortObjectives(row.question?.learningObjectiveIds);

  const actionCell = document.createElement("td");
  actionCell.append(createPerRowSelect(row.decision, [["allow", "Allow"], ["exclude", "Force-exclude"]], onDecisionChange));

  tr.append(idCell, promptCell, difficultyCell, pointsCell, sourcesCell, objectivesCell, actionCell);
  return tr;
}

function createDiffPane(row) {
  const wrap = document.createElement("div");
  wrap.className = "diff-pane";

  if (!row.vintageSnapshot) {
    const note = document.createElement("p");
    note.className = "helper-copy";
    note.textContent = "No vintage snapshot recorded — cannot show before/after.";
    wrap.append(note);
    return wrap;
  }

  for (const field of row.changedFields) {
    const label = DIFF_FIELD_LABELS[field] || field;
    const fieldRow = document.createElement("div");
    fieldRow.className = "diff-row";

    const fieldName = document.createElement("div");
    fieldName.className = "diff-row__field";
    fieldName.textContent = label;

    const before = document.createElement("div");
    before.className = "diff-row__was";
    before.textContent = diffValueToText(field, row.vintageSnapshot[field]);

    const after = document.createElement("div");
    after.className = "diff-row__now";
    after.textContent = diffValueToText(field, row.currentNormalized?.[field]);

    fieldRow.append(fieldName, before, after);
    wrap.append(fieldRow);
  }
  return wrap;
}

function createModifiedRow(row, onDecisionChange) {
  const tr = document.createElement("tr");

  const idCell = document.createElement("td");
  idCell.className = "cell-mono";
  idCell.textContent = row.id;

  const fieldsCell = document.createElement("td");
  fieldsCell.className = "cell-copy";
  if (row.changedFields.length === 0) {
    fieldsCell.textContent = "(content changed — fields unknown)";
  } else {
    for (const field of row.changedFields) {
      const badge = document.createElement("span");
      badge.className = "status-badge status-badge--filtered diff-field-badge";
      badge.textContent = DIFF_FIELD_LABELS[field] || field;
      fieldsCell.append(badge);
    }
  }

  const promptCell = document.createElement("td");
  promptCell.className = "cell-copy";
  promptCell.textContent = truncate(stripRichTextMarkup(row.currentQuestion?.question ?? ""), 120) || "—";

  const diffCell = document.createElement("td");
  if (row.vintageSnapshot && row.changedFields.length > 0) {
    const details = document.createElement("details");
    details.className = "diff-details";
    const summary = document.createElement("summary");
    summary.textContent = "Show diff";
    details.append(summary);
    details.append(createDiffPane(row));
    diffCell.append(details);
  } else if (!row.vintageSnapshot) {
    diffCell.textContent = "Vintage content not recorded";
    diffCell.className = "helper-copy";
  } else {
    diffCell.textContent = "—";
  }

  const actionCell = document.createElement("td");
  actionCell.append(createPerRowSelect(row.decision, [["keep", "Keep current"], ["exclude", "Force-exclude"]], onDecisionChange));

  tr.append(idCell, fieldsCell, promptCell, diffCell, actionCell);
  return tr;
}

function createRemovedRow(row) {
  const tr = document.createElement("tr");
  const idCell = document.createElement("td");
  idCell.className = "cell-mono";
  idCell.textContent = row.id;

  const snapshot = row.vintageSnapshot;
  const promptCell = document.createElement("td");
  promptCell.className = "cell-copy";
  promptCell.textContent = snapshot
    ? truncate(stripRichTextMarkup(snapshot.question ?? ""), 140) || "—"
    : "Vintage content not recorded";

  const difficultyCell = document.createElement("td");
  difficultyCell.textContent = snapshot?.difficulty ?? "—";

  const pointsCell = document.createElement("td");
  pointsCell.textContent = snapshot?.points ?? "—";

  const sourcesCell = document.createElement("td");
  sourcesCell.className = "cell-copy";
  sourcesCell.textContent = shortVintageSources(snapshot);

  const objectivesCell = document.createElement("td");
  objectivesCell.className = "cell-copy";
  objectivesCell.textContent = shortObjectives(snapshot?.learningObjectiveIds);

  tr.append(idCell, promptCell, difficultyCell, pointsCell, sourcesCell, objectivesCell);
  return tr;
}

function renderAddedRows() {
  const fragment = document.createDocumentFragment();
  for (const row of pendingRestore.addedRows) {
    fragment.append(createAddedRow(row, (value) => {
      row.decision = value;
    }));
  }
  elements.restoreDiffAddedBody.replaceChildren(fragment);
}

function renderModifiedRows() {
  const fragment = document.createDocumentFragment();
  for (const row of pendingRestore.modifiedRows) {
    fragment.append(createModifiedRow(row, (value) => {
      row.decision = value;
    }));
  }
  elements.restoreDiffModifiedBody.replaceChildren(fragment);
}

function renderRemovedRows() {
  const fragment = document.createDocumentFragment();
  for (const row of pendingRestore.removedRows) {
    fragment.append(createRemovedRow(row));
  }
  elements.restoreDiffRemovedBody.replaceChildren(fragment);
}

function applyBulkDecision(target, decision) {
  if (!pendingRestore) return;
  if (target === "added") {
    for (const row of pendingRestore.addedRows) row.decision = decision;
    renderAddedRows();
  } else if (target === "modified") {
    for (const row of pendingRestore.modifiedRows) row.decision = decision;
    renderModifiedRows();
  }
}

function openRestoreDiffModal(examSetId, summary, selection, diff) {
  const {
    addedRows,
    modifiedRows,
    removedRows,
    added,
    removed,
    modified,
    vintageSource,
    hashesAvailable,
    snapshotAvailable,
    addedScope,
    currentScopeCount,
    currentPoolCount,
  } = diff;
  pendingRestore = {
    examSetId,
    selection,
    addedRows,
    modifiedRows,
    removedRows,
    added,
    removed,
    modified,
    vintageSource,
    hashesAvailable,
    snapshotAvailable,
  };

  const summaryParts = [];
  summaryParts.push(`Restoring from <strong>${escapeHtml(examSetId)}</strong>.`);
  if (vintageSource === "snapshot") {
    summaryParts.push("This exam recorded a full pool snapshot, so the diff below is exact.");
  } else {
    summaryParts.push("This exam predates pool snapshots, so its diff is reconstructed from the filter scope it was generated under (less precise — see Detection mode below).");
  }
  elements.restoreDiffSummary.innerHTML = summaryParts.join(" ");

  const generatedAt = summary?.generatedAt || selection.generatedAt || "";
  elements.restoreDiffPreviewWhen.textContent = generatedAt ? new Date(generatedAt).toLocaleString() : "Unknown";
  elements.restoreDiffPreviewSeed.textContent = (selection.generationSeed && String(selection.generationSeed)) || "—";
  elements.restoreDiffPreviewCount.textContent = String(selection.questionCount ?? "—");
  elements.restoreDiffPreviewVariants.textContent = String(selection.variantCount ?? "—");
  elements.restoreDiffPreviewSources.textContent = describeFilterList(selection.sources, "Any source");
  elements.restoreDiffPreviewDifficulties.textContent = describeFilterList(
    Array.isArray(selection.difficulties) ? selection.difficulties.map(String) : [],
    "Any difficulty",
  );
  elements.restoreDiffPreviewObjectives.textContent = describeObjectives(selection.learningObjectiveIds);
  const vintagePoolSize = Array.isArray(selection.poolQuestionIds) ? selection.poolQuestionIds.length : null;
  elements.restoreDiffPreviewPool.textContent = vintagePoolSize !== null
    ? `${vintagePoolSize} questions then · ${currentPoolCount} now`
    : `Unknown then · ${currentPoolCount} now`;
  const vintageAvailable = Array.isArray(selection.availableQuestionIds) ? selection.availableQuestionIds.length : null;
  const vintageSelected = Array.isArray(selection.selectedQuestionIds) ? selection.selectedQuestionIds.length : null;
  const eligibleParts = [];
  if (vintageAvailable !== null) eligibleParts.push(`${vintageAvailable} eligible then`);
  eligibleParts.push(`${currentScopeCount} match the saved filters now`);
  if (vintageSelected !== null) eligibleParts.push(`${vintageSelected} chosen for the exam`);
  elements.restoreDiffPreviewEligible.textContent = eligibleParts.join(" · ");

  const detectionLines = [];
  if (vintageSource === "snapshot") {
    detectionLines.push("Pool changes: exact. Additions and removals are detected against the full pool snapshot saved with this exam.");
  } else {
    detectionLines.push("Pool changes: scope-restricted. We re-apply the exam's saved filters to today's pool and compare against the filtered IDs that were recorded. Questions outside the saved filter scope are ignored here, even if they were added to the pool since.");
  }
  if (hashesAvailable) {
    detectionLines.push("Edits: exact. Each question's current content hash is compared against the hash recorded when this exam was generated.");
  } else {
    detectionLines.push("Edits: unavailable. Content hashes were not recorded for this exam, so we can't tell which questions have been edited. Edited questions will silently use their current content if you allow them.");
  }
  elements.restoreDiffDetection.textContent = detectionLines.join(" ");

  const addedHeadingText = addedScope === "pool"
    ? "New questions added to the pool"
    : "Newly eligible questions (filter-scope estimate)";
  elements.restoreDiffAddedHeading.firstChild.nodeValue = `${addedHeadingText} `;
  elements.restoreDiffAddedDescription.textContent = addedScope === "pool"
    ? "These IDs exist in the pool today but were not in the pool when this exam was generated."
    : "These questions match this exam's saved filters today but were not in its recorded filter result. They are most likely additions to the pool, but could also include questions whose metadata was edited so they now match the filters.";

  elements.restoreDiffAddedCount.textContent = `(${added.length})`;
  elements.restoreDiffRemovedCount.textContent = `(${removed.length})`;
  elements.restoreDiffModifiedCount.textContent = hashesAvailable ? `(${modified.length})` : "(unavailable)";

  elements.restoreDiffAddedBlock.hidden = added.length === 0;
  elements.restoreDiffRemovedBlock.hidden = removed.length === 0;
  elements.restoreDiffModifiedBlock.hidden = !hashesAvailable || modified.length === 0;

  renderAddedRows();
  renderModifiedRows();
  renderRemovedRows();

  elements.restoreDiffModal.classList.add("is-open");
  elements.restoreDiffModal.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
}

function applyRestoredSelection(examSetId, selection, applyOptions) {
  const {
    strictExclusionIds = [],
    excludeModifiedIds = [],
    addedCount = 0,
    removedCount = 0,
    modifiedCount = 0,
    keptEditsCount = 0,
    vintageSource,
    hashesAvailable = true,
  } = applyOptions;
  const filterOpts = filterOptions();
  const sourceSet = new Set(filterOpts.sources);
  const difficultySet = new Set(filterOpts.difficulties);
  const objectiveSet = new Set(filterOpts.learningObjectives.map((objective) => objective.id));
  const questionIdSet = new Set(state.quiz.questions.map((question) => question.id));

  const restoredCount = Number.parseInt(selection.questionCount, 10);
  state.selection.questionCount = Math.min(
    MAX_QUESTIONS_PER_EXAM,
    Math.max(1, Number.isFinite(restoredCount) ? restoredCount : state.selection.questionCount),
  );
  const restoredVariants = Number.parseInt(selection.variantCount, 10);
  state.selection.variantCount = Math.max(
    1,
    Number.isFinite(restoredVariants) ? restoredVariants : state.selection.variantCount,
  );
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
  const includeIds = Array.isArray(selection.includeQuestionIds) ? selection.includeQuestionIds : [];
  const excludeIds = Array.isArray(selection.excludeQuestionIds) ? selection.excludeQuestionIds : [];
  for (const id of includeIds) {
    if (questionIdSet.has(id)) {
      state.selection.overrides[id] = "include";
    }
  }
  for (const id of excludeIds) {
    if (questionIdSet.has(id)) {
      state.selection.overrides[id] = "exclude";
    }
  }
  let strictAppliedCount = 0;
  for (const id of strictExclusionIds) {
    if (questionIdSet.has(id) && state.selection.overrides[id] !== "include") {
      state.selection.overrides[id] = "exclude";
      strictAppliedCount += 1;
    }
  }
  let modifiedAppliedCount = 0;
  for (const id of excludeModifiedIds) {
    if (questionIdSet.has(id) && state.selection.overrides[id] !== "include") {
      state.selection.overrides[id] = "exclude";
      modifiedAppliedCount += 1;
    }
  }
  state.selection.excludeFromExamSetId = "";
  state.derivedExcludeQuestionIds = new Set();

  elements.questionCount.value = String(state.selection.questionCount);
  elements.variantCount.value = String(state.selection.variantCount);
  elements.generationSeed.value = state.selection.generationSeed;

  renderPreviousExamSelect();
  renderPoolState();
  scheduleDraftSave();

  const statusParts = [`Restored settings from ${examSetId}.`];
  if (strictAppliedCount > 0) {
    statusParts.push(`Force-excluded ${strictAppliedCount} new question${strictAppliedCount === 1 ? "" : "s"} (post-vintage).`);
  } else if (addedCount > 0) {
    statusParts.push(`${addedCount} newer question${addedCount === 1 ? "" : "s"} left eligible.`);
  }
  if (modifiedAppliedCount > 0) {
    statusParts.push(`Force-excluded ${modifiedAppliedCount} edited question${modifiedAppliedCount === 1 ? "" : "s"}.`);
  } else if (modifiedCount > 0) {
    statusParts.push(`${keptEditsCount || modifiedCount} edited question${(keptEditsCount || modifiedCount) === 1 ? " was" : "s were"} kept with current content.`);
  }
  if (removedCount > 0) {
    statusParts.push(`${removedCount} vintage question${removedCount === 1 ? "" : "s"} no longer in the pool were skipped.`);
  }
  if (vintageSource === "fallback") {
    statusParts.push("Added/removed diff was approximated (legacy exam without pool snapshot).");
  }
  if (!hashesAvailable) {
    statusParts.push("Edit detection was unavailable for this exam.");
  }
  setStatus(statusParts.join(" "));
}

async function restoreFromExamSet(examSetId) {
  if (!examSetId || !state.quiz) {
    return;
  }
  setStatus(`Restoring settings from ${examSetId}...`);
  let payload;
  try {
    const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}`);
    if (!response.ok) {
      setStatus(`Could not restore from ${examSetId}.`, true);
      return;
    }
    payload = await response.json();
  } catch (error) {
    console.error(error);
    setStatus(`Could not restore from ${examSetId}: ${error.message}`, true);
    return;
  }

  const selection = payload.examSet?.selection ?? {};
  const summary = payload.summary ?? payload.examSet ?? null;
  const diff = computeRestoreDiff(selection);

  if (diff.added.length === 0 && diff.removed.length === 0 && diff.modified.length === 0) {
    applyRestoredSelection(examSetId, selection, {
      addedCount: 0,
      removedCount: 0,
      modifiedCount: 0,
      vintageSource: diff.vintageSource,
      hashesAvailable: diff.hashesAvailable,
    });
    return;
  }

  openRestoreDiffModal(examSetId, summary, selection, diff);
}

async function loadPreviousExams() {
  try {
    const response = await fetch("/api/exams");
    if (!response.ok) {
      state.previousExams = [];
      return;
    }
    const payload = await response.json();
    state.previousExams = Array.isArray(payload.examSets) ? payload.examSets : [];
  } catch (error) {
    console.warn("Could not load previous exams", error);
    state.previousExams = [];
  }
}

async function refreshDerivedExcludes() {
  const examSetId = state.selection.excludeFromExamSetId;
  if (!examSetId) {
    state.derivedExcludeQuestionIds = new Set();
    return;
  }
  try {
    const response = await fetch(`/api/exams/set/${encodeURIComponent(examSetId)}`);
    if (!response.ok) {
      state.derivedExcludeQuestionIds = new Set();
      state.selection.excludeFromExamSetId = "";
      return;
    }
    const payload = await response.json();
    const ids = Array.isArray(payload.examSet?.selection?.selectedQuestionIds)
      ? payload.examSet.selection.selectedQuestionIds
      : [];
    const pool = new Set(state.quiz?.questions?.map((question) => question.id) ?? []);
    state.derivedExcludeQuestionIds = new Set(ids.filter((id) => pool.has(id)));
  } catch (error) {
    console.warn("Could not load previous exam questions", error);
    state.derivedExcludeQuestionIds = new Set();
  }
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
      excludeFromExamSetId: state.selection.excludeFromExamSetId,
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
  state.selection.excludeFromExamSetId = typeof selection.excludeFromExamSetId === "string"
    ? selection.excludeFromExamSetId.trim()
    : "";

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
  state.questionHashes = payload.questionHashes && typeof payload.questionHashes === "object"
    ? payload.questionHashes
    : {};
  state.dbPath = payload.dbPath;
  state.selection.questionCount = Math.max(1, Math.min(10, state.quiz.questions.length));
  state.selection.questionCount = Math.min(state.selection.questionCount, MAX_QUESTIONS_PER_EXAM);
  state.selection.variantCount = 1;
  const draft = await loadGeneratorDraft();
  applyGeneratorDraft(draft);
  await restoreGeneratedRunFromDraft(draft);
  await loadPreviousExams();
  if (state.selection.excludeFromExamSetId
    && !state.previousExams.some((summary) => summary.examSetId === state.selection.excludeFromExamSetId)) {
    state.selection.excludeFromExamSetId = "";
  }
  await refreshDerivedExcludes();
  elements.dbPath.textContent = payload.projectPath ?? state.dbPath;
  elements.questionCount.value = String(state.selection.questionCount);
  elements.variantCount.value = String(state.selection.variantCount);
  elements.generationSeed.value = state.selection.generationSeed;
  renderPreviousExamSelect();
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

  const manualIncludes = selectedOverrideIds("include");
  const manualIncludeSet = new Set(manualIncludes);
  const manualExcludes = selectedOverrideIds("exclude");
  const mergedExcludes = new Set(manualExcludes);
  for (const id of state.derivedExcludeQuestionIds) {
    if (!manualIncludeSet.has(id)) {
      mergedExcludes.add(id);
    }
  }

  const payload = {
    questionCount: Number.parseInt(elements.questionCount.value, 10),
    variantCount: Number.parseInt(elements.variantCount.value, 10),
    generationSeed: elements.generationSeed.value.trim(),
    sources: [...state.selection.sources],
    difficulties: [...state.selection.difficulties],
    learningObjectiveIds: [...state.selection.learningObjectiveIds],
    includeQuestionIds: manualIncludes,
    excludeQuestionIds: [...mergedExcludes],
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
  await loadPreviousExams();
  renderPreviousExamSelect();
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

  elements.poolSearch.addEventListener("input", (event) => {
    state.poolSearch = event.target.value;
    if (state.quiz) {
      renderPoolTable();
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

  elements.restoreFromExam.addEventListener("change", async (event) => {
    const examSetId = event.target.value;
    event.target.value = "";
    if (examSetId) {
      await restoreFromExamSet(examSetId);
    }
  });

  elements.excludeFromExam.addEventListener("change", async (event) => {
    state.selection.excludeFromExamSetId = event.target.value;
    await refreshDerivedExcludes();
    renderPreviousExamSelect();
    renderPoolState();
    scheduleDraftSave();
    if (state.selection.excludeFromExamSetId) {
      setStatus(`Excluding ${state.derivedExcludeQuestionIds.size} questions used in ${state.selection.excludeFromExamSetId}.`);
    } else {
      setStatus("Previous-exam exclusion cleared.");
    }
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

  elements.restoreDiffApply.addEventListener("click", () => {
    if (!pendingRestore) {
      closeRestoreDiffModal();
      return;
    }
    const { examSetId, selection, addedRows, modifiedRows, removedRows, vintageSource, hashesAvailable } = pendingRestore;
    const excludeAdded = addedRows.filter((row) => row.decision === "exclude").map((row) => row.id);
    const excludeModified = modifiedRows.filter((row) => row.decision === "exclude").map((row) => row.id);
    const keptEdits = modifiedRows.filter((row) => row.decision === "keep").length;
    closeRestoreDiffModal();
    applyRestoredSelection(examSetId, selection, {
      strictExclusionIds: excludeAdded,
      excludeModifiedIds: excludeModified,
      addedCount: addedRows.length,
      removedCount: removedRows.length,
      modifiedCount: modifiedRows.length,
      keptEditsCount: keptEdits,
      vintageSource,
      hashesAvailable,
    });
  });

  for (const button of elements.restoreDiffModal.querySelectorAll("[data-bulk]")) {
    button.addEventListener("click", () => {
      const target = button.dataset.bulk;
      const decision = button.dataset.decision;
      applyBulkDecision(target, decision);
    });
  }

  elements.restoreDiffCancel.addEventListener("click", () => {
    closeRestoreDiffModal();
    setStatus("Restore cancelled.");
  });

  elements.restoreDiffBackdrop.addEventListener("click", () => {
    closeRestoreDiffModal();
    setStatus("Restore cancelled.");
  });

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (pendingRestore) {
        closeRestoreDiffModal();
        setStatus("Restore cancelled.");
        return;
      }
      if (state.activePoolQuestionId) {
        closePoolQuestionModal();
      }
    }
  });
}

wireEvents();
loadQuiz().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
