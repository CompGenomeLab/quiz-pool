import { renderRichTextHtml } from "./rich-text.js";

const state = {
  annotationResults: null,
  busyMessage: "",
  busyTitle: "",
  dbPath: "",
  examStorePath: "",
  gradingFiles: [],
  annotateModalOpen: false,
  isAnnotating: false,
  isDeletingGradingRun: false,
  isGrading: false,
  results: null,
  savedRuns: [],
  selectedSavedRunId: "",
  gradingFormula: {
    mode: "none",
    wrongPenalty: 0,
  },
  detailModalOpen: false,
  correctSortDirection: "desc",
  selectedRowIndex: null,
  statusIsError: false,
  statusMessage: "Loading grading tools...",
  validationErrors: [],
};

const elements = {
  annotateGradedPdfs: document.querySelector("#annotate-graded-pdfs"),
  annotateModal: document.querySelector("#annotate-modal"),
  annotateModalBackdrop: document.querySelector("#annotate-modal-backdrop"),
  browseGradingFile: document.querySelector("#browse-grading-file"),
  browseGradingFolder: document.querySelector("#browse-grading-folder"),
  cancelAnnotateModal: document.querySelector("#cancel-annotate-modal"),
  confirmAnnotateModal: document.querySelector("#confirm-annotate-modal"),
  closeGradingDetail: document.querySelector("#close-grading-detail"),
  dbPath: document.querySelector("#db-path"),
  deleteGradingRun: document.querySelector("#delete-grading-run"),
  errorList: document.querySelector("#grading-error-list"),
  exportGradingCsv: document.querySelector("#export-grading-csv"),
  errorPanel: document.querySelector("#grading-errors"),
  examStorePath: document.querySelector("#exam-store-path"),
  gradingHeading: document.querySelector("#grading-heading"),
  gradingDuplicateCount: document.querySelector("#grading-duplicate-count"),
  gradingInputPath: document.querySelector("#grading-input-path"),
  gradingPdfUpload: document.querySelector("#grading-pdf-upload"),
  gradingFolderUpload: document.querySelector("#grading-folder-upload"),
  gradingKnownCount: document.querySelector("#grading-known-count"),
  gradingMismatchCount: document.querySelector("#grading-mismatch-count"),
  gradingOErrorCount: document.querySelector("#grading-omr-error-count"),
  gradingProcessedCount: document.querySelector("#grading-processed-count"),
  gradingResultPath: document.querySelector("#grading-result-path"),
  gradingResults: document.querySelector("#grading-results"),
  gradingStatus: document.querySelector("#grading-status"),
  gradingTableBody: document.querySelector("#grading-table-body"),
  gradingDetailList: document.querySelector("#grading-detail-list"),
  gradingDetailModal: document.querySelector("#grading-detail-modal"),
  gradingDetailBackdrop: document.querySelector("#grading-detail-backdrop"),
  gradingFormulaMode: document.querySelector("#grading-formula-mode"),
  gradingFixedPenalty: document.querySelector("#grading-fixed-penalty"),
  gradingFixedPenaltyField: document.querySelector("#grading-fixed-penalty-field"),
  gradingFormulaSummary: document.querySelector("#grading-formula-summary"),
  gradingObjectiveReportBody: document.querySelector("#grading-objective-report-body"),
  gradingTotalScore: document.querySelector("#grading-total-score"),
  gradingTotalWrong: document.querySelector("#grading-total-wrong"),
  recalculateGrading: document.querySelector("#recalculate-grading"),
  runGrading: document.querySelector("#run-grading"),
  savedGradingList: document.querySelector("#saved-grading-list"),
  savedGradingSelect: document.querySelector("#saved-grading-select"),
  sortCorrect: document.querySelector("#sort-correct"),
  busyCopy: document.querySelector("#grading-busy-copy"),
  busyOverlay: document.querySelector("#grading-busy-overlay"),
  busyTitle: document.querySelector("#grading-busy-title"),
};

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  elements.gradingStatus.textContent = message;
  elements.gradingStatus.style.color = isError ? "var(--danger-strong)" : "var(--muted)";
}

function renderBusyOverlay() {
  const isBusy = state.isGrading || state.isAnnotating;
  elements.busyOverlay.classList.toggle("hidden", !isBusy);
  elements.busyTitle.textContent = state.busyTitle || "Working";
  elements.busyCopy.textContent = state.busyMessage || "Please wait while the OMR process completes.";
  document.body.style.overflow = isBusy ? "hidden" : "";
}

function setBusyState(isBusy, title = "", message = "") {
  state.busyTitle = title;
  state.busyMessage = message;
  if (!isBusy) {
    state.isGrading = false;
    state.isAnnotating = false;
  }
  renderBusyOverlay();
}

function renderAnnotateModal() {
  elements.annotateModal.classList.toggle("is-open", state.annotateModalOpen);
  elements.annotateModal.setAttribute("aria-hidden", String(!state.annotateModalOpen));
}

function openAnnotateModal() {
  if (!state.results) {
    return;
  }
  state.annotateModalOpen = true;
  renderAnnotateModal();
  window.setTimeout(() => {
    elements.confirmAnnotateModal.focus();
  }, 0);
}

function closeAnnotateModal() {
  state.annotateModalOpen = false;
  renderAnnotateModal();
}

function renderGradingDetailModal() {
  const isOpen = state.detailModalOpen && Boolean(state.results);
  elements.gradingDetailModal.classList.toggle("is-open", isOpen);
  elements.gradingDetailModal.setAttribute("aria-hidden", String(!isOpen));
}

function openGradingDetail(rowIndex) {
  if (!state.results) {
    return;
  }
  state.selectedRowIndex = rowIndex;
  state.detailModalOpen = true;
  renderSummary();
  window.setTimeout(() => {
    elements.closeGradingDetail.focus();
  }, 0);
}

function closeGradingDetail() {
  state.detailModalOpen = false;
  renderGradingDetailModal();
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

function renderQuestionImageHtml(question = {}) {
  const imageAssetIds = Array.isArray(question.imageAssetIds)
    ? question.imageAssetIds.filter((assetId) => typeof assetId === "string" && assetId.trim() !== "")
    : [];
  if (imageAssetIds.length === 0) {
    return "";
  }
  return `
    <div class="question-image-list">
      ${imageAssetIds.map((assetId, index) => `
        <img class="question-image-preview" src="${assetUrl(assetId)}" alt="Question image ${index + 1}" loading="lazy" />
      `).join("")}
    </div>
  `;
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

function statusTone(row) {
  if (row.status === "omr_error") return "exclude";
  if (row.status === "mismatch") return "filtered";
  return "eligible";
}

function statusLabel(row) {
  if (row.status === "omr_error") return "OMR Error";
  if (row.status === "mismatch") return "Needs Review";
  return "Ready";
}

function formatScore(value) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number)) {
    return "0";
  }
  if (Number.isInteger(number)) {
    return String(number);
  }
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: 3,
    minimumFractionDigits: 0,
  }).format(number);
}

function gradingFormulaDescription(formula = {}) {
  if (formula.description) {
    return formula.description;
  }
  if (formula.mode === "fixed") {
    return `Correct answers earn question points; incorrect or invalid answers subtract ${formatScore(formula.wrongPenalty)} point(s).`;
  }
  if (formula.mode === "choice_weighted") {
    return "Correct answers earn question points; incorrect or invalid answers subtract question points divided by answer choices minus 1.";
  }
  return "Correct answers earn question points; blank, incorrect, and invalid answers receive no penalty.";
}

function currentGradingFormula() {
  const mode = elements.gradingFormulaMode.value;
  const wrongPenalty = Math.max(0, Number.parseFloat(elements.gradingFixedPenalty.value || "0"));
  return {
    mode,
    wrongPenalty: Number.isFinite(wrongPenalty) ? wrongPenalty : 0,
  };
}

function populateGradingFormula(formula = {}) {
  const mode = ["none", "fixed", "choice_weighted"].includes(formula.mode) ? formula.mode : "none";
  elements.gradingFormulaMode.value = mode;
  elements.gradingFixedPenalty.value = String(formula.wrongPenalty ?? 0);
  state.gradingFormula = currentGradingFormula();
  renderFormulaControls();
}

function renderFormulaControls() {
  const mode = elements.gradingFormulaMode.value;
  const isFixed = mode === "fixed";
  elements.gradingFixedPenaltyField.classList.toggle("hidden", !isFixed);
  elements.gradingFixedPenalty.disabled = !isFixed || state.isGrading || state.isAnnotating;
}

function objectiveSummaryText(items = []) {
  if (!Array.isArray(items) || items.length === 0) {
    return "—";
  }
  return items.map((item) => {
    const label = item.id ? `${item.id}` : "Objective";
    return `${label}: ${formatScore(item.earnedPoints)}/${formatScore(item.possiblePoints)} C${item.correctCount ?? 0} W${item.wrongCount ?? 0}`;
  }).join(" | ");
}

function objectiveBlankMissingCount(item = {}) {
  return Number(item.blankCount ?? 0) + Number(item.missingCount ?? 0);
}

function createObjectiveMetric(label, value, tone) {
  const metric = document.createElement("span");
  metric.className = `objective-chip__metric objective-chip__metric--${tone}`;

  const metricLabel = document.createElement("span");
  metricLabel.className = "objective-chip__metric-label";
  metricLabel.textContent = label;

  const metricValue = document.createElement("strong");
  metricValue.textContent = value;

  metric.append(metricLabel, metricValue);
  return metric;
}

function createObjectiveSummaryNode(items = []) {
  const wrap = document.createElement("div");
  wrap.className = "objective-chip-list";
  if (!Array.isArray(items) || items.length === 0) {
    const empty = document.createElement("span");
    empty.className = "objective-chip-list__empty";
    empty.textContent = "—";
    wrap.append(empty);
    return wrap;
  }

  for (const item of items) {
    const chip = document.createElement("span");
    chip.className = "objective-chip";

    const name = document.createElement("strong");
    name.className = "objective-chip__name";
    name.textContent = item.id || "Objective";

    const score = createObjectiveMetric(
      "Score",
      `${formatScore(item.earnedPoints)}/${formatScore(item.possiblePoints)}`,
      "score",
    );
    const correct = createObjectiveMetric("C", String(item.correctCount ?? 0), "correct");
    const wrong = createObjectiveMetric("W", String(item.wrongCount ?? 0), "wrong");
    chip.append(name, score, correct, wrong);

    const blankMissing = objectiveBlankMissingCount(item);
    if (blankMissing > 0) {
      chip.append(createObjectiveMetric("B/M", String(blankMissing), "blank"));
    }

    wrap.append(chip);
  }
  return wrap;
}

function createCompactObjectiveSummaryNode(items = []) {
  const wrap = document.createElement("div");
  wrap.className = "objective-chip-list objective-chip-list--compact";
  if (!Array.isArray(items) || items.length === 0) {
    const empty = document.createElement("span");
    empty.className = "objective-chip-list__empty";
    empty.textContent = "—";
    wrap.append(empty);
    return wrap;
  }

  for (const item of items) {
    const chip = document.createElement("span");
    chip.className = "objective-compact-chip";
    chip.title = `${item.id || "Objective"}: ${formatScore(item.earnedPoints)}/${formatScore(item.possiblePoints)}; C${item.correctCount ?? 0}; W${item.wrongCount ?? 0}; B/M${objectiveBlankMissingCount(item)}`;

    const name = document.createElement("strong");
    name.className = "objective-compact-chip__name";
    name.textContent = item.id || "Obj";

    const score = document.createElement("span");
    score.className = "objective-compact-chip__score";
    score.textContent = `${formatScore(item.earnedPoints)}/${formatScore(item.possiblePoints)}`;

    const correct = document.createElement("span");
    correct.className = "objective-compact-chip__metric objective-compact-chip__metric--correct";
    correct.textContent = `C${item.correctCount ?? 0}`;

    const wrong = document.createElement("span");
    wrong.className = "objective-compact-chip__metric objective-compact-chip__metric--wrong";
    wrong.textContent = `W${item.wrongCount ?? 0}`;

    chip.append(name, score, correct, wrong);

    const blankMissing = objectiveBlankMissingCount(item);
    if (blankMissing > 0) {
      const blank = document.createElement("span");
      blank.className = "objective-compact-chip__metric objective-compact-chip__metric--blank";
      blank.textContent = `B${blankMissing}`;
      chip.append(blank);
    }

    wrap.append(chip);
  }
  return wrap;
}

function createScorePill(value, tone) {
  const pill = document.createElement("span");
  pill.className = `score-pill score-pill--${tone}`;
  pill.textContent = value;
  return pill;
}

function sourcePdfUrl(row) {
  if (!state.results?.gradingRunId || !row.rowIndex) {
    return "";
  }
  return `/api/gradings/source/${encodeURIComponent(state.results.gradingRunId)}/${encodeURIComponent(row.rowIndex)}.pdf`;
}

function csvCell(value) {
  const normalized = String(value ?? "");
  if (normalized.includes('"') || normalized.includes(",") || normalized.includes("\n") || normalized.includes("\r")) {
    return `"${normalized.replaceAll('"', '""')}"`;
  }
  return normalized;
}

function scoreText(earnedPoints, possiblePoints) {
  return `${formatScore(earnedPoints)}/${formatScore(possiblePoints)}`;
}

function countText(...values) {
  return String(values.reduce((total, value) => total + Number(value ?? 0), 0));
}

function getSortedRows(rows) {
  return [...rows].sort((left, right) => {
    const scoreDelta = right.summary.correctCount - left.summary.correctCount;
    if (scoreDelta !== 0) {
      return state.correctSortDirection === "desc" ? scoreDelta : -scoreDelta;
    }
    const pointsDelta = right.summary.earnedPoints - left.summary.earnedPoints;
    if (pointsDelta !== 0) {
      return state.correctSortDirection === "desc" ? pointsDelta : -pointsDelta;
    }
    return String(left.displayStudentId).localeCompare(String(right.displayStudentId))
      || String(left.sourcePdf || "").localeCompare(String(right.sourcePdf || ""));
  });
}

function updateSortButton() {
  const directionLabel = state.correctSortDirection === "desc" ? "highest first" : "lowest first";
  elements.sortCorrect.textContent = `Correct ${state.correctSortDirection === "desc" ? "↓" : "↑"}`;
  elements.sortCorrect.setAttribute("aria-label", `Sort by correct answers, ${directionLabel}`);
}

function selectedPdfFiles(fileList) {
  return [...(fileList ?? [])]
    .filter((file) => file.name.toLowerCase().endsWith(".pdf"))
    .sort((left, right) => {
      const leftName = left.webkitRelativePath || left.name;
      const rightName = right.webkitRelativePath || right.name;
      return leftName.localeCompare(rightName);
    });
}

function formatFileSelection(files) {
  if (files.length === 0) {
    return "";
  }
  if (files.length === 1) {
    return files[0].webkitRelativePath || files[0].name;
  }
  const sample = files.slice(0, 2).map((file) => file.name).join(", ");
  return files.length === 2 ? sample : `${files.length} PDFs selected (${sample}, ...)`;
}

function setGradingFiles(files) {
  state.gradingFiles = files;
  state.results = null;
  state.annotationResults = null;
  state.validationErrors = [];
  state.selectedSavedRunId = "";
  elements.gradingInputPath.value = formatFileSelection(files);
  renderErrors();
  renderSummary();
  renderSavedRuns();
  setStatus(files.length ? `${files.length} PDF${files.length === 1 ? "" : "s"} selected.` : "No PDFs selected.");
}

function gradingRunLabel(summary) {
  const when = summary.gradedAt ? new Date(summary.gradedAt).toLocaleString() : "Unknown time";
  const score = `${formatScore(summary.earnedPoints)}/${formatScore(summary.possiblePoints)}`;
  return `${when} · ${summary.processedCount ?? 0} PDF(s) · ${score}`;
}

function renderSavedRuns() {
  const selectFragment = document.createDocumentFragment();
  if (state.savedRuns.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No saved grading runs";
    selectFragment.append(option);
  } else {
    for (const summary of state.savedRuns) {
      const option = document.createElement("option");
      option.value = summary.gradingRunId;
      option.textContent = gradingRunLabel(summary);
      option.selected = summary.gradingRunId === state.selectedSavedRunId;
      selectFragment.append(option);
    }
  }
  elements.savedGradingSelect.replaceChildren(selectFragment);
  elements.savedGradingSelect.disabled = state.savedRuns.length === 0 || state.isGrading || state.isAnnotating;

  const listFragment = document.createDocumentFragment();
  for (const summary of state.savedRuns) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "question-tile";
    if (summary.gradingRunId === state.selectedSavedRunId) {
      button.classList.add("is-selected");
    }

    const id = document.createElement("span");
    id.className = "question-tile__id";
    id.textContent = summary.inputPath || summary.gradingRunId;

    const text = document.createElement("span");
    text.className = "question-tile__text";
    text.textContent = gradingRunLabel(summary);

    button.append(id, text);
    button.disabled = state.isGrading || state.isAnnotating;
    button.addEventListener("click", async () => {
      await loadSavedGradingRun(summary.gradingRunId);
    });
    listFragment.append(button);
  }
  elements.savedGradingList.replaceChildren(listFragment);
}

function downloadBlob(blob, filename) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function downloadGradingCsv() {
  if (!state.results) {
    return;
  }

  const rows = getSortedRows(state.results.rows);
  const reportTotal = state.results.report?.total ?? {};
  const formulaDescription = gradingFormulaDescription(state.results.gradingFormula);
  const objectiveColumns = (state.results.report?.learningObjectives ?? []).map((objective, index) => {
    const id = String(objective.id || `Objective ${index + 1}`).trim();
    return {
      id,
      label: String(objective.label || "").trim(),
      report: objective,
    };
  });
  const objectiveHeaders = objectiveColumns.flatMap((objective) => [
    `${objective.id} Label`,
    `${objective.id} Run Questions`,
    `${objective.id} Run Score`,
    `${objective.id} Score`,
    `${objective.id} Correct`,
    `${objective.id} Wrong`,
    `${objective.id} Blank / Missing`,
  ]);
  const header = [
    "Grading Run ID",
    "Graded At",
    "Input",
    "Formula",
    "Run Processed PDFs",
    "Run Total Score",
    "Run Total Correct",
    "Run Total Wrong",
    "Row",
    "Student ID",
    "Source PDF",
    "Exam Set",
    "Variant",
    "Detected Questions",
    "Variant Questions",
    "Score",
    "Earned Points",
    "Possible Points",
    "Correct",
    "Wrong",
    "Blank",
    "Missing",
    "Blank / Missing",
    "Invalid",
    "Penalty",
    "Status",
    "Issues",
    "Learning Objectives",
    ...objectiveHeaders,
  ];
  const lines = [header.map(csvCell).join(",")];

  for (const row of (rows.length > 0 ? rows : [null])) {
    const summary = row?.summary ?? {};
    const rowObjectives = new Map();
    for (const objective of row?.learningObjectiveSummary ?? []) {
      if (objective?.id) {
        rowObjectives.set(String(objective.id), objective);
      }
    }
    const objectiveValues = objectiveColumns.flatMap((objective) => {
      const item = rowObjectives.get(objective.id);
      return [
        objective.label,
        String(objective.report.questionCount ?? ""),
        scoreText(objective.report.earnedPoints, objective.report.possiblePoints),
        item ? scoreText(item.earnedPoints, item.possiblePoints) : "",
        item ? String(item.correctCount ?? 0) : "",
        item ? String(item.wrongCount ?? 0) : "",
        item ? String(objectiveBlankMissingCount(item)) : "",
      ];
    });
    lines.push([
      state.results.gradingRunId || "",
      state.results.gradedAt || "",
      state.results.inputPath || "",
      formulaDescription,
      String(state.results.summary?.processedCount ?? ""),
      scoreText(reportTotal.earnedPoints, reportTotal.possiblePoints),
      String(reportTotal.correctCount ?? 0),
      String(reportTotal.wrongCount ?? 0),
      String(row?.rowIndex ?? ""),
      row?.displayStudentId || row?.studentId || "",
      row?.sourcePdf || "",
      row?.examSetId || "",
      row?.variantId || "",
      String(row?.detectedQuestionCount ?? ""),
      String(row?.variantQuestionCount ?? ""),
      row ? scoreText(summary.earnedPoints, summary.possiblePoints) : "",
      row ? formatScore(summary.earnedPoints) : "",
      row ? formatScore(summary.possiblePoints) : "",
      row ? String(summary.correctCount ?? 0) : "",
      row ? String(summary.wrongCount ?? summary.incorrectCount ?? 0) : "",
      row ? String(summary.blankCount ?? 0) : "",
      row ? String(summary.missingCount ?? 0) : "",
      row ? countText(summary.blankCount, summary.missingCount) : "",
      row ? String(summary.invalidCount ?? 0) : "",
      row ? formatScore(summary.penaltyPoints) : "",
      row ? statusLabel(row) : "",
      row && Array.isArray(row.issues) ? row.issues.join(" | ") : "",
      row && Array.isArray(row.learningObjectiveSummary) && row.learningObjectiveSummary.length > 0
        ? objectiveSummaryText(row.learningObjectiveSummary)
        : "",
      ...objectiveValues,
    ].map(csvCell).join(","));
  }

  const blob = new Blob([lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
  const timestamp = new Date().toISOString().replaceAll(":", "-");
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `grading-results-${timestamp}.csv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function renderSummary() {
  const result = state.results;
  const hasResult = Boolean(result);
  elements.gradingResults.classList.toggle("hidden", !hasResult);
  const isBusy = state.isGrading || state.isAnnotating || state.isDeletingGradingRun;
  elements.annotateGradedPdfs.disabled = !hasResult || isBusy;
  elements.deleteGradingRun.disabled = !hasResult || isBusy || !result?.gradingRunId;
  elements.exportGradingCsv.disabled = !hasResult || isBusy;
  elements.recalculateGrading.disabled = !hasResult || isBusy;
  elements.runGrading.disabled = isBusy || state.gradingFiles.length === 0;
  elements.browseGradingFile.disabled = isBusy;
  elements.browseGradingFolder.disabled = isBusy;
  elements.gradingPdfUpload.disabled = isBusy;
  elements.gradingFolderUpload.disabled = isBusy;
  elements.gradingInputPath.disabled = isBusy;
  elements.gradingFormulaMode.disabled = isBusy;
  elements.sortCorrect.disabled = isBusy;
  renderFormulaControls();
  renderSavedRuns();
  updateSortButton();
  if (!hasResult) {
    elements.gradingTableBody.replaceChildren();
    elements.gradingDetailList.replaceChildren();
    elements.gradingObjectiveReportBody.replaceChildren();
    state.selectedRowIndex = null;
    state.detailModalOpen = false;
    renderGradingDetailModal();
    return;
  }

  const reportTotal = result.report?.total ?? {};
  const formulaDescription = gradingFormulaDescription(result.gradingFormula);
  elements.gradingHeading.textContent = `Grading Results (${result.summary.processedCount} PDF${result.summary.processedCount === 1 ? "" : "s"})`;
  elements.gradingResultPath.textContent = result.inputPath;
  elements.gradingProcessedCount.textContent = String(result.summary.processedCount);
  elements.gradingKnownCount.textContent = String(result.summary.knownStudentCount);
  elements.gradingDuplicateCount.textContent = String(result.summary.duplicateStudentIdCount ?? 0);
  elements.gradingOErrorCount.textContent = String(result.summary.omrErrorCount);
  elements.gradingMismatchCount.textContent = String(result.summary.mismatchCount);
  elements.gradingTotalScore.textContent = `${formatScore(reportTotal.earnedPoints)}/${formatScore(reportTotal.possiblePoints)}`;
  elements.gradingTotalWrong.textContent = String(reportTotal.wrongCount ?? 0);
  elements.gradingFormulaSummary.textContent = formulaDescription;

  const objectiveFragment = document.createDocumentFragment();
  const objectives = Array.isArray(result.report?.learningObjectives) ? result.report.learningObjectives : [];
  if (objectives.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 6;
    cell.textContent = "No learning-objective data is available for this grading run.";
    row.append(cell);
    objectiveFragment.append(row);
  } else {
    for (const objective of objectives) {
      const row = document.createElement("tr");
      for (const value of [
        `${objective.id} · ${objective.label}`,
        String(objective.questionCount ?? 0),
      ]) {
        const cell = document.createElement("td");
        cell.textContent = value;
        row.append(cell);
      }

      const scoreCell = document.createElement("td");
      scoreCell.append(createScorePill(
        `${formatScore(objective.earnedPoints)}/${formatScore(objective.possiblePoints)}`,
        "score",
      ));
      row.append(scoreCell);

      const correctCell = document.createElement("td");
      correctCell.append(createScorePill(String(objective.correctCount ?? 0), "correct"));
      row.append(correctCell);

      const wrongCell = document.createElement("td");
      wrongCell.append(createScorePill(String(objective.wrongCount ?? 0), "wrong"));
      row.append(wrongCell);

      const blankCell = document.createElement("td");
      blankCell.append(createScorePill(String(objectiveBlankMissingCount(objective)), "blank"));
      row.append(blankCell);

      objectiveFragment.append(row);
    }
  }
  elements.gradingObjectiveReportBody.replaceChildren(objectiveFragment);

  const tableFragment = document.createDocumentFragment();
  const sortedRows = getSortedRows(result.rows);
  if (!sortedRows.some((row) => row.rowIndex === state.selectedRowIndex)) {
    state.selectedRowIndex = null;
    state.detailModalOpen = false;
  }

  for (const row of sortedRows) {
    const tableRow = document.createElement("tr");
    tableRow.className = "pool-table__row grading-table__row";
    if (row.rowIndex === state.selectedRowIndex) {
      tableRow.classList.add("is-selected");
    }
    tableRow.addEventListener("click", (event) => {
      if (event.target.closest("a, button")) {
        return;
      }
      openGradingDetail(row.rowIndex);
    });

    for (const value of [
      String(row.rowIndex ?? "—"),
      row.displayStudentId,
    ]) {
      const cell = document.createElement("td");
      cell.textContent = value;
      tableRow.append(cell);
    }

    const sourceCell = document.createElement("td");
    const sourceUrl = sourcePdfUrl(row);
    if (sourceUrl) {
      const link = document.createElement("a");
      link.href = sourceUrl;
      link.target = "_blank";
      link.rel = "noopener";
      link.textContent = row.sourcePdf || "Open PDF";
      sourceCell.append(link);
    } else {
      sourceCell.textContent = row.sourcePdf || "—";
    }
    tableRow.append(sourceCell);

    for (const value of [
      row.examSetId || "—",
      row.variantId || "—",
      row.variantQuestionCount ? `${row.detectedQuestionCount}/${row.variantQuestionCount}` : String(row.detectedQuestionCount),
    ]) {
      const cell = document.createElement("td");
      cell.textContent = value;
      tableRow.append(cell);
    }

    const scoreCell = document.createElement("td");
    scoreCell.append(createScorePill(
      `${formatScore(row.summary.earnedPoints)}/${formatScore(row.summary.possiblePoints)}`,
      "score",
    ));
    tableRow.append(scoreCell);

    const correctCell = document.createElement("td");
    correctCell.append(createScorePill(String(row.summary.correctCount), "correct"));
    tableRow.append(correctCell);

    const wrongCell = document.createElement("td");
    wrongCell.append(createScorePill(String(row.summary.wrongCount ?? row.summary.incorrectCount ?? 0), "wrong"));
    tableRow.append(wrongCell);

    const blankCell = document.createElement("td");
    blankCell.append(createScorePill(String(row.summary.blankCount + row.summary.missingCount), "blank"));
    tableRow.append(blankCell);

    const objectiveCell = document.createElement("td");
    objectiveCell.append(createCompactObjectiveSummaryNode(row.learningObjectiveSummary));
    tableRow.append(objectiveCell);

    const formulaCell = document.createElement("td");
    formulaCell.textContent = gradingFormulaDescription(row.gradingFormula ?? result.gradingFormula);
    tableRow.append(formulaCell);

    const statusCell = document.createElement("td");
    const badge = document.createElement("span");
    badge.className = `status-badge status-badge--${statusTone(row)}`;
    badge.textContent = statusLabel(row);
    statusCell.append(badge);
    tableRow.append(statusCell);

    const detailCell = document.createElement("td");
    const detailButton = document.createElement("button");
    detailButton.type = "button";
    detailButton.className = "button button--tiny button--ghost grading-detail-button";
    detailButton.textContent = "Details";
    detailButton.setAttribute("aria-label", `Open question-level review for student ${row.displayStudentId}`);
    detailButton.addEventListener("click", (event) => {
      event.stopPropagation();
      openGradingDetail(row.rowIndex);
    });
    detailCell.append(detailButton);
    tableRow.append(detailCell);
    tableFragment.append(tableRow);
  }

  elements.gradingTableBody.replaceChildren(tableFragment);
  renderSelectedDetail(sortedRows);
  renderGradingDetailModal();
}

function renderSelectedDetail(rows) {
  const selectedRow = rows.find((row) => row.rowIndex === state.selectedRowIndex) ?? null;
  if (!selectedRow) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.innerHTML = "<p>Select a student row to inspect question-level grading details.</p>";
    elements.gradingDetailList.replaceChildren(empty);
    return;
  }

  const detail = document.createElement("article");
  detail.className = "grading-detail-card";
  const issueMarkup = selectedRow.issues.length > 0
    ? `<ul class="grading-issue-list">${selectedRow.issues.map((issue) => `<li>${escapeHtml(issue)}</li>`).join("")}</ul>`
    : "<p class=\"helper-copy\">No run-level issues detected.</p>";
  const questionRows = selectedRow.questionDetails.length > 0
    ? selectedRow.questionDetails.map((question) => `
      <tr>
        <td>${question.position}</td>
        <td class="cell-copy">${renderRichTextHtml(question.prompt || "—")}${renderQuestionImageHtml(question)}</td>
        <td>${escapeHtml(question.allowedChoices.join(", ") || "—")}</td>
        <td>${escapeHtml(question.markedAnswers.join(", ") || "—")}</td>
        <td>${escapeHtml(question.correctAnswers.join(", ") || "—")}</td>
        <td>${formatScore(question.earnedPoints ?? 0)}/${formatScore(question.points ?? 1)}${question.penaltyPoints ? ` (-${formatScore(question.penaltyPoints)})` : ""}</td>
        <td>${escapeHtml(question.status)}</td>
        <td class="cell-copy">${escapeHtml(question.issues.join(" ") || "—")}</td>
      </tr>
    `).join("")
    : "<tr><td colspan=\"8\">No question-level data available.</td></tr>";

  detail.innerHTML = `
    <div class="grading-detail-card__head">
      <div>
        <p class="eyebrow">Student ${escapeHtml(selectedRow.displayStudentId)}</p>
        <h3>${escapeHtml(selectedRow.sourcePdf || "Unknown Source")}</h3>
        <p class="grading-detail-card__meta">${escapeHtml(selectedRow.examName || "Unknown Exam")} · ${escapeHtml(selectedRow.variantId || "No Variant")}</p>
      </div>
      <span class="status-badge status-badge--${statusTone(selectedRow)}">${escapeHtml(statusLabel(selectedRow))}</span>
    </div>
    <div class="grading-detail-metrics">
      <div class="metric"><span class="metric__label">Score</span><span class="metric__value">${formatScore(selectedRow.summary.earnedPoints)} / ${formatScore(selectedRow.summary.possiblePoints)}</span></div>
      <div class="metric"><span class="metric__label">Correct</span><span class="metric__value">${selectedRow.summary.correctCount}</span></div>
      <div class="metric"><span class="metric__label">Wrong</span><span class="metric__value">${selectedRow.summary.wrongCount ?? selectedRow.summary.incorrectCount}</span></div>
      <div class="metric"><span class="metric__label">Penalty</span><span class="metric__value">${formatScore(selectedRow.summary.penaltyPoints)}</span></div>
      <div class="metric"><span class="metric__label">Blank</span><span class="metric__value">${selectedRow.summary.blankCount}</span></div>
      <div class="metric"><span class="metric__label">Missing</span><span class="metric__value">${selectedRow.summary.missingCount}</span></div>
      <div class="metric"><span class="metric__label">Invalid</span><span class="metric__value">${selectedRow.summary.invalidCount}</span></div>
    </div>
    <div class="grading-detail-card__issues">
      <h4>Formula</h4>
      <p class="helper-copy">${escapeHtml(gradingFormulaDescription(selectedRow.gradingFormula ?? state.results?.gradingFormula))}</p>
    </div>
    <div class="grading-detail-card__issues">
      <h4>Learning Objectives</h4>
      <div class="grading-detail-objectives"></div>
    </div>
    <div class="grading-detail-card__issues">
      <h4>Run-Level Checks</h4>
      ${issueMarkup}
    </div>
    <div class="table-wrap">
      <table class="pool-table grading-question-table">
        <thead>
          <tr>
            <th>Q#</th>
            <th>Prompt</th>
            <th>Allowed</th>
            <th>Marked</th>
            <th>Correct</th>
            <th>Points</th>
            <th>Status</th>
            <th>Issue</th>
          </tr>
        </thead>
        <tbody>${questionRows}</tbody>
      </table>
    </div>
  `;
  const objectiveDetail = detail.querySelector(".grading-detail-objectives");
  if (objectiveDetail) {
    objectiveDetail.append(createObjectiveSummaryNode(selectedRow.learningObjectiveSummary));
  }
  elements.gradingDetailList.replaceChildren(detail);
}

async function loadPaths() {
  setStatus("Loading grading tools...");
  const response = await fetch("/api/quiz");
  if (!response.ok) {
    throw new Error(`Could not load quiz metadata (${response.status})`);
  }
  const payload = await response.json();
  state.dbPath = payload.dbPath;
  state.examStorePath = payload.examStorePath;
  elements.dbPath.textContent = state.dbPath;
  elements.examStorePath.textContent = state.examStorePath;
  setStatus("Ready to grade filled OMR PDFs.");
}

async function loadSavedGradingRuns({ loadLatest = false } = {}) {
  const response = await fetch("/api/gradings");
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not load saved grading runs (${response.status})`);
  }
  state.savedRuns = Array.isArray(payload.gradingRuns) ? payload.gradingRuns : [];
  renderSavedRuns();
  if (loadLatest && state.savedRuns.length > 0 && !state.results) {
    await loadSavedGradingRun(state.savedRuns[0].gradingRunId);
  }
}

async function loadSavedGradingRun(gradingRunId) {
  if (!gradingRunId) {
    return;
  }
  setStatus("Loading saved grading run...");
  const response = await fetch(`/api/gradings/run/${encodeURIComponent(gradingRunId)}`);
  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<grading>", message: "Could not load saved grading run" }];
    renderErrors();
    setStatus("Saved grading run load failed.", true);
    return;
  }
  state.results = payload.gradingRun;
  state.annotationResults = null;
  state.selectedRowIndex = null;
  state.detailModalOpen = false;
  state.selectedSavedRunId = gradingRunId;
  state.gradingFiles = [];
  elements.gradingInputPath.value = "";
  populateGradingFormula(state.results.gradingFormula);
  state.validationErrors = [];
  renderErrors();
  renderSummary();
  renderSavedRuns();
  setStatus(`Loaded saved grading run ${gradingRunId}.`);
}

async function runGrading() {
  if (state.gradingFiles.length === 0) {
    state.validationErrors = [{ path: "pdfs", message: "Choose at least one PDF before running grading." }];
    renderErrors();
    setStatus("Choose at least one PDF before running grading.", true);
    return;
  }

  state.validationErrors = [];
  renderErrors();
  state.isGrading = true;
  setBusyState(
    true,
    "Running Grading",
    "Please wait while omr-grade reads the PDF files and validates them against the stored exam variants.",
  );
  renderSummary();
  setStatus("Running omr-grade...");

  try {
    const formData = new FormData();
    for (const file of state.gradingFiles) {
      formData.append("pdfs", file, file.webkitRelativePath || file.name);
    }

    const response = await fetch("/api/exams/grade-upload", {
      method: "POST",
      headers: {
        "X-Quiz-Pool-Grading-Formula": JSON.stringify(currentGradingFormula()),
      },
      body: formData,
    });

    const payload = await response.json();
    if (!response.ok) {
      state.results = null;
      state.validationErrors = payload.errors ?? [{ path: "<unknown>", message: "Grading failed" }];
      renderErrors();
      renderSummary();
      setStatus("Grading failed. Review the messages.", true);
      return;
    }

    state.results = payload;
    state.selectedSavedRunId = payload.gradingRunId ?? "";
    state.annotationResults = null;
    state.selectedRowIndex = null;
    state.detailModalOpen = false;
    populateGradingFormula(payload.gradingFormula);
    state.validationErrors = [];
    renderErrors();
    renderSummary();
    await loadSavedGradingRuns();
    setStatus(`Processed ${payload.summary.processedCount} PDF(s).`);
  } finally {
    setBusyState(false);
    renderSummary();
  }
}

async function annotateGradedPdfs() {
  if (!state.results) {
    return;
  }

  closeAnnotateModal();
  state.validationErrors = [];
  renderErrors();
  state.isAnnotating = true;
  setBusyState(
    true,
    "Annotating PDFs",
    "Please wait while omr-annotate packages annotated review PDFs.",
  );
  renderSummary();
  setStatus("Running omr-annotate...");

  try {
    const response = await fetch("/api/exams/annotate-upload", {
      method: "POST",
    });

    if (!response.ok) {
      const payload = await response.json();
      state.validationErrors = payload.errors ?? [{ path: "<unknown>", message: "Annotation failed" }];
      renderErrors();
      setStatus("Annotation failed. Review the messages.", true);
      return;
    }

    const blob = await response.blob();
    const annotatedCount = response.headers.get("X-Quiz-Pool-Annotated-Count");
    downloadBlob(blob, "annotated-pdfs.zip");
    state.annotationResults = { annotatedCount };
    state.validationErrors = [];
    renderErrors();
    setStatus(`Downloaded annotated PDF package${annotatedCount ? ` (${annotatedCount} PDF${annotatedCount === "1" ? "" : "s"})` : ""}.`);
  } finally {
    setBusyState(false);
    renderSummary();
  }
}

async function deleteCurrentGradingRun() {
  if (!state.results?.gradingRunId || state.isDeletingGradingRun) {
    return;
  }

  const gradingRunId = state.results.gradingRunId;
  const label = state.results.inputPath || gradingRunId;
  if (!window.confirm(`Delete grading run "${label}"? This cannot be undone.`)) {
    return;
  }

  state.validationErrors = [];
  state.isDeletingGradingRun = true;
  renderErrors();
  renderSummary();
  setStatus("Deleting grading run...");

  try {
    const response = await fetch(`/api/gradings/run/${encodeURIComponent(gradingRunId)}`, {
      method: "DELETE",
    });
    let payload = {};
    try {
      payload = await response.json();
    } catch {
      payload = {};
    }

    if (!response.ok) {
      state.validationErrors = payload.errors ?? [{ path: "<grading>", message: "Could not delete grading run" }];
      renderErrors();
      setStatus("Delete failed. Review the messages.", true);
      return;
    }

    state.results = null;
    state.annotationResults = null;
    state.selectedSavedRunId = "";
    state.selectedRowIndex = null;
    state.detailModalOpen = false;
    state.validationErrors = [];
    renderErrors();
    await loadSavedGradingRuns();
    setStatus(`Deleted grading run ${gradingRunId}.`);
  } finally {
    state.isDeletingGradingRun = false;
    renderSummary();
  }
}

async function recalculateCurrentRun() {
  if (!state.results?.gradingRunId) {
    return;
  }
  state.validationErrors = [];
  renderErrors();
  setStatus("Recalculating grading scores...");
  const gradingRunId = state.results.gradingRunId;
  const response = await fetch(`/api/gradings/run/${encodeURIComponent(gradingRunId)}/formula`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      gradingFormula: currentGradingFormula(),
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<grading>", message: "Could not recalculate grading scores" }];
    renderErrors();
    setStatus("Recalculation failed. Review the messages.", true);
    return;
  }
  state.results = payload.gradingRun;
  state.selectedSavedRunId = gradingRunId;
  populateGradingFormula(state.results.gradingFormula);
  renderSummary();
  await loadSavedGradingRuns();
  setStatus("Recalculated scores for the current grading run.");
}

function wireEvents() {
  if (!("webkitdirectory" in elements.gradingFolderUpload)) {
    elements.browseGradingFolder.hidden = true;
  }

  elements.browseGradingFile.addEventListener("click", () => {
    elements.gradingPdfUpload.click();
  });

  elements.browseGradingFolder.addEventListener("click", () => {
    elements.gradingFolderUpload.click();
  });

  elements.gradingPdfUpload.addEventListener("change", (event) => {
    const files = selectedPdfFiles(event.target.files);
    event.target.value = "";
    setGradingFiles(files);
  });

  elements.gradingFolderUpload.addEventListener("change", (event) => {
    const files = selectedPdfFiles(event.target.files);
    event.target.value = "";
    setGradingFiles(files);
  });

  elements.annotateGradedPdfs.addEventListener("click", () => {
    openAnnotateModal();
  });
  elements.cancelAnnotateModal.addEventListener("click", () => {
    closeAnnotateModal();
  });
  elements.confirmAnnotateModal.addEventListener("click", async () => {
    await annotateGradedPdfs();
  });
  elements.annotateModalBackdrop.addEventListener("click", () => {
    closeAnnotateModal();
  });
  elements.closeGradingDetail.addEventListener("click", () => {
    closeGradingDetail();
  });
  elements.gradingDetailBackdrop.addEventListener("click", () => {
    closeGradingDetail();
  });
  elements.runGrading.addEventListener("click", async () => {
    await runGrading();
  });
  elements.recalculateGrading.addEventListener("click", async () => {
    await recalculateCurrentRun();
  });
  elements.deleteGradingRun.addEventListener("click", async () => {
    await deleteCurrentGradingRun();
  });
  elements.gradingFormulaMode.addEventListener("change", () => {
    state.gradingFormula = currentGradingFormula();
    renderFormulaControls();
    if (state.results) {
      setStatus("Grading formula changed. Recalculate the current run to update saved scores.");
    }
  });
  elements.gradingFixedPenalty.addEventListener("input", () => {
    state.gradingFormula = currentGradingFormula();
    if (state.results) {
      setStatus("Grading formula changed. Recalculate the current run to update saved scores.");
    }
  });
  elements.savedGradingSelect.addEventListener("change", async (event) => {
    await loadSavedGradingRun(event.target.value);
  });
  elements.exportGradingCsv.addEventListener("click", () => {
    downloadGradingCsv();
  });
  elements.sortCorrect.addEventListener("click", () => {
    state.correctSortDirection = state.correctSortDirection === "desc" ? "asc" : "desc";
    renderSummary();
  });
  window.addEventListener("keydown", async (event) => {
    if (event.key === "Escape") {
      if (state.annotateModalOpen) {
        closeAnnotateModal();
        return;
      }
      if (state.detailModalOpen) {
        closeGradingDetail();
      }
      return;
    }
    if (state.annotateModalOpen && event.key === "Enter") {
      event.preventDefault();
      await annotateGradedPdfs();
    }
  });
}

wireEvents();
renderAnnotateModal();
renderGradingDetailModal();
populateGradingFormula();
renderSummary();
loadPaths()
  .then(() => loadSavedGradingRuns({ loadLatest: true }))
  .catch((error) => {
    console.error(error);
    setStatus(error.message, true);
  });
