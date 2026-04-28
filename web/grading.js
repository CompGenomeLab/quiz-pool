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
  isGrading: false,
  results: null,
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
  dbPath: document.querySelector("#db-path"),
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
  runGrading: document.querySelector("#run-grading"),
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

function csvCell(value) {
  const normalized = String(value ?? "");
  if (normalized.includes('"') || normalized.includes(",") || normalized.includes("\n")) {
    return `"${normalized.replaceAll('"', '""')}"`;
  }
  return normalized;
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
  elements.gradingInputPath.value = formatFileSelection(files);
  renderErrors();
  renderSummary();
  setStatus(files.length ? `${files.length} PDF${files.length === 1 ? "" : "s"} selected.` : "No PDFs selected.");
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
  const lines = [
    [
      "Row",
      "Student ID",
      "Source PDF",
      "Exam Set",
      "Variant",
      "Questions",
      "Score",
      "Correct",
      "Blank",
      "Status",
    ].map(csvCell).join(","),
  ];

  for (const row of rows) {
    const questionCount = row.variantQuestionCount
      ? `${row.detectedQuestionCount}/${row.variantQuestionCount}`
      : String(row.detectedQuestionCount);
    lines.push([
      String(row.rowIndex ?? ""),
      row.displayStudentId,
      row.sourcePdf || "—",
      row.examSetId || "—",
      row.variantId || "—",
      questionCount,
      `${row.summary.earnedPoints ?? 0}/${row.summary.possiblePoints ?? 0}`,
      String(row.summary.correctCount),
      String(row.summary.blankCount + row.summary.missingCount),
      statusLabel(row),
    ].map(csvCell).join(","));
  }

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
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
  const isBusy = state.isGrading || state.isAnnotating;
  elements.annotateGradedPdfs.disabled = !hasResult || isBusy;
  elements.exportGradingCsv.disabled = !hasResult || isBusy;
  elements.runGrading.disabled = isBusy || state.gradingFiles.length === 0;
  elements.browseGradingFile.disabled = isBusy;
  elements.browseGradingFolder.disabled = isBusy;
  elements.gradingPdfUpload.disabled = isBusy;
  elements.gradingFolderUpload.disabled = isBusy;
  elements.gradingInputPath.disabled = isBusy;
  elements.sortCorrect.disabled = isBusy;
  updateSortButton();
  if (!hasResult) {
    elements.gradingTableBody.replaceChildren();
    elements.gradingDetailList.replaceChildren();
    state.selectedRowIndex = null;
    return;
  }

  elements.gradingHeading.textContent = `Grading Results (${result.summary.processedCount} PDF${result.summary.processedCount === 1 ? "" : "s"})`;
  elements.gradingResultPath.textContent = result.inputPath;
  elements.gradingProcessedCount.textContent = String(result.summary.processedCount);
  elements.gradingKnownCount.textContent = String(result.summary.knownStudentCount);
  elements.gradingDuplicateCount.textContent = String(result.summary.duplicateStudentIdCount ?? 0);
  elements.gradingOErrorCount.textContent = String(result.summary.omrErrorCount);
  elements.gradingMismatchCount.textContent = String(result.summary.mismatchCount);

  const tableFragment = document.createDocumentFragment();
  const sortedRows = getSortedRows(result.rows);
  if (!sortedRows.some((row) => row.rowIndex === state.selectedRowIndex)) {
    state.selectedRowIndex = sortedRows[0]?.rowIndex ?? null;
  }

  for (const row of sortedRows) {
    const tableRow = document.createElement("tr");
    tableRow.className = "pool-table__row grading-table__row";
    if (row.rowIndex === state.selectedRowIndex) {
      tableRow.classList.add("is-selected");
    }
    tableRow.addEventListener("click", () => {
      state.selectedRowIndex = row.rowIndex;
      renderSummary();
    });

    for (const value of [
      String(row.rowIndex ?? "—"),
      row.displayStudentId,
      row.sourcePdf || "—",
      row.examSetId || "—",
      row.variantId || "—",
      row.variantQuestionCount ? `${row.detectedQuestionCount}/${row.variantQuestionCount}` : String(row.detectedQuestionCount),
      `${row.summary.earnedPoints ?? 0}/${row.summary.possiblePoints ?? 0}`,
      String(row.summary.correctCount),
      String(row.summary.blankCount + row.summary.missingCount),
    ]) {
      const cell = document.createElement("td");
      cell.textContent = value;
      tableRow.append(cell);
    }
    const statusCell = document.createElement("td");
    const badge = document.createElement("span");
    badge.className = `status-badge status-badge--${statusTone(row)}`;
    badge.textContent = statusLabel(row);
    statusCell.append(badge);
    tableRow.append(statusCell);
    tableFragment.append(tableRow);
  }

  elements.gradingTableBody.replaceChildren(tableFragment);
  renderSelectedDetail(sortedRows);
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
        <td>${question.earnedPoints ?? 0}/${question.points ?? 1}</td>
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
      <div class="metric"><span class="metric__label">Score</span><span class="metric__value">${selectedRow.summary.earnedPoints ?? 0} / ${selectedRow.summary.possiblePoints ?? 0}</span></div>
      <div class="metric"><span class="metric__label">Correct</span><span class="metric__value">${selectedRow.summary.correctCount}</span></div>
      <div class="metric"><span class="metric__label">Incorrect</span><span class="metric__value">${selectedRow.summary.incorrectCount}</span></div>
      <div class="metric"><span class="metric__label">Blank</span><span class="metric__value">${selectedRow.summary.blankCount}</span></div>
      <div class="metric"><span class="metric__label">Missing</span><span class="metric__value">${selectedRow.summary.missingCount}</span></div>
      <div class="metric"><span class="metric__label">Invalid</span><span class="metric__value">${selectedRow.summary.invalidCount}</span></div>
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
    state.annotationResults = null;
    state.selectedRowIndex = payload.rows[0]?.rowIndex ?? null;
    state.validationErrors = [];
    renderErrors();
    renderSummary();
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
  elements.runGrading.addEventListener("click", async () => {
    await runGrading();
  });
  elements.exportGradingCsv.addEventListener("click", () => {
    downloadGradingCsv();
  });
  elements.sortCorrect.addEventListener("click", () => {
    state.correctSortDirection = state.correctSortDirection === "desc" ? "asc" : "desc";
    renderSummary();
  });
  window.addEventListener("keydown", async (event) => {
    if (!state.annotateModalOpen) {
      return;
    }
    if (event.key === "Escape") {
      closeAnnotateModal();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      await annotateGradedPdfs();
    }
  });
}

wireEvents();
renderAnnotateModal();
renderSummary();
loadPaths().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
