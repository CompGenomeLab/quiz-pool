const state = {
  annotationResults: null,
  busyMessage: "",
  busyTitle: "",
  dbPath: "",
  examStorePath: "",
  isAnnotating: false,
  isGrading: false,
  results: null,
  correctSortDirection: "desc",
  statusIsError: false,
  statusMessage: "Loading grading tools...",
  validationErrors: [],
};

const elements = {
  annotateGradedPdfs: document.querySelector("#annotate-graded-pdfs"),
  dbPath: document.querySelector("#db-path"),
  errorList: document.querySelector("#grading-error-list"),
  exportGradingCsv: document.querySelector("#export-grading-csv"),
  errorPanel: document.querySelector("#grading-errors"),
  examStorePath: document.querySelector("#exam-store-path"),
  gradingHeading: document.querySelector("#grading-heading"),
  gradingDuplicateCount: document.querySelector("#grading-duplicate-count"),
  gradingInputPath: document.querySelector("#grading-input-path"),
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

function defaultAnnotationOutputPath(inputPath) {
  const normalized = String(inputPath || "").trim();
  if (normalized === "") {
    return "";
  }
  if (normalized.toLowerCase().endsWith(".pdf")) {
    const lastSlash = Math.max(normalized.lastIndexOf("/"), normalized.lastIndexOf("\\"));
    const baseDir = lastSlash >= 0 ? normalized.slice(0, lastSlash) : ".";
    return `${baseDir}/annotated`;
  }
  return `${normalized.replace(/[\\/]+$/u, "")}/annotated`;
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
  elements.runGrading.disabled = isBusy;
  elements.gradingInputPath.disabled = isBusy;
  elements.sortCorrect.disabled = isBusy;
  updateSortButton();
  if (!hasResult) {
    elements.gradingTableBody.replaceChildren();
    elements.gradingDetailList.replaceChildren();
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
  const detailFragment = document.createDocumentFragment();
  const sortedRows = getSortedRows(result.rows);

  for (const row of sortedRows) {
    const tableRow = document.createElement("tr");
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

    const detail = document.createElement("article");
    detail.className = "grading-detail-card";
    const issueMarkup = row.issues.length > 0
      ? `<ul class="grading-issue-list">${row.issues.map((issue) => `<li>${escapeHtml(issue)}</li>`).join("")}</ul>`
      : "<p class=\"helper-copy\">No run-level issues detected.</p>";
    const questionRows = row.questionDetails.length > 0
      ? row.questionDetails.map((question) => `
        <tr>
          <td>${question.position}</td>
          <td class="cell-copy">${escapeHtml(question.prompt || "—")}</td>
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
          <p class="eyebrow">Student ${escapeHtml(row.displayStudentId)}</p>
          <h3>${escapeHtml(row.sourcePdf || "Unknown Source")}</h3>
          <p class="grading-detail-card__meta">${escapeHtml(row.examName || "Unknown Exam")} · ${escapeHtml(row.variantId || "No Variant")}</p>
        </div>
        <span class="status-badge status-badge--${statusTone(row)}">${escapeHtml(statusLabel(row))}</span>
      </div>
      <div class="grading-detail-metrics">
        <div class="metric"><span class="metric__label">Score</span><span class="metric__value">${row.summary.earnedPoints ?? 0} / ${row.summary.possiblePoints ?? 0}</span></div>
        <div class="metric"><span class="metric__label">Correct</span><span class="metric__value">${row.summary.correctCount}</span></div>
        <div class="metric"><span class="metric__label">Incorrect</span><span class="metric__value">${row.summary.incorrectCount}</span></div>
        <div class="metric"><span class="metric__label">Blank</span><span class="metric__value">${row.summary.blankCount}</span></div>
        <div class="metric"><span class="metric__label">Missing</span><span class="metric__value">${row.summary.missingCount}</span></div>
        <div class="metric"><span class="metric__label">Invalid</span><span class="metric__value">${row.summary.invalidCount}</span></div>
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
    detailFragment.append(detail);
  }

  elements.gradingTableBody.replaceChildren(tableFragment);
  elements.gradingDetailList.replaceChildren(detailFragment);
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
  const inputPath = elements.gradingInputPath.value.trim();
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
    const response = await fetch("/api/exams/grade", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ inputPath }),
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

  const suggestedPath = defaultAnnotationOutputPath(state.results.inputPath);
  const outputPath = window.prompt("Output folder path for omr-annotate:", suggestedPath);
  if (outputPath === null) {
    return;
  }
  const trimmedOutputPath = outputPath.trim();
  if (trimmedOutputPath === "") {
    setStatus("Annotation cancelled. Output folder path is required.", true);
    return;
  }

  state.validationErrors = [];
  renderErrors();
  state.isAnnotating = true;
  setBusyState(
    true,
    "Annotating PDFs",
    "Please wait while omr-annotate writes annotated review PDFs to the selected output folder.",
  );
  renderSummary();
  setStatus("Running omr-annotate...");

  try {
    const response = await fetch("/api/exams/annotate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        inputPath: state.results.inputPath,
        outputPath: trimmedOutputPath,
      }),
    });

    const payload = await response.json();
    if (!response.ok) {
      state.validationErrors = payload.errors ?? [{ path: "<unknown>", message: "Annotation failed" }];
      renderErrors();
      setStatus("Annotation failed. Review the messages.", true);
      return;
    }

    state.annotationResults = payload;
    state.validationErrors = [];
    renderErrors();
    setStatus(`Annotated ${payload.summary.annotatedCount} PDF(s) into ${payload.outputPath}.`);
  } finally {
    setBusyState(false);
    renderSummary();
  }
}

function wireEvents() {
  elements.annotateGradedPdfs.addEventListener("click", async () => {
    await annotateGradedPdfs();
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
}

wireEvents();
loadPaths().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
