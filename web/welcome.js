const state = {
  currentDbPath: "",
  currentExamStorePath: "",
  defaultDbPath: "",
  defaultExamStorePath: "",
  schemaPath: "",
  statusIsError: false,
  statusMessage: "Loading current paths...",
  validationErrors: [],
};

const elements = {
  applyPaths: document.querySelector("#welcome-apply-paths"),
  currentDbPath: document.querySelector("#welcome-current-db-path"),
  currentExamStorePath: document.querySelector("#welcome-current-exam-store-path"),
  dbPath: document.querySelector("#welcome-db-path"),
  defaultDbPath: document.querySelector("#welcome-default-db-path"),
  defaultExamStorePath: document.querySelector("#welcome-default-exam-store-path"),
  errorList: document.querySelector("#welcome-error-list"),
  errorPanel: document.querySelector("#welcome-errors"),
  examStorePath: document.querySelector("#welcome-exam-store-path"),
  schemaPath: document.querySelector("#welcome-schema-path"),
  status: document.querySelector("#welcome-status"),
  useDefaults: document.querySelector("#welcome-use-defaults"),
};

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  elements.status.textContent = message;
  elements.status.style.color = isError ? "var(--danger-strong)" : "var(--muted)";
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

function render() {
  elements.dbPath.value = state.currentDbPath;
  elements.examStorePath.value = state.currentExamStorePath;
  elements.currentDbPath.textContent = state.currentDbPath;
  elements.currentExamStorePath.textContent = state.currentExamStorePath;
  elements.defaultDbPath.textContent = state.defaultDbPath;
  elements.defaultExamStorePath.textContent = state.defaultExamStorePath;
  elements.schemaPath.textContent = state.schemaPath;
  renderErrors();
}

async function loadSessionPaths() {
  setStatus("Loading current paths...");
  const response = await fetch("/api/session-paths");
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not load paths (${response.status})`);
  }

  state.currentDbPath = payload.dbPath;
  state.currentExamStorePath = payload.examStorePath;
  state.defaultDbPath = payload.defaultDbPath;
  state.defaultExamStorePath = payload.defaultExamStorePath;
  state.schemaPath = payload.schemaPath;
  state.validationErrors = [];
  render();
  setStatus("Ready.");
}

async function applyPaths() {
  const dbPath = elements.dbPath.value.trim();
  const examStorePath = elements.examStorePath.value.trim();
  state.validationErrors = [];
  renderErrors();
  setStatus("Applying selected paths...");

  const response = await fetch("/api/session-paths", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ dbPath, examStorePath }),
  });
  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<paths>", message: "Could not apply paths" }];
    renderErrors();
    setStatus("Path update failed. Review the messages.", true);
    return;
  }

  state.currentDbPath = payload.dbPath;
  state.currentExamStorePath = payload.examStorePath;
  state.defaultDbPath = payload.defaultDbPath;
  state.defaultExamStorePath = payload.defaultExamStorePath;
  state.schemaPath = payload.schemaPath;
  state.validationErrors = [];
  render();
  setStatus("Active paths updated.");
}

function wireEvents() {
  elements.useDefaults.addEventListener("click", () => {
    elements.dbPath.value = state.defaultDbPath;
    elements.examStorePath.value = state.defaultExamStorePath;
    setStatus("Default paths loaded into the form.");
  });

  elements.applyPaths.addEventListener("click", async () => {
    await applyPaths();
  });
}

wireEvents();
loadSessionPaths().catch((error) => {
  console.error(error);
  state.validationErrors = [{ path: "<welcome>", message: error.message }];
  renderErrors();
  setStatus(error.message, true);
});
