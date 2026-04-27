import { browseFile } from "./file-browser.js";

const state = {
  currentProjectPath: "",
  defaultProjectPath: "",
  selectedProjectPath: "",
  statusIsError: false,
  statusMessage: "Loading current project...",
  validationErrors: [],
};

const elements = {
  applyProject: document.querySelector("#welcome-apply-project"),
  browseProject: document.querySelector("#welcome-browse-project"),
  currentProjectPath: document.querySelector("#welcome-current-project-path"),
  defaultProjectPath: document.querySelector("#welcome-default-project-path"),
  errorList: document.querySelector("#welcome-error-list"),
  errorPanel: document.querySelector("#welcome-errors"),
  selectedProjectPath: document.querySelector("#welcome-selected-project-path"),
  status: document.querySelector("#welcome-status"),
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
  elements.currentProjectPath.textContent = state.currentProjectPath;
  elements.defaultProjectPath.textContent = state.defaultProjectPath;
  elements.selectedProjectPath.textContent = state.selectedProjectPath || state.currentProjectPath;
  renderErrors();
}

async function loadProject() {
  setStatus("Loading current project...");
  const response = await fetch("/api/project");
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not load project (${response.status})`);
  }

  state.currentProjectPath = payload.projectPath;
  state.selectedProjectPath = payload.projectPath;
  state.defaultProjectPath = payload.defaultProjectPath ?? "";
  state.validationErrors = [];
  render();
  setStatus("Ready.");
}

async function openProject() {
  const projectPath = state.selectedProjectPath || state.currentProjectPath;
  state.validationErrors = [];
  renderErrors();
  setStatus("Opening selected project...");

  const response = await fetch("/api/project/open", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ projectPath }),
  });
  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<project>", message: "Could not open project" }];
    renderErrors();
    setStatus("Project open failed. Review the messages.", true);
    return;
  }

  state.currentProjectPath = payload.projectPath;
  state.selectedProjectPath = payload.projectPath;
  state.defaultProjectPath = payload.defaultProjectPath ?? state.defaultProjectPath;
  render();
  setStatus("Active project updated.");
}

function wireEvents() {
  elements.browseProject.addEventListener("click", async () => {
    const selectedPath = await browseFile({
      title: "Open Project DB",
      purpose: "project",
      startPath: state.selectedProjectPath || state.currentProjectPath,
    });
    if (!selectedPath) {
      return;
    }
    state.selectedProjectPath = selectedPath;
    render();
    setStatus("Project selected. Open it to switch the session.");
  });

  elements.applyProject.addEventListener("click", async () => {
    await openProject();
  });
}

wireEvents();
loadProject().catch((error) => {
  console.error(error);
  state.validationErrors = [{ path: "<welcome>", message: error.message }];
  renderErrors();
  setStatus(error.message, true);
});
