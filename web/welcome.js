import { openSystemFileDialog } from "./system-file-dialog.js";

const state = {
  currentProjectPath: "",
  hasActiveProject: false,
  selectedProjectPath: "",
  statusIsError: false,
  statusMessage: "Loading current project...",
  validationErrors: [],
};

const elements = {
  applyProject: document.querySelector("#welcome-apply-project"),
  browseProject: document.querySelector("#welcome-browse-project"),
  createProject: document.querySelector("#welcome-create-project"),
  errorList: document.querySelector("#welcome-error-list"),
  errorPanel: document.querySelector("#welcome-errors"),
  projectLinks: [...document.querySelectorAll("[data-project-link]")],
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
  elements.selectedProjectPath.value = state.selectedProjectPath || state.currentProjectPath;
  for (const link of elements.projectLinks) {
    if (!link.dataset.projectHref) {
      link.dataset.projectHref = link.getAttribute("href") ?? "";
    }
    link.classList.toggle("is-disabled", !state.hasActiveProject);
    link.setAttribute("aria-disabled", String(!state.hasActiveProject));
    if (state.hasActiveProject) {
      link.setAttribute("href", link.dataset.projectHref);
      link.removeAttribute("tabindex");
    } else {
      link.removeAttribute("href");
      link.setAttribute("tabindex", "-1");
    }
  }
  renderErrors();
}

async function loadProject() {
  setStatus("Loading current project...");
  const response = await fetch("/api/project");
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not load project (${response.status})`);
  }

  state.hasActiveProject = Boolean(payload.hasActiveProject);
  state.currentProjectPath = payload.projectPath ?? "";
  state.selectedProjectPath = payload.projectPath ?? "";
  state.validationErrors = [];
  render();
  setStatus(
    state.hasActiveProject
      ? "Ready."
      : "No active project. Create a project DB or open an existing one.",
  );
}

async function activateProject(projectPath, {
  pendingStatus = "Opening selected project...",
  successStatus = "Active project updated.",
} = {}) {
  state.validationErrors = [];
  renderErrors();
  setStatus(pendingStatus);

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

  state.hasActiveProject = Boolean(payload.hasActiveProject ?? true);
  state.currentProjectPath = payload.projectPath;
  state.selectedProjectPath = payload.projectPath;
  render();
  setStatus(successStatus);
}

async function openProject() {
  const projectPath = elements.selectedProjectPath.value.trim() || state.selectedProjectPath || state.currentProjectPath;
  await activateProject(projectPath);
}

function wireEvents() {
  elements.createProject.addEventListener("click", async () => {
    try {
      setStatus("Choose where to create the project DB...");
      const selectedPath = await openSystemFileDialog({
        title: "Create Project DB",
        purpose: "project",
        mode: "save-file",
        startPath: state.selectedProjectPath || state.currentProjectPath,
      });
      if (!selectedPath) {
        setStatus("Project creation canceled.");
        return;
      }
      state.selectedProjectPath = selectedPath;
      elements.selectedProjectPath.value = selectedPath;
      await activateProject(selectedPath, {
        pendingStatus: "Creating project DB...",
        successStatus: "Project DB created and activated.",
      });
    } catch (error) {
      state.validationErrors = [{ path: "<project>", message: error.message }];
      renderErrors();
      setStatus(error.message, true);
    }
  });

  elements.browseProject.addEventListener("click", async () => {
    try {
      setStatus("Opening system file picker...");
      const selectedPath = await openSystemFileDialog({
        title: "Open Project DB",
        purpose: "project",
        mode: "file",
        startPath: state.selectedProjectPath || state.currentProjectPath,
      });
      if (!selectedPath) {
        setStatus("Project selection canceled.");
        return;
      }
      state.selectedProjectPath = selectedPath;
      elements.selectedProjectPath.value = selectedPath;
      state.validationErrors = [];
      render();
      setStatus("Project selected. Open it to switch the session.");
    } catch (error) {
      state.validationErrors = [{ path: "<project>", message: error.message }];
      renderErrors();
      setStatus(error.message, true);
    }
  });

  elements.applyProject.addEventListener("click", async () => {
    await openProject();
  });

  elements.selectedProjectPath.addEventListener("input", (event) => {
    state.selectedProjectPath = event.target.value;
  });
}

wireEvents();
loadProject().catch((error) => {
  console.error(error);
  state.validationErrors = [{ path: "<welcome>", message: error.message }];
  renderErrors();
  setStatus(error.message, true);
});
