import { browseFile } from "./file-browser.js";
import { hasRichTextMarkup, renderRichTextIntoElement, stripRichTextMarkup } from "./rich-text.js";

const state = {
  dbPath: "",
  isDirty: false,
  lastSavedSnapshot: "",
  metaPanelCollapsed: false,
  pendingNavigationHref: "",
  quiz: null,
  selectedQuestionIndex: 0,
  showReloadNotice: false,
  statusIsError: false,
  statusMessage: "Loading quiz data...",
  validationErrors: [],
};

const CHOICE_KEYS = ["A", "B", "C", "D", "E"];
const MIN_CHOICE_COUNT = 2;
const MAX_CHOICE_COUNT = CHOICE_KEYS.length;

const elements = {
  addChoice: document.querySelector("#add-choice"),
  addObjective: document.querySelector("#add-objective"),
  addQuestion: document.querySelector("#add-question"),
  addLocation: document.querySelector("#add-location"),
  locations: document.querySelector("#locations"),
  cancelReload: document.querySelector("#cancel-reload"),
  choicesEditor: document.querySelector("#choices-editor"),
  confirmUnsavedNav: document.querySelector("#confirm-unsaved-nav"),
  confirmReload: document.querySelector("#confirm-reload"),
  dbPath: document.querySelector("#db-path"),
  deleteQuestion: document.querySelector("#delete-question"),
  emptyState: document.querySelector("#empty-state"),
  errorList: document.querySelector("#error-list"),
  errorPanel: document.querySelector("#error-panel"),
  learningObjectives: document.querySelector("#learning-objectives"),
  importQuizJson: document.querySelector("#import-quiz-json"),
  metaPanelBody: document.querySelector("#meta-panel-body"),
  questionDifficulty: document.querySelector("#question-difficulty"),
  questionEditor: document.querySelector("#question-editor"),
  questionHeading: document.querySelector("#question-heading"),
  questionId: document.querySelector("#question-id"),
  questionImages: document.querySelector("#question-images"),
  questionImageUpload: document.querySelector("#question-image-upload"),
  questionTextPreview: document.querySelector("#question-text-preview"),
  questionPoints: document.querySelector("#question-points"),
  questionList: document.querySelector("#question-list"),
  questionObjectives: document.querySelector("#question-objectives"),
  questionShuffle: document.querySelector("#question-shuffle"),
  questionText: document.querySelector("#question-text"),
  questionExplanation: document.querySelector("#question-explanation"),
  questionExplanationPreview: document.querySelector("#question-explanation-preview"),
  quizDescription: document.querySelector("#quiz-description"),
  quizTitle: document.querySelector("#quiz-title"),
  reloadQuiz: document.querySelector("#reload-quiz"),
  reloadNotice: document.querySelector("#reload-notice"),
  removeChoice: document.querySelector("#remove-choice"),
  removeObjective: document.querySelector("#remove-objective"),
  unsavedNavBackdrop: document.querySelector("#unsaved-nav-backdrop"),
  unsavedNavModal: document.querySelector("#unsaved-nav-modal"),
  saveQuiz: document.querySelector("#save-quiz"),
  saveStatus: document.querySelector("#save-status"),
  cancelUnsavedNav: document.querySelector("#cancel-unsaved-nav"),
  toggleMetaPanel: document.querySelector("#toggle-meta-panel"),
};

const templates = {
  choice: document.querySelector("#choice-template"),
  location: document.querySelector("#location-template"),
  objective: document.querySelector("#objective-template"),
  objectiveLink: document.querySelector("#objective-link-template"),
};

const META_PANEL_STORAGE_KEY = "quiz-pool:meta-panel-collapsed";

function hasUnsavedChanges() {
  return state.isDirty;
}

function renderUnsavedNavModal() {
  const isOpen = Boolean(state.pendingNavigationHref);
  elements.unsavedNavModal.classList.toggle("is-open", isOpen);
  elements.unsavedNavModal.setAttribute("aria-hidden", String(!isOpen));
}

function openUnsavedNavModal(href) {
  state.pendingNavigationHref = href;
  renderUnsavedNavModal();
}

function closeUnsavedNavModal() {
  state.pendingNavigationHref = "";
  renderUnsavedNavModal();
}

function wireNavigationGuards() {
  for (const link of document.querySelectorAll(".page-link")) {
    link.addEventListener("click", (event) => {
      if (!hasUnsavedChanges()) {
        return;
      }
      event.preventDefault();
      openUnsavedNavModal(link.href);
    });
  }

  window.addEventListener("beforeunload", (event) => {
    if (!hasUnsavedChanges()) {
      return;
    }
    event.preventDefault();
    event.returnValue = "";
  });
}

function getSelectedQuestion() {
  if (!state.quiz) {
    return null;
  }
  return state.quiz.questions[state.selectedQuestionIndex] ?? null;
}

function snapshotQuiz(quiz) {
  return JSON.stringify(quiz);
}

function renderStatus() {
  const suffix = state.isDirty ? " Unsaved changes." : "";
  elements.saveStatus.textContent = `${state.statusMessage}${suffix}`.trim();
  elements.saveStatus.style.color = state.statusIsError
    ? "var(--danger-strong)"
    : state.isDirty
      ? "var(--accent-strong)"
      : "var(--muted)";
}

function setStatus(message, isError = false) {
  state.statusMessage = message;
  state.statusIsError = isError;
  renderStatus();
}

function refreshDirtyState() {
  state.isDirty = Boolean(state.quiz) && snapshotQuiz(state.quiz) !== state.lastSavedSnapshot;
  if (!state.isDirty) {
    state.showReloadNotice = false;
  }
  renderStatus();
  renderReloadNotice();
}

function recordQuizMutation() {
  state.showReloadNotice = false;
  refreshDirtyState();
}

function renderReloadNotice() {
  elements.reloadNotice.classList.toggle("hidden", !(state.showReloadNotice && state.isDirty));
}

function normalizeChoices(rawChoices) {
  const choiceMap = new Map(
    (Array.isArray(rawChoices) ? rawChoices : [])
      .filter((choice) => choice && CHOICE_KEYS.includes(choice.key))
      .map((choice) => [choice.key, choice.text ?? ""]),
  );
  const highestIndex = CHOICE_KEYS.reduce((index, key, currentIndex) => (
    choiceMap.has(key) ? currentIndex : index
  ), MIN_CHOICE_COUNT - 1);
  const count = Math.max(MIN_CHOICE_COUNT, Math.min(MAX_CHOICE_COUNT, highestIndex + 1));
  return CHOICE_KEYS.slice(0, count).map((key) => ({ key, text: choiceMap.get(key) ?? "" }));
}

function normalizeCorrectAnswers(rawAnswers, activeChoices) {
  const activeKeys = new Set(activeChoices.map((choice) => choice.key));
  return CHOICE_KEYS.filter((key) => (
    activeKeys.has(key) && Array.isArray(rawAnswers) && rawAnswers.includes(key)
  ));
}

function normalizeQuestionDraft(question = {}) {
  const choices = normalizeChoices(question.choices);
  const rawLocations = Array.isArray(question.locations) && question.locations.length > 0
    ? question.locations
    : Array.isArray(question.bookLocations) && question.bookLocations.length > 0
      ? question.bookLocations
      : [defaultLocation()];

  return {
    id: question.id ?? nextQuestionId(),
    question: question.question ?? "",
    choices,
    shuffleChoices: Boolean(question.shuffleChoices),
    learningObjectiveIds: Array.isArray(question.learningObjectiveIds) ? [...question.learningObjectiveIds] : [],
    correctAnswers: normalizeCorrectAnswers(question.correctAnswers, choices),
    locations: rawLocations.map((location) => ({
      source: location.source ?? "",
      chapter: location.chapter ?? "",
      section: location.section ?? "",
      page: location.page ?? "",
      url: location.url ?? "",
      reference: location.reference ?? "",
    })),
    points: clampPoints(question.points),
    difficulty: clampDifficulty(question.difficulty),
    explanation: question.explanation ?? "",
    imageAssetIds: Array.isArray(question.imageAssetIds)
      ? question.imageAssetIds.filter((assetId) => typeof assetId === "string" && assetId.trim() !== "")
      : [],
  };
}

function defaultLocation() {
  return {
    source: "",
    chapter: "",
    section: "",
    page: "",
    url: "",
    reference: "",
  };
}

function clampDifficulty(value) {
  const numeric = Number.parseInt(value, 10);
  if (Number.isNaN(numeric)) {
    return 1;
  }
  return Math.min(5, Math.max(1, numeric));
}

function clampPoints(value) {
  const numeric = Number.parseInt(value, 10);
  if (Number.isNaN(numeric)) {
    return 1;
  }
  return Math.max(1, numeric);
}

function nextQuestionId() {
  const ids = new Set((state.quiz?.questions ?? []).map((question) => question.id));
  let counter = (state.quiz?.questions.length ?? 0) + 1;
  while (ids.has(`Q${counter}`)) {
    counter += 1;
  }
  return `Q${counter}`;
}

function nextLearningObjectiveId() {
  const ids = new Set((state.quiz?.learningObjectives ?? []).map((objective) => objective.id));
  let counter = 1;
  while (ids.has(`LO${counter}`)) {
    counter += 1;
  }
  return `LO${counter}`;
}

function currentQuestionIndex() {
  return state.selectedQuestionIndex;
}

function initialQuestionIdFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const questionId = params.get("questionId");
  return typeof questionId === "string" ? questionId.trim() : "";
}

function syncEditorUrl() {
  const url = new URL(window.location.href);
  const question = getSelectedQuestion();
  if (question?.id) {
    url.searchParams.set("questionId", question.id);
  } else {
    url.searchParams.delete("questionId");
  }
  window.history.replaceState({}, "", url);
}

function updateSelectedQuestion(updater) {
  if (!state.quiz) {
    return null;
  }
  const index = currentQuestionIndex();
  if (index === -1) {
    return null;
  }
  const current = normalizeQuestionDraft(state.quiz.questions[index]);
  state.quiz.questions[index] = updater(current);
  recordQuizMutation();
  return state.quiz.questions[index];
}

function renderQuestionList() {
  const fragment = document.createDocumentFragment();
  for (const question of state.quiz.questions) {
    const index = state.quiz.questions.indexOf(question);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "question-tile";
    if (index === state.selectedQuestionIndex) {
      button.classList.add("is-selected");
    }

    const id = document.createElement("span");
    id.className = "question-tile__id";
    id.textContent = question.id || "NO ID";

    const text = document.createElement("span");
    text.className = "question-tile__text";
    text.textContent = stripRichTextMarkup(question.question) || "Untitled question";

    button.append(id, text);
    button.addEventListener("click", () => {
      state.selectedQuestionIndex = index;
      syncEditorUrl();
      render();
    });

    fragment.append(button);
  }

  elements.questionList.replaceChildren(fragment);
}

function renderLearningObjectives() {
  const fragment = document.createDocumentFragment();
  for (const objective of state.quiz.learningObjectives) {
    const node = templates.objective.content.firstElementChild.cloneNode(true);
    const label = node.querySelector("span");
    const input = node.querySelector("input");
    label.textContent = objective.id;
    input.value = objective.label;
    input.addEventListener("input", (event) => {
      objective.label = event.target.value;
      renderQuestionObjectiveLinks();
      renderQuestionList();
      recordQuizMutation();
    });
    fragment.append(node);
  }
  elements.learningObjectives.replaceChildren(fragment);
  elements.removeObjective.disabled = state.quiz.learningObjectives.length === 0;
}

function renderQuestionObjectiveLinks() {
  const question = getSelectedQuestion();
  const fragment = document.createDocumentFragment();

  for (const objective of state.quiz.learningObjectives) {
    const node = templates.objectiveLink.content.firstElementChild.cloneNode(true);
    const input = node.querySelector("input");
    const label = node.querySelector("span");
    input.checked = Boolean(question?.learningObjectiveIds?.includes(objective.id));
    input.disabled = !question;
    label.textContent = `${objective.id} - ${objective.label}`;
    input.addEventListener("change", (event) => {
      updateSelectedQuestion((draft) => {
        const selected = new Set(draft.learningObjectiveIds);
        if (event.target.checked) {
          selected.add(objective.id);
        } else {
          selected.delete(objective.id);
        }
        draft.learningObjectiveIds = [...selected];
        return draft;
      });
    });
    fragment.append(node);
  }

  elements.questionObjectives.replaceChildren(fragment);
}

function renderChoices(question) {
  const fragment = document.createDocumentFragment();
  const previewSyncers = [];
  for (const choice of question.choices) {
    const node = templates.choice.content.firstElementChild.cloneNode(true);
    const checkbox = node.querySelector('input[type="checkbox"]');
    const badge = node.querySelector(".choice-row__correct span");
    const textInput = node.querySelector('.field input');
    const previewField = node.querySelector(".choice-row__preview-field");
    const preview = node.querySelector(".choice-row__preview");

    badge.textContent = `${choice.key} Correct`;
    checkbox.checked = question.correctAnswers.includes(choice.key);
    textInput.value = choice.text;

    const syncPreview = (value) => {
      const isRichText = hasRichTextMarkup(value);
      previewField.classList.toggle("hidden", !isRichText);
      if (isRichText) {
        renderRichTextIntoElement(preview, value || "—");
      } else {
        preview.replaceChildren();
      }
    };

    previewSyncers.push(() => syncPreview(textInput.value));

    checkbox.addEventListener("change", (event) => {
      updateSelectedQuestion((draft) => {
        const answers = new Set(draft.correctAnswers);
        if (event.target.checked) {
          answers.add(choice.key);
        } else {
          answers.delete(choice.key);
        }
        draft.correctAnswers = [...answers];
        return draft;
      });
    });

    textInput.addEventListener("input", (event) => {
      updateSelectedQuestion((draft) => {
        const item = draft.choices.find((entry) => entry.key === choice.key);
        item.text = event.target.value;
        return draft;
      });
      syncPreview(event.target.value);
    });

    fragment.append(node);
  }

  elements.choicesEditor.replaceChildren(fragment);
  for (const syncPreview of previewSyncers) {
    syncPreview();
  }
}

function renderChoiceActions(question) {
  const choiceCount = question?.choices?.length ?? 0;
  const hasQuestion = Boolean(question);
  elements.addChoice.disabled = !hasQuestion || choiceCount >= MAX_CHOICE_COUNT;
  elements.removeChoice.disabled = !hasQuestion || choiceCount <= MIN_CHOICE_COUNT;
}

function renderLocations(question) {
  const fragment = document.createDocumentFragment();
  question.locations.forEach((location, index) => {
    const node = templates.location.content.firstElementChild.cloneNode(true);
    const removeButton = node.querySelector("button");
    const inputs = node.querySelectorAll("[data-key]");

    removeButton.disabled = question.locations.length === 1;
    removeButton.addEventListener("click", () => {
      updateSelectedQuestion((draft) => {
        if (draft.locations.length === 1) {
          draft.locations[0] = defaultLocation();
          return draft;
        }
        draft.locations.splice(index, 1);
        return draft;
      });
      render();
    });

    for (const input of inputs) {
      const key = input.dataset.key;
      input.value = location[key] ?? "";
      input.addEventListener("input", (event) => {
        updateSelectedQuestion((draft) => {
          draft.locations[index][key] = event.target.value;
          return draft;
        });
      });
    }

    fragment.append(node);
  });

  elements.locations.replaceChildren(fragment);
}

function assetUrl(assetId) {
  return `/api/assets/${encodeURIComponent(assetId)}`;
}

function renderQuestionImages(question) {
  const fragment = document.createDocumentFragment();
  const imageAssetIds = question?.imageAssetIds ?? [];

  if (imageAssetIds.length === 0) {
    const empty = document.createElement("p");
    empty.className = "helper-copy";
    empty.textContent = "No images attached.";
    fragment.append(empty);
  }

  imageAssetIds.forEach((assetId, index) => {
    const item = document.createElement("div");
    item.className = "question-image-item";

    const image = document.createElement("img");
    image.className = "question-image-item__preview";
    image.src = assetUrl(assetId);
    image.alt = `Question image ${index + 1}`;
    image.loading = "lazy";

    const actions = document.createElement("div");
    actions.className = "question-image-item__actions";

    const moveUp = document.createElement("button");
    moveUp.type = "button";
    moveUp.className = "button button--tiny button--ghost";
    moveUp.textContent = "Up";
    moveUp.disabled = index === 0;
    moveUp.addEventListener("click", () => {
      updateSelectedQuestion((draft) => {
        [draft.imageAssetIds[index - 1], draft.imageAssetIds[index]] = [
          draft.imageAssetIds[index],
          draft.imageAssetIds[index - 1],
        ];
        return draft;
      });
      render();
    });

    const moveDown = document.createElement("button");
    moveDown.type = "button";
    moveDown.className = "button button--tiny button--ghost";
    moveDown.textContent = "Down";
    moveDown.disabled = index === imageAssetIds.length - 1;
    moveDown.addEventListener("click", () => {
      updateSelectedQuestion((draft) => {
        [draft.imageAssetIds[index], draft.imageAssetIds[index + 1]] = [
          draft.imageAssetIds[index + 1],
          draft.imageAssetIds[index],
        ];
        return draft;
      });
      render();
    });

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "button button--tiny button--danger";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      updateSelectedQuestion((draft) => {
        draft.imageAssetIds.splice(index, 1);
        return draft;
      });
      render();
    });

    actions.append(moveUp, moveDown, remove);
    item.append(image, actions);
    fragment.append(item);
  });

  elements.questionImages.replaceChildren(fragment);
}

function renderQuestionPreviews(question) {
  renderRichTextIntoElement(elements.questionTextPreview, question?.question || "—");
  renderRichTextIntoElement(elements.questionExplanationPreview, question?.explanation || "—");
}

function renderQuestionEditor() {
  const question = getSelectedQuestion();
  const hasQuestion = Boolean(question);

  elements.emptyState.classList.toggle("hidden", hasQuestion);
  elements.questionEditor.classList.toggle("hidden", !hasQuestion);
  elements.deleteQuestion.disabled = !hasQuestion;

  if (!hasQuestion) {
    elements.questionHeading.textContent = "Select a question";
    renderChoiceActions(null);
    renderQuestionObjectiveLinks();
    elements.choicesEditor.replaceChildren();
    elements.locations.replaceChildren();
    elements.questionImages.replaceChildren();
    renderQuestionPreviews(null);
    return;
  }

  const draft = normalizeQuestionDraft(question);
  elements.questionHeading.textContent = draft.id || "Untitled question";
  elements.questionId.value = draft.id;
  elements.questionDifficulty.value = String(draft.difficulty);
  elements.questionPoints.value = String(draft.points);
  elements.questionShuffle.checked = draft.shuffleChoices;
  elements.questionText.value = draft.question;
  elements.questionExplanation.value = draft.explanation;

  renderChoiceActions(draft);
  renderChoices(draft);
  renderQuestionObjectiveLinks();
  renderLocations(draft);
  renderQuestionImages(draft);
  renderQuestionPreviews(draft);
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

function renderMeta() {
  elements.quizTitle.value = state.quiz.title ?? "";
  elements.quizDescription.value = state.quiz.description ?? "";
  elements.dbPath.textContent = state.dbPath;
}

function renderMetaPanel() {
  elements.metaPanelBody.classList.toggle("hidden", state.metaPanelCollapsed);
  elements.toggleMetaPanel.textContent = state.metaPanelCollapsed ? "Expand" : "Collapse";
  elements.toggleMetaPanel.setAttribute("aria-expanded", String(!state.metaPanelCollapsed));
}

function render() {
  if (!state.quiz) {
    return;
  }
  renderMeta();
  renderMetaPanel();
  renderQuestionList();
  renderLearningObjectives();
  renderQuestionEditor();
  renderErrors();
  renderReloadNotice();
  renderStatus();
  syncEditorUrl();
}

async function loadQuiz() {
  setStatus("Loading quiz data...");
  const response = await fetch("/api/quiz");
  if (!response.ok) {
    throw new Error(`Could not load quiz (${response.status})`);
  }
  const payload = await response.json();
  state.dbPath = payload.dbPath;
  state.quiz = payload.quiz;
  state.lastSavedSnapshot = snapshotQuiz(payload.quiz);
  state.isDirty = false;
  const requestedQuestionId = initialQuestionIdFromUrl();
  const requestedIndex = requestedQuestionId
    ? state.quiz.questions.findIndex((question) => question.id === requestedQuestionId)
    : -1;
  state.selectedQuestionIndex = requestedIndex >= 0 ? requestedIndex : 0;
  state.showReloadNotice = false;
  state.validationErrors = [];
  render();
  setStatus("Loaded from disk.");
}

async function saveQuiz() {
  if (!state.quiz) {
    return;
  }

  setStatus("Saving...");
  const response = await fetch("/api/quiz", {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(state.quiz),
  });

  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<unknown>", message: "Save failed" }];
    renderErrors();
    setStatus("Save failed. Fix validation issues and try again.", true);
    return;
  }

  state.lastSavedSnapshot = snapshotQuiz(state.quiz);
  state.isDirty = false;
  state.showReloadNotice = false;
  state.validationErrors = [];
  renderErrors();
  renderReloadNotice();
  setStatus(`Saved at ${new Date().toLocaleTimeString()}.`);
}

async function importQuizJson(path) {
  setStatus("Importing quiz JSON...");
  const response = await fetch("/api/quiz/import-json", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ path }),
  });
  const payload = await response.json();
  if (!response.ok) {
    state.validationErrors = payload.errors ?? [{ path: "<import>", message: "Import failed" }];
    renderErrors();
    setStatus("Import failed. Review the validation messages.", true);
    return;
  }
  state.quiz = payload.quiz;
  state.dbPath = payload.projectPath ?? payload.dbPath;
  state.lastSavedSnapshot = snapshotQuiz(payload.quiz);
  state.isDirty = false;
  state.selectedQuestionIndex = 0;
  state.validationErrors = [];
  render();
  setStatus("Quiz JSON imported into the project DB.");
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.slice(result.indexOf(",") + 1) : result);
    }, { once: true });
    reader.addEventListener("error", () => {
      reject(reader.error || new Error("Could not read image file."));
    }, { once: true });
    reader.readAsDataURL(file);
  });
}

async function uploadImageAsset(file) {
  const dataBase64 = await readFileAsBase64(file);
  const response = await fetch("/api/assets", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      filename: file.name,
      mimeType: file.type,
      dataBase64,
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Image upload failed (${response.status})`);
  }
  return payload.asset;
}

function wireGlobalFields() {
  elements.cancelUnsavedNav.addEventListener("click", () => {
    closeUnsavedNavModal();
  });

  elements.confirmUnsavedNav.addEventListener("click", () => {
    const href = state.pendingNavigationHref;
    closeUnsavedNavModal();
    if (href) {
      window.location.href = href;
    }
  });

  elements.unsavedNavBackdrop.addEventListener("click", () => {
    closeUnsavedNavModal();
  });

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && state.pendingNavigationHref) {
      closeUnsavedNavModal();
    }
  });

  elements.quizTitle.addEventListener("input", (event) => {
    if (!state.quiz) {
      return;
    }
    state.quiz.title = event.target.value;
    recordQuizMutation();
  });

  elements.quizDescription.addEventListener("input", (event) => {
    if (!state.quiz) {
      return;
    }
    state.quiz.description = event.target.value;
    recordQuizMutation();
  });

  elements.toggleMetaPanel.addEventListener("click", () => {
    state.metaPanelCollapsed = !state.metaPanelCollapsed;
    window.localStorage.setItem(META_PANEL_STORAGE_KEY, String(state.metaPanelCollapsed));
    renderMetaPanel();
  });

  elements.questionId.addEventListener("input", (event) => {
    const updated = updateSelectedQuestion((draft) => {
      draft.id = event.target.value;
      return draft;
    });
    elements.questionHeading.textContent = updated?.id || "Untitled question";
    renderQuestionList();
  });

  elements.questionDifficulty.addEventListener("change", (event) => {
    updateSelectedQuestion((draft) => {
      draft.difficulty = clampDifficulty(event.target.value);
      return draft;
    });
  });

  elements.questionPoints.addEventListener("input", (event) => {
    updateSelectedQuestion((draft) => {
      draft.points = clampPoints(event.target.value);
      return draft;
    });
  });

  elements.questionShuffle.addEventListener("change", (event) => {
    updateSelectedQuestion((draft) => {
      draft.shuffleChoices = event.target.checked;
      return draft;
    });
  });

  elements.questionText.addEventListener("input", (event) => {
    updateSelectedQuestion((draft) => {
      draft.question = event.target.value;
      return draft;
    });
    renderQuestionList();
    renderQuestionPreviews(getSelectedQuestion());
  });

  elements.questionExplanation.addEventListener("input", (event) => {
    updateSelectedQuestion((draft) => {
      draft.explanation = event.target.value;
      return draft;
    });
    renderQuestionPreviews(getSelectedQuestion());
  });

  elements.importQuizJson.addEventListener("click", async () => {
    const selectedPath = await browseFile({
      title: "Import Quiz JSON",
      purpose: "quiz-json",
    });
    if (!selectedPath) {
      return;
    }
    await importQuizJson(selectedPath);
  });

  elements.questionImageUpload.addEventListener("change", async (event) => {
    const files = [...(event.target.files ?? [])];
    event.target.value = "";
    if (files.length === 0 || !state.quiz || !getSelectedQuestion()) {
      return;
    }
    try {
      setStatus("Uploading image...");
      const uploadedAssets = [];
      for (const file of files) {
        uploadedAssets.push(await uploadImageAsset(file));
      }
      updateSelectedQuestion((draft) => {
        draft.imageAssetIds.push(...uploadedAssets.map((asset) => asset.assetId));
        return draft;
      });
      render();
      setStatus("Image uploaded. Save the quiz to persist the attachment.");
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  elements.addObjective.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    state.quiz.learningObjectives.push({
      id: nextLearningObjectiveId(),
      label: "",
    });
    recordQuizMutation();
    render();
    setStatus("Learning objective added. Save when ready.");
  });

  elements.removeObjective.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    if (state.quiz.learningObjectives.length === 0) {
      return;
    }
    if (state.quiz.learningObjectives.length === 1 && state.quiz.questions.length > 0) {
      setStatus("At least one learning objective must remain while questions exist.", true);
      return;
    }

    const removedObjective = state.quiz.learningObjectives.pop();
    const fallbackObjectiveId = state.quiz.learningObjectives[0]?.id ?? "";
    for (const question of state.quiz.questions) {
      if (!Array.isArray(question.learningObjectiveIds)) {
        question.learningObjectiveIds = fallbackObjectiveId ? [fallbackObjectiveId] : [];
        continue;
      }
      question.learningObjectiveIds = question.learningObjectiveIds.filter((id) => id !== removedObjective.id);
      if (question.learningObjectiveIds.length === 0 && fallbackObjectiveId) {
        question.learningObjectiveIds = [fallbackObjectiveId];
      }
    }

    recordQuizMutation();
    render();
    setStatus(`Learning objective ${removedObjective.id} removed. Save to persist the change.`);
  });

  elements.addQuestion.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    if (state.quiz.learningObjectives.length === 0) {
      state.quiz.learningObjectives.push({
        id: nextLearningObjectiveId(),
        label: "",
      });
    }
    const newQuestion = normalizeQuestionDraft({
      id: nextQuestionId(),
      learningObjectiveIds: state.quiz.learningObjectives.slice(0, 1).map((objective) => objective.id),
      choices: CHOICE_KEYS.slice(0, MIN_CHOICE_COUNT).map((key) => ({ key, text: "" })),
      correctAnswers: ["A"],
      locations: [defaultLocation()],
      points: 1,
    });
    state.quiz.questions.push(newQuestion);
    state.selectedQuestionIndex = state.quiz.questions.length - 1;
    recordQuizMutation();
    render();
    setStatus("New question added. Save when ready.");
  });

  elements.deleteQuestion.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    const index = currentQuestionIndex();
    if (index === -1) {
      return;
    }
    state.quiz.questions.splice(index, 1);
    state.selectedQuestionIndex = state.quiz.questions.length === 0
      ? 0
      : Math.max(0, Math.min(index, state.quiz.questions.length - 1));
    recordQuizMutation();
    render();
    setStatus("Question removed. Save to persist the deletion.");
  });

  elements.addChoice.addEventListener("click", () => {
    updateSelectedQuestion((draft) => {
      if (draft.choices.length >= MAX_CHOICE_COUNT) {
        return draft;
      }
      draft.choices.push({
        key: CHOICE_KEYS[draft.choices.length],
        text: "",
      });
      return draft;
    });
    render();
  });

  elements.removeChoice.addEventListener("click", () => {
    updateSelectedQuestion((draft) => {
      if (draft.choices.length <= MIN_CHOICE_COUNT) {
        return draft;
      }
      const removedChoice = draft.choices.pop();
      draft.correctAnswers = draft.correctAnswers.filter((answer) => answer !== removedChoice.key);
      if (draft.correctAnswers.length === 0) {
        draft.correctAnswers = [draft.choices[0].key];
      }
      return draft;
    });
    render();
  });

  elements.addLocation.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    updateSelectedQuestion((draft) => {
      draft.locations.push(defaultLocation());
      return draft;
    });
    render();
  });

  elements.reloadQuiz.addEventListener("click", async () => {
    if (state.isDirty) {
      state.showReloadNotice = true;
      renderReloadNotice();
      setStatus("Reload requested. Confirm before discarding local edits.");
      return;
    }
    await loadQuiz();
  });

  elements.cancelReload.addEventListener("click", () => {
    state.showReloadNotice = false;
    renderReloadNotice();
    setStatus("Reload canceled.");
  });

  elements.confirmReload.addEventListener("click", async () => {
    state.showReloadNotice = false;
    renderReloadNotice();
    await loadQuiz();
  });

  elements.saveQuiz.addEventListener("click", async () => {
    await saveQuiz();
  });
}

state.metaPanelCollapsed = window.localStorage.getItem(META_PANEL_STORAGE_KEY) === "true";
wireGlobalFields();
wireNavigationGuards();
loadQuiz().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
