const state = {
  dbPath: "",
  isDirty: false,
  lastSavedSnapshot: "",
  quiz: null,
  selectedQuestionIndex: 0,
  showReloadNotice: false,
  statusIsError: false,
  statusMessage: "Loading quiz data...",
  validationErrors: [],
};

const elements = {
  addQuestion: document.querySelector("#add-question"),
  addLocation: document.querySelector("#add-location"),
  bookLocations: document.querySelector("#book-locations"),
  cancelReload: document.querySelector("#cancel-reload"),
  choicesEditor: document.querySelector("#choices-editor"),
  confirmReload: document.querySelector("#confirm-reload"),
  dbPath: document.querySelector("#db-path"),
  deleteQuestion: document.querySelector("#delete-question"),
  emptyState: document.querySelector("#empty-state"),
  errorList: document.querySelector("#error-list"),
  errorPanel: document.querySelector("#error-panel"),
  learningObjectives: document.querySelector("#learning-objectives"),
  questionDifficulty: document.querySelector("#question-difficulty"),
  questionEditor: document.querySelector("#question-editor"),
  questionHeading: document.querySelector("#question-heading"),
  questionId: document.querySelector("#question-id"),
  questionList: document.querySelector("#question-list"),
  questionObjectives: document.querySelector("#question-objectives"),
  questionShuffle: document.querySelector("#question-shuffle"),
  questionText: document.querySelector("#question-text"),
  questionExplanation: document.querySelector("#question-explanation"),
  quizDescription: document.querySelector("#quiz-description"),
  quizTitle: document.querySelector("#quiz-title"),
  reloadQuiz: document.querySelector("#reload-quiz"),
  reloadNotice: document.querySelector("#reload-notice"),
  saveQuiz: document.querySelector("#save-quiz"),
  saveStatus: document.querySelector("#save-status"),
};

const templates = {
  choice: document.querySelector("#choice-template"),
  location: document.querySelector("#location-template"),
  objective: document.querySelector("#objective-template"),
  objectiveLink: document.querySelector("#objective-link-template"),
};

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

function normalizeQuestionDraft(question = {}) {
  const choiceMap = new Map((question.choices ?? []).map((choice) => [choice.key, choice.text ?? ""]));
  const rawLocations = Array.isArray(question.bookLocations) && question.bookLocations.length > 0
    ? question.bookLocations
    : [defaultBookLocation()];

  return {
    id: question.id ?? nextQuestionId(),
    question: question.question ?? "",
    choices: ["A", "B", "C", "D"].map((key) => ({ key, text: choiceMap.get(key) ?? "" })),
    shuffleChoices: Boolean(question.shuffleChoices),
    learningObjectiveIds: Array.isArray(question.learningObjectiveIds) ? [...question.learningObjectiveIds] : [],
    correctAnswers: Array.isArray(question.correctAnswers) ? [...question.correctAnswers] : [],
    bookLocations: rawLocations.map((location) => ({
      chapter: location.chapter ?? "",
      section: location.section ?? "",
      page: location.page ?? "",
      reference: location.reference ?? "",
    })),
    difficulty: clampDifficulty(question.difficulty),
    explanation: question.explanation ?? "",
  };
}

function defaultBookLocation() {
  return {
    chapter: "",
    section: "",
    page: "",
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

function nextQuestionId() {
  const ids = new Set((state.quiz?.questions ?? []).map((question) => question.id));
  let counter = (state.quiz?.questions.length ?? 0) + 1;
  while (ids.has(`Q${counter}`)) {
    counter += 1;
  }
  return `Q${counter}`;
}

function currentQuestionIndex() {
  return state.selectedQuestionIndex;
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
    text.textContent = question.question || "Untitled question";

    button.append(id, text);
    button.addEventListener("click", () => {
      state.selectedQuestionIndex = index;
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
  for (const choice of question.choices) {
    const node = templates.choice.content.firstElementChild.cloneNode(true);
    const checkbox = node.querySelector('input[type="checkbox"]');
    const badge = node.querySelector(".choice-row__correct span");
    const textInput = node.querySelector('.field input');

    badge.textContent = `${choice.key} Correct`;
    checkbox.checked = question.correctAnswers.includes(choice.key);
    textInput.value = choice.text;

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
    });

    fragment.append(node);
  }

  elements.choicesEditor.replaceChildren(fragment);
}

function renderBookLocations(question) {
  const fragment = document.createDocumentFragment();
  question.bookLocations.forEach((location, index) => {
    const node = templates.location.content.firstElementChild.cloneNode(true);
    const removeButton = node.querySelector("button");
    const inputs = node.querySelectorAll("[data-key]");

    removeButton.disabled = question.bookLocations.length === 1;
    removeButton.addEventListener("click", () => {
      updateSelectedQuestion((draft) => {
        if (draft.bookLocations.length === 1) {
          draft.bookLocations[0] = defaultBookLocation();
          return draft;
        }
        draft.bookLocations.splice(index, 1);
        return draft;
      });
      render();
    });

    for (const input of inputs) {
      const key = input.dataset.key;
      input.value = location[key] ?? "";
      input.addEventListener("input", (event) => {
        updateSelectedQuestion((draft) => {
          draft.bookLocations[index][key] = event.target.value;
          return draft;
        });
      });
    }

    fragment.append(node);
  });

  elements.bookLocations.replaceChildren(fragment);
}

function renderQuestionEditor() {
  const question = getSelectedQuestion();
  const hasQuestion = Boolean(question);

  elements.emptyState.classList.toggle("hidden", hasQuestion);
  elements.questionEditor.classList.toggle("hidden", !hasQuestion);
  elements.deleteQuestion.disabled = !hasQuestion || state.quiz.questions.length === 1;

  if (!hasQuestion) {
    elements.questionHeading.textContent = "Select a question";
    renderQuestionObjectiveLinks();
    elements.choicesEditor.replaceChildren();
    elements.bookLocations.replaceChildren();
    return;
  }

  const draft = normalizeQuestionDraft(question);
  elements.questionHeading.textContent = draft.id || "Untitled question";
  elements.questionId.value = draft.id;
  elements.questionDifficulty.value = String(draft.difficulty);
  elements.questionShuffle.checked = draft.shuffleChoices;
  elements.questionText.value = draft.question;
  elements.questionExplanation.value = draft.explanation;

  renderChoices(draft);
  renderQuestionObjectiveLinks();
  renderBookLocations(draft);
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

function render() {
  if (!state.quiz) {
    return;
  }
  renderMeta();
  renderQuestionList();
  renderLearningObjectives();
  renderQuestionEditor();
  renderErrors();
  renderReloadNotice();
  renderStatus();
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
  state.selectedQuestionIndex = 0;
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

function wireGlobalFields() {
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
  });

  elements.questionExplanation.addEventListener("input", (event) => {
    updateSelectedQuestion((draft) => {
      draft.explanation = event.target.value;
      return draft;
    });
  });

  elements.addQuestion.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    const newQuestion = normalizeQuestionDraft({
      id: nextQuestionId(),
      learningObjectiveIds: state.quiz.learningObjectives.slice(0, 1).map((objective) => objective.id),
      correctAnswers: ["A"],
      bookLocations: [defaultBookLocation()],
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
    if (state.quiz.questions.length === 1) {
      setStatus("At least one question must remain in the quiz.", true);
      return;
    }
    const index = currentQuestionIndex();
    if (index === -1) {
      return;
    }
    state.quiz.questions.splice(index, 1);
    state.selectedQuestionIndex = Math.max(0, Math.min(index, state.quiz.questions.length - 1));
    recordQuizMutation();
    render();
    setStatus("Question removed. Save to persist the deletion.");
  });

  elements.addLocation.addEventListener("click", () => {
    if (!state.quiz) {
      return;
    }
    updateSelectedQuestion((draft) => {
      draft.bookLocations.push(defaultBookLocation());
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

wireGlobalFields();
loadQuiz().catch((error) => {
  console.error(error);
  setStatus(error.message, true);
});
