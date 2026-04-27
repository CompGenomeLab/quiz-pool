let activeDialog = null;

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function fetchListing(path, purpose) {
  const params = new URLSearchParams({ purpose });
  if (path) {
    params.set("path", path);
  }
  const response = await fetch(`/api/file-browser?${params.toString()}`);
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not list files (${response.status})`);
  }
  return payload;
}

function closeActiveDialog(result = null) {
  if (!activeDialog) {
    return;
  }
  const { modal, resolve } = activeDialog;
  activeDialog = null;
  modal.remove();
  resolve(result);
}

function createDialog(title) {
  const modal = document.createElement("div");
  modal.className = "file-browser-modal";
  modal.innerHTML = `
    <div class="file-browser-modal__backdrop"></div>
    <section class="file-browser-modal__panel" role="dialog" aria-modal="true">
      <div class="panel__head">
        <div>
          <p class="eyebrow">Files</p>
          <h2>${escapeHtml(title)}</h2>
        </div>
        <button class="button button--ghost" type="button" data-file-browser-close>Close</button>
      </div>
      <p class="status-card__value" data-file-browser-path></p>
      <div class="file-browser-actions">
        <button class="button button--ghost" type="button" data-file-browser-home>Home</button>
        <button class="button button--ghost" type="button" data-file-browser-up>Up</button>
        <button class="button button--primary" type="button" data-file-browser-select-folder>Select This Folder</button>
      </div>
      <div class="file-browser-list" data-file-browser-list></div>
    </section>
  `;
  document.body.append(modal);
  return modal;
}

function renderListing(modal, listing, navigate, select) {
  modal.querySelector("[data-file-browser-path]").textContent = listing.currentPath;
  const list = modal.querySelector("[data-file-browser-list]");
  const fragment = document.createDocumentFragment();

  if (listing.entries.length === 0) {
    const empty = document.createElement("p");
    empty.className = "helper-copy";
    empty.textContent = "No matching files in this folder.";
    fragment.append(empty);
  }

  for (const entry of listing.entries) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "file-browser-row";
    button.innerHTML = `
      <span class="file-browser-row__icon">${entry.isDirectory ? "Folder" : "File"}</span>
      <span class="file-browser-row__name">${escapeHtml(entry.name)}</span>
    `;
    button.addEventListener("click", () => {
      if (entry.isDirectory) {
        navigate(entry.path);
      } else {
        select(entry.path);
      }
    });
    fragment.append(button);
  }
  list.replaceChildren(fragment);

  modal.querySelector("[data-file-browser-up]").disabled = !listing.parentPath;
  modal.querySelector("[data-file-browser-up]").onclick = () => navigate(listing.parentPath);
  modal.querySelector("[data-file-browser-home]").onclick = () => navigate(listing.homePath);
  const selectFolderButton = modal.querySelector("[data-file-browser-select-folder]");
  const canSelectFolder = listing.purpose === "directory" || listing.purpose === "pdf-or-dir";
  selectFolderButton.classList.toggle("hidden", !canSelectFolder);
  selectFolderButton.onclick = () => select(listing.currentPath);
}

function browseFile({ title, purpose, startPath = "" }) {
  if (activeDialog) {
    closeActiveDialog(null);
  }
  const modal = createDialog(title);
  const promise = new Promise((resolve) => {
    activeDialog = { modal, resolve };
  });

  async function navigate(path) {
    try {
      const listing = await fetchListing(path, purpose);
      renderListing(modal, listing, navigate, (selectedPath) => closeActiveDialog(selectedPath));
    } catch (error) {
      const list = modal.querySelector("[data-file-browser-list]");
      list.innerHTML = `<p class="helper-copy">${escapeHtml(error.message)}</p>`;
    }
  }

  modal.querySelector("[data-file-browser-close]").addEventListener("click", () => closeActiveDialog(null));
  modal.querySelector(".file-browser-modal__backdrop").addEventListener("click", () => closeActiveDialog(null));
  window.addEventListener("keydown", function onKeydown(event) {
    if (!activeDialog || activeDialog.modal !== modal) {
      window.removeEventListener("keydown", onKeydown);
      return;
    }
    if (event.key === "Escape") {
      closeActiveDialog(null);
      window.removeEventListener("keydown", onKeydown);
    }
  });
  void navigate(startPath);
  return promise;
}

export { browseFile };
