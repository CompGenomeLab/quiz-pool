// Global project-DB snapshot control. Self-contained: injects a "Snapshots"
// button into the page nav and a modal listing snapshots with restore/delete,
// plus a "Take snapshot now" action. Loaded on every tool page.

const KIND_LABELS = {
  baseline: "Startup baseline",
  auto: "Auto",
  manual: "Manual",
  prerestore: "Before restore",
};

function formatTimestamp(iso) {
  const date = new Date(iso);
  return Number.isNaN(date.getTime()) ? iso : date.toLocaleString();
}

function formatSize(bytes) {
  if (!Number.isFinite(bytes)) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function buildModal() {
  const modal = document.createElement("div");
  modal.id = "snapshots-modal";
  modal.className = "detail-modal";
  modal.setAttribute("aria-hidden", "true");
  modal.innerHTML = `
    <div id="snapshots-backdrop" class="detail-modal__backdrop"></div>
    <section class="detail-modal__panel" role="dialog" aria-modal="true" aria-labelledby="snapshots-title">
      <div class="panel__head">
        <div>
          <p class="eyebrow">Project DB</p>
          <h2 id="snapshots-title">Snapshots</h2>
        </div>
      </div>
      <p class="question-detail-sheet__copy">
        Restore the entire project DB (quiz, exams, drafts, grading runs, and
        images) to an earlier point. A snapshot is taken automatically at startup,
        every 15 minutes when the project changes, and just before any restore.
      </p>
      <div class="panel__subhead">
        <div></div>
        <div class="panel__actions">
          <button id="snapshots-take" class="button button--primary" type="button">Take snapshot now</button>
        </div>
      </div>
      <p id="snapshots-status" class="helper-copy" hidden></p>
      <div class="table-wrap">
        <table class="pool-table">
          <thead>
            <tr><th>Created</th><th>Type</th><th>Size</th><th>Actions</th></tr>
          </thead>
          <tbody id="snapshots-body"></tbody>
        </table>
      </div>
      <div class="panel__actions">
        <button id="snapshots-close" class="button button--ghost" type="button">Close</button>
      </div>
    </section>`;
  document.body.append(modal);
  return modal;
}

function injectButton(onClick) {
  const nav = document.querySelector(".page-links");
  if (!nav) return;
  const button = document.createElement("button");
  button.type = "button";
  button.id = "snapshots-open";
  button.className = "page-link";
  button.textContent = "Snapshots";
  button.addEventListener("click", onClick);
  nav.append(button);
}

function setStatus(el, message, isError = false) {
  if (!message) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.textContent = message;
  el.style.color = isError ? "var(--danger, #b00020)" : "";
}

async function apiJson(url, options) {
  const response = await fetch(url, options);
  let payload = {};
  try {
    payload = await response.json();
  } catch (error) {
    payload = {};
  }
  return { ok: response.ok && payload.ok !== false, payload };
}

function init() {
  const modal = buildModal();
  const body = modal.querySelector("#snapshots-body");
  const status = modal.querySelector("#snapshots-status");

  const open = () => {
    modal.classList.add("is-open");
    modal.setAttribute("aria-hidden", "false");
    refresh();
  };
  const close = () => {
    modal.classList.remove("is-open");
    modal.setAttribute("aria-hidden", "true");
  };

  async function refresh() {
    setStatus(status, "Loading snapshots...");
    const { ok, payload } = await apiJson("/api/snapshots");
    if (!ok) {
      setStatus(status, "Could not load snapshots.", true);
      return;
    }
    setStatus(status, "");
    render(payload.snapshots || []);
  }

  function render(snapshots) {
    body.replaceChildren();
    if (snapshots.length === 0) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 4;
      td.textContent = "No snapshots yet.";
      tr.append(td);
      body.append(tr);
      return;
    }
    for (const snap of snapshots) {
      const tr = document.createElement("tr");
      const created = document.createElement("td");
      created.textContent = formatTimestamp(snap.createdAt);
      const kind = document.createElement("td");
      kind.textContent = KIND_LABELS[snap.kind] || snap.kind;
      const size = document.createElement("td");
      size.textContent = formatSize(snap.sizeBytes);
      const actions = document.createElement("td");
      const restore = document.createElement("button");
      restore.type = "button";
      restore.className = "button";
      restore.textContent = "Restore";
      restore.addEventListener("click", () => doRestore(snap));
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "button button--ghost";
      remove.textContent = "Delete";
      remove.addEventListener("click", () => doDelete(snap));
      actions.append(restore, remove);
      tr.append(created, kind, size, actions);
      body.append(tr);
    }
  }

  async function take() {
    setStatus(status, "Taking snapshot...");
    const { ok, payload } = await apiJson("/api/snapshots", { method: "POST" });
    if (!ok) {
      setStatus(status, "Snapshot failed.", true);
      return;
    }
    render(payload.snapshots || []);
    setStatus(status, "Snapshot saved.");
  }

  async function doDelete(snap) {
    if (!window.confirm(`Delete this ${KIND_LABELS[snap.kind] || snap.kind} snapshot from ${formatTimestamp(snap.createdAt)}?`)) {
      return;
    }
    const { ok, payload } = await apiJson(`/api/snapshots/${encodeURIComponent(snap.id)}`, { method: "DELETE" });
    if (!ok) {
      setStatus(status, "Delete failed.", true);
      return;
    }
    render(payload.snapshots || []);
    setStatus(status, "Snapshot deleted.");
  }

  async function doRestore(snap) {
    if (!window.confirm(
      `Restore the entire project DB to the snapshot from ${formatTimestamp(snap.createdAt)}?\n\n`
      + "This replaces the current quiz, exams, drafts, grading runs, and images. "
      + "A safety snapshot of the current state is taken first. The page will reload.",
    )) {
      return;
    }
    setStatus(status, "Restoring...");
    const { ok } = await apiJson("/api/snapshots/restore", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: snap.id }),
    });
    if (!ok) {
      setStatus(status, "Restore failed.", true);
      return;
    }
    window.location.reload();
  }

  injectButton(open);
  modal.querySelector("#snapshots-close").addEventListener("click", close);
  modal.querySelector("#snapshots-backdrop").addEventListener("click", close);
  modal.querySelector("#snapshots-take").addEventListener("click", take);
}

init();
