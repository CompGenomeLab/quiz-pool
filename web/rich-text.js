function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

const MATH_TAG_PATTERN = /\[math\]([\s\S]*?)\[\/math\]/giu;
const MATH_NODE_SELECTOR = ".qp-math[data-qp-math-source]";
const MATHJAX_SCRIPT_PATH = "/vendor/mathjax/tex-svg-full.js";

let mathJaxPromise = null;
let mutationObserver = null;
let flushPromise = null;
const scheduledRoots = new Set();

function renderPlainTextSegment(value) {
  return escapeHtml(String(value ?? "").replace(/\r\n?/gu, "\n")).replaceAll("\n", "<br />");
}

function renderMathExpression(value) {
  const source = String(value ?? "").trim();
  if (!source) {
    return "";
  }
  return (
    `<span class="qp-math" data-qp-math-source="${escapeHtml(source)}" data-qp-math-state="pending">`
    + `<code class="qp-math__source">${escapeHtml(source)}</code>`
    + "</span>"
  );
}

function renderRichTextHtml(value) {
  const source = String(value ?? "");
  let cursor = 0;
  let output = "";

  for (const match of source.matchAll(MATH_TAG_PATTERN)) {
    const start = match.index ?? 0;
    output += renderPlainTextSegment(source.slice(cursor, start));
    output += renderMathExpression(match[1] ?? "");
    cursor = start + match[0].length;
  }

  output += renderPlainTextSegment(source.slice(cursor));
  return output;
}

function stripRichTextMarkup(value) {
  return String(value ?? "")
    .replace(MATH_TAG_PATTERN, (_, expression) => ` ${String(expression ?? "").trim()} `)
    .replace(/\s+/gu, " ")
    .trim();
}

function markMathRenderFailure(node, source) {
  node.dataset.qpMathState = "error";
  node.classList.add("qp-math--error");
  node.replaceChildren();
  const fallback = document.createElement("code");
  fallback.className = "qp-math__source";
  fallback.textContent = source;
  node.append(fallback);
}

function normalizeRoot(root) {
  if (root instanceof Element || root instanceof DocumentFragment || root instanceof Document) {
    return root;
  }
  return document;
}

function collectPendingMathNodes(root) {
  const nodes = [];
  if (root instanceof Element && root.matches(MATH_NODE_SELECTOR)) {
    nodes.push(root);
  }
  if (root instanceof DocumentFragment || root instanceof Element || root instanceof Document) {
    nodes.push(...root.querySelectorAll(MATH_NODE_SELECTOR));
  }
  return nodes.filter((node) => node.dataset.qpMathState !== "done" && node.dataset.qpMathState !== "rendering");
}

function ensureMathJaxConfig() {
  if (!window.MathJax || !window.MathJax.startup) {
    window.MathJax = {
      startup: {
        typeset: false,
      },
      tex: {
        packages: {
          "[-]": ["noerrors", "noundefined"],
        },
      },
      svg: {
        fontCache: "local",
      },
    };
    return;
  }

  window.MathJax.startup.typeset = false;
  const packages = window.MathJax.tex?.packages;
  if (!packages) {
    window.MathJax.tex = { ...(window.MathJax.tex ?? {}), packages: { "[-]": ["noerrors", "noundefined"] } };
  }
}

function ensureMathJax() {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return Promise.resolve(null);
  }
  if (window.MathJax?.tex2svgPromise && window.MathJax?.startup?.promise) {
    return window.MathJax.startup.promise.then(() => window.MathJax);
  }
  if (mathJaxPromise) {
    return mathJaxPromise;
  }

  ensureMathJaxConfig();
  mathJaxPromise = new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-quiz-pool-mathjax="true"]');
    if (existing) {
      existing.addEventListener("load", () => {
        Promise.resolve(window.MathJax?.startup?.promise)
          .then(() => resolve(window.MathJax))
          .catch(reject);
      }, { once: true });
      existing.addEventListener("error", () => {
        reject(new Error("Could not load the local MathJax bundle."));
      }, { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = MATHJAX_SCRIPT_PATH;
    script.async = true;
    script.dataset.quizPoolMathjax = "true";
    script.addEventListener("load", () => {
      Promise.resolve(window.MathJax?.startup?.promise)
        .then(() => resolve(window.MathJax))
        .catch(reject);
    }, { once: true });
    script.addEventListener("error", () => {
      reject(new Error("Could not load the local MathJax bundle."));
    }, { once: true });
    document.head.append(script);
  }).catch((error) => {
    mathJaxPromise = null;
    throw error;
  });

  return mathJaxPromise;
}

async function renderPendingMathNodes(nodes) {
  if (nodes.length === 0) {
    return;
  }

  let mathJax = null;
  try {
    mathJax = await ensureMathJax();
  } catch {
    for (const node of nodes) {
      markMathRenderFailure(node, node.dataset.qpMathSource ?? "");
    }
    return;
  }

  for (const node of nodes) {
    const source = node.dataset.qpMathSource ?? "";
    if (!source.trim()) {
      node.remove();
      continue;
    }
    node.dataset.qpMathState = "rendering";
    try {
      const rendered = await mathJax.tex2svgPromise(source, { display: false });
      if (rendered.querySelector?.('[data-mml-node="merror"], [data-mjx-error], mjx-merror')) {
        markMathRenderFailure(node, source);
        continue;
      }
      node.classList.remove("qp-math--error");
      node.replaceChildren();
      node.append(rendered);
      node.dataset.qpMathState = "done";
    } catch {
      markMathRenderFailure(node, source);
    }
  }
}

function flushScheduledMath() {
  const roots = [...scheduledRoots];
  scheduledRoots.clear();
  flushPromise = null;

  const nodes = [];
  const seen = new Set();
  for (const root of roots) {
    for (const node of collectPendingMathNodes(root)) {
      if (seen.has(node)) {
        continue;
      }
      seen.add(node);
      nodes.push(node);
    }
  }

  return renderPendingMathNodes(nodes);
}

function scheduleMathTypeset(root = document) {
  if (typeof document === "undefined") {
    return Promise.resolve();
  }
  scheduledRoots.add(normalizeRoot(root));
  if (!flushPromise) {
    flushPromise = Promise.resolve().then(flushScheduledMath);
  }
  return flushPromise;
}

function ensureMathObserver() {
  if (mutationObserver || typeof document === "undefined" || typeof MutationObserver === "undefined") {
    return;
  }

  const attach = () => {
    if (mutationObserver || !document.documentElement) {
      return;
    }
    mutationObserver = new MutationObserver((records) => {
      for (const record of records) {
        for (const node of record.addedNodes) {
          if (node instanceof Element || node instanceof DocumentFragment) {
            scheduleMathTypeset(node);
          }
        }
      }
    });
    mutationObserver.observe(document.documentElement, {
      childList: true,
      subtree: true,
    });
    scheduleMathTypeset(document);
  };

  if (document.readyState === "loading") {
    window.addEventListener("DOMContentLoaded", attach, { once: true });
  } else {
    attach();
  }
}

function renderRichTextIntoElement(element, value) {
  const source = String(value ?? "");
  element.dataset.richTextSource = source;
  element.innerHTML = renderRichTextHtml(source);
  ensureMathObserver();
  scheduleMathTypeset(element);
}

function renderRichTextTargets(root = document) {
  const normalizedRoot = normalizeRoot(root);
  const elements = [];
  if (normalizedRoot instanceof Element && normalizedRoot.matches("[data-rich-text]")) {
    elements.push(normalizedRoot);
  }
  if (
    normalizedRoot instanceof Document
    || normalizedRoot instanceof DocumentFragment
    || normalizedRoot instanceof Element
  ) {
    elements.push(...normalizedRoot.querySelectorAll("[data-rich-text]"));
  }
  for (const element of elements) {
    const source = element.dataset.richTextSource ?? element.textContent ?? "";
    renderRichTextIntoElement(element, source);
  }
}

ensureMathObserver();

export {
  escapeHtml,
  renderRichTextHtml,
  renderRichTextIntoElement,
  renderRichTextTargets,
  stripRichTextMarkup,
};
