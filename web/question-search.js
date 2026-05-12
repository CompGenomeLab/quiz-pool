import { stripRichTextMarkup } from "./rich-text.js";

function safeText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

function stripRich(value) {
  return stripRichTextMarkup(safeText(value));
}

export function buildQuestionSearchText(question, options = {}) {
  const { objectiveLabel } = options;
  const parts = [];

  parts.push(safeText(question.id));
  parts.push(stripRich(question.question));
  parts.push(stripRich(question.explanation));
  parts.push(safeText(question.difficulty));
  parts.push(safeText(question.points));

  for (const choice of question.choices || []) {
    if (!choice) continue;
    parts.push(safeText(choice.key));
    parts.push(stripRich(choice.text));
  }

  for (const answer of question.correctAnswers || []) {
    parts.push(safeText(answer));
  }

  for (const objectiveId of question.learningObjectiveIds || []) {
    parts.push(safeText(objectiveId));
    if (typeof objectiveLabel === "function") {
      parts.push(safeText(objectiveLabel(objectiveId)));
    }
  }

  const locations = Array.isArray(question.locations)
    ? question.locations
    : Array.isArray(question.bookLocations)
      ? question.bookLocations
      : [];
  for (const location of locations) {
    if (!location || typeof location !== "object") continue;
    parts.push(safeText(location.chapter));
    parts.push(safeText(location.section));
    parts.push(safeText(location.page));
    parts.push(safeText(location.source));
    parts.push(safeText(location.url));
    parts.push(safeText(location.reference));
  }

  for (const assetId of question.imageAssetIds || []) {
    parts.push(safeText(assetId));
  }

  return parts.join("  ").toLowerCase();
}

function subsequenceScore(haystack, needle) {
  let hi = 0;
  let score = 0;
  let prevIndex = -2;
  for (let ni = 0; ni < needle.length; ni += 1) {
    const ch = needle[ni];
    let found = -1;
    while (hi < haystack.length) {
      if (haystack[hi] === ch) {
        found = hi;
        hi += 1;
        break;
      }
      hi += 1;
    }
    if (found === -1) {
      return -1;
    }
    score += found === prevIndex + 1 ? 3 : 1;
    prevIndex = found;
  }
  return score;
}

function tokenScore(haystack, token) {
  if (!token) return 0;
  const idx = haystack.indexOf(token);
  if (idx !== -1) {
    let score = token.length * 6;
    if (idx === 0) score += 8;
    else if (haystack[idx - 1] === " " || haystack[idx - 1] === "") score += 4;
    return score;
  }
  return subsequenceScore(haystack, token);
}

export function fuzzyQueryScore(haystack, query) {
  const q = (query || "").trim().toLowerCase();
  if (!q) return 1;
  const tokens = q.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return 1;
  let total = 0;
  for (const token of tokens) {
    const score = tokenScore(haystack, token);
    if (score <= 0) return 0;
    total += score;
  }
  return total;
}
