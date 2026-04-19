from __future__ import annotations

import argparse
import asyncio
import base64
import html
import io
import json
import math
import mimetypes
import random
import re
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from uuid import uuid4
import zipfile

from jsonschema import Draft202012Validator
from omr import SheetConfig, annotate_pdf, generate_omr_sheet, grade_path
from pypdf import PdfReader, PdfWriter
from pyppeteer import launch
import segno


PACKAGE_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PACKAGE_ROOT.parent
ROOT = SRC_ROOT.parent
WEB_ROOT = ROOT / "web"
DEFAULT_DB = ROOT / "sample_quiz.json"
DEFAULT_SCHEMA = ROOT / "scheme.json"
DEFAULT_EXAM_STORE_NAME = "generated_exams.json"
DISPLAY_KEYS = ("A", "B", "C", "D", "E")
DEFAULT_PRINTABLE_FOLDER = "exam-printables"
QUESTION_POOL_PRINTABLE_NAME = "question-pool.pdf"
LATEX_STUDENT_TEMPLATE = ROOT / "tex_templates" / "exam_template.tex"
LATEX_QUESTION_POOL_TEMPLATE = ROOT / "tex_templates" / "pool_template.tex"
LATEX_ENGINES = ("xelatex", "lualatex", "pdflatex")
LATEX_VARIANT_QR_ASSET_NAME = "variant-qr.pdf"
QUESTION_PAGE_CAPACITY = 130
MAX_QUESTIONS_PER_EXAM = 100
DEFAULT_OMR_INSTRUCTIONS = (
    "Fully fill bubbles. Do not leave any Student ID field blank. Include leading zeros."
)
DEFAULT_EXAM_RULES = [
    DEFAULT_OMR_INSTRUCTIONS,
    "Read every question carefully and select all correct answers for each question.",
    "Mark answers clearly and keep your paper neat for printing, photocopying, and scanning.",
    "Do not communicate with other students or use unauthorized materials during the exam.",
    "Remain seated until instructed to stop and submit your paper.",
]
GRADE_ALLOWED_LABELS = set(DISPLAY_KEYS)
MATH_TAG_PATTERN = re.compile(r"\[math\]([\s\S]*?)\[/math\]", re.IGNORECASE)
LATEX_TEXT_ESCAPES = {
    "\\": r"\textbackslash{}",
    "{": r"\{",
    "}": r"\}",
    "#": r"\#",
    "$": r"\$",
    "%": r"\%",
    "&": r"\&",
    "_": r"\_",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def default_exam_rules(omr_instructions: str = DEFAULT_OMR_INSTRUCTIONS) -> list[str]:
    return [
        omr_instructions,
        "Read every question carefully and select all correct answers for each question.",
        "Mark answers clearly and keep your paper neat for printing, photocopying, and scanning.",
        "Do not communicate with other students or use unauthorized materials during the exam.",
        "Remain seated until instructed to stop and submit your paper.",
    ]


@dataclass
class AppState:
    db_path: Path
    schema_path: Path
    exam_store_path: Path
    validator: Draft202012Validator


def default_exam_store_path_for(db_path: Path) -> Path:
    return (db_path.parent / DEFAULT_EXAM_STORE_NAME).resolve()


def display_default_db_path() -> str:
    return str(DEFAULT_DB.relative_to(ROOT))


def display_default_exam_store_path() -> str:
    return str(default_exam_store_path_for(DEFAULT_DB).relative_to(ROOT))


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "w",
        encoding="utf-8",
        dir=path.parent,
        delete=False,
    ) as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
        handle.write("\n")
        temp_path = Path(handle.name)
    temp_path.replace(path)


def validation_errors(
    validator: Draft202012Validator, payload: dict[str, Any]
) -> list[dict[str, str]]:
    errors: list[dict[str, str]] = []
    for error in sorted(validator.iter_errors(payload), key=lambda item: list(item.path)):
        path = ".".join(str(part) for part in error.path) or "<root>"
        errors.append({"path": path, "message": error.message})
    return errors


def load_exam_store(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"examSets": []}
    if path.stat().st_size == 0:
        return {"examSets": []}

    payload = load_json(path)
    if not isinstance(payload, dict):
        raise ValueError("Exam store must be a JSON object")

    exam_sets = payload.get("examSets")
    if not isinstance(exam_sets, list):
        raise ValueError("Exam store must contain an 'examSets' array")

    return payload


def validate_quiz_file(path: Path, validator: Draft202012Validator) -> None:
    if not path.is_file():
        raise ValueError(f"Quiz file not found: {path}")

    quiz = load_json(path)
    errors = validation_errors(validator, quiz)
    if errors:
        message = "; ".join(f"{item['path']}: {item['message']}" for item in errors)
        raise ValueError(f"Quiz file does not match the schema: {message}")


def validate_exam_store_file(path: Path) -> None:
    if path.exists():
        load_exam_store(path)


def set_active_paths(state: AppState, *, db_path: Path, exam_store_path: Path) -> None:
    validate_quiz_file(db_path, state.validator)
    validate_exam_store_file(exam_store_path)
    state.db_path = db_path
    state.exam_store_path = exam_store_path


def append_exam_set(path: Path, exam_set: dict[str, Any]) -> None:
    store = load_exam_store(path)
    store["examSets"].append(exam_set)
    write_json_atomic(path, store)


def delete_exam_set(path: Path, exam_set_id: str) -> bool:
    store = load_exam_store(path)
    exam_sets = store.get("examSets", [])
    remaining_exam_sets: list[Any] = []
    deleted = False

    for exam_set in exam_sets:
        if (
            not deleted
            and isinstance(exam_set, dict)
            and exam_set.get("examSetId") == exam_set_id
        ):
            deleted = True
            continue
        remaining_exam_sets.append(exam_set)

    if not deleted:
        return False

    store["examSets"] = remaining_exam_sets
    write_json_atomic(path, store)
    return True


def find_variant(path: Path, variant_id: str) -> tuple[dict[str, Any], dict[str, Any]] | None:
    store = load_exam_store(path)
    for exam_set in store["examSets"]:
        if not isinstance(exam_set, dict):
            continue
        for variant in exam_set.get("variants", []):
            if isinstance(variant, dict) and variant.get("variantId") == variant_id:
                return exam_set, variant
    return None


def find_exam_set(path: Path, exam_set_id: str) -> dict[str, Any] | None:
    store = load_exam_store(path)
    for exam_set in store["examSets"]:
        if isinstance(exam_set, dict) and exam_set.get("examSetId") == exam_set_id:
            return exam_set
    return None


def list_exam_sets(path: Path) -> list[dict[str, Any]]:
    store = load_exam_store(path)
    exam_sets = [exam_set for exam_set in store.get("examSets", []) if isinstance(exam_set, dict)]
    exam_sets.sort(key=lambda item: str(item.get("generatedAt") or ""), reverse=True)
    return exam_sets


def dedupe_preserve_order(items: list[Any]) -> list[Any]:
    seen: set[Any] = set()
    deduped: list[Any] = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        deduped.append(item)
    return deduped


def load_web_text_asset(name: str) -> str:
    return (WEB_ROOT / name).read_text(encoding="utf-8")


def render_rich_text_css() -> str:
    return """
      .qp-math {
        align-items: flex-end;
        display: inline-flex;
        margin: 0 0.08em;
        max-width: 100%;
        vertical-align: -0.08em;
        white-space: nowrap;
      }

      .qp-mrow {
        align-items: flex-end;
        display: inline-flex;
        gap: 0.04em;
        white-space: nowrap;
      }

      .qp-mi,
      .qp-mn,
      .qp-mo,
      .qp-mtext,
      .qp-script__base,
      .qp-script__sup,
      .qp-script__sub,
      .qp-frac__top,
      .qp-frac__bottom,
      .qp-root__index,
      .qp-root__radical,
      .qp-root__body {
        line-height: 1;
      }

      .qp-mi {
        font-style: italic;
      }

      .qp-mi--normal,
      .qp-mtext {
        font-style: normal;
      }

      .qp-mspace {
        display: inline-block;
        width: var(--qp-space-width, 0.2em);
      }

      .qp-accent {
        align-items: center;
        display: inline-flex;
        flex-direction: column;
        margin: 0 0.03em;
        white-space: nowrap;
      }

      .qp-accent__glyph,
      .qp-accent__body {
        line-height: 1;
      }

      .qp-accent__glyph {
        display: block;
      }

      .qp-accent--bar .qp-accent__glyph {
        border-top: 1px solid currentColor;
        margin-bottom: 0.08em;
        min-width: 100%;
        width: 100%;
      }

      .qp-accent--glyph .qp-accent__glyph {
        font-size: 0.68em;
        margin-bottom: -0.08em;
        transform: translateY(0.06em) scaleX(1.08);
      }

      .qp-script {
        align-items: baseline;
        display: inline-flex;
        white-space: nowrap;
      }

      .qp-script--sup,
      .qp-script--sub {
        gap: 0.05em;
      }

      .qp-script__sup,
      .qp-script__sub {
        display: inline-block;
        font-size: 0.68em;
      }

      .qp-script--sup .qp-script__sup {
        transform: translateY(-0.45em);
      }

      .qp-script--sub .qp-script__sub {
        transform: translateY(0.4em);
      }

      .qp-script__stack {
        display: inline-flex;
        flex-direction: column;
        margin-left: 0.05em;
      }

      .qp-script__stack .qp-script__sup {
        transform: translateY(-0.08em);
      }

      .qp-script__stack .qp-script__sub {
        transform: translateY(0.14em);
      }

      .qp-frac {
        align-items: center;
        display: inline-flex;
        flex-direction: column;
        justify-content: center;
        margin: 0 0.08em;
        vertical-align: middle;
      }

      .qp-frac__top,
      .qp-frac__bottom {
        display: block;
        font-size: 0.82em;
        padding: 0 0.14em;
      }

      .qp-frac__bar {
        border-top: 1px solid currentColor;
        display: block;
        margin: 0.06em 0 0.04em;
        min-width: 100%;
        width: 100%;
      }

      .qp-root {
        align-items: flex-end;
        display: inline-flex;
        margin: 0 0.04em;
        white-space: nowrap;
      }

      .qp-root__index {
        font-size: 0.56em;
        margin-right: 0.02em;
        transform: translateY(-0.52em);
      }

      .qp-root__radical {
        font-size: 1.15em;
      }

      .qp-root__body {
        border-top: 1px solid currentColor;
        margin-left: -0.05em;
        padding: 0.08em 0 0 0.14em;
      }

      .qp-math__source {
        font-family: "Courier New", monospace;
        font-size: 0.92em;
      }

      .qp-math[data-qp-math-state="pending"] .qp-math__source,
      .qp-math[data-qp-math-state="rendering"] .qp-math__source {
        opacity: 0.68;
      }

      .qp-math mjx-container[jax="SVG"] {
        display: inline-block !important;
        margin: 0 !important;
        max-width: 100%;
      }

      .qp-math mjx-container[jax="SVG"] > svg {
        display: block;
        max-width: 100%;
        overflow: visible;
      }

      .qp-math--error {
        background: rgba(201, 106, 106, 0.08);
        border: 1px solid rgba(159, 77, 77, 0.18);
        border-radius: 10px;
        font-family: "Courier New", monospace;
        font-size: 0.92em;
        padding: 0.16em 0.4em;
      }
"""


def render_inline_rich_text_module(script_body: str) -> str:
    return f"""
    <script type="module">
{load_web_text_asset("rich-text.js")}
{script_body}
    </script>
"""


def strip_math_markup(value: Any) -> str:
    text = str(value or "")
    without_tags = MATH_TAG_PATTERN.sub(
        lambda match: f" {match.group(1).strip()} ",
        text,
    )
    return " ".join(without_tags.split())


def latex_escape_text_segment(value: Any, *, preserve_linebreaks: bool = True) -> str:
    text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    if not preserve_linebreaks:
        text = re.sub(r"\s+", " ", text)
    escaped = "".join(LATEX_TEXT_ESCAPES.get(character, character) for character in text)
    if preserve_linebreaks:
        escaped = escaped.replace("\n", r"\newline ")
    return escaped


def render_rich_text_latex(value: Any, *, preserve_linebreaks: bool = True) -> str:
    source = str(value or "")
    if not source:
        return ""

    cursor = 0
    parts: list[str] = []
    for match in MATH_TAG_PATTERN.finditer(source):
        start = match.start()
        parts.append(
            latex_escape_text_segment(
                source[cursor:start],
                preserve_linebreaks=preserve_linebreaks,
            )
        )
        expression = (match.group(1) or "").strip()
        if expression:
            parts.append(rf"\({expression}\)")
        cursor = match.end()

    parts.append(
        latex_escape_text_segment(
            source[cursor:],
            preserve_linebreaks=preserve_linebreaks,
        )
    )
    return "".join(parts)


def render_latex_text_or_dash(
    value: Any,
    *,
    preserve_linebreaks: bool = False,
    fallback: str = "—",
) -> str:
    if strip_math_markup(value).strip():
        return (
            render_rich_text_latex(value, preserve_linebreaks=preserve_linebreaks).strip()
            or latex_escape_text_segment(fallback, preserve_linebreaks=False)
        )
    return latex_escape_text_segment(fallback, preserve_linebreaks=False)


def latex_placeholder_value(value: Any, *, preserve_linebreaks: bool = False) -> str:
    return render_latex_text_or_dash(value, preserve_linebreaks=preserve_linebreaks)


def latex_placeholder_value_or_blank(value: Any, *, preserve_linebreaks: bool = False) -> str:
    if strip_math_markup(value).strip():
        return render_rich_text_latex(value, preserve_linebreaks=preserve_linebreaks).strip()
    return ""


def render_latex_rules(rules: list[Any]) -> str:
    rendered_rules = []
    for rule in rules:
        content = render_rich_text_latex(rule, preserve_linebreaks=False).strip()
        if content:
            rendered_rules.append(f"  \\item {content}")
    return "\n".join(rendered_rules)


def render_latex_choice_rows(choices: list[dict[str, Any]]) -> str:
    rows: list[str] = []
    for index in range(0, len(choices), 2):
        left = choices[index] if index < len(choices) else None
        right = choices[index + 1] if index + 1 < len(choices) else None

        left_cell = "~"
        if isinstance(left, dict):
            left_cell = (
                f"\\textbf{{{latex_escape_text_segment(left.get('key', ''), preserve_linebreaks=False)}.}} "
                + render_rich_text_latex(left.get("text", ""), preserve_linebreaks=True)
            ).strip()

        right_cell = "~"
        if isinstance(right, dict):
            right_cell = (
                f"\\textbf{{{latex_escape_text_segment(right.get('key', ''), preserve_linebreaks=False)}.}} "
                + render_rich_text_latex(right.get("text", ""), preserve_linebreaks=True)
            ).strip()

        rows.append(f"{left_cell} & {right_cell}")

    return " \\\\\n".join(rows)


def render_student_template_choice_rows(choices: list[dict[str, Any]]) -> str:
    rows: list[str] = []
    for choice in choices:
        if not isinstance(choice, dict):
            continue
        key = latex_escape_text_segment(choice.get("key", ""), preserve_linebreaks=False)
        text = render_rich_text_latex(choice.get("text", ""), preserve_linebreaks=True).strip() or "~"
        rows.append(f"{key}. & {text}")
    return " \\\\\n      ".join(rows) or "~ & ~"


def variant_question_heading(variant: dict[str, Any]) -> str:
    questions = [question for question in variant.get("questions", []) if isinstance(question, dict)]
    if not questions:
        return "Answer each question."

    has_multiple_correct = any(
        len(
            [
                answer
                for answer in question.get("sourceCorrectAnswers", question.get("displayCorrectAnswers", []))
                if isinstance(answer, str) and answer.strip()
            ]
        ) > 1
        for question in questions
    )
    if has_multiple_correct:
        return "Choose all correct answers for each question."
    return "Choose the single best answer for each question."


def render_student_question_blocks_latex(variant: dict[str, Any]) -> str:
    blocks: list[str] = []
    for question in variant.get("questions", []):
        if not isinstance(question, dict):
            continue
        choices = [
            choice
            for choice in question.get("displayChoices", [])
            if isinstance(choice, dict)
        ]
        choice_rows = render_student_template_choice_rows(choices)
        prompt = render_rich_text_latex(question.get("question", ""), preserve_linebreaks=True).strip() or "—"
        blocks.append(
            "\n".join(
                [
                    r"\mcqitem",
                    rf"  {{{prompt}}}",
                    rf"  {{{choice_rows}}}",
                ]
            )
        )
    return "\n\n".join(blocks)


def render_question_pool_blocks_latex(question_pool: list[dict[str, Any]]) -> str:
    blocks: list[str] = []
    for question in question_pool:
        if not isinstance(question, dict):
            continue
        chapters = ", ".join(str(chapter).strip() for chapter in question.get("chapters", []) if str(chapter).strip()) or "—"
        correct_answers = ", ".join(
            str(answer).strip() for answer in question.get("sourceCorrectAnswers", []) if str(answer).strip()
        ) or "—"
        objectives = ", ".join(
            str(objective.get("label") or objective.get("id") or "").strip()
            for objective in question.get("learningObjectives", [])
            if isinstance(objective, dict) and str(objective.get("label") or objective.get("id") or "").strip()
        ) or "—"
        prompt = render_rich_text_latex(question.get("question", ""), preserve_linebreaks=True).strip() or "—"
        choices = [
            choice
            for choice in question.get("choices", [])
            if isinstance(choice, dict)
        ]
        meta_bits = [
            latex_escape_text_segment(str(question.get("sourceQuestionId") or "—"), preserve_linebreaks=False),
            f"Difficulty {int(question.get('difficulty') or 0)}",
            f"{int(question.get('points') or 1)} pt",
            f"Chapters: {latex_escape_text_segment(chapters, preserve_linebreaks=False)}",
        ]
        blocks.append(
            "\n".join(
                [
                    r"\item",
                    rf"  {{\small\textbf{{{' \\textbullet{} '.join(meta_bits)}}}}}\par",
                    f"  {prompt}",
                    r"  \vspace{0.2em}",
                    r"  {\renewcommand{\arraystretch}{1.12}%",
                    r"  \begin{tabularx}{\linewidth}{@{}YY@{}}",
                    "  " + (render_latex_choice_rows(choices) or "~ & ~"),
                    r"  \end{tabularx}}",
                    (
                        r"  {\small\textbf{Correct:} "
                        + latex_escape_text_segment(correct_answers, preserve_linebreaks=False)
                        + r"\quad\textbf{Objectives:} "
                        + latex_escape_text_segment(objectives, preserve_linebreaks=False)
                        + "}"
                    ),
                ]
            )
        )
    return "\n\n".join(blocks)


def apply_latex_template(template_path: Path, replacements: dict[str, str]) -> str:
    content = template_path.read_text(encoding="utf-8")
    for placeholder, value in replacements.items():
        content = content.replace(placeholder, value)
    return content


def build_question_index(
    quiz: dict[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[str, int], list[dict[str, str]]]:
    question_by_id: dict[str, dict[str, Any]] = {}
    order_by_id: dict[str, int] = {}
    errors: list[dict[str, str]] = []

    for index, question in enumerate(quiz.get("questions", [])):
        question_id = question.get("id")
        if not isinstance(question_id, str) or not question_id.strip():
            errors.append(
                {
                    "path": f"questions.{index}.id",
                    "message": "Question id must be a non-empty string for exam generation",
                }
            )
            continue
        if question_id in question_by_id:
            errors.append(
                {
                    "path": f"questions.{index}.id",
                    "message": f"Duplicate question id: {question_id}",
                }
            )
            continue
        question_by_id[question_id] = question
        order_by_id[question_id] = index

    return question_by_id, order_by_id, errors


def normalize_positive_int(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> int | None:
    value = payload.get(key)
    if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
        errors.append({"path": key, "message": f"{key} must be a positive integer"})
        return None
    return value


def normalize_string_list(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> list[str]:
    value = payload.get(key, [])
    if not isinstance(value, list):
        errors.append({"path": key, "message": f"{key} must be an array of strings"})
        return []

    items: list[str] = []
    for index, item in enumerate(value):
        if not isinstance(item, str):
            errors.append({"path": f"{key}.{index}", "message": "Must be a string"})
            continue
        normalized = item.strip()
        if not normalized:
            errors.append({"path": f"{key}.{index}", "message": "Must not be empty"})
            continue
        items.append(normalized)

    return dedupe_preserve_order(items)


def normalize_difficulty_list(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> list[int]:
    value = payload.get(key, [])
    if not isinstance(value, list):
        errors.append({"path": key, "message": f"{key} must be an array of integers"})
        return []

    items: list[int] = []
    for index, item in enumerate(value):
        if not isinstance(item, int) or isinstance(item, bool) or item < 1 or item > 5:
            errors.append({"path": f"{key}.{index}", "message": "Difficulty must be an integer from 1 to 5"})
            continue
        items.append(item)

    return dedupe_preserve_order(items)


def normalize_optional_string(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> str:
    value = payload.get(key, "")
    if value is None:
        return ""
    if not isinstance(value, str):
        errors.append({"path": key, "message": f"{key} must be a string"})
        return ""
    return value.strip()


def normalize_optional_positive_int(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> int | None:
    value = payload.get(key)
    if value in (None, ""):
        return None
    if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
        errors.append({"path": key, "message": f"{key} must be a positive integer"})
        return None
    return value


def normalize_rule_list(
    payload: dict[str, Any], key: str, errors: list[dict[str, str]]
) -> list[str]:
    value = payload.get(key, [])
    if value in (None, ""):
        return []

    if isinstance(value, str):
        items = [line.strip() for line in value.splitlines() if line.strip()]
        return dedupe_preserve_order(items)

    if not isinstance(value, list):
        errors.append({"path": key, "message": f"{key} must be an array of strings or a newline-delimited string"})
        return []

    items: list[str] = []
    for index, item in enumerate(value):
        if not isinstance(item, str):
            errors.append({"path": f"{key}.{index}", "message": "Must be a string"})
            continue
        normalized = item.strip()
        if not normalized:
            errors.append({"path": f"{key}.{index}", "message": "Must not be empty"})
            continue
        items.append(normalized)

    return dedupe_preserve_order(items)


def extract_question_chapters(question: dict[str, Any]) -> list[str]:
    chapters: list[str] = []
    for location in question.get("bookLocations", []):
        if not isinstance(location, dict):
            continue
        chapter = location.get("chapter")
        if isinstance(chapter, str) and chapter.strip():
            chapters.append(chapter.strip())
    return dedupe_preserve_order(chapters)


def question_matches_filters(question: dict[str, Any], request: dict[str, Any]) -> bool:
    selected_chapters = set(request["chapters"])
    if selected_chapters:
        if not selected_chapters.intersection(extract_question_chapters(question)):
            return False

    selected_difficulties = set(request["difficulties"])
    if selected_difficulties and question.get("difficulty") not in selected_difficulties:
        return False

    selected_objectives = set(request["learningObjectiveIds"])
    if selected_objectives and not selected_objectives.intersection(question.get("learningObjectiveIds", [])):
        return False

    return True


def normalize_generation_request(
    payload: Any, quiz: dict[str, Any]
) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
    if not isinstance(payload, dict):
        return None, [{"path": "<body>", "message": "Generation payload must be a JSON object"}]

    question_by_id, _, question_errors = build_question_index(quiz)
    errors = list(question_errors)

    question_count = normalize_positive_int(payload, "questionCount", errors)
    variant_count = normalize_positive_int(payload, "variantCount", errors)
    chapters = normalize_string_list(payload, "chapters", errors)
    difficulties = normalize_difficulty_list(payload, "difficulties", errors)
    learning_objective_ids = normalize_string_list(payload, "learningObjectiveIds", errors)
    include_question_ids = normalize_string_list(payload, "includeQuestionIds", errors)
    exclude_question_ids = normalize_string_list(payload, "excludeQuestionIds", errors)
    institution_name = normalize_optional_string(payload, "institutionName", errors)
    exam_name = normalize_optional_string(payload, "examName", errors)
    course_name = normalize_optional_string(payload, "courseName", errors)
    exam_date = normalize_optional_string(payload, "examDate", errors)
    start_time = normalize_optional_string(payload, "startTime", errors)
    total_time_minutes = normalize_optional_positive_int(payload, "totalTimeMinutes", errors)
    instructor = normalize_optional_string(payload, "instructor", errors)
    allowed_materials = normalize_optional_string(payload, "allowedMaterials", errors)
    omr_instructions = normalize_optional_string(payload, "omrInstructions", errors)
    exam_rules = normalize_rule_list(payload, "examRules", errors)

    known_objective_ids = {
        objective["id"]
        for objective in quiz.get("learningObjectives", [])
        if isinstance(objective, dict) and isinstance(objective.get("id"), str)
    }
    for objective_id in learning_objective_ids:
        if objective_id not in known_objective_ids:
            errors.append(
                {
                    "path": "learningObjectiveIds",
                    "message": f"Unknown learning objective id: {objective_id}",
                }
            )

    for question_id in include_question_ids:
        if question_id not in question_by_id:
            errors.append({"path": "includeQuestionIds", "message": f"Unknown question id: {question_id}"})
    for question_id in exclude_question_ids:
        if question_id not in question_by_id:
            errors.append({"path": "excludeQuestionIds", "message": f"Unknown question id: {question_id}"})

    overlap = set(include_question_ids).intersection(exclude_question_ids)
    if overlap:
        overlap_list = ", ".join(sorted(overlap))
        errors.append(
            {
                "path": "includeQuestionIds",
                "message": f"Question ids cannot be both included and excluded: {overlap_list}",
            }
        )

    if question_count is not None and question_count > MAX_QUESTIONS_PER_EXAM:
        errors.append(
            {
                "path": "questionCount",
                "message": f"questionCount cannot be greater than {MAX_QUESTIONS_PER_EXAM}",
            }
        )

    if errors:
        return None, errors

    return (
        {
            "questionCount": question_count,
            "variantCount": variant_count,
            "chapters": chapters,
            "difficulties": difficulties,
            "learningObjectiveIds": learning_objective_ids,
            "includeQuestionIds": include_question_ids,
            "excludeQuestionIds": exclude_question_ids,
            "institutionName": institution_name,
            "examName": exam_name,
            "courseName": course_name,
            "examDate": exam_date,
            "startTime": start_time,
            "totalTimeMinutes": total_time_minutes,
            "instructor": instructor,
            "allowedMaterials": allowed_materials,
            "omrInstructions": omr_instructions,
            "examRules": exam_rules,
        },
        [],
    )


def normalize_grading_request(payload: Any) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
    if not isinstance(payload, dict):
        return None, [{"path": "<body>", "message": "Grading payload must be a JSON object"}]

    errors: list[dict[str, str]] = []
    input_path = normalize_optional_string(payload, "inputPath", errors)
    if not input_path:
        errors.append({"path": "inputPath", "message": "inputPath must be a non-empty path to a PDF or directory"})
    if errors:
        return None, errors
    return ({"inputPath": input_path}, [])


def normalize_annotation_request(payload: Any) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
    if not isinstance(payload, dict):
        return None, [{"path": "<body>", "message": "Annotation payload must be a JSON object"}]

    errors: list[dict[str, str]] = []
    input_path = normalize_optional_string(payload, "inputPath", errors)
    output_path = normalize_optional_string(payload, "outputPath", errors)
    if not input_path:
        errors.append({"path": "inputPath", "message": "inputPath must be a non-empty path to a PDF or directory"})
    if not output_path:
        errors.append({"path": "outputPath", "message": "outputPath must be a non-empty output directory path"})
    if errors:
        return None, errors
    return ({"inputPath": input_path, "outputPath": output_path}, [])


def run_omr_grade(input_path: Path) -> list[dict[str, Any]]:
    try:
        payload = grade_path(input_path)
    except Exception as error:
        if input_path.is_file():
            return [
                {
                    "source_pdf": input_path.name,
                    "qr_data": None,
                    "student_id": "",
                    "marked_answers": {},
                    "omr_error": str(error),
                }
            ]
        raise ValueError(str(error)) from error

    if isinstance(payload, list):
        return [asdict(item) for item in payload]

    source_name = input_path.name if input_path.is_file() else str(input_path)
    return [{"source_pdf": source_name, **asdict(payload)}]


def run_omr_annotate(
    input_path: Path,
    output_path: Path,
    *,
    correct_answers: dict[str, list[str]] | None = None,
) -> dict[str, Any]:
    try:
        return asdict(
            annotate_pdf(
                input_path,
                output_path,
                correct_answers=correct_answers,
            )
        )
    except Exception as error:
        raise ValueError(str(error)) from error


def normalize_marked_answers(raw_value: Any) -> dict[str, list[str]]:
    if not isinstance(raw_value, dict):
        return {}

    normalized: dict[str, list[str]] = {}
    for raw_key, raw_answers in raw_value.items():
        key = str(raw_key).strip()
        if not key:
            continue
        labels: list[str] = []
        if isinstance(raw_answers, list):
            for answer in raw_answers:
                if not isinstance(answer, str):
                    continue
                label = answer.strip().upper()
                if label:
                    labels.append(label)
        normalized[key] = dedupe_preserve_order(labels)
    return normalized


def earns_full_credit(marked: list[str], correct: list[str]) -> bool:
    return bool(marked) and set(marked) == set(correct)


def build_variant_lookup(store: dict[str, Any]) -> dict[str, tuple[dict[str, Any], dict[str, Any]]]:
    lookup: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {}
    for exam_set in store.get("examSets", []):
        if not isinstance(exam_set, dict):
            continue
        for variant in exam_set.get("variants", []):
            if not isinstance(variant, dict):
                continue
            variant_id = variant.get("variantId")
            if isinstance(variant_id, str) and variant_id.strip():
                lookup[variant_id] = (exam_set, variant)
    return lookup


def analyze_grade_result(
    result: dict[str, Any],
    variant_lookup: dict[str, tuple[dict[str, Any], dict[str, Any]]],
) -> dict[str, Any]:
    source_pdf = str(result.get("source_pdf") or "")
    student_id = str(result.get("student_id") or "").strip()
    qr_data = result.get("qr_data")
    marked_answers = normalize_marked_answers(result.get("marked_answers"))
    omr_error = str(result.get("omr_error") or "").strip()
    issues: list[str] = []
    question_details: list[dict[str, Any]] = []
    exam_set_id = ""
    variant_id = ""
    exam_name = ""
    variant_question_count = 0
    detected_question_count = len(marked_answers)
    summary = {
        "correctCount": 0,
        "incorrectCount": 0,
        "blankCount": 0,
        "missingCount": 0,
        "invalidCount": 0,
        "earnedPoints": 0,
        "possiblePoints": 0,
    }
    has_mismatch = False

    if omr_error:
        issues.append(f"OMR error: {omr_error}")

    if not isinstance(qr_data, dict):
        issues.append("QR data is missing or could not be decoded as JSON.")
    else:
        exam_set_id = str(qr_data.get("examSetId") or "").strip()
        variant_id = str(qr_data.get("variantId") or "").strip()
        if not exam_set_id:
            issues.append("QR data is missing examSetId.")
        if not variant_id:
            issues.append("QR data is missing variantId.")

    matched_record = variant_lookup.get(variant_id) if variant_id else None
    if variant_id and matched_record is None:
        issues.append(f"Variant {variant_id} was not found in the exam store.")

    if matched_record is not None:
        exam_set, variant = matched_record
        exam_name = str(get_print_settings(exam_set).get("examName") or "")
        stored_exam_set_id = str(exam_set.get("examSetId") or "")
        if exam_set_id and exam_set_id != stored_exam_set_id:
            issues.append(
                f"QR examSetId {exam_set_id} does not match stored exam set {stored_exam_set_id} for variant {variant_id}."
            )
        exam_set_id = stored_exam_set_id
        questions = [question for question in variant.get("questions", []) if isinstance(question, dict)]
        variant_question_count = len(questions)
        if detected_question_count != variant_question_count:
            issues.append(
                f"Detected {detected_question_count} question rows but variant expects {variant_question_count}."
            )
        expected_positions = {str(question.get("position")) for question in questions}
        unexpected_positions = sorted(key for key in marked_answers if key not in expected_positions)
        if unexpected_positions:
            issues.append(
                "Detected unexpected question rows: " + ", ".join(unexpected_positions) + "."
            )

        question_by_position = {
            str(question.get("position")): question
            for question in questions
            if isinstance(question.get("position"), int)
        }
        for position in range(1, variant_question_count + 1):
            position_key = str(position)
            question = question_by_position.get(position_key, {})
            allowed = [
                str(choice.get("key"))
                for choice in question.get("displayChoices", [])
                if isinstance(choice, dict) and isinstance(choice.get("key"), str)
            ]
            correct = [
                str(label)
                for label in question.get("displayCorrectAnswers", [])
                if isinstance(label, str)
            ]
            points = int(question.get("points") or 1)
            marked = marked_answers.get(position_key)
            status = "missing"
            earned_points = 0
            detail_issues: list[str] = []
            summary["possiblePoints"] += points
            if marked is None:
                summary["missingCount"] += 1
                detail_issues.append("Question row was not detected.")
            else:
                invalid_labels = [label for label in marked if label not in allowed or label not in GRADE_ALLOWED_LABELS]
                if invalid_labels:
                    status = "invalid"
                    summary["invalidCount"] += 1
                    detail_issues.append(
                        f"Marked invalid choice(s): {', '.join(invalid_labels)}. Allowed choices: {', '.join(allowed) or 'none'}."
                    )
                elif not marked:
                    status = "blank"
                    summary["blankCount"] += 1
                elif earns_full_credit(marked, correct):
                    status = "correct"
                    summary["correctCount"] += 1
                    earned_points = points
                    summary["earnedPoints"] += points
                else:
                    status = "incorrect"
                    summary["incorrectCount"] += 1
            if detail_issues:
                has_mismatch = True
            question_details.append(
                {
                    "position": position,
                    "prompt": str(question.get("question") or ""),
                    "allowedChoices": allowed,
                    "correctAnswers": correct,
                    "markedAnswers": marked if marked is not None else [],
                    "points": points,
                    "earnedPoints": earned_points,
                    "status": status,
                    "issues": detail_issues,
                }
            )

    if issues:
        has_mismatch = True

    if omr_error:
        row_status = "omr_error"
    elif has_mismatch:
        row_status = "mismatch"
    else:
        row_status = "ok"

    return {
        "sourcePdf": source_pdf,
        "studentId": student_id,
        "displayStudentId": student_id or "Unknown",
        "qrData": qr_data,
        "examSetId": exam_set_id,
        "variantId": variant_id,
        "examName": exam_name,
        "omrError": omr_error,
        "issues": issues,
        "hasMismatch": has_mismatch,
        "status": row_status,
        "detectedQuestionCount": detected_question_count,
        "variantQuestionCount": variant_question_count,
        "summary": summary,
        "questionDetails": question_details,
    }


def grade_exam_pdfs(state: AppState, input_path: Path) -> dict[str, Any]:
    if not input_path.exists():
        raise ValueError(f"Input path not found: {input_path}")
    if input_path.is_file() and input_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF")
    if not input_path.is_dir() and not input_path.is_file():
        raise ValueError("Input path must be a PDF file or a directory containing PDFs")

    store = load_exam_store(state.exam_store_path)
    variant_lookup = build_variant_lookup(store)
    raw_results = run_omr_grade(input_path)
    rows = [analyze_grade_result(result, variant_lookup) for result in raw_results]
    duplicate_student_rows: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        student_id = row["studentId"]
        if not student_id:
            continue
        duplicate_student_rows.setdefault(student_id, []).append(row)

    duplicated_student_ids = {
        student_id: student_rows
        for student_id, student_rows in duplicate_student_rows.items()
        if len(student_rows) > 1
    }
    for student_id, student_rows in duplicated_student_ids.items():
        source_pdfs = ", ".join(
            sorted({row["sourcePdf"] or "<unknown source>" for row in student_rows})
        )
        issue = (
            f"Duplicate student ID {student_id} appears in multiple graded PDFs: {source_pdfs}."
        )
        for row in student_rows:
            row["issues"].append(issue)
            row["hasMismatch"] = True
            if row["status"] == "ok":
                row["status"] = "mismatch"

    rows.sort(key=lambda row: (row["studentId"] == "", row["studentId"], row["sourcePdf"]))
    for index, row in enumerate(rows, start=1):
        row["rowIndex"] = index

    return {
        "inputPath": str(input_path),
        "examStorePath": str(state.exam_store_path),
        "rows": rows,
        "summary": {
            "processedCount": len(rows),
            "knownStudentCount": sum(1 for row in rows if row["studentId"]),
            "duplicateStudentIdCount": len(duplicated_student_ids),
            "omrErrorCount": sum(1 for row in rows if row["omrError"]),
            "mismatchCount": sum(1 for row in rows if row["hasMismatch"]),
        },
    }


def build_annotation_answer_key(row: dict[str, Any]) -> dict[str, list[str]]:
    answer_key: dict[str, list[str]] = {}
    for question in row.get("questionDetails", []):
        if not isinstance(question, dict):
            continue
        position = question.get("position")
        correct_answers = question.get("correctAnswers")
        if not isinstance(position, int) or not isinstance(correct_answers, list):
            continue
        normalized = [str(label).strip().upper() for label in correct_answers if isinstance(label, str) and str(label).strip()]
        answer_key[str(position)] = dedupe_preserve_order(normalized)
    return answer_key


def resolve_annotation_source_path(input_path: Path, row: dict[str, Any]) -> Path:
    if input_path.is_file():
        return input_path

    source_pdf = str(row.get("sourcePdf") or "").strip()
    if not source_pdf:
        raise ValueError("A graded row is missing its source PDF name, so it cannot be annotated.")

    source_path = (input_path / source_pdf).resolve()
    if not source_path.is_file():
        raise ValueError(f"Graded source PDF not found for annotation: {source_path}")
    return source_path


def annotation_output_filename(row: dict[str, Any]) -> str:
    row_index = row.get("rowIndex")
    student_id = str(row.get("studentId") or "").strip() or "unknown"
    normalized_student_id = "".join(
        character if character.isalnum() or character in {"-", "_"} else "-"
        for character in student_id
    ).strip("-") or "unknown"
    if not isinstance(row_index, int) or row_index <= 0:
        raise ValueError("Annotated file naming requires a valid grading row index.")
    return f"{row_index}-{normalized_student_id}-annotated.pdf"


def annotate_exam_pdfs(state: AppState, input_path: Path, output_path: Path) -> dict[str, Any]:
    grading_result = grade_exam_pdfs(state, input_path)
    output_path.mkdir(parents=True, exist_ok=True)

    annotation_rows: list[dict[str, Any]] = []
    for row in grading_result["rows"]:
        source_path = resolve_annotation_source_path(input_path, row)
        answer_key = build_annotation_answer_key(row)
        target_output_path = output_path / annotation_output_filename(row)
        annotate_payload = run_omr_annotate(
            source_path,
            target_output_path,
            correct_answers=answer_key or None,
        )
        annotated_pdf = str(annotate_payload.get("annotated_pdf") or "").strip()
        omr_error = str(annotate_payload.get("omr_error") or "").strip()
        issues = list(row.get("issues", []))
        if not answer_key:
            issues.append("Annotation used no answer key because the exam variant could not be matched.")
        if omr_error:
            issues.append(f"omr-annotate reported an OMR error: {omr_error}")

        annotation_rows.append(
            {
                "rowIndex": row.get("rowIndex", 0),
                "sourcePdf": row.get("sourcePdf", ""),
                "annotatedPdf": annotated_pdf,
                "studentId": row.get("studentId", ""),
                "displayStudentId": row.get("displayStudentId", "Unknown"),
                "examSetId": row.get("examSetId", ""),
                "variantId": row.get("variantId", ""),
                "usedAnswerKey": bool(answer_key),
                "omrError": omr_error,
                "issues": issues,
                "status": "omr_error" if omr_error else ("review" if issues else "ok"),
            }
        )

    return {
        "inputPath": str(input_path),
        "outputPath": str(output_path),
        "rows": annotation_rows,
        "summary": {
            "processedCount": len(annotation_rows),
            "annotatedCount": sum(1 for row in annotation_rows if row["annotatedPdf"]),
            "omrErrorCount": sum(1 for row in annotation_rows if row["omrError"]),
            "usedAnswerKeyCount": sum(1 for row in annotation_rows if row["usedAnswerKey"]),
        },
    }


def unrank_permutation(items: list[Any], rank: int) -> list[Any]:
    available = list(items)
    result: list[Any] = []

    for size in range(len(items), 0, -1):
        factorial = math.factorial(size - 1)
        index, rank = divmod(rank, factorial)
        result.append(available.pop(index))

    return result


def sample_unique_ranks(total: int, count: int, rng: random.SystemRandom) -> list[int]:
    if total <= 50000 and count > total // 3:
        pool = list(range(total))
        rng.shuffle(pool)
        return pool[:count]

    seen: set[int] = set()
    sampled: list[int] = []
    while len(sampled) < count:
        rank = rng.randrange(total)
        if rank in seen:
            continue
        seen.add(rank)
        sampled.append(rank)
    return sampled


def build_variant_signature(questions: list[dict[str, Any]]) -> str:
    parts = [",".join(question["sourceQuestionId"] for question in questions)]
    for question in questions:
        if question["shuffleChoices"]:
            parts.append(
                f"{question['sourceQuestionId']}:{''.join(choice['sourceKey'] for choice in question['displayChoices'])}"
            )
    return "|".join(parts)


def build_question_pool_entry(
    question: dict[str, Any], objective_labels: dict[str, str]
) -> dict[str, Any]:
    return {
        "sourceQuestionId": question["id"],
        "question": question["question"],
        "difficulty": question["difficulty"],
        "points": question["points"],
        "chapters": extract_question_chapters(question),
        "learningObjectiveIds": list(question["learningObjectiveIds"]),
        "learningObjectives": [
            {"id": objective_id, "label": objective_labels.get(objective_id, objective_id)}
            for objective_id in question["learningObjectiveIds"]
        ],
        "shuffleChoices": bool(question["shuffleChoices"]),
        "bookLocations": question["bookLocations"],
        "choices": [
            {"key": choice["key"], "text": choice["text"]}
            for choice in question["choices"]
        ],
        "sourceCorrectAnswers": list(question["correctAnswers"]),
        "explanation": question.get("explanation", ""),
    }


def build_exam_set_summary(exam_set: dict[str, Any]) -> dict[str, Any]:
    print_settings = get_print_settings(exam_set)
    selection = exam_set.get("selection", {})
    variants = [variant for variant in exam_set.get("variants", []) if isinstance(variant, dict)]
    return {
        "examSetId": str(exam_set.get("examSetId") or ""),
        "generatedAt": str(exam_set.get("generatedAt") or ""),
        "quiz": exam_set.get("quiz", {}),
        "printSettings": print_settings,
        "selectedQuestionCount": len(selection.get("selectedQuestionIds", [])),
        "variantCount": len(variants),
    }


def get_print_settings(exam_set: dict[str, Any]) -> dict[str, Any]:
    raw = exam_set.get("printSettings")
    if not isinstance(raw, dict):
        raw = {}

    institution_name = raw.get("institutionName")
    exam_name = raw.get("examName")
    course_name = raw.get("courseName")
    exam_date = raw.get("examDate")
    start_time = raw.get("startTime")
    total_time_minutes = raw.get("totalTimeMinutes")
    instructor = raw.get("instructor")
    allowed_materials = raw.get("allowedMaterials")
    omr_instructions = raw.get("omrInstructions")
    exam_rules = raw.get("examRules")

    normalized_institution_name = institution_name.strip() if isinstance(institution_name, str) else ""
    normalized_exam_name = exam_name.strip() if isinstance(exam_name, str) else ""
    normalized_course_name = course_name.strip() if isinstance(course_name, str) else ""
    normalized_exam_date = exam_date.strip() if isinstance(exam_date, str) else ""
    normalized_start_time = start_time.strip() if isinstance(start_time, str) else ""
    normalized_instructor = instructor.strip() if isinstance(instructor, str) else ""
    normalized_allowed_materials = (
        allowed_materials.strip() if isinstance(allowed_materials, str) else ""
    )
    normalized_total_time = ""
    if isinstance(total_time_minutes, int) and total_time_minutes > 0:
        normalized_total_time = str(total_time_minutes)
    elif isinstance(total_time_minutes, str) and total_time_minutes.strip():
        normalized_total_time = total_time_minutes.strip()
    normalized_omr_instructions = (
        omr_instructions.strip() if isinstance(omr_instructions, str) and omr_instructions.strip() else DEFAULT_OMR_INSTRUCTIONS
    )
    normalized_rules: list[str] = []
    if isinstance(exam_rules, list):
        normalized_rules = [
            rule.strip() for rule in exam_rules if isinstance(rule, str) and rule.strip()
        ]
    elif isinstance(exam_rules, str) and exam_rules.strip():
        normalized_rules = [line.strip() for line in exam_rules.splitlines() if line.strip()]

    fallback_title = exam_set.get("quiz", {}).get("title", "")
    if not normalized_exam_name and isinstance(fallback_title, str):
        normalized_exam_name = fallback_title.strip()
    if not normalized_rules:
        normalized_rules = default_exam_rules(normalized_omr_instructions)

    return {
        "institutionName": normalized_institution_name or "Institution Name",
        "examName": normalized_exam_name,
        "courseName": normalized_course_name,
        "examDate": normalized_exam_date,
        "startTime": normalized_start_time,
        "totalTimeMinutes": normalized_total_time,
        "instructor": normalized_instructor,
        "allowedMaterials": normalized_allowed_materials,
        "omrInstructions": normalized_omr_instructions,
        "examRules": normalized_rules,
    }


def build_variant(
    selected_questions: list[dict[str, Any]],
    rank: int,
    exam_set_id: str,
    objective_labels: dict[str, str],
) -> dict[str, Any]:
    question_space = math.factorial(len(selected_questions))
    question_rank = rank % question_space
    choice_rank_state = rank // question_space

    rendered_choices_by_question: dict[str, tuple[list[dict[str, str]], list[str]]] = {}
    for question in selected_questions:
        source_correct_answers = list(question["correctAnswers"])
        if question["shuffleChoices"]:
            choices = list(question["choices"])
            choice_space = math.factorial(len(choices))
            choice_rank = choice_rank_state % choice_space
            choice_rank_state //= choice_space
            shuffled_choices = unrank_permutation(choices, choice_rank)
            display_choices = [
                {"key": DISPLAY_KEYS[index], "text": choice["text"], "sourceKey": choice["key"]}
                for index, choice in enumerate(shuffled_choices)
            ]
            display_correct_answers = [
                choice["key"] for choice in display_choices if choice["sourceKey"] in source_correct_answers
            ]
        else:
            display_choices = [
                {"key": choice["key"], "text": choice["text"], "sourceKey": choice["key"]}
                for choice in question["choices"]
            ]
            display_correct_answers = list(source_correct_answers)

        rendered_choices_by_question[question["id"]] = (display_choices, display_correct_answers)

    ordered_questions = unrank_permutation(selected_questions, question_rank)
    rendered_questions: list[dict[str, Any]] = []
    for position, question in enumerate(ordered_questions, start=1):
        display_choices, display_correct_answers = rendered_choices_by_question[question["id"]]
        rendered_questions.append(
            {
                "position": position,
                "sourceQuestionId": question["id"],
                "question": question["question"],
                "difficulty": question["difficulty"],
                "points": question["points"],
                "chapters": extract_question_chapters(question),
                "learningObjectiveIds": list(question["learningObjectiveIds"]),
                "learningObjectives": [
                    {"id": objective_id, "label": objective_labels.get(objective_id, objective_id)}
                    for objective_id in question["learningObjectiveIds"]
                ],
                "shuffleChoices": bool(question["shuffleChoices"]),
                "bookLocations": question["bookLocations"],
                "displayChoices": display_choices,
                "displayCorrectAnswers": display_correct_answers,
                "sourceCorrectAnswers": list(question["correctAnswers"]),
                "explanation": question.get("explanation", ""),
            }
        )

    variant = {
        "variantId": str(uuid4()),
        "examSetId": exam_set_id,
        "questions": rendered_questions,
    }
    variant["signature"] = build_variant_signature(rendered_questions)
    return variant


def build_variant_printable_filename(position: int, total: int) -> str:
    width = max(2, len(str(max(total, 1))))
    return f"student-variant-{position:0{width}d}.pdf"


def annotate_variant_printables(variants: list[dict[str, Any]]) -> None:
    total = len(variants)
    for index, variant in enumerate(variants, start=1):
        variant["printableOrdinal"] = index
        variant["printableFileName"] = build_variant_printable_filename(index, total)


def build_variant_qr_payload(exam_set_id: str, variant_id: str) -> str:
    return json.dumps(
        {"examSetId": exam_set_id, "variantId": variant_id},
        ensure_ascii=False,
        separators=(",", ":"),
    )


def render_variant_qr_svg(exam_set_id: str, variant_id: str) -> str:
    qr_code = segno.make(build_variant_qr_payload(exam_set_id, variant_id))
    return qr_code.svg_inline(scale=5, border=2, omitsize=True)


def build_variant_qr_pdf_bytes(exam_set_id: str, variant_id: str) -> bytes:
    qr_code = segno.make(build_variant_qr_payload(exam_set_id, variant_id))
    buffer = io.BytesIO()
    qr_code.save(buffer, kind="pdf", scale=3, border=2)
    return buffer.getvalue()


def estimate_wrapped_line_count(text: str, chars_per_line: int) -> int:
    normalized = strip_math_markup(text)
    if not normalized:
        return 1
    return max(1, math.ceil(len(normalized) / chars_per_line))


def estimate_question_print_units(question: dict[str, Any]) -> int:
    units = 4 + estimate_wrapped_line_count(question.get("question", ""), 92)
    for choice in question.get("displayChoices", []):
        units += 1 + estimate_wrapped_line_count(choice.get("text", ""), 84)
    return units + 1


def build_variant_print_layout(variant: dict[str, Any]) -> dict[str, Any]:
    pages: list[dict[str, Any]] = [{"pageNumber": 1, "kind": "cover", "questionPositions": []}]
    current_positions: list[int] = []
    used_units = 0

    for question in variant.get("questions", []):
        question_units = estimate_question_print_units(question)
        if current_positions and used_units + question_units > QUESTION_PAGE_CAPACITY:
            pages.append(
                {
                    "pageNumber": len(pages) + 1,
                    "kind": "questions",
                    "questionPositions": current_positions,
                }
            )
            current_positions = [question["position"]]
            used_units = question_units
            continue

        current_positions.append(question["position"])
        used_units += question_units

    if current_positions:
        pages.append(
            {
                "pageNumber": len(pages) + 1,
                "kind": "questions",
                "questionPositions": current_positions,
            }
        )

    return {
        "questionCount": len(variant.get("questions", [])),
        "totalPages": len(pages),
        "pages": pages,
    }


def annotate_variant_print_layouts(variants: list[dict[str, Any]]) -> None:
    for variant in variants:
        variant["printLayout"] = build_variant_print_layout(variant)


def generate_exam_run(state: AppState, quiz: dict[str, Any], request: dict[str, Any]) -> dict[str, Any]:
    question_by_id, order_by_id, question_errors = build_question_index(quiz)
    if question_errors:
        raise ValueError(question_errors[0]["message"])

    ordered_questions = [
        question_by_id[question_id]
        for question_id, _ in sorted(order_by_id.items(), key=lambda item: item[1])
    ]

    filtered_questions = [
        question for question in ordered_questions if question_matches_filters(question, request)
    ]
    filtered_question_ids = [question["id"] for question in filtered_questions]
    filtered_question_id_set = set(filtered_question_ids)

    include_set = set(request["includeQuestionIds"])
    exclude_set = set(request["excludeQuestionIds"])

    if len(include_set) > request["questionCount"]:
        raise ValueError("Force-included questions exceed the selected question count")

    available_questions = [
        question
        for question in ordered_questions
        if question["id"] not in exclude_set
        and (question["id"] in include_set or question["id"] in filtered_question_id_set)
    ]

    if len(available_questions) < request["questionCount"]:
        raise ValueError(
            f"Only {len(available_questions)} questions are available after filters and overrides, "
            f"but {request['questionCount']} were requested"
        )

    forced_questions = [question for question in ordered_questions if question["id"] in include_set]
    remaining_candidates = [
        question for question in available_questions if question["id"] not in include_set
    ]
    remaining_slots = request["questionCount"] - len(forced_questions)

    rng = random.SystemRandom()
    sampled_questions = rng.sample(remaining_candidates, remaining_slots)
    selected_question_ids = {question["id"] for question in forced_questions + sampled_questions}
    selected_questions = [
        question for question in ordered_questions if question["id"] in selected_question_ids
    ]

    shuffleable_questions = [question for question in selected_questions if question["shuffleChoices"]]
    max_unique_variants = math.factorial(len(selected_questions)) * (
        math.factorial(4) ** len(shuffleable_questions)
    )
    if request["variantCount"] > max_unique_variants:
        raise ValueError(
            f"Requested {request['variantCount']} unique variants, but the selected exam set only supports "
            f"{max_unique_variants}"
        )

    objective_labels = {
        objective["id"]: objective["label"]
        for objective in quiz.get("learningObjectives", [])
        if isinstance(objective, dict) and isinstance(objective.get("id"), str)
    }

    exam_set_id = str(uuid4())
    generated_at = datetime.now(timezone.utc).isoformat()
    ranks = sample_unique_ranks(max_unique_variants, request["variantCount"], rng)
    variants = [
        build_variant(selected_questions, rank, exam_set_id, objective_labels) for rank in ranks
    ]
    annotate_variant_printables(variants)
    annotate_variant_print_layouts(variants)

    return {
        "examSetId": exam_set_id,
        "generatedAt": generated_at,
        "quiz": {
            "title": quiz.get("title", ""),
            "description": quiz.get("description", ""),
            "dbPath": str(state.db_path),
        },
        "printSettings": {
            "institutionName": request["institutionName"],
            "examName": request["examName"] or quiz.get("title", ""),
            "courseName": request["courseName"],
            "examDate": request["examDate"],
            "startTime": request["startTime"],
            "totalTimeMinutes": request["totalTimeMinutes"],
            "instructor": request["instructor"],
            "allowedMaterials": request["allowedMaterials"],
            "omrInstructions": request["omrInstructions"],
            "examRules": request["examRules"],
        },
        "printableFolderName": DEFAULT_PRINTABLE_FOLDER,
        "questionPoolFileName": QUESTION_POOL_PRINTABLE_NAME,
        "selection": {
            "questionCount": request["questionCount"],
            "variantCount": request["variantCount"],
            "chapters": request["chapters"],
            "difficulties": request["difficulties"],
            "learningObjectiveIds": request["learningObjectiveIds"],
            "includeQuestionIds": request["includeQuestionIds"],
            "excludeQuestionIds": request["excludeQuestionIds"],
            "filteredQuestionIds": filtered_question_ids,
            "availableQuestionIds": [question["id"] for question in available_questions],
            "selectedQuestionIds": [question["id"] for question in selected_questions],
            "maxUniqueVariants": str(max_unique_variants),
        },
        "questionPool": [
            build_question_pool_entry(question, objective_labels) for question in selected_questions
        ],
        "variants": variants,
    }


def get_question_pool_for_export(state: AppState, exam_set: dict[str, Any]) -> list[dict[str, Any]]:
    question_pool = exam_set.get("questionPool")
    if isinstance(question_pool, list) and question_pool:
        return question_pool

    quiz = load_json(state.db_path)
    question_by_id, _, _ = build_question_index(quiz)
    objective_labels = {
        objective["id"]: objective["label"]
        for objective in quiz.get("learningObjectives", [])
        if isinstance(objective, dict) and isinstance(objective.get("id"), str)
    }

    rebuilt_pool: list[dict[str, Any]] = []
    for question_id in exam_set.get("selection", {}).get("selectedQuestionIds", []):
        question = question_by_id.get(question_id)
        if question is None:
            continue
        rebuilt_pool.append(build_question_pool_entry(question, objective_labels))
    return rebuilt_pool


def variant_label_for_print(variant: dict[str, Any]) -> str:
    ordinal = int(variant.get("printableOrdinal") or 0)
    if ordinal > 0:
        width = max(2, len(str(ordinal)))
        return f"Variant {ordinal:0{width}d}"
    return str(variant.get("printableFileName") or variant.get("variantId") or "Variant")


def build_student_latex_document(exam_set: dict[str, Any], variant: dict[str, Any]) -> str:
    print_settings = get_print_settings(exam_set)
    total_points = sum(int(question.get("points") or 1) for question in variant.get("questions", []))
    distinct_points = sorted(
        {
            int(question.get("points") or 1)
            for question in variant.get("questions", [])
            if isinstance(question, dict)
        }
    )
    if len(distinct_points) == 1:
        points_per_question = str(distinct_points[0])
    else:
        points_per_question = "Varies"

    total_time = str(print_settings["totalTimeMinutes"]).strip()
    if total_time:
        total_time = f"{total_time} minutes"

    replacements = {
        "%%QUIZPOOL_INSTITUTION%%": latex_placeholder_value_or_blank(print_settings["institutionName"]),
        "%%QUIZPOOL_DEPARTMENT%%": "",
        "%%QUIZPOOL_COURSE%%": latex_placeholder_value_or_blank(print_settings["courseName"]),
        "%%QUIZPOOL_EXAM_NAME%%": latex_placeholder_value_or_blank(print_settings["examName"]),
        "%%QUIZPOOL_SEMESTER%%": "",
        "%%QUIZPOOL_EXAM_DATE%%": latex_placeholder_value_or_blank(print_settings["examDate"]),
        "%%QUIZPOOL_TOTAL_TIME%%": latex_placeholder_value_or_blank(total_time),
        "%%QUIZPOOL_QUESTION_COUNT%%": latex_escape_text_segment(
            str(len(variant.get("questions", []))),
            preserve_linebreaks=False,
        ),
        "%%QUIZPOOL_TOTAL_POINTS%%": latex_escape_text_segment(
            str(total_points),
            preserve_linebreaks=False,
        ),
        "%%QUIZPOOL_INSTRUCTOR%%": latex_placeholder_value_or_blank(
            print_settings["instructor"]
        ),
        "%%QUIZPOOL_ALLOWED_MATERIALS%%": latex_placeholder_value_or_blank(
            print_settings["allowedMaterials"]
        ),
        "%%QUIZPOOL_POINTS_PER_QUESTION%%": latex_escape_text_segment(
            points_per_question,
            preserve_linebreaks=False,
        ),
        "%%QUIZPOOL_RULES%%": render_latex_rules(print_settings["examRules"]),
        "%%QUIZPOOL_QUESTION_HEADING%%": latex_escape_text_segment(
            variant_question_heading(variant),
            preserve_linebreaks=False,
        ),
        "%%QUIZPOOL_PAGE_QR%%": rf"\includegraphics[width=0.42in]{{{LATEX_VARIANT_QR_ASSET_NAME}}}",
        "%%QUIZPOOL_QUESTION_BLOCKS%%": render_student_question_blocks_latex(variant),
    }
    return apply_latex_template(LATEX_STUDENT_TEMPLATE, replacements)


def build_student_latex_assets(exam_set: dict[str, Any], variant: dict[str, Any]) -> dict[str, bytes]:
    exam_set_id = str(exam_set.get("examSetId") or "").strip()
    variant_id = str(variant.get("variantId") or "").strip()
    if not exam_set_id or not variant_id:
        return {}
    return {
        LATEX_VARIANT_QR_ASSET_NAME: build_variant_qr_pdf_bytes(exam_set_id, variant_id),
    }


def build_question_pool_latex_document(
    exam_set: dict[str, Any],
    question_pool: list[dict[str, Any]],
) -> str:
    print_settings = get_print_settings(exam_set)
    replacements = {
        "%%QUIZPOOL_INSTITUTION%%": latex_placeholder_value(print_settings["institutionName"]),
        "%%QUIZPOOL_EXAM_NAME%%": latex_placeholder_value(print_settings["examName"]),
        "%%QUIZPOOL_COURSE%%": latex_placeholder_value(print_settings["courseName"]),
        "%%QUIZPOOL_EXAM_DATE%%": latex_placeholder_value(print_settings["examDate"]),
        "%%QUIZPOOL_INSTRUCTOR%%": latex_placeholder_value_or_blank(
            print_settings["instructor"]
        ),
        "%%QUIZPOOL_ALLOWED_MATERIALS%%": latex_placeholder_value_or_blank(
            print_settings["allowedMaterials"]
        ),
        "%%QUIZPOOL_EXAM_SET_ID%%": latex_placeholder_value(exam_set.get("examSetId")),
        "%%QUIZPOOL_SELECTED_COUNT%%": latex_escape_text_segment(
            str(len(question_pool)),
            preserve_linebreaks=False,
        ),
        "%%QUIZPOOL_QUESTION_BLOCKS%%": render_question_pool_blocks_latex(question_pool),
    }
    return apply_latex_template(LATEX_QUESTION_POOL_TEMPLATE, replacements)


def summarize_latex_failure(output: str) -> str:
    lines: list[str] = []
    for line in output.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("!") or "Error" in stripped or "Emergency stop" in stripped:
            lines.append(stripped)
    if lines:
        return "; ".join(dedupe_preserve_order(lines)[:8])
    condensed = " ".join(output.split())
    return condensed[-800:] if condensed else "No LaTeX engine output was captured."


def latex_engine_candidates() -> list[str]:
    return [engine for engine in LATEX_ENGINES if shutil.which(engine)]


def compile_latex_with_engine(
    tex_content: str,
    *,
    job_name: str,
    engine: str,
    assets: dict[str, bytes] | None = None,
) -> bytes:
    try:
        with tempfile.TemporaryDirectory(prefix="quiz-pool-latex-") as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            tex_path = temp_dir / f"{job_name}.tex"
            pdf_path = temp_dir / f"{job_name}.pdf"
            log_path = temp_dir / f"{job_name}.log"
            tex_path.write_text(tex_content, encoding="utf-8")
            for asset_name, asset_bytes in (assets or {}).items():
                (temp_dir / asset_name).write_bytes(asset_bytes)

            command = [
                engine,
                "-interaction=nonstopmode",
                "-halt-on-error",
                "-no-shell-escape",
                f"-jobname={job_name}",
                tex_path.name,
            ]
            for _ in range(2):
                completed = subprocess.run(
                    command,
                    cwd=temp_dir,
                    check=False,
                    capture_output=True,
                    text=True,
                )
                if completed.returncode != 0:
                    log_output = log_path.read_text(encoding="utf-8", errors="replace") if log_path.exists() else ""
                    detail = summarize_latex_failure(
                        "\n".join(part for part in [completed.stdout, completed.stderr, log_output] if part)
                    )
                    raise ValueError(f"LaTeX compilation failed with {engine}: {detail}")

            if not pdf_path.exists():
                raise ValueError(f"LaTeX compilation with {engine} finished without producing a PDF.")
            return pdf_path.read_bytes()
    except FileNotFoundError as error:
        raise OSError(f"{engine} was not found in PATH.") from error


def render_latex_to_pdf(
    tex_content: str,
    *,
    job_name: str,
    assets: dict[str, bytes] | None = None,
) -> bytes:
    engines = latex_engine_candidates()
    if not engines:
        raise OSError(
            "No LaTeX engine was found in PATH. Install TeX Live or another LaTeX distribution."
        )

    failures: list[str] = []
    for engine in engines:
        try:
            return compile_latex_with_engine(
                tex_content,
                job_name=job_name,
                engine=engine,
                assets=assets,
            )
        except ValueError as error:
            failures.append(str(error))
            continue

    detail = "; ".join(failures)
    raise ValueError(detail or "LaTeX compilation failed for all available engines.")


def render_printable_html(title: str, subtitle: str, body_html: str) -> str:
    escaped_title = html.escape(title)
    escaped_subtitle = html.escape(subtitle)
    return f"""<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{escaped_title}</title>
    <style>
      :root {{
        color-scheme: light;
        --ink: #182026;
        --muted: #5c6460;
        --border: #d7d7d0;
        --panel: #ffffff;
        --paper: #f8f5ec;
      }}

      * {{
        box-sizing: border-box;
      }}

      body {{
        background: var(--paper);
        color: var(--ink);
        font-family: "Avenir Next", "Trebuchet MS", sans-serif;
        margin: 0;
        padding: 24px;
      }}

      .page {{
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 18px;
        margin: 0 auto;
        max-width: 920px;
        padding: 28px;
        position: relative;
      }}

      .eyebrow {{
        color: #146c59;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.14em;
        margin: 0 0 10px;
        text-transform: uppercase;
      }}

      h1,
      h2,
      h3 {{
        font-family: "Rockwell", "Georgia", serif;
        margin: 0;
      }}

      h1 {{
        margin-bottom: 8px;
      }}

      .subtitle {{
        color: var(--muted);
        margin: 0 0 22px;
      }}

      .meta-grid {{
        display: grid;
        gap: 12px;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        margin-bottom: 22px;
      }}

      .meta-card {{
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 12px 14px;
      }}

      .meta-card strong {{
        display: block;
        font-size: 0.82rem;
        margin-bottom: 6px;
        text-transform: uppercase;
      }}

      .question {{
        border: 1px solid var(--border);
        border-radius: 16px;
        margin-top: 16px;
        padding: 16px;
      }}

      .question-head {{
        color: var(--muted);
        font-size: 0.9rem;
        margin-bottom: 10px;
      }}

      .question-title {{
        font-size: 1.02rem;
        line-height: 1.45;
        margin: 0 0 12px;
      }}

      .choice-list {{
        list-style: none;
        margin: 0;
        padding: 0;
      }}

      .choice-list li {{
        border-radius: 12px;
        margin-top: 8px;
        padding: 8px 10px;
      }}

      .choice-list li:nth-child(odd) {{
        background: rgba(20, 108, 89, 0.07);
      }}

{render_rich_text_css()}

      @media print {{
        body {{
          background: #fff;
          padding: 12mm;
        }}

        .page {{
          border: 0;
          border-radius: 0;
          box-shadow: none;
          max-width: none;
          padding: 0;
        }}
      }}
    </style>
  </head>
  <body>
    <article class="page">
      <p class="eyebrow">Quiz Pool</p>
      <h1>{escaped_title}</h1>
      <p class="subtitle">{escaped_subtitle}</p>
      {body_html}
    </article>
{render_inline_rich_text_module("""
      window.__quizPoolPrintableReady = false;

      function renderPrintableRichText() {
        renderRichTextTargets(document);
        window.__quizPoolPrintableReady = true;
      }

      if (document.readyState === "loading") {
        window.addEventListener("DOMContentLoaded", renderPrintableRichText, { once: true });
      } else {
        renderPrintableRichText();
      }

      window.addEventListener("load", renderPrintableRichText, { once: true });
""")}
  </body>
</html>
"""


def render_meta_cards(items: list[tuple[str, str]]) -> str:
    cards = []
    for label, value in items:
        normalized = value.strip()
        if not normalized:
            continue
        cards.append(
            f"""
  <div class="meta-card">
    <strong>{html.escape(label)}</strong>
    <span>{html.escape(normalized)}</span>
  </div>
"""
        )

    if not cards:
        return ""

    return "<section class=\"meta-grid\">" + "".join(cards) + "\n</section>\n"


def render_question_pool_html(exam_set: dict[str, Any], question_pool: list[dict[str, Any]]) -> str:
    print_settings = get_print_settings(exam_set)
    summary = render_meta_cards(
        [
            ("Exam Name", print_settings["examName"]),
            ("Course Name", print_settings["courseName"]),
            ("Exam Date", print_settings["examDate"]),
            ("Exam Set", exam_set["examSetId"]),
            ("Selected Questions", str(len(question_pool))),
        ]
    )

    question_sections: list[str] = []
    for index, question in enumerate(question_pool, start=1):
        choices = "".join(
            (
                "<li>"
                f"<span>{html.escape(choice['key'])}.</span> "
                f"<span data-rich-text>{html.escape(choice['text'])}</span>"
                "</li>"
            )
            for choice in question["choices"]
        )
        chapters = ", ".join(question["chapters"]) or "—"
        question_sections.append(
            f"""
<section class="question">
  <div class="question-head">Question {index} · {html.escape(question['sourceQuestionId'])} · Difficulty {question['difficulty']} · {question['points']} pt · Chapters {html.escape(chapters)}</div>
  <p class="question-title" data-rich-text>{html.escape(question['question'])}</p>
  <ul class="choice-list">{choices}</ul>
</section>
"""
        )

    body = summary + "".join(question_sections)
    return render_printable_html(
        title=f"{print_settings['examName']} Question Pool",
        subtitle="Shared source question set before variant-specific shuffling.",
        body_html=body,
    )


def render_variant_qr_markup(exam_set: dict[str, Any], variant: dict[str, Any]) -> str:
    qr_svg = render_variant_qr_svg(exam_set["examSetId"], variant["variantId"])
    return f'<div class="sheet-header__qr" aria-label="Variant tracking QR code">{qr_svg}</div>'


def render_exam_detail(label: str, value: str, *, escape_value: bool = True) -> str:
    escaped_label = html.escape(label)
    escaped_value = html.escape(value) if escape_value else value
    blank_class = " exam-detail__value--blank" if not value else ""
    body = escaped_value if value else "&nbsp;"
    return f"""
<div class="exam-detail">
  <span class="exam-detail__label">{escaped_label}</span>
  <span class="exam-detail__value{blank_class}">{body}</span>
</div>
"""


def render_student_line_field(label: str, *, wide: bool = False) -> str:
    wide_class = " student-line-field--wide" if wide else ""
    return f"""
<div class="student-line-field{wide_class}">
  <span class="student-line-field__label">{html.escape(label)}</span>
  <span class="student-line-field__line"></span>
</div>
"""


def render_student_print_css() -> str:
    return """      :root {
        color-scheme: light;
        --paper: #ffffff;
        --ink: #111111;
        --muted: #444444;
        --rule: #222222;
        --border: #1f1f1f;
      }

      * {
        box-sizing: border-box;
      }

      @page {
        size: A4 portrait;
        margin: 0;
      }

      html,
      body {
        background: #f0f0f0;
        color: var(--ink);
        font-family: "Helvetica Neue", Arial, sans-serif;
        margin: 0;
        padding: 0;
      }

      body {
        display: flex;
        flex-direction: column;
        gap: 10mm;
        padding: 8mm 0;
      }

      .sheet-page {
        background: var(--paper);
        box-shadow: 0 6mm 16mm rgba(0, 0, 0, 0.12);
        break-after: page;
        break-inside: avoid;
        display: flex;
        flex-direction: column;
        margin: 0 auto;
        page-break-after: always;
        page-break-inside: avoid;
        height: 297mm;
        padding: 16mm 17mm 15mm;
        width: 210mm;
      }

      .sheet-page:last-child {
        break-after: auto;
        page-break-after: auto;
      }

      .sheet-header {
        align-items: start;
        display: grid;
        gap: 5mm;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1.3fr) 27mm;
      }

      .sheet-header__institution,
      .sheet-header__title {
        min-height: 27mm;
      }

      .sheet-header__institution {
        font-size: 10.2pt;
        font-weight: 700;
        letter-spacing: 0.04em;
        line-height: 1.35;
        padding-top: 1.5mm;
        text-transform: uppercase;
      }

      .sheet-header__title {
        align-items: center;
        display: flex;
        font-size: 15pt;
        font-weight: 700;
        justify-content: center;
        letter-spacing: 0.03em;
        line-height: 1.2;
        padding: 0 4mm;
        text-align: center;
        text-transform: uppercase;
      }

      .sheet-header__qr {
        align-items: center;
        border: 0.3mm solid var(--rule);
        display: flex;
        height: 27mm;
        justify-content: center;
        justify-self: end;
        overflow: hidden;
        padding: 1.2mm;
        width: 27mm;
      }

      .sheet-header__qr svg {
        display: block;
        height: 100% !important;
        max-height: 100%;
        max-width: 100%;
        width: 100% !important;
      }

      .sheet-header-rule {
        border-bottom: 0.35mm solid var(--rule);
        margin: 3mm 0 5mm;
      }

      .sheet-body {
        display: flex;
        flex: 1;
        flex-direction: column;
        gap: 4.5mm;
      }

      .sheet-footer {
        border-top: 0.2mm solid #666666;
        color: var(--muted);
        font-size: 9pt;
        margin-top: 6mm;
        padding-top: 3mm;
        text-align: center;
      }

      .section-block {
        border: 0.3mm solid var(--border);
        padding: 4.5mm;
      }

      .section-block__title {
        font-size: 10.5pt;
        font-weight: 700;
        letter-spacing: 0.05em;
        margin: 0 0 3.5mm;
        text-transform: uppercase;
      }

      .student-line-grid {
        display: grid;
        gap: 4mm;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .student-line-field {
        display: flex;
        flex-direction: column;
        gap: 2.3mm;
      }

      .student-line-field--wide {
        grid-column: 1 / -1;
      }

      .student-line-field__label {
        font-size: 9pt;
        font-weight: 700;
        text-transform: uppercase;
      }

      .student-line-field__line {
        border-bottom: 0.3mm solid var(--rule);
        display: block;
        min-height: 7mm;
      }

      .exam-detail-grid {
        display: grid;
        gap: 3mm 5mm;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .exam-detail {
        border: 0.2mm solid #777777;
        min-height: 16mm;
        padding: 2.8mm 3.2mm;
      }

      .exam-detail__label {
        display: block;
        font-size: 8.6pt;
        font-weight: 700;
        margin-bottom: 1.6mm;
        text-transform: uppercase;
      }

      .exam-detail__value {
        border-bottom: 0.25mm solid var(--rule);
        display: block;
        font-size: 10.3pt;
        min-height: 5.6mm;
        padding-bottom: 0.7mm;
      }

      .exam-detail__value--blank {
        color: transparent;
      }

      .rules-list {
        font-size: 10.2pt;
        line-height: 1.45;
        margin: 0;
        padding-left: 5.4mm;
      }

      .rules-list li + li {
        margin-top: 1.5mm;
      }

      .instruction-line {
        border-top: 0.2mm solid #777777;
        font-size: 10.7pt;
        font-weight: 700;
        margin: 4mm 0 0;
        padding-top: 3.2mm;
      }

      .question-stack {
        display: flex;
        flex-direction: column;
        gap: 3.2mm;
      }

      .question-block {
        border: 0.2mm solid var(--border);
        break-inside: avoid;
        page-break-inside: avoid;
        padding: 2.8mm 3.2mm;
      }

      .question-block__prompt {
        font-size: 10.6pt;
        line-height: 1.38;
        margin: 0;
      }

      .question-block__number {
        font-size: 0.94em;
        font-weight: 700;
        margin-right: 1.6mm;
      }

      .question-block__text {
        font-size: 1em;
      }

      .choice-list {
        display: flex;
        flex-direction: column;
        gap: 1.2mm;
        list-style: none;
        margin: 2.4mm 0 0;
        padding: 0;
      }

      .choice-list li {
        align-items: start;
        display: grid;
        gap: 1.8mm;
        grid-template-columns: 5mm minmax(0, 1fr);
        padding: 0;
      }

      .choice-key {
        font-weight: 700;
      }

      .choice-text {
        line-height: 1.4;
      }

""" + render_rich_text_css() + """

      .sheet-page--question {
        padding: 8mm 10mm 8mm;
      }

      .sheet-page--question .sheet-header {
        gap: 2.2mm;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1.5fr) 16mm;
      }

      .sheet-page--question .sheet-header__institution,
      .sheet-page--question .sheet-header__title {
        min-height: 16mm;
      }

      .sheet-page--question .sheet-header__institution {
        font-size: 7.7pt;
        line-height: 1.16;
        padding-top: 0.5mm;
      }

      .sheet-page--question .sheet-header__title {
        font-size: 10.5pt;
        letter-spacing: 0.02em;
        line-height: 1.08;
        padding: 0 1.2mm;
      }

      .sheet-page--question .sheet-header__qr {
        height: 16mm;
        padding: 0.6mm;
        width: 16mm;
      }

      .sheet-page--question .sheet-header-rule {
        margin: 1.5mm 0 2mm;
      }

      .sheet-page--question .sheet-footer {
        font-size: 7.2pt;
        margin-top: 2mm;
        padding-top: 1.2mm;
      }

      .sheet-page--question .question-stack {
        gap: 1.6mm;
      }

      .sheet-page--question .question-block {
        padding: 1.7mm 2mm;
      }

      .sheet-page--question .question-block__prompt {
        font-size: 9.35pt;
        line-height: 1.2;
      }

      .sheet-page--question .question-block__number {
        margin-right: 1mm;
      }

      .sheet-page--question .choice-list {
        gap: 0.8mm;
        margin-top: 1.5mm;
      }

      .sheet-page--question .choice-list li {
        gap: 1.4mm;
        grid-template-columns: 4.6mm minmax(0, 1fr);
      }

      .sheet-page--question .choice-key,
      .sheet-page--question .choice-text {
        font-size: 8.75pt;
        line-height: 1.14;
      }

      .pagination-measure {
        left: -1000vw;
        pointer-events: none;
        position: absolute;
        top: 0;
        visibility: hidden;
      }

      .hidden-print-buffer {
        display: none;
      }

      .sheet-page--omr {
        padding: 0;
        position: relative;
      }

      .sheet-page--omr .sheet-footer {
        background: rgba(255, 255, 255, 0.92);
        bottom: 0;
        left: 0;
        margin: 0;
        padding: 2mm 0 3mm;
        position: absolute;
        right: 0;
      }

      .omr-sheet-image {
        display: block;
        height: 100%;
        object-fit: fill;
        width: 100%;
      }

      @media print {
        html,
        body {
          background: #fff;
          padding: 0;
        }

        body {
          display: block;
        }

        .sheet-page {
          box-shadow: none;
          margin: 0;
        }

        .sheet-page--omr .sheet-footer {
          background: rgba(255, 255, 255, 0.96);
        }
      }
"""


def render_student_cover_page(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
    page_number: int,
) -> str:
    print_settings = get_print_settings(exam_set)
    rules_markup = "".join(
        f"<li>{html.escape(rule)}</li>" for rule in print_settings["examRules"]
    )
    student_info = "".join(
        [
            render_student_line_field("Name", wide=True),
            render_student_line_field("ID"),
            render_student_line_field("Class / Section"),
            render_student_line_field("Signature", wide=True),
        ]
    )
    exam_details = "".join(
        [
            render_exam_detail("Exam Name", print_settings["examName"]),
            render_exam_detail("Course / Subject", print_settings["courseName"]),
            render_exam_detail("Exam Date", print_settings["examDate"]),
            render_exam_detail("Start Time", print_settings["startTime"]),
            render_exam_detail("Total Time in Minutes", print_settings["totalTimeMinutes"]),
            render_exam_detail("Number of Questions", str(len(variant.get("questions", [])))),
            render_exam_detail("Total Points", str(sum(int(question.get("points") or 1) for question in variant.get("questions", [])))),
            render_exam_detail(
                "Number of Pages",
                '<span class="js-total-pages">1</span>',
                escape_value=False,
            ),
        ]
    )

    return f"""
<section class="sheet-page">
  <header class="sheet-header">
    <div class="sheet-header__institution">{html.escape(print_settings["institutionName"])}</div>
    <div class="sheet-header__title">{html.escape(print_settings["examName"])}</div>
    {render_variant_qr_markup(exam_set, variant)}
  </header>
  <div class="sheet-header-rule"></div>
  <main class="sheet-body">
    <section class="section-block">
      <h2 class="section-block__title">Student Information</h2>
      <div class="student-line-grid">{student_info}</div>
    </section>
    <section class="section-block">
      <h2 class="section-block__title">Exam Information</h2>
      <div class="exam-detail-grid">{exam_details}</div>
    </section>
    <section class="section-block">
      <h2 class="section-block__title">Exam Rules</h2>
      <ol class="rules-list">{rules_markup}</ol>
    </section>
  </main>
  <footer class="sheet-footer">Page <span class="js-page-number">{page_number}</span> of <span class="js-total-pages">1</span></footer>
</section>
"""


def render_student_question_blocks(variant: dict[str, Any]) -> str:
    question_by_position = {
        question["position"]: question for question in variant.get("questions", [])
    }
    question_blocks: list[str] = []

    for position in sorted(question_by_position):
        question = question_by_position[position]
        choices = "".join(
            f"""
<li>
  <span class="choice-key">{html.escape(choice['key'])}.</span>
  <span class="choice-text" data-rich-text>{html.escape(choice['text'])}</span>
</li>
"""
            for choice in question["displayChoices"]
        )
        question_blocks.append(
            f"""
<article class="question-block" data-question-position="{question['position']}">
  <p class="question-block__prompt">
    <span class="question-block__number">{question['position']}.</span>
    <span class="question-block__text" data-rich-text>{html.escape(question['question'])}</span>
  </p>
  <ul class="choice-list">{choices}</ul>
</article>
"""
        )

    return "".join(question_blocks)


def render_student_question_page_template(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
) -> str:
    print_settings = get_print_settings(exam_set)
    return f"""
<section class="sheet-page sheet-page--question">
  <header class="sheet-header">
    <div class="sheet-header__institution">{html.escape(print_settings["institutionName"])}</div>
    <div class="sheet-header__title">{html.escape(print_settings["examName"])}</div>
    {render_variant_qr_markup(exam_set, variant)}
  </header>
  <div class="sheet-header-rule"></div>
  <main class="sheet-body">
    <div class="question-stack"></div>
  </main>
  <footer class="sheet-footer">Page <span class="js-page-number">2</span> of <span class="js-total-pages">1</span></footer>
</section>
"""


def build_omr_sheet_config(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
) -> SheetConfig:
    print_settings = get_print_settings(exam_set)
    questions = [question for question in variant.get("questions", []) if isinstance(question, dict)]
    choice_count = max(
        (len(question.get("displayChoices", [])) for question in questions),
        default=0,
    )
    return SheetConfig(
        question_count=len(questions),
        choice_count=choice_count,
        exam_set_id=str(exam_set["examSetId"]),
        variant_id=str(variant["variantId"]),
        title=print_settings["examName"] or "Optical Mark Recognition Sheet",
        instructions=print_settings["omrInstructions"],
    )


def build_omr_sheet_pdf_bytes(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
) -> bytes:
    config = build_omr_sheet_config(exam_set, variant)
    output = io.BytesIO()
    generate_omr_sheet(config, output)
    return output.getvalue()


def build_omr_sheet_page_images(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
) -> list[str]:
    with tempfile.TemporaryDirectory() as temp_dir_raw:
        temp_dir = Path(temp_dir_raw)
        pdf_path = temp_dir / "omr-sheet.pdf"
        image_prefix = temp_dir / "omr-page"
        pdf_path.write_bytes(build_omr_sheet_pdf_bytes(exam_set, variant))
        subprocess.run(
            [
                "pdftoppm",
                "-png",
                "-r",
                "300",
                str(pdf_path),
                str(image_prefix),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        image_paths = sorted(temp_dir.glob("omr-page-*.png"))
        return [
            "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")
            for path in image_paths
        ]


def render_student_omr_pages(exam_set: dict[str, Any], variant: dict[str, Any]) -> str:
    print_settings = get_print_settings(exam_set)
    omr_images = build_omr_sheet_page_images(exam_set, variant)
    pages: list[str] = []
    for index, image_data_uri in enumerate(omr_images):
        pages.append(
            f"""
<section class="sheet-page sheet-page--omr">
  <img class="omr-sheet-image" src="{image_data_uri}" alt="OMR answer sheet page {index + 1}" />
</section>
"""
        )
    return "".join(pages)


def render_student_pagination_script() -> str:
    return """
    <script type="module">
""" + load_web_text_asset("rich-text.js") + """
      (() => {
        window.__quizPoolPrintableReady = false;

        function renderPrintableRichText() {
          const bank = document.querySelector('#question-bank');
          if (bank instanceof HTMLTemplateElement) {
            renderRichTextTargets(bank.content);
          }
          renderRichTextTargets(document);
        }

        function createQuestionPage() {
          const template = document.querySelector('#question-page-template');
          if (!(template instanceof HTMLTemplateElement)) return null;
          const fragment = template.content.cloneNode(true);
          return fragment.firstElementChild;
        }

        function updatePageCounters(totalPages) {
          document.querySelectorAll('.js-total-pages').forEach((node) => {
            node.textContent = String(totalPages);
          });
        }

        function paginateStudentView() {
          const root = document.querySelector('#question-pages');
          const measure = document.querySelector('#pagination-measure');
          const bank = document.querySelector('#question-bank');
          if (!(root instanceof HTMLElement) || !(measure instanceof HTMLElement) || !(bank instanceof HTMLTemplateElement)) {
            return;
          }

          root.innerHTML = '';
          measure.innerHTML = '';

          const sourceBlocks = Array.from(bank.content.querySelectorAll('.question-block'));
          const finalPages = [];
          let currentPage = createQuestionPage();
          if (!(currentPage instanceof HTMLElement)) return;
          measure.appendChild(currentPage);
          let currentStack = currentPage.querySelector('.question-stack');
          if (!(currentStack instanceof HTMLElement)) return;

          for (const sourceBlock of sourceBlocks) {
            const block = sourceBlock.cloneNode(true);
            currentStack.appendChild(block);
            const overflows = currentPage.scrollHeight > currentPage.clientHeight;
            if (!overflows) continue;

            currentStack.removeChild(block);
            if (currentStack.children.length === 0) {
              currentStack.appendChild(block);
              finalPages.push(currentPage);
              currentPage = createQuestionPage();
              if (!(currentPage instanceof HTMLElement)) return;
              measure.appendChild(currentPage);
              currentStack = currentPage.querySelector('.question-stack');
              if (!(currentStack instanceof HTMLElement)) return;
              continue;
            }

            finalPages.push(currentPage);
            currentPage = createQuestionPage();
            if (!(currentPage instanceof HTMLElement)) return;
            measure.appendChild(currentPage);
            currentStack = currentPage.querySelector('.question-stack');
            if (!(currentStack instanceof HTMLElement)) return;
            currentStack.appendChild(block);
          }

          if (currentStack.children.length > 0) {
            finalPages.push(currentPage);
          } else {
            currentPage.remove();
          }

          finalPages.forEach((page, index) => {
            const pageNumber = index + 2;
            const pageNumberNode = page.querySelector('.js-page-number');
            if (pageNumberNode) pageNumberNode.textContent = String(pageNumber);
            root.appendChild(page);
          });

          const omrPages = Array.from(document.querySelectorAll('#omr-pages .sheet-page--omr'));
          omrPages.forEach((page) => {
            root.appendChild(page);
          });

          updatePageCounters(finalPages.length + 1);
          window.__quizPoolPrintableReady = true;
        }

        let scheduled = false;
        function schedulePagination() {
          if (scheduled) return;
          window.__quizPoolPrintableReady = false;
          scheduled = true;
          requestAnimationFrame(() => {
            scheduled = false;
            renderPrintableRichText();
            paginateStudentView();
          });
        }

        if (document.readyState === 'loading') {
          window.addEventListener('DOMContentLoaded', schedulePagination, { once: true });
        } else {
          schedulePagination();
        }
        window.addEventListener('load', schedulePagination, { once: true });
        window.addEventListener('resize', schedulePagination);
        window.addEventListener('beforeprint', () => {
          renderPrintableRichText();
          paginateStudentView();
        });
      })();
    </script>
"""


def render_variant_html(
    exam_set: dict[str, Any],
    variant: dict[str, Any],
    *,
    include_omr_pages: bool = True,
) -> str:
    print_settings = get_print_settings(exam_set)
    omr_pages = render_student_omr_pages(exam_set, variant) if include_omr_pages else ""
    return f"""<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{html.escape(print_settings["examName"])}</title>
    <style>
{render_student_print_css()}
    </style>
  </head>
  <body>
    {render_student_cover_page(exam_set, variant, 1)}
    <div id="question-pages"></div>
    <template id="question-bank">
      {render_student_question_blocks(variant)}
    </template>
    <template id="question-page-template">
      {render_student_question_page_template(exam_set, variant)}
    </template>
    <div id="omr-pages" class="hidden-print-buffer">
      {omr_pages}
    </div>
    <div id="pagination-measure" class="pagination-measure" aria-hidden="true"></div>
{render_student_pagination_script()}
  </body>
</html>
"""


async def render_html_to_pdf_bytes(html_content: str) -> bytes:
    browser = await launch(
        headless=True,
        handleSIGINT=False,
        handleSIGTERM=False,
        handleSIGHUP=False,
        args=[
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
        ],
    )
    try:
        page = await browser.newPage()
        await page.emulateMedia("print")
        await page.setContent(html_content)
        if "__quizPoolPrintableReady" in html_content:
            await page.waitForFunction("window.__quizPoolPrintableReady === true")
        return await page.pdf(
            {
                "printBackground": True,
                "preferCSSPageSize": True,
            }
        )
    finally:
        await browser.close()


def render_html_to_pdf(html_content: str) -> bytes:
    return asyncio.run(render_html_to_pdf_bytes(html_content))


def append_pdf_documents(base_pdf: bytes, extra_pdf: bytes) -> bytes:
    writer = PdfWriter()
    for source in (base_pdf, extra_pdf):
        reader = PdfReader(io.BytesIO(source))
        for page in reader.pages:
            writer.add_page(page)
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


def build_printable_zip(state: AppState, exam_set: dict[str, Any]) -> bytes:
    question_pool = get_question_pool_for_export(state, exam_set)
    buffer = io.BytesIO()
    variants = [variant for variant in exam_set.get("variants", []) if isinstance(variant, dict)]
    annotate_variant_printables(variants)
    annotate_variant_print_layouts(variants)
    base_folder = str(exam_set.get("printableFolderName") or DEFAULT_PRINTABLE_FOLDER)
    question_pool_name = str(exam_set.get("questionPoolFileName") or QUESTION_POOL_PRINTABLE_NAME)

    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        question_pool_pdf = render_latex_to_pdf(
            build_question_pool_latex_document(exam_set, question_pool),
            job_name="question-pool",
        )
        archive.writestr(
            f"{base_folder}/{question_pool_name}",
            question_pool_pdf,
        )
        for variant in variants:
            variant_file_name = str(
                variant.get("printableFileName")
                or build_variant_printable_filename(variant.get("printableOrdinal", 1), len(variants))
            )
            student_printable_pdf = render_latex_to_pdf(
                build_student_latex_document(exam_set, variant),
                job_name=f"student-variant-{int(variant.get('printableOrdinal') or 1):02d}",
                assets=build_student_latex_assets(exam_set, variant),
            )
            omr_pdf = build_omr_sheet_pdf_bytes(exam_set, variant)
            variant_pdf = append_pdf_documents(student_printable_pdf, omr_pdf)
            archive.writestr(
                f"{base_folder}/{variant_file_name}",
                variant_pdf,
            )

    return buffer.getvalue()


def build_handler(state: AppState) -> type[BaseHTTPRequestHandler]:
    class QuizRequestHandler(BaseHTTPRequestHandler):
        server_version = "QuizPool/0.2"

        def do_GET(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path == "/api/session-paths":
                self.handle_get_session_paths()
                return
            if parsed.path == "/api/quiz":
                self.handle_get_quiz()
                return
            if parsed.path == "/api/exams":
                self.handle_list_exam_sets()
                return
            if parsed.path.startswith("/api/exams/set/"):
                self.handle_get_exam_set(parsed.path)
                return
            if parsed.path.startswith("/api/exams/variant-qr/"):
                self.handle_get_variant_qr(parsed.path)
                return
            if parsed.path.startswith("/api/exams/variant/"):
                self.handle_get_variant(parsed.path)
                return
            if parsed.path.startswith("/api/exams/export/"):
                self.handle_export_exam_set(parsed.path)
                return
            self.serve_static(parsed.path)

        def do_PUT(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path != "/api/quiz":
                self.send_error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")
                return
            self.handle_put_quiz()

        def do_POST(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path == "/api/session-paths":
                self.handle_update_session_paths()
                return
            if parsed.path == "/api/quiz":
                self.handle_put_quiz()
                return
            if parsed.path == "/api/exams/generate":
                self.handle_generate_exams()
                return
            if parsed.path == "/api/exams/grade":
                self.handle_grade_exams()
                return
            if parsed.path == "/api/exams/annotate":
                self.handle_annotate_exams()
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")

        def do_DELETE(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path.startswith("/api/exams/set/"):
                self.handle_delete_exam_set(parsed.path)
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")

        def log_message(self, format: str, *args: object) -> None:
            return

        def read_json_body(self) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length <= 0:
                return None, [{"path": "<body>", "message": "Empty request body"}]

            raw_body = self.rfile.read(content_length)
            try:
                payload = json.loads(raw_body.decode("utf-8"))
            except json.JSONDecodeError as error:
                return None, [{"path": "<body>", "message": f"Invalid JSON: {error.msg}"}]

            if not isinstance(payload, dict):
                return None, [{"path": "<body>", "message": "Top-level payload must be a JSON object"}]

            return payload, []

        def handle_get_quiz(self) -> None:
            quiz = load_json(state.db_path)
            self.send_json(
                {
                    "quiz": quiz,
                    "dbPath": str(state.db_path),
                    "schemaPath": str(state.schema_path),
                    "examStorePath": str(state.exam_store_path),
                    "defaultDbPath": display_default_db_path(),
                    "defaultExamStorePath": display_default_exam_store_path(),
                }
            )

        def handle_get_session_paths(self) -> None:
            self.send_json(
                {
                    "dbPath": str(state.db_path),
                    "schemaPath": str(state.schema_path),
                    "examStorePath": str(state.exam_store_path),
                    "defaultDbPath": display_default_db_path(),
                    "defaultExamStorePath": display_default_exam_store_path(),
                }
            )

        def handle_update_session_paths(self) -> None:
            payload, errors = self.read_json_body()
            if errors:
                self.send_json({"ok": False, "errors": errors}, status=HTTPStatus.BAD_REQUEST)
                return

            raw_db_path = payload.get("dbPath")
            raw_exam_store_path = payload.get("examStorePath")
            request_errors: list[dict[str, str]] = []
            if not isinstance(raw_db_path, str) or not raw_db_path.strip():
                request_errors.append({"path": "dbPath", "message": "Quiz DB path is required"})
            if not isinstance(raw_exam_store_path, str) or not raw_exam_store_path.strip():
                request_errors.append({"path": "examStorePath", "message": "Exam DB path is required"})
            if request_errors:
                self.send_json({"ok": False, "errors": request_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            db_path = Path(raw_db_path).expanduser().resolve()
            exam_store_path = Path(raw_exam_store_path).expanduser().resolve()
            try:
                set_active_paths(state, db_path=db_path, exam_store_path=exam_store_path)
            except ValueError as error:
                self.send_json(
                    {"ok": False, "errors": [{"path": "<paths>", "message": str(error)}]},
                    status=HTTPStatus.BAD_REQUEST,
                )
                return

            self.send_json(
                {
                    "ok": True,
                    "dbPath": str(state.db_path),
                    "schemaPath": str(state.schema_path),
                    "examStorePath": str(state.exam_store_path),
                    "defaultDbPath": display_default_db_path(),
                    "defaultExamStorePath": display_default_exam_store_path(),
                }
            )

        def handle_list_exam_sets(self) -> None:
            try:
                summaries = [build_exam_set_summary(item) for item in list_exam_sets(state.exam_store_path)]
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            self.send_json(
                {
                    "examStorePath": str(state.exam_store_path),
                    "examSets": summaries,
                }
            )

        def handle_get_exam_set(self, path: str) -> None:
            exam_set_id = unquote(path.removeprefix("/api/exams/set/")).strip()
            if not exam_set_id:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            try:
                exam_set = find_exam_set(state.exam_store_path, exam_set_id)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            if exam_set is None:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            variants = [variant for variant in exam_set.get("variants", []) if isinstance(variant, dict)]
            annotate_variant_printables(variants)
            annotate_variant_print_layouts(variants)
            question_pool = exam_set.get("questionPool")
            if not isinstance(question_pool, list):
                question_pool = get_question_pool_for_export(state, exam_set)

            self.send_json(
                {
                    "examStorePath": str(state.exam_store_path),
                    "summary": build_exam_set_summary(exam_set),
                    "examSet": {
                        "examSetId": exam_set["examSetId"],
                        "generatedAt": exam_set["generatedAt"],
                        "quiz": exam_set["quiz"],
                        "printSettings": get_print_settings(exam_set),
                        "selection": exam_set["selection"],
                        "questionPool": question_pool,
                        "variants": variants,
                    },
                }
            )

        def handle_delete_exam_set(self, path: str) -> None:
            exam_set_id = unquote(path.removeprefix("/api/exams/set/")).strip()
            if not exam_set_id:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            try:
                deleted = delete_exam_set(state.exam_store_path, exam_set_id)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            if not deleted:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            self.send_json({"ok": True, "examSetId": exam_set_id})

        def handle_get_variant(self, path: str) -> None:
            variant_id = unquote(path.removeprefix("/api/exams/variant/")).strip()
            if not variant_id:
                self.send_error(HTTPStatus.NOT_FOUND, "Variant not found")
                return

            try:
                record = find_variant(state.exam_store_path, variant_id)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            if record is None:
                self.send_error(HTTPStatus.NOT_FOUND, "Variant not found")
                return

            exam_set, variant = record
            annotate_variant_printables([variant])
            annotate_variant_print_layouts([variant])
            self.send_json(
                {
                    "examStorePath": str(state.exam_store_path),
                    "examSetId": exam_set["examSetId"],
                    "generatedAt": exam_set["generatedAt"],
                    "quiz": exam_set["quiz"],
                    "printSettings": get_print_settings(exam_set),
                    "selection": exam_set["selection"],
                    "variant": variant,
                }
            )

        def handle_get_variant_qr(self, path: str) -> None:
            variant_id = unquote(path.removeprefix("/api/exams/variant-qr/")).removesuffix(".svg").strip()
            if not variant_id:
                self.send_error(HTTPStatus.NOT_FOUND, "Variant not found")
                return

            try:
                record = find_variant(state.exam_store_path, variant_id)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            if record is None:
                self.send_error(HTTPStatus.NOT_FOUND, "Variant not found")
                return

            exam_set, variant = record
            self.send_text(
                render_variant_qr_svg(exam_set["examSetId"], variant["variantId"]),
                content_type="image/svg+xml; charset=utf-8",
            )

        def handle_export_exam_set(self, path: str) -> None:
            exam_set_id = unquote(path.removeprefix("/api/exams/export/")).removesuffix(".zip").strip()
            if not exam_set_id:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            try:
                exam_set = find_exam_set(state.exam_store_path, exam_set_id)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            if exam_set is None:
                self.send_error(HTTPStatus.NOT_FOUND, "Exam set not found")
                return

            try:
                payload = build_printable_zip(state, exam_set)
            except (OSError, ValueError) as error:
                self.send_json(
                    {"errors": [{"path": "<export>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            self.send_bytes(
                payload,
                content_type="application/zip",
                filename=f"exam-set-{exam_set_id}-printables.zip",
            )

        def handle_put_quiz(self) -> None:
            payload, errors = self.read_json_body()
            if errors:
                self.send_json({"ok": False, "errors": errors}, status=HTTPStatus.BAD_REQUEST)
                return

            validation = validation_errors(state.validator, payload)
            if validation:
                self.send_json({"ok": False, "errors": validation}, status=HTTPStatus.BAD_REQUEST)
                return

            write_json_atomic(state.db_path, payload)
            self.send_json({"ok": True})

        def handle_generate_exams(self) -> None:
            payload, body_errors = self.read_json_body()
            if body_errors:
                self.send_json({"errors": body_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            quiz = load_json(state.db_path)
            quiz_errors = validation_errors(state.validator, quiz)
            if quiz_errors:
                self.send_json(
                    {
                        "errors": [
                            {
                                "path": "<quiz>",
                                "message": "The quiz database on disk is invalid. Fix it before generating exams.",
                            },
                            *quiz_errors,
                        ]
                    },
                    status=HTTPStatus.BAD_REQUEST,
                )
                return

            request, request_errors = normalize_generation_request(payload, quiz)
            if request_errors:
                self.send_json({"errors": request_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            try:
                exam_run = generate_exam_run(state, quiz, request)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<generation>", "message": str(error)}]},
                    status=HTTPStatus.BAD_REQUEST,
                )
                return

            try:
                append_exam_set(state.exam_store_path, exam_run)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return
            except OSError as error:
                self.send_json(
                    {"errors": [{"path": "<store>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            exam_run["examStorePath"] = str(state.exam_store_path)
            self.send_json(exam_run)

        def handle_grade_exams(self) -> None:
            payload, body_errors = self.read_json_body()
            if body_errors:
                self.send_json({"errors": body_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            request, request_errors = normalize_grading_request(payload)
            if request_errors:
                self.send_json({"errors": request_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            input_path = Path(request["inputPath"]).expanduser().resolve()
            try:
                result = grade_exam_pdfs(state, input_path)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<grading>", "message": str(error)}]},
                    status=HTTPStatus.BAD_REQUEST,
                )
                return
            except OSError as error:
                self.send_json(
                    {"errors": [{"path": "<grading>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            self.send_json(result)

        def handle_annotate_exams(self) -> None:
            payload, body_errors = self.read_json_body()
            if body_errors:
                self.send_json({"errors": body_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            request, request_errors = normalize_annotation_request(payload)
            if request_errors:
                self.send_json({"errors": request_errors}, status=HTTPStatus.BAD_REQUEST)
                return

            input_path = Path(request["inputPath"]).expanduser().resolve()
            output_path = Path(request["outputPath"]).expanduser().resolve()
            try:
                result = annotate_exam_pdfs(state, input_path, output_path)
            except ValueError as error:
                self.send_json(
                    {"errors": [{"path": "<annotation>", "message": str(error)}]},
                    status=HTTPStatus.BAD_REQUEST,
                )
                return
            except OSError as error:
                self.send_json(
                    {"errors": [{"path": "<annotation>", "message": str(error)}]},
                    status=HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return

            self.send_json(result)

        def serve_static(self, raw_path: str) -> None:
            request_path = raw_path or "/"
            if request_path == "/":
                relative = "welcome.html"
            elif request_path == "/welcome":
                relative = "welcome.html"
            elif request_path == "/editor":
                relative = "index.html"
            elif request_path == "/generator":
                relative = "generator.html"
            elif request_path == "/viewer":
                relative = "viewer.html"
            elif request_path == "/grading":
                relative = "grading.html"
            else:
                relative = request_path.lstrip("/")
            target = (WEB_ROOT / relative).resolve()

            try:
                target.relative_to(WEB_ROOT.resolve())
            except ValueError:
                self.send_error(HTTPStatus.FORBIDDEN, "Invalid path")
                return

            if not target.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "Not found")
                return

            mime_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
            with target.open("rb") as handle:
                payload = handle.read()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", mime_type)
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
            body = json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def send_bytes(
            self,
            payload: bytes,
            *,
            content_type: str,
            filename: str,
            status: HTTPStatus = HTTPStatus.OK,
        ) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def send_text(
            self,
            payload: str,
            *,
            content_type: str,
            status: HTTPStatus = HTTPStatus.OK,
        ) -> None:
            body = payload.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    return QuizRequestHandler


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Quiz pool browser editor")
    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB,
        help="Path to the quiz JSON file to edit",
    )
    parser.add_argument(
        "--schema",
        type=Path,
        default=DEFAULT_SCHEMA,
        help="Path to the JSON schema used for validation",
    )
    parser.add_argument(
        "--exam-store",
        type=Path,
        help="Path to the generated exam store JSON file",
    )
    parser.add_argument("--host", default="127.0.0.1", help="Host interface to bind")
    parser.add_argument("--port", type=int, default=8000, help="Port to listen on")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    db_path = args.db.resolve()
    schema_path = args.schema.resolve()
    exam_store_path = (
        args.exam_store.resolve()
        if args.exam_store is not None
        else default_exam_store_path_for(db_path)
    )

    if not schema_path.is_file():
        raise SystemExit(f"Schema file not found: {schema_path}")
    if not WEB_ROOT.is_dir():
        raise SystemExit(f"Web assets not found: {WEB_ROOT}")

    schema = load_json(schema_path)
    validator = Draft202012Validator(schema)
    try:
        validate_quiz_file(db_path, validator)
    except ValueError as error:
        raise SystemExit(str(error)) from error

    try:
        validate_exam_store_file(exam_store_path)
    except ValueError as error:
        raise SystemExit(f"Exam store is invalid: {error}") from error

    app_state = AppState(
        db_path=db_path,
        schema_path=schema_path,
        exam_store_path=exam_store_path,
        validator=validator,
    )
    handler = build_handler(app_state)
    server = ThreadingHTTPServer((args.host, args.port), handler)

    print(f"Quiz editor running at http://{args.host}:{args.port}")
    print(f"Editing database: {db_path}")
    print(f"Exam store: {exam_store_path}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
