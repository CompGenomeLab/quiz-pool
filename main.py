from __future__ import annotations

import argparse
import html
import io
import json
import math
import mimetypes
import random
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from uuid import uuid4
import zipfile

from jsonschema import Draft202012Validator
import segno


ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"
DEFAULT_DB = ROOT / "sample_quiz.json"
DEFAULT_SCHEMA = ROOT / "scheme.json"
DEFAULT_EXAM_STORE_NAME = "generated_exams.json"
DISPLAY_KEYS = ("A", "B", "C", "D")
DEFAULT_PRINTABLE_FOLDER = "exam-printables"
QUESTION_POOL_PRINTABLE_NAME = "question-pool.html"


@dataclass
class AppState:
    db_path: Path
    schema_path: Path
    exam_store_path: Path
    validator: Draft202012Validator


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

    payload = load_json(path)
    if not isinstance(payload, dict):
        raise ValueError("Exam store must be a JSON object")

    exam_sets = payload.get("examSets")
    if not isinstance(exam_sets, list):
        raise ValueError("Exam store must contain an 'examSets' array")

    return payload


def append_exam_set(path: Path, exam_set: dict[str, Any]) -> None:
    store = load_exam_store(path)
    store["examSets"].append(exam_set)
    write_json_atomic(path, store)


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


def dedupe_preserve_order(items: list[Any]) -> list[Any]:
    seen: set[Any] = set()
    deduped: list[Any] = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        deduped.append(item)
    return deduped


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
    exam_name = normalize_optional_string(payload, "examName", errors)
    course_name = normalize_optional_string(payload, "courseName", errors)
    exam_date = normalize_optional_string(payload, "examDate", errors)

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
            "examName": exam_name,
            "courseName": course_name,
            "examDate": exam_date,
        },
        [],
    )


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


def get_print_settings(exam_set: dict[str, Any]) -> dict[str, str]:
    raw = exam_set.get("printSettings")
    if not isinstance(raw, dict):
        raw = {}

    exam_name = raw.get("examName")
    course_name = raw.get("courseName")
    exam_date = raw.get("examDate")

    normalized_exam_name = exam_name.strip() if isinstance(exam_name, str) else ""
    normalized_course_name = course_name.strip() if isinstance(course_name, str) else ""
    normalized_exam_date = exam_date.strip() if isinstance(exam_date, str) else ""

    fallback_title = exam_set.get("quiz", {}).get("title", "")
    if not normalized_exam_name and isinstance(fallback_title, str):
        normalized_exam_name = fallback_title.strip()

    return {
        "examName": normalized_exam_name,
        "courseName": normalized_course_name,
        "examDate": normalized_exam_date,
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
    return f"student-variant-{position:0{width}d}.html"


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
    return qr_code.svg_inline(scale=4, border=1)


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

    return {
        "examSetId": exam_set_id,
        "generatedAt": generated_at,
        "quiz": {
            "title": quiz.get("title", ""),
            "description": quiz.get("description", ""),
            "dbPath": str(state.db_path),
        },
        "printSettings": {
            "examName": request["examName"] or quiz.get("title", ""),
            "courseName": request["courseName"],
            "examDate": request["examDate"],
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

      .page--student {{
        padding-right: 172px;
      }}

      .student-qr {{
        align-items: center;
        background: rgba(248, 245, 236, 0.96);
        border: 1px solid var(--border);
        border-radius: 16px;
        display: flex;
        flex-direction: column;
        gap: 8px;
        padding: 12px;
        position: absolute;
        right: 28px;
        text-align: center;
        top: 28px;
        width: 126px;
      }}

      .student-qr svg {{
        display: block;
        height: auto;
        width: 100%;
      }}

      .student-qr__label {{
        color: #146c59;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.12em;
        margin: 0;
        text-transform: uppercase;
      }}

      .student-qr__copy {{
        color: var(--muted);
        font-size: 0.78rem;
        line-height: 1.35;
        margin: 0;
      }}

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

        .page--student {{
          padding-right: 42mm;
        }}

        .student-qr {{
          background: #fff;
          border-radius: 12px;
          padding: 8px;
          position: fixed;
          right: 12mm;
          top: 12mm;
          width: 30mm;
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
            f"<li>{html.escape(choice['key'])}. {html.escape(choice['text'])}</li>"
            for choice in question["choices"]
        )
        chapters = ", ".join(question["chapters"]) or "—"
        question_sections.append(
            f"""
<section class="question">
  <div class="question-head">Question {index} · {html.escape(question['sourceQuestionId'])} · Difficulty {question['difficulty']} · Chapters {html.escape(chapters)}</div>
  <p class="question-title">{html.escape(question['question'])}</p>
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
    return f"""
<aside class="student-qr" aria-label="Variant tracking QR code">
  {qr_svg}
  <p class="student-qr__label">Variant QR</p>
  <p class="student-qr__copy">Repeated on every printed page for grading lookup.</p>
</aside>
"""


def render_variant_html(exam_set: dict[str, Any], variant: dict[str, Any]) -> str:
    print_settings = get_print_settings(exam_set)
    summary = render_meta_cards(
        [
            ("Exam Name", print_settings["examName"]),
            ("Course Name", print_settings["courseName"]),
            ("Exam Date", print_settings["examDate"]),
        ]
    )

    question_sections: list[str] = []
    for question in variant["questions"]:
        choices = "".join(
            f"<li>{html.escape(choice['key'])}. {html.escape(choice['text'])}</li>"
            for choice in question["displayChoices"]
        )
        question_sections.append(
            f"""
<section class="question">
  <div class="question-head">Question {question['position']} · {html.escape(question['sourceQuestionId'])} · Difficulty {question['difficulty']}</div>
  <p class="question-title">{html.escape(question['question'])}</p>
  <ul class="choice-list">{choices}</ul>
</section>
"""
        )

    body = summary + "".join(question_sections)
    html_output = render_printable_html(
        title=print_settings["examName"],
        subtitle="Printable student form with QR-based grading lookup.",
        body_html=body,
    )
    qr_markup = render_variant_qr_markup(exam_set, variant)
    html_output = html_output.replace('<body>', "<body>\n" + qr_markup, 1)
    html_output = html_output.replace('class="page"', 'class="page page--student"', 1)
    return html_output


def build_printable_zip(state: AppState, exam_set: dict[str, Any]) -> bytes:
    question_pool = get_question_pool_for_export(state, exam_set)
    buffer = io.BytesIO()
    variants = [variant for variant in exam_set.get("variants", []) if isinstance(variant, dict)]
    annotate_variant_printables(variants)
    base_folder = str(exam_set.get("printableFolderName") or DEFAULT_PRINTABLE_FOLDER)
    question_pool_name = str(exam_set.get("questionPoolFileName") or QUESTION_POOL_PRINTABLE_NAME)

    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            f"{base_folder}/{question_pool_name}",
            render_question_pool_html(exam_set, question_pool),
        )
        for variant in variants:
            variant_file_name = str(
                variant.get("printableFileName")
                or build_variant_printable_filename(variant.get("printableOrdinal", 1), len(variants))
            )
            archive.writestr(
                f"{base_folder}/{variant_file_name}",
                render_variant_html(exam_set, variant),
            )

    return buffer.getvalue()


def build_handler(state: AppState) -> type[BaseHTTPRequestHandler]:
    class QuizRequestHandler(BaseHTTPRequestHandler):
        server_version = "QuizPool/0.2"

        def do_GET(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path == "/api/quiz":
                self.handle_get_quiz()
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
            if parsed.path == "/api/quiz":
                self.handle_put_quiz()
                return
            if parsed.path == "/api/exams/generate":
                self.handle_generate_exams()
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
                }
            )

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
            except OSError as error:
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

        def serve_static(self, raw_path: str) -> None:
            request_path = raw_path or "/"
            if request_path == "/":
                relative = "index.html"
            elif request_path == "/generator":
                relative = "generator.html"
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
        else (db_path.parent / DEFAULT_EXAM_STORE_NAME).resolve()
    )

    if not db_path.is_file():
        raise SystemExit(f"Quiz file not found: {db_path}")
    if not schema_path.is_file():
        raise SystemExit(f"Schema file not found: {schema_path}")
    if not WEB_ROOT.is_dir():
        raise SystemExit(f"Web assets not found: {WEB_ROOT}")

    schema = load_json(schema_path)
    validator = Draft202012Validator(schema)
    initial_quiz = load_json(db_path)
    errors = validation_errors(validator, initial_quiz)
    if errors:
        raise SystemExit(
            "Quiz file does not match the schema:\n"
            + "\n".join(f"- {item['path']}: {item['message']}" for item in errors)
        )

    if exam_store_path.exists():
        try:
            load_exam_store(exam_store_path)
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
