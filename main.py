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
DISPLAY_KEYS = ("A", "B", "C", "D", "E")
DEFAULT_PRINTABLE_FOLDER = "exam-printables"
QUESTION_POOL_PRINTABLE_NAME = "question-pool.html"
QUESTION_PAGE_CAPACITY = 38
DEFAULT_EXAM_RULES = [
    "Complete the student information block before the exam begins.",
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
            "examRules": exam_rules,
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
    exam_rules = raw.get("examRules")

    normalized_institution_name = institution_name.strip() if isinstance(institution_name, str) else ""
    normalized_exam_name = exam_name.strip() if isinstance(exam_name, str) else ""
    normalized_course_name = course_name.strip() if isinstance(course_name, str) else ""
    normalized_exam_date = exam_date.strip() if isinstance(exam_date, str) else ""
    normalized_start_time = start_time.strip() if isinstance(start_time, str) else ""
    normalized_total_time = ""
    if isinstance(total_time_minutes, int) and total_time_minutes > 0:
        normalized_total_time = str(total_time_minutes)
    elif isinstance(total_time_minutes, str) and total_time_minutes.strip():
        normalized_total_time = total_time_minutes.strip()
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
        normalized_rules = list(DEFAULT_EXAM_RULES)

    return {
        "institutionName": normalized_institution_name or "Institution Name",
        "examName": normalized_exam_name,
        "courseName": normalized_course_name,
        "examDate": normalized_exam_date,
        "startTime": normalized_start_time,
        "totalTimeMinutes": normalized_total_time,
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
    return qr_code.svg_inline(scale=5, border=2, omitsize=True)


def estimate_wrapped_line_count(text: str, chars_per_line: int) -> int:
    normalized = " ".join(str(text).split())
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
        display: flex;
        flex-direction: column;
        margin: 0 auto;
        height: 297mm;
        padding: 16mm 17mm 15mm;
        width: 210mm;
      }

      .sheet-page:last-child {
        break-after: auto;
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
        gap: 4.8mm;
      }

      .question-block {
        border: 0.3mm solid var(--border);
        break-inside: avoid;
        page-break-inside: avoid;
        padding: 4.2mm 4.6mm;
      }

      .question-block__number {
        font-size: 10pt;
        font-weight: 700;
        margin: 0 0 2.2mm;
        text-transform: uppercase;
      }

      .question-block__text {
        font-size: 11pt;
        line-height: 1.45;
        margin: 0;
      }

      .choice-list {
        display: flex;
        flex-direction: column;
        gap: 2.4mm;
        list-style: none;
        margin: 4mm 0 0;
        padding: 0;
      }

      .choice-list li {
        align-items: start;
        border: 0.2mm solid #888888;
        display: grid;
        gap: 3mm;
        grid-template-columns: 7mm minmax(0, 1fr);
        padding: 2.6mm 3mm;
      }

      .choice-key {
        font-weight: 700;
      }

      .choice-text {
        line-height: 1.4;
      }

      .sheet-page--question {
        padding: 11mm 14mm 10mm;
      }

      .sheet-page--question .sheet-header {
        gap: 3mm;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1.45fr) 19mm;
      }

      .sheet-page--question .sheet-header__institution,
      .sheet-page--question .sheet-header__title {
        min-height: 19mm;
      }

      .sheet-page--question .sheet-header__institution {
        font-size: 8.4pt;
        line-height: 1.25;
        padding-top: 0.8mm;
      }

      .sheet-page--question .sheet-header__title {
        font-size: 11.5pt;
        letter-spacing: 0.02em;
        line-height: 1.15;
        padding: 0 2mm;
      }

      .sheet-page--question .sheet-header__qr {
        height: 19mm;
        padding: 0.8mm;
        width: 19mm;
      }

      .sheet-page--question .sheet-header-rule {
        margin: 2mm 0 3mm;
      }

      .sheet-page--question .sheet-footer {
        font-size: 8pt;
        margin-top: 3.5mm;
        padding-top: 2mm;
      }

      .sheet-page--question .question-stack {
        gap: 2.4mm;
      }

      .sheet-page--question .question-block {
        padding: 2.6mm 3mm;
      }

      .sheet-page--question .question-block__number {
        font-size: 8.8pt;
        margin-bottom: 1.4mm;
      }

      .sheet-page--question .question-block__text {
        font-size: 9.8pt;
        line-height: 1.28;
      }

      .sheet-page--question .choice-list {
        gap: 1.2mm;
        margin-top: 2.2mm;
      }

      .sheet-page--question .choice-list li {
        gap: 2mm;
        grid-template-columns: 5.4mm minmax(0, 1fr);
        padding: 1.5mm 1.8mm;
      }

      .sheet-page--question .choice-key,
      .sheet-page--question .choice-text {
        font-size: 9.2pt;
        line-height: 1.22;
      }

      .pagination-measure {
        left: -1000vw;
        pointer-events: none;
        position: absolute;
        top: 0;
        visibility: hidden;
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
            render_student_line_field("Student Name", wide=True),
            render_student_line_field("Student ID"),
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
  <span class="choice-text">{html.escape(choice['text'])}</span>
</li>
"""
            for choice in question["displayChoices"]
        )
        question_blocks.append(
            f"""
<article class="question-block" data-question-position="{question['position']}">
  <p class="question-block__number">Question {question['position']}</p>
  <p class="question-block__text">{html.escape(question['question'])}</p>
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


def render_student_pagination_script() -> str:
    return """
    <script>
      (() => {
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

          updatePageCounters(finalPages.length + 1);
        }

        let scheduled = false;
        function schedulePagination() {
          if (scheduled) return;
          scheduled = true;
          requestAnimationFrame(() => {
            scheduled = false;
            paginateStudentView();
          });
        }

        window.addEventListener('DOMContentLoaded', schedulePagination, { once: true });
        window.addEventListener('load', schedulePagination, { once: true });
        window.addEventListener('resize', schedulePagination);
        window.addEventListener('beforeprint', paginateStudentView);
      })();
    </script>
"""


def render_variant_html(exam_set: dict[str, Any], variant: dict[str, Any]) -> str:
    print_settings = get_print_settings(exam_set)
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
    <div id="pagination-measure" class="pagination-measure" aria-hidden="true"></div>
{render_student_pagination_script()}
  </body>
</html>
"""


def build_printable_zip(state: AppState, exam_set: dict[str, Any]) -> bytes:
    question_pool = get_question_pool_for_export(state, exam_set)
    buffer = io.BytesIO()
    variants = [variant for variant in exam_set.get("variants", []) if isinstance(variant, dict)]
    annotate_variant_printables(variants)
    annotate_variant_print_layouts(variants)
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
