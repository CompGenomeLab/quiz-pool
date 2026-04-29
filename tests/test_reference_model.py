import unittest
import tempfile
from pathlib import Path

from jsonschema import Draft202012Validator

from src.quiz_pool.main import (
    AppState,
    build_question_pool_entry,
    extract_question_source_labels,
    generate_exam_run,
    load_internal_schema,
    normalize_generation_request,
)


def sample_question() -> dict:
    return {
        "id": "Q1",
        "question": "Sample question",
        "choices": [
            {"key": "A", "text": "Alpha"},
            {"key": "B", "text": "Beta"},
        ],
        "shuffleChoices": True,
        "learningObjectiveIds": ["LO1"],
        "correctAnswers": ["A"],
        "locations": [
            {
                "source": "Example Biology Text",
                "chapter": "Chapter 4",
                "section": "4.2",
                "page": "88-89",
                "reference": "Key derivation",
            },
            {
                "url": "https://example.com/evolution",
            },
            {
                "reference": "Figure 2",
            },
        ],
        "points": 2,
        "difficulty": 3,
        "explanation": "Alpha is correct.",
    }


def sample_quiz() -> dict:
    return {
        "learningObjectives": [
            {"id": "LO1", "label": "Explain the concept"},
        ],
        "questions": [sample_question()],
    }


class ReferenceModelTests(unittest.TestCase):
    def test_extract_question_source_labels_supports_multiple_reference_kinds(self) -> None:
        labels = extract_question_source_labels(sample_question())
        self.assertEqual(labels, ["Chapter 4", "https://example.com/evolution", "Figure 2"])

    def test_normalize_generation_request_accepts_sources(self) -> None:
        request, errors = normalize_generation_request(
            {
                "questionCount": 1,
                "variantCount": 1,
                "sources": ["Chapter 4"],
                "difficulties": [3],
                "learningObjectiveIds": ["LO1"],
                "includeQuestionIds": [],
                "excludeQuestionIds": [],
            },
            sample_quiz(),
        )

        self.assertEqual(errors, [])
        self.assertEqual(request["sources"], ["Chapter 4"])
        self.assertEqual(request["chapters"], ["Chapter 4"])

    def test_build_question_pool_entry_includes_sources_alias(self) -> None:
        entry = build_question_pool_entry(sample_question(), {"LO1": "Explain the concept"})
        self.assertEqual(entry["sources"], ["Chapter 4", "https://example.com/evolution", "Figure 2"])
        self.assertEqual(entry["chapters"], ["Chapter 4", "https://example.com/evolution", "Figure 2"])

    def test_generate_exam_run_uses_saved_seed_for_repeatable_selection_and_variants(self) -> None:
        quiz = seeded_quiz()
        request, errors = normalize_generation_request(
            {
                "questionCount": 2,
                "variantCount": 2,
                "sources": [],
                "difficulties": [],
                "learningObjectiveIds": [],
                "includeQuestionIds": [],
                "excludeQuestionIds": [],
                "generationSeed": "repeatable-seed",
            },
            quiz,
        )
        self.assertEqual(errors, [])
        with tempfile.TemporaryDirectory() as temp_dir:
            project_path = Path(temp_dir) / "course.quizpool"
            state = AppState(
                db_path=project_path,
                exam_store_path=project_path,
                project_path=project_path,
                validator=Draft202012Validator(load_internal_schema()),
            )
            first = generate_exam_run(state, quiz, request)
            second = generate_exam_run(state, quiz, request)

        self.assertEqual(first["generationSeed"], "repeatable-seed")
        self.assertEqual(first["selection"]["selectedQuestionIds"], second["selection"]["selectedQuestionIds"])
        self.assertEqual(
            [variant["signature"] for variant in first["variants"]],
            [variant["signature"] for variant in second["variants"]],
        )


def seeded_quiz() -> dict:
    questions = []
    for index in range(1, 4):
        question = sample_question()
        question["id"] = f"Q{index}"
        question["question"] = f"Seeded question {index}"
        question["shuffleChoices"] = False
        questions.append(question)
    return {
        "learningObjectives": [
            {"id": "LO1", "label": "Explain the concept"},
        ],
        "questions": questions,
    }


if __name__ == "__main__":
    unittest.main()
