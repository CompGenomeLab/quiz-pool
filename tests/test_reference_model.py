import unittest
import tempfile
from pathlib import Path

from jsonschema import Draft202012Validator

import random

from src.quiz_pool.main import (
    AppState,
    build_question_pool_entry,
    extract_question_source_labels,
    generate_exam_run,
    load_internal_schema,
    normalize_generation_request,
    question_selection_weight,
    weighted_sample_without_replacement,
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


    def test_normalize_generation_request_accepts_source_weights(self) -> None:
        request, errors = normalize_generation_request(
            {
                "questionCount": 1,
                "variantCount": 1,
                "sources": ["Chapter 4"],
                "sourceWeights": {"Chapter 4": 2},
                "difficulties": [],
                "learningObjectiveIds": [],
                "includeQuestionIds": [],
                "excludeQuestionIds": [],
            },
            sample_quiz(),
        )

        self.assertEqual(errors, [])
        self.assertEqual(request["sourceWeights"], {"Chapter 4": 2.0})
        # chapterWeights alias mirrors sourceWeights for symmetry with chapters.
        self.assertEqual(request["chapterWeights"], {"Chapter 4": 2.0})

    def test_normalize_generation_request_rejects_non_positive_weight(self) -> None:
        _, errors = normalize_generation_request(
            {
                "questionCount": 1,
                "variantCount": 1,
                "sources": ["Chapter 4"],
                "sourceWeights": {"Chapter 4": 0},
                "difficulties": [],
                "learningObjectiveIds": [],
                "includeQuestionIds": [],
                "excludeQuestionIds": [],
            },
            sample_quiz(),
        )

        self.assertTrue(any(error["path"] == "sourceWeights.Chapter 4" for error in errors))

    def test_question_selection_weight_uses_max_of_matching_sources(self) -> None:
        question = sample_question()  # belongs to "Chapter 4" among others
        selected = {"Chapter 4", "Figure 2"}
        weights = {"Chapter 4": 3.0}
        # Max of matching selected sources, default 1 for the unweighted match.
        self.assertEqual(question_selection_weight(question, selected, weights), 3.0)
        # No matching selected source falls back to the default weight.
        self.assertEqual(question_selection_weight(question, {"Nope"}, weights), 1.0)

    def test_weighted_sample_without_replacement_returns_distinct_items(self) -> None:
        rng = random.Random(7)
        items = ["a", "b", "c", "d"]
        drawn = weighted_sample_without_replacement(items, [1, 1, 1, 1], 3, rng)
        self.assertEqual(len(drawn), 3)
        self.assertEqual(len(set(drawn)), 3)
        self.assertTrue(set(drawn).issubset(set(items)))

    def test_generate_exam_run_weights_bias_selection_toward_heavier_source(self) -> None:
        quiz = two_source_quiz()
        heavy_picks = 0
        light_picks = 0
        # Weighting is probabilistic, not a quota, so check the aggregate bias
        # across many seeds rather than any single deterministic outcome.
        for seed in range(200):
            request, errors = normalize_generation_request(
                {
                    "questionCount": 4,
                    "variantCount": 1,
                    "sources": ["Heavy", "Light"],
                    "sourceWeights": {"Heavy": 4},
                    "difficulties": [],
                    "learningObjectiveIds": [],
                    "includeQuestionIds": [],
                    "excludeQuestionIds": [],
                    "generationSeed": f"seed-{seed}",
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
                result = generate_exam_run(state, quiz, request)
            for question_id in result["selection"]["selectedQuestionIds"]:
                if question_id.startswith("H"):
                    heavy_picks += 1
                else:
                    light_picks += 1

        # Equal pool sizes; a 4x weight must draw the heavy source far more often.
        self.assertGreater(heavy_picks, light_picks * 2)

    def test_generate_exam_run_weighted_selection_is_repeatable(self) -> None:
        quiz = two_source_quiz()
        payload = {
            "questionCount": 4,
            "variantCount": 1,
            "sources": ["Heavy", "Light"],
            "sourceWeights": {"Heavy": 3},
            "difficulties": [],
            "learningObjectiveIds": [],
            "includeQuestionIds": [],
            "excludeQuestionIds": [],
            "generationSeed": "weighted-seed",
        }
        request, errors = normalize_generation_request(payload, quiz)
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
        self.assertEqual(
            first["selection"]["selectedQuestionIds"],
            second["selection"]["selectedQuestionIds"],
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


def two_source_quiz() -> dict:
    questions = []
    for source, prefix in (("Heavy", "H"), ("Light", "L")):
        for index in range(1, 11):
            question = sample_question()
            question["id"] = f"{prefix}{index}"
            question["question"] = f"{source} question {index}"
            question["shuffleChoices"] = False
            question["locations"] = [{"chapter": source}]
            questions.append(question)
    return {
        "learningObjectives": [
            {"id": "LO1", "label": "Explain the concept"},
        ],
        "questions": questions,
    }


if __name__ == "__main__":
    unittest.main()
