import unittest

from src.quiz_pool.main import (
    build_question_pool_entry,
    extract_question_source_labels,
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


if __name__ == "__main__":
    unittest.main()
