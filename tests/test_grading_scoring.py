import unittest

from src.quiz_pool.main import (
    analyze_grade_result,
    build_grading_report,
    earns_full_credit,
    normalize_grading_formula_payload,
    recalculate_grading_row,
)


class EarnsFullCreditTests(unittest.TestCase):
    def test_exact_match_earns_full_credit(self) -> None:
        self.assertTrue(earns_full_credit(["A", "C"], ["C", "A"]))

    def test_missing_correct_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit(["A"], ["A", "C"]))

    def test_extra_wrong_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit(["A", "C", "D"], ["A", "C"]))

    def test_blank_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit([], ["A", "C"]))


class GradingFormulaTests(unittest.TestCase):
    def test_default_formula_applies_no_wrong_answer_penalty(self) -> None:
        formula, errors = normalize_grading_formula_payload(None)
        self.assertEqual(errors, [])
        row = recalculate_grading_row(sample_scoring_row(), formula)

        self.assertEqual(row["summary"]["earnedPoints"], 2)
        self.assertEqual(row["summary"]["penaltyPoints"], 0)
        self.assertEqual(row["summary"]["wrongCount"], 2)
        self.assertEqual(row["summary"]["questionCount"], 4)
        self.assertEqual(row["summary"]["scorePercent"], 33.333333)
        self.assertEqual(row["summary"]["wrongPercent"], 50)

    def test_fixed_wrong_answer_penalty_subtracts_per_wrong_or_invalid_answer(self) -> None:
        formula, errors = normalize_grading_formula_payload(
            {"mode": "fixed", "wrongPenalty": 0.25}
        )
        self.assertEqual(errors, [])
        row = recalculate_grading_row(sample_scoring_row(), formula)

        self.assertEqual(row["summary"]["earnedPoints"], 1.5)
        self.assertEqual(row["summary"]["penaltyPoints"], 0.5)
        self.assertEqual(row["learningObjectiveSummary"][0]["earnedPoints"], 1.75)
        self.assertEqual(row["learningObjectiveSummary"][1]["earnedPoints"], 0)
        self.assertEqual(row["learningObjectiveSummary"][0]["scorePercent"], 43.75)
        self.assertEqual(row["learningObjectiveSummary"][0]["wrongPercent"], 50)
        self.assertEqual(row["learningObjectiveSummary"][1]["scorePercent"], 0)
        self.assertEqual(row["learningObjectiveSummary"][1]["blankOrMissingPercent"], 50)

    def test_choice_weighted_penalty_uses_question_points_and_choice_count(self) -> None:
        formula, errors = normalize_grading_formula_payload({"mode": "choice_weighted"})
        self.assertEqual(errors, [])
        row = recalculate_grading_row(sample_scoring_row(), formula)

        self.assertEqual(row["summary"]["earnedPoints"], 0.833333)
        self.assertEqual(row["summary"]["penaltyPoints"], 1.166667)

    def test_penalties_do_not_make_scores_negative(self) -> None:
        formula, errors = normalize_grading_formula_payload(
            {"mode": "fixed", "wrongPenalty": 5}
        )
        self.assertEqual(errors, [])
        row = recalculate_grading_row(sample_scoring_row(), formula)

        self.assertEqual(row["summary"]["earnedPoints"], 0)
        self.assertEqual(row["summary"]["scorePercent"], 0)
        self.assertEqual(row["learningObjectiveSummary"][1]["earnedPoints"], 0)
        self.assertEqual(row["questionDetails"][1]["earnedPoints"], 0)
        self.assertEqual(row["questionDetails"][2]["earnedPoints"], 0)

    def test_grading_report_totals_include_normalized_percentages(self) -> None:
        formula, errors = normalize_grading_formula_payload(None)
        self.assertEqual(errors, [])
        row = recalculate_grading_row(sample_scoring_row(), formula)
        report = build_grading_report([row], formula)

        self.assertEqual(report["total"]["questionCount"], 4)
        self.assertEqual(report["total"]["scorePercent"], 33.333333)
        self.assertEqual(report["total"]["wrongPercent"], 50)
        self.assertEqual(report["learningObjectives"][0]["correctPercent"], 50)

    def test_grading_objectives_are_read_from_matched_variant_question(self) -> None:
        exam_set = {
            "examSetId": "exam-set-001",
            "quiz": {"title": "Original Quiz"},
            "printSettings": {"examName": "Midterm"},
        }
        variant = {
            "variantId": "variant-001",
            "questions": [
                {
                    "position": 1,
                    "question": "Stored variant question",
                    "points": 2,
                    "displayChoices": [
                        {"key": "A", "text": "Alpha"},
                        {"key": "B", "text": "Beta"},
                    ],
                    "displayCorrectAnswers": ["A"],
                    "learningObjectiveIds": ["LO-from-variant"],
                    "learningObjectives": [
                        {"id": "LO-from-variant", "label": "Variant objective"}
                    ],
                }
            ],
        }

        row = analyze_grade_result(
            {
                "source_pdf": "scan.pdf",
                "qr_data": {
                    "examSetId": "exam-set-001",
                    "variantId": "variant-001",
                },
                "student_id": "123",
                "marked_answers": {"1": ["A"]},
            },
            {"variant-001": (exam_set, variant)},
        )

        self.assertEqual(
            row["questionDetails"][0]["learningObjectives"],
            [{"id": "LO-from-variant", "label": "Variant objective"}],
        )
        self.assertEqual(row["learningObjectiveSummary"][0]["id"], "LO-from-variant")
        self.assertEqual(row["learningObjectiveSummary"][0]["earnedPoints"], 2)


def sample_scoring_row() -> dict:
    return {
        "questionDetails": [
            {
                "status": "correct",
                "points": 2,
                "allowedChoices": ["A", "B", "C", "D"],
                "learningObjectives": [{"id": "LO1", "label": "Objective 1"}],
            },
            {
                "status": "incorrect",
                "points": 2,
                "allowedChoices": ["A", "B", "C", "D"],
                "learningObjectives": [{"id": "LO1", "label": "Objective 1"}],
            },
            {
                "status": "invalid",
                "points": 1,
                "allowedChoices": ["A", "B", "C"],
                "learningObjectives": [{"id": "LO2", "label": "Objective 2"}],
            },
            {
                "status": "blank",
                "points": 1,
                "allowedChoices": ["A", "B", "C", "D"],
                "learningObjectives": [{"id": "LO2", "label": "Objective 2"}],
            },
        ]
    }


if __name__ == "__main__":
    unittest.main()
