import io
import shutil
import unittest

from pypdf import PdfReader

from src.quiz_pool.main import (
    build_question_pool_latex_document,
    build_student_latex_assets,
    build_student_latex_document,
    latex_engine_candidates,
    latex_escape_text_segment,
    render_latex_choice_rows,
    render_latex_to_pdf,
    render_rich_text_latex,
)


def sample_exam_set() -> dict:
    return {
        "examSetId": "exam-set-001",
        "quiz": {"title": "Population Genetics"},
        "printSettings": {
            "institutionName": "Sabanci University",
            "examName": "BIO460/560 Midterm",
            "courseName": "BIO460/560",
            "examDate": "2026-04-27",
            "startTime": "10:30",
            "totalTimeMinutes": 120,
            "instructor": "Dr. Ada Lovelace",
            "allowedMaterials": "One handwritten note sheet",
            "omrInstructions": "Fully fill bubbles.",
            "examRules": [
                "Fully fill bubbles. Do not leave any Student ID field blank. Include leading zeros.",
                "Read every question carefully and select all correct answers for each question.",
            ],
        },
    }


def sample_variant() -> dict:
    return {
        "variantId": "variant-001",
        "printableOrdinal": 1,
        "questions": [
            {
                "position": 1,
                "question": "For [math]\\Delta\\bar{z} = h^2 S[/math], what does [math]h^2[/math] represent?",
                "points": 1,
                "displayChoices": [
                    {"key": "A", "text": "Heritability"},
                    {"key": "B", "text": "Selection differential"},
                    {"key": "C", "text": "Mutation rate"},
                    {"key": "D", "text": "Migration rate"},
                    {"key": "E", "text": "Effective population size [math]N_e[/math]"},
                ],
            }
        ],
    }


def sample_question_pool() -> list[dict]:
    return [
        {
            "sourceQuestionId": "C07Q07",
            "question": "In the breeder's equation, [math]\\Delta\\bar{z} = h^2 S[/math].",
            "difficulty": 2,
            "points": 1,
            "chapters": ["Chapter 7"],
            "learningObjectives": [{"id": "LO1", "label": "Quantify selection response"}],
            "choices": [
                {"key": "A", "text": "Option [math]p_2[/math]"},
                {"key": "B", "text": "Option B"},
                {"key": "C", "text": "Option C"},
                {"key": "D", "text": "Option D"},
            ],
            "sourceCorrectAnswers": ["A"],
        }
    ]


def latex_engine_is_usable() -> bool:
    if not latex_engine_candidates():
        return False
    return shutil.which("pdflatex") is not None


class LatexExportTests(unittest.TestCase):
    def test_latex_escape_text_segment_escapes_special_characters(self) -> None:
        escaped = latex_escape_text_segment(r"100% A_B\C#D{E}~^", preserve_linebreaks=False)
        self.assertEqual(
            escaped,
            r"100\% A\_B\textbackslash{}C\#D\{E\}\textasciitilde{}\textasciicircum{}",
        )

    def test_render_rich_text_latex_preserves_math_blocks(self) -> None:
        rendered = render_rich_text_latex(
            r"Area [math]A = \pi r^{2}[/math] grows with [math]r[/math].",
            preserve_linebreaks=False,
        )
        self.assertIn(r"\(A = \pi r^{2}\)", rendered)
        self.assertIn(r"\(r\)", rendered)
        self.assertNotIn("[math]", rendered)

    def test_render_latex_choice_rows_is_row_major(self) -> None:
        rows = render_latex_choice_rows(
            [
                {"key": "A", "text": "Alpha"},
                {"key": "B", "text": "Beta"},
                {"key": "C", "text": "Gamma"},
                {"key": "D", "text": "Delta"},
                {"key": "E", "text": "Epsilon"},
            ]
        )
        self.assertIn(r"\textbf{A.} Alpha & \textbf{B.} Beta", rows)
        self.assertIn(r"\textbf{C.} Gamma & \textbf{D.} Delta", rows)
        self.assertIn(r"\textbf{E.} Epsilon & ~", rows)

    def test_latex_documents_include_instructor_and_materials(self) -> None:
        student_document = build_student_latex_document(sample_exam_set(), sample_variant())
        question_pool_document = build_question_pool_latex_document(
            sample_exam_set(),
            sample_question_pool(),
        )

        self.assertIn(r"\newcommand{\instructor}{Dr. Ada Lovelace}", student_document)
        self.assertIn(
            r"\newcommand{\allowedmaterials}{One handwritten note sheet}",
            student_document,
        )
        self.assertIn(r"\newcommand{\instructor}{Dr. Ada Lovelace}", question_pool_document)
        self.assertIn(
            r"\newcommand{\allowedmaterials}{One handwritten note sheet}",
            question_pool_document,
        )

    @unittest.skipUnless(latex_engine_is_usable(), "no LaTeX engine is available")
    def test_student_and_question_pool_documents_compile(self) -> None:
        student_pdf = render_latex_to_pdf(
            build_student_latex_document(sample_exam_set(), sample_variant()),
            job_name="student-smoke",
            assets=build_student_latex_assets(sample_exam_set(), sample_variant()),
        )
        question_pool_pdf = render_latex_to_pdf(
            build_question_pool_latex_document(sample_exam_set(), sample_question_pool()),
            job_name="question-pool-smoke",
        )
        self.assertGreaterEqual(len(PdfReader(io.BytesIO(student_pdf)).pages), 1)
        self.assertGreaterEqual(len(PdfReader(io.BytesIO(question_pool_pdf)).pages), 1)


if __name__ == "__main__":
    unittest.main()
