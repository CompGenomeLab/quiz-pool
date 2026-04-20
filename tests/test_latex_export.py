import io
import shutil
import unittest

from pypdf import PdfReader
from omr.layout import PageLayout

from src.quiz_pool.main import (
    LATEX_ENGINE,
    build_omr_sheet_config,
    build_question_pool_latex_document,
    build_latex_font_assets,
    build_student_latex_assets,
    build_student_latex_document,
    build_variant_qr_pdf_bytes,
    latex_engine_available,
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
            "sources": ["Chapter 7"],
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
    if not latex_engine_available():
        return False
    return shutil.which(LATEX_ENGINE) is not None


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

    def test_latex_documents_use_vendored_lm_fonts(self) -> None:
        student_document = build_student_latex_document(sample_exam_set(), sample_variant())
        question_pool_document = build_question_pool_latex_document(
            sample_exam_set(),
            sample_question_pool(),
        )

        self.assertIn(r"\usepackage{fontspec}", student_document)
        self.assertIn(r"\usepackage{unicode-math}", student_document)
        self.assertIn(r"\setmainfont{lmroman10-regular.otf}", student_document)
        self.assertIn(r"\setmathfont{latinmodern-math.otf}", student_document)
        self.assertIn(r"\setmainfont{lmroman10-regular.otf}", question_pool_document)
        self.assertIn(r"\setmathfont{latinmodern-math.otf}", question_pool_document)

    def test_build_latex_font_assets_includes_vendored_fonts(self) -> None:
        assets = build_latex_font_assets()
        self.assertIn("lmroman10-regular.otf", assets)
        self.assertIn("lmroman10-bold.otf", assets)
        self.assertIn("lmroman10-italic.otf", assets)
        self.assertIn("lmroman10-bolditalic.otf", assets)
        self.assertIn("latinmodern-math.otf", assets)

    def test_variant_qr_pdf_matches_omr_qr_box_size(self) -> None:
        qr_pdf = build_variant_qr_pdf_bytes("exam-set-001", "variant-001")
        reader = PdfReader(io.BytesIO(qr_pdf))
        self.assertEqual(len(reader.pages), 1)
        page = reader.pages[0]
        layout = PageLayout()
        self.assertAlmostEqual(float(page.mediabox.width), layout.qr_size, places=2)
        self.assertAlmostEqual(float(page.mediabox.height), layout.qr_size, places=2)

    def test_omr_sheet_title_uses_course_name_and_exam_name(self) -> None:
        config = build_omr_sheet_config(sample_exam_set(), sample_variant())
        self.assertEqual(config.title, "BIO460/560 - BIO460/560 Midterm")

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
