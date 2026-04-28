import base64
import tempfile
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator

from src.quiz_pool.main import (
    AppState,
    UploadedFile,
    clear_grading_uploads,
    grading_upload_label,
    import_quiz_json_content_into_project,
    import_quiz_json_into_project,
    initialize_empty_project,
    load_active_quiz,
    load_internal_schema,
    load_project_generator_draft,
    normalize_system_file_dialog_request,
    parse_multipart_uploads,
    replace_grading_uploads,
    system_file_dialog_allowed,
    store_project_asset,
    write_project_generator_draft,
)


ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
)


class ProjectStorageTests(unittest.TestCase):
    def test_empty_project_starts_without_questions_and_can_import_quiz(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project_path = Path(temp_dir) / "course.quizpool"
            db_path = Path("sample_quiz.json").resolve()
            validator = Draft202012Validator(load_internal_schema())

            initialize_empty_project(project_path)
            state = AppState(
                db_path=project_path,
                exam_store_path=project_path,
                project_path=project_path,
                validator=validator,
            )

            quiz = load_active_quiz(state)
            self.assertEqual(quiz["questions"], [])

            import_quiz_json_into_project(
                project_path=project_path,
                quiz_path=db_path,
                validator=validator,
            )
            quiz = load_active_quiz(state)
            self.assertGreater(len(quiz["questions"]), 0)

            write_project_generator_draft(project_path, {"selection": {"questionCount": 3}})
            draft = load_project_generator_draft(project_path)
            self.assertEqual(draft["selection"]["questionCount"], 3)

    def test_import_quiz_json_content_into_project(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project_path = Path(temp_dir) / "course.quizpool"
            validator = Draft202012Validator(load_internal_schema())
            initialize_empty_project(project_path)
            state = AppState(
                db_path=project_path,
                exam_store_path=project_path,
                project_path=project_path,
                validator=validator,
            )

            import_quiz_json_content_into_project(
                project_path=project_path,
                content=Path("sample_quiz.json").read_text(encoding="utf-8"),
                validator=validator,
            )

            quiz = load_active_quiz(state)
            self.assertGreater(len(quiz["questions"]), 0)

    def test_project_asset_upload_records_png_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project_path = Path(temp_dir) / "course.quizpool"
            asset = store_project_asset(
                project_path,
                filename="dot.png",
                mime_type="image/png",
                data=ONE_PIXEL_PNG,
            )

            self.assertEqual(asset["mimeType"], "image/png")
            self.assertEqual(asset["width"], 1)
            self.assertEqual(asset["height"], 1)
            self.assertEqual(asset["sizeBytes"], len(ONE_PIXEL_PNG))

    def test_system_file_dialog_allows_expected_selection_types(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            project_path = temp_path / "course.quizpool"
            quiz_json_path = temp_path / "quiz.json"
            pdf_path = temp_path / "scan.pdf"
            project_path.write_text("", encoding="utf-8")
            quiz_json_path.write_text("{}", encoding="utf-8")
            pdf_path.write_bytes(b"%PDF-1.7\n")

            self.assertTrue(system_file_dialog_allowed(project_path, "project", "file"))
            self.assertTrue(system_file_dialog_allowed(quiz_json_path, "quiz-json", "file"))
            self.assertTrue(system_file_dialog_allowed(pdf_path, "pdf-or-dir", "file"))
            self.assertTrue(system_file_dialog_allowed(temp_path, "pdf-or-dir", "directory"))
            self.assertTrue(system_file_dialog_allowed(temp_path, "directory", "directory"))
            self.assertFalse(system_file_dialog_allowed(project_path, "quiz-json", "file"))
            self.assertFalse(system_file_dialog_allowed(temp_path, "project", "directory"))

    def test_system_file_dialog_request_rejects_unsupported_mode(self) -> None:
        request, errors = normalize_system_file_dialog_request(
            {"purpose": "project", "mode": "directory"},
            fallback_path=Path.cwd(),
        )

        self.assertIsNone(request)
        self.assertEqual(errors[0]["path"], "mode")

    def test_parse_multipart_uploads_reads_browser_file_fields(self) -> None:
        boundary = "----quizpool-test"
        body = (
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="pdfs"; filename="scan.pdf"\r\n'
            "Content-Type: application/pdf\r\n"
            "\r\n"
        ).encode("utf-8") + b"%PDF-1.7\n" + f"\r\n--{boundary}--\r\n".encode("utf-8")

        files = parse_multipart_uploads(f"multipart/form-data; boundary={boundary}", body)

        self.assertEqual(len(files), 1)
        self.assertEqual(files[0].field_name, "pdfs")
        self.assertEqual(files[0].filename, "scan.pdf")
        self.assertEqual(files[0].data, b"%PDF-1.7\n")

    def test_replace_grading_uploads_writes_unique_pdf_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project_path = Path(temp_dir) / "course.quizpool"
            state = AppState(
                db_path=project_path,
                exam_store_path=project_path,
                project_path=project_path,
                validator=Draft202012Validator(load_internal_schema()),
            )
            try:
                upload_path = replace_grading_uploads(
                    state,
                    [
                        UploadedFile("pdfs", "scan.pdf", "application/pdf", b"%PDF-1.7\n"),
                        UploadedFile("pdfs", "nested/scan.pdf", "application/pdf", b"%PDF-1.7\n"),
                    ],
                )

                self.assertTrue((upload_path / "scan.pdf").is_file())
                self.assertTrue((upload_path / "scan-2.pdf").is_file())
                self.assertEqual(grading_upload_label(state.grading_upload_files), "Uploaded PDFs (2 files)")
            finally:
                clear_grading_uploads(state)


if __name__ == "__main__":
    unittest.main()
