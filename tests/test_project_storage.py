import base64
import tempfile
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator

from src.quiz_pool.main import (
    AppState,
    import_quiz_json_into_project,
    initialize_empty_project,
    load_active_quiz,
    load_internal_schema,
    load_project_generator_draft,
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


if __name__ == "__main__":
    unittest.main()
