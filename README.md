# Quiz Pool Editor

This repo contains a schema-driven quiz JSON file, a browser editor for updating the pool, and an exam
variant generator for building printable shuffled exam forms with persistent IDs.

## Run

```bash
uv run python main.py --db sample_quiz.json
```

Then open:

- `http://127.0.0.1:8000` for the question editor
- `http://127.0.0.1:8000/generator.html` for exam generation

## What It Does

- Loads a quiz JSON database from disk
- Lets you edit quiz metadata, learning objective labels, and questions
- Supports adding and deleting questions
- Supports editing choices, correct answers, linked learning objectives, book locations, difficulty, and explanation
- Generates one shared exam set and multiple unique shuffled variants of that set
- Persists `examSetId` and `variantId` records for later answer lookup and grading
- Lets you customize printable metadata such as institution name, exam name, course name, exam date, start time, total time, and exam rules
- Renders printable student variants with repeated QR codes instead of visible exam/variant IDs, plus a teacher answer-key summary
- Downloads a ZIP of printable HTML documents for the shared question pool and each student variant, using ordinal student filenames inside the archive
- Exports the current generated exam run as JSON
- Validates every save against `scheme.json`
- Writes changes back to the JSON file atomically

## Useful Options

```bash
uv run python main.py --db path/to/quiz.json --schema scheme.json --exam-store generated_exams.json --host 127.0.0.1 --port 8000
```

## Notes

- The app expects the JSON file to already exist and to be valid against the schema on startup.
- Invalid edits are rejected on save and shown in the UI.
- Generated exams are stored in `generated_exams.json` by default, next to the quiz database.
- The generator always builds one shared question set per run, then creates unique variant views of that set.
- Student printables repeat a QR on every page with the raw JSON payload `{"examSetId":"...","variantId":"..."}`.
- Student printables use a fixed A4 portrait header/footer, a cover sheet on page 1, and start questions on page 2 while keeping each question block together on one page.
