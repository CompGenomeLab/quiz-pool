# OMR Sheet Generator

This project provides three separate entry points:

- `omr`: generate a blank OMR sheet PDF
- `omr-grade`: read filled sheet PDFs and return JSON
- `omr-annotate`: write annotated PDFs using grading results and an optional answer key

It also exposes matching Python APIs.

## Requirements

- Python `3.12+`
- [`uv`](https://docs.astral.sh/uv/)
- Poppler installed so `pdftoppm` is available on `PATH`

## Install

```bash
uv sync
```

CLI help:

```bash
uv run omr --help
uv run omr-grade --help
uv run omr-annotate --help
```

## Sheet Layout

Generated sheets are:

- A4 portrait
- marker-based for registration
- 8-digit student ID
- handwritten template fields for `Name`, `ID`, and `Signature`
- 5 answer columns on a single page
- question columns with at most `20` questions per column
- up to `100` questions on a single page
- `2` to `5` answer options per question, fixed across the whole sheet
- displayed PDF text uses Latin Modern and falls back to Computer Modern, resolved from the system or cached locally on first use
- QR-backed with `examSetId` and `variantId`

The QR payload format is:

```json
{
  "examSetId": "f6adcc63-71dc-412c-9c8d-a4609df454ff",
  "variantId": "37e3d65f-e540-4e34-b438-549e731be3b0"
}
```

## CLI Usage

### 1. Generate a Blank Sheet

```bash
uv run omr \
  --questions 100 \
  --choices 5 \
  --exam-set-id f6adcc63-71dc-412c-9c8d-a4609df454ff \
  --variant-id 37e3d65f-e540-4e34-b438-549e731be3b0 \
  --output omr-sheet.pdf
```

With custom title and instructions:

```bash
uv run omr \
  --questions 40 \
  --choices 5 \
  --exam-set-id exam-set-001 \
  --variant-id variant-a \
  --output exam-a.pdf \
  --title "Midterm Exam A" \
  --instructions "Fill bubbles completely. Use a dark pencil."
```

Arguments:

- `--questions`: required total number of questions to print
- `--choices`: required number of choices for every question
- `--exam-set-id`: required
- `--variant-id`: required
- `--output`: optional, defaults to `omr-sheet.pdf`
- `--title`: optional
- `--instructions`: optional, defaults to `Fill bubbles completely.`

Rules:

- question count must be at least `1`
- question count must not exceed `100`
- choice count must be between `2` and `5`
- every generated question uses the same printed choices
- questions are numbered from `1`
- the single-page layout uses `5` columns with `20` rows each
- the sheet holds at most `100` questions on one page
- the final sheet does not include page numbering
- `--choices 4` prints `A-D`; `--choices 5` prints `A-E`

### 2. Grade a Filled Sheet

Single PDF:

```bash
uv run omr-grade filled-sheet.pdf
```

Directory of PDFs:

```bash
uv run omr-grade filled-sheets/
```

`omr-grade` prints JSON only. It does not write PDFs.

### 3. Annotate a Filled Sheet

Single PDF:

```bash
uv run omr-annotate filled-sheet.pdf --output annotated.pdf
```

Directory of PDFs:

```bash
uv run omr-annotate filled-sheets/ --output reviewed/
```

With an answer key from a file:

```bash
uv run omr-annotate filled-sheet.pdf \
  --output annotated.pdf \
  --correct-answers tests/answer-key.json
```

With an inline answer key:

```bash
uv run omr-annotate filled-sheet.pdf \
  --output annotated.pdf \
  --correct-answers '{"1":["D"],"3":["B","C"]}'
```

Arguments:

- `path`: required PDF or directory of PDFs
- `--output`: required output PDF path or directory
- `--correct-answers`: optional JSON string or JSON file path

Output path behavior:

- for a single input PDF:
  - if `--output` ends with `.pdf`, that exact file is written
  - otherwise it is treated as a directory and `<source>-annotated.pdf` is written inside it
- for a directory input:
  - `--output` is treated as a directory
  - each input file produces `<source-stem>-annotated.pdf`

## Python API

The package exports:

```python
from omr import (
    SheetConfig,
    generate_omr_sheet,
    grade_pdf,
    grade_path,
    grade_directory,
    annotate_pdf,
    annotate_path,
    annotate_directory,
)
```

### Generate a Sheet

```python
from omr import SheetConfig, generate_omr_sheet

config = SheetConfig(
    question_count=100,
    choice_count=4,
    exam_set_id="f6adcc63-71dc-412c-9c8d-a4609df454ff",
    variant_id="37e3d65f-e540-4e34-b438-549e731be3b0",
    title="Optical Mark Recognition Sheet",
    instructions="Fill bubbles completely.",
)

generate_omr_sheet(config, "omr-sheet.pdf")
```

`generate_omr_sheet(config, destination)` accepts:

- `config`: `SheetConfig`
- `destination`: file path or binary output stream

### Grade One PDF

```python
from omr import grade_pdf

result = grade_pdf("filled-sheet.pdf")

print(result.qr_data)
print(result.student_id)
print(result.marked_answers)
print(result.omr_error)
```

Return type: `GradeResult`

Fields:

- `qr_data: dict | str | None`
- `student_id: str`
- `marked_answers: dict[str, list[str]]`
- `omr_error: str`

Notes:

- grading raises an OMR error if any student ID column is empty or has multiple marked digits
- the default A4 layout uses 5 columns and 20 rows, capped at 100 questions on one page

### Grade a Directory

```python
from omr import grade_directory

results = grade_directory("filled-sheets")
for result in results:
    print(result.source_pdf, result.student_id, result.omr_error)
```

Return type: `list[BatchGradeResult]`

Fields per item:

- `source_pdf: str`
- `qr_data: dict | str | None`
- `student_id: str`
- `marked_answers: dict[str, list[str]]`
- `omr_error: str`

Batch grading behavior:

- only `.pdf` files are processed
- failures do not abort the batch
- failed files still return one result entry
- empty or multiply marked student ID columns are reported through `omr_error`
- failed entries use:
  - `qr_data = None`
  - `student_id = ""`
  - `marked_answers = {}`
  - `omr_error = "<message>"`

### Annotate One PDF

```python
from omr import annotate_pdf

result = annotate_pdf(
    "filled-sheet.pdf",
    "annotated.pdf",
    correct_answers={
        "1": ["D"],
        "3": ["B", "C"],
    },
)

print(result.annotated_pdf)
```

Return type: `AnnotateResult`

Fields:

- `qr_data: dict | str | None`
- `student_id: str`
- `marked_answers: dict[str, list[str]]`
- `omr_error: str`
- `annotated_pdf: str`

### Annotate a Directory

```python
from omr import annotate_directory

results = annotate_directory(
    "filled-sheets",
    "reviewed",
    correct_answers={"1": ["D"]},
)

for result in results:
    print(result.source_pdf, result.annotated_pdf)
```

Return type: `list[BatchAnnotateResult]`

Fields per item:

- `source_pdf: str`
- `qr_data: dict | str | None`
- `student_id: str`
- `marked_answers: dict[str, list[str]]`
- `omr_error: str`
- `annotated_pdf: str`

## JSON Output

### `omr-grade` Single PDF

```json
{
  "qr_data": {
    "examSetId": "f6adcc63-71dc-412c-9c8d-a4609df454ff",
    "variantId": "37e3d65f-e540-4e34-b438-549e731be3b0"
  },
  "student_id": "63620147",
  "marked_answers": {
    "1": ["D"],
    "2": ["C"],
    "3": ["B", "C"],
    "4": ["C"],
    "5": ["A"],
    "6": ["E"]
  },
  "omr_error": ""
}
```

### `omr-grade` Directory

```json
[
  {
    "source_pdf": "student-a.pdf",
    "qr_data": {
      "examSetId": "f6adcc63-71dc-412c-9c8d-a4609df454ff",
      "variantId": "37e3d65f-e540-4e34-b438-549e731be3b0"
    },
    "student_id": "63620147",
    "marked_answers": {
      "1": ["D"]
    },
    "omr_error": ""
  },
  {
    "source_pdf": "student-b.pdf",
    "qr_data": null,
    "student_id": "",
    "marked_answers": {},
    "omr_error": "filled-sheets/student-b.pdf: Required alignment marker 'top_right' was not detected"
  }
]
```

### `omr-annotate` Single PDF

```json
{
  "qr_data": {
    "examSetId": "f6adcc63-71dc-412c-9c8d-a4609df454ff",
    "variantId": "37e3d65f-e540-4e34-b438-549e731be3b0"
  },
  "student_id": "63620147",
  "marked_answers": {
    "1": ["D"]
  },
  "omr_error": "",
  "annotated_pdf": "/tmp/annotated.pdf"
}
```

### `marked_answers` Shape

```json
{
  "1": ["B"],
  "2": ["B", "D"],
  "3": []
}
```

Rules:

- keys are question numbers as strings
- values are arrays of selected option labels
- multiple marked choices are allowed

## Correct Answer JSON

Accepted formats:

```json
{
  "1": ["D"],
  "2": ["B", "D"],
  "3": ["A"]
}
```

or:

```json
{
  "answers": {
    "1": ["D"],
    "2": ["B", "D"]
  }
}
```

Rules:

- question keys may be strings or numbers
- values may be a string or a list of strings
- labels are normalized to uppercase
- supported labels are `A` through `E`
- answer keys affect annotation only, not grading

## Practical Examples

### Full Flow

```bash
uv run omr \
  --questions 100 \
  --choices 5 \
  --exam-set-id exam-set-001 \
  --variant-id variant-a \
  --output sheet.pdf

uv run omr-grade sheet-filled.pdf > result.json

uv run omr-annotate \
  sheet-filled.pdf \
  --output sheet-filled-annotated.pdf \
  --correct-answers tests/answer-key.json
```

### Batch Grade

```bash
uv run omr-grade scans/ > results.json
```

### Batch Annotate

```bash
uv run omr-annotate scans/ --output reviewed/
```

## Limitations

- grading requires the current marker-based sheet format
- question count is capped at `100`
- student ID is fixed to 8 digits
- student ID columns must have exactly one marked digit each
- option labels are limited to `A-E`
- geometric registration is marker-first
- QR is used for payload extraction, not for primary alignment

## Tests

```bash
uv run --with pytest pytest -q
```

Tests generate PDF fixtures dynamically and clean them up automatically.
