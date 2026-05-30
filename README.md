# CAC Status Report Extractor

A CLI tool that batch-processes CAC (Corporate Affairs Commission) Status Report PDFs, extracts key business and proprietor fields via regex, and writes the results to a consolidated Excel file.

---

## Requirements

- Python 3.10+
- [uv](https://docs.astral.sh/uv/) package manager

---

## Installation

**1. Clone the repo and navigate into it:**

```bash
git clone <your-repo-url>
cd cac-extractor
```

**2. Create a virtual environment and install dependencies:**

```bash
uv venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
uv pip install pdfplumber pandas openpyxl
```

Or if you have a `pyproject.toml`:

```bash
uv sync
```

---

## Usage

**1. Set the input folder** in the script (or via env):

```python
PDF_FOLDER = r"/path/to/your/pdfs/"
```

**2. Run the extractor:**

```bash
uv run main.py
```

The script will:
- Scan `PDF_FOLDER` for PDF files
- Skip files that don't contain `"Status Report"` in the filename
- Skip files already listed in `processed_files.txt`
- Extract fields and append each row to an Excel file
- Run a post-processing cleanup pass on the output

---

## Output

| File | Description |
|---|---|
| `CAC_REGISTERED:<timestamp>.xlsx` | Extracted data, one row per PDF |
| `processed_files.txt` | Log of already-processed filenames (prevents duplicates) |

---

## Extracted Fields

### Business (Page 1)

| Field | Description |
|---|---|
| `Business Name` | Registered business name |
| `BN Number` | Business registration number |
| `Date of Registration` | Registration date |
| `Business Type` | e.g. Sole Proprietorship |
| `Business Email` | Business contact email |
| `Principal Place of Business` | Address (handles line-wrapped layout) |
| `Principal Business Activity` | Nature of business |

### Proprietor (Page 2)

| Field | Description |
|---|---|
| `SURNAME` | Proprietor surname |
| `FIRSTNAME` | Proprietor first name |
| `OTHER NAME` | Middle name / other names |
| `EMAIL` | Personal email |
| `PHONE NUMBER` | Contact number |
| `GENDER` | Gender |
| `DATE OF BIRTH` | Date of birth |
| `STATUS` | `ACTIVE`, `INACTIVE`, or `STRUCK OFF` |

> **Note:** Fields with a value of `NIL` in the PDF are normalised to empty cells in Excel.

---

## Project Structure

```
cac-extractor/
├── main.py               # Entry point — runs batch extraction and cleanup
├── processed_files.txt   # Auto-generated: tracks processed PDFs
└── README.md
```

---

## Configuration

Edit these constants at the top of `main.py`:

| Constant | Default | Description |
|---|---|---|
| `PDF_FOLDER` | `/home/spot/Downloads/Queried recovered/` | Folder containing input PDFs |
| `OUTPUT_FILE` | `./CAC_REGISTERED:<timestamp>.xlsx` | Output Excel path |
| `PROCESSED_LOG` | `./processed_files.txt` | Duplicate-prevention log |

---

## Notes

- Only files with `"Status Report"` in the filename are processed; certificate PDFs are skipped automatically.
- The `Principal Place of Business` field uses a multiline regex to handle addresses that wrap across lines in the PDF layout.
- Re-running the script is safe — already-processed files are skipped via `processed_files.txt`. Delete that file to reprocess everything from scratch.
