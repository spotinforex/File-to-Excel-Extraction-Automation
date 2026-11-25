import os
import re
import time
from datetime import datetime

import pandas as pd
import pdfplumber

# CONFIG
PDF_FOLDER = r"C:\Users\VOSTRO 3590\Documents\MCIPP\cac_downloads\status_reports"
OUTPUT_FILE = f'CAC_REGISTERED_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
PROCESSED_LOG = r"processed_files.txt"

# Layout settings (fractions/points of the PDF page)
SPLIT_RATIO = 0.33   # words starting left of this x (as a fraction of page width) are labels
MAX_DIST = 16        # max vertical distance (pt) between a value line and its label
LINE_TOL = 3         # words whose tops differ by <= this are on the same line

# LABELS (normalised to UPPERCASE) -> output column
BUSINESS_LABELS = {
    "BUSINESS NAME": "Business Name",
    "BN NUMBER": "BN Number",
    "DATE OF REGISTRATION": "Date of Registration",
    "BUSINESS TYPE": "Business Type",
    "PRINCIPAL PLACE OF BUSINESS": "Principal Place of Business",
    "BRANCH ADDRESS": "Branch Address",
    "POST CODE": "Post Code",
    "EMAIL": "Business Email",
    "STATUS": "Business Status",
    "PRINCIPAL BUSINESS ACTIVITY": "Principal Business Activity",
    "DUE DATE OF ACCOUNTS": "Due Date of Accounts",
}
PROPRIETOR_LABELS = {
    "ROLE TYPE": "ROLE TYPE",
    "SURNAME": "SURNAME",
    "FIRSTNAME": "FIRSTNAME",
    "OTHER NAME": "OTHER NAME",
    "EMAIL": "EMAIL",
    "PHONE NUMBER": "PHONE NUMBER",
    "NATIONALITY": "NATIONALITY",
    "SERVICE ADDRESS": "SERVICE ADDRESS",
    "RESIDENTIAL ADDRESS": "RESIDENTIAL ADDRESS",
    "GENDER": "GENDER",
    "DATE OF BIRTH": "DATE OF BIRTH",
    "STATUS": "STATUS",
}
PROPRIETOR_START = "ROLE TYPE"   # the business section ends where this label appears

COLUMNS = (
    ["Business Name", "BN Number", "Date of Registration", "Business Type",
     "Principal Place of Business", "Branch Address", "Post Code", "Business Email",
     "Business Status", "Principal Business Activity", "Due Date of Accounts"]
    + ["ROLE TYPE", "SURNAME", "FIRSTNAME", "OTHER NAME", "EMAIL", "PHONE NUMBER",
       "NATIONALITY", "SERVICE ADDRESS", "RESIDENTIAL ADDRESS", "GENDER",
       "DATE OF BIRTH", "STATUS"]
)
REQUIRED = ["Business Name", "Principal Place of Business", "SURNAME", "FIRSTNAME"]


# PROCESSED-FILE TRACKING
def load_processed_files():
    if os.path.exists(PROCESSED_LOG):
        with open(PROCESSED_LOG, "r") as f:
            return set(line.strip() for line in f if line.strip())
    return set()


def save_processed_files(filenames):
    with open(PROCESSED_LOG, "a") as f:
        for name in filenames:
            f.write(name + "\n")


# COORDINATE-BASED PARSING
def group_lines(words):
    """Cluster words into lines by vertical position. Returns lines top to bottom."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(w["top"] - lines[-1]["top"]) <= LINE_TOL:
            lines[-1]["words"].append(w)
        else:
            lines.append({"top": w["top"], "words": [w]})
    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
        line["y"] = sum((w["top"] + w["bottom"]) / 2 for w in line["words"]) / len(line["words"])
        line["text"] = " ".join(w["text"] for w in line["words"])
    return lines


def normalise_label(text):
    # drop numbering such as "1. 1. ROLE TYPE"
    return re.sub(r"^(?:\d+\.\s*)+", "", text).strip().upper()


def clean_value(value):
    value = re.sub(r"\s+", " ", value).strip()
    return None if not value or value.upper() == "NIL" else value


def extract_fields(pdf_path):
    fields = {}
    in_proprietor = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            split_x = page.width * SPLIT_RATIO
            left = [w for w in words if w["x0"] < split_x]
            right = [w for w in words if w["x0"] >= split_x]

            # 1) find label anchors in the left column, in reading order
            anchors = []
            for line in group_lines(left):
                key = normalise_label(line["text"])
                if key == PROPRIETOR_START:
                    in_proprietor = True
                table = PROPRIETOR_LABELS if in_proprietor else BUSINESS_LABELS
                if key in table:
                    anchors.append({"column": table[key], "y": line["y"], "words": []})

            if not anchors:
                continue

            # 2) give each right-column word to the vertically nearest label
            for w in right:
                wy = (w["top"] + w["bottom"]) / 2
                nearest = min(anchors, key=lambda a: abs(a["y"] - wy))
                if abs(nearest["y"] - wy) <= MAX_DIST:
                    nearest["words"].append(w)

            # 3) build values (multi-line values are joined top to bottom)
            for a in anchors:
                value = " ".join(line["text"] for line in group_lines(a["words"]))
                if a["column"] == "Principal Business Activity":
                    value = re.sub(r"^\d+\.\s*", "", value)   # "1. FASHION DESIGNING"
                fields.setdefault(a["column"], clean_value(value))  # first occurrence wins

    return fields


# PROCESS PDFs
def is_status_report(filename):
    return "STATUS REPORT" in filename.upper().replace("_", " ")


def process_pdf(pdf_path):
    filename = os.path.basename(pdf_path)
    try:
        data = extract_fields(pdf_path)
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return None

    if not data:
        print(f"No fields found in {filename}")
        return None

    missing = [c for c in REQUIRED if not data.get(c)]
    if missing:
        print(f"  WARNING {filename}: missing {', '.join(missing)}")

    row = {c: data.get(c) for c in COLUMNS}
    row["Source File"] = filename
    row["Check"] = ("MISSING: " + ", ".join(missing)) if missing else ""
    return row


def process_existing_pdfs():
    processed = load_processed_files()
    rows, done = [], []

    for file in sorted(os.listdir(PDF_FOLDER)):
        if not file.lower().endswith(".pdf"):
            continue
        if file in processed:
            print(f"Skipping already processed: {file}")
            continue
        if not is_status_report(file):
            print(f"Skipping Certificate: {file}")
            continue

        print(f"Processing: {file}")
        row = process_pdf(os.path.join(PDF_FOLDER, file))
        if row:
            rows.append(row)
            done.append(file)

    if not rows:
        return None

    # write once, then mark files processed (so a crash doesn't lose rows)
    pd.DataFrame(rows, columns=COLUMNS + ["Source File", "Check"]).to_excel(OUTPUT_FILE, index=False)
    save_processed_files(done)
    print(f"Total Processed: {len(rows)}")
    return OUTPUT_FILE


if __name__ == "__main__":
    start = time.time()
    os.makedirs(PDF_FOLDER, exist_ok=True)

    print("Checking existing PDFs...")
    out = process_existing_pdfs()
    if out:
        print(f"Saved: {out}")
    else:
        print("No Status Reports were processed.")
    print(f"Process Completed. Time Taken: {time.time() - start:.2f} seconds")