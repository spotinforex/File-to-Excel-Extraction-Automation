# =========================
# PDF Folder Auto-Extractor
# =========================
# What it does:
# 2. Opens each PDF
# 3. Extracts data using regex
# 4. Saves/appends results into Excel
# 5. Skips already processed files

import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime

# =========================
# CONFIG
# =========================
PDF_FOLDER = r"/home/spot/Downloads/Cac_cert/"       # Folder to watch
OUTPUT_FILE = rf"./CAC REGISTERED:{datetime.now()}.xlsx"      # Excel output file
PROCESSED_LOG = r"./processed_files.txt"             # Tracks processed PDFs

# =========================
# REGEX PATTERNS
# --------------------------
# The PDF layout has labels and values on the SAME line, separated by spaces.
# Pattern format:  LABEL<spaces>VALUE<end-of-line>
# re.MULTILINE is used so $ matches end of each line.
# =========================
PATTERNS = {
    # Page 2 — Proprietor fields
    "SURNAME":       r"^SURNAME\s+(.+)$",
    "FIRSTNAME":     r"^FIRSTNAME\s+(.+)$",
    "OTHER NAME":    r"^OTHER NAME\s+(.+)$",
    "EMAIL":         r"^EMAIL\s+(.+)$",
    "PHONE NUMBER":  r"^PHONE NUMBER\s+(.+)$",
    "GENDER":        r"^GENDER\s+(.+)$",
    "DATE OF BIRTH": r"^DATE OF BIRTH\s+(.+)$",
    "STATUS":        r"^STATUS\s+(ACTIVE|INACTIVE|STRUCK OFF)$",

    # Page 1 — Business fields
    # "Business Name DETAILS" is the section header — require value not equal to "DETAILS"
    "Business Name":               r"^Business Name\s+(?!DETAILS$)(.+)$",
    "BN Number":                   r"^BN Number\s+(.+)$",
    "Date of Registration":        r"^Date of Registration\s+(.+)$",
    "Business Type":               r"^Business Type\s+(.+)$",
    "Business Email":              r"^Email\s+(.+)$",
    # Address wraps: first part is on the line BEFORE the label, second part on line AFTER
    "Principal Place of Business": r"(.+)\nPrincipal Place of Business\s*\n(.+)",
    "Principal Business Activity": r"^Principal Business Activity\s+(.+)$",
}

# =========================
# LOAD PROCESSED FILES
# =========================
def load_processed_files():
    if os.path.exists(PROCESSED_LOG):
        with open(PROCESSED_LOG, "r") as f:
            return set(line.strip() for line in f)
    return set()

def save_processed_file(filename):
    with open(PROCESSED_LOG, "a") as f:
        f.write(filename + "\n")

# =========================
# EXTRACT TEXT FROM PDF
# =========================
def extract_pdf_text(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return text

# =========================
# APPLY REGEX EXTRACTION
# =========================
def extract_data(text):
    extracted = {}
    for field, pattern in PATTERNS.items():
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            # Principal Place of Business has 2 groups (wrapped address)
            if match.lastindex and match.lastindex >= 2:
                value = " ".join(g.strip() for g in match.groups() if g)
            else:
                value = match.group(1).strip()
        else:
            value = None
        # Normalise "NIL" to actual None so Excel cells are empty
        extracted[field] = None if value and value.upper() == "NIL" else value
    return extracted

# =========================
# APPEND TO EXCEL
# =========================
def append_to_excel(data):
    new_df = pd.DataFrame([data])

    if os.path.exists(OUTPUT_FILE):
        existing_df = pd.read_excel(OUTPUT_FILE)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    combined_df.to_excel(OUTPUT_FILE, index=False)

# =========================
# CHECKS IF STATUS REPORT
# =========================
def checker(filename):
    return "Status Report" in filename

# =========================
# PROCESS PDF
# =========================
def process_pdf(pdf_path):
    filename = os.path.basename(pdf_path)

    processed_files = load_processed_files()
    if filename in processed_files:
        print(f"Skipping already processed: {filename}")
        return

    if not checker(filename):
        print(f"Skipping Certificate: {filename}")
        return

    print(f"Processing: {filename}")

    text = extract_pdf_text(pdf_path)
    if not text.strip():
        print(f"No text found in {filename}")
        return

    data = extract_data(text)
    data["Source File"] = filename

    append_to_excel(data)
    save_processed_file(filename)

    print(f"Done: {filename}")

# =========================
# INITIAL BACKLOG PROCESSING
# =========================
def process_existing_pdfs():
    for file in os.listdir(PDF_FOLDER):
        if file.lower().endswith(".pdf"):
            process_pdf(os.path.join(PDF_FOLDER, file))

# =========================
# MAIN
# =========================
if __name__ == "__main__":
    os.makedirs(PDF_FOLDER, exist_ok=True)

    print("Checking existing PDFs...")
    process_existing_pdfs()
    print("Process Completed")
