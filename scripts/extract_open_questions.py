#!/usr/bin/env python3
"""Extract free-text (open-ended) question columns from a VJC survey export.

Usage:
    python3 extract_open_questions.py <input.xlsx> [output.xlsx]

Reads the "Worksheet" sheet of a raw VJC survey export (two header rows:
question text, then field code, followed by response rows), drops every
closed-choice column (Yes/No, rating scales, agree/disagree, select-from-list),
and drops redundant "-textfield" duplicate columns that never carry data the
primary column lacks. The result keeps only genuine open-ended text answers
plus the cohort_id/student_id/uploaded_time identifiers.
"""

import sys

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

SOURCE_SHEET = "Worksheet"
OUTPUT_SHEET = "Open Questions"

# code -> readable label, in the order they should appear in the output.
# Columns not listed here (closed-choice questions, and "-textfield" columns
# that are pure duplicates of their primary column) are dropped.
KEEP_COLUMNS = {
    "cohort_id": "cohort_id",
    "student_id": "student_id",
    "uploaded_time": "uploaded_time",
    "v4-general-4": "What did you enjoy most about being a VJC leader?",
    "v4-general-5": "Was there any part of being a Virtual Justice Club leader you disliked and if so what part?",
    "v4-general-9": "Are there any other human rights-related topics you would like to learn more about?",
    "v4-reporting-1-textfield": "Since you became a VJC leader, have you helped anyone with a defilement situation? (details)",
    "v4-reporting-1-supplemental": "Since you became a VJC leader, have you helped anyone with a defilement situation? (additional details)",
    "v3-technology-2": "What were the things you disliked about having the device?",
    "v4-conclusion-1": "Is there anything else you would like to tell us about your experience with the Virtual Justice Club?",
    "v4-newsletter-1": "Please tell us all the things you liked about delivering the 160 Girls Justice Journal in your community.",
    "v4-newsletter-2": "Please tell us all the things you disliked about delivering the 160 Girls Justice Journal in your community.",
    "v4-newsletter-5": "Please describe both positive and negative feedback you received about the 160 Girls Justice Journal.",
    "v4-technology-1": "What were the things you liked about having the device?",
    "v4-technology-4": "Is there anything about the device training that could be improved?",
}

IDENTIFIER_CODES = {"cohort_id", "student_id", "uploaded_time"}


def extract(input_path, output_path):
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb[SOURCE_SHEET]
    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    codes = rows[1]
    data = rows[2:]

    keep_idx = sorted(i for i, c in enumerate(codes) if c in KEEP_COLUMNS)
    out_codes = [codes[i] for i in keep_idx]
    out_labels = [KEEP_COLUMNS[c] for c in out_codes]

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = OUTPUT_SHEET

    ws2.append(out_labels)
    ws2.append(out_codes)
    for r in data:
        ws2.append([r[i] for i in keep_idx])

    bold = Font(name="Arial", bold=True)
    regular = Font(name="Arial")
    align = Alignment(wrap_text=True, vertical="top")

    for row in ws2.iter_rows(min_row=1, max_row=2):
        for cell in row:
            cell.font = bold
            cell.alignment = align
    for row in ws2.iter_rows(min_row=3):
        for cell in row:
            cell.font = regular
            cell.alignment = align

    for i, code in enumerate(out_codes, start=1):
        letter = get_column_letter(i)
        ws2.column_dimensions[letter].width = 16 if code in IDENTIFIER_CODES else 45

    ws2.freeze_panes = "A3"
    ws2.row_dimensions[1].height = 45

    wb2.save(output_path)
    return out_codes, len(data)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "Post2025OpenQuestions.xlsx"

    out_codes, n_rows = extract(input_path, output_path)
    print(f"Saved {output_path}")
    print(f"Columns ({len(out_codes)}): {out_codes}")
    print(f"Rows: {n_rows}")


if __name__ == "__main__":
    main()
