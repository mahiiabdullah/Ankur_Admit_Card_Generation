import argparse
import sys
import pandas as pd
from docx import Document
try:
    from docx2pdf import convert
    _HAS_DOCX2PDF = True
except Exception:
    _HAS_DOCX2PDF = False
import os


def replace_placeholders_in_doc(doc: Document, data_row: pd.Series) -> None:
    """Replace all placeholders in a docx Document with values from data_row.

    Placeholders are expected in the form {{ColumnName}}.
    Performs case-insensitive matching and formats text as Times New Roman, not bold.
    """
    import re
    from docx.shared import Pt
    
    # Create a mapping of lowercase placeholder to actual value
    placeholder_map = {}
    for key, value in data_row.items():
        placeholder_map[f"{{{{{key}}}}}".lower()] = str(value)
    
    # Replace in paragraphs
    for p in doc.paragraphs:
        for run in p.runs:
            original_text = run.text
            for placeholder_lower, value in placeholder_map.items():
                # Case-insensitive replacement
                pattern = re.compile(re.escape(placeholder_lower), re.IGNORECASE)
                if pattern.search(original_text):
                    run.text = pattern.sub(value, original_text)
                    # Format: Times New Roman, not bold
                    run.font.name = 'Times New Roman'
                    run.font.bold = False
                    run.font.size = Pt(12)  # Default size

    # Replace in tables
    for table in doc.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        original_text = run.text
                        for placeholder_lower, value in placeholder_map.items():
                            # Case-insensitive replacement
                            pattern = re.compile(re.escape(placeholder_lower), re.IGNORECASE)
                            if pattern.search(original_text):
                                run.text = pattern.sub(value, original_text)
                                # Format: Times New Roman, not bold
                                run.font.name = 'Times New Roman'
                                run.font.bold = False
                                run.font.size = Pt(12)  # Default size


def main(template: str, data_file: str, output_dir: str):
    # Normalize and check paths
    template = os.path.expanduser(template)
    data_file = os.path.expanduser(data_file)
    output_dir = os.path.expanduser(output_dir)

    if not os.path.isfile(template):
        print(f"Template file not found: {template}")
        sys.exit(2)
    if not os.path.isfile(data_file):
        print(f"Data file not found: {data_file}")
        sys.exit(2)

    os.makedirs(output_dir, exist_ok=True)

    # Read Excel file with header in row 2 (index 1)
    df = pd.read_excel(data_file, header=1)

    for idx, row in df.iterrows():
        doc = Document(template)
        replace_placeholders_in_doc(doc, row)

        roll = str(row.get("Exam Roll", idx + 1))
        safe_roll = "".join(c for c in roll if c.isalnum() or c in ("-", "_")) or str(idx + 1)
        output_docx = os.path.join(output_dir, f"Admit_{safe_roll}.docx")
        doc.save(output_docx)

    if _HAS_DOCX2PDF:
        try:
            convert(output_dir)
        except Exception as e:
            print(f"Warning: failed to convert docx to pdf: {e}")
    else:
        print("docx2pdf not available; skipping conversion to PDF.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate admit cards from an Excel file and a Word template.")
    parser.add_argument("--template", default=r"C:\Users\Lenovo\Downloads\Telegram Desktop\ASP Admit Card Soft Copy (5).docx",
                        help="Path to the .docx template")
    parser.add_argument("--data", default=r"C:\Users\Lenovo\Downloads\Telegram Desktop\ASP Online Reg. Boy (2).xlsx",
                        help="Path to the Excel file with data")
    parser.add_argument("--out", default="Generated_Admit_Cards", help="Output directory")
    args = parser.parse_args()

    main(args.template, args.data, args.out)
