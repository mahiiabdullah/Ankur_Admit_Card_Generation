"""
Generate 2026 admit cards (DOCX + PDF) from Admit_Card_Data_2026.xlsx.

1. Builds Admit_Template_2026.docx from 'Demo Admit.docx' by swapping the sample
   student's values for {{placeholders}} (the demo file is never modified).
   Bangla text in the value fields uses the Kalpurush font; English stays
   Times New Roman 12pt as in the demo.
2. Fills the template once per row of the Admit_Data sheet. Replacement is done
   directly in the document XML, so formatting is preserved exactly.
3. Converts every DOCX to PDF through the Word-compatible COM server
   (WPS Office on this PC), restarting it every batch to stay stable.

Output: Admit_Cards_2026/DOCX/<File Name>.docx and Admit_Cards_2026/PDF/<File Name>.pdf
"""
import argparse
import os
import re
import shutil
import sys
import time
import zipfile
from xml.sax.saxutils import escape

import pandas as pd

DEMO = "Demo Admit.docx"
TEMPLATE = "Admit_Template_2026.docx"

# Sample value in the demo card -> placeholder (each must occur exactly once)
DEMO_VALUES = {
    "Navila Tahsin Tasha": "{{Name (English)}}",
    "Md. Abdur Rahim": "{{Fathers Name}}",
    "Purba bakalia city corporation girls high school and college": "{{Name of School}}",
    "10": "{{Class}}",
    "+8801825067781": "{{Mobile (Online)}}",
    "11225031": "{{Exam Roll}}",
    "Chattogram Collegiate School": "{{Exam Centre}}",
}
BANGLA_FONT = "Kalpurush"
BANGLA_RE = re.compile(r"[ঀ-৿]")


def build_template(demo: str, template: str) -> None:
    with zipfile.ZipFile(demo) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        for value, placeholder in DEMO_VALUES.items():
            needle = f"<w:t>{value}</w:t>"
            count = xml.count(needle)
            if count != 1:
                sys.exit(f"Template build failed: '{value}' found {count} times in {demo}")
            pos = xml.index(needle)
            # Give this run a Bangla-capable complex-script font
            rpr_start = xml.rindex("<w:rPr>", 0, pos)
            rpr = xml[rpr_start:pos].replace(
                'w:cs="Times New Roman"', f'w:cs="{BANGLA_FONT}"')
            xml = xml[:rpr_start] + rpr + f"<w:t xml:space=\"preserve\">{placeholder}</w:t>" + xml[pos + len(needle):]
        write_docx(z, template, xml)


def write_docx(src_zip: zipfile.ZipFile, path: str, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src_zip.infolist():
            data = document_xml.encode("utf-8") if item.filename == "word/document.xml" else src_zip.read(item)
            # Fresh ZipInfo: writestr() mutates the one it is given, which would break src_zip
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            out.writestr(info, data)


def fill_cards(template: str, data: pd.DataFrame, docx_dir: str) -> list:
    os.makedirs(docx_dir, exist_ok=True)
    paths = []
    with zipfile.ZipFile(template) as z:
        base = z.read("word/document.xml").decode("utf-8")
        for _, row in data.iterrows():
            xml = base
            for placeholder in DEMO_VALUES.values():
                field = placeholder[2:-2]
                value = "" if pd.isna(row[field]) else str(row[field])
                if BANGLA_RE.search(value):
                    # Mark the run as complex script so Bangla conjuncts shape correctly
                    pos = xml.index(placeholder)
                    end = xml.rindex("</w:rPr>", 0, pos)
                    xml = xml[:end] + '<w:cs/><w:lang w:bidi="bn-IN"/>' + xml[end:]
                xml = xml.replace(placeholder, escape(value))
            if "{{" in xml:
                sys.exit(f"Unfilled placeholder left for order {row['Order ID']}")
            path = os.path.abspath(os.path.join(docx_dir, f"{row['File Name']}.docx"))
            write_docx(z, path, xml)
            paths.append(path)
    return paths


class _Converter:
    """Word-compatible COM server that is restarted whenever it dies."""

    def __init__(self):
        import pythoncom
        pythoncom.CoInitialize()
        self.app = None

    def _start(self):
        import win32com.client
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = 0

    def stop(self):
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None
            time.sleep(1)

    def convert(self, docx: str, pdf: str) -> None:
        if self.app is None:
            self._start()
        doc = self.app.Documents.Open(docx, False, True)  # ConfirmConversions, ReadOnly
        try:
            doc.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF
        finally:
            try:
                doc.Close(False)
            except Exception:
                pass


def convert_to_pdf(docx_paths: list, pdf_dir: str, restart_every: int = 20, attempts: int = 3) -> list:
    os.makedirs(pdf_dir, exist_ok=True)
    conv = _Converter()
    failed = []
    total = len(docx_paths)
    try:
        for i, docx in enumerate(docx_paths, 1):
            pdf = os.path.abspath(os.path.join(pdf_dir, os.path.basename(docx)[:-5] + ".pdf"))
            for attempt in range(1, attempts + 1):
                try:
                    conv.convert(docx, pdf)
                    if not os.path.isfile(pdf) or os.path.getsize(pdf) == 0:
                        raise RuntimeError("PDF not written")
                    print(f"  [{i}/{total}] {os.path.basename(pdf)}")
                    break
                except Exception as e:
                    conv.stop()  # server likely crashed - start a fresh one
                    if attempt == attempts:
                        failed.append((docx, str(e)))
                        print(f"  [{i}/{total}] FAILED {os.path.basename(docx)}: {e}")
            if i % restart_every == 0:
                conv.stop()
    finally:
        conv.stop()
    return failed


def main(xlsx: str, out_dir: str, only: list, skip_pdf: bool) -> None:
    data = pd.read_excel(xlsx, sheet_name="Admit_Data", dtype=str)
    missing = data[data["Exam Roll"].isna()]
    if len(missing):
        sys.exit(f"{len(missing)} rows have no Exam Roll - run assign_rolls_2026.py first")
    if only:
        data = data[data["Order ID"].isin(only) | data["Exam Roll"].isin(only)]

    build_template(DEMO, TEMPLATE)
    print(f"Template: {TEMPLATE}")

    docx_dir = os.path.join(out_dir, "DOCX")
    pdf_dir = os.path.join(out_dir, "PDF")
    if not only:
        for d in (docx_dir, pdf_dir):
            shutil.rmtree(d, ignore_errors=True)

    paths = fill_cards(TEMPLATE, data, docx_dir)
    print(f"DOCX created: {len(paths)}  ->  {docx_dir}")

    if skip_pdf:
        return
    print("Converting to PDF...")
    failed = convert_to_pdf(paths, pdf_dir)
    print(f"\nPDF created: {len(paths) - len(failed)}/{len(paths)}  ->  {pdf_dir}")
    if failed:
        print("Failed:")
        for f, e in failed:
            print(f"  {os.path.basename(f)}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--xlsx", default="Admit_Card_Data_2026.xlsx")
    p.add_argument("--out", default="Admit_Cards_2026")
    p.add_argument("--only", nargs="*", default=[], help="Order IDs or Exam Rolls to (re)generate")
    p.add_argument("--skip-pdf", action="store_true", help="Create DOCX files only")
    a = p.parse_args()
    main(a.xlsx, a.out, a.only, a.skip_pdf)
