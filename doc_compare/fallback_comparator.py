import os
import subprocess
import tempfile
import shutil
import datetime
from difflib import HtmlDiff
from typing import Optional, Dict


def _soffice_convert_to_txt(input_path: str, output_dir: str) -> str:
    cmd = ["soffice", "--headless", "--convert-to", "txt:Text", "--outdir", output_dir, input_path]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"soffice conversion failed: {res.stderr.strip()}")
    base = os.path.splitext(os.path.basename(input_path))[0] + ".txt"
    return os.path.join(output_dir, base)


def compare_texts(orig_text: str, rev_text: str, out_html_path: str) -> str:
    """Create an HTML side-by-side diff of two text strings and write to out_html_path."""
    diff = HtmlDiff(wrapcolumn=80)
    html = diff.make_file(orig_text.splitlines(), rev_text.splitlines(), fromdesc="Original", todesc="Revised")
    os.makedirs(os.path.dirname(out_html_path) or '.', exist_ok=True)
    with open(out_html_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_html_path


def compare_with_libreoffice(original_path: str, revised_path: str, out_dir: str, prefix: str = "libre_cmp") -> Dict[str, Optional[str]]:
    """Fallback comparator that uses LibreOffice to convert documents to text and generates an HTML diff.

    Returns a dict with keys: html, pdf (if created), timestamp.
    """
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    tmpdir = tempfile.mkdtemp()
    try:
        orig_txt = _soffice_convert_to_txt(original_path, tmpdir)
        rev_txt = _soffice_convert_to_txt(revised_path, tmpdir)
        with open(orig_txt, "r", encoding="utf-8", errors="ignore") as f:
            orig_text = f.read()
        with open(rev_txt, "r", encoding="utf-8", errors="ignore") as f:
            rev_text = f.read()

        out_html = os.path.join(out_dir, f"{prefix}_{ts}.html")
        compare_texts(orig_text, rev_text, out_html)

        pdf_path = None
        # attempt to export revised doc to PDF for artifact parity
        try:
            cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, revised_path]
            subprocess.run(cmd, capture_output=True, text=True, check=True)
            pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(revised_path))[0] + ".pdf")
            if not os.path.exists(pdf_path):
                pdf_path = None
        except Exception:
            pdf_path = None

        return {"html": out_html, "pdf": pdf_path, "timestamp": ts}
    finally:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass


if __name__ == "__main__":
    print("This module provides a fallback comparator for environments without MS Word.")
