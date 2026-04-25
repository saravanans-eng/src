import os
import time
from typing import Dict, Tuple, Any

try:
    import win32com.client
except Exception:
    win32com = None

from . import config
from . import fallback_comparator


def _extract_revisions_from_docx(path: str) -> Dict[str, Any]:
    """Open a Word document and extract insertion/deletion texts from Revisions.

    Returns a summary dict: {"insertions": [...], "deletions": [...], "counts": {"ins": n, "del": m}}
    """
    summary = {"insertions": [], "deletions": [], "counts": {"ins": 0, "del": 0}}
    if win32com is None:
        return summary

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(path, ReadOnly=True)
        # WD constants: 1 = wdRevisionInsert, 2 = wdRevisionDelete
        for rev in doc.Revisions:
            try:
                t = rev.Type
            except Exception:
                t = None
            text = getattr(rev, "Range").Text if hasattr(rev, "Range") else ""
            if t == 1:
                summary["insertions"].append(text)
                summary["counts"]["ins"] += 1
            elif t == 2:
                summary["deletions"].append(text)
                summary["counts"]["del"] += 1
            else:
                # unknown revision type: ignore
                pass
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass

    return summary


def compare_documents(original_path: str, revised_path: str, out_dir: str, prefix: str = "comparison") -> Dict[str, str]:
    """Compare two DOCX files using Microsoft Word COM if available.

    Returns a dict with keys: docx, html, pdf, details_html and a revision summary in 'summary'.
    """
    os.makedirs(out_dir, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    base_name = f"{prefix}_{ts}"
    out_docx = os.path.join(out_dir, base_name + ".docx")
    out_html = os.path.join(out_dir, base_name + ".html")
    out_pdf = os.path.join(out_dir, base_name + ".pdf")
    details_html = os.path.join(out_dir, base_name + "_details.html")

    # If Word COM is not available, fall back to LibreOffice-based comparator
    if win32com is None:
        fb = fallback_comparator.compare_with_libreoffice(original_path, revised_path, out_dir, prefix=prefix)
        return {"docx": None, "html": fb.get("html"), "pdf": fb.get("pdf"), "details": fb.get("html"), "summary": {"note": "fallback_libreoffice"}}

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    # Open documents (read-only)
    doc_orig = word.Documents.Open(original_path, ReadOnly=True)
    doc_rev = word.Documents.Open(revised_path, ReadOnly=True)

    # Perform compare: call Compare on revised document against original.
    # This creates revisions in the revised document and sets it active.
    try:
        doc_rev.Compare(Name=original_path)
        compared = word.ActiveDocument
        # Save comparison as docx
        compared.SaveAs(out_docx)
        # Save HTML and PDF
        compared.SaveAs(out_html, FileFormat=config.WD_FORMAT_HTML)
        compared.SaveAs(out_pdf, FileFormat=config.WD_FORMAT_PDF)

        # Extract revisions by operating on the saved comparison file
        rev_summary = _extract_revisions_from_docx(out_docx)

        # build a simple details HTML
        with open(details_html, "w", encoding="utf-8") as f:
            f.write("<html><head><meta charset='utf-8'><title>Comparison Details</title></head><body>")
            f.write(f"<h1>Comparison details for {os.path.basename(revised_path)}</h1>")
            f.write(f"<p>Inserted: {rev_summary['counts']['ins']}, Deleted: {rev_summary['counts']['del']}</p>")
            f.write("<h2>Insertions</h2><ul>")
            for it in rev_summary["insertions"]:
                f.write(f"<li>{(it or '').replace('<','&lt;').replace('>','&gt;')}</li>")
            f.write("</ul><h2>Deletions</h2><ul>")
            for it in rev_summary["deletions"]:
                f.write(f"<li>{(it or '').replace('<','&lt;').replace('>','&gt;')}</li>")
            f.write("</ul>")
            f.write(f"<p><a href=\"{os.path.basename(out_docx)}\">DOCX</a> | <a href=\"{os.path.basename(out_html)}\">HTML</a> | <a href=\"{os.path.basename(out_pdf)}\">PDF</a></p>")
            f.write("</body></html>")

    except Exception:
        # If Word compare fails unexpectedly, fall back to LibreOffice-based comparator
        try:
            fb = fallback_comparator.compare_with_libreoffice(original_path, revised_path, out_dir, prefix=prefix)
            return {"docx": None, "html": fb.get("html"), "pdf": fb.get("pdf"), "details": fb.get("html"), "summary": {"note": "fallback_libreoffice_after_error"}}
        except Exception:
            raise
    finally:
        # cleanup
        try:
            doc_orig.Close(False)
        except Exception:
            pass
        try:
            doc_rev.Close(False)
        except Exception:
            pass
        try:
            # close compared if still open
            if 'compared' in locals():
                compared.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass

    return {"docx": out_docx, "html": out_html, "pdf": out_pdf, "details": details_html, "summary": rev_summary}
