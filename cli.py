import os
import sys
import argparse
import time
from datetime import datetime

from src.doc_compare import parser as filename_parser
from src.doc_compare import comparator, report, config
from src.doc_compare.watcher import start_watching


def find_original(jid: str, aid: str) -> str:
    # UNC: \\tnqfs01\CUTOOL\ELSEVIER\<JID>\<AID>\TUD_Output\<JID>_<AID>_tud.docx
    name = f"{jid}_{aid}_tud.docx"
    return os.path.join(config.UNC_BASE, jid, aid, "TUD_Output", name)


def process_file(path: str):
    print("Processing:", path)
    parsed = filename_parser.parse_filename(path)
    if not parsed:
        print("Filename not recognized, skipping:", path)
        return
    jid, aid = parsed
    original = find_original(jid, aid)
    if not os.path.exists(original):
        print("Original not found:", original)
        # write entry to report indicating missing original
        ts = datetime.now().isoformat()
        report.append_report(config.REPORT_DIR, jid, aid, ts, {}, "Original not found")
        return

    out_dir = os.path.join(config.REPORT_DIR, jid, aid)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        artifacts = comparator.compare_documents(original, path, out_dir)
        summary = "Comparison generated"
    except Exception as e:
        artifacts = {}
        summary = f"Comparison failed: {e}"

    report.append_report(config.REPORT_DIR, jid, aid, ts, artifacts, summary)
    print("Done. Artifacts:", artifacts)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--watch", action="store_true", help="Watch configured folder for new files")
    ap.add_argument("--once", help="Process a single file path and exit")
    args = ap.parse_args()
    if args.once:
        process_file(args.once)
        return
    if args.watch:
        print("Watching", config.DEFAULT_WATCH_DIR)
        start_watching(config.DEFAULT_WATCH_DIR, process_file)
        return
    ap.print_help()


if __name__ == "__main__":
    main()
