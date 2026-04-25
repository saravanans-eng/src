from pathlib import Path
import sys
# ensure repo root is on sys.path so imports work when running the script directly
repo_root = Path(__file__).resolve().parents[1]
if str(repo_root) not in sys.path:
    sys.path.insert(0, str(repo_root))

from doc_compare import fallback_comparator

def main():
    orig = "Hello\nworld\nThis is a demo document."
    rev = "Hello\neveryone\nThis is a demo document!"
    out = Path(__file__).resolve().parent / "demo_out.html"
    fallback_comparator.compare_texts(orig, rev, str(out))
    print("Wrote demo HTML to:", out)

if __name__ == '__main__':
    main()
