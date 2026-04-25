import os
from pathlib import Path

from doc_compare import fallback_comparator


def test_compare_texts_creates_html(tmp_path):
    orig = "Hello\nworld\nThis is a test"
    rev = "Hello\neveryone\nThis is a test!"
    out_html = tmp_path / "out.html"
    fallback_comparator.compare_texts(orig, rev, str(out_html))
    assert out_html.exists()
    content = out_html.read_text(encoding='utf-8')
    assert "everyone" in content
    assert "Hello" in content
