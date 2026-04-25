import re
import os
from typing import Optional, Tuple

FILENAME_RE = re.compile(r"^(?P<jid>[A-Za-z]+)_(?P<aid>\d{8})_tud_ACE_For_S100_Conversion\.docx$")


def parse_filename(path: str) -> Optional[Tuple[str, str]]:
    """Parse revised filename and return (jid, aid) or None if not matching.

    Example filename: JAC_00033182_tud_ACE_For_S100_Conversion.docx
    """
    base = os.path.basename(path)
    m = FILENAME_RE.match(base)
    if not m:
        return None
    return m.group("jid"), m.group("aid")


if __name__ == "__main__":
    # quick manual test
    print(parse_filename("JAC_00033182_tud_ACE_For_S100_Conversion.docx"))
