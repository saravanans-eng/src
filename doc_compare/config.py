import os

# Default directories (change as needed)
DEFAULT_WATCH_DIR = r"D:\WORK\ELSEVIER\FOR-S100ACTXML"
REPORT_DIR = os.path.join(DEFAULT_WATCH_DIR, "reports")
UNC_BASE = r"\\tnqfs01\CUTOOL\ELSEVIER"

# Filename patterns
JID_REGEX = r"^[A-Za-z]+$"
AID_REGEX = r"^\d{8}$"

# Word SaveAs format constants (used by pywin32 when available)
WD_FORMAT_PDF = 17
WD_FORMAT_HTML = 8
