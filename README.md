# DOC Compare

Automates DOCX comparisons: when revised files are placed in the watch folder, the service finds the original from the UNC path, runs a comparison (Microsoft Word COM when available; LibreOffice fallback otherwise), and stores timestamped comparison artifacts (DOCX/HTML/PDF) with a central HTML index.

Prerequisites
- Windows with Microsoft Word installed (recommended for best fidelity)
- Python 3.10+
- Install Python requirements: `pip install -r requirements.txt`
- Optional: LibreOffice (`soffice`) in PATH for fallback comparator

Local run (once)
1. Process a single file:
```powershell
python cli.py --once "C:\path\to\JAC_00033182_tud_ACE_For_S100_Conversion.docx"
```

Run as watcher (interactive)
```powershell
python cli.py --watch
```

Install as Windows service (recommended: run PowerShell as Administrator)
1. Use the included NSSM helper which downloads NSSM and registers the service:
```powershell
.\scripts\install_service.ps1 -ServiceName DocCompareService -PythonExe "C:\Users\3874\AppData\Local\Programs\Python\Python313\python.exe" -Module "src.doc_compare.service"
```
2. If you prefer `sc` or manual registration, register the Python module as the service host (requires Admin):
```powershell
sc create "DocCompareService" binPath= "\"C:\Users\3874\AppData\Local\Programs\Python\Python313\python.exe\" -m src.doc_compare.service" start= auto DisplayName= "DocCompare Service"
sc start "DocCompareService"
```

Notes
- The service must run under an account that can access the UNC path `\\tnqfs01\CUTOOL\ELSEVIER`.
- Word COM automation may require an interactive desktop session in some server configurations. If Word COM fails, the fallback uses LibreOffice headless text conversion and a text diff (less accurate for tracked changes).

Configuration
- Edit `doc_compare/config.py` to set `DEFAULT_WATCH_DIR`, `REPORT_DIR`, and `UNC_BASE`.

Development
- Tests: `pytest -q`
