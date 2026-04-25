"""Windows service wrapper to run the watcher as a service.

This module uses pywin32's win32serviceutil.ServiceFramework. Installation and
running require Administrator privileges. This file is optional; if pywin32 is not
available the module will still import but cannot install the service.
"""
try:
    import win32serviceutil
    import win32service
    import win32event
except Exception:
    win32serviceutil = None

import sys
from src.doc_compare.watcher import start_watching
from src.doc_compare import config


class DocCompareService(win32serviceutil.ServiceFramework if win32serviceutil else object):
    if win32serviceutil:
        _svc_name_ = "DocCompareService"
        _svc_display_name_ = "DOCX Compare Service"

    def __init__(self, args):
        if win32serviceutil:
            win32serviceutil.ServiceFramework.__init__(self, args)
            self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        else:
            # fallback: allow module import on non-Windows
            pass

    def SvcStop(self):
        if win32serviceutil:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        # Start watching (blocking). Service control will signal stop via event.
        start_watching(config.DEFAULT_WATCH_DIR, lambda p: None)


def run_service():
    if win32serviceutil is None:
        print("pywin32 not available; cannot install/run as Windows service")
        sys.exit(1)
    win32serviceutil.HandleCommandLine(DocCompareService)


if __name__ == "__main__":
    run_service()
