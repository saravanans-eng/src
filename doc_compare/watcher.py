import time
import os
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler


class StableFileHandler(PatternMatchingEventHandler):
    def __init__(self, patterns, on_stable, wait_seconds=1.0):
        super().__init__(patterns=patterns, ignore_directories=True)
        self.on_stable = on_stable
        self.wait_seconds = wait_seconds

    def _handle(self, src_path):
        # wait until filesize stable
        last = -1
        stable_count = 0
        while stable_count < 2:
            try:
                size = os.path.getsize(src_path)
            except OSError:
                size = -1
            if size == last:
                stable_count += 1
            else:
                stable_count = 0
                last = size
            time.sleep(self.wait_seconds)
        self.on_stable(src_path)

    def on_created(self, event):
        self._handle(event.src_path)

    def on_modified(self, event):
        self._handle(event.src_path)


def start_watching(path: str, on_file_ready, patterns=("*_tud_ACE_For_S100_Conversion.docx",)):
    observer = Observer()
    handler = StableFileHandler(patterns=patterns, on_stable=on_file_ready)
    observer.schedule(handler, path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
