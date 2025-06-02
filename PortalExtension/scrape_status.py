# scrape_status.py

from threading import Lock

class ScrapeStatus:
    def __init__(self):
        self.total_rows = 0
        self.rows_processed = 0
        self.status = "idle"
        self.lock = Lock()

    def start(self):
        with self.lock:
            self.status = "scraping"
            self.total_rows = 0
            self.rows_processed = 0

    def set_total(self, total):
        with self.lock:
            self.total_rows = total

    def increment(self):
        with self.lock:
            self.rows_processed += 1

    def finish(self):
        with self.lock:
            self.status = "done"

    def get_status(self):
        with self.lock:
            return {
                "status": self.status,
                "rows_processed": self.rows_processed,
                "total_rows": self.total_rows
            }

scrape_status = ScrapeStatus()