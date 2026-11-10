import logging
from pathlib import Path
from contextlib import contextmanager
from threading import Event, Thread


class LogService:
    MAX_LOG_FILES = 6
    # TODO: maybe change this to track days instead of number of files, but for now this is simpler

    def __init__(self, log_path):
        self.logger = logging.getLogger("sharepoint_textbook_cataloging")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False

        # Ensure repeated runs in one Python process use the new logfile.
        for old_handler in list(self.logger.handlers):
            self.logger.removeHandler(old_handler)
            old_handler.close()

        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(message)s\n")
        )
        self.logger.addHandler(handler)

    def write(self, *values, level=logging.INFO):
        message = " ".join(str(value) for value in values)
        print(message)
        self.logger.log(level, message)

    def exception(self, message):
        print(message)
        self.logger.exception(message)

    def prune_old_logfiles(self, current_log_path):
        current_log_path = Path(current_log_path)
        log_files = sorted(
            current_log_path.parent.glob("textbook_cataloging_*.log"),
            key=lambda path: path.name,
        )
        current_log_resolved = current_log_path.resolve()
        old_logs = [
            path for path in log_files if path.resolve() != current_log_resolved
        ]

        for old_log in old_logs:
            if len(log_files) <= self.MAX_LOG_FILES:
                break

            try:
                old_log.unlink()
                log_files.remove(old_log)
                self.write(
                    f"Removed old logfile to keep the most recent "
                    f"{self.MAX_LOG_FILES}: {old_log.name}"
                )
            except OSError as exc:
                self.write(
                    f"Could not remove old logfile {old_log}: {exc}",
                    level=logging.WARNING,
                )

        if len(log_files) > self.MAX_LOG_FILES:
            self.write(
                f"There are {len(log_files)} matching logfiles; "
                f"the limit is {self.MAX_LOG_FILES}.",
                level=logging.WARNING,
            )

    def log_run_summary(
        self,
        successful_folder_updates=None,
        unsuccessful_folder_updates=None,
    ):
        successful = successful_folder_updates or []
        unsuccessful = unsuccessful_folder_updates or []

        self.write("")
        self.write("=" * 72)
        self.write("RUN SUMMARY")
        self.write("=" * 72)
        self.write(f"Total folder outcomes: {len(successful) + len(unsuccessful)}")
        self.write(f"Successful changes: {len(successful)}")

        if successful:
            for index, entry in enumerate(successful, start=1):
                isbn = entry.get("isbn") or "unknown"
                formats = entry.get("formats") or []

                self.write(f"{index}) ISBN {isbn}")
                self.write(f"   Title: {entry.get('title') or 'unknown'}")
                self.write(
                    f"   Folder: {entry.get('original_folder') or 'unknown'} "
                    f"-> {entry.get('final_folder') or 'unknown'}"
                )
                self.write(
                    f"   Applying detected formats for ISBN {isbn!r}: " f"{formats!r}"
                )
        else:
            self.write("No successful changes were recorded.")

        self.write(f"Unsuccessful or skipped changes: {len(unsuccessful)}")

        if unsuccessful:
            for index, entry in enumerate(unsuccessful, start=1):
                folder = entry.get("original_folder") or "unknown"
                reason = entry.get("reason") or "not recorded"
                isbn = entry.get("isbn") or "unknown"
                formats = entry.get("formats") or []

                if reason == "No eligible catalog result or parent folder row.":
                    self.write(
                        f"{index}) Folder rename skipped for {folder!r}.",
                        level=logging.WARNING,
                    )
                else:
                    self.write(
                        f"{index}) Folder update failed for {folder!r}.",
                        level=logging.ERROR,
                    )

                self.write(f"   ISBN: {isbn}")
                self.write(f"   Title: {entry.get('title') or 'unknown'}")
                self.write(
                    f"   Applying detected formats for ISBN {isbn!r}: " f"{formats!r}"
                )
                self.write(f"   Reason: {reason}")
        else:
            self.write("No unsuccessful changes were recorded.")

        self.write("=" * 72)

    @contextmanager
    def spinner(self, message="searching..."):
        stop_event = Event()
        frames = ("|", "/", "-", "\\")

        def animate():
            frame_index = 0
            while not stop_event.is_set():
                frame = frames[frame_index % len(frames)]
                print(f"\r{message} {frame}", end="", flush=True)
                frame_index += 1
                if stop_event.wait(0.12):
                    break

        spinner_thread = Thread(target=animate, daemon=True)
        spinner_thread.start()
        try:
            yield
        finally:
            stop_event.set()
            spinner_thread.join()
            print(f"\r{' ' * (len(message) + 2)}\r", end="", flush=True)
