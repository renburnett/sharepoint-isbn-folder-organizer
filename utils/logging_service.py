import logging
from pathlib import Path
from contextlib import contextmanager
from threading import Event, Thread
import atexit


class LogService:
    MAX_LOG_FILES = 8

    def __init__(self, log_path):
        requested_log_path = Path(log_path)

        if requested_log_path.parent.name.casefold() != "log":
            requested_log_path = (
                requested_log_path.parent / "log" / requested_log_path.name
            )

        requested_log_path.parent.mkdir(parents=True, exist_ok=True)
        self.log_path = requested_log_path

        self.logger = logging.getLogger("sharepoint_textbook_cataloging")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False

        for old_handler in list(self.logger.handlers):
            self.logger.removeHandler(old_handler)
            old_handler.close()

        handler = logging.FileHandler(self.log_path, encoding="utf-8")
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(message)s\n")
        )
        self.logger.addHandler(handler)

        self.unrenamed_folders = []
        self._summary_written = False

        previous_callback = getattr(LogService, "_active_summary_callback", None)
        if previous_callback:
            atexit.unregister(previous_callback)

        self._summary_callback = self._write_unrenamed_folders_summary
        LogService._active_summary_callback = self._summary_callback
        atexit.register(self._summary_callback)

        self.prune_old_logfiles()

    def write(self, *values, level=logging.INFO):
        message = " ".join(str(value) for value in values)
        message_lower = message.casefold()

        failure_markers = (
            "no eligible catalog result or parent folder row",
            "sharepoint could not set the folder name to isbn-13",
        )
        if any(marker in message_lower for marker in failure_markers):
            if message not in self.unrenamed_folders:
                self.unrenamed_folders.append(message)

        print(message)
        self.logger.log(level, message)

    def _write_unrenamed_folders_summary(self):
        if self._summary_written:
            return

        self._summary_written = True
        heading = "UNRENAMED FOLDERS"

        print(f"\n{heading}")
        self.logger.info(heading)

        if not self.unrenamed_folders:
            message = "None."
            print(message)
            self.logger.info(message)
        else:
            for index, details in enumerate(self.unrenamed_folders, start=1):
                message = f"{index}) {details}"
                print(message)
                self.logger.error(message)

    def exception(self, message):
        print(message)
        self.logger.exception(message)

    def prune_old_logfiles(self, current_log_path=None):
        current_log_path = Path(current_log_path or self.log_path)

        if current_log_path.parent.name.casefold() != "log":
            current_log_path = current_log_path.parent / "log" / current_log_path.name

        log_files = list(current_log_path.parent.glob("textbook_cataloging_*.log"))
        current_log_resolved = current_log_path.resolve()

        # Keep the active logfile; remove the oldest inactive ones first.
        old_logs = sorted(
            (path for path in log_files if path.resolve() != current_log_resolved),
            key=lambda path: (path.stat().st_mtime, path.name),
        )

        files_to_remove = max(0, len(log_files) - self.MAX_LOG_FILES)

        for old_log in old_logs[:files_to_remove]:
            try:
                old_log.unlink()
                self.write(f"Removed old logfile: {old_log.name}")
            except OSError as exc:
                self.write(
                    f"Could not remove old logfile {old_log}: {exc}",
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
        frames = ("⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏")
        magenta = "\033[95m"
        reset_color = "\033[0m"

        def animate():
            frame_index = 0
            while not stop_event.is_set():
                frame = frames[frame_index % len(frames)]
                print(
                    f"\r{message} {magenta}{frame}{reset_color} ",
                    end="",
                    flush=True,
                )
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
            print(f"\r{' ' * (len(message) + 3)}\r", end="", flush=True)
