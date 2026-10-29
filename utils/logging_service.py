import logging


class LogService:
    def __init__(self, log_path):
        self.logger = logging.getLogger("sharepoint_textbook_cataloging")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False

        if not self.logger.handlers:
            handler = logging.FileHandler(log_path, encoding="utf-8")
            handler.setFormatter(
                logging.Formatter("%(asctime)s %(levelname)s %(message)s")
            )
            self.logger.addHandler(handler)

    def write(self, *values, level=logging.INFO):
        message = " ".join(str(value) for value in values)
        print(message)
        self.logger.log(level, message)

    def exception(self, message):
        print(message)
        self.logger.exception(message)
