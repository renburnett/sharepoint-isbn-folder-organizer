import time
import requests
import logging


class ApiService:
    API_TIMEOUT = (5, 10)
    # TODO: set this to True once authentication and response schema are set up and handled respectively
    GOOGLE_BOOKS_FALLBACK_ENABLED = False

    def __init__(self, minimum_interval=1.1, log_service=None):
        self.minimum_interval = minimum_interval
        self.log_service = log_service
        self.last_request_at = None

    def wait(self):
        if self.last_request_at is not None:
            elapsed = time.monotonic() - self.last_request_at
            delay = self.minimum_interval - elapsed
            if delay > 0:
                if self.log_service:
                    self.log_service.write(
                        f"API rate limit: waiting {delay:.1f} seconds."
                    )
                time.sleep(delay)

        self.last_request_at = time.monotonic()

    def lookup_isbn(self, isbn):
        url = f"https://openlibrary.org/isbn/{isbn}.json"

        try:
            self.wait()
            response = requests.get(url, timeout=self.API_TIMEOUT)

            if response.status_code == 404:
                self.log_service.write(
                    f"ISBN {isbn} not found in Open Library.",
                    level=logging.WARNING,
                )
            elif response.status_code == 429:
                self.log_service.write(
                    "Open Library rate limited this request.",
                    level=logging.WARNING,
                )
            else:
                response.raise_for_status()
                return response.json()

        except requests.RequestException as exc:
            self.log_service.write(
                f"Open Library lookup failed for ISBN {isbn}: {exc}",
                level=logging.WARNING,
            )

        if not self.GOOGLE_BOOKS_FALLBACK_ENABLED:
            return None

        self.log_service.write(f"Trying Google Books for ISBN {isbn}.")

        response = requests.get(
            "https://www.googleapis.com/books/v1/volumes",
            params={"q": f"isbn:{isbn}", "maxResults": 1},
            timeout=self.API_TIMEOUT,
        )
        if response.status_code == 200:
            data = response.json()
            if data.get("totalItems", 0) > 0:
                return data["items"][0]["volumeInfo"]
            else:
                self.log_service.write(
                    f"ISBN {isbn} not found in Google Books.",
                    level=logging.WARNING,
                )
        else:
            self.log_service.write(
                f"Google Books lookup failed for ISBN {isbn}: "
                f"HTTP {response.status_code}",
                level=logging.WARNING,
            )

        return None
