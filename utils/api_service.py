from contextlib import nullcontext
import re
import time
import requests
import logging
import xml.etree.ElementTree as ET
import os
from dotenv import load_dotenv

load_dotenv()


class ApiService:
    API_TIMEOUT = (9, 13)  # (connect timeout, read timeout)

    def __init__(self, minimum_interval=1.1, logging_service=None):
        self.minimum_interval = minimum_interval
        self.logging_service = logging_service
        self.last_request_at = None

    def _spinner(self, message):
        if self.logging_service and hasattr(self.logging_service, "spinner"):
            return self.logging_service.spinner(message)
        return nullcontext()

    def wait(self):
        if self.last_request_at is not None:
            elapsed = time.monotonic() - self.last_request_at
            delay = self.minimum_interval - elapsed
            if delay > 0:
                if self.logging_service:
                    self.logging_service.write(
                        f"API rate limit: waiting {delay:.1f} seconds."
                    )
                with self._spinner("Waiting for API rate limit"):
                    time.sleep(delay)

        self.last_request_at = time.monotonic()

    def lookup_author(self, author_key):
        match = re.fullmatch(r"/?authors/(OL\d+A)", str(author_key or "").strip())
        if not match:
            if self.logging_service:
                self.logging_service.write(
                    f"Invalid Open Library author key: {author_key!r}",
                    level=logging.WARNING,
                )
            return None

        normalized_key = f"/authors/{match.group(1)}"
        url = f"https://openlibrary.org/authors/{match.group(1)}.json"
        max_attempts = 3

        for attempt in range(1, max_attempts + 1):
            try:
                self.wait()
                with self._spinner(
                    f"Looking up author in Open Library " f"({attempt}/{max_attempts})"
                ):
                    response = requests.get(url, timeout=self.API_TIMEOUT)

                if response.status_code == 404:
                    if self.logging_service:
                        self.logging_service.write(
                            f"Open Library author record not found: {normalized_key}",
                            level=logging.WARNING,
                        )
                    return None

                response.raise_for_status()
                return response.json()

            except requests.RequestException as exc:
                if attempt < max_attempts:
                    if self.logging_service:
                        self.logging_service.write(
                            f"Open Library author lookup attempt "
                            f"{attempt}/{max_attempts} failed for "
                            f"{normalized_key}: {exc}; retrying.",
                            level=logging.WARNING,
                        )
                    time.sleep(attempt)
                else:
                    if self.logging_service:
                        self.logging_service.write(
                            f"Open Library author lookup failed after "
                            f"{max_attempts} attempts for "
                            f"{normalized_key}: {exc}",
                            level=logging.WARNING,
                        )

        return None

    def _lookup_loc_catalog(self, isbn, log):
        normalized_isbn = re.sub(r"[^0-9Xx]", "", str(isbn or "")).upper()
        if not normalized_isbn:
            return None

        log(f"Trying Library of Congress for ISBN {isbn}.")

        response = None
        max_attempts = 1

        for attempt in range(1, max_attempts + 1):
            try:
                self.wait()
                with self._spinner(
                    f"Looking up ISBN in Library of Congress "
                    f"({attempt}/{max_attempts})"
                ):
                    response = requests.get(
                        "https://lx2.loc.gov/sru/lcdb",
                        params={
                            "version": "1.1",
                            "operation": "searchRetrieve",
                            "query": f"bath.isbn={normalized_isbn}",
                            "maximumRecords": 1,
                            "recordSchema": "mods",
                        },
                        timeout=self.API_TIMEOUT,
                    )

            except requests.RequestException as exc:
                log(
                    f"Library of Congress request attempt "
                    f"{attempt}/{max_attempts} failed for ISBN {isbn}: {exc}",
                    level=logging.WARNING,
                )

                if attempt == max_attempts:
                    return None

                retry_delay = 2 ** (attempt - 1)
                with self._spinner(
                    f"Retrying Library of Congress in {retry_delay} second(s)"
                ):
                    time.sleep(retry_delay)
                continue

            if response.status_code == 404:
                log(
                    f"No Library of Congress catalog record found for ISBN {isbn}.",
                    level=logging.WARNING,
                )
                return None

            retryable = response.status_code == 429 or response.status_code >= 500

            if retryable:
                if attempt == max_attempts:
                    log(
                        f"Library of Congress remained unavailable after "
                        f"{max_attempts} attempts for ISBN {isbn}. "
                        f"Last HTTP status: {response.status_code}.",
                        level=logging.WARNING,
                    )
                    return None

                retry_after = response.headers.get("Retry-After")
                try:
                    retry_delay = max(1, min(int(retry_after), 30))
                except (TypeError, ValueError):
                    retry_delay = 2 ** (attempt - 1)

                log(
                    f"Library of Congress returned HTTP {response.status_code}; "
                    f"retrying in {retry_delay} second(s).",
                    level=logging.WARNING,
                )
                with self._spinner(
                    f"Retrying Library of Congress in {retry_delay} second(s)"
                ):
                    time.sleep(retry_delay)
                continue

            try:
                response.raise_for_status()
                break
            except requests.RequestException as exc:
                log(
                    f"Library of Congress lookup failed for ISBN {isbn}: {exc}",
                    level=logging.WARNING,
                )
                return None

        if response is None:
            return None

        try:
            root = ET.fromstring(response.content)

            def local_name(element):
                return element.tag.rsplit("}", 1)[-1]

            mods_record = next(
                (element for element in root.iter() if local_name(element) == "mods"),
                None,
            )
            if mods_record is None:
                log(
                    f"No Library of Congress catalog record found for ISBN {isbn}.",
                    level=logging.WARNING,
                )
                return None

            def child_text(parent, wanted_name):
                for child in parent:
                    if local_name(child) == wanted_name:
                        value = " ".join("".join(child.itertext()).split())
                        if value:
                            return value
                return None

            title_info = next(
                (
                    element
                    for element in mods_record.iter()
                    if local_name(element) == "titleInfo"
                ),
                None,
            )
            title = child_text(title_info, "title") if title_info is not None else None
            subtitle = (
                child_text(title_info, "subTitle") if title_info is not None else None
            )

            author_names = []
            for name_element in mods_record.iter():
                if local_name(name_element) != "name":
                    continue

                roles = [
                    " ".join("".join(role.itertext()).split()).casefold()
                    for role in name_element.iter()
                    if local_name(role) == "roleTerm"
                ]
                if roles and not any(
                    role in {"author", "aut", "creator"} for role in roles
                ):
                    continue

                name_parts = [
                    " ".join("".join(part.itertext()).split())
                    for part in name_element.iter()
                    if local_name(part) == "namePart"
                ]
                name = " ".join(part for part in name_parts if part).strip()
                if name and name not in author_names:
                    author_names.append(name)

            publishers = []
            publish_dates = []
            editions = []
            responsibility = None

            for element in mods_record.iter():
                tag = local_name(element)
                value = " ".join("".join(element.itertext()).split())
                if not value:
                    continue

                if tag == "publisher" and value not in publishers:
                    publishers.append(value)
                elif tag == "dateIssued" and value not in publish_dates:
                    publish_dates.append(value)
                elif tag == "edition" and value not in editions:
                    editions.append(value)
                elif (
                    tag == "note"
                    and element.get("type", "").casefold()
                    == "statement of responsibility"
                ):
                    responsibility = value

            isbn_10 = []
            isbn_13 = []

            for element in mods_record.iter():
                if local_name(element) != "identifier":
                    continue

                identifier_type = str(element.get("type") or "").casefold()
                if "isbn" not in identifier_type:
                    continue

                value = re.sub(
                    r"[^0-9Xx]",
                    "",
                    "".join(element.itertext()),
                ).upper()

                if len(value) == 13 and value.startswith(("978", "979")):
                    if value not in isbn_13:
                        isbn_13.append(value)
                elif len(value) == 10 and value not in isbn_10:
                    isbn_10.append(value)

            result = {
                "title": title,
                "subtitle": subtitle,
                "authors": author_names,
                "by_statement": responsibility,
                "publishers": publishers,
                "publish_date": publish_dates[0] if publish_dates else None,
                "edition_name": editions[0] if editions else None,
                "isbn_10": isbn_10,
                "isbn_13": isbn_13,
            }

            if not isbn_13:
                log(
                    f"Library of Congress returned a record for ISBN {isbn}, "
                    "but it did not include an ISBN-13.",
                    level=logging.WARNING,
                )
                return None

            return result

        except (ET.ParseError, ValueError) as exc:
            log(
                f"Could not parse the Library of Congress response for "
                f"ISBN {isbn}: {exc}",
                level=logging.WARNING,
            )
            return None

    def lookup_google_books(self, isbn):
        api_key = os.getenv("GOOGLE_BOOKS_API_KEY", "").strip()
        if not api_key:
            raise ValueError("GOOGLE_BOOKS_API_KEY is not set.")

        response = requests.get(
            "https://www.googleapis.com/books/v1/volumes",
            params={
                "q": f"isbn:{isbn}",
                "maxResults": 5,
                "key": api_key,
            },
            timeout=self.API_TIMEOUT,
        )
        response.raise_for_status()
        return response.json()

    def lookup_isbn(self, isbn):
        max_attempts = 3

        def log(message, level=logging.INFO):
            if self.logging_service:
                self.logging_service.write(message, level=level)
            else:
                print(message)

        normalized_isbn = re.sub(r"[^0-9Xx]", "", str(isbn or "")).upper()
        if not normalized_isbn:
            log("Cannot look up an empty or invalid ISBN.", level=logging.WARNING)
            return None

        # Try Open Library first.
        url = f"https://openlibrary.org/isbn/{normalized_isbn}.json"

        for attempt in range(1, max_attempts + 1):
            try:
                self.wait()
                with self._spinner(
                    f"Looking up ISBN in Open Library ({attempt}/{max_attempts})"
                ):
                    response = requests.get(url, timeout=self.API_TIMEOUT)

                if response.status_code == 404:
                    log(
                        f"ISBN {isbn} not found in Open Library.",
                        level=logging.WARNING,
                    )
                    break

                response.raise_for_status()
                return response.json()

            except requests.HTTPError as exc:
                status_code = (
                    exc.response.status_code if exc.response is not None else None
                )
                retryable = status_code == 429 or (
                    status_code is not None and status_code >= 500
                )

                log(
                    f"Open Library lookup attempt {attempt}/{max_attempts} "
                    f"failed for ISBN {isbn}: {exc}",
                    level=logging.WARNING,
                )

                if not retryable or attempt == max_attempts:
                    break

            except requests.RequestException as exc:
                log(
                    f"Open Library lookup attempt {attempt}/{max_attempts} "
                    f"failed for ISBN {isbn}: {exc}",
                    level=logging.WARNING,
                )

                if attempt == max_attempts:
                    break

            retry_delay = 2 ** (attempt - 1)
            with self._spinner(f"Waiting {retry_delay} second(s) before retry"):
                time.sleep(retry_delay)

        # If Open Library has no record, try LOC.
        loc_result = self._lookup_loc_catalog(normalized_isbn, log)
        if loc_result:
            return loc_result

        # Google Books runs only when its command-line flag is enabled.
        if not self.GOOGLE_BOOKS_FALLBACK_ENABLED:
            return None

        if not os.getenv("GOOGLE_BOOKS_API_KEY", "").strip():
            log(
                "Google Books fallback is enabled, but "
                "GOOGLE_BOOKS_API_KEY is not set.",
                level=logging.WARNING,
            )
            return None

        log(f"Trying Google Books for ISBN {normalized_isbn}.")

        try:
            self.wait()
            with self._spinner("Looking up ISBN in Google Books"):
                data = self.lookup_google_books(normalized_isbn)

            for item in data.get("items") or []:
                volume_info = item.get("volumeInfo") or {}
                identifiers = volume_info.get("industryIdentifiers") or []

                result_isbns = {
                    re.sub(
                        r"[^0-9Xx]",
                        "",
                        str(identifier.get("identifier") or ""),
                    ).upper()
                    for identifier in identifiers
                    if isinstance(identifier, dict)
                }

                if normalized_isbn in result_isbns:
                    return volume_info

            log(
                f"Google Books returned no exact ISBN match for ISBN "
                f"{normalized_isbn}.",
                level=logging.WARNING,
            )

        except (requests.RequestException, ValueError) as exc:
            log(
                f"Google Books lookup failed for ISBN {normalized_isbn}: {exc}",
                level=logging.WARNING,
            )

        return None
