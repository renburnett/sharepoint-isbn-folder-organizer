from utils.logging_service import LogService
from utils.api_service import ApiService
from utils.doc_scan_service import DocScanService
import time
from datetime import datetime
import ctypes
import os
import re
from pathlib import Path
import logging
import pytesseract
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
    NoAlertPresentException,
    UnexpectedAlertPresentException,
    WebDriverException,
)
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from utils.constants import (
    PHONE_PATTERN,
    SHAREPOINT_URL,
    ISBN_13_PATTERN,
    ISBN_10_PATTERN,
    EXTENSION_PRIORITY,
    SCANNABLE_EXTENSIONS,
    SUPPORTED_EXTENSIONS,
    FORMAT_PRIORITY,
    FORMAT_NAMES,
)


class TextbookCatalogingService:
    MINIMUM_ROWS_TO_LOAD = 12

    def get_listing_signature(self, driver):
        elements = driver.find_elements(
            By.CSS_SELECTOR,
            '[role="row"][data-automationid^="row-"] '
            '[data-automationid="field-LinkFilename"] '
            '[data-id="heroField"]',
        )

        names = []
        for element in elements:
            name = (element.get_attribute("title") or element.text).strip()
            if name:
                names.append(name)

        return tuple(names)

    def wait_for_folder_listing(self, driver, previous_signature=None, timeout=13):
        deadline = time.monotonic() + timeout
        last_signature = None
        stable_since = None

        empty_phrases = (
            "this folder is empty",
            "no items found",
            "no files found",
            "no items to show",
        )
        busy_selector = (
            '[aria-busy="true"], [role="progressbar"], '
            '[data-automationid*="loading"]'
        )
        see_more_selector = 'div[role="gridcell"][class*="seeMoreCell"] button'

        with self.logging_service.spinner("Waiting for SharePoint folder listing"):
            while time.monotonic() < deadline:
                try:
                    signature = self.get_listing_signature(driver)
                    body_text = driver.find_element(By.TAG_NAME, "body").text.casefold()

                    busy_elements = driver.find_elements(
                        By.CSS_SELECTOR,
                        busy_selector,
                    )
                    busy = any(element.is_displayed() for element in busy_elements)

                    see_more_buttons = driver.find_elements(
                        By.CSS_SELECTOR,
                        see_more_selector,
                    )
                    see_more_visible = any(
                        button.is_displayed() for button in see_more_buttons
                    )

                    empty = any(phrase in body_text for phrase in empty_phrases)

                except StaleElementReferenceException:
                    time.sleep(0.5)
                    continue

                has_results_or_empty_message = (
                    bool(signature) or empty or see_more_visible
                )
                changed_from_parent = (
                    previous_signature is None
                    or not previous_signature
                    or signature != previous_signature
                )

                # A visible See more button means SharePoint has presented a
                # usable batch, even if its broad loading indicator remains set.
                listing_usable = not busy or see_more_visible

                if (
                    has_results_or_empty_message
                    and changed_from_parent
                    and listing_usable
                ):
                    if signature == last_signature:
                        if (
                            stable_since is not None
                            and time.monotonic() - stable_since >= 2
                        ):
                            return signature
                    else:
                        last_signature = signature
                        stable_since = time.monotonic()
                else:
                    last_signature = None
                    stable_since = None

                time.sleep(0.5)

        raise TimeoutError(
            "SharePoint did not present a stable folder listing or See more button."
        )

    def inspect_open_folder(self, driver, wait, current_folder):
        self._detected_formats = []
        self._detected_edition = None
        self._selected_isbn = None

        folders = self.get_folders_to_update(driver)
        traversed_items = []
        file_locations = []
        root_url = driver.current_url
        root_signature = self.get_listing_signature(driver)

        def edition_from_text(text):
            match = re.search(
                r"(?<![A-Za-z0-9])"
                r"(?:(?P<shorthand>\d{1,2}\s*e)|"
                r"(?P<ordinal>\d{1,2}\s*-?\s*(?:st|nd|rd|th))"
                r"[\s-]+(?:edition|ed\.?))"
                r"(?![A-Za-z0-9])",
                str(text or ""),
                flags=re.IGNORECASE,
            )
            if not match:
                return None

            edition = match.group("shorthand") or match.group("ordinal")
            return re.sub(r"[\s-]+", "", edition).casefold()

        def record_editions(items):
            if self._detected_edition:
                return

            for item in items:
                edition = edition_from_text(item.get("title"))
                if edition:
                    self._detected_edition = edition
                    self.logging_service.write(
                        f"Found edition marker in item title "
                        f"{item.get('title')!r}: {edition}"
                    )
                    return

        def formats_in(item):
            formats = set()
            title = str(item.get("title") or "")

            if item.get("is_folder"):
                formats.update(
                    self.doc_scan_service.title_tokens(item) & SUPPORTED_EXTENSIONS
                )
                return formats

            extension = os.path.splitext(title)[1].casefold().lstrip(".")
            if extension in SUPPORTED_EXTENSIONS:
                formats.add(extension)
                return formats

            format_text = str(item.get("format") or "").casefold()
            formats.update(
                token
                for token in re.findall(r"[a-z0-9]+", format_text)
                if token in SUPPORTED_EXTENSIONS
            )

            if not formats and extension:
                formats.add("other")

            return formats

        def record_formats(items):
            record_editions(items)

            formats = set(self._detected_formats)
            file_extensions = set(current_folder.get("file_extensions") or [])
            listing_extensions = set()

            for item in items:
                formats.update(formats_in(item))
                if item.get("is_folder"):
                    continue

                extension = (
                    os.path.splitext(str(item.get("title") or ""))[1]
                    .casefold()
                    .lstrip(".")
                )
                if extension:
                    file_extensions.add(extension)
                    listing_extensions.add(extension)

            formats.update(file_extensions)
            current_folder["file_extensions"] = sorted(file_extensions)
            current_folder["format"] = sorted(formats)
            self._detected_formats = sorted(formats)

            self.logging_service.write(
                "File extensions in this listing:",
                sorted(listing_extensions) if listing_extensions else "none",
            )
            self.logging_service.write(
                "All file extensions found so far:",
                current_folder["file_extensions"],
            )

        def format_for(item):
            formats = formats_in(item)
            for extension in FORMAT_PRIORITY:
                if extension in formats:
                    return extension
            return "other" if "other" in formats else None

        def listing(items):
            return [
                f"{'[folder]' if item.get('is_folder') else '[file]'} "
                f"{item.get('title') or '(untitled)'}"
                for item in items
            ]

        def isbn_from_title(item, pattern):
            title = str(item.get("title") or "")
            file_format = format_for(item)
            item["format"] = file_format

            for match in pattern.finditer(title):
                if PHONE_PATTERN.fullmatch(match.group()):
                    continue

                isbn = re.sub(r"[^0-9Xx]", "", match.group()).upper()
                return isbn, file_format

            return None, file_format

        def lookup_catalog(isbn):
            self._selected_isbn = isbn
            catalog = self.api_service.lookup_isbn(isbn)

            if not catalog:
                self.logging_service.write(
                    f"No complete catalog record was found from the enabled "
                    f"sources for ISBN {isbn}.",
                    level=logging.WARNING,
                )

            return catalog

        def is_scannable_file(item):
            if item.get("is_folder"):
                return False

            extension = (
                os.path.splitext(str(item.get("title") or ""))[1].casefold().lstrip(".")
            )
            return extension in SCANNABLE_EXTENSIONS

        def register_files(items, folder_path):
            for item in items:
                if not item.get("is_folder"):
                    file_locations.append((item, tuple(folder_path)))

        def open_and_read_folder(folder, purpose):
            try:
                self.open_folder(driver, wait, folder)
                items = self.get_folders_to_update(driver)
                traversed_items.extend(items)
                record_formats(items)
                return items
            except Exception as exc:
                self.logging_service.exception(
                    f"Could not open/read nested folder {folder.get('title')!r} "
                    f"while {purpose}; skipping this folder. "
                    f"Error: {type(exc).__name__}: {exc}"
                )
                return None

        def return_to_listing(target_url, target_signature):
            try:
                driver.back()
                wait.until(
                    lambda d: d.current_url == target_url
                    and self.get_listing_signature(d) == target_signature
                )
                self.wait_for_folder_listing(driver)
                return True
            except Exception as exc:
                self.logging_service.exception(
                    f"Could not return to the parent SharePoint listing. "
                    f"Error: {type(exc).__name__}: {exc}"
                )
                return False

        def descend_single_child_chain(first_folder, first_items):
            path = [str(first_folder.get("title") or "")]
            navigation_stack = [(root_url, root_signature)]
            items = first_items
            max_depth = 20

            for depth in range(max_depth):
                child_folders = [item for item in items if item.get("is_folder")]

                if not child_folders:
                    register_files(items, path)

                    while navigation_stack:
                        target_url, target_signature = navigation_stack.pop()
                        if not return_to_listing(target_url, target_signature):
                            return None

                    return items

                # Keep following a single-folder-only chain as before.
                if len(items) == 1 and len(child_folders) == 1:
                    child_folder = child_folders[0]
                    navigation_stack.append(
                        (driver.current_url, self.get_listing_signature(driver))
                    )
                    path.append(str(child_folder.get("title") or ""))

                    self.logging_service.write(
                        f"Opening single-child nested folder "
                        f"(level {depth + 1}): {child_folder.get('title')!r}"
                    )

                    items = open_and_read_folder(
                        child_folder,
                        f"following the single-child folder chain under "
                        f"{first_folder.get('title')!r}",
                    )
                    if items is None:
                        return None
                    continue

                # Also allow up to four shallow sibling folders at this level,
                # alongside any files in the current listing.
                if len(child_folders) > 4:
                    self.logging_service.write(
                        f"Cannot follow nested folders under "
                        f"{first_folder.get('title')!r}: found "
                        f"{len(child_folders)} child folders; at most four "
                        f"shallow child folders are allowed. Contents: "
                        f"{listing(items)}",
                        level=logging.WARNING,
                    )
                    return None

                register_files(items, path)

                for child_folder in child_folders:
                    child_title = str(child_folder.get("title") or "")
                    parent_url = driver.current_url
                    parent_signature = self.get_listing_signature(driver)

                    live_child = next(
                        (
                            candidate
                            for candidate in self.get_folders_to_update(driver)
                            if candidate.get("is_folder")
                            and str(candidate.get("title") or "") == child_title
                        ),
                        None,
                    )
                    if not live_child:
                        self.logging_service.write(
                            f"Could not find shallow child folder "
                            f"{child_title!r} under "
                            f"{first_folder.get('title')!r}; skipping this "
                            "parent folder.",
                            level=logging.WARNING,
                        )
                        return None

                    self.logging_service.write(
                        f"Opening shallow child folder: {child_title!r}"
                    )
                    child_items = open_and_read_folder(
                        live_child,
                        f"inspecting a shallow child folder under "
                        f"{first_folder.get('title')!r}",
                    )
                    if child_items is None:
                        return None

                    if any(item.get("is_folder") for item in child_items):
                        self.logging_service.write(
                            f"Cannot follow shallow child folder "
                            f"{child_title!r} under "
                            f"{first_folder.get('title')!r}: it contains "
                            f"another folder. Contents: {listing(child_items)}",
                            level=logging.WARNING,
                        )
                        return None

                    register_files(child_items, path + [child_title])
                    if not return_to_listing(parent_url, parent_signature):
                        return None

                while navigation_stack:
                    target_url, target_signature = navigation_stack.pop()
                    if not return_to_listing(target_url, target_signature):
                        return None

                return items

            self.logging_service.write(
                f"Stopped nested-folder traversal under "
                f"{first_folder.get('title')!r} after {max_depth} levels.",
                level=logging.WARNING,
            )
            return None

        record_editions([current_folder])
        record_formats(folders)
        register_files(folders, ())

        child_folders = [item for item in folders if item.get("is_folder")]

        if len(child_folders) == 1:
            child_folder = child_folders[0]
            if formats_in(child_folder):
                self.logging_service.write(
                    "Opening format folder:", child_folder["title"]
                )
            else:
                self.logging_service.write(
                    "Unexpected subfolder; inspecting its listing:",
                    child_folder.get("title"),
                    level=logging.INFO,
                )

            child_items = open_and_read_folder(
                child_folder,
                "inspecting its child folder",
            )
            if child_items is None:
                return None

            if descend_single_child_chain(child_folder, child_items) is None:
                return None

        elif child_folders:
            if len(child_folders) > 4:
                self.logging_service.write(
                    f"Skipping folder {current_folder.get('title')!r}: found "
                    f"{len(child_folders)} immediate child folders; at most four "
                    "are supported in this layout.",
                    level=logging.WARNING,
                )
                return None

            for child_folder in child_folders:
                child_title = str(child_folder.get("title") or "")
                parent_url = driver.current_url
                parent_signature = self.get_listing_signature(driver)

                live_child = next(
                    (
                        item
                        for item in self.get_folders_to_update(driver)
                        if item.get("is_folder")
                        and str(item.get("title") or "") == child_title
                    ),
                    None,
                )
                if not live_child:
                    self.logging_service.write(
                        f"Could not find child folder {child_title!r} in the "
                        "current listing; skipping this parent folder.",
                        level=logging.WARNING,
                    )
                    return None

                self.logging_service.write("Opening sibling folder:", child_title)
                child_items = open_and_read_folder(
                    live_child,
                    "inspecting a sibling folder",
                )
                if child_items is None:
                    return None

                child_signature = self.get_listing_signature(driver)
                child_url = driver.current_url
                nested_folders = [item for item in child_items if item.get("is_folder")]
                child_is_pdf = "pdf" in formats_in(child_folder)

                if not nested_folders:
                    register_files(child_items, (child_title,))
                    if not return_to_listing(parent_url, parent_signature):
                        return None
                    continue

                if child_is_pdf and len(child_items) == 1 and len(nested_folders) == 1:
                    nested_folder = nested_folders[0]
                    nested_title = str(nested_folder.get("title") or "")

                    nested_items = open_and_read_folder(
                        nested_folder,
                        f"inspecting the single nested folder inside "
                        f"PDF folder {child_title!r}",
                    )
                    if nested_items is None:
                        return None

                    nested_files = [
                        item for item in nested_items if not item.get("is_folder")
                    ]
                    nested_has_only_pdf_files = (
                        bool(nested_files)
                        and not any(item.get("is_folder") for item in nested_items)
                        and all(
                            os.path.splitext(str(item.get("title") or ""))[1].casefold()
                            == ".pdf"
                            for item in nested_files
                        )
                    )

                    if not nested_has_only_pdf_files:
                        self.logging_service.write(
                            f"Skipping folder {current_folder.get('title')!r}: "
                            f"nested folder {nested_title!r} inside PDF folder "
                            f"{child_title!r} does not contain only PDF files. "
                            f"Contents: {listing(nested_items)}",
                            level=logging.WARNING,
                        )
                        return None

                    register_files(
                        nested_items,
                        (child_title, nested_title),
                    )

                    if not return_to_listing(child_url, child_signature):
                        return None
                    if not return_to_listing(parent_url, parent_signature):
                        return None
                    continue

                self.logging_service.write(
                    f"Skipping folder {current_folder.get('title')!r}: "
                    f"child folder {child_title!r} contains an unsupported deeper "
                    f"structure. Contents: {listing(child_items)}",
                    level=logging.WARNING,
                )
                return None

        search_items = sorted(
            folders + traversed_items,
            key=lambda item: (
                0
                if format_for(item) == "pdf"
                else 1 if format_for(item) in SUPPORTED_EXTENSIONS else 2
            ),
        )

        isbn_13_candidates = []
        isbn_10_candidates = []

        for item in search_items:
            isbn, _ = isbn_from_title(item, ISBN_13_PATTERN)
            if isbn and isbn not in isbn_13_candidates:
                isbn_13_candidates.append(isbn)

            isbn, _ = isbn_from_title(item, ISBN_10_PATTERN)
            if isbn and isbn not in isbn_10_candidates:
                isbn_10_candidates.append(isbn)

        # If an ISBN-13 is already visible in an item title, look it up first.
        if isbn_13_candidates:
            isbn = isbn_13_candidates[0]
            self.logging_service.write(f"Found ISBN-13 in item titles: {isbn}")
            return lookup_catalog(isbn)

        if not child_folders:
            self.logging_service.write(
                "No supported-format folder was found; checking the current folder.",
                level=logging.WARNING,
            )

        # Try a title-found ISBN-10 immediately. If the catalog sources cannot
        # produce a complete record, scan files to continue looking for ISBN-13.
        isbn_10_lookup_attempted = False
        if isbn_10_candidates:
            isbn = isbn_10_candidates[0]
            self.logging_service.write(f"Found ISBN-10 in item titles: {isbn}")
            isbn_10_lookup_attempted = True

            catalog = lookup_catalog(isbn)
            if catalog:
                return catalog

            self.logging_service.write(
                "ISBN-10 lookup returned no complete catalog record; "
                "resuming file scanning to look for an ISBN-13.",
                level=logging.WARNING,
            )

        all_scannable_files = [
            (item, path) for item, path in file_locations if is_scannable_file(item)
        ]

        priority_files = [
            entry
            for entry in all_scannable_files
            if self.doc_scan_service.matter_priority(entry[0]) is not None
        ]

        priority_files.sort(
            key=lambda entry: (
                self.doc_scan_service.matter_priority(entry[0]),
                EXTENSION_PRIORITY.get(
                    os.path.splitext(str(entry[0].get("title") or ""))[1].casefold(),
                    99,
                ),
            )
        )

        def file_key(entry):
            item, folder_path = entry
            return (
                str(item.get("title") or "").casefold(),
                tuple(folder_path),
            )

        selected_keys = {file_key(entry) for entry in priority_files}

        # Add at most one remaining file per extension as a fallback. This lets
        # an unnamed EPUB/DOCX/TIFF get scanned without downloading every chapter.
        fallback_by_extension = {}

        for entry in all_scannable_files:
            if file_key(entry) in selected_keys:
                continue

            filename = str(entry[0].get("title") or "")
            extension = os.path.splitext(filename)[1].casefold()

            if extension not in fallback_by_extension:
                fallback_by_extension[extension] = entry

        fallback_files = sorted(
            fallback_by_extension.values(),
            key=lambda entry: (
                EXTENSION_PRIORITY.get(
                    os.path.splitext(str(entry[0].get("title") or ""))[1].casefold(),
                    99,
                ),
                str(entry[0].get("title") or "").casefold(),
            ),
        )

        scannable_files = priority_files + fallback_files

        for index, (matter_file, matter_path) in enumerate(scannable_files):
            filename = str(matter_file.get("title") or "")
            extension = os.path.splitext(filename)[1].casefold()
            matter_type = self.doc_scan_service.matter_type(matter_file)

            downloaded_path = None

            self.logging_service.write(
                f"Trying file {index + 1}/{len(scannable_files)} "
                f"({extension or 'unknown'}): {filename!r}"
            )

            try:
                # Return to this book's listing before trying another candidate.
                if index > 0:
                    driver.get(root_url)
                    self.wait_for_folder_listing(driver)

                for folder_title in matter_path:
                    live_items = self.get_folders_to_update(driver)
                    live_folder = next(
                        (
                            item
                            for item in live_items
                            if item.get("is_folder")
                            and str(item.get("title") or "") == folder_title
                        ),
                        None,
                    )
                    if not live_folder:
                        raise TimeoutError(
                            f"Could not reopen folder {folder_title!r} "
                            f"to scan {filename!r}."
                        )
                    self.open_folder(driver, wait, live_folder)

                live_items = self.get_folders_to_update(driver)
                live_file = next(
                    (
                        item
                        for item in live_items
                        if not item.get("is_folder")
                        and str(item.get("title") or "") == filename
                    ),
                    None,
                )
                if not live_file:
                    raise TimeoutError(
                        f"Could not find {filename!r} in its folder listing."
                    )

                downloaded_path = self.download_sharepoint_file(
                    driver,
                    wait,
                    live_file,
                )
                self.logging_service.write("Downloaded:", downloaded_path)

                downloaded_isbn = self.doc_scan_service.find_isbn(
                    downloaded_path,
                    matter_type,
                )

            except Exception as exc:
                self.logging_service.exception(
                    f"Failed to download or scan {filename!r} "
                    f"for folder {current_folder.get('title')!r}; "
                    f"trying the next candidate. "
                    f"Error: {type(exc).__name__}: {exc}"
                )
                continue

            finally:
                if downloaded_path is not None and downloaded_path.is_file():
                    try:
                        downloaded_path.unlink()
                        self.logging_service.write(
                            f"Deleted temporary downloaded file: {downloaded_path}"
                        )
                    except OSError as cleanup_error:
                        self.logging_service.write(
                            f"Could not delete temporary downloaded file "
                            f"{downloaded_path}: {cleanup_error}",
                            level=logging.WARNING,
                        )

            if downloaded_isbn and len(downloaded_isbn) == 13:
                self.logging_service.write(
                    f"Found ISBN-13 in downloaded file "
                    f"{filename!r}: {downloaded_isbn}"
                )
                return lookup_catalog(downloaded_isbn)

            if downloaded_isbn and len(downloaded_isbn) == 10:
                self.logging_service.write(
                    f"Found ISBN-10 in downloaded file {filename!r}: "
                    f"{downloaded_isbn}"
                )

                # Look up the first ISBN-10 immediately. If all enabled catalog
                # sources return no complete record, resume with the next file.
                if not isbn_10_lookup_attempted:
                    isbn_10_lookup_attempted = True
                    catalog = lookup_catalog(downloaded_isbn)
                    if catalog:
                        return catalog

                    self.logging_service.write(
                        "ISBN-10 lookup returned no complete catalog record; "
                        "resuming file scanning to look for an ISBN-13.",
                        level=logging.WARNING,
                    )
                else:
                    self.logging_service.write(
                        "An ISBN-10 lookup has already failed; continuing "
                        "through the remaining files to look for an ISBN-13.",
                        level=logging.WARNING,
                    )
            else:
                self.logging_service.write(
                    f"No ISBN-13 or ISBN-10 was found in {filename!r}; "
                    "trying the next candidate.",
                    level=logging.WARNING,
                )

        self.logging_service.write(
            "No ISBN-13 was found in the remaining scannable files.",
            level=logging.WARNING,
        )
        return None

    def enter_grid_view(self, driver, wait, folder_title):
        def normalized(value):
            return " ".join(str(value or "").split()).casefold()

        grid_view_deadline = time.monotonic() + 30

        def wait_for_grid_action(condition):
            remaining = grid_view_deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutException(
                    "Timed out locating the Edit in grid view command."
                )
            return WebDriverWait(driver, remaining).until(condition)

        def get_folder_row(driver):
            folder = self.get_folder_by_title(driver, folder_title)
            if not folder:
                return False

            try:
                row = folder["_elements"]["row"]
                rowheader = row.find_element(
                    By.CSS_SELECTOR,
                    '[data-automationid="field-LinkFilename"]',
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});",
                    rowheader,
                )
                ActionChains(driver).move_to_element(rowheader).pause(0.3).perform()

                return row, rowheader
            except StaleElementReferenceException:
                return False

        def get_row_more_actions(driver):
            row_data = get_folder_row(driver)
            if not row_data:
                return False

            row, _ = row_data
            try:
                buttons = row.find_elements(
                    By.CSS_SELECTOR,
                    'button[data-automationid="moreActionsHeroField"]',
                )
                return next(
                    (
                        button
                        for button in buttons
                        if button.is_displayed() and button.is_enabled()
                    ),
                    False,
                )
            except StaleElementReferenceException:
                return False

        def find_edit_grid_button(driver):
            try:
                candidates = driver.find_elements(
                    By.CSS_SELECTOR,
                    'button, [role="button"], [role="menuitem"]',
                )
                for candidate in candidates:
                    try:
                        if not candidate.is_displayed() or not candidate.is_enabled():
                            continue

                        labels = (
                            candidate.get_attribute("data-automationid"),
                            candidate.get_attribute("aria-label"),
                            candidate.get_attribute("title"),
                            candidate.get_attribute("data-command-key"),
                            candidate.text,
                        )
                        if any(
                            "edit in grid view" in normalized(label) for label in labels
                        ):
                            return candidate
                    except StaleElementReferenceException:
                        continue
            except StaleElementReferenceException:
                return False

            return False

        def grid_view_is_active(driver):
            more_actions = get_row_more_actions(driver)
            if not more_actions:
                return False

            try:
                driver.execute_script("arguments[0].click();", more_actions)

                exit_grid_locator = (
                    By.CSS_SELECTOR,
                    'button[data-automationid="exitGridCommand"]',
                )
                try:
                    WebDriverWait(driver, 5).until(
                        EC.visibility_of_element_located(exit_grid_locator)
                    )
                    return True
                except TimeoutException:
                    return False
                finally:
                    try:
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    except WebDriverException:
                        pass
            except (StaleElementReferenceException, WebDriverException):
                return False

        try:
            wait_for_grid_action(
                lambda d: self.get_folder_by_title(d, folder_title) or False
            )

            more_actions = wait_for_grid_action(get_row_more_actions)
            driver.execute_script("arguments[0].click();", more_actions)

            edit_button = wait_for_grid_action(find_edit_grid_button)
            driver.execute_script("arguments[0].click();", edit_button)

            wait_for_grid_action(grid_view_is_active)

            self.logging_service.write(
                f"Re-entered grid view for folder {folder_title!r} "
                "after saving its ISBN rename."
            )
            return True

        except (
            TimeoutException,
            StaleElementReferenceException,
            WebDriverException,
        ) as exc:
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except WebDriverException:
                pass

            self.logging_service.write(
                f"Could not re-enter grid view for folder {folder_title!r} "
                f"after saving its ISBN rename: {type(exc).__name__}: {exc}. "
                "Skipping its catalog-cell updates.",
                level=logging.ERROR,
            )
            return False

    def open_folder(self, driver, wait, folder):
        old_row = folder["_elements"]["row"]
        title_element = folder["_elements"]["title"]
        old_url = driver.current_url
        old_signature = self.get_listing_signature(driver)

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
            title_element,
        )
        wait.until(EC.element_to_be_clickable(title_element))

        try:
            title_element.click()
        except ElementClickInterceptedException:
            self.logging_service.write(
                f"Folder title click was intercepted for {folder.get('title')!r}; "
                "retrying with a DOM click because a SharePoint header or "
                "overlay covered the title.",
                level=logging.WARNING,
            )
            driver.execute_script("arguments[0].click();", title_element)

        wait.until(EC.any_of(EC.staleness_of(old_row), EC.url_changes(old_url)))
        self.wait_for_folder_listing(driver, previous_signature=old_signature)

    def open_first_eligible_folder(
        self, driver, wait, open_subfolders=True, excluded_titles=None
    ):
        self.wait_for_folder_listing(driver)

        wait.until(
            lambda d: len(
                d.find_elements(
                    By.CSS_SELECTOR,
                    '[role="row"][data-automationid^="row-"]',
                )
            )
            > self.MINIMUM_ROWS_TO_LOAD
        )

        items = self.get_folders_to_update(driver)
        excluded_names = ("abueg", "izaiah", "burnett", "-sa")

        passed_excluded_titles = excluded_titles
        excluded_titles = {str(title).casefold() for title in (excluded_titles or ())}
        parent_listing_url = driver.current_url

        while True:
            eligible_folder = None

            for item in items:
                modified_by = (item.get("modified_by") or "").casefold()
                formats = (item.get("format") or "").strip()
                book_title = (item.get("book_title") or "").strip()
                title = str(item.get("title") or "")
                title_key = title.casefold()
                excluded = any(name in modified_by for name in excluded_names)

                if (
                    item.get("is_folder")
                    and title_key not in excluded_titles
                    and not excluded
                    and not formats
                    and not book_title
                ):
                    eligible_folder = item
                    break

            if eligible_folder is None:
                self.logging_service.write(
                    f"No folder matches the conditions for opening. Folders: {items}",
                    level=logging.WARNING,
                )
                return None

            if not open_subfolders:
                return eligible_folder

            title = str(eligible_folder.get("title") or "")
            title_key = title.casefold()

            try:
                self.open_folder(driver, wait, eligible_folder)
                return eligible_folder

            except (
                TimeoutException,
                TimeoutError,
                StaleElementReferenceException,
                WebDriverException,
            ) as exc:
                excluded_titles.add(title_key)
                if isinstance(passed_excluded_titles, set):
                    passed_excluded_titles.add(title_key)

                self.logging_service.write(
                    f"Could not open folder {title!r}: "
                    f"{type(exc).__name__}: {exc}. "
                    "Skipping it and trying the next eligible folder.",
                    level=logging.WARNING,
                )

                recovered = False
                for recovery_attempt in range(1, 3):
                    try:
                        driver.get(parent_listing_url)
                        self.wait_for_folder_listing(driver, timeout=20)

                        see_more_locator = (
                            By.CSS_SELECTOR,
                            'div[role="gridcell"][class*="seeMoreCell"] button',
                        )
                        see_more_buttons = driver.find_elements(*see_more_locator)
                        if any(button.is_displayed() for button in see_more_buttons):
                            self.click_safely(
                                driver,
                                WebDriverWait(driver, 10),
                                see_more_locator,
                            )
                            self.wait_for_folder_listing(driver, timeout=20)

                        WebDriverWait(driver, 15).until(
                            lambda d: len(
                                d.find_elements(
                                    By.CSS_SELECTOR,
                                    '[role="row"][data-automationid^="row-"]',
                                )
                            )
                            > self.MINIMUM_ROWS_TO_LOAD
                        )

                        items = self.get_folders_to_update(driver)
                        recovered = True
                        break

                    except (
                        TimeoutException,
                        TimeoutError,
                        StaleElementReferenceException,
                        WebDriverException,
                    ) as recovery_exc:
                        self.logging_service.write(
                            f"Could not reload the parent folder listing "
                            f"(attempt {recovery_attempt}/2): "
                            f"{type(recovery_exc).__name__}: {recovery_exc}.",
                            level=logging.WARNING,
                        )

                if not recovered:
                    self.logging_service.write(
                        "Could not restore the parent folder listing after the "
                        "open-folder failure. Stopping folder selection.",
                        level=logging.ERROR,
                    )
                    return None

    def click_safely(self, driver, wait, locator):
        try:
            wait.until(EC.element_to_be_clickable(locator)).click()
        except ElementClickInterceptedException:
            # Clear the profile card that can appear on hover.
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            wait.until(
                EC.invisibility_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        'lpc-card[role="dialog"][aria-label="Profile Card"]',
                    )
                )
            )
            wait.until(EC.element_to_be_clickable(locator)).click()

    def read_cell(self, driver, row, selector, use_title_attribute=False):
        matches = row.find_elements(By.CSS_SELECTOR, selector)
        if not matches:
            return None, None
        element = matches[0]
        text = element.get_attribute("title") if use_title_attribute else None
        if not text:
            text = driver.execute_script(
                "return (arguments[0].innerText || "
                "arguments[0].textContent || '').trim();",
                element,
            )
        return text or None, element

    def get_folders_to_update(self, driver):
        for attempt in range(3):
            try:
                return self._read_folders_to_update_once(driver)
            except StaleElementReferenceException:
                if attempt == 2:
                    raise
                time.sleep(0.5)

    def _read_folders_to_update_once(self, driver):
        rows = driver.find_elements(
            By.CSS_SELECTOR, '[role="row"][data-automationid^="row-"]'
        )

        items = []
        for row in rows:
            title, title_element = self.read_cell(
                driver,
                row,
                '[data-automationid="field-LinkFilename"] ' '[data-id="heroField"]',
                use_title_attribute=True,
            )
            if not title_element:
                continue
            modified_by, modified_by_cell = self.read_cell(
                driver, row, '[data-automationid="field-Editor"]'
            )
            formats, formats_cell = self.read_cell(
                driver, row, '[data-automationid="field-Formats"]'
            )
            book_title, book_title_cell = self.read_cell(
                driver, row, '[data-automationid="field-BookTitle"]'
            )
            author, author_cell = self.read_cell(
                driver, row, '[data-automationid="field-Author0"]'
            )
            publisher, publisher_cell = self.read_cell(
                driver, row, '[data-automationid="field-Publisher"]'
            )
            edition, edition_cell = self.read_cell(
                driver, row, '[data-automationid="field-Edition"]'
            )
            is_folder = bool(
                row.find_elements(
                    By.CSS_SELECTOR,
                    '[data-automationid="field-DocIcon"] '
                    '[aria-label="Yellow folder"]',
                )
            )
            items.append(
                {
                    "title": title,
                    "modified_by": modified_by,
                    "format": formats,
                    "book_title": book_title,
                    "author": author,
                    "publisher": publisher,
                    "edition": edition,
                    "is_folder": is_folder,
                    "_elements": {
                        "row": row,
                        "title": title_element,
                        "modified_by": modified_by_cell,
                        "format": formats_cell,
                        "book_title": book_title_cell,
                        "author": author_cell,
                        "publisher": publisher_cell,
                        "edition": edition_cell,
                    },
                }
            )
        return items

    def get_folder_by_title(self, driver, title):
        return next(
            (
                item
                for item in self.get_folders_to_update(driver)
                if item.get("is_folder") and item.get("title") == title
            ),
            None,
        )

    def rename_sharepoint_folder(
        self,
        driver,
        wait,
        folder,
        new_title,
        allow_duplicate=True,
        is_rollback=False,
        original_title=None,
        ignored_messages=(),
    ):
        old_title = folder["title"]
        original_title = original_title or old_title
        ignored_messages = {str(message).casefold() for message in ignored_messages}

        original_url = driver.current_url
        original_signature = self.get_listing_signature(driver)

        refreshed_folder = self.get_folder_by_title(driver, old_title)
        if refreshed_folder:
            folder.update(refreshed_folder)
        row_ref = [folder["_elements"]["row"]]

        editor_selectors = (
            'input:not([type="checkbox"])',
            "textarea",
            '[contenteditable="true"]',
            '[role="textbox"]',
        )

        def find_visible_editor(driver):
            containers = [row_ref[0]]

            for locator in (
                (By.ID, "CommonEditorCalloutId"),
                (By.CSS_SELECTOR, '[role="dialog"]'),
            ):
                containers.extend(
                    element
                    for element in driver.find_elements(*locator)
                    if element.is_displayed()
                )

            for container in containers:
                try:
                    for selector in editor_selectors:
                        for element in container.find_elements(
                            By.CSS_SELECTOR, selector
                        ):
                            if element.is_displayed():
                                return element
                except StaleElementReferenceException:
                    continue

            return False

        title_cell = row_ref[0].find_element(
            By.CSS_SELECTOR,
            '[data-automationid="field-LinkFilename"]',
        )
        title_cell.click()

        try:
            editor = WebDriverWait(driver, 2).until(find_visible_editor)
        except TimeoutException:
            if (
                driver.current_url != original_url
                or self.get_listing_signature(driver) != original_signature
            ):
                raise TimeoutException(
                    "Clicking the title changed the listing instead of opening its editor."
                )

            current_folder = self.get_folder_by_title(driver, old_title)
            if not current_folder:
                raise TimeoutException(
                    f"Could not find the folder {old_title!r} after selecting it."
                )

            folder.update(current_folder)
            row_ref[0] = folder["_elements"]["row"]

            title_cell = row_ref[0].find_element(
                By.CSS_SELECTOR,
                '[data-automationid="field-LinkFilename"]',
            )
            title_cell.click()
            editor = wait.until(find_visible_editor)

        editor.click()
        editor.send_keys(Keys.CONTROL, "a")
        editor.send_keys(new_title)
        editor.send_keys(Keys.ENTER)

        # Limit the grid exit and rename confirmation to 7 seconds total.
        rename_deadline = time.monotonic() + 7

        def wait_for_rename(condition):
            remaining = rename_deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutException(
                    "The folder rename did not finish within 7 seconds."
                )

            try:
                return WebDriverWait(
                    driver,
                    remaining,
                    poll_frequency=0.2,
                ).until(condition)
            except TimeoutException as exc:
                raise TimeoutException(
                    "The folder rename did not finish within 7 seconds."
                ) from exc

        # Exit grid view to submit the ISBN rename before editing other cells.
        with self.logging_service.spinner("Submitting folder rename"):
            try:
                # Prefer the source row because the destination folder may already exist.
                current_folder = self.get_folder_by_title(
                    driver, old_title
                ) or self.get_folder_by_title(driver, new_title)
                if current_folder:
                    folder.update(current_folder)
                    row_ref[0] = folder["_elements"]["row"]

                title_element = folder["_elements"]["title"]
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});",
                    title_element,
                )
                ActionChains(driver).move_to_element(title_element).pause(0.3).perform()

                folder_row = title_element.find_element(
                    By.XPATH,
                    "./ancestor::*[@role='row'][1]",
                )
                more_actions = wait_for_rename(
                    lambda d: next(
                        (
                            button
                            for button in folder_row.find_elements(
                                By.CSS_SELECTOR,
                                'button[data-automationid="moreActionsHeroField"]',
                            )
                            if button.is_displayed() and button.is_enabled()
                        ),
                        False,
                    )
                )
                driver.execute_script("arguments[0].click();", more_actions)

                exit_grid_locator = (
                    By.CSS_SELECTOR,
                    'button[data-automationid="exitGridCommand"]',
                )
                exit_grid_button = wait_for_rename(
                    EC.visibility_of_element_located(exit_grid_locator)
                )
                driver.execute_script("arguments[0].click();", exit_grid_button)
                wait_for_rename(EC.invisibility_of_element_located(exit_grid_locator))

                self.logging_service.write(
                    f"Exited grid view to save ISBN folder rename "
                    f"from {old_title!r} to {new_title!r}."
                )
            except (TimeoutException, StaleElementReferenceException) as exc:
                self.logging_service.write(
                    f"Could not submit the ISBN folder rename for {old_title!r} "
                    f"to {new_title!r} by exiting grid view: "
                    f"{type(exc).__name__}: {exc}. Skipping this folder.",
                    level=logging.ERROR,
                )
                try:
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                except WebDriverException:
                    pass
                return False

        error_message_locator = (
            By.CSS_SELECTOR,
            '[data-automationid="error-icon-button"][data-error-message], '
            '[data-automationid="list-info-callout-container"] '
            '[data-automationid="AsReactChildren"], '
            '[class*="errorCallout"] [data-automationid="AsReactChildren"], '
            '[role="alert"]',
        )

        def rename_result(driver):
            try:
                for message_element in driver.find_elements(*error_message_locator):
                    if not message_element.is_displayed():
                        continue

                    message = (
                        message_element.get_attribute("data-error-message")
                        or message_element.text
                        or ""
                    ).strip()
                    if not message:
                        continue

                    message_lower = message.casefold()
                    if message_lower in ignored_messages:
                        continue

                    if (
                        "already exists" in message_lower
                        or "duplicate" in message_lower
                    ):
                        return "collision", message

                    return "error", message

                current_titles = self.get_listing_signature(driver)
                if old_title == new_title and new_title in current_titles:
                    return "renamed", None
                if old_title not in current_titles and new_title in current_titles:
                    return "renamed", None

                return False

            except StaleElementReferenceException:
                return False

        try:
            with self.logging_service.spinner(
                "Waiting for SharePoint to confirm folder rename"
            ):
                result, message = wait_for_rename(rename_result)
        except TimeoutException as exc:
            self.logging_service.write(
                f"SharePoint did not confirm the folder rename from "
                f"{old_title!r} to {new_title!r} within 7 seconds. "
                f"Error: {exc}",
                level=logging.ERROR,
            )
            return False

        if result in ("collision", "error"):
            context = getattr(self, "_catalog_log_context", None)
            description = (
                "Name collision"
                if result == "collision"
                else "SharePoint rejected the folder rename"
            )
            self.logging_service.write(
                f"{description}. Current folder: {old_title!r}; "
                f"requested title: {new_title!r}; "
                f"SharePoint message: {message!r}; "
                f"catalog context: {context!r}. "
                "Skipping this folder and continuing.",
                level=logging.ERROR,
            )
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except WebDriverException:
                pass

            ignored_messages.add(str(message or "").casefold())

            if (
                result == "collision"
                and allow_duplicate
                and not new_title.casefold().endswith("-duplicate")
            ):
                duplicate_title = f"{new_title}-DUPLICATE"
                self.logging_service.write(
                    f"Retrying the folder rename with duplicate title "
                    f"{duplicate_title!r}.",
                    level=logging.WARNING,
                )
                retry_folder = self.get_folder_by_title(driver, old_title) or folder
                return self.rename_sharepoint_folder(
                    driver,
                    wait,
                    retry_folder,
                    duplicate_title,
                    allow_duplicate=False,
                    is_rollback=False,
                    original_title=original_title,
                    ignored_messages=ignored_messages,
                )

            if not is_rollback:
                self.logging_service.write(
                    f"Restoring folder title to its original value "
                    f"{original_title!r} after the rename error.",
                    level=logging.WARNING,
                )
                rollback_folder = self.get_folder_by_title(driver, old_title) or folder
                restored = self.rename_sharepoint_folder(
                    driver,
                    wait,
                    rollback_folder,
                    original_title,
                    allow_duplicate=False,
                    is_rollback=True,
                    original_title=original_title,
                    ignored_messages=ignored_messages,
                )
                if restored:
                    self.logging_service.write(
                        f"Restored original folder title {original_title!r}; "
                        "skipping this folder's catalog updates."
                    )
                else:
                    self.logging_service.write(
                        f"Could not confirm restoration of original folder title "
                        f"{original_title!r}.",
                        level=logging.ERROR,
                    )
                return False

            return False

        folder["title"] = new_title
        updated_folder = self.get_folder_by_title(driver, new_title)
        if updated_folder:
            folder.update(updated_folder)

        return True

    def update_folder_formats(self, driver, wait, folder, formats):
        format_aliases = {
            "pdf": ("PDF",),
            "epub": ("Epub",),
            "docx": ("Docx",),
            "doc": ("Docx",),
            "rtf": ("Other",),
            "tif": ("Tif",),
            "other": ("Other",),
        }
        canonical_formats = {"tiff": "tif"}

        extensions = list(
            dict.fromkeys(
                canonical_formats.get(str(ext).casefold(), str(ext).casefold())
                for ext in formats
            )
        )
        extensions = [
            extension for extension in extensions if extension in format_aliases
        ]
        if not extensions:
            return

        folder = self.get_folder_by_title(driver, folder["title"]) or folder
        folder_title = folder["title"]
        original_url = driver.current_url
        original_signature = self.get_listing_signature(driver)
        format_cell = folder["_elements"]["format"]

        try:
            wait.until(EC.element_to_be_clickable(format_cell)).click()
            WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.ID, "CommonEditorCalloutId"))
            )
        except TimeoutException:
            if (
                driver.current_url != original_url
                or self.get_listing_signature(driver) != original_signature
            ):
                raise TimeoutException(
                    "Clicking the Formats cell changed the listing instead "
                    "of opening its editor."
                )

            refreshed_folder = self.get_folder_by_title(driver, folder_title)
            if not refreshed_folder:
                raise TimeoutException(
                    f"Could not find folder {folder_title!r} after selecting "
                    "the Formats cell."
                )

            folder.update(refreshed_folder)
            format_cell = folder["_elements"]["format"]
            wait.until(EC.element_to_be_clickable(format_cell)).click()
            wait.until(
                EC.visibility_of_element_located((By.ID, "CommonEditorCalloutId"))
            )

        current_tokens = set(
            re.findall(r"[a-z0-9]+", str(folder.get("format") or "").casefold())
        )
        already_selected = {
            extension
            for extension, aliases in format_aliases.items()
            if any(alias.casefold() in current_tokens for alias in aliases)
        }
        selected_titles = set()

        for extension in extensions:
            if extension in already_selected:
                continue

            callout = wait.until(
                EC.visibility_of_element_located((By.ID, "CommonEditorCalloutId"))
            )
            options = callout.find_elements(By.CSS_SELECTOR, '[role="option"]')
            aliases = {alias.casefold() for alias in format_aliases[extension]}

            def option_matches(candidate):
                labels = (
                    candidate.get_attribute("title"),
                    candidate.get_attribute("aria-label"),
                    candidate.text,
                )
                option_tokens = {
                    token
                    for label in labels
                    for token in re.findall(r"[a-z0-9]+", str(label or "").casefold())
                }
                return bool(aliases & option_tokens)

            option = next(
                (candidate for candidate in options if option_matches(candidate)),
                None,
            )

            if option is None:
                self.logging_service.write(
                    f"No SharePoint Formats choice found for {extension!r}; "
                    f"available choices: "
                    f"{[option.text for option in options]!r}; skipping it.",
                    level=logging.WARNING,
                )
                continue

            option_title = (
                (option.get_attribute("title") or option.text).strip().casefold()
            )

            if option_title not in selected_titles:
                driver.execute_script("arguments[0].click();", option)
                selected_titles.add(option_title)

        callout = wait.until(
            EC.visibility_of_element_located((By.ID, "CommonEditorCalloutId"))
        )
        save_button = callout.find_element(
            By.ID,
            "acceptButtonInCommonEditorCallout",
        )
        driver.execute_script("arguments[0].click();", save_button)
        wait.until(EC.invisibility_of_element_located((By.ID, "CommonEditorCalloutId")))

    def update_folder_text_field(self, driver, wait, folder, field_name, value):
        if not value:
            return

        folder = self.get_folder_by_title(driver, folder["title"]) or folder
        folder_title = folder["title"]
        original_url = driver.current_url
        original_signature = self.get_listing_signature(driver)
        row_ref = [folder["_elements"]["row"]]

        editor_selectors = (
            'input:not([type="checkbox"])',
            "textarea",
            '[contenteditable="true"]',
            '[role="textbox"]',
        )

        def find_visible_editor(driver):
            containers = []

            for locator in (
                (By.ID, "CommonEditorCalloutId"),
                (By.CSS_SELECTOR, '[role="dialog"]'),
            ):
                containers.extend(
                    element
                    for element in driver.find_elements(*locator)
                    if element.is_displayed()
                )

            try:
                if row_ref[0].is_displayed():
                    containers.append(row_ref[0])
            except StaleElementReferenceException:
                refreshed_folder = self.get_folder_by_title(driver, folder_title)
                if refreshed_folder:
                    folder.update(refreshed_folder)
                    row_ref[0] = folder["_elements"]["row"]
                    containers.append(row_ref[0])

            for container in containers:
                try:
                    for selector in editor_selectors:
                        for candidate in container.find_elements(
                            By.CSS_SELECTOR, selector
                        ):
                            if candidate.is_displayed():
                                return candidate
                except StaleElementReferenceException:
                    continue

            return False

        cell = folder["_elements"][field_name]
        wait.until(EC.element_to_be_clickable(cell)).click()

        try:
            editor = WebDriverWait(driver, 2).until(find_visible_editor)
        except TimeoutException:
            if (
                driver.current_url != original_url
                or self.get_listing_signature(driver) != original_signature
            ):
                raise TimeoutException(
                    f"Clicking the {field_name} cell changed the listing "
                    "instead of opening its editor."
                )

            refreshed_folder = self.get_folder_by_title(driver, folder_title)
            if not refreshed_folder:
                raise TimeoutException(
                    f"Could not find folder {folder_title!r} after selecting "
                    f"the {field_name} cell."
                )

            folder.update(refreshed_folder)
            row_ref[0] = folder["_elements"]["row"]
            cell = folder["_elements"][field_name]

            wait.until(EC.element_to_be_clickable(cell)).click()
            editor = wait.until(find_visible_editor)

        editor.click()
        editor.send_keys(Keys.CONTROL, "a")
        editor.send_keys(str(value))
        editor.send_keys(Keys.ENTER)

        # Tab leaves the inline editor; persistence is checked after reload.
        ActionChains(driver).send_keys(Keys.TAB).perform()

        wait.until(
            EC.any_of(
                EC.staleness_of(editor),
                EC.invisibility_of_element(editor),
            )
        )

        expected_value = " ".join(str(value).split()).casefold()

        def field_value_visible(driver):
            updated_folder = self.get_folder_by_title(driver, folder_title)
            if not updated_folder:
                return False

            actual_value = " ".join(
                str(updated_folder.get(field_name) or "").split()
            ).casefold()

            if actual_value != expected_value:
                return False

            folder.update(updated_folder)
            return True

        try:
            wait.until(field_value_visible)
        except TimeoutException as exc:
            updated_folder = self.get_folder_by_title(driver, folder_title)
            actual_value = updated_folder.get(field_name) if updated_folder else None
            raise TimeoutException(
                f"The current SharePoint view did not reflect {field_name!r} "
                f"for folder {folder_title!r}. Expected {value!r}; "
                f"found {actual_value!r}."
            ) from exc

        self.logging_service.write(
            f"SharePoint's current view shows {field_name!r} for "
            f"{folder_title!r}: ",
            value,
        )

    def verify_folder_catalog_update(
        self,
        driver,
        wait,
        original_title,
        expected_title,
        expected_formats,
        expected_fields,
        max_retries=2,
    ):

        expected_format_set = {
            FORMAT_NAMES[str(extension).casefold()]
            for extension in expected_formats
            if str(extension).casefold() in FORMAT_NAMES
        }

        def normalized(value):
            return " ".join(str(value or "").split()).casefold()

        def normalized_formats(value):
            tokens = re.findall(r"[a-z0-9]+", str(value or "").casefold())
            return {FORMAT_NAMES[token] for token in tokens if token in FORMAT_NAMES}

        def handle_reload_warning():
            try:
                alert = driver.switch_to.alert
            except NoAlertPresentException:
                return False

            alert_text = alert.text.casefold()
            if not any(
                phrase in alert_text
                for phrase in ("may not be saved", "changes you made", "unsaved")
            ):
                alert.dismiss()
                raise TimeoutException(
                    f"Unexpected browser confirmation during save verification: "
                    f"{alert_text!r}"
                )

            self.logging_service.write(
                "SharePoint warned about pending changes. Reloading to check "
                "what persisted; missing values will be reapplied.",
                level=logging.WARNING,
            )
            alert.accept()
            return True

        def find_target_folder(driver):
            try:
                return self.get_folder_by_title(
                    driver, expected_title
                ) or self.get_folder_by_title(driver, original_title)
            except StaleElementReferenceException:
                return False

        def has_visible_see_more(driver):
            try:
                return any(
                    button.is_displayed()
                    for button in driver.find_elements(
                        By.CSS_SELECTOR,
                        'div[role="gridcell"][class*="seeMoreCell"] button',
                    )
                )
            except StaleElementReferenceException:
                return False

        def exit_grid_view(current_folder):
            current_folder = (
                self.get_folder_by_title(driver, expected_title) or current_folder
            )
            title_element = current_folder["_elements"]["title"]
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                title_element,
            )
            ActionChains(driver).move_to_element(title_element).pause(0.3).perform()

            folder_row = title_element.find_element(
                By.XPATH,
                "./ancestor::*[@role='row'][1]",
            )
            more_actions = wait.until(
                lambda d: next(
                    (
                        button
                        for button in folder_row.find_elements(
                            By.CSS_SELECTOR,
                            'button[data-automationid="moreActionsHeroField"]',
                        )
                        if button.is_displayed() and button.is_enabled()
                    ),
                    False,
                )
            )
            driver.execute_script("arguments[0].click();", more_actions)

            exit_grid_locator = (
                By.CSS_SELECTOR,
                'button[data-automationid="exitGridCommand"]',
            )
            exit_grid_button = wait.until(
                EC.visibility_of_element_located(exit_grid_locator)
            )
            driver.execute_script("arguments[0].click();", exit_grid_button)
            wait.until(EC.invisibility_of_element_located(exit_grid_locator))
            self.logging_service.write(
                f"Exited grid view after retrying catalog values for "
                f"{expected_title!r}."
            )
            return self.get_folder_by_title(driver, expected_title) or current_folder

        def reload_and_read_folder():
            try:
                driver.refresh()
            except UnexpectedAlertPresentException:
                if not handle_reload_warning():
                    raise TimeoutException(
                        "SharePoint blocked the reload, so saved values "
                        "could not be verified."
                    )

            try:
                handle_reload_warning()
                wait.until(
                    lambda d: d.execute_script("return document.readyState")
                    == "complete"
                )
            except UnexpectedAlertPresentException:
                if not handle_reload_warning():
                    raise TimeoutException(
                        "SharePoint blocked the reload, so saved values "
                        "could not be verified."
                    )
                wait.until(
                    lambda d: d.execute_script("return document.readyState")
                    == "complete"
                )

            navigation_type = driver.execute_script(
                "const entry = performance.getEntriesByType('navigation')[0]; "
                "return entry ? entry.type : null;"
            )
            if navigation_type != "reload":
                raise TimeoutException(
                    "The browser did not complete a fresh reload; "
                    "SharePoint values cannot be confirmed."
                )

            self.wait_for_folder_listing(driver, timeout=30)

            see_more_locator = (
                By.CSS_SELECTOR,
                'div[role="gridcell"][class*="seeMoreCell"] button',
            )
            pagination_wait = WebDriverWait(driver, 15)

            # Load successive batches because the ISBN rename may move the row
            # farther down the sorted SharePoint listing.
            for _ in range(100):
                current_folder = find_target_folder(driver)
                if current_folder:
                    return current_folder

                if not has_visible_see_more(driver):
                    break

                previous_signature = self.get_listing_signature(driver)
                self.click_safely(driver, pagination_wait, see_more_locator)

                try:
                    self.wait_for_folder_listing(
                        driver,
                        previous_signature=previous_signature,
                        timeout=20,
                    )
                except TimeoutError:
                    current_folder = find_target_folder(driver)
                    if current_folder:
                        return current_folder
                    raise TimeoutException(
                        "SharePoint did not load more rows while verifying "
                        f"folder {expected_title!r}."
                    )

            current_folder = find_target_folder(driver)
            if current_folder:
                return current_folder

            raise TimeoutException(
                f"Could not find folder {expected_title!r} or "
                f"{original_title!r} in the refreshed SharePoint listing."
            )

        for verification_round in range(max_retries + 1):
            current_folder = reload_and_read_folder()

            mismatched_fields = {
                field_name: expected_value
                for field_name, expected_value in expected_fields.items()
                if normalized(current_folder.get(field_name))
                != normalized(expected_value)
            }
            title_mismatch = normalized(current_folder.get("title")) != normalized(
                expected_title
            )
            actual_formats = normalized_formats(current_folder.get("format"))
            missing_formats = expected_format_set - actual_formats
            grid_view_open = False

            if not title_mismatch and not mismatched_fields and not missing_formats:
                self.logging_service.write(
                    f"VERIFIED after reload: folder {expected_title!r}, "
                    f"formats {sorted(actual_formats)}, and fields "
                    f"{expected_fields!r} persisted in SharePoint."
                )
                return current_folder

            details = {
                "folder_title": current_folder.get("title"),
                "formats": current_folder.get("format"),
                **{
                    field_name: current_folder.get(field_name)
                    for field_name in expected_fields
                },
            }

            if verification_round == max_retries:
                message = (
                    f"SharePoint values did not persist after {max_retries} "
                    f"retries. Expected folder {expected_title!r}, formats "
                    f"{sorted(expected_format_set)}, and fields "
                    f"{expected_fields!r}; reloaded values: {details!r}."
                )
                self.logging_service.write(message, level=logging.ERROR)
                raise TimeoutException(message)

            self.logging_service.write(
                f"Reload verification found unsaved values; retry "
                f"{verification_round + 1} of {max_retries}. "
                f"Current values: {details!r}",
                level=logging.WARNING,
            )

            if title_mismatch:
                if not self.rename_sharepoint_folder(
                    driver, wait, current_folder, expected_title
                ):
                    self.logging_service.write(
                        f"Stopping verification for "
                        f"{current_folder.get('title')!r} because SharePoint "
                        "rejected the rename. This folder will be skipped.",
                        level=logging.WARNING,
                    )
                    return current_folder

                current_folder = (
                    self.get_folder_by_title(driver, expected_title) or current_folder
                )

            if missing_formats:
                self.update_folder_formats(
                    driver,
                    wait,
                    current_folder,
                    sorted(missing_formats),
                )
                grid_view_open = True

            for field_name, expected_value in mismatched_fields.items():
                self.update_folder_text_field(
                    driver,
                    wait,
                    current_folder,
                    field_name,
                    expected_value,
                )
                grid_view_open = True

            if grid_view_open:
                current_folder = exit_grid_view(current_folder)

    def update_folder_from_catalog(self, driver, wait, folder, catalog):
        isbn_13_value = catalog.get("isbn_13")

        if not isbn_13_value:
            isbn_13_value = [
                identifier.get("identifier")
                for identifier in catalog.get("industryIdentifiers", [])
                if isinstance(identifier, dict) and identifier.get("type") == "ISBN_13"
            ]

        if not isinstance(isbn_13_value, (list, tuple)):
            isbn_13_value = [isbn_13_value]

        selected_isbn = str(getattr(self, "_selected_isbn", "") or "")
        selected_isbn = re.sub(r"[^0-9Xx]", "", selected_isbn).upper()

        catalog_isbn_13_values = []
        for value in isbn_13_value:
            candidate = re.sub(r"[^0-9]", "", str(value or ""))
            if (
                len(candidate) == 13
                and candidate.startswith(("978", "979"))
                and candidate not in catalog_isbn_13_values
            ):
                catalog_isbn_13_values.append(candidate)

        isbn = None
        if (
            len(selected_isbn) == 13
            and selected_isbn.startswith(("978", "979"))
            and selected_isbn in catalog_isbn_13_values
        ):
            isbn = selected_isbn
        elif catalog_isbn_13_values:
            # Keep catalog order when the requested ISBN-13 is not present.
            isbn = catalog_isbn_13_values[0]
        elif len(selected_isbn) == 13 and selected_isbn.startswith(("978", "979")):
            # Retain the existing fallback for records without a catalog ISBN-13.
            isbn = selected_isbn

        if not isbn:
            self.logging_service.write(
                f"Catalog lookup ISBN {selected_isbn!r} did not produce a valid "
                f"ISBN-13 for folder {folder.get('title')!r}. "
                f"Catalog title: "
                f"{catalog.get('title') or catalog.get('full_title')!r}; "
                f"catalog ISBN-13 values: {isbn_13_value!r}. "
                "The folder will not be renamed or updated.",
                level=logging.ERROR,
            )
            self._catalog_log_context = None
            return False

        title = (
            catalog.get("title")
            or catalog.get("full_title")
            or catalog.get("short_title")
        )
        subtitle = str(catalog.get("subtitle") or "").strip()

        if title:
            title = str(title).strip()
            if subtitle and subtitle.casefold() not in title.casefold():
                title = f"{title}: {subtitle}"

        print("isbn", isbn)
        print("title", title)

        # Preserve detected formats before the rename refreshes the SharePoint row.
        formats_to_apply = []
        for format_values in (
            folder.get("format"),
            folder.get("file_extensions"),
        ):
            if isinstance(format_values, str):
                format_values = re.findall(
                    r"[a-z0-9]+",
                    format_values.casefold(),
                )
            elif not isinstance(format_values, (list, tuple, set)):
                format_values = [format_values] if format_values else []

            for format_value in format_values:
                extension = str(format_value or "").strip().casefold()
                if extension and extension not in formats_to_apply:
                    formats_to_apply.append(extension)

        self._catalog_log_context = {
            "isbn": isbn,
            "catalog_title": title,
            "folder_details": {
                key: folder.get(key)
                for key in (
                    "title",
                    "modified_by",
                    "format",
                    "book_title",
                    "author",
                    "publisher",
                    "edition",
                    "is_folder",
                )
            },
        }

        if isbn != folder["title"]:
            if not self.rename_sharepoint_folder(driver, wait, folder, isbn):
                self.logging_service.write(
                    f"Skipping catalog updates for {folder['title']!r}; "
                    f"SharePoint could not set the folder name to ISBN-13 {isbn!r}.",
                    level=logging.WARNING,
                )
                self._catalog_log_context = None
                return False

            # The rename function exits grid view to submit and check the ISBN.
            # Re-enter grid view on the renamed folder before editing metadata.
            if not self.enter_grid_view(driver, wait, folder["title"]):
                self._catalog_log_context = None
                return False

        def attempt_cell_update(label, update, max_attempts=1):
            for attempt in range(1, max_attempts + 1):
                try:
                    update()

                    if attempt > 1:
                        self.logging_service.write(
                            f"Updated {label} for folder {folder.get('title')!r} "
                            f"on attempt {attempt}/{max_attempts}."
                        )
                    return True

                except (TimeoutException, TimeoutError, WebDriverException) as exc:
                    if attempt < max_attempts:
                        self.logging_service.write(
                            f"Could not update {label} for folder "
                            f"{folder.get('title')!r} on attempt "
                            f"{attempt}/{max_attempts}: "
                            f"{type(exc).__name__}: {exc}. Retrying.",
                            level=logging.WARNING,
                        )
                        try:
                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        except WebDriverException:
                            pass

                        time.sleep(attempt)
                        continue

                    self.logging_service.write(
                        f"Could not update {label} for folder "
                        f"{folder.get('title')!r} after "
                        f"{max_attempts} attempt(s): "
                        f"{type(exc).__name__}: {exc}. "
                        "Continuing with the remaining fields.",
                        level=logging.ERROR,
                    )
                    try:
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    except WebDriverException:
                        pass
                    return False

            return False

        self.logging_service.write(
            f"Applying detected formats for ISBN {isbn!r}: ",
            formats_to_apply,
        )
        attempt_cell_update(
            "Formats",
            lambda: self.update_folder_formats(
                driver,
                wait,
                folder,
                formats_to_apply,
            ),
        )

        author_entries = catalog.get("authors") or []
        if not isinstance(author_entries, (list, tuple)):
            author_entries = [author_entries]

        def clean_author_credit(value):
            return re.sub(
                r"\[\s*edited\s+by\s*\]\s*",
                "",
                str(value or ""),
                flags=re.IGNORECASE,
            ).strip()

        author_names = []
        seen_names = set()

        # Limit author lookups to the first four author entries.
        for author in author_entries[:4]:
            name = None

            if isinstance(author, dict):
                name = author.get("name") or author.get("personal_name")

                if not name and author.get("key"):
                    author_record = self.api_service.lookup_author(author["key"])
                    if author_record:
                        name = author_record.get("name") or author_record.get(
                            "personal_name"
                        )
            elif author is not None:
                name = str(author)

            name = clean_author_credit(name)
            if not name:
                continue

            dedupe_key = name.casefold()
            if dedupe_key not in seen_names:
                seen_names.add(dedupe_key)
                author_names.append(name)

        if author_names:
            authors = "; ".join(author_names)
        else:
            authors = catalog.get("by_statement")
            if not authors:
                authors = catalog.get("contributions")

            if isinstance(authors, (list, tuple)):
                authors = "; ".join(
                    str(author).strip() for author in authors if str(author).strip()
                )

            authors = clean_author_credit(authors) if authors else None

        publishers = catalog.get("publishers") or catalog.get("publisher")
        if isinstance(publishers, (list, tuple)):
            publishers = "; ".join(
                str(publisher).strip()
                for publisher in publishers
                if str(publisher).strip()
            )

        # Edition is optional. Use a detected title edition or explicit catalog metadata.
        edition = (
            getattr(self, "_detected_edition", None)
            or catalog.get("edition_name")
            or catalog.get("edition")
        )

        field_attempts = {
            "book_title": 2,
            "author": 2,
            "publisher": 2,
            "edition": 2,
        }

        for field_name, value in (
            ("book_title", title),
            ("author", authors),
            ("publisher", publishers),
            ("edition", edition),
        ):
            if value:
                attempt_cell_update(
                    field_name,
                    lambda field_name=field_name, value=value: (
                        self.update_folder_text_field(
                            driver,
                            wait,
                            folder,
                            field_name,
                            value,
                        )
                    ),
                    max_attempts=field_attempts[field_name],
                )
            elif field_name == "author":
                self.logging_service.write(
                    f"No author value was available to write for folder "
                    f"{folder.get('title')!r}.",
                    level=logging.WARNING,
                )

        # Submit the metadata edits by exiting grid view after the last cell.
        folder = self.get_folder_by_title(driver, folder["title"]) or folder
        folder_title_element = folder["_elements"]["title"]

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            folder_title_element,
        )
        ActionChains(driver).move_to_element(folder_title_element).pause(0.3).perform()

        folder_row = folder_title_element.find_element(
            By.XPATH,
            "./ancestor::*[@role='row'][1]",
        )
        more_actions = wait.until(
            lambda d: next(
                (
                    button
                    for button in folder_row.find_elements(
                        By.CSS_SELECTOR,
                        'button[data-automationid="moreActionsHeroField"]',
                    )
                    if button.is_displayed() and button.is_enabled()
                ),
                False,
            )
        )
        driver.execute_script("arguments[0].click();", more_actions)

        exit_grid_locator = (
            By.CSS_SELECTOR,
            'button[data-automationid="exitGridCommand"]',
        )
        exit_grid_button = wait.until(
            EC.visibility_of_element_located(exit_grid_locator)
        )
        driver.execute_script("arguments[0].click();", exit_grid_button)
        wait.until(EC.invisibility_of_element_located(exit_grid_locator))

        self.logging_service.write(
            f"Exited grid view for folder {isbn!r}; "
            "continuing on the current listing."
        )
        self.logging_service.write(
            f"Post-update verification skipped for folder "
            f"{folder.get('title')!r} (ISBN {isbn!r})."
        )

        self._catalog_log_context = None
        return True

    def download_sharepoint_file(self, driver, wait, item):
        with self.logging_service.spinner(
            f"Downloading {item.get('title') or 'SharePoint file'}"
        ):
            download_wait = WebDriverWait(driver, 35)
            before = set(self.download_dir.iterdir())
            title_element = item["_elements"]["title"]

            download_wait.until(EC.visibility_of(title_element))
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", title_element
            )
            ActionChains(driver).move_to_element(title_element).pause(0.3).perform()

            row = title_element.find_element(
                By.XPATH,
                "./ancestor::*[@role='row'][1]",
            )
            more_actions = row.find_element(
                By.CSS_SELECTOR,
                'button[data-automationid="moreActionsHeroField"]',
            )

            download_wait.until(EC.visibility_of(more_actions))
            driver.execute_script("arguments[0].click();", more_actions)

            download_button_locator = (
                By.CSS_SELECTOR,
                'button[data-automationid="downloadCommand"]',
            )
            download_button = download_wait.until(
                EC.visibility_of_element_located(download_button_locator)
            )
            driver.execute_script("arguments[0].click();", download_button)

            def new_download_files():
                return [
                    path
                    for path in self.download_dir.iterdir()
                    if path not in before and path.is_file()
                ]

            deadline = time.monotonic() + 180
            while time.monotonic() < deadline:
                run_download_files = new_download_files()
                partial_download = any(
                    path.name.endswith((".crdownload", ".tmp"))
                    for path in run_download_files
                )
                completed_files = [
                    path
                    for path in run_download_files
                    if not path.name.endswith((".crdownload", ".tmp"))
                ]

                if completed_files and not partial_download:
                    return max(
                        completed_files,
                        key=lambda path: path.stat().st_mtime,
                    )

                time.sleep(0.5)

            # These files were created during this download attempt and are
            # confined to this run's dedicated directory.
            for path in new_download_files():
                try:
                    path.unlink()
                    self.logging_service.write(
                        f"Removed incomplete SharePoint download: {path.name}",
                        level=logging.WARNING,
                    )
                except OSError as exc:
                    self.logging_service.write(
                        f"Could not remove incomplete SharePoint download "
                        f"{path}: {exc}",
                        level=logging.WARNING,
                    )

            raise TimeoutError(
                f"SharePoint did not finish downloading {item['title']!r} "
                "within 180 seconds."
            )

    def __init__(self):
        run_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        self.download_dir = Path.home() / "Downloads" / "sharepoint_textbook_cataloging"
        log_path = Path(__file__).with_name(f"textbook_cataloging_{run_timestamp}.log")
        self.logging_service = LogService(log_path)
        self.api_service = ApiService(logging_service=self.logging_service)
        self.doc_scan_service = DocScanService(logging_service=self.logging_service)
        self.skipped_folder_renames = 0
        driver = None
        processed_titles = set()

        def valid_google_books_result(result):
            if (
                not isinstance(result, dict)
                or not str(result.get("title") or "").strip()
            ):
                return False

            identifiers = result.get("industryIdentifiers")
            if not isinstance(identifiers, (list, tuple)):
                return False

            for item in identifiers:
                if not isinstance(item, dict):
                    continue

                identifier_type = item.get("type")
                identifier = re.sub(
                    r"[^0-9Xx]", "", str(item.get("identifier") or "")
                ).upper()

                if identifier_type == "ISBN_13":
                    if (
                        len(identifier) == 13
                        and identifier.startswith(("978", "979"))
                        and identifier.isdigit()
                    ):
                        checksum = sum(
                            int(digit) * (1 if index % 2 == 0 else 3)
                            for index, digit in enumerate(identifier[:12])
                        )
                        if (10 - checksum % 10) % 10 == int(identifier[-1]):
                            return True

                elif identifier_type == "ISBN_10":
                    if (
                        len(identifier) == 10
                        and identifier[:9].isdigit()
                        and re.fullmatch(r"\d{9}[\dX]", identifier)
                    ):
                        checksum = sum(
                            (10 - index) * int(digit)
                            for index, digit in enumerate(identifier[:9])
                        )
                        checksum += 10 if identifier[-1] == "X" else int(identifier[-1])
                        if checksum % 11 == 0:
                            return True

            return False

        try:
            self.download_dir.mkdir(parents=True, exist_ok=True)

            options = Options()
            options.add_experimental_option(
                "prefs",
                {
                    "download.default_directory": str(self.download_dir.resolve()),
                    "download.prompt_for_download": False,
                },
            )

            driver = webdriver.Edge(options=options)
            wait = WebDriverWait(driver, 10)

            screen_width = ctypes.windll.user32.GetSystemMetrics(0)
            screen_height = ctypes.windll.user32.GetSystemMetrics(1)

            driver.get(SHAREPOINT_URL)
            driver.set_window_rect(
                x=0,
                y=0,
                width=screen_width,
                height=screen_height,
            )
            self.logging_service.write("\nEdge opened SharePoint folder.")

            see_more = (
                By.CSS_SELECTOR,
                'div[role="gridcell"][class*="seeMoreCell"] button',
            )
            self.click_safely(driver, wait, see_more)

            for attempt in range(1, self.MAX_FOLDERS_TO_RENAME + 1):
                self.logging_service.write(
                    f"\nStarting folder pass {attempt} of "
                    f"{self.MAX_FOLDERS_TO_RENAME}."
                )

                book_folder = self.open_first_eligible_folder(
                    driver,
                    wait,
                    excluded_titles=processed_titles,
                )
                if not book_folder:
                    self.logging_service.write("No more eligible folders to process.")
                    break

                processed_titles.add(str(book_folder.get("title") or "").casefold())
                result = self.inspect_open_folder(driver, wait, book_folder)

                detected_formats = list(
                    getattr(self, "_detected_formats", [])
                    or book_folder.get("file_extensions")
                    or book_folder.get("format")
                    or []
                )
                self.logging_service.write(
                    "Detected formats to carry back to the parent folder:",
                    detected_formats,
                )
                self.logging_service.write("Catalog lookup result:", result)

                driver.get(SHAREPOINT_URL)
                self.click_safely(driver, wait, see_more)
                self.wait_for_folder_listing(driver)
                wait.until(
                    lambda d: len(
                        d.find_elements(
                            By.CSS_SELECTOR,
                            '[role="row"][data-automationid^="row-"]',
                        )
                    )
                    > self.MINIMUM_ROWS_TO_LOAD
                )
                parent_folder = self.get_folder_by_title(
                    driver,
                    book_folder["title"],
                )

                open_library_result = isinstance(result, dict) and str(
                    result.get("key", "")
                ).startswith("/books/")

                loc_result = (
                    not open_library_result
                    and isinstance(result, dict)
                    and bool(str(result.get("title") or "").strip())
                    and bool(result.get("isbn_13"))
                )

                google_books_result = (
                    not open_library_result
                    and not loc_result
                    and self.api_service.GOOGLE_BOOKS_FALLBACK_ENABLED
                    and valid_google_books_result(result)
                )

                self.logging_service.write(
                    "Open Library result:",
                    open_library_result,
                )
                self.logging_service.write(
                    "LOC result:",
                    loc_result,
                )
                self.logging_service.write(
                    "Google Books result:",
                    google_books_result,
                )

                if parent_folder and (
                    open_library_result or loc_result or google_books_result
                ):
                    parent_folder["format"] = detected_formats
                    parent_folder["file_extensions"] = list(
                        book_folder.get("file_extensions") or []
                    )
                    original_title = parent_folder.get("title")
                    update_succeeded = self.update_folder_from_catalog(
                        driver,
                        wait,
                        parent_folder,
                        result,
                    )

                    if not update_succeeded:
                        current_parent_title = str(
                            parent_folder.get("title") or ""
                        ).casefold()
                        if current_parent_title:
                            processed_titles.add(current_parent_title)

                    if parent_folder.get("title") == original_title:
                        self.skipped_folder_renames += 1
                        self.logging_service.write(
                            f"Folder rename skipped for {original_title!r}."
                        )
                else:
                    self.skipped_folder_renames += 1
                    self.logging_service.write(
                        f"Folder rename skipped for "
                        f"{book_folder.get('title')!r}: no eligible catalog "
                        f"result or parent folder row."
                    )

        except (TimeoutException, TimeoutError) as exc:
            catalog_context = getattr(self, "_catalog_log_context", None)
            if catalog_context:
                self.logging_service.write(
                    f"SharePoint processing timed out; ISBN, title, and folder "
                    f"details: {catalog_context!r}",
                    level=logging.ERROR,
                )

            timeout_detail = str(exc).strip()
            if not timeout_detail:
                timeout_detail = (
                    "The Selenium wait expired without providing more details."
                )

            self.logging_service.write(
                f"Timed out waiting for SharePoint "
                f"({type(exc).__name__}): {timeout_detail}",
                level=logging.ERROR,
            )
            self.logging_service.write("Closing Edge and stopping the script.")
            input("\nPress Enter to continue...")

        except Exception as exc:
            catalog_context = getattr(self, "_catalog_log_context", None)
            if catalog_context:
                self.logging_service.write(
                    f"Catalog update failed; ISBN, title, and folder details: "
                    f"{catalog_context!r}",
                    level=logging.ERROR,
                )
            self.logging_service.exception(
                f"Unexpected error ({type(exc).__name__}): {exc}"
            )

        finally:
            self.logging_service.write(
                f"Total folder renames skipped: {self.skipped_folder_renames}"
            )

            if driver:
                input("\nPress Enter to close Edge and finish the script...")
                driver.quit()

            try:
                self.download_dir.rmdir()
                self.logging_service.write(
                    f"Removed empty run-specific download directory: "
                    f"{self.download_dir}"
                )
            except FileNotFoundError:
                pass
            except OSError as exc:
                self.logging_service.write(
                    f"Run-specific download directory was not empty or could "
                    f"not be removed: {self.download_dir}: {exc}",
                    level=logging.WARNING,
                )


def main():
    sid = os.getenv("SID", "").strip()
    if not sid or Path(sid).name != sid:
        raise SystemExit("Set SID in .env to your Windows user-folder name.")

    google_books_fallback_value = (
        os.getenv("GOOGLE_BOOKS_FALLBACK", "false").strip().casefold()
    )
    if google_books_fallback_value in {"1", "true", "yes", "on"}:
        google_books_fallback = True
    elif google_books_fallback_value in {"", "0", "false", "no", "off"}:
        google_books_fallback = False
    else:
        raise SystemExit("GOOGLE_BOOKS_FALLBACK in .env must be true or false.")

    if google_books_fallback and not os.getenv("GOOGLE_BOOKS_API_KEY", "").strip():
        raise SystemExit(
            "GOOGLE_BOOKS_FALLBACK is enabled, but GOOGLE_BOOKS_API_KEY "
            "is not set in .env."
        )

    try:
        folder_rename_cap = int(os.getenv("FOLDER_RENAME_CAP", "50").strip())
    except ValueError:
        raise SystemExit("FOLDER_RENAME_CAP in .env must be a whole number.")

    tesseract_paths = [
        Path("C:/Users") / sid / "AppData/Local/Programs/Tesseract-OCR/tesseract.exe",
        Path("C:/Users") / sid / "AppData/Local/Tesseract-OCR/tesseract.exe",
        Path("C:/Program Files/Tesseract-OCR/tesseract.exe"),
    ]

    tesseract_path = next(
        (path for path in tesseract_paths if path.is_file()),
        None,
    )

    if tesseract_path is None:
        checked_paths = ", ".join(str(path) for path in tesseract_paths)
        raise SystemExit(f"Tesseract executable not found. Checked: {checked_paths}")

    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)

    TextbookCatalogingService.MAX_FOLDERS_TO_RENAME = max(1, min(folder_rename_cap, 50))
    ApiService.GOOGLE_BOOKS_FALLBACK_ENABLED = google_books_fallback

    TextbookCatalogingService()


if __name__ == "__main__":
    main()
