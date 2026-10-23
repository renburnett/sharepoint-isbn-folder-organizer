import time
import ctypes
import os
import re
from pathlib import Path

import ebooklib
import pymupdf
import requests
import pytesseract
pytesseract.pytesseract.tesseract_cmd = (
  r"C:\Users\006888191SA\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)

from bs4 import BeautifulSoup
from docx import Document
from ebooklib import epub
from PIL import Image
from pypdf import PdfReader

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
  ElementClickInterceptedException,
  StaleElementReferenceException,
)
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


class TextbookCatalogingService:
  SHAREPOINT_URL = (
    "https://csulb.sharepoint.com/sites/DSA-BMAC/Shared%20Documents/"
    "Forms/AllItems.aspx?id=%2Fsites%2FDSA-BMAC%2FShared%20Documents"
    "%2F00%20Programs%20Workflow%2FMaster%20Documents%2FATP"
    "%2FAlternative%20Formats%2FTextbooks%2FArchives"
    "&sortField=BookTitle&isAscending=true"
  )

  SUPPORTED_EXTENSIONS = {"pdf", "tif", "tiff", "epub", "docx"}

  PHONE_PATTERN = re.compile(
    r"(?<!\d)\(?\d{3}\)?[^\dA-Za-z]+\d{3}[^\dA-Za-z]+\d{4}(?!\d)"
  )

  ISBN_PATTERN = re.compile(
    r"(?<!\d)(?:97[89][^\dA-Za-z]?)?(?:\d[^\dA-Za-z]?){9}[\dXx](?!\d)"
  )

  def lookup_isbn(self, isbn):
    url = f"https://openlibrary.org/isbn/{isbn}.json"
    response = requests.get(url, timeout=10)

    print("---")
    print('lookup_isbn :response.json()', response.json())
    print("---")

    if response.status_code == 404:
      print(f"ISBN {isbn} not found in Open Library.")
      return None

    response.raise_for_status()
    return response.json()

  def get_listing_signature(self, driver):
    return tuple(driver.execute_script("""
      return Array.from(document.querySelectorAll(
        '[role="row"][data-automationid^="row-"] '
        + '[data-automationid="field-LinkFilename"] '
        + '[data-id="heroField"]'
      ))
      .map(el => (
        el.getAttribute('title')
        || el.innerText
        || el.textContent
        || ''
      ).trim())
      .filter(Boolean);
    """))

  def wait_for_folder_listing(
      self, driver, previous_signature=None, timeout=60
  ):
    deadline = time.monotonic() + timeout
    last_signature = None
    stable_since = None

    while time.monotonic() < deadline:
      signature = self.get_listing_signature(driver)

      page_state = driver.execute_script("""
        const body = (document.body.innerText || '').toLowerCase();

        const busyElement = document.querySelector(
          '[aria-busy="true"], [role="progressbar"], '
          + '[data-automationid*="loading"]'
        );

        const busy = Boolean(
          busyElement && busyElement.getClientRects().length
        );

        const empty = [
          'this folder is empty',
          'no items found',
          'no files found',
          'no items to show',
          'drag files here'
        ].some(phrase => body.includes(phrase));

        return { busy, empty };
      """)

      has_results_or_empty_message = (
        bool(signature) or page_state["empty"]
      )

      changed_from_parent = (
        previous_signature is None
        or not previous_signature
        or signature != previous_signature
      )

      if (
        has_results_or_empty_message
        and changed_from_parent
        and not page_state["busy"]
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

      time.sleep(0.4)

    raise TimeoutError(
      "The SharePoint folder listing did not finish loading."
    )

  def inspect_open_folder(self, driver, wait, current_folder):
    print(f"\nInspecting folder: {current_folder['title']}")

    # Get the items shown in the currently open folder.
    folders = self.get_folders_to_update(driver)

    # First check files in this folder for an ISBN in the filename.
    file_with_isbn = next(
      (
        item for item in folders
        if not item.get("is_folder")
        and self.ISBN_PATTERN.search(item.get("title") or "")
      ),
      None
    )

    if file_with_isbn:
      match = self.ISBN_PATTERN.search(file_with_isbn["title"])
      isbn = re.sub(r"[^0-9Xx]", "", match.group()).upper()
      return self.lookup_isbn(isbn)

    def normalized_title(item):
      return (
        (item.get("title") or "")
        .casefold()
        .replace("_", " ")
        .replace("-", " ")
      )

    # Prefer Split PDF over a folder named TIF, PDF, EPUB, etc.
    split_pdf_folder = next(
      (
        item for item in folders
        if item.get("is_folder")
        and "split" in normalized_title(item)
        and "pdf" in normalized_title(item)
      ),
      None
    )

    # If Split PDF isn't present, use the first supported-format folder.
    supported_folder = next(
      (
        item for item in folders
        if item.get("is_folder")
        and (
          (item.get("format") or "").strip().casefold()
            in self.SUPPORTED_EXTENSIONS
          or normalized_title(item).strip()
            in self.SUPPORTED_EXTENSIONS
          or os.path.splitext(
            normalized_title(item)
          )[1].casefold().lstrip(".")
            in self.SUPPORTED_EXTENSIONS
        )
      ),
      None
    )

    first_folder = split_pdf_folder or supported_folder

    if not first_folder:
      print("No Split PDF folder or supported-format folder was found.")
      return current_folder

    print("Opening:", first_folder["title"])
    self.open_folder(driver, wait, first_folder)

    contents = self.get_folders_to_update(driver)

    # If Split PDF was opened, look for and open its Front Matter subfolder.
    if first_folder is split_pdf_folder:
      front_matter_folder = next(
        (
          item for item in contents
          if item.get("is_folder")
          and "front" in normalized_title(item)
          and "matter" in normalized_title(item)
        ),
        None
      )

      if front_matter_folder:
        print("Opening:", front_matter_folder["title"])
        self.open_folder(driver, wait, front_matter_folder)
        contents = self.get_folders_to_update(driver)

    # Check files inside the opened folder for an ISBN in the filename.
    isbn_file = next(
      (
        item for item in contents
        if not item.get("is_folder")
        and self.ISBN_PATTERN.search(item.get("title") or "")
      ),
      None
    )

    print(
      "ISBN found in file title:",
      isbn_file["title"] if isbn_file else "None"
    )

    if isbn_file:
      match = self.ISBN_PATTERN.search(isbn_file["title"])
      isbn = re.sub(r"[^0-9Xx]", "", match.group()).upper()
      return self.lookup_isbn(isbn)

    # Look for a front-matter file first.
    front_matter_file = next(
      (
        item for item in contents
        if not item.get("is_folder")
        and "front" in normalized_title(item)
        and "matter" in normalized_title(item)
      ),
      None
    )

    end_matter_file = next(
      (
        item for item in contents
        if not item.get("is_folder")
        and (
          (
            "end" in normalized_title(item)
            and "matter" in normalized_title(item)
          )
          or (
            "back" in normalized_title(item)
            and "matter" in normalized_title(item)
          )
        )
      ),
      None
    )

    matter_file = front_matter_file or end_matter_file

    # If neither was named explicitly, look for any file with "matter".
    if not matter_file:
      matter_file = next(
        (
          item for item in contents
          if not item.get("is_folder")
          and "matter" in normalized_title(item)
        ),
        None
      )

    if not matter_file:
      print("No front-matter or end-matter file was found.")
      return current_folder

    print("Found a possible matter file:", matter_file["title"])

    # Download the file using its row's More Actions menu.
    downloaded_path = self.download_sharepoint_file(
      driver, wait, matter_file
    )
    print("Downloaded:", downloaded_path)

    title = normalized_title(matter_file)
    matter_type = (
      "end"
      if (
        ("end" in title and "matter" in title)
        or ("back" in title and "matter" in title)
      )
      else "front"
    )

    isbn = self.find_isbn_in_downloaded_file(
      downloaded_path,
      matter_type
    )

    if isbn:
      print("Found ISBN:", isbn)
      return self.lookup_isbn(isbn)

    print("No ISBN found in the selected pages.")
    return None

  def open_folder(self, driver, wait, folder):
    old_row = folder["_elements"]["row"]
    old_url = driver.current_url
    old_signature = self.get_listing_signature(driver)

    wait.until(
      EC.element_to_be_clickable(folder["_elements"]["title"])
    ).click()

    wait.until(EC.any_of(
      EC.staleness_of(old_row),
      EC.url_changes(old_url)
    ))

    self.wait_for_folder_listing(
      driver,
      previous_signature=old_signature
    )

  def open_first_eligible_folder(self, driver, wait):
    self.wait_for_folder_listing(driver)

    wait.until(lambda d: len(d.find_elements(
      By.CSS_SELECTOR,
      '[role="row"][data-automationid^="row-"]'
    )) > 12)

    items = self.get_folders_to_update(driver)
    excluded_names = ("abueg", "izaiah", "burnett", "-sa")

    for item in items:
      modified_by = (item.get("modified_by") or "").casefold()
      formats = (item.get("format") or "").strip()
      book_title = (item.get("book_title") or "").strip()

      excluded = any(name in modified_by for name in excluded_names)

      if (
        item.get("is_folder")
        and not excluded
        and not formats
        and not book_title
      ):
        self.open_folder(driver, wait, item)
        return item

    print("No folder matches those conditions.")
    return None

  def click_safely(self, driver, wait, locator):
    try:
      wait.until(EC.element_to_be_clickable(locator)).click()
    except ElementClickInterceptedException:
      # Clear the profile card that can appear on hover.
      ActionChains(driver).send_keys(Keys.ESCAPE).perform()

      wait.until(EC.invisibility_of_element_located((
        By.CSS_SELECTOR,
        'lpc-card[role="dialog"][aria-label="Profile Card"]'
      )))

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
        element
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
      By.CSS_SELECTOR,
      '[role="row"][data-automationid^="row-"]'
    )

    print("DOM rows:", len(rows))
    items = []

    for row in rows:
      title, title_element = self.read_cell(
        driver,
        row,
        '[data-automationid="field-LinkFilename"] '
        '[data-id="heroField"]',
        use_title_attribute=True
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

      is_folder = bool(row.find_elements(
        By.CSS_SELECTOR,
        '[data-automationid="field-DocIcon"] '
        '[aria-label="Yellow folder"]'
      ))

      items.append({
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
        }
      })

    print(f"Found {len(items)} item(s).")
    return items

  def download_sharepoint_file(self, driver, wait, item):
    before = set(self.download_dir.iterdir())
    title_element = item["_elements"]["title"]

    wait.until(EC.visibility_of(title_element))
    driver.execute_script(
      "arguments[0].scrollIntoView({block: 'center'});",
      title_element
    )
    ActionChains(driver).move_to_element(title_element).pause(0.3).perform()

    row = title_element.find_element(
      By.XPATH,
      "./ancestor::*[@role='row'][1]"
    )
    more_actions = row.find_element(
      By.CSS_SELECTOR,
      'button[data-automationid="moreActionsHeroField"]'
    )
    print("Clicking More Actions for:", item["title"])
    print('more_actions', more_actions)
    wait.until(EC.visibility_of(more_actions))
    driver.execute_script("arguments[0].click();", more_actions)

    download_button_locator = (
        By.CSS_SELECTOR,
        'button[data-automationid="downloadCommand"]'
    )
    download_button = wait.until(
        EC.visibility_of_element_located(download_button_locator)
    )
    driver.execute_script("arguments[0].click();", download_button)

    deadline = time.monotonic() + 90

    while time.monotonic() < deadline:
      new_files = [
        path for path in self.download_dir.iterdir()
        if path not in before and path.is_file()
      ]

      partial_download = any(
        path.name.endswith((".crdownload", ".tmp"))
        for path in new_files
      )
      completed_files = [
        path for path in new_files
        if not path.name.endswith((".crdownload", ".tmp"))
      ]

      if completed_files and not partial_download:
        return max(
          completed_files,
          key=lambda path: path.stat().st_mtime
        )

      time.sleep(0.5)

    raise TimeoutError("The SharePoint file did not finish downloading.")

  def find_isbn_in_downloaded_file(self, file_path, matter_type):
    extension = file_path.suffix.casefold()
    text_parts = []

    if extension == ".pdf":
      reader = PdfReader(str(file_path))
      pdf_for_ocr = pymupdf.open(str(file_path))

      if matter_type == "end":
        start = max(0, len(reader.pages) - 8)
        page_numbers = range(start, len(reader.pages))
      else:
        page_numbers = range(min(8, len(reader.pages)))

      for page_number in page_numbers:
        text = reader.pages[page_number].extract_text() or ""

        # OCR the page if it has no extractable text.
        if not text.strip():
          pixmap = pdf_for_ocr[page_number].get_pixmap(dpi=200)
          image = Image.frombytes(
            "RGB",
            (pixmap.width, pixmap.height),
            pixmap.samples
          )
          text = pytesseract.image_to_string(image)

        text_parts.append(text)

    elif extension in {".tif", ".tiff"}:
      with Image.open(file_path) as image:
        frame_count = getattr(image, "n_frames", 1)

        if matter_type == "end":
          start = max(0, frame_count - 8)
          frame_numbers = range(start, frame_count)
        else:
          frame_numbers = range(min(8, frame_count))

        for frame_number in frame_numbers:
          image.seek(frame_number)
          text_parts.append(
            pytesseract.image_to_string(image.copy())
          )

    elif extension == ".docx":
      document = Document(str(file_path))
      text_parts.extend(
        paragraph.text for paragraph in document.paragraphs
      )

      for table in document.tables:
        for row in table.rows:
          text_parts.append(
            " ".join(cell.text for cell in row.cells)
          )

    elif extension == ".epub":
      book = epub.read_epub(str(file_path))

      for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
        soup = BeautifulSoup(item.get_content(), "html.parser")
        text_parts.append(soup.get_text(" ", strip=True))

    else:
      raise ValueError(f"Unsupported file type: {extension}")

    text = "\n".join(text_parts)

    for match in self.ISBN_PATTERN.finditer(text):
      if self.PHONE_PATTERN.fullmatch(match.group()):
        continue

      return re.sub(r"[^0-9Xx]", "", match.group()).upper()

    return None

  def __init__(self):
    # Use the Windows Downloads folder.
    self.download_dir = Path.home() / "Downloads"
    self.download_dir.mkdir(exist_ok=True)

    options = Options()
    options.add_experimental_option("prefs", {
      "download.default_directory": str(self.download_dir.resolve()),
      "download.prompt_for_download": False,
    })

    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 60)

    screen_width = ctypes.windll.user32.GetSystemMetrics(0)
    screen_height = ctypes.windll.user32.GetSystemMetrics(1)

    try:
      driver.get(self.SHAREPOINT_URL)

      driver.set_window_rect(
        x=0,
        y=0,
        width=screen_width,
        height=screen_height
      )

      print("\nEdge opened your SharePoint folder.")

      # Load more folders.
      see_more = (
        By.CSS_SELECTOR,
        'div[role="gridcell"][class*="seeMoreCell"] button'
      )
      self.click_safely(driver, wait, see_more)

      first_folder = self.open_first_eligible_folder(driver, wait)

      if first_folder:
        result = self.inspect_open_folder(
          driver, wait, first_folder
        )
        print("Catalog lookup result:", result)

      input("\nPress Enter to close Edge...")

    finally:
      driver.quit()


if __name__ == "__main__":
  TextbookCatalogingService()