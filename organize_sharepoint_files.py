import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import ctypes
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import requests
import os
from selenium.common.exceptions import StaleElementReferenceException
import re
from pathlib import Path

class TextbookCatalogingService:
  SHAREPOINT_URL = 'https://csulb.sharepoint.com/sites/DSA-BMAC/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FDSA-BMAC%2FShared%20Documents%2F00%20Programs%20Workflow%2FMaster%20Documents%2FATP%2FAlternative%20Formats%2FTextbooks%2FArchives&sortField=BookTitle&isAscending=true'
  SUPPORTED_EXTENSIONS = {"pdf", "tif", "tiff", "epub", "docx"}
  ISBN_PATTERN = re.compile(
    r"(?<!\d)(?:97[89][ -]?)?(?:\d[ -]?){9}[\dXx](?!\d)"
  )

  def lookup_isbn(self, isbn):
    url = f"https://openlibrary.org/isbn/{isbn}.json"
    response = requests.get(url, timeout=10)

    if response.status_code == 404:
      print(f"ISBN {isbn} not found in Open Library.")
      return None  # ISBN not found in Open Library

    response.raise_for_status()
    return response.json()

  def inspect_open_folder(self, driver, wait, current_folder):
    print(f"\nInspecting folder: {current_folder}")
    # Get all items shown in the currently open folder

    folders = self.get_folders_to_update(driver)

    first_folder_with_isbn_title = next(
      (
        folder for folder in folders
        if self.ISBN_PATTERN.search(folder.get("title") or "")
      ),
      None
    )

    if first_folder_with_isbn_title:
      match = self.ISBN_PATTERN.search(
          first_folder_with_isbn_title.get("title") or ""
      )
      isbn = re.sub(r"[^0-9Xx]", "", match.group()).upper()
    else:
      isbn = None

    if isbn:
      return self.lookup_isbn(isbn)

    #if no ISBN found, dive into the first folder that has a supported extension in its title or format to look for front_matter files
    first_folder = next(
      (
        folder for folder in folders
        if folder.get("is_folder")
        and (
          (folder.get("format") or "").strip().casefold()
            in self.SUPPORTED_EXTENSIONS
          or (folder.get("title") or "").strip().casefold()
            in self.SUPPORTED_EXTENSIONS
          or os.path.splitext(
            (folder.get("title") or "").strip()
          )[1].casefold().lstrip(".") in self.SUPPORTED_EXTENSIONS
        )
      ),
      None
    )

    if first_folder:
      print("Opening:", first_folder["title"])
      self.open_folder(driver, wait, first_folder)

      contents = self.get_folders_to_update(driver)

      # Look for an ISBN in a file title.
      isbn_file = next(
        (
          item for item in contents
          if not item.get("is_folder")
          and self.ISBN_PATTERN.search(item.get("title") or "")
        ),
        None
      )
      print("ISBN FOUDN IN FILE TITLE:", isbn_file["title"] if isbn_file else "None")

    if isbn_file:
      match = self.ISBN_PATTERN.search(isbn_file["title"])
      isbn = re.sub(r"[^0-9Xx]", "", match.group()).upper()
      return self.lookup_isbn(isbn)

    # No ISBN in a file title: look for Front Matter, then Matter.
    matter_file = next(
      (
        item for item in contents
        if not item.get("is_folder")
        and ("front" in (item.get("title") or "").casefold() and "matter" in (item.get("title") or "").casefold())
      ),
      None
    )

    if not matter_file:
      matter_file = next(
        (
          item for item in contents
          if not item.get("is_folder")
          and "matter" in (item.get("title") or "").casefold()
        ),
        None
      )

    if matter_file:
      print("Found a possible front matter file:", matter_file["title"])
      wait.until(
        EC.element_to_be_clickable(matter_file["_elements"]["title"])
      ).click()

    return current_folder

  def open_folder(self, driver, wait, folder):
    old_row = folder["_elements"]["row"]
    old_url = driver.current_url

    wait.until(
        EC.element_to_be_clickable(folder["_elements"]["title"])
    ).click()

    wait.until(EC.any_of(
        EC.staleness_of(old_row),
        EC.url_changes(old_url)
    ))

  def open_first_eligible_folder(self, driver, wait):
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
      # clear the profile card that appears on hover
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
        "return (arguments[0].innerText || arguments[0].textContent || '').trim();",
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
        '[data-automationid="field-LinkFilename"] [data-id="heroField"]',
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
        '[data-automationid="field-DocIcon"] [aria-label="Yellow folder"]'
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

  def download_matter_file(self, wait, matter_file):
    before = set(self.download_dir.iterdir())

    # Open More Actions for this specific file row.
    more_actions = matter_file["_elements"]["row"].find_element(
        By.CSS_SELECTOR,
        'button[data-automationid="moreActionsHeroField"]'
    )
    wait.until(EC.element_to_be_clickable(more_actions)).click()

    # Click Download in the menu.
    download_button = (
        By.CSS_SELECTOR,
        'button[data-automationid="downloadCommand"]'
    )
    wait.until(EC.element_to_be_clickable(download_button)).click()

    # Wait for the new download to finish.
    deadline = time.monotonic() + 90
    while time.monotonic() < deadline:
        current_files = list(self.download_dir.iterdir())

        partial_download = any(
            path.name.endswith((".crdownload", ".tmp"))
            for path in current_files
        )
        new_files = [
            path for path in current_files
            if path not in before
            and path.is_file()
            and not path.name.endswith((".crdownload", ".tmp"))
        ]

        if new_files and not partial_download:
            return max(new_files, key=lambda path: path.stat().st_mtime)

        time.sleep(0.5)

    raise TimeoutError("The file did not finish downloading.")

  def __init__(self):
    self.download_dir = Path(__file__).with_name("downloads")
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
      wait = WebDriverWait(driver, 60)

      driver.set_window_rect(
        x=0,
        y=0,
        width=screen_width,
        height=screen_height
      )
      print("\nEdge opened your SharePoint folder.")
      # print("Sign in with your CSULB account if prompted.")
      # input("Once the folder contents are visible, press Enter here...")

      # sort from A-Z
      header_menu = (
        By.CSS_SELECTOR,
        'div[role="columnheader"][data-automationid="field-BookTitle"] '
        'div[role="button"][aria-haspopup="menu"]'
      )

      ## NOTE: add back if URL sort not possible
      # self.click_safely(driver, wait, header_menu)
      # sort_ascending = (By.CSS_SELECTOR, 'button[data-automationid="sortAsc"]')
      # self.click_safely(driver, wait, sort_ascending)

      # load more folders
      see_more = (
        By.CSS_SELECTOR,
        'div[role="gridcell"][class*="seeMoreCell"] button'
      )
      self.click_safely(driver, wait, see_more)

      first_folder = self.open_first_eligible_folder(driver, wait)
      self.inspect_open_folder(driver, wait, first_folder)
      input("\nPress Enter to close Edge...")

    finally:
      driver.quit()

# init
if __name__ == "__main__":
  TextbookCatalogingService()