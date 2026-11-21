import ebooklib
import pymupdf
from bs4 import BeautifulSoup
from docx import Document
from ebooklib import epub
from PIL import Image, ImageOps
from pypdf import PdfReader
import pytesseract
import re
from pathlib import Path
from utils.constants import (
    ISBN_13_PATTERN,
    ISBN_10_PATTERN,
    PHONE_PATTERN,
    SCAN_PRIORITY_KEYWORDS,
)


class DocScanService:
    MAX_FRONT_PAGES_TO_SCAN = 10
    MAX_BACK_PAGES_TO_SCAN = 6

    def __init__(self, logging_service=None):
        self.logging_service = logging_service

    def contains_isbn_13(self, text):
        for match in ISBN_13_PATTERN.finditer(text or ""):
            if not PHONE_PATTERN.fullmatch(match.group()):
                return True
        return False

    def contains_valid_isbn(self, text):
        for pattern in (ISBN_13_PATTERN, ISBN_10_PATTERN):
            for match in pattern.finditer(text or ""):
                if not PHONE_PATTERN.fullmatch(match.group()):
                    return True
        return False

    def pages_to_scan(self, page_count, matter_type):
        if matter_type == "end":
            return list(
                range(
                    page_count - 1,
                    max(-1, page_count - self.MAX_BACK_PAGES_TO_SCAN - 1),
                    -1,
                )
            )

        front_pages = list(range(min(self.MAX_FRONT_PAGES_TO_SCAN, page_count)))

        if matter_type != "full":
            return front_pages

        scanned_pages = set(front_pages)
        back_start = max(0, page_count - self.MAX_BACK_PAGES_TO_SCAN)
        back_pages = [
            page_number
            for page_number in range(page_count - 1, back_start - 1, -1)
            if page_number not in scanned_pages
        ]
        return front_pages + back_pages

    def normalized_title(self, item):
        title = item.get("title") if isinstance(item, dict) else item
        return str(title or "").casefold().replace("_", " ").replace("-", " ")

    def title_tokens(self, item):
        return set(re.findall(r"[a-z0-9]+", self.normalized_title(item)))

    def is_front_matter(self, item):
        return bool(self.title_tokens(item) & {"front", "frontmatter", "fm"})

    def is_end_matter(self, item):
        return bool(
            self.title_tokens(item)
            & {
                "end",
                "back",
                "bm",
                "endmatter",
                "backmatter",
                "endnote",
                "endnotes",
            }
        )

    def is_chapter_one(self, item):
        return bool(
            re.search(
                r"\b(?:chapter|ch)\s*0*1\b",
                self.normalized_title(item),
            )
        )

    def matter_priority(self, item):
        tokens = self.title_tokens(item)

        if self.is_front_matter(item):
            return 0
        if self.is_end_matter(item):
            return 1
        if tokens & SCAN_PRIORITY_KEYWORDS:
            return 2
        if self.is_chapter_one(item):
            return 3

        return None

    def matter_type(self, item):
        tokens = self.title_tokens(item)

        if self.is_end_matter(item) or "index" in tokens:
            return "end"

        is_chapter_file = bool(
            re.search(
                r"\b(?:chapter|ch)\s*0*\d+\b",
                self.normalized_title(item),
            )
        )

        if (
            self.is_front_matter(item)
            or is_chapter_file
            or tokens & SCAN_PRIORITY_KEYWORDS
        ):
            return "front"

        return "full"

    def is_matter_file(self, file_path):
        filename_tokens = self.title_tokens(Path(file_path).stem)

        is_front_matter_file = (
            {"front", "matter"}.issubset(filename_tokens)
            or "frontmatter" in filename_tokens
            or "fm" in filename_tokens
        )
        is_end_matter_file = {"end", "matter"}.issubset(
            filename_tokens
        ) or "endmatter" in filename_tokens
        is_back_matter_file = (
            {"back", "matter"}.issubset(filename_tokens)
            or "backmatter" in filename_tokens
            or "bm" in filename_tokens
        )

        return is_front_matter_file or is_end_matter_file or is_back_matter_file

    def find_isbn(self, file_path, matter_type):
        is_matter_file = self.is_matter_file(file_path)

        def first_isbn_in_text(text):
            candidates = []

            for preference, pattern in enumerate((ISBN_13_PATTERN, ISBN_10_PATTERN)):
                for match in pattern.finditer(text or ""):
                    if PHONE_PATTERN.fullmatch(match.group()):
                        continue

                    candidates.append(
                        (
                            match.start(),
                            preference,
                            re.sub(r"[^0-9Xx]", "", match.group()).upper(),
                        )
                    )

            if not candidates:
                return None

            # Return the first ISBN in reading order. If both patterns start
            # at the same position, prefer ISBN-13.
            return min(candidates, key=lambda item: (item[0], item[1]))[2]

        with self.logging_service.spinner(f"Scanning {file_path.name} for an ISBN"):
            extension = file_path.suffix.casefold()
            text_parts = []

            def add_text(text):
                text_parts.append(text or "")
                return first_isbn_in_text("\n".join(text_parts))

            if extension == ".pdf":
                reader = PdfReader(str(file_path))
                pdf_for_ocr = pymupdf.open(str(file_path))

                try:
                    for page_number in self.pages_to_scan(
                        len(reader.pages), matter_type
                    ):
                        page_text = reader.pages[page_number].extract_text() or ""

                        isbn = add_text(page_text)
                        if isbn:
                            return isbn

                        self.logging_service.write(
                            f"PDF page {page_number + 1}: running first OCR pass."
                        )

                        pixmap = pdf_for_ocr[page_number].get_pixmap(dpi=200)
                        image = Image.frombytes(
                            "RGB",
                            (pixmap.width, pixmap.height),
                            pixmap.samples,
                        )
                        first_pass_text = pytesseract.image_to_string(image)

                        isbn = add_text(first_pass_text)
                        if isbn:
                            return isbn

                        if is_matter_file:
                            self.logging_service.write(
                                f"PDF page {page_number + 1}: first pass found "
                                "no ISBN; retrying matter file with enhanced OCR."
                            )

                            enhanced_image = ImageOps.autocontrast(image.convert("L"))
                            enlarged_image = enhanced_image.resize(
                                (
                                    enhanced_image.width * 2,
                                    enhanced_image.height * 2,
                                ),
                                Image.Resampling.LANCZOS,
                            )
                            second_pass_text = pytesseract.image_to_string(
                                enlarged_image,
                                config="--psm 11",
                            )

                            isbn = add_text(second_pass_text)
                            if isbn:
                                self.logging_service.write(
                                    f"PDF page {page_number + 1}: enhanced OCR "
                                    f"found ISBN {isbn}."
                                )
                                return isbn

                            self.logging_service.write(
                                f"PDF page {page_number + 1}: enhanced OCR "
                                "found no ISBN."
                            )
                finally:
                    pdf_for_ocr.close()

            elif extension in {".tif", ".tiff"}:
                with Image.open(file_path) as image:
                    frame_count = getattr(image, "n_frames", 1)

                    for frame_number in self.pages_to_scan(frame_count, matter_type):
                        image.seek(frame_number)
                        isbn = add_text(pytesseract.image_to_string(image.copy()))
                        if isbn:
                            return isbn

            elif extension in {".jpg", ".jpeg"}:
                with Image.open(file_path) as image:
                    return first_isbn_in_text(
                        pytesseract.image_to_string(image.convert("RGB"))
                    )

            elif extension == ".docx":
                document = Document(str(file_path))

                for paragraph in document.paragraphs:
                    isbn = add_text(paragraph.text)
                    if isbn:
                        return isbn

                for table in document.tables:
                    for row in table.rows:
                        isbn = add_text(" ".join(cell.text for cell in row.cells))
                        if isbn:
                            return isbn

            elif extension in {".doc", ".rtf"}:
                try:
                    import win32com.client
                except ImportError as exc:
                    raise RuntimeError(
                        "Reading .doc and .rtf files requires pywin32. "
                        "Install it with: python -m pip install pywin32"
                    ) from exc

                word = None
                document = None

                try:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                    document = word.Documents.Open(
                        FileName=str(Path(file_path).resolve()),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        NoEncodingDialog=True,
                    )
                    return first_isbn_in_text(document.Content.Text)
                finally:
                    if document is not None:
                        try:
                            document.Close(SaveChanges=False)
                        except Exception:
                            pass

                    if word is not None:
                        try:
                            word.Quit(SaveChanges=False)
                        except Exception:
                            pass

            elif extension == ".epub":
                book = epub.read_epub(str(file_path))

                for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
                    soup = BeautifulSoup(
                        item.get_content(),
                        "html.parser",
                    )
                    isbn = add_text(soup.get_text(" ", strip=True))
                    if isbn:
                        return isbn

            else:
                raise ValueError(f"Unsupported file type: {extension}")

            return None
