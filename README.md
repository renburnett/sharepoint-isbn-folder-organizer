# SharePoint Textbook Cataloging

> **Note:** Using a WebDriver and browser to update these fields is _not_ best practice. This tool is a proof of concept. Ideally, we would have direct access to the BMAC SharePoint API and an approved CSULB BMAC account to access the Google Books API.

This tool scans eligible textbook folders in the SharePoint document library archive. It looks for ISBNs, retrieves book details, and updates the folder name and catalog fields in SharePoint. It uses Open Library first, then the Library of Congress. Google Books is an optional fallback.

The tool changes SharePoint folder names and fields. Confirm that the Edge window is signed in to the correct account and site before allowing it to run.

## Before you start

- Use a Windows computer with Microsoft Edge, Python, and access to the SharePoint library.
- Keep `organize_sharepoint_files.py` and the `utils` folder together. The `utils` folder must contain `api_service.py` and `logging_service.py`.
- Install Tesseract OCR. The script builds its expected path from `SID` in `.env`:
  `C:\Users\<SID>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe`
  [INSTALL https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/tesseract-ocr/tesseract/releases/download/5.5.3/tesseract-ocr-w64-setup-5.5.3.20260724.exe)
  If Tesseract is installed elsewhere, update the path in `main()`.
- Install Microsoft Word and `pywin32` to scan old `.doc` and `.rtf` files. Word's Trust Center settings may block some `.doc` files. `.docx` files do not need Word.
- The script can scan PDF, EPUB, DOCX, TIF/TIFF, DOC, RTF, and JPG/JPEG files. It cannot read KES files.
- Tesseract is needed when a PDF has no extractable text and for image-based TIFF/JPG scans. PDF files with embedded text can be read without OCR.

## One-time setup

1. [Download Python 3 for Windows: https://www.python.org/downloads](https://www.python.org/ftp/python/3.15.2/python-3.15.2-amd64.exe)

2. [Download and Unzip these Files https://github.com/renburnett/sharepoint-isbn-folder-organizer](https://github.com/renburnett/sharepoint-isbn-folder-organizer/archive/refs/heads/main.zip)

3. Create a project folder somewhere memorable, such as under Documents. Open PowerShell in the project folder:

```powershell
cd "C:\Users\<YOUR SID>\Documents\Projects\SharePointTextbookCataloging"
python -m venv .venv
.venv\Scripts\activate.bat
```

The virtual environment keeps this project's Python packages separate from other projects on your computer. Create it once, then activate it whenever you open a new PowerShell window to work on this project. When it is active, the prompt starts with `(.venv)`.

You can confirm that the active Python is from the virtual environment:

```powershell
python -c "import sys; print(sys.executable)"
```

The path should end with:

```text
SharePointTextbookCataloging\.venv\Scripts\python.exe
```

With the virtual environment active, install the Python packages:

```powershell
python -m pip install --upgrade pip
python -m pip install requests selenium ebooklib PyMuPDF pytesseract beautifulsoup4 python-docx Pillow pypdf python-dotenv fonttools pywin32
```

Create a `.env` file in the project folder with the settings below. Set `GOOGLE_BOOKS_FALLBACK=true` only if you want to use Google Books; the API key is required when it is enabled.

```env
SID=your_windows_user_folder_name
FOLDER_RENAME_CAP=50
GOOGLE_BOOKS_FALLBACK=false
GOOGLE_BOOKS_API_KEY=your_api_key_here
```

## Run the tool

With the virtual environment active and `.env` configured, run:

```powershell
python organize_sharepoint_files.py
```

The processing limit defaults to 50. Values above 50 are capped at 50. Set `FOLDER_RENAME_CAP` in `.env` to use a smaller limit. Google Books remains disabled unless `GOOGLE_BOOKS_FALLBACK=true` is set in `.env`.

## How ISBN scanning works

The script first checks folder and file names for an ISBN-13. If it does not find one, it scans candidate files, prioritizing names that suggest front matter, end matter, copyright, or title pages. Within each priority group, it tries PDF, EPUB, DOCX, TIF/TIFF, DOC, RTF, and JPG/JPEG. If it finds no ISBN in the prioritized files, it tries at most one remaining file per extension, so it can check an unnamed EPUB or DOCX without downloading every chapter. It prefers an ISBN-13; if none is found, it can use an ISBN-10.

PDF scanning checks up to the first eight pages, or the last eight pages for an end-matter file. TIFF images are OCR-scanned. DOCX and EPUB text are extracted directly. DOC and RTF are read through Microsoft Word automation. KES files are not supported.

After finding an ISBN, the tool looks up catalog information in Open Library. If Open Library does not return a record, it tries the Library of Congress. Google Books is tried only when `GOOGLE_BOOKS_FALLBACK=true` and neither previous source returns a usable result.

## While it runs

Edge opens SharePoint and the script processes folders one at a time. Leave Edge open while it runs. A spinner appears during API requests, downloads, and other longer waits. The script writes timestamped logs in the `log` folder and downloads files to a run-specific folder in your Windows Downloads directory. Temporary downloaded files are removed after scanning. When processing finishes, press Enter in the terminal to close Edge.

The console and log include a final list of folders that could not be renamed.

## If something goes wrong

- **A package is missing:** Activate `.venv` and install the package named in the error.
- **Tesseract cannot be found:** Check that `SID` in `.env` exactly matches your Windows user-folder name and that the expected `tesseract.exe` path exists. If Tesseract is elsewhere, update the path in `main()`.
- **An old `.doc` or `.rtf` file cannot be scanned:** Confirm Microsoft Word and `pywin32` are installed. A Trust Center policy may block `.doc`; if permitted, convert the file to `.docx` or ask your administrator about the policy.
- **SharePoint asks you to sign in:** Sign in to Edge using an account with access to the library.
- **A folder is skipped or a field is not updated:** Check the timestamped log for the folder, ISBN, and error details.
- **A PDF produces `fontTools` warnings:** Confirm `fonttools` is installed in the active virtual environment:
  ```powershell
  python -m pip install fonttools
  ```
- **Hard-quitting the script from the console:** Press `Ctrl+C`.
