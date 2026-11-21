import re

PHONE_PATTERN = re.compile(
    r"(?<!\d)\(?\d{3}\)?[^\dA-Za-z]+\d{3}[^\dA-Za-z]+\d{4}(?!\d)"
)
ISBN_13_PATTERN = re.compile(r"(?<!\d)97[89](?:[^\dA-Za-z]?\d){10}(?!\d)")
ISBN_10_PATTERN = re.compile(r"(?<!\d)(?:\d[^\dA-Za-z]?){9}[\dXx](?!\d)")
SHAREPOINT_URL = (
    "https://csulb.sharepoint.com/sites/DSA-BMAC/Shared%20Documents/"
    "Forms/AllItems.aspx?id=%2Fsites%2FDSA-BMAC%2FShared%20Documents"
    "%2F00%20Programs%20Workflow%2FMaster%20Documents%2FATP"
    "%2FAlternative%20Formats%2FTextbooks%2FArchives"
    "&sortField=BookTitle&isAscending=true"
)
SCAN_PRIORITY_KEYWORDS = frozenset(
    {
        "matter",
        "index",
        "title",
        "copyright",
        "isbn",
        "catalog",
        "cataloging",
        "bibliographic",
        "publication",
    }
)
EXTENSION_CONFIG = {
    # extension: (scan_priority, SharePoint format)
    "pdf": (0, "pdf"),
    "epub": (1, "epub"),
    "docx": (2, "docx"),
    "tif": (3, "tif"),
    "tiff": (3, "tif"),
    "doc": (4, "docx"),
    "rtf": (5, "other"),
    "jpg": (6, None),
    "jpeg": (6, None),
}

SCANNABLE_EXTENSIONS = frozenset(EXTENSION_CONFIG)

# Extensions with a SharePoint format mapping are considered supported formats.
SUPPORTED_EXTENSIONS = frozenset(
    extension
    for extension, (_, sharepoint_format) in EXTENSION_CONFIG.items()
    if sharepoint_format is not None
)

EXTENSION_PRIORITY = {
    extension: scan_priority
    for extension, (scan_priority, _) in EXTENSION_CONFIG.items()
}

# Keep the same relative order, omitting scannable-only formats like JPG/JPEG.
FORMAT_PRIORITY = tuple(
    extension
    for extension, (_, sharepoint_format) in sorted(
        EXTENSION_CONFIG.items(),
        key=lambda item: item[1][0],
    )
    if sharepoint_format is not None
)

FORMAT_NAMES = {
    extension: sharepoint_format
    for extension, (_, sharepoint_format) in EXTENSION_CONFIG.items()
    if sharepoint_format is not None
}

FORMAT_NAMES["word"] = "docx"
FORMAT_NAMES["other"] = "other"
