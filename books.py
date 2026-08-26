import os
import re
import json
import time
import shutil
import base64
import random
from io import BytesIO
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from PIL import Image, ImageOps
from openai import OpenAI

# =========================
# CONFIG
# =========================
BASE_FOLDER = Path(r"C:\Users\megha\OneDrive\Pictures\Samsung Gallery\Gallery\Books")
INCOMING_FOLDER = BASE_FOLDER
PROCESSED_FOLDER = BASE_FOLDER / "processed"
SKIPPED_FOLDER = BASE_FOLDER / "skipped"

BOOKS_2_ROOT = Path(r"C:\Users\megha\OneDrive\Pictures\Samsung Gallery\DCIM\Books 2.0")

OUTPUT_XLSX = BASE_FOLDER / "books_output.xlsx"
PROGRESS_CSV = BASE_FOLDER / "_progress_all_books.csv"
ERROR_LOG = BASE_FOLDER / "_errors.csv"
SYNC_LOG = BASE_FOLDER / "_sync_log.csv"

MODEL = "gpt-4.1-mini"
MAX_WORKERS = 4
CONFIDENCE_THRESHOLD = 0.75
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}

# TEST MODE
TEST_LIMIT = 10
RANDOM_SAMPLE = True

# If True, compare files by filename only
COMPARE_BY_FILENAME_ONLY = True

# Metadata enrichment
ENABLE_METADATA_ENRICHMENT = True
ENABLE_ASIN_LOOKUP = True
ENABLE_CONTENT_ENRICHMENT = True

LOOKUP_TIMEOUT_SECONDS = 20
LOOKUP_SLEEP_SECONDS = 0.2
AMAZON_LOOKUP_SLEEP_SECONDS = 0.8
AI_ENRICH_SLEEP_SECONDS = 0.2

TEXT_COLUMNS = [
    "Image", "Title", "Author", "Needs Review",
    "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
    "Description", "Genre", "AgeRange", "Tropes", "Triggers",
    "Metadata Enriched",
    "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed",
    "DuplicateKey", "Last Enriched"
]

ALL_BOOKS_COLUMNS = [
    "Image", "Title", "Author", "Confidence", "Needs Review",
    "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
    "Description", "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
    "Metadata Enriched",
    "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed",
    "DuplicateKey", "Last Enriched"
]

# =========================
# ENRICHMENT SKIP CONFIG
# =========================
# Fields counted when calculating enrichment completeness percentage.
ENRICHMENT_SCORED_FIELDS = [
    "ISBN_13", "ASIN", "Description", "Genre",
    "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
]

# Skip re-enrichment if the row was enriched within this many days.
# Set to 0 to disable the date check entirely.
ENRICHMENT_MIN_AGE_DAYS = 180  # 6 months

# Skip re-enrichment if the row is already this % complete across
# ENRICHMENT_SCORED_FIELDS (0.0–1.0).  Set to 1.0 to disable.
ENRICHMENT_SKIP_THRESHOLD = 0.70  # 70 %

# =========================
# API CLIENT
# =========================
api_key = os.environ.get("OPENAI_API_KEY")
if not api_key:
    raise RuntimeError("OPENAI_API_KEY is not set. Set it first, then run again.")

client = OpenAI(api_key=api_key)

# =========================
# SETUP
# =========================
PROCESSED_FOLDER.mkdir(exist_ok=True)
SKIPPED_FOLDER.mkdir(exist_ok=True)

# =========================
# HELPERS
# =========================
def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_text(value: str) -> str:
    value = clean_text(value).lower()
    value = re.sub(r"[^\w\s]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_csv_list(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    parts = [p.strip() for p in re.split(r"[;,|]+", text) if p.strip()]
    seen = set()
    out = []
    for part in parts:
        key = part.lower()
        if key not in seen:
            seen.add(key)
            out.append(part)
    return ", ".join(out)


def build_duplicate_key(title: str, author: str) -> str:
    return f"{normalize_text(title)}|{normalize_text(author)}"


def needs_review(title: str, author: str, confidence: float) -> str:
    if confidence < CONFIDENCE_THRESHOLD:
        return "Yes"
    if not title or len(title.strip()) < 3:
        return "Yes"
    if not author:
        return "Yes"
    return "No"


def is_supported_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS


def preprocess_image_to_base64(image_path: Path, max_size=(1200, 1200), jpeg_quality=75) -> str:
    with Image.open(image_path) as img:
        img = ImageOps.exif_transpose(img)
        img = img.convert("RGB")
        img.thumbnail(max_size)

        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True)
        return base64.b64encode(buffer.getvalue()).decode("utf-8")


def page_count_to_length_category(page_count) -> str:
    try:
        p = int(float(page_count))
    except Exception:
        return ""
    if p < 200:
        return "Short"
    if p < 400:
        return "Medium"
    if p < 700:
        return "Long"
    return "Epic"


def metadata_enriched_yes(value) -> bool:
    return clean_text(value).lower() == "yes"


def enrichment_fill_pct(row) -> float:
    """Return the fraction of ENRICHMENT_SCORED_FIELDS that are non-empty (0.0–1.0)."""
    if not ENRICHMENT_SCORED_FIELDS:
        return 0.0
    filled = 0
    for field in ENRICHMENT_SCORED_FIELDS:
        val = clean_text(row.get(field, ""))
        # PageCount=0 counts as missing
        if field == "PageCount":
            try:
                if int(float(val)) > 0:
                    filled += 1
                continue
            except Exception:
                pass
        if val:
            filled += 1
    return filled / len(ENRICHMENT_SCORED_FIELDS)


def enrichment_last_date(row) -> "datetime.date | None":
    """Parse the Last Enriched cell; return a date or None if blank/invalid."""
    import datetime as _dt
    raw = clean_text(row.get("Last Enriched", ""))
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def today_str() -> str:
    """Return today's date as YYYY-MM-DD for the Last Enriched column."""
    import datetime as _dt
    return _dt.date.today().isoformat()


def is_enriched_enough(row) -> bool:
    """
    Skip re-enrichment if EITHER condition is true:

    1. Last Enriched date is set AND was within ENRICHMENT_MIN_AGE_DAYS days.
    2. Fill percentage across ENRICHMENT_SCORED_FIELDS >= ENRICHMENT_SKIP_THRESHOLD.

    The old hard-coded check (identifier + genre + tropes) is still honoured
    as a fallback so existing callers in amazon_owned_books.py keep working.
    """
    import datetime as _dt

    # Date-based skip
    if ENRICHMENT_MIN_AGE_DAYS > 0:
        last = enrichment_last_date(row)
        if last is not None:
            age_days = (_dt.date.today() - last).days
            if age_days < ENRICHMENT_MIN_AGE_DAYS:
                return True

    # Percentage-based skip
    if enrichment_fill_pct(row) >= ENRICHMENT_SKIP_THRESHOLD:
        return True

    # Legacy fallback (keeps amazon_owned_books.py fast on well-filled rows)
    has_identifier = bool(clean_text(row.get("ISBN_13")) or clean_text(row.get("ASIN")))
    has_genre = bool(clean_text(row.get("Genre")))
    has_tropes = bool(clean_text(row.get("Tropes")))
    return has_identifier and has_genre and has_tropes


def append_progress(rows: list[dict]) -> None:
    if not rows:
        return
    df = pd.DataFrame(rows)
    write_header = not PROGRESS_CSV.exists()
    df.to_csv(PROGRESS_CSV, mode="a", index=False, header=write_header)


def get_completed_duplicate_keys() -> set[str]:
    if not OUTPUT_XLSX.exists():
        return set()
    try:
        df = pd.read_excel(OUTPUT_XLSX, sheet_name="All Books")
    except Exception:
        return set()

    if "DuplicateKey" not in df.columns or "StoryGraph Completed" not in df.columns:
        return set()

    df["DuplicateKey"] = df["DuplicateKey"].apply(clean_text)
    df["StoryGraph Completed"] = df["StoryGraph Completed"].apply(clean_text)

    return set(
        df[
            (df["DuplicateKey"] != "") &
            (df["StoryGraph Completed"].str.lower() == "yes")
        ]["DuplicateKey"].tolist()
    )


def append_error(image_name: str, message: str) -> None:
    df = pd.DataFrame([{"Image": image_name, "Error": message}])
    write_header = not ERROR_LOG.exists()
    df.to_csv(ERROR_LOG, mode="a", index=False, header=write_header)


def append_sync_log(rows: list[dict]) -> None:
    if not rows:
        return
    df = pd.DataFrame(rows)
    write_header = not SYNC_LOG.exists()
    df.to_csv(SYNC_LOG, mode="a", index=False, header=write_header)

# =========================
# IMAGE EXTRACTION
# =========================
def extract_books_from_image(image_path: Path) -> list[dict]:
    b64 = preprocess_image_to_base64(image_path)

    schema = {
        "type": "object",
        "properties": {
            "books": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string"},
                        "author": {"type": "string"},
                        "confidence": {"type": "number"}
                    },
                    "required": ["title", "author", "confidence"],
                    "additionalProperties": False
                }
            }
        },
        "required": ["books"],
        "additionalProperties": False
    }

    prompt = """
Extract every visible book from this screenshot.

Rules:
- Return every clearly visible book.
- A single image may contain multiple books.
- Ignore UI text such as prices, ratings, buttons, badges, tabs, ads, and navigation text.
- If the author is not visible, use an empty string.
- confidence must be from 0.0 to 1.0.
- If no books are present, return an empty list.
- Do not guess wildly. Mark uncertain items with lower confidence.
"""

    response = client.responses.create(
        model=MODEL,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {"type": "input_image", "image_url": f"data:image/jpeg;base64,{b64}"}
                ],
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "book_extraction",
                "schema": schema,
                "strict": True
            }
        },
    )

    parsed = json.loads(response.output_text)
    return parsed.get("books", [])

# =========================
# FILE PROCESSING
# =========================
def process_one_file(image_path: Path) -> tuple[str, list[dict], str | None]:
    filename = image_path.name

    try:
        books = extract_books_from_image(image_path)

        if not books:
            return filename, [], "No books found"

        rows = []
        for book in books:
            title = clean_text(book.get("title"))
            author = clean_text(book.get("author"))
            confidence = float(book.get("confidence", 0.5))

            if not title:
                continue

            rows.append({
                "Image": filename,
                "Title": title,
                "Author": author,
                "Confidence": round(confidence, 3),
                "Needs Review": needs_review(title, author, confidence),
                "ISBN_10": "",
                "ISBN_13": "",
                "ASIN": "",
                "Lookup Source": "",
                "Description": "",
                "Genre": "",
                "PageCount": "",
                "LengthCategory": "",
                "AgeRange": "",
                "Tropes": "",
                "StoryGraph Status": "",
                "StoryGraph Matched Query": "",
                "StoryGraph Notes": "",
                "StoryGraph Completed": "",
                "Triggers": "",
                "Metadata Enriched": "",
                "DuplicateKey": build_duplicate_key(title, author),
            })

        if not rows:
            return filename, [], "No valid rows returned"

        return filename, rows, None

    except Exception as e:
        return filename, [], str(e)


def move_safe(src: Path, dest_folder: Path) -> None:
    dest = dest_folder / src.name
    if dest.exists():
        stem = dest.stem
        suffix = dest.suffix
        counter = 1
        while True:
            candidate = dest_folder / f"{stem}_{counter}{suffix}"
            if not candidate.exists():
                dest = candidate
                break
            counter += 1
    shutil.move(str(src), str(dest))


def get_incoming_files() -> list[Path]:
    files = []
    for p in INCOMING_FOLDER.iterdir():
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(p)
    return sorted(files)

# =========================
# SYNC
# =========================
def get_existing_book_keys(root: Path) -> set[str]:
    keys = set()

    for path in root.rglob("*"):
        if not is_supported_file(path):
            continue

        if COMPARE_BY_FILENAME_ONLY:
            key = path.name.lower()
        else:
            try:
                key = str(path.relative_to(root)).lower()
            except ValueError:
                key = path.name.lower()

        keys.add(key)

    return keys


def sync_from_books_2() -> tuple[list[dict], list[dict], list[dict]]:
    copied = []
    skipped_existing = []
    failed = []

    if not BOOKS_2_ROOT.exists():
        print(f"Books 2.0 source folder not found — skipping sync step: {BOOKS_2_ROOT}")
        return copied, skipped_existing, failed

    existing_keys = get_existing_book_keys(BASE_FOLDER)

    for src in BOOKS_2_ROOT.rglob("*"):
        if not is_supported_file(src):
            continue

        if COMPARE_BY_FILENAME_ONLY:
            src_key = src.name.lower()
        else:
            try:
                src_key = str(src.relative_to(BOOKS_2_ROOT)).lower()
            except ValueError:
                src_key = src.name.lower()

        if src_key in existing_keys:
            skipped_existing.append({
                "Source": str(src),
                "Destination": "",
                "Status": "Already existed"
            })
            continue

        dest = BASE_FOLDER / src.name

        if dest.exists():
            stem = dest.stem
            suffix = dest.suffix
            counter = 1
            while True:
                candidate = BASE_FOLDER / f"{stem}_{counter}{suffix}"
                if not candidate.exists():
                    dest = candidate
                    break
                counter += 1

        try:
            shutil.copy2(src, dest)
            copied.append({
                "Source": str(src),
                "Destination": str(dest),
                "Status": "Copied"
            })
            existing_keys.add(dest.name.lower())
            print(f"Copied from Books 2.0: {src.name}")
        except Exception as e:
            failed.append({
                "Source": str(src),
                "Destination": str(dest),
                "Status": f"Copy failed: {e}"
            })
            print(f"Failed to copy {src}: {e}")

    append_sync_log(copied + skipped_existing + failed)
    return copied, skipped_existing, failed

# =========================
# LOOKUPS
# =========================
def extract_isbns_from_identifiers(identifiers: list[dict]) -> tuple[str, str]:
    isbn_10 = ""
    isbn_13 = ""

    for item in identifiers or []:
        id_type = clean_text(item.get("type")).upper()
        id_value = clean_text(item.get("identifier"))

        if id_type == "ISBN_10" and not isbn_10:
            isbn_10 = id_value
        elif id_type == "ISBN_13" and not isbn_13:
            isbn_13 = id_value

    return isbn_10, isbn_13


def authors_match(candidate_authors, target_author: str) -> bool:
    target_norm = normalize_text(target_author)
    if not target_norm:
        return True

    if isinstance(candidate_authors, list):
        joined = " ".join(str(x) for x in candidate_authors)
    else:
        joined = str(candidate_authors or "")

    candidate_norm = normalize_text(joined)
    return target_norm in candidate_norm or candidate_norm in target_norm


def title_match(candidate_title: str, target_title: str) -> bool:
    c = normalize_text(candidate_title)
    t = normalize_text(target_title)
    if not c or not t:
        return False
    return c == t or c in t or t in c


def lookup_google_books(title: str, author: str) -> dict:
    try:
        q = f'intitle:"{title}" inauthor:"{author}"' if author else f'intitle:"{title}"'
        r = requests.get(
            "https://www.googleapis.com/books/v1/volumes",
            params={"q": q, "maxResults": 5, "printType": "books"},
            timeout=LOOKUP_TIMEOUT_SECONDS
        )
        r.raise_for_status()
        data = r.json()

        for item in data.get("items", []):
            info = item.get("volumeInfo", {})
            candidate_title = clean_text(info.get("title"))
            candidate_authors = info.get("authors", [])
            if title_match(candidate_title, title) and authors_match(candidate_authors, author):
                isbn_10, isbn_13 = extract_isbns_from_identifiers(info.get("industryIdentifiers", []))
                categories = info.get("categories", []) or []
                description = clean_text(info.get("description"))
                page_count = info.get("pageCount", "")
                maturity = clean_text(info.get("maturityRating"))
                return {
                    "ISBN_10": isbn_10,
                    "ISBN_13": isbn_13,
                    "ASIN": "",
                    "Lookup Source": "Google Books",
                    "Description": description,
                    "Genre": normalize_csv_list(", ".join([clean_text(x) for x in categories])),
                    "PageCount": page_count,
                    "LengthCategory": page_count_to_length_category(page_count),
                    "AgeRange": "Adult" if maturity == "MATURE" else "",
                }

        for item in data.get("items", []):
            info = item.get("volumeInfo", {})
            isbn_10, isbn_13 = extract_isbns_from_identifiers(info.get("industryIdentifiers", []))
            if isbn_10 or isbn_13:
                categories = info.get("categories", []) or []
                description = clean_text(info.get("description"))
                page_count = info.get("pageCount", "")
                maturity = clean_text(info.get("maturityRating"))
                return {
                    "ISBN_10": isbn_10,
                    "ISBN_13": isbn_13,
                    "ASIN": "",
                    "Lookup Source": "Google Books (fallback)",
                    "Description": description,
                    "Genre": normalize_csv_list(", ".join([clean_text(x) for x in categories])),
                    "PageCount": page_count,
                    "LengthCategory": page_count_to_length_category(page_count),
                    "AgeRange": "Adult" if maturity == "MATURE" else "",
                }
    except Exception:
        pass

    return {
        "ISBN_10": "", "ISBN_13": "", "ASIN": "", "Lookup Source": "",
        "Description": "", "Genre": "", "PageCount": "", "LengthCategory": "", "AgeRange": ""
    }


def lookup_open_library(title: str, author: str) -> dict:
    try:
        q = f"{title} {author}".strip()
        r = requests.get(
            "https://openlibrary.org/search.json",
            params={"q": q, "limit": 5},
            timeout=LOOKUP_TIMEOUT_SECONDS
        )
        r.raise_for_status()
        data = r.json()

        for doc in data.get("docs", []):
            candidate_title = clean_text(doc.get("title"))
            candidate_authors = doc.get("author_name", [])
            if title_match(candidate_title, title) and authors_match(candidate_authors, author):
                isbn_list = doc.get("isbn", []) or []
                isbn_10 = ""
                isbn_13 = ""
                for isbn in isbn_list:
                    clean = re.sub(r"[^0-9Xx]", "", str(isbn))
                    if len(clean) == 10 and not isbn_10:
                        isbn_10 = clean.upper()
                    elif len(clean) == 13 and not isbn_13:
                        isbn_13 = clean
                subjects = doc.get("subject", []) or []
                page_count = doc.get("number_of_pages_median", "")
                return {
                    "ISBN_10": isbn_10,
                    "ISBN_13": isbn_13,
                    "ASIN": "",
                    "Lookup Source": "Open Library",
                    "Description": "",
                    "Genre": normalize_csv_list(", ".join([clean_text(x) for x in subjects[:8]])),
                    "PageCount": page_count,
                    "LengthCategory": page_count_to_length_category(page_count),
                    "AgeRange": "",
                }
    except Exception:
        pass

    return {
        "ISBN_10": "", "ISBN_13": "", "ASIN": "", "Lookup Source": "",
        "Description": "", "Genre": "", "PageCount": "", "LengthCategory": "", "AgeRange": ""
    }


def lookup_amazon_asin(title: str, author: str) -> dict:
    if not ENABLE_ASIN_LOOKUP:
        return {"ASIN": "", "Lookup Source": ""}

    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9",
        }

        q = f"{title} {author} kindle".strip()
        r = requests.get(
            "https://www.amazon.com/s",
            params={"k": q, "i": "digital-text"},
            headers=headers,
            timeout=LOOKUP_TIMEOUT_SECONDS
        )
        r.raise_for_status()
        html = r.text

        patterns = [
            r'data-asin="([A-Z0-9]{10})"',
            r'"/dp/([A-Z0-9]{10})"',
            r'/dp/([A-Z0-9]{10})',
            r'"asin":"([A-Z0-9]{10})"',
        ]

        found = []
        for pattern in patterns:
            matches = re.findall(pattern, html)
            for m in matches:
                if m and len(m) == 10 and m not in found:
                    found.append(m)

        for asin in found:
            if asin.upper().startswith("B"):
                return {"ASIN": asin.upper(), "Lookup Source": "Amazon search"}

        if found:
            return {"ASIN": found[0].upper(), "Lookup Source": "Amazon search"}
    except Exception:
        pass

    return {"ASIN": "", "Lookup Source": ""}


def ai_content_enrichment(title: str, author: str, description: str, genre: str, age_range: str, page_count) -> dict:
    if not ENABLE_CONTENT_ENRICHMENT:
        return {"Genre": genre, "AgeRange": age_range, "Tropes": "", "Triggers": ""}

    schema = {
        "type": "object",
        "properties": {
            "genre": {"type": "string"},
            "age_range": {"type": "string"},
            "tropes": {"type": "string"},
            "triggers": {"type": "string"}
        },
        "required": ["genre", "age_range", "tropes", "triggers"],
        "additionalProperties": False
    }

    prompt = f"""
You are enriching a book spreadsheet.

Return only JSON with:
- genre: short comma-separated list of genres/subgenres
- age_range: one of "Children", "Middle Grade", "YA", "New Adult", "Adult", "General", or ""
- tropes: comma-separated list of likely reading tropes
- triggers: comma-separated list of likely content warnings/triggers, or "" if unknown

Rules:
- Be conservative. Do not invent highly specific triggers if unsupported.
- Use the provided metadata first.
- Keep each field concise.
- If uncertain, return a shorter list.

Title: {title}
Author: {author}
Existing genre: {genre}
Existing age range: {age_range}
Page count: {page_count}
Description:
{description}
"""

    try:
        response = client.responses.create(
            model=MODEL,
            input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "book_content_enrichment",
                    "schema": schema,
                    "strict": True
                }
            },
        )
        parsed = json.loads(response.output_text)
        return {
            "Genre": normalize_csv_list(parsed.get("genre", "")),
            "AgeRange": clean_text(parsed.get("age_range", "")),
            "Tropes": normalize_csv_list(parsed.get("tropes", "")),
            "Triggers": normalize_csv_list(parsed.get("triggers", "")),
        }
    except Exception:
        return {
            "Genre": normalize_csv_list(genre),
            "AgeRange": clean_text(age_range),
            "Tropes": "",
            "Triggers": "",
        }


def lookup_book_metadata(title: str, author: str) -> dict:
    result = {
        "ISBN_10": "", "ISBN_13": "", "ASIN": "", "Lookup Source": "",
        "Description": "", "Genre": "", "PageCount": "", "LengthCategory": "", "AgeRange": "",
        "Tropes": "", "Triggers": ""
    }

    google_result = lookup_google_books(title, author)
    if any(clean_text(google_result.get(k)) for k in ["ISBN_10", "ISBN_13", "Description", "Genre", "PageCount", "AgeRange"]):
        result.update(google_result)

    if not clean_text(result["ISBN_10"]) and not clean_text(result["ISBN_13"]) and not clean_text(result["Description"]) and not clean_text(result["Genre"]):
        time.sleep(LOOKUP_SLEEP_SECONDS)
        openlib_result = lookup_open_library(title, author)
        result.update({k: v for k, v in openlib_result.items() if clean_text(v)})

    if ENABLE_ASIN_LOOKUP:
        time.sleep(AMAZON_LOOKUP_SLEEP_SECONDS)
        asin_result = lookup_amazon_asin(title, author)
        if clean_text(asin_result.get("ASIN")):
            result["ASIN"] = asin_result.get("ASIN", "")
            if not clean_text(result["Lookup Source"]):
                result["Lookup Source"] = asin_result.get("Lookup Source", "")

    ai_result = ai_content_enrichment(
        title=title,
        author=author,
        description=clean_text(result.get("Description")),
        genre=clean_text(result.get("Genre")),
        age_range=clean_text(result.get("AgeRange")),
        page_count=result.get("PageCount", "")
    )

    if clean_text(ai_result.get("Genre")):
        result["Genre"] = ai_result["Genre"]
    if clean_text(ai_result.get("AgeRange")) and not clean_text(result.get("AgeRange")):
        result["AgeRange"] = ai_result["AgeRange"]
    if clean_text(ai_result.get("Tropes")):
        result["Tropes"] = ai_result["Tropes"]
    if clean_text(ai_result.get("Triggers")):
        result["Triggers"] = ai_result["Triggers"]

    if not clean_text(result.get("LengthCategory")) and clean_text(result.get("PageCount")):
        result["LengthCategory"] = page_count_to_length_category(result["PageCount"])

    return result

# =========================
# LOAD / NORMALIZE
# =========================
def ensure_all_books_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in ALL_BOOKS_COLUMNS:
        if col not in df.columns:
            if col == "Confidence":
                df[col] = 0.0
            else:
                df[col] = ""

    for col in TEXT_COLUMNS:
        df[col] = df[col].apply(clean_text)

    for col in ["Genre", "Tropes", "Triggers"]:
        df[col] = df[col].apply(normalize_csv_list)

    df["Confidence"] = pd.to_numeric(df["Confidence"], errors="coerce").fillna(0.0)
    df["PageCount"] = pd.to_numeric(df["PageCount"], errors="coerce").fillna("")
    df["LengthCategory"] = df.apply(
        lambda row: clean_text(row["LengthCategory"]) or page_count_to_length_category(row["PageCount"]),
        axis=1
    )

    return df[ALL_BOOKS_COLUMNS]


def load_existing_all_books() -> pd.DataFrame:
    if not OUTPUT_XLSX.exists():
        return pd.DataFrame(columns=ALL_BOOKS_COLUMNS)

    try:
        df = pd.read_excel(OUTPUT_XLSX, sheet_name="All Books")
        return ensure_all_books_columns(df)
    except Exception as e:
        print(f"Warning: could not read existing workbook All Books sheet: {e}")
        return pd.DataFrame(columns=ALL_BOOKS_COLUMNS)


def load_progress_csv() -> pd.DataFrame:
    if not PROGRESS_CSV.exists():
        return pd.DataFrame(columns=ALL_BOOKS_COLUMNS)

    try:
        df = pd.read_csv(PROGRESS_CSV)
        return ensure_all_books_columns(df)
    except Exception as e:
        print(f"Warning: could not read progress CSV normally: {e}")
        print("Trying recovery mode...")
        try:
            df = pd.read_csv(PROGRESS_CSV, engine="python", on_bad_lines="skip")
            df = ensure_all_books_columns(df)
            print(f"Recovered {len(df)} rows from progress CSV.")
            return df
        except Exception as e2:
            print(f"Recovery failed: {e2}")
            return pd.DataFrame(columns=ALL_BOOKS_COLUMNS)


def metadata_fill_score(df: pd.DataFrame) -> pd.Series:
    return (
        (df["ISBN_10"].astype(str).str.strip() != "").astype(int) +
        (df["ISBN_13"].astype(str).str.strip() != "").astype(int) +
        (df["ASIN"].astype(str).str.strip() != "").astype(int) +
        (df["Lookup Source"].astype(str).str.strip() != "").astype(int) +
        (df["Description"].astype(str).str.strip() != "").astype(int) +
        (df["Genre"].astype(str).str.strip() != "").astype(int) +
        (df["PageCount"].astype(str).str.strip() != "").astype(int) +
        (df["LengthCategory"].astype(str).str.strip() != "").astype(int) +
        (df["AgeRange"].astype(str).str.strip() != "").astype(int) +
        (df["Tropes"].astype(str).str.strip() != "").astype(int) +
        (df["Triggers"].astype(str).str.strip() != "").astype(int) +
        (df["StoryGraph Status"].astype(str).str.strip() != "").astype(int) +
        (df["StoryGraph Completed"].astype(str).str.strip() != "").astype(int) +
        (df["Metadata Enriched"].astype(str).str.strip() != "").astype(int)
    )

# =========================
# ENRICHMENT
# =========================
def enrich_unique_books(df_unique: pd.DataFrame) -> pd.DataFrame:
    if df_unique.empty:
        return df_unique

    df_unique = df_unique.copy()

    for col in [
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source", "Description", "Genre",
        "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers", "Metadata Enriched"
    ]:
        if col not in df_unique.columns:
            df_unique[col] = ""

    total = len(df_unique)

    for idx, row in df_unique.iterrows():
        title = clean_text(row.get("Title"))
        author = clean_text(row.get("Author"))

        # Fast skip if prior run marked it enriched
        if metadata_enriched_yes(row.get("Metadata Enriched")):
            continue

        # Smart skip if row is already good enough
        if is_enriched_enough(row):
            df_unique.at[idx, "Metadata Enriched"] = "Yes"
            continue

        print(f"Enriching {idx + 1}/{total}: {title} | {author}")
        result = lookup_book_metadata(title, author)

        for field in [
            "ISBN_10", "ISBN_13", "ASIN", "Lookup Source", "Description", "Genre",
            "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers"
        ]:
            current_value = clean_text(df_unique.at[idx, field])
            new_value = clean_text(result.get(field))
            if not current_value and new_value:
                df_unique.at[idx, field] = new_value

        if not clean_text(df_unique.at[idx, "LengthCategory"]):
            df_unique.at[idx, "LengthCategory"] = page_count_to_length_category(df_unique.at[idx, "PageCount"])

        if is_enriched_enough(df_unique.loc[idx]):
            df_unique.at[idx, "Metadata Enriched"] = "Yes"

        # Stamp the date whenever this row was touched by enrichment.
        # reenrich_existing.py uses this date to skip recently enriched rows
        # — more reliable than the Yes/No Metadata Enriched flag which goes stale.
        df_unique.at[idx, "Last Enriched"] = today_str()

        time.sleep(AI_ENRICH_SLEEP_SECONDS)

    df_unique["Genre"] = df_unique["Genre"].apply(normalize_csv_list)
    df_unique["Tropes"] = df_unique["Tropes"].apply(normalize_csv_list)
    df_unique["Triggers"] = df_unique["Triggers"].apply(normalize_csv_list)
    return df_unique

# =========================
# OUTPUT
# =========================
def build_excel_from_progress() -> None:
    df_progress = load_progress_csv()
    existing_all = load_existing_all_books()

    if df_progress.empty and existing_all.empty:
        print("No progress CSV or existing workbook data available.")
        return

    combined_raw = pd.concat([existing_all, df_progress], ignore_index=True)
    combined_raw = ensure_all_books_columns(combined_raw)

    combined_raw["_fill_score"] = metadata_fill_score(combined_raw)
    combined_raw = combined_raw.sort_values(
        ["_fill_score", "Confidence"],
        ascending=[False, False]
    ).drop_duplicates(
        subset=["Image", "DuplicateKey"],
        keep="first"
    ).reset_index(drop=True)

    duplicate_summary = (
        combined_raw.groupby(["DuplicateKey", "Title", "Author"], dropna=False)
        .size()
        .reset_index(name="Count")
        .sort_values(["Count", "Title"], ascending=[False, True])
    )

    combined_raw["_fill_score"] = metadata_fill_score(combined_raw)
    best_per_book = combined_raw.sort_values(
        ["_fill_score", "Confidence", "Title", "Author"],
        ascending=[False, False, True, True]
    ).drop_duplicates(
        subset=["DuplicateKey"],
        keep="first"
    ).reset_index(drop=True)

    df_unique = best_per_book.copy()

    if ENABLE_METADATA_ENRICHMENT:
        df_unique = enrich_unique_books(df_unique)

    unique_metadata = df_unique[[
        "DuplicateKey", "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
        "Description", "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
        "Metadata Enriched",
        "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed"
    ]].copy()

    best_per_book = best_per_book.drop(
        columns=[
            "ISBN_10", "ISBN_13", "ASIN", "Lookup Source", "Description", "Genre",
            "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
            "Metadata Enriched",
            "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed"
        ],
        errors="ignore"
    )
    best_per_book = best_per_book.merge(unique_metadata, how="left", on="DuplicateKey")

    duplicate_summary = duplicate_summary.merge(unique_metadata, how="left", on="DuplicateKey")

    review_only = (
        best_per_book[best_per_book["Needs Review"] == "Yes"]
        .sort_values(["Confidence", "Image"], ascending=[True, True])
        .reset_index(drop=True)
    )

    output_unique = df_unique.loc[:, [
        "Title", "Author", "Confidence", "Needs Review",
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
        "Description", "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
        "Metadata Enriched",
        "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed"
    ]]

    output_all = best_per_book.loc[:, [
        "Image", "Title", "Author", "Confidence", "Needs Review",
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
        "Description", "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
        "Metadata Enriched",
        "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed",
        "DuplicateKey"
    ]]

    output_review = review_only.loc[:, [
        "Image", "Title", "Author", "Confidence", "Needs Review",
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
        "Description", "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
        "Metadata Enriched",
        "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed",
        "DuplicateKey"
    ]]

    output_duplicates = duplicate_summary.loc[:, [
        "Title", "Author", "Count",
        "ISBN_10", "ISBN_13", "ASIN", "Lookup Source",
        "Genre", "PageCount", "LengthCategory", "AgeRange", "Tropes", "Triggers",
        "Metadata Enriched",
        "StoryGraph Status", "StoryGraph Matched Query", "StoryGraph Notes", "StoryGraph Completed",
        "DuplicateKey"
    ]]

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        output_all.to_excel(writer, sheet_name="All Books", index=False)
        output_unique.to_excel(writer, sheet_name="Unique Books", index=False)
        output_duplicates.to_excel(writer, sheet_name="Duplicate Summary", index=False)
        output_review.to_excel(writer, sheet_name="Needs Review", index=False)

# =========================
# MAIN
# =========================
def main():
    print("Starting sync from Books 2.0...")
    copied, skipped_existing, failed = sync_from_books_2()

    print("\nSync summary")
    print(f"Copied new files: {len(copied)}")
    print(f"Already existed: {len(skipped_existing)}")
    print(f"Copy failures: {len(failed)}")

    incoming_files = get_incoming_files()

    if TEST_LIMIT:
        if RANDOM_SAMPLE:
            incoming_files = random.sample(incoming_files, min(TEST_LIMIT, len(incoming_files)))
            print(f"\nTEST MODE: Random sample of {len(incoming_files)} images")
        else:
            incoming_files = incoming_files[:TEST_LIMIT]
            print(f"\nTEST MODE: First {len(incoming_files)} images")

    if not incoming_files:
        print("No incoming image files found in Books folder.")
        build_excel_from_progress()
        return

    print(f"Found {len(incoming_files)} file(s) to process.")

    completed_duplicate_keys = get_completed_duplicate_keys()
    if completed_duplicate_keys:
        print(f"Found {len(completed_duplicate_keys)} completed StoryGraph duplicate key(s) to skip.")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {executor.submit(process_one_file, path): path for path in incoming_files}

        for future in as_completed(future_map):
            image_path = future_map[future]
            filename = image_path.name

            try:
                filename, rows, error_message = future.result()

                if rows:
                    rows_to_append = []
                    for row in rows:
                        dup_key = clean_text(row.get("DuplicateKey"))
                        if dup_key and dup_key in completed_duplicate_keys:
                            continue
                        rows_to_append.append(row)

                    if rows_to_append:
                        append_progress(rows_to_append)
                        print(f"Processed: {filename} ({len(rows_to_append)} new book(s))")
                    else:
                        print(f"Processed: {filename} (0 new books; all matched completed duplicates)")

                    move_safe(image_path, PROCESSED_FOLDER)
                else:
                    append_error(filename, error_message or "Unknown error")
                    move_safe(image_path, SKIPPED_FOLDER)
                    print(f"Skipped: {filename} - {error_message}")

            except Exception as e:
                append_error(filename, str(e))
                if image_path.exists():
                    move_safe(image_path, SKIPPED_FOLDER)
                print(f"Skipped: {filename} - {e}")

    build_excel_from_progress()
    print(f"Done. Excel created: {OUTPUT_XLSX}")


if __name__ == "__main__":
    start = time.time()
    main()
    elapsed = round(time.time() - start, 1)
    print(f"Elapsed: {elapsed} seconds")
