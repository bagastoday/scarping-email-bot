import os
import re
import time
import asyncio
import logging
import io
from datetime import datetime
from urllib.parse import urljoin, urlparse

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# dnspython untuk cek MX record email domain (opsional, bot tetap jalan tanpa ini)
try:
    import dns.resolver
    DNS_AVAILABLE = True
except ImportError:
    DNS_AVAILABLE = False

# Playwright dipakai untuk JS rendering (opsional, bot tetap jalan tanpa ini)
try:
    from playwright.async_api import async_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

# PyMuPDF untuk ekstrak email dari PDF publik
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Document,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# =====================================================
# LOAD ENV
# =====================================================

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
SERPAPI_KEY = os.getenv("SERPAPI_KEY")

# Verifalia email verification API
# Isi di file .env:
# VERIFALIA_USERNAME=isi_username_api
# VERIFALIA_PASSWORD=isi_password_api
VERIFALIA_USERNAME = os.getenv("VERIFALIA_USERNAME")
VERIFALIA_PASSWORD = os.getenv("VERIFALIA_PASSWORD")

# =====================================================
# SETTINGS
# =====================================================

WEB_MAX_PAGES = 120
WEB_MAX_DEPTH = 3
MAPS_MAX_BUSINESSES = 50
MAPS_WEBSITE_MAX_PAGES = 8
DORK_MAX_RESULTS = 10           # hasil URL per query dork
DORK_MAX_PAGES_PER_URL = 5      # halaman yang discan dari tiap URL hasil dork
DIRECTORY_MAX_PAGES = 30        # halaman listing direktori yang discan
DIRECTORY_DETAIL_MAX = 5        # halaman detail per listing direktori
REQUEST_TIMEOUT = 15
REQUEST_DELAY = 0.8
FETCH_RETRIES = 2
DEFAULT_TARGET_EMAILS = 150
PARTIAL_SAVE_EVERY = 50
PROGRESS_UPDATE_EVERY = 3

# =====================================================
# EMAIL CHECKER SETTINGS
# =====================================================

# True = pakai MX check + Verifalia sebelum email masuk XLSX.
# Jika Verifalia limit / kredensial kosong / error, bot otomatis fallback ke local checker + MX.
EMAIL_CHECKER_ENABLED = True
VERIFALIA_QUALITY = "Standard"       # Hemat credit. Bisa diganti: "High" / "Extreme" jika perlu.
VERIFALIA_ACCEPT_RISKY = True         # True = Risky tetap masuk; False = hanya Deliverable.
VERIFALIA_LIMIT_REACHED = False
VERIFALIA_CACHE = {}
MX_CACHE = {}

PRIORITY_LINK_KEYWORDS = (
    "contact", "kontak", "hubungi", "about", "tentang", "team", "staff",
    "profile", "profil", "company", "perusahaan", "support", "customer-service",
    "cs", "marketing", "sales", "career", "karir", "privacy", "legal"
)

# Direktori bisnis publik yang akan di-scrape
DIRECTORIES = {
    "yellowpages_id": {
        "name": "Yellowpages Indonesia",
        "search_url": "https://www.yellowpages.co.id/search?q={keyword}&l={location}&page={page}",
        "listing_selector": "div.listing",
        "name_selector": "h3.listing-name",
        "url_selector": "a.listing-url",
    },
    "clutch": {
        "name": "Clutch.co",
        "search_url": "https://clutch.co/agencies?q={keyword}&page={page}",
        "listing_selector": "li.provider-row",
        "name_selector": "h3.company-name",
        "url_selector": "a.website-link",
    },
}

OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

WEB_MASTER_FILE = "master_web_scraping.xlsx"
MAPS_MASTER_FILE = "master_maps_scraping.xlsx"
DORK_MASTER_FILE = "master_dork_scraping.xlsx"
DIRECTORY_MASTER_FILE = "master_directory_scraping.xlsx"

# =====================================================
# LOGGING
# =====================================================

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

# =====================================================
# EMAIL RULES
# =====================================================

EMAIL_REGEX = re.compile(
    r"\b[A-Za-z0-9][A-Za-z0-9._%+-]{0,38}[A-Za-z0-9]@[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}\b"
)

BLOCKED_EMAILS = {
    "example@example.com", "test@test.com", "admin@example.com",
    "email@example.com", "your@email.com", "you@example.com",
}

BLOCKED_LOCAL_PARTS = {
    "privacy", "policy", "legal", "terms", "gdpr", "dpo",
    "noreply", "no-reply", "no_reply", "donotreply", "do-not-reply",
    "mailer-daemon", "postmaster", "abuse", "root", "hostmaster",
    "webmaster", "example", "test", "testing", "email", "your", "you",
}

BLOCKED_LOCAL_KEYWORDS = (
    "noreply", "no-reply", "donotreply", "do-not-reply",
    "sentry", "notification", "automated", "bounce", "mailer-daemon",
)

BLOCKED_EMAIL_DOMAINS = {
    "linktr.ee", "sentry.wixpress.com", "wixpress.com", "wix.com",
    "wixsite.com", "sentry.io", "example.com", "test.com",
}

BLOCKED_EMAIL_DOMAIN_SUFFIXES = (
    ".wixpress.com", ".wixsite.com", ".sentry.io",
)

HEX_HASH_REGEX = re.compile(r"^[a-f0-9]{16,}$", re.IGNORECASE)
LONG_RANDOM_REGEX = re.compile(r"^[a-z0-9]{24,}$", re.IGNORECASE)

BLOCKED_EXTENSIONS = (
    ".png", ".jpg", ".jpeg", ".webp", ".svg", ".gif",
    ".bmp", ".ico", ".pdf", ".css", ".js",
)

# =====================================================
# GOOGLE DORK QUERIES TEMPLATE
# =====================================================

DORK_TEMPLATES = [
    '"{keyword}" "{location}" "@gmail.com" OR "@yahoo.com" OR "@hotmail.com"',
    '"{keyword}" "{location}" "email" "kontak"',
    'site:yellowpages.co.id "{keyword}" "{location}"',
    'site:clutch.co "{keyword}"',
    'filetype:pdf "{keyword}" "{location}" "email"',
    'inurl:contact "{keyword}" "{location}"',
    'inurl:about "{keyword}" "{location}" email',
    '"{keyword}" "{location}" "@" site:.co.id OR site:.com',
    '"{keyword}" "Jakarta" OR "Surabaya" OR "Bandung" email contact',
    'intitle:"{keyword}" "{location}" email',
]


# =====================================================
# HELPER: EMAIL UTILS
# =====================================================

def clean_email(email: str) -> str:
    return email.strip().lower().replace("mailto:", "")


def get_email_domain(email: str) -> str:
    email = clean_email(email)
    if "@" not in email:
        return ""
    return email.split("@", 1)[1].strip().lower()


def validate_email(email: str) -> str:
    email = clean_email(email)

    # Harus ada @ dan minimal satu karakter sebelum dan sesudahnya
    if "@" not in email:
        return "Invalid"

    # Blok format @domain.com tanpa local-part (bug yang dilaporkan)
    if email.startswith("@"):
        return "Invalid"

    # Regex fullmatch — local-part minimal 2 karakter, tidak boleh mulai/akhir dengan titik
    if not EMAIL_REGEX.fullmatch(email):
        return "Invalid"

    if email in BLOCKED_EMAILS:
        return "Invalid"

    if email.endswith(BLOCKED_EXTENSIONS):
        return "Invalid"

    local_part, domain = email.rsplit("@", 1)
    local_part = local_part.strip().lower()
    domain = domain.strip().lower()

    # Validasi domain
    if "." not in domain:
        return "Invalid"
    if domain.startswith(".") or domain.endswith("."):
        return "Invalid"
    if ".." in domain:
        return "Invalid"

    # TLD minimal 2 karakter, maksimal 10 (hindari string acak panjang)
    tld = domain.rsplit(".", 1)[-1]
    if len(tld) < 2 or len(tld) > 10:
        return "Invalid"

    # Blok domain platform/teknis
    if domain in BLOCKED_EMAIL_DOMAINS:
        return "Invalid"
    if any(domain.endswith(suffix) for suffix in BLOCKED_EMAIL_DOMAIN_SUFFIXES):
        return "Invalid"

    # Validasi local-part
    if not local_part or len(local_part) < 2 or len(local_part) > 40:
        return "Invalid"

    # Local-part tidak boleh mulai atau akhiri dengan titik
    if local_part.startswith(".") or local_part.endswith("."):
        return "Invalid"

    # Tidak boleh ada titik berurutan di local-part
    if ".." in local_part:
        return "Invalid"

    normalized_local = local_part.replace(".", "").replace("_", "").replace("-", "")

    if local_part in BLOCKED_LOCAL_PARTS or normalized_local in BLOCKED_LOCAL_PARTS:
        return "Invalid"

    if any(keyword in local_part for keyword in BLOCKED_LOCAL_KEYWORDS):
        return "Invalid"

    # Blok hash/string acak
    if HEX_HASH_REGEX.fullmatch(normalized_local):
        return "Invalid"

    if LONG_RANDOM_REGEX.fullmatch(normalized_local) and not any(ch in local_part for ch in [".", "-", "_"]):
        return "Invalid"

    # Blok local-part yang mayoritas angka (tracking ID)
    digit_count = sum(ch.isdigit() for ch in normalized_local)
    if len(normalized_local) >= 12 and digit_count / max(len(normalized_local), 1) > 0.45:
        return "Invalid"

    # Blok email yang local-part-nya hanya angka semua (contoh: 123456@domain.com)
    if normalized_local.isdigit():
        return "Invalid"

    return "Valid"


def extract_emails_from_text(text: str):
    found = EMAIL_REGEX.findall(text or "")
    results = []
    seen = set()
    for email in found:
        email = clean_email(email)

        # Skip duplikat dalam satu halaman
        if email in seen:
            continue
        seen.add(email)

        # Skip langsung kalau mulai dengan @ (format rusak)
        if email.startswith("@"):
            continue

        # Skip kalau tidak ada karakter huruf sama sekali di local-part
        local = email.split("@")[0]
        if not any(c.isalpha() for c in local):
            continue

        status = validate_email(email)
        results.append((email, status))
    return results



# =====================================================
# HELPER: EMAIL CHECKER API (MX + VERIFALIA)
# =====================================================

def has_mx_record(domain: str) -> bool:
    """Cek apakah domain email punya MX record. Jika dnspython belum ada, fallback True agar bot tetap jalan."""
    domain = (domain or "").strip().lower()
    if not domain:
        return False

    if domain in MX_CACHE:
        return MX_CACHE[domain]

    if not DNS_AVAILABLE:
        # Fallback: jangan mematikan bot hanya karena dnspython belum diinstall.
        MX_CACHE[domain] = True
        return True

    try:
        dns.resolver.resolve(domain, "MX")
        MX_CACHE[domain] = True
        return True
    except Exception:
        MX_CACHE[domain] = False
        return False


def get_verifalia_status_text() -> str:
    if not EMAIL_CHECKER_ENABLED:
        return "OFF"
    if VERIFALIA_LIMIT_REACHED:
        return "LIMIT/FALLBACK"
    if not VERIFALIA_USERNAME or not VERIFALIA_PASSWORD:
        return "NO API/FALLBACK"
    return "ACTIVE"


async def check_verifalia(session: aiohttp.ClientSession, email: str) -> str:
    """
    Return: Deliverable, Risky, Undeliverable, Unknown, atau Skipped.
    Skipped artinya bot tetap lanjut pakai local checker + MX.
    """
    global VERIFALIA_LIMIT_REACHED

    email = clean_email(email)

    if not EMAIL_CHECKER_ENABLED:
        return "Skipped"

    if VERIFALIA_LIMIT_REACHED:
        return "Skipped"

    if not VERIFALIA_USERNAME or not VERIFALIA_PASSWORD:
        return "Skipped"

    if email in VERIFALIA_CACHE:
        return VERIFALIA_CACHE[email]

    url = "https://api.verifalia.com/v2.7/email-validations"
    payload = {
        "entries": [
            {"inputData": email}
        ],
        "quality": VERIFALIA_QUALITY,
    }

    try:
        async with session.post(
            url,
            json=payload,
            auth=aiohttp.BasicAuth(VERIFALIA_USERNAME, VERIFALIA_PASSWORD),
            timeout=25,
        ) as response:
            # Limit / quota habis. Jangan matikan bot; fallback ke local + MX.
            if response.status == 429:
                VERIFALIA_LIMIT_REACHED = True
                logging.warning("Verifalia limit tercapai. Fallback ke local checker + MX.")
                return "Skipped"

            # Kredensial salah / akses ditolak. Jangan matikan bot.
            if response.status in (402, 429):
                logging.warning("Verifalia username/password salah atau akses ditolak. Fallback ke local checker + MX.")
                return "Skipped"

            if response.status >= 400:
                logging.warning(f"Verifalia HTTP error: {response.status}")
                return "Skipped"

            data = await response.json()
            entries = data.get("entries", {}).get("data", [])

            if not entries:
                return "Unknown"

            result = entries[0].get("classification", "Unknown")
            VERIFALIA_CACHE[email] = result
            return result

    except asyncio.TimeoutError:
        logging.warning(f"Verifalia timeout untuk {email}")
        return "Unknown"
    except Exception as e:
        logging.warning(f"Verifalia error untuk {email}: {e}")
        return "Unknown"


async def is_email_allowed_by_checker(session: aiohttp.ClientSession, email: str) -> tuple[bool, str]:
    """
    Checker final sebelum email masuk Excel.
    - Selalu cek MX jika dnspython tersedia.
    - Jika Verifalia aktif, hanya simpan Deliverable/Risky sesuai setting.
    - Jika Verifalia limit/kredensial kosong, fallback: local checker + MX tetap boleh masuk.
    """
    email = clean_email(email)
    email_domain = get_email_domain(email)

    if not has_mx_record(email_domain):
        return False, "No MX"

    verifalia_status = await check_verifalia(session, email)

    # Fallback mode: Verifalia tidak tersedia/limit, tetapi MX valid.
    if verifalia_status == "Skipped":
        return True, "Local+MX"

    if verifalia_status == "Deliverable":
        return True, "Deliverable"

    if VERIFALIA_ACCEPT_RISKY and verifalia_status == "Risky":
        return True, "Risky"

    return False, verifalia_status


# =====================================================
# HELPER: URL UTILS
# =====================================================

def normalize_url(url: str) -> str:
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


def is_valid_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        return parsed.scheme in ["http", "https"] and bool(parsed.netloc)
    except Exception:
        return False


def extract_links(base_url: str, html: str):
    links = set()
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("a", href=True):
        href = tag.get("href", "").strip()
        if not href or href.startswith(("mailto:", "tel:", "#", "javascript:")):
            continue
        full_url = urljoin(base_url, href)
        parsed = urlparse(full_url)
        if parsed.scheme not in ["http", "https"]:
            continue
        clean_url = parsed._replace(fragment="").geturl()
        links.add(clean_url)
    return links


def same_domain(url1: str, url2: str) -> bool:
    return (
        urlparse(url1).netloc.lower().replace("www.", "")
        == urlparse(url2).netloc.lower().replace("www.", "")
    )


def normalize_crawl_url(url: str) -> str:
    parsed = urlparse(url)
    return parsed._replace(fragment="").geturl().rstrip("/")


def link_priority(url: str) -> int:
    lowered = url.lower()
    if any(keyword in lowered for keyword in PRIORITY_LINK_KEYWORDS):
        return 0
    if lowered.count("/") <= 3:
        return 1
    return 2


# =====================================================
# HELPER: PROGRESS & FILE
# =====================================================

def make_progress_bar(percent: int, size: int = 10) -> str:
    filled = max(0, min(size, round((percent / 100) * size)))
    return "█" * filled + "░" * (size - filled)


def get_columns(scraping_type: str):
    if scraping_type == "web":
        return ["No", "Source URL", "Email", "Status", "Found At", "Scraping Type"]
    elif scraping_type == "maps":
        return ["No", "Business Name", "Search Keyword", "Location", "Website URL", "Email", "Status", "Found At", "Scraping Type"]
    elif scraping_type == "dork":
        return ["No", "Dork Query", "Source URL", "Email", "Status", "Found At", "Scraping Type"]
    elif scraping_type == "directory":
        return ["No", "Directory", "Business Name", "Website URL", "Email", "Status", "Found At", "Scraping Type"]
    return ["No", "Source URL", "Email", "Status", "Found At", "Scraping Type"]


def get_master_filename(scraping_type: str):
    return {
        "web": WEB_MASTER_FILE,
        "maps": MAPS_MASTER_FILE,
        "dork": DORK_MASTER_FILE,
        "directory": DIRECTORY_MASTER_FILE,
    }.get(scraping_type, WEB_MASTER_FILE)


def safe_xlsx_name(name: str, scraping_type: str) -> str:
    name = (name or "").strip()
    name = os.path.basename(name)
    name = re.sub(r"[^A-Za-z0-9_. -]", "_", name)
    name = name.strip(" ._")
    if not name:
        name = f"{scraping_type}_scraping_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def list_xlsx_files(scraping_type: str | None = None):
    files = []
    if not os.path.exists(OUTPUT_DIR):
        return files
    for filename in sorted(os.listdir(OUTPUT_DIR)):
        if not filename.lower().endswith(".xlsx"):
            continue
        lower = filename.lower()
        if scraping_type == "web" and any(x in lower for x in ["maps", "dork", "directory"]):
            continue
        if scraping_type == "maps" and "maps" not in lower:
            continue
        if scraping_type == "dork" and "dork" not in lower:
            continue
        if scraping_type == "directory" and "directory" not in lower:
            continue
        path = os.path.join(OUTPUT_DIR, filename)
        try:
            size_kb = os.path.getsize(path) / 1024
            modified = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
            try:
                rows_count = len(pd.read_excel(path))
            except Exception:
                rows_count = 0
            files.append({
                "filename": filename,
                "path": path,
                "size_kb": size_kb,
                "modified": modified,
                "rows": rows_count,
            })
        except Exception:
            continue
    return files


def load_existing_master(scraping_type: str, filename: str | None = None):
    filename = safe_xlsx_name(filename or get_master_filename(scraping_type), scraping_type)
    master_path = os.path.join(OUTPUT_DIR, filename)
    columns = get_columns(scraping_type)

    if not os.path.exists(master_path):
        return pd.DataFrame(columns=columns), set(), set(), master_path

    try:
        df = pd.read_excel(master_path)
        for column in columns:
            if column not in df.columns:
                df[column] = ""
        df = df[columns]
        email_col = "Email"
        if email_col in df.columns:
            df[email_col] = df[email_col].fillna("").astype(str).str.strip().str.lower()
            df = df[df[email_col].apply(lambda x: validate_email(x) == "Valid")]
        existing_emails = set(df[email_col].dropna().astype(str).str.strip().str.lower())
        existing_emails.discard("")
        existing_domains = {get_email_domain(e) for e in existing_emails if get_email_domain(e)}
        return df, existing_emails, existing_domains, master_path
    except Exception as e:
        logging.warning(f"Gagal membaca file master {master_path}: {e}")
        return pd.DataFrame(columns=columns), set(), set(), master_path


def append_to_master_excel(new_rows: list, scraping_type: str, filename: str | None = None):
    columns = get_columns(scraping_type)
    old_df, existing_emails, existing_domains, master_path = load_existing_master(scraping_type, filename)

    if not new_rows:
        old_df.to_excel(master_path, index=False)
        return master_path, 0, len(old_df)

    # Tentukan index kolom email dan status berdasarkan tipe scraping
    email_idx = {"web": 2, "maps": 5, "dork": 3, "directory": 4}.get(scraping_type, 2)
    status_idx = {"web": 3, "maps": 6, "dork": 4, "directory": 5}.get(scraping_type, 3)

    unique_rows = []
    for row in new_rows:
        row_email = str(row[email_idx]).strip().lower() if len(row) > email_idx else ""
        row_status = str(row[status_idx]).strip() if len(row) > status_idx else ""

        if not row_email or row_status != "Valid" or validate_email(row_email) != "Valid":
            continue

        row_domain = get_email_domain(row_email)
        if not row_domain:
            continue
        if row_email in existing_emails:
            continue
        if row_domain in existing_domains:
            continue

        existing_emails.add(row_email)
        existing_domains.add(row_domain)
        unique_rows.append(row)

    new_df = pd.DataFrame(unique_rows, columns=columns)
    combined_df = pd.concat([old_df, new_df], ignore_index=True)
    combined_df["No"] = range(1, len(combined_df) + 1)
    combined_df.to_excel(master_path, index=False)

    return master_path, len(unique_rows), len(combined_df)


def save_excel(rows: list, filename: str, scraping_type: str):
    filepath = os.path.join(OUTPUT_DIR, filename)
    columns = get_columns(scraping_type)
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(filepath, index=False)
    return filepath


# =====================================================
# HELPER: HTTP FETCH
# =====================================================

async def fetch_html(session: aiohttp.ClientSession, url: str):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; EmailScraperBot/1.0; +legal-public-email-scan)"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    last_error = None
    for attempt in range(1, FETCH_RETRIES + 2):
        try:
            async with session.get(
                url, headers=headers, timeout=REQUEST_TIMEOUT, allow_redirects=True,
            ) as response:
                if response.status >= 400:
                    return None, f"HTTP {response.status}"
                content_type = response.headers.get("content-type", "").lower()
                if "text/html" not in content_type and "application/xhtml" not in content_type:
                    return None, "Not HTML"
                html = await response.text(errors="ignore")
                return html, None
        except asyncio.TimeoutError:
            last_error = "Timeout"
        except aiohttp.ClientError as e:
            last_error = str(e)
        except Exception as e:
            last_error = str(e)
        if attempt <= FETCH_RETRIES:
            await asyncio.sleep(0.5 * attempt)
    return None, last_error or "Unknown error"


async def fetch_pdf_bytes(session: aiohttp.ClientSession, url: str):
    """Download PDF bytes untuk diekstrak emailnya."""
    try:
        async with session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True) as response:
            if response.status >= 400:
                return None
            content_type = response.headers.get("content-type", "").lower()
            if "pdf" not in content_type and not url.lower().endswith(".pdf"):
                return None
            return await response.read()
    except Exception:
        return None


async def fetch_html_playwright(url: str):
    """Fetch halaman yang butuh JS rendering via Playwright."""
    if not PLAYWRIGHT_AVAILABLE:
        return None, "Playwright tidak tersedia"
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()
            await page.goto(url, timeout=20000, wait_until="networkidle")
            html = await page.content()
            await browser.close()
            return html, None
    except Exception as e:
        return None, str(e)


def extract_emails_from_pdf_bytes(pdf_bytes: bytes):
    """Ekstrak email dari bytes PDF menggunakan PyMuPDF."""
    if not PYMUPDF_AVAILABLE or not pdf_bytes:
        return []
    results = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            text = page.get_text()
            for email, status in extract_emails_from_text(text):
                results.append((email, status))
        doc.close()
    except Exception as e:
        logging.warning(f"Gagal ekstrak PDF: {e}")
    return results


# =====================================================
# HELPER: MENU & MESSAGE
# =====================================================

def main_menu():
    keyboard = [
        [
            InlineKeyboardButton("🌐 Scraping Web", callback_data="scrape_web"),
            InlineKeyboardButton("🗺️ Google Maps", callback_data="scrape_maps"),
        ],
        [
            InlineKeyboardButton("🔍 Google Dorking", callback_data="scrape_dork"),
            InlineKeyboardButton("📋 Direktori Bisnis", callback_data="scrape_directory"),
        ],
        [
            InlineKeyboardButton("📦 Batch URL (.txt)", callback_data="scrape_batch"),
            InlineKeyboardButton("📁 Kelola XLSX", callback_data="manage_xlsx"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


def output_choice_menu(scraping_type: str):
    keyboard = [
        [
            InlineKeyboardButton("🆕 Buat XLSX Baru", callback_data=f"xlsx_new:{scraping_type}"),
            InlineKeyboardButton("📂 Pakai XLSX Lama", callback_data=f"xlsx_old:{scraping_type}"),
        ],
        [InlineKeyboardButton("⬅️ Kembali", callback_data="back_main")],
    ]
    return InlineKeyboardMarkup(keyboard)


def manage_xlsx_menu():
    keyboard = [
        [
            InlineKeyboardButton("📄 Web", callback_data="list_xlsx:web"),
            InlineKeyboardButton("🗺️ Maps", callback_data="list_xlsx:maps"),
            InlineKeyboardButton("🔍 Dork", callback_data="list_xlsx:dork"),
        ],
        [
            InlineKeyboardButton("📋 Direktori", callback_data="list_xlsx:directory"),
        ],
        [
            InlineKeyboardButton("🗑️ Hapus File", callback_data="delete_xlsx:all"),
        ],
        [InlineKeyboardButton("⬅️ Kembali", callback_data="back_main")],
    ]
    return InlineKeyboardMarkup(keyboard)


def build_xlsx_file_menu(files: list, action: str, scraping_type: str):
    keyboard = []
    for idx, fileinfo in enumerate(files[:20]):
        label = f"{idx + 1}. {fileinfo['filename']} ({fileinfo['rows']} email)"
        keyboard.append([InlineKeyboardButton(label, callback_data=f"{action}:{scraping_type}:{idx}")])
    keyboard.append([InlineKeyboardButton("⬅️ Kembali", callback_data="manage_xlsx")])
    return InlineKeyboardMarkup(keyboard)


def get_selected_output_filename(context: ContextTypes.DEFAULT_TYPE, scraping_type: str):
    return context.user_data.get(f"selected_{scraping_type}_xlsx") or get_master_filename(scraping_type)


async def safe_edit_message(message, text, reply_markup=None):
    try:
        await message.edit_text(text, reply_markup=reply_markup)
    except Exception:
        pass


# =====================================================
# SERPAPI HELPERS
# =====================================================

async def serpapi_search_maps(keyword_location: str):
    if not SERPAPI_KEY:
        raise ValueError("SERPAPI key belum diatur di file .env")
    params = {
        "engine": "google_maps",
        "q": keyword_location,
        "type": "search",
        "api_key": SERPAPI_KEY,
    }
    async with aiohttp.ClientSession() as session:
        async with session.get("https://serpapi.com/search.json", params=params, timeout=REQUEST_TIMEOUT) as response:
            if response.status == 401:
                raise ValueError("SERPAPI key tidak valid.")
            if response.status == 429:
                raise ValueError("Limit SerpAPI sudah tercapai.")
            data = await response.json()
    if "error" in data:
        raise ValueError(data["error"])
    return data.get("local_results", [])[:MAPS_MAX_BUSINESSES]


async def serpapi_search_web(query: str):
    """Gunakan SerpAPI untuk Google Web Search (dorking)."""
    if not SERPAPI_KEY:
        raise ValueError("SERPAPI key belum diatur di file .env")
    params = {
        "engine": "google",
        "q": query,
        "num": DORK_MAX_RESULTS,
        "api_key": SERPAPI_KEY,
    }
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get("https://serpapi.com/search.json", params=params, timeout=REQUEST_TIMEOUT) as response:
                if response.status != 200:
                    return []
                data = await response.json()
        results = data.get("organic_results", [])
        urls = [r.get("link") for r in results if r.get("link")]
        return urls
    except Exception as e:
        logging.warning(f"SerpAPI web search error: {e}")
        return []


def split_keyword_location(text: str):
    parts = text.strip().split()
    if len(parts) <= 1:
        return text.strip(), ""
    location = parts[-1]
    keyword = " ".join(parts[:-1])
    return keyword, location


def parse_target_from_input(text: str):
    """
    Parse input yang mungkin menyertakan target email.
    Format: 'keyword atau URL | 500'
    Kalau tidak ada pipe, pakai DEFAULT_TARGET_EMAILS.
    """
    if "|" in text:
        parts = text.rsplit("|", 1)
        main_input = parts[0].strip()
        try:
            target = int(parts[1].strip())
            target = max(1, min(target, 10000))
        except ValueError:
            target = DEFAULT_TARGET_EMAILS
    else:
        main_input = text.strip()
        target = DEFAULT_TARGET_EMAILS
    return main_input, target


# =====================================================
# CORE: WEB SCRAPER
# =====================================================

async def scan_single_url(session, url: str, max_pages: int, start_url: str | None = None,
                          found_email_set: set = None, found_domain_set: set = None,
                          use_playwright_fallback: bool = False):
    """
    Scan satu URL (dan halaman internalnya) untuk email.
    Mengembalikan list (email, status, source_url).
    """
    if found_email_set is None:
        found_email_set = set()
    if found_domain_set is None:
        found_domain_set = set()

    root = start_url or url
    queue = [(normalize_crawl_url(url), 0)]
    queued = {normalize_crawl_url(url)}
    visited = set()
    found = []
    pages_scanned = 0

    while queue and pages_scanned < max_pages:
        queue.sort(key=lambda item: (item[1], link_priority(item[0])))
        current_url, depth = queue.pop(0)
        queued.discard(current_url)

        if current_url in visited:
            continue

        visited.add(current_url)
        pages_scanned += 1

        html, error = await fetch_html(session, current_url)

        # Fallback ke Playwright kalau HTML kosong dan halaman kemungkinan butuh JS
        if not html and use_playwright_fallback and PLAYWRIGHT_AVAILABLE:
            html, error = await fetch_html_playwright(current_url)

        if not html:
            # Coba ekstrak dari PDF kalau URL-nya .pdf
            if current_url.lower().endswith(".pdf") and PYMUPDF_AVAILABLE:
                pdf_bytes = await fetch_pdf_bytes(session, current_url)
                if pdf_bytes:
                    for email, status in extract_emails_from_pdf_bytes(pdf_bytes):
                        found.append((email, status, current_url))
            continue

        for email, status in extract_emails_from_text(html):
            found.append((email, status, current_url))

        # Scan link PDF yang ada di halaman
        if PYMUPDF_AVAILABLE:
            soup = BeautifulSoup(html, "html.parser")
            for a_tag in soup.find_all("a", href=True):
                href = a_tag.get("href", "")
                if href.lower().endswith(".pdf"):
                    pdf_url = urljoin(current_url, href)
                    if pdf_url not in visited:
                        pdf_bytes = await fetch_pdf_bytes(session, pdf_url)
                        if pdf_bytes:
                            for email, status in extract_emails_from_pdf_bytes(pdf_bytes):
                                found.append((email, status, pdf_url))
                        visited.add(pdf_url)

        if depth < WEB_MAX_DEPTH:
            for link in sorted(extract_links(current_url, html), key=link_priority):
                clean_link = normalize_crawl_url(link)
                if not same_domain(root, clean_link):
                    continue
                if clean_link in visited or clean_link in queued:
                    continue
                queue.append((clean_link, depth + 1))
                queued.add(clean_link)

    return found, pages_scanned


async def scrape_website(start_url: str, progress_message, target_emails=DEFAULT_TARGET_EMAILS,
                         output_filename: str | None = None):
    visited_urls = set()
    queued_urls = [(normalize_crawl_url(start_url), 0)]
    queued_url_set = {normalize_crawl_url(start_url)}

    _, existing_master_emails, existing_master_domains, _ = load_existing_master("web", output_filename)
    found_email_set = set(existing_master_emails)
    found_domain_set = set(existing_master_domains)
    session_email_count_start = len(found_email_set)
    rows = []

    scanned_pages = 0
    failed_pages = 0
    skipped_duplicate = 0
    latest_preview = []
    skipped_preview = []
    skipped_invalid_email = 0
    skipped_duplicate_email = 0
    skipped_same_domain = 0
    last_error = "-"
    start_time = time.time()
    stop_reason = "Selesai"

    async with aiohttp.ClientSession() as session:
        while queued_urls:
            queued_urls.sort(key=lambda item: (item[1], link_priority(item[0])))
            current_url, depth = queued_urls.pop(0)
            queued_url_set.discard(current_url)

            if current_url in visited_urls:
                skipped_duplicate += 1
                continue

            if scanned_pages >= WEB_MAX_PAGES:
                stop_reason = "Batas halaman tercapai"
                break

            if (len(found_email_set) - session_email_count_start) >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            visited_urls.add(current_url)
            scanned_pages += 1
            extracted = []

            html, error = await fetch_html(session, current_url)

            # Playwright fallback untuk halaman JS-heavy
            if not html and PLAYWRIGHT_AVAILABLE:
                html, error = await fetch_html_playwright(current_url)

            if html:
                extracted = extract_emails_from_text(html)

                # Scan PDF yang ditemukan di halaman
                if PYMUPDF_AVAILABLE:
                    soup = BeautifulSoup(html, "html.parser")
                    for a_tag in soup.find_all("a", href=True):
                        href = a_tag.get("href", "")
                        if href.lower().endswith(".pdf"):
                            pdf_url = urljoin(current_url, href)
                            if pdf_url not in visited_urls:
                                pdf_bytes = await fetch_pdf_bytes(session, pdf_url)
                                if pdf_bytes:
                                    for e, s in extract_emails_from_pdf_bytes(pdf_bytes):
                                        extracted.append((e, s))
                                visited_urls.add(pdf_url)

                for email, status in extracted:
                    email_domain = get_email_domain(email)
                    if status != "Valid":
                        skipped_invalid_email += 1
                        skipped_preview.insert(0, f"🚫 {email}")
                        skipped_preview = skipped_preview[:5]
                        continue
                    if email in found_email_set:
                        skipped_duplicate_email += 1
                        continue
                    if email_domain in found_domain_set:
                        skipped_same_domain += 1
                        continue

                    allowed, checker_status = await is_email_allowed_by_checker(session, email)
                    if not allowed:
                        skipped_invalid_email += 1
                        skipped_preview.insert(0, f"🚫 {email} | Checker: {checker_status}")
                        skipped_preview = skipped_preview[:5]
                        continue

                    found_email_set.add(email)
                    found_domain_set.add(email_domain)
                    rows.append([
                        len(rows) + 1, current_url, email, "Valid",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Web Scraping",
                    ])
                    latest_preview.insert(0, f"✅ {email} | {email_domain}")
                    latest_preview = latest_preview[:10]

                    if rows and len(rows) % PARTIAL_SAVE_EVERY == 0:
                        save_excel(rows, f"partial_web_{int(time.time())}.xlsx", "web")

                if depth < WEB_MAX_DEPTH:
                    for link in sorted(extract_links(current_url, html), key=link_priority):
                        clean_link = normalize_crawl_url(link)
                        if not same_domain(start_url, clean_link):
                            continue
                        if clean_link in visited_urls or clean_link in queued_url_set:
                            skipped_duplicate += 1
                            continue
                        queued_urls.append((clean_link, depth + 1))
                        queued_url_set.add(clean_link)
            else:
                failed_pages += 1
                last_error = f"{current_url} -> {error}"

            new_email_count = len(found_email_set) - session_email_count_start
            percent = min(max(
                int((scanned_pages / WEB_MAX_PAGES) * 100),
                int((new_email_count / max(target_emails, 1)) * 100)
            ), 100)
            elapsed = int(time.time() - start_time)

            should_update = (
                scanned_pages == 1
                or scanned_pages % PROGRESS_UPDATE_EVERY == 0
                or bool(extracted)
                or not queued_urls
            )

            if should_update:
                progress_text = (
                    "🌐 *Progress Scraping Web*\n\n"
                    f"{make_progress_bar(percent)} {percent}%\n\n"
                    f"📄 Halaman: {current_url}\n"
                    f"📊 Discan: {scanned_pages}/{WEB_MAX_PAGES} | Kedalaman: {depth}/{WEB_MAX_DEPTH}\n"
                    f"✅ Valid: {len(rows)} | ❌ Invalid: {skipped_invalid_email}\n"
                    f"♻️ Dup email: {skipped_duplicate_email} | 🏷️ Dup domain: {skipped_same_domain}\n"
                    f"📧 Baru: {new_email_count}/{target_emails}\n"
                    f"⏱️ {elapsed}s | Error: {last_error[:60]}\n"
                    f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                    f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
                )
                await safe_edit_message(progress_message, progress_text)

            await asyncio.sleep(REQUEST_DELAY)

    final_text = (
        "✅ *Scraping Web Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Halaman discan: {scanned_pages}\n"
        f"Email valid: {len(rows)}\n"
        f"Email baru: {len(found_email_set) - session_email_count_start}"
    )
    await safe_edit_message(progress_message, final_text)
    return rows


# =====================================================
# CORE: GOOGLE MAPS SCRAPER
# =====================================================

async def scan_business_website(session, website: str, max_pages: int = MAPS_WEBSITE_MAX_PAGES):
    website = normalize_url(website)
    found, pages_scanned = await scan_single_url(session, website, max_pages, start_url=website)
    return found, pages_scanned, "-"


async def scrape_maps(keyword_location: str, progress_message, target_emails=DEFAULT_TARGET_EMAILS,
                      output_filename: str | None = None):
    listings = await serpapi_search_maps(keyword_location)
    keyword, location = split_keyword_location(keyword_location)
    rows = []
    _, existing_master_emails, existing_master_domains, _ = load_existing_master("maps", output_filename)
    found_email_set = set(existing_master_emails)
    found_domain_set = set(existing_master_domains)
    session_email_count_start = len(found_email_set)

    total_checked = 0
    websites_scanned = 0
    pages_scanned = 0
    valid_count = 0
    no_website_count = 0
    no_email_count = 0
    error_count = 0
    skipped_invalid_email = 0
    skipped_duplicate_email = 0
    skipped_same_domain = 0
    latest_preview = []
    start_time = time.time()
    stop_reason = "Selesai"

    async with aiohttp.ClientSession() as session:
        for business in listings:
            if (len(found_email_set) - session_email_count_start) >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            total_checked += 1
            business_name = business.get("title", "Unknown Business")
            website = business.get("website")
            business_address = business.get("address", location)

            if not website:
                no_website_count += 1
            else:
                websites_scanned += 1
                extracted, wp, last_error = await scan_business_website(session, website)
                pages_scanned += wp

                if extracted:
                    for email, status, source_url in extracted:
                        email_domain = get_email_domain(email)
                        if status != "Valid":
                            skipped_invalid_email += 1
                            continue
                        if email in found_email_set:
                            skipped_duplicate_email += 1
                            continue
                        if email_domain in found_domain_set:
                            skipped_same_domain += 1
                            continue

                        allowed, checker_status = await is_email_allowed_by_checker(session, email)
                        if not allowed:
                            skipped_invalid_email += 1
                            continue

                        found_email_set.add(email)
                        found_domain_set.add(email_domain)
                        valid_count += 1
                        rows.append([
                            len(rows) + 1, business_name, keyword, business_address,
                            source_url, email, "Valid",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Google Maps Scraping",
                        ])
                        latest_preview.insert(0, f"✅ {email} | {business_name}")
                        latest_preview = latest_preview[:10]

                        if rows and len(rows) % PARTIAL_SAVE_EVERY == 0:
                            save_excel(rows, f"partial_maps_{int(time.time())}.xlsx", "maps")

                        if (len(found_email_set) - session_email_count_start) >= target_emails:
                            stop_reason = "Target email baru tercapai"
                            break
                else:
                    no_email_count += 1

            new_email_count = len(found_email_set) - session_email_count_start
            percent = min(max(
                int((total_checked / max(len(listings), 1)) * 100),
                int((new_email_count / max(target_emails, 1)) * 100)
            ), 100)
            elapsed = int(time.time() - start_time)

            progress_text = (
                "🗺️ *Progress Scraping Google Maps*\n\n"
                f"{make_progress_bar(percent)} {percent}%\n\n"
                f"🏢 {business_name}\n"
                f"📊 Bisnis: {total_checked}/{len(listings)} | Website discan: {websites_scanned}\n"
                f"✅ Valid: {valid_count} | ❌ Invalid: {skipped_invalid_email}\n"
                f"♻️ Dup email: {skipped_duplicate_email} | 🏷️ Dup domain: {skipped_same_domain}\n"
                f"📧 Baru: {new_email_count}/{target_emails}\n"
                f"⏱️ {elapsed}s\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )
            await safe_edit_message(progress_message, progress_text)
            await asyncio.sleep(REQUEST_DELAY)

    final_text = (
        "✅ *Scraping Google Maps Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Bisnis dicek: {total_checked}/{len(listings)}\n"
        f"Email valid: {valid_count}\n"
        f"Email baru: {len(found_email_set) - session_email_count_start}"
    )
    await safe_edit_message(progress_message, final_text)
    return rows


# =====================================================
# CORE: GOOGLE DORKING
# =====================================================

async def scrape_dork(keyword_location: str, progress_message, target_emails=DEFAULT_TARGET_EMAILS,
                      output_filename: str | None = None):
    """
    Gunakan SerpAPI Google Web Search dengan dork query untuk menemukan URL yang mengandung email.
    Kemudian crawl tiap URL hasil dork untuk ekstrak email.
    """
    keyword, location = split_keyword_location(keyword_location)

    _, existing_master_emails, existing_master_domains, _ = load_existing_master("dork", output_filename)
    found_email_set = set(existing_master_emails)
    found_domain_set = set(existing_master_domains)
    session_email_count_start = len(found_email_set)
    rows = []

    total_queries = 0
    total_urls_found = 0
    total_urls_scanned = 0
    skipped_invalid_email = 0
    skipped_duplicate_email = 0
    skipped_same_domain = 0
    latest_preview = []
    start_time = time.time()
    stop_reason = "Selesai"

    async with aiohttp.ClientSession() as session:
        for template in DORK_TEMPLATES:
            if (len(found_email_set) - session_email_count_start) >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            query = template.format(keyword=keyword, location=location or "Indonesia")
            total_queries += 1

            new_email_count = len(found_email_set) - session_email_count_start
            percent = min(int((new_email_count / max(target_emails, 1)) * 100), 100)
            elapsed = int(time.time() - start_time)

            progress_text = (
                "🔍 *Progress Google Dorking*\n\n"
                f"{make_progress_bar(percent)} {percent}%\n\n"
                f"🔎 Query {total_queries}/{len(DORK_TEMPLATES)}:\n{query[:80]}\n\n"
                f"🌐 URL ditemukan: {total_urls_found} | Discan: {total_urls_scanned}\n"
                f"✅ Valid: {len(rows)} | ❌ Invalid: {skipped_invalid_email}\n"
                f"♻️ Dup email: {skipped_duplicate_email} | 🏷️ Dup domain: {skipped_same_domain}\n"
                f"📧 Baru: {new_email_count}/{target_emails}\n"
                f"⏱️ {elapsed}s\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )
            await safe_edit_message(progress_message, progress_text)

            urls = await serpapi_search_web(query)
            total_urls_found += len(urls)

            for url in urls:
                if (len(found_email_set) - session_email_count_start) >= target_emails:
                    stop_reason = "Target email baru tercapai"
                    break

                if not is_valid_url(url):
                    continue

                total_urls_scanned += 1

                extracted, _ = await scan_single_url(
                    session, url, DORK_MAX_PAGES_PER_URL, start_url=url,
                    use_playwright_fallback=True
                )

                for email, status, source_url in extracted:
                    email_domain = get_email_domain(email)
                    if status != "Valid":
                        skipped_invalid_email += 1
                        continue
                    if email in found_email_set:
                        skipped_duplicate_email += 1
                        continue
                    if email_domain in found_domain_set:
                        skipped_same_domain += 1
                        continue

                    allowed, checker_status = await is_email_allowed_by_checker(session, email)
                    if not allowed:
                        skipped_invalid_email += 1
                        continue

                    found_email_set.add(email)
                    found_domain_set.add(email_domain)
                    rows.append([
                        len(rows) + 1, query, source_url, email, "Valid",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Google Dorking",
                    ])
                    latest_preview.insert(0, f"✅ {email} | {email_domain}")
                    latest_preview = latest_preview[:10]

                    if rows and len(rows) % PARTIAL_SAVE_EVERY == 0:
                        save_excel(rows, f"partial_dork_{int(time.time())}.xlsx", "dork")

                await asyncio.sleep(REQUEST_DELAY)

            await asyncio.sleep(1.0)  # jeda antar query dork

    final_text = (
        "✅ *Google Dorking Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Query dijalankan: {total_queries}/{len(DORK_TEMPLATES)}\n"
        f"URL ditemukan: {total_urls_found} | Discan: {total_urls_scanned}\n"
        f"Email valid: {len(rows)}\n"
        f"Email baru: {len(found_email_set) - session_email_count_start}"
    )
    await safe_edit_message(progress_message, final_text)
    return rows


# =====================================================
# CORE: DIRECTORY SCRAPER
# =====================================================

async def scrape_directory(keyword_location: str, progress_message, target_emails=DEFAULT_TARGET_EMAILS,
                           output_filename: str | None = None):
    """
    Scrape direktori bisnis publik (Yellowpages Indonesia, Clutch.co) untuk mendapatkan
    website bisnis, lalu crawl website tersebut untuk mendapatkan email.
    """
    keyword, location = split_keyword_location(keyword_location)

    _, existing_master_emails, existing_master_domains, _ = load_existing_master("directory", output_filename)
    found_email_set = set(existing_master_emails)
    found_domain_set = set(existing_master_domains)
    session_email_count_start = len(found_email_set)
    rows = []

    total_listings = 0
    total_websites_scanned = 0
    skipped_invalid_email = 0
    skipped_duplicate_email = 0
    skipped_same_domain = 0
    latest_preview = []
    start_time = time.time()
    stop_reason = "Selesai"

    async with aiohttp.ClientSession() as session:
        for dir_key, dir_config in DIRECTORIES.items():
            if (len(found_email_set) - session_email_count_start) >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            dir_name = dir_config["name"]

            for page in range(1, DIRECTORY_MAX_PAGES + 1):
                if (len(found_email_set) - session_email_count_start) >= target_emails:
                    stop_reason = "Target email baru tercapai"
                    break

                listing_url = dir_config["search_url"].format(
                    keyword=keyword.replace(" ", "+"),
                    location=location.replace(" ", "+"),
                    page=page,
                )

                new_email_count = len(found_email_set) - session_email_count_start
                percent = min(int((new_email_count / max(target_emails, 1)) * 100), 100)
                elapsed = int(time.time() - start_time)

                progress_text = (
                    "📋 *Progress Scraping Direktori*\n\n"
                    f"{make_progress_bar(percent)} {percent}%\n\n"
                    f"📁 Direktori: {dir_name} | Hal: {page}\n"
                    f"🏢 Listing ditemukan: {total_listings}\n"
                    f"🌐 Website discan: {total_websites_scanned}\n"
                    f"✅ Valid: {len(rows)} | ❌ Invalid: {skipped_invalid_email}\n"
                    f"♻️ Dup email: {skipped_duplicate_email} | 🏷️ Dup domain: {skipped_same_domain}\n"
                    f"📧 Baru: {new_email_count}/{target_emails}\n"
                    f"⏱️ {elapsed}s\n\n"
                    f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
                )
                await safe_edit_message(progress_message, progress_text)

                html, error = await fetch_html(session, listing_url)
                if not html:
                    break  # tidak ada halaman listing, lanjut direktori berikutnya

                soup = BeautifulSoup(html, "html.parser")
                listings_found = soup.select(dir_config["listing_selector"])

                if not listings_found:
                    break  # halaman kosong, berhenti paginasi

                for listing in listings_found:
                    total_listings += 1

                    # Ambil nama bisnis
                    name_tag = listing.select_one(dir_config["name_selector"])
                    business_name = name_tag.get_text(strip=True) if name_tag else "Unknown"

                    # Ambil URL website bisnis
                    url_tag = listing.select_one(dir_config["url_selector"])
                    if not url_tag:
                        continue
                    website = url_tag.get("href", "").strip()
                    if not website or not is_valid_url(website):
                        continue

                    total_websites_scanned += 1
                    extracted, _ = await scan_single_url(
                        session, website, DIRECTORY_DETAIL_MAX, start_url=website
                    )

                    for email, status, source_url in extracted:
                        email_domain = get_email_domain(email)
                        if status != "Valid":
                            skipped_invalid_email += 1
                            continue
                        if email in found_email_set:
                            skipped_duplicate_email += 1
                            continue
                        if email_domain in found_domain_set:
                            skipped_same_domain += 1
                            continue

                        allowed, checker_status = await is_email_allowed_by_checker(session, email)
                        if not allowed:
                            skipped_invalid_email += 1
                            continue

                        found_email_set.add(email)
                        found_domain_set.add(email_domain)
                        rows.append([
                            len(rows) + 1, dir_name, business_name, source_url, email, "Valid",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Directory Scraping",
                        ])
                        latest_preview.insert(0, f"✅ {email} | {business_name}")
                        latest_preview = latest_preview[:10]

                        if rows and len(rows) % PARTIAL_SAVE_EVERY == 0:
                            save_excel(rows, f"partial_dir_{int(time.time())}.xlsx", "directory")

                        if (len(found_email_set) - session_email_count_start) >= target_emails:
                            stop_reason = "Target email baru tercapai"
                            break

                    await asyncio.sleep(REQUEST_DELAY)

                await asyncio.sleep(1.0)  # jeda antar halaman listing

    final_text = (
        "✅ *Scraping Direktori Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Listing ditemukan: {total_listings}\n"
        f"Website discan: {total_websites_scanned}\n"
        f"Email valid: {len(rows)}\n"
        f"Email baru: {len(found_email_set) - session_email_count_start}"
    )
    await safe_edit_message(progress_message, final_text)
    return rows


# =====================================================
# CORE: BATCH URL SCRAPER
# =====================================================

async def scrape_batch(urls: list, progress_message, target_emails=DEFAULT_TARGET_EMAILS,
                       output_filename: str | None = None):
    """
    Scrape banyak URL sekaligus dari file .txt.
    Setiap URL di-crawl dengan crawler web biasa, hasilnya digabung ke satu XLSX.
    """
    _, existing_master_emails, existing_master_domains, _ = load_existing_master("web", output_filename)
    found_email_set = set(existing_master_emails)
    found_domain_set = set(existing_master_domains)
    session_email_count_start = len(found_email_set)
    all_rows = []

    total_urls = len(urls)
    scanned_urls = 0
    total_pages = 0
    skipped_invalid_email = 0
    skipped_duplicate_email = 0
    skipped_same_domain = 0
    latest_preview = []
    start_time = time.time()

    async with aiohttp.ClientSession() as session:
        for url in urls:
            url = normalize_url(url.strip())
            if not is_valid_url(url):
                continue

            scanned_urls += 1
            new_email_count = len(found_email_set) - session_email_count_start
            percent = min(max(
                int((scanned_urls / max(total_urls, 1)) * 100),
                int((new_email_count / max(target_emails, 1)) * 100)
            ), 100)
            elapsed = int(time.time() - start_time)

            progress_text = (
                "📦 *Progress Batch URL Scraping*\n\n"
                f"{make_progress_bar(percent)} {percent}%\n\n"
                f"🌐 URL {scanned_urls}/{total_urls}:\n{url[:80]}\n\n"
                f"📄 Total halaman discan: {total_pages}\n"
                f"✅ Valid: {len(all_rows)} | ❌ Invalid: {skipped_invalid_email}\n"
                f"♻️ Dup email: {skipped_duplicate_email} | 🏷️ Dup domain: {skipped_same_domain}\n"
                f"📧 Baru: {new_email_count}/{target_emails}\n"
                f"⏱️ {elapsed}s\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )
            await safe_edit_message(progress_message, progress_text)

            extracted, pages_scanned = await scan_single_url(
                session, url, WEB_MAX_PAGES // max(total_urls, 1) + 10,
                start_url=url, use_playwright_fallback=True
            )
            total_pages += pages_scanned

            for email, status, source_url in extracted:
                email_domain = get_email_domain(email)
                if status != "Valid":
                    skipped_invalid_email += 1
                    continue
                if email in found_email_set:
                    skipped_duplicate_email += 1
                    continue
                if email_domain in found_domain_set:
                    skipped_same_domain += 1
                    continue

                allowed, checker_status = await is_email_allowed_by_checker(session, email)
                if not allowed:
                    skipped_invalid_email += 1
                    continue

                found_email_set.add(email)
                found_domain_set.add(email_domain)
                all_rows.append([
                    len(all_rows) + 1, url, email, "Valid",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Batch Web Scraping",
                ])
                latest_preview.insert(0, f"✅ {email} | {email_domain}")
                latest_preview = latest_preview[:10]

            if (len(found_email_set) - session_email_count_start) >= target_emails:
                break

            await asyncio.sleep(REQUEST_DELAY)

    final_text = (
        "✅ *Batch Scraping Selesai*\n\n"
        f"URL diproses: {scanned_urls}/{total_urls}\n"
        f"Halaman discan: {total_pages}\n"
        f"Email valid: {len(all_rows)}\n"
        f"Email baru: {len(found_email_set) - session_email_count_start}"
    )
    await safe_edit_message(progress_message, final_text)
    return all_rows


# =====================================================
# TELEGRAM HANDLERS
# =====================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    features = []
    if PLAYWRIGHT_AVAILABLE:
        features.append("✅ Playwright (JS rendering)")
    else:
        features.append("⚠️ Playwright tidak tersedia (install: pip install playwright && playwright install chromium)")
    if PYMUPDF_AVAILABLE:
        features.append("✅ PyMuPDF (PDF scraping)")
    else:
        features.append("⚠️ PyMuPDF tidak tersedia (install: pip install pymupdf)")

    if DNS_AVAILABLE:
        features.append("✅ MX checker (dnspython)")
    else:
        features.append("⚠️ MX checker tidak tersedia (install: pip install dnspython)")

    features.append(f"📡 Verifalia: {get_verifalia_status_text()}")

    await update.message.reply_text(
        "🤖 *Email Scraper Bot*\n\n"
        "Fitur aktif:\n" + "\n".join(features) + "\n\n"
        "Pilih mode scraping:",
        reply_markup=main_menu(),
        parse_mode="Markdown",
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "back_main":
        context.user_data.clear()
        await query.edit_message_text("Pilih mode scraping:", reply_markup=main_menu())
        return

    if data == "manage_xlsx":
        await query.edit_message_text(
            "📁 Kelola file XLSX hasil scraping.",
            reply_markup=manage_xlsx_menu(),
        )
        return

    if data.startswith("list_xlsx:"):
        scraping_type = data.split(":", 1)[1]
        files = list_xlsx_files(scraping_type)
        if not files:
            await query.edit_message_text(
                f"Belum ada file XLSX untuk {scraping_type.upper()}.",
                reply_markup=manage_xlsx_menu()
            )
            return
        lines = [f"📄 Daftar XLSX {scraping_type.upper()}:\n"]
        for idx, fileinfo in enumerate(files[:20], start=1):
            lines.append(
                f"{idx}. {fileinfo['filename']}\n"
                f"   Email: {fileinfo['rows']} | Ukuran: {fileinfo['size_kb']:.1f} KB | Update: {fileinfo['modified']}"
            )
        await query.edit_message_text("\n".join(lines), reply_markup=manage_xlsx_menu())
        return

    if data.startswith("delete_xlsx:"):
        scraping_type = data.split(":", 1)[1]
        files = list_xlsx_files(None)  # semua file
        context.user_data[f"delete_files_{scraping_type}"] = files
        if not files:
            await query.edit_message_text("Belum ada file XLSX.", reply_markup=manage_xlsx_menu())
            return
        await query.edit_message_text(
            "🗑️ Pilih file XLSX yang mau dihapus:",
            reply_markup=build_xlsx_file_menu(files, "confirm_delete_xlsx", scraping_type),
        )
        return

    if data.startswith("confirm_delete_xlsx:"):
        _, scraping_type, idx_text = data.split(":")
        files = context.user_data.get(f"delete_files_{scraping_type}", [])
        try:
            fileinfo = files[int(idx_text)]
        except Exception:
            await query.edit_message_text("File tidak ditemukan.", reply_markup=manage_xlsx_menu())
            return
        context.user_data["delete_target"] = fileinfo
        keyboard = [
            [InlineKeyboardButton("✅ Ya, hapus", callback_data="do_delete_xlsx")],
            [InlineKeyboardButton("❌ Batal", callback_data="manage_xlsx")],
        ]
        await query.edit_message_text(
            f"Yakin hapus?\n{fileinfo['filename']}\nEmail: {fileinfo['rows']}",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return

    if data == "do_delete_xlsx":
        fileinfo = context.user_data.get("delete_target")
        if not fileinfo:
            await query.edit_message_text("Tidak ada file yang dipilih.", reply_markup=manage_xlsx_menu())
            return
        try:
            path = fileinfo["path"]
            if os.path.dirname(os.path.abspath(path)) == os.path.abspath(OUTPUT_DIR) and path.lower().endswith(".xlsx"):
                os.remove(path)
                await query.edit_message_text(f"✅ File dihapus:\n{fileinfo['filename']}", reply_markup=manage_xlsx_menu())
            else:
                await query.edit_message_text("File tidak aman untuk dihapus.", reply_markup=manage_xlsx_menu())
        except Exception as e:
            await query.edit_message_text(f"Gagal hapus file:\n{e}", reply_markup=manage_xlsx_menu())
        return

    # Mode scraping
    for mode_key, label, prompt in [
        ("scrape_web", "web", "🌐 *Scraping Web*\n\nFormat input:\n`https://example.com` atau\n`https://example.com | 300` (dengan target email)\n\nPilih output XLSX:"),
        ("scrape_maps", "maps", "🗺️ *Scraping Google Maps*\n\nFormat input:\n`digital agency Jakarta` atau\n`digital agency Jakarta | 200`\n\nPilih output XLSX:"),
        ("scrape_dork", "dork", "🔍 *Google Dorking*\n\nFormat input:\n`digital agency Jakarta` atau\n`digital agency Jakarta | 500`\n\nPilih output XLSX:"),
        ("scrape_directory", "directory", "📋 *Scraping Direktori Bisnis*\n\nFormat input:\n`digital agency Jakarta` atau\n`digital agency Jakarta | 300`\n\nPilih output XLSX:"),
        ("scrape_batch", "web", "📦 *Batch URL Scraping*\n\nKirim file .txt berisi daftar URL (1 URL per baris).\nAtau ketik daftar URL langsung (pisah enter).\n\nPilih output XLSX dulu:"),
    ]:
        if data == mode_key:
            context.user_data.clear()
            context.user_data["pending_mode"] = label
            context.user_data["actual_mode"] = mode_key.replace("scrape_", "")
            if mode_key == "scrape_batch":
                context.user_data["awaiting_batch_file"] = True
            await query.edit_message_text(prompt, reply_markup=output_choice_menu(label), parse_mode="Markdown")
            return

    if data.startswith("xlsx_new:"):
        scraping_type = data.split(":", 1)[1]
        context.user_data["pending_mode"] = scraping_type
        context.user_data["awaiting_new_xlsx_name"] = True
        await query.edit_message_text(
            "🆕 Kirim nama file XLSX baru.\nContoh: leads_jakarta.xlsx"
        )
        return

    if data.startswith("xlsx_old:"):
        scraping_type = data.split(":", 1)[1]
        files = list_xlsx_files(scraping_type)
        context.user_data[f"select_files_{scraping_type}"] = files
        if not files:
            await query.edit_message_text(
                f"Belum ada XLSX lama untuk {scraping_type.upper()}. Buat file baru dulu.",
                reply_markup=output_choice_menu(scraping_type)
            )
            return
        await query.edit_message_text(
            f"📂 Pilih XLSX lama untuk output {scraping_type.upper()}:",
            reply_markup=build_xlsx_file_menu(files, "select_xlsx", scraping_type),
        )
        return

    if data.startswith("select_xlsx:"):
        _, scraping_type, idx_text = data.split(":")
        files = context.user_data.get(f"select_files_{scraping_type}", [])
        try:
            fileinfo = files[int(idx_text)]
        except Exception:
            await query.edit_message_text("File tidak ditemukan.", reply_markup=output_choice_menu(scraping_type))
            return

        context.user_data["mode"] = scraping_type
        context.user_data[f"selected_{scraping_type}_xlsx"] = fileinfo["filename"]

        actual_mode = context.user_data.get("actual_mode", scraping_type)
        if actual_mode == "batch":
            await query.edit_message_text(
                f"✅ Output: {fileinfo['filename']}\n\n"
                "📦 Sekarang kirim file .txt berisi daftar URL\n"
                "atau ketik URL langsung (1 per baris).\n\n"
                "Format opsional tambahkan target di baris pertama:\n"
                "`TARGET:500`"
            )
        else:
            prompts = {
                "web": "🌐 Kirim URL website.\nContoh: https://example.com\nAtau: https://example.com | 300",
                "maps": "🗺️ Kirim keyword + lokasi.\nContoh: digital agency Jakarta\nAtau: digital agency Jakarta | 200",
                "dork": "🔍 Kirim keyword + lokasi untuk dorking.\nContoh: digital agency Jakarta | 500",
                "directory": "📋 Kirim keyword + lokasi.\nContoh: digital agency Jakarta | 300",
            }
            await query.edit_message_text(
                f"✅ Output: {fileinfo['filename']}\n\n" + prompts.get(actual_mode, "Kirim input:")
            )
        return


async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode")
    actual_mode = context.user_data.get("actual_mode", mode)
    user_input = update.message.text.strip() if update.message.text else ""

    # ------------------------------------------------
    # Handle nama file XLSX baru
    # ------------------------------------------------
    if context.user_data.get("awaiting_new_xlsx_name"):
        scraping_type = context.user_data.get("pending_mode")
        filename = safe_xlsx_name(user_input, scraping_type)
        filepath = os.path.join(OUTPUT_DIR, filename)

        if os.path.exists(filepath):
            await update.message.reply_text(
                "Nama file sudah ada. Pilih XLSX lama atau kirim nama berbeda.",
                reply_markup=output_choice_menu(scraping_type),
            )
            context.user_data.pop("awaiting_new_xlsx_name", None)
            return

        pd.DataFrame(columns=get_columns(scraping_type)).to_excel(filepath, index=False)
        context.user_data["awaiting_new_xlsx_name"] = False
        context.user_data["mode"] = scraping_type
        context.user_data[f"selected_{scraping_type}_xlsx"] = filename

        am = context.user_data.get("actual_mode", scraping_type)
        prompts = {
            "web": "🌐 Kirim URL website.\nContoh: https://example.com | 300",
            "maps": "🗺️ Kirim keyword + lokasi.\nContoh: digital agency Jakarta | 200",
            "dork": "🔍 Kirim keyword + lokasi.\nContoh: digital agency Jakarta | 500",
            "directory": "📋 Kirim keyword + lokasi.\nContoh: digital agency Jakarta | 300",
            "batch": "📦 Kirim file .txt atau ketik URL langsung (1 per baris).",
        }
        await update.message.reply_text(
            f"✅ File baru: {filename}\n\n" + prompts.get(am, "Kirim input:")
        )
        return

    # ------------------------------------------------
    # Handle file dokumen (batch .txt atau PDF)
    # ------------------------------------------------
    if update.message.document:
        doc: Document = update.message.document
        file_name = doc.file_name or ""

        # Batch URL dari .txt
        if file_name.lower().endswith(".txt") and actual_mode == "batch":
            file_obj = await doc.get_file()
            file_bytes = await file_obj.download_as_bytearray()
            content = file_bytes.decode("utf-8", errors="ignore")
            lines = [l.strip() for l in content.splitlines() if l.strip()]

            target = DEFAULT_TARGET_EMAILS
            urls = []
            for line in lines:
                if line.upper().startswith("TARGET:"):
                    try:
                        target = int(line.split(":", 1)[1].strip())
                    except ValueError:
                        pass
                elif is_valid_url(normalize_url(line)):
                    urls.append(line)

            if not urls:
                await update.message.reply_text("Tidak ada URL valid di file .txt.")
                return

            progress_message = await update.message.reply_text(
                f"📦 Memulai batch scraping {len(urls)} URL... Target: {target} email"
            )
            output_filename = get_selected_output_filename(context, "web")
            rows = await scrape_batch(urls, progress_message, target_emails=target, output_filename=output_filename)

            if not rows:
                await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added_rows, total_rows = append_to_master_excel(rows, "web", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Batch selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
            context.user_data.clear()
            return

        await update.message.reply_text(
            "File tidak dikenali. Kirim .txt untuk batch URL.",
            reply_markup=main_menu(),
        )
        return

    # ------------------------------------------------
    # Handle teks biasa
    # ------------------------------------------------
    if not mode:
        await update.message.reply_text("Pilih menu terlebih dahulu.", reply_markup=main_menu())
        return

    # Batch URL dari teks langsung
    if actual_mode == "batch":
        lines = [l.strip() for l in user_input.splitlines() if l.strip()]
        target = DEFAULT_TARGET_EMAILS
        urls = []
        for line in lines:
            if line.upper().startswith("TARGET:"):
                try:
                    target = int(line.split(":", 1)[1].strip())
                except ValueError:
                    pass
            elif is_valid_url(normalize_url(line)):
                urls.append(line)

        if not urls:
            await update.message.reply_text(
                "Tidak ada URL valid. Kirim daftar URL (1 per baris) atau file .txt."
            )
            return

        progress_message = await update.message.reply_text(
            f"📦 Memulai batch scraping {len(urls)} URL... Target: {target} email"
        )
        output_filename = get_selected_output_filename(context, "web")
        rows = await scrape_batch(urls, progress_message, target_emails=target, output_filename=output_filename)

        if not rows:
            await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
            context.user_data.clear()
            return

        filepath, added_rows, total_rows = append_to_master_excel(rows, "web", output_filename)
        await update.message.reply_document(
            document=open(filepath, "rb"),
            filename=os.path.basename(filepath),
            caption=f"✅ Batch selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
        )
        await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        context.user_data.clear()
        return

    # Web scraping
    if actual_mode == "web":
        main_input, target = parse_target_from_input(user_input)
        url = normalize_url(main_input)

        if not is_valid_url(url):
            await update.message.reply_text(
                "URL tidak valid. Contoh: https://example.com\nAtau: https://example.com | 300",
                reply_markup=main_menu(),
            )
            return

        progress_message = await update.message.reply_text(
            f"🌐 Memulai scraping web... Target: {target} email"
        )
        try:
            output_filename = get_selected_output_filename(context, "web")
            rows = await scrape_website(url, progress_message, target_emails=target, output_filename=output_filename)

            if not rows:
                await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added_rows, total_rows = append_to_master_excel(rows, "web", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Web selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        except Exception as e:
            logging.exception(e)
            await update.message.reply_text(f"Error: {e}", reply_markup=main_menu())
        context.user_data.clear()

    # Maps scraping
    elif actual_mode == "maps":
        if not SERPAPI_KEY:
            await update.message.reply_text("SERPAPI_KEY belum diatur di .env.", reply_markup=main_menu())
            context.user_data.clear()
            return

        main_input, target = parse_target_from_input(user_input)
        progress_message = await update.message.reply_text(
            f"🗺️ Memulai scraping Google Maps... Target: {target} email"
        )
        try:
            output_filename = get_selected_output_filename(context, "maps")
            rows = await scrape_maps(main_input, progress_message, target_emails=target, output_filename=output_filename)

            if not rows:
                await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added_rows, total_rows = append_to_master_excel(rows, "maps", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Maps selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        except Exception as e:
            logging.exception(e)
            await update.message.reply_text(f"Error: {e}", reply_markup=main_menu())
        context.user_data.clear()

    # Dork scraping
    elif actual_mode == "dork":
        if not SERPAPI_KEY:
            await update.message.reply_text("SERPAPI_KEY belum diatur di .env.", reply_markup=main_menu())
            context.user_data.clear()
            return

        main_input, target = parse_target_from_input(user_input)
        progress_message = await update.message.reply_text(
            f"🔍 Memulai Google Dorking... Target: {target} email"
        )
        try:
            output_filename = get_selected_output_filename(context, "dork")
            rows = await scrape_dork(main_input, progress_message, target_emails=target, output_filename=output_filename)

            if not rows:
                await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added_rows, total_rows = append_to_master_excel(rows, "dork", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Dorking selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        except Exception as e:
            logging.exception(e)
            await update.message.reply_text(f"Error: {e}", reply_markup=main_menu())
        context.user_data.clear()

    # Directory scraping
    elif actual_mode == "directory":
        main_input, target = parse_target_from_input(user_input)
        progress_message = await update.message.reply_text(
            f"📋 Memulai scraping direktori bisnis... Target: {target} email"
        )
        try:
            output_filename = get_selected_output_filename(context, "directory")
            rows = await scrape_directory(main_input, progress_message, target_emails=target, output_filename=output_filename)

            if not rows:
                await progress_message.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added_rows, total_rows = append_to_master_excel(rows, "directory", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Direktori selesai.\nEmail baru: {added_rows} | Total: {total_rows}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        except Exception as e:
            logging.exception(e)
            await update.message.reply_text(f"Error: {e}", reply_markup=main_menu())
        context.user_data.clear()


# =====================================================
# MAIN APP
# =====================================================

def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN belum diatur di file .env")

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(
        (filters.TEXT | filters.Document.ALL) & ~filters.COMMAND,
        message_handler
    ))

    print("Bot berjalan...")
    print(f"Playwright: {'✅ Aktif' if PLAYWRIGHT_AVAILABLE else '⚠️ Tidak tersedia'}")
    print(f"PyMuPDF: {'✅ Aktif' if PYMUPDF_AVAILABLE else '⚠️ Tidak tersedia'}")
    print(f"MX checker: {'✅ Aktif' if DNS_AVAILABLE else '⚠️ Tidak tersedia'}")
    print(f"Verifalia: {get_verifalia_status_text()}")
    app.run_polling()


if __name__ == "__main__":
    main()