import os
import re
import time
import asyncio
import logging
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

try:
    import dns.resolver
    DNS_AVAILABLE = True
except ImportError:
    DNS_AVAILABLE = False

try:
    from playwright.async_api import async_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Document
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, ContextTypes, filters,
)

# =====================================================
# LOAD ENV
# =====================================================

load_dotenv()

TELEGRAM_BOT_TOKEN   = os.getenv("TELEGRAM_BOT_TOKEN")
SERPAPI_KEY          = os.getenv("SERPAPI_KEY")
VERIFALIA_USERNAME   = os.getenv("VERIFALIA_USERNAME")
VERIFALIA_PASSWORD   = os.getenv("VERIFALIA_PASSWORD")

# =====================================================
# SETTINGS
# =====================================================

WEB_MAX_PAGES           = 120
WEB_MAX_DEPTH           = 3
MAPS_MAX_BUSINESSES     = 50
MAPS_WEBSITE_MAX_PAGES  = 8
DORK_MAX_RESULTS        = 10
DORK_MAX_PAGES_PER_URL  = 5
DIRECTORY_MAX_PAGES     = 30
DIRECTORY_DETAIL_MAX    = 5
REQUEST_TIMEOUT         = 15
REQUEST_DELAY           = 0.8
FETCH_RETRIES           = 2
DEFAULT_TARGET_EMAILS   = 150
PARTIAL_SAVE_EVERY      = 50
PROGRESS_UPDATE_EVERY   = 3

# =====================================================
# EMAIL CHECKER SETTINGS
# =====================================================

EMAIL_CHECKER_ENABLED   = True
VERIFALIA_QUALITY       = "Standard"
VERIFALIA_ACCEPT_RISKY  = True
VERIFALIA_LIMIT_REACHED = False
VERIFALIA_CACHE: dict   = {}
MX_CACHE: dict          = {}

# =====================================================
# PRIORITY LINK KEYWORDS
# =====================================================

PRIORITY_LINK_KEYWORDS = (
    "contact", "kontak", "hubungi", "about", "tentang", "team", "staff",
    "profile", "profil", "company", "perusahaan", "support", "customer-service",
    "cs", "marketing", "sales", "career", "karir", "privacy", "legal",
)

# =====================================================
# DIRECTORY CONFIG
# =====================================================

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

# =====================================================
# OUTPUT FILES
# =====================================================

OUTPUT_DIR           = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

WEB_MASTER_FILE       = "master_web_scraping.xlsx"
MAPS_MASTER_FILE      = "master_maps_scraping.xlsx"
DORK_MASTER_FILE      = "master_dork_scraping.xlsx"
DIRECTORY_MASTER_FILE = "master_directory_scraping.xlsx"

# =====================================================
# LOGGING
# =====================================================

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)

# =====================================================
# EMAIL DOMAIN RULES
# =====================================================

FREE_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "yahoo.co.id", "hotmail.com",
    "outlook.com", "live.com", "icloud.com", "me.com",
    "ymail.com", "rocketmail.com", "mail.com",
    "googlemail.com", "msn.com",
}

SMART_MULTI_ENTITY_SUFFIXES = (
    ".sch.id", ".ac.id", ".go.id", ".or.id", ".desa.id", ".ponpes.id",
)

# =====================================================
# ▼▼▼ GEO-FILTER INDONESIA (TAMBAHAN BARU) ▼▼▼
# =====================================================

# TLD & domain Indonesia yang diizinkan
ALLOWED_ID_TLDS = (
    ".co.id", ".go.id", ".ac.id", ".sch.id", ".or.id",
    ".net.id", ".web.id", ".biz.id", ".my.id", ".desa.id",
    ".ponpes.id", ".id",
)

# Domain global umum yang tetap diizinkan (perusahaan Indonesia sering pakai .com/.net/.org)
ALLOWED_GLOBAL_TLDS = (
    ".com", ".net", ".org", ".io", ".co", ".info", ".biz",
)

# Domain country-code asing yang DIBLOKIR total
# Semua TLD negara lain di luar Indonesia akan otomatis terblokir
# karena hanya ALLOWED_ID_TLDS + ALLOWED_GLOBAL_TLDS yang lolos
BLOCKED_FOREIGN_CC_TLDS = (
    # Asia Timur & Tenggara
    ".jp", ".co.jp", ".go.jp", ".ac.jp", ".ne.jp", ".or.jp",
    ".kr", ".co.kr", ".go.kr", ".ac.kr",
    ".cn", ".com.cn", ".net.cn",
    ".tw", ".com.tw",
    ".hk", ".com.hk",
    ".sg", ".com.sg",
    ".my", ".com.my",
    ".th", ".co.th", ".go.th",
    ".vn", ".com.vn",
    ".ph", ".com.ph",
    ".mm", ".com.mm",
    ".kh",
    # Asia Selatan & Tengah
    ".in", ".co.in",
    ".pk", ".com.pk",
    ".bd", ".com.bd",
    ".lk",
    # Timur Tengah
    ".ir",           # Iran (terlihat di screenshot: n.saatchi@cra.ir)
    ".sa", ".com.sa",
    ".ae", ".com.ae",
    ".tr", ".com.tr",
    ".il",
    # Eropa
    ".uk", ".co.uk",
    ".de",
    ".fr",
    ".it",
    ".es",
    ".nl",
    ".ru",
    ".pl",
    ".se",
    ".no",
    ".fi",
    ".dk",
    ".be",
    ".at",
    ".ch",
    ".cz",
    ".pt",
    ".gr",
    ".hu",
    ".ro",
    # Amerika & Oseania
    ".us",
    ".ca",
    ".au", ".com.au",
    ".nz",
    ".br", ".com.br",
    ".mx", ".com.mx",
    ".ar", ".com.ar",
    # Afrika
    ".za", ".co.za",
    ".ng", ".com.ng",
    ".ke",
    ".eg",
    # Domain negara lain yang umum dipakai
    ".eu",
    ".gov",   # domain pemerintah AS (bukan .go.id)
    ".mil",
    ".edu",   # domain universitas AS (bukan .ac.id)
)

# Pattern domain sekolah Indonesia yang dimiliki individu (noise, bukan B2B korporat)
# Format: prefix sekolah diikuti tanda - atau .
BLOCKED_SCHOOL_DOMAIN_PATTERNS = (
    re.compile(r"^sma[\-\.]"),    # sma-xxx.co.id
    re.compile(r"^smk[\-\.]"),    # smk-xxx.co.id
    re.compile(r"^sman[\-\.]"),   # sman1-xxx.sch.id
    re.compile(r"^smkn[\-\.]"),   # smkn2-xxx.sch.id
    re.compile(r"^sd[\-\.]"),     # sd-xxx.sch.id
    re.compile(r"^sdn[\-\.]"),    # sdn1-xxx.sch.id
    re.compile(r"^smp[\-\.]"),    # smp-xxx.sch.id
    re.compile(r"^smpn[\-\.]"),   # smpn3-xxx.sch.id
    re.compile(r"^mts[\-\.]"),    # mts-xxx.sch.id
    re.compile(r"^man[\-\.]"),    # man1-xxx.sch.id
    re.compile(r"^mi[\-\.]"),     # mi-xxx.sch.id
)


def is_indonesia_or_global_domain(domain: str) -> bool:
    """
    Mengembalikan True jika domain termasuk:
    1. Domain TLD Indonesia (.co.id, .go.id, .ac.id, .id, dst.)
    2. Domain global umum yang BISA dipakai perusahaan Indonesia (.com, .net, .org, .io, dst.)
    3. Free email domain (gmail, yahoo, dst.) — tetap diizinkan untuk outreach personal

    Mengembalikan False (diblokir) jika:
    - Domain berakhiran TLD negara asing (.jp, .kr, .au, .ir, dst.)
    - Domain sekolah dasar/menengah Indonesia (pattern sma-, smk-, sd-, dst.)
    """
    domain = (domain or "").strip().lower()
    if not domain:
        return False

    # Free email domain selalu lolos
    if domain in FREE_EMAIL_DOMAINS:
        return True

    # Cek TLD Indonesia (prioritas tinggi)
    if any(domain.endswith(tld) for tld in ALLOWED_ID_TLDS):
        # Tapi blokir pattern sekolah dasar/menengah (sma-, smk-, dst.)
        domain_base = domain.split(".")[0]  # ambil bagian pertama sebelum titik
        for pat in BLOCKED_SCHOOL_DOMAIN_PATTERNS:
            if pat.match(domain_base):
                return False
        return True

    # Cek domain global umum
    if any(domain.endswith(tld) for tld in ALLOWED_GLOBAL_TLDS):
        return True

    # Semua TLD asing — blokir
    return False

# =====================================================
# ▲▲▲ END GEO-FILTER ▲▲▲
# =====================================================

# =====================================================
# EMAIL LABEL
# =====================================================

def get_email_label(email: str) -> str:
    domain = get_email_domain(email)
    if not domain:
        return "Valid"
    if domain in ("gmail.com", "googlemail.com"):
        return "Valid-Gmail"
    if domain in ("yahoo.com", "yahoo.co.id", "ymail.com", "rocketmail.com"):
        return "Valid-Yahoo"
    if domain in ("hotmail.com", "outlook.com", "live.com", "msn.com"):
        return "Valid-Microsoft"
    if domain in ("icloud.com", "me.com"):
        return "Valid-Apple"
    if domain in FREE_EMAIL_DOMAINS:
        return "Valid-FreeEmail"
    if domain.endswith(".go.id"):
        return "Valid-ID-Gov"
    if domain.endswith((".ac.id", ".sch.id")):
        return "Valid-ID-Edu"
    if domain.endswith(".or.id"):
        return "Valid-ID-Org"
    if any(domain.endswith(tld) for tld in (".co.id", ".net.id", ".web.id", ".biz.id", ".my.id")):
        return "Valid-ID-Corporate"
    if domain.endswith(".id"):
        return "Valid-ID"
    return "Valid-Corporate"

# =====================================================
# EMAIL VALIDATION RULES (B2B HIGH-PERFORMANCE)
# =====================================================

EMAIL_REGEX = re.compile(
    r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
)

HEX_HASH_REGEX    = re.compile(r"^[a-f0-9]{16,}$", re.IGNORECASE)
LONG_RANDOM_REGEX = re.compile(r"^[a-z0-9]{24,}$", re.IGNORECASE)

GOOD_BUSINESS_LOCAL_PARTS = {
    "admin", "info", "contact", "kontak", "sales", "marketing",
    "cs", "hr", "humas", "office", "hello", "halo", "partnership",
    "kerjasama", "b2b", "procurement", "purchasing", "vendor", "inquiry", "enquiry"
}

BLOCKED_EMAILS = {
    "example@example.com", "test@test.com", "admin@example.com",
    "contoh@email.com", "email@example.com", "your@email.com",
    "you@example.com", "halo@kontak.com",
    "help.lordsmobile.android@igg.com",
    "contact@lemon8-app.com",
    "service@pubgmobile.com",
    "wwwwa.l.r.u.scv.kd@lulle.sakura.ne.jp",
    "sekretariat@rektor.unair.ac.id",
    "gossipharbor-service@microfun.com",
    "support.happycolor@x-flow.app",
    "hrd_recruitment@arnotts.com",
    "recruitment.group@mensa-group.com",
    "recruitment_indonesia@goodyear.com",
    "info@eeoc.gov", "ivona@juicebox.com.au",
    "ier@usdoj.gov", "hqaffirmativeaction@marriott.com",
    "inquiry-lgq1@logique.co.id",
}

BLOCKED_LOCAL_PARTS = {
    "privacy", "policy", "legal", "terms", "gdpr", "dpo",
    "noreply", "no-reply", "no_reply", "donotreply", "do-not-reply",
    "mailer-daemon", "postmaster", "abuse", "root", "hostmaster",
    "webmaster", "example", "test", "testing", "email", "your", "you", "rektor",
}

BLOCKED_LOCAL_KEYWORDS = (
    "noreply", "no-reply", "donotreply", "do-not-reply",
    "sentry", "notification", "automated", "bounce", "mailer-daemon",
    "android", "ios", "apk", "appstore", "playstore", "mobilegame",
)

BLOCKED_EMAIL_DOMAINS = {
    "linktr.ee", "sentry.wixpress.com", "wixpress.com", "wix.com",
    "wixsite.com", "sentry.io", "example.com", "test.com",
    "support.whatsapp.com", "igg.com", "lemon8-app.com",
    "pubgmobile.com", "lulle.sakura.ne.jp", "rektor.unair.ac.id",
    "microfun.com", "x-flow.app", "kemen",
}

BLOCKED_EMAIL_DOMAIN_SUFFIXES = (
    ".wixpress.com", ".wixsite.com", ".sentry.io",
    ".go.id",  # Blokir semua email instansi pemerintah/kementerian
)

BLOCKED_EXTENSIONS = (
    ".png", ".jpg", ".jpeg", ".webp", ".svg", ".gif",
    ".bmp", ".ico", ".pdf", ".css", ".js",
)

DISPOSABLE_EMAIL_DOMAINS = {
    "mailinator.com", "tempmail.com", "temp-mail.org", "10minutemail.com",
    "guerrillamail.com", "guerrillamail.net", "yopmail.com", "trashmail.com",
    "getnada.com", "fakeinbox.com", "emailondeck.com", "sharklasers.com",
    "dispostable.com", "maildrop.cc", "moakt.com", "mintemail.com",
    "throwawaymail.com", "tempail.com", "tempmailo.com",
}

BAD_DOMAIN_KEYWORDS = (
    "sentry", "wixpress", "example", "test", "dummy", "placeholder",
    "localhost", "invalid", "fake", "tempmail", "mailinator",
)

STRICT_BLOCKED_DOMAIN_KEYWORDS = (
    "pubg", "pubgmobile", "mobile", "game", "games",
    "lemon8", "microfun", "x-flow", "happycolor", "gossipharbor",
    "sakura.ne.jp", "rektor", "unair.ac.id",
    "mailgun", "sendgrid", "amazonses", "sparkpost", "mandrillapp",
    "mailchimp", "brevo", "sendinblue", "constantcontact",
)

ROLE_EMAIL_LOCAL_PARTS = {
    "privacy", "legal", "terms", "gdpr", "dpo", "abuse",
    "postmaster", "mailer-daemon", "root", "hostmaster", "webmaster",
    "sekretariat", "rektor", "rektorat", "akademik", "kemahasiswaan",
    "student", "alumni", "prakerin", "magang", "bkk", "perpus", "perpustakaan",
    "recruitment", "recruiting", "career", "careers", "jobs", "job",
    "vacancy", "vacancies", "hrd", "talent", "headhunter", "hiring", "recruit",
    "complaint", "complaints", "pengaduan", "helpdesk", "billing", "invoice",
    "finance", "accounting", "bendahara", "tatausaha", "tu"
}

STRICT_BLOCKED_LOCAL_PARTS = {
    "go", "hi", "hey", "mail", "email", "user", "users",
    "contactus", "help", "support", "service", "services",
    "customer", "customers", "owner", "demo", "null", "none",
}

STRICT_BLOCKED_LOCAL_PREFIXES = (
    "help.", "support.", "privacy.", "legal.", "terms.",
    "noreply.", "no-reply.", "notification.", "admin.",
)

STRICT_APP_NOISE_KEYWORDS = (
    "android", "ios", "apk", "appstore", "playstore",
    "mobile", "mobilegame", "game", "games",
    "pubg", "pubgmobile", "lordsmobile",
    "happycolor", "gossipharbor", "lemon8",
    "facebook", "instagram", "tiktok", "telegram",
    "discord", "whatsapp", "line", "snapchat",
    "notification", "notify", "noreply", "no-reply",
    "mailer", "daemon", "system", "automated", "bot", "bounce",
    "shopee", "tokopedia", "lazada", "bukalapak",
    "grab", "gojek", "traveloka", "booking", "airbnb",
    "amazon", "paypal", "stripe",
    "microfun", "x-flow", "sakura",
)

BAD_LOCAL_PATTERNS = (
    "example", "dummy", "sample", "testing", "testmail",
    "yourname", "username", "firstname", "lastname", "namehere",
)

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

    # --- Tambahan: TLD Indonesia spesifik ---
    'site:.co.id "{keyword}" "{location}" "email"',
    'site:.id "{keyword}" "{location}" "hubungi kami"',
    'site:.or.id "{keyword}" "email" "kontak"',
    'site:.web.id "{keyword}" "{location}" email',

    # --- Tambahan: Halaman kontak & profil Indonesia ---
    'inurl:kontak "{keyword}" "{location}" "@"',
    'inurl:hubungi "{keyword}" "{location}" email',
    'inurl:profil "{keyword}" "{location}" email',
    'inurl:tentang-kami "{keyword}" email "@"',
    'inurl:about-us "{keyword}" "{location}" "@"',

    # --- Tambahan: Direktori & listing bisnis Indonesia ---
    'site:indonetwork.co.id "{keyword}" "{location}"',
    'site:tokomesin.com "{keyword}" email',
    'site:ralali.com "{keyword}" "{location}" email',
    'site:indotrading.com "{keyword}" "{location}"',
    'site:b2b.id "{keyword}" "{location}"',

    # --- Tambahan: Format email langsung di SERP ---
    '"{keyword}" "{location}" intext:"@{location}.co.id"',
    '"{keyword}" "{location}" intext:"info@" OR intext:"contact@" OR intext:"sales@"',
    '"{keyword}" "{location}" intext:"cs@" OR intext:"marketing@" OR intext:"admin@"',
    '"{keyword}" "{location}" intext:"humas@" OR intext:"office@" OR intext:"kontak@"',

    # --- Tambahan: Filetype dokumen bisnis ---
    'filetype:pdf site:.co.id "{keyword}" "email"',
    'filetype:doc "{keyword}" "{location}" "email" "kontak"',
    'filetype:xls "{keyword}" "{location}" email',

    # --- Tambahan: Media & berita bisnis Indonesia ---
    'site:bisnis.com "{keyword}" "{location}" email',
    'site:kontan.co.id "{keyword}" email',
    'site:detik.com "{keyword}" "{location}" email kontak',

    # --- Tambahan: Kota besar Indonesia lainnya ---
    '"{keyword}" "Medan" OR "Makassar" OR "Semarang" email kontak',
    '"{keyword}" "Yogyakarta" OR "Palembang" OR "Pekanbaru" email',
    '"{keyword}" "Balikpapan" OR "Denpasar" OR "Manado" email',
    '"{keyword}" "Samarinda" OR "Batam" OR "Banjarmasin" email kontak',

    # --- Tambahan: Kombinasi profil perusahaan ---
    '"{keyword}" "{location}" "profil perusahaan" email',
    '"{keyword}" "{location}" "company profile" email "@"',
    '"{keyword}" "{location}" "visi misi" email kontak',
    '"{keyword}" "{location}" filetype:pdf "company profile" email',
]

# =====================================================
# EMAIL UTILS
# =====================================================

def clean_email(email: str) -> str:
    # Tambahan: strip trailing punctuation yang sering menempel di HTML
    email = email.strip().lower().replace("mailto:", "")
    email = email.rstrip(".,;:!?\"'")
    return email


def get_email_domain(email: str) -> str:
    email = clean_email(email)
    if "@" not in email:
        return ""
    return email.split("@", 1)[1].strip().lower()


def validate_email(email: str) -> str:
    email = clean_email(email)

    if "@" not in email or email.startswith("@"):
        return "Invalid"
    if not EMAIL_REGEX.fullmatch(email):
        return "Invalid"
    if email in BLOCKED_EMAILS:
        return "Invalid"
    if any(email.endswith(ext) for ext in BLOCKED_EXTENSIONS):
        return "Invalid"

    local_part, domain = email.rsplit("@", 1)
    local_part = local_part.strip().lower()
    domain     = domain.strip().lower()

    # --- Validasi Domain ---
    if "." not in domain or domain.startswith(".") or domain.endswith("."):
        return "Invalid"
    if ".." in domain or "--" in domain or "-." in domain or ".-" in domain:
        return "Invalid"

    parts = domain.split(".")
    if parts[-2].endswith("-") or parts[0].startswith("-"):
        return "Invalid"

    tld = parts[-1]
    if len(tld) < 2 or len(tld) > 10:
        return "Invalid"

    if domain in BLOCKED_EMAIL_DOMAINS:
        return "Invalid"
    if any(domain.endswith(s) for s in BLOCKED_EMAIL_DOMAIN_SUFFIXES):
        return "Invalid"
    if domain in DISPOSABLE_EMAIL_DOMAINS:
        return "Invalid"
    if any(kw in domain for kw in BAD_DOMAIN_KEYWORDS):
        return "Invalid"
    if any(kw in domain for kw in STRICT_BLOCKED_DOMAIN_KEYWORDS):
        return "Invalid"

    # -------------------------------------------------------
    # ▼▼▼ GEO-FILTER INDONESIA — CUKUP SATU BARIS INI ▼▼▼
    # Blokir semua domain negara asing (.jp, .kr, .au, .ir, dst.)
    # Izinkan: TLD Indonesia + domain global umum + free email
    # -------------------------------------------------------
    if not is_indonesia_or_global_domain(domain):
        return "Invalid"
    # ▲▲▲ END GEO-FILTER ▲▲▲

    # --- Validasi Local Part (Sistem Jalur Ganda B2B) ---
    if not local_part or len(local_part) < 2 or len(local_part) > 40:
        return "Invalid"
    if local_part.startswith(".") or local_part.endswith(".") or ".." in local_part:
        return "Invalid"

    normalized = local_part.replace(".", "").replace("_", "").replace("-", "")

    # Jalur Fast-Track B2B
    is_priority_b2b = local_part in GOOD_BUSINESS_LOCAL_PARTS

    if is_priority_b2b:
        pass
    else:
        if local_part in BLOCKED_LOCAL_PARTS or normalized in BLOCKED_LOCAL_PARTS:
            return "Invalid"
        if local_part in ROLE_EMAIL_LOCAL_PARTS or normalized in ROLE_EMAIL_LOCAL_PARTS:
            return "Invalid"
        if local_part in STRICT_BLOCKED_LOCAL_PARTS or normalized in STRICT_BLOCKED_LOCAL_PARTS:
            return "Invalid"
        if any(kw in local_part for kw in BLOCKED_LOCAL_KEYWORDS):
            return "Invalid"
        if any(kw in local_part for kw in STRICT_APP_NOISE_KEYWORDS):
            return "Invalid"
        if any(pat in local_part for pat in BAD_LOCAL_PATTERNS):
            return "Invalid"
        if any(local_part.startswith(pfx) for pfx in STRICT_BLOCKED_LOCAL_PREFIXES):
            return "Invalid"

        if HEX_HASH_REGEX.fullmatch(normalized):
            return "Invalid"
        if LONG_RANDOM_REGEX.fullmatch(normalized) and not any(c in local_part for c in [".", "-", "_"]):
            return "Invalid"
        if normalized.isdigit():
            return "Invalid"

        if len(normalized) >= 4 and len(set(normalized[:4])) == 1:
            return "Invalid"

        digit_count = sum(c.isdigit() for c in normalized)
        if len(normalized) >= 8 and (digit_count / len(normalized)) > 0.35:
            return "Invalid"

    return "Valid"


def extract_emails_from_text(text: str) -> list[tuple[str, str]]:
    results, seen = [], set()
    for email in EMAIL_REGEX.findall(text or ""):
        email = clean_email(email)
        if email in seen or email.startswith("@"):
            continue
        seen.add(email)
        local = email.split("@")[0]
        if not any(c.isalpha() for c in local):
            continue
        results.append((email, validate_email(email)))
    return results

# =====================================================
# DOMAIN DEDUP
# =====================================================

def is_free_email_domain(domain: str) -> bool:
    return (domain or "").strip().lower() in FREE_EMAIL_DOMAINS


def should_skip_domain(email_domain: str, found_domain_set: set) -> bool:
    domain = (email_domain or "").strip().lower()
    if not domain:
        return True
    if is_free_email_domain(domain):
        return False
    return domain in found_domain_set

# =====================================================
# EMAIL CHECKER: MX + VERIFALIA
# =====================================================

def has_mx_record(domain: str) -> bool:
    domain = (domain or "").strip().lower()
    if not domain:
        return False
    if domain in MX_CACHE:
        return MX_CACHE[domain]
    if not DNS_AVAILABLE:
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
    global VERIFALIA_LIMIT_REACHED
    email = clean_email(email)

    if not EMAIL_CHECKER_ENABLED or VERIFALIA_LIMIT_REACHED:
        return "Skipped"
    if not VERIFALIA_USERNAME or not VERIFALIA_PASSWORD:
        return "Skipped"
    if email in VERIFALIA_CACHE:
        return VERIFALIA_CACHE[email]

    try:
        async with session.post(
            "https://api.verifalia.com/v2.7/email-validations",
            json={"entries": [{"inputData": email}], "quality": VERIFALIA_QUALITY},
            auth=aiohttp.BasicAuth(VERIFALIA_USERNAME, VERIFALIA_PASSWORD),
            timeout=25,
        ) as resp:
            if resp.status in (402, 429):
                VERIFALIA_LIMIT_REACHED = True
                logging.warning("Verifalia credit/limit habis. Fallback ke MX.")
                return "Skipped"
            if resp.status >= 400:
                return "Skipped"
            data    = await resp.json()
            entries = data.get("entries", {}).get("data", [])
            if not entries:
                return "Unknown"
            result = entries[0].get("classification", "Unknown")
            VERIFALIA_CACHE[email] = result
            return result
    except Exception as e:
        logging.warning(f"Verifalia error {email}: {e}")
        return "Unknown"


async def is_email_allowed_by_checker(
    session: aiohttp.ClientSession, email: str
) -> tuple[bool, str]:
    email_domain = get_email_domain(clean_email(email))
    if not has_mx_record(email_domain):
        return False, "No MX"

    status = await check_verifalia(session, email)
    if status == "Skipped":
        return True, "Local+MX"
    if status == "Deliverable":
        return True, "Deliverable"
    if VERIFALIA_ACCEPT_RISKY and status == "Risky":
        return True, "Risky"
    return False, status

# =====================================================
# URL UTILS
# =====================================================

def normalize_url(url: str) -> str:
    url = (url or "").strip().replace("[", "").replace("]", "").replace("\\", "")
    if "google.com/url" in url and "q=" in url:
        try:
            qs = parse_qs(urlparse(url).query)
            if "q" in qs and qs["q"]:
                url = qs["q"][0].strip()
        except Exception:
            pass
    if url and not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


def is_valid_url(url: str) -> bool:
    try:
        p = urlparse(normalize_url(url))
        return p.scheme in ("http", "https") and bool(p.netloc) and "[" not in p.netloc
    except Exception:
        return False


def extract_links(base_url: str, html: str) -> set[str]:
    links = set()
    for tag in BeautifulSoup(html, "html.parser").find_all("a", href=True):
        href = tag["href"].strip()
        if not href or href.startswith(("mailto:", "tel:", "#", "javascript:")):
            continue
        full = urljoin(base_url, href)
        p = urlparse(full)
        if p.scheme in ("http", "https"):
            links.add(p._replace(fragment="").geturl())
    return links


def same_domain(url1: str, url2: str) -> bool:
    return (
        urlparse(url1).netloc.lower().replace("www.", "")
        == urlparse(url2).netloc.lower().replace("www.", "")
    )


def normalize_crawl_url(url: str) -> str:
    return urlparse(url)._replace(fragment="").geturl().rstrip("/")


def link_priority(url: str) -> int:
    low = url.lower()
    if any(k in low for k in PRIORITY_LINK_KEYWORDS):
        return 0
    return 1 if low.count("/") <= 3 else 2

# =====================================================
# PROGRESS & FILE HELPERS
# =====================================================

def make_progress_bar(percent: int, size: int = 10) -> str:
    filled = max(0, min(size, round((percent / 100) * size)))
    return "█" * filled + "░" * (size - filled)


def get_columns(scraping_type: str) -> list[str]:
    return {
        "web":       ["No", "Source URL", "Email", "Status", "Found At", "Scraping Type"],
        "maps":      ["No", "Business Name", "Search Keyword", "Location", "Website URL", "Email", "Status", "Found At", "Scraping Type"],
        "dork":      ["No", "Dork Query", "Source URL", "Email", "Status", "Found At", "Scraping Type"],
        "directory": ["No", "Directory", "Business Name", "Website URL", "Email", "Status", "Found At", "Scraping Type"],
    }.get(scraping_type, ["No", "Source URL", "Email", "Status", "Found At", "Scraping Type"])


def get_master_filename(scraping_type: str) -> str:
    return {
        "web":       WEB_MASTER_FILE,
        "maps":      MAPS_MASTER_FILE,
        "dork":      DORK_MASTER_FILE,
        "directory": DIRECTORY_MASTER_FILE,
    }.get(scraping_type, WEB_MASTER_FILE)


def safe_xlsx_name(name: str, scraping_type: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_. -]", "_", os.path.basename((name or "").strip())).strip(" ._")
    if not name:
        name = f"{scraping_type}_scraping_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return name if name.lower().endswith(".xlsx") else name + ".xlsx"


def list_xlsx_files(scraping_type: str | None = None) -> list[dict]:
    files = []
    if not os.path.exists(OUTPUT_DIR):
        return files
    for filename in sorted(os.listdir(OUTPUT_DIR)):
        if not filename.lower().endswith(".xlsx"):
            continue
        path = os.path.join(OUTPUT_DIR, filename)
        try:
            size_kb  = os.path.getsize(path) / 1024
            modified = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
            try:
                rows_count = len(pd.read_excel(path))
            except Exception:
                rows_count = 0
            files.append({
                "filename": filename, "path": path,
                "size_kb": size_kb, "modified": modified, "rows": rows_count,
            })
        except Exception:
            continue
    return files


def load_existing_master(
    scraping_type: str, filename: str | None = None
) -> tuple[pd.DataFrame, set, set, str]:
    master_path = os.path.join(
        OUTPUT_DIR,
        safe_xlsx_name(filename or get_master_filename(scraping_type), scraping_type)
    )
    columns = get_columns(scraping_type)

    if not os.path.exists(master_path):
        return pd.DataFrame(columns=columns), set(), set(), master_path

    try:
        df = pd.read_excel(master_path)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        df = df[columns]
        if "Email" in df.columns:
            df["Email"] = df["Email"].fillna("").astype(str).str.strip().str.lower()
            df = df[df["Email"].apply(lambda x: validate_email(x) == "Valid")]
        existing_emails  = set(df["Email"].dropna().astype(str).str.strip().str.lower()) - {""}
        existing_domains = {get_email_domain(e) for e in existing_emails if get_email_domain(e)}
        return df, existing_emails, existing_domains, master_path
    except Exception as e:
        logging.warning(f"Gagal baca master {master_path}: {e}")
        return pd.DataFrame(columns=columns), set(), set(), master_path


def append_to_master_excel(
    new_rows: list, scraping_type: str, filename: str | None = None
) -> tuple[str, int, int]:
    columns = get_columns(scraping_type)
    old_df, existing_emails, existing_domains, master_path = load_existing_master(
        scraping_type, filename
    )

    if not new_rows:
        old_df.to_excel(master_path, index=False)
        return master_path, 0, len(old_df)

    email_idx  = {"web": 2, "maps": 5, "dork": 3, "directory": 4}.get(scraping_type, 2)
    status_idx = {"web": 3, "maps": 6, "dork": 4, "directory": 5}.get(scraping_type, 3)

    unique_rows = []
    for row in new_rows:
        row_email  = str(row[email_idx]).strip().lower()  if len(row) > email_idx  else ""
        row_status = str(row[status_idx]).strip()         if len(row) > status_idx else ""

        if (
            not row_email
            or not row_status.startswith("Valid")
            or validate_email(row_email) != "Valid"
            or row_email in existing_emails
        ):
            continue

        row_domain = get_email_domain(row_email)
        if not row_domain or should_skip_domain(row_domain, existing_domains):
            continue

        existing_emails.add(row_email)
        if not is_free_email_domain(row_domain):
            existing_domains.add(row_domain)
        unique_rows.append(row)

    combined_df = pd.concat(
        [old_df, pd.DataFrame(unique_rows, columns=columns)], ignore_index=True
    )
    combined_df["No"] = range(1, len(combined_df) + 1)
    combined_df.to_excel(master_path, index=False)
    return master_path, len(unique_rows), len(combined_df)


def save_excel(rows: list, filename: str, scraping_type: str) -> str:
    filepath = os.path.join(OUTPUT_DIR, filename)
    pd.DataFrame(rows, columns=get_columns(scraping_type)).to_excel(filepath, index=False)
    return filepath


# =====================================================
# HTTP FETCH HELPERS
# =====================================================

async def fetch_html(session: aiohttp.ClientSession, url: str) -> tuple[str | None, str | None]:
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; EmailScraperBot/1.0; +legal-public-email-scan)",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }

    last_error = None
    for attempt in range(1, FETCH_RETRIES + 2):
        try:
            async with session.get(
                url,
                headers=headers,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            ) as resp:
                if resp.status >= 400:
                    return None, f"HTTP {resp.status}"

                content_type = resp.headers.get("content-type", "").lower()
                if "text/html" not in content_type and "application/xhtml" not in content_type:
                    return None, "Not HTML"

                return await resp.text(errors="ignore"), None

        except asyncio.TimeoutError:
            last_error = "Timeout"
        except aiohttp.ClientError as e:
            last_error = str(e)
        except Exception as e:
            last_error = str(e)

        if attempt <= FETCH_RETRIES:
            await asyncio.sleep(0.5 * attempt)

    return None, last_error or "Unknown error"


async def fetch_pdf_bytes(session: aiohttp.ClientSession, url: str) -> bytes | None:
    try:
        async with session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True) as resp:
            if resp.status >= 400:
                return None

            content_type = resp.headers.get("content-type", "").lower()
            if "pdf" not in content_type and not url.lower().endswith(".pdf"):
                return None

            return await resp.read()
    except Exception:
        return None


async def fetch_html_playwright(url: str) -> tuple[str | None, str | None]:
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


def extract_emails_from_pdf_bytes(pdf_bytes: bytes) -> list[tuple[str, str]]:
    if not PYMUPDF_AVAILABLE or not pdf_bytes:
        return []

    results = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            results.extend(extract_emails_from_text(page.get_text()))
        doc.close()
    except Exception as e:
        logging.warning(f"Gagal ekstrak PDF: {e}")

    return results


# =====================================================
# TELEGRAM MENU HELPERS
# =====================================================

def main_menu() -> InlineKeyboardMarkup:
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


def output_choice_menu(scraping_type: str) -> InlineKeyboardMarkup:
    keyboard = [
        [
            InlineKeyboardButton("🆕 Buat XLSX Baru", callback_data=f"xlsx_new:{scraping_type}"),
            InlineKeyboardButton("📂 Pakai XLSX Lama", callback_data=f"xlsx_old:{scraping_type}"),
        ],
        [InlineKeyboardButton("⬅️ Kembali", callback_data="back_main")],
    ]
    return InlineKeyboardMarkup(keyboard)


def manage_xlsx_menu() -> InlineKeyboardMarkup:
    keyboard = [
        [
            InlineKeyboardButton("📄 Web", callback_data="list_xlsx:web"),
            InlineKeyboardButton("🗺️ Maps", callback_data="list_xlsx:maps"),
            InlineKeyboardButton("🔍 Dork", callback_data="list_xlsx:dork"),
        ],
        [InlineKeyboardButton("📋 Direktori", callback_data="list_xlsx:directory")],
        [InlineKeyboardButton("🗑️ Hapus File", callback_data="delete_xlsx:all")],
        [InlineKeyboardButton("⬅️ Kembali", callback_data="back_main")],
    ]
    return InlineKeyboardMarkup(keyboard)


def build_xlsx_file_menu(files: list, action: str, scraping_type: str) -> InlineKeyboardMarkup:
    keyboard = []

    for idx, fileinfo in enumerate(files[:20]):
        label = f"{idx + 1}. {fileinfo['filename']} ({fileinfo['rows']} email)"
        keyboard.append([
            InlineKeyboardButton(
                label,
                callback_data=f"{action}:{scraping_type}:{idx}",
            )
        ])

    keyboard.append([InlineKeyboardButton("⬅️ Kembali", callback_data="manage_xlsx")])
    return InlineKeyboardMarkup(keyboard)


def get_selected_output_filename(context: ContextTypes.DEFAULT_TYPE, scraping_type: str) -> str:
    return context.user_data.get(f"selected_{scraping_type}_xlsx") or get_master_filename(scraping_type)


async def safe_edit_message(message, text: str, reply_markup=None) -> None:
    try:
        await message.edit_text(text, reply_markup=reply_markup, parse_mode="Markdown")
    except Exception:
        try:
            await message.edit_text(text, reply_markup=reply_markup)
        except Exception:
            pass


# =====================================================
# SERPAPI HELPERS
# =====================================================

async def serpapi_search_maps(keyword_location: str) -> list:
    if not SERPAPI_KEY:
        raise ValueError("SERPAPI_KEY belum diatur di file .env")

    params = {
        "engine": "google_maps",
        "q": keyword_location,
        "type": "search",
        "api_key": SERPAPI_KEY,
    }

    async with aiohttp.ClientSession() as session:
        async with session.get("https://serpapi.com/search.json", params=params, timeout=REQUEST_TIMEOUT) as resp:
            if resp.status == 401:
                raise ValueError("SERPAPI_KEY tidak valid.")
            if resp.status == 429:
                raise ValueError("Limit SerpAPI sudah tercapai.")
            data = await resp.json()

    if "error" in data:
        raise ValueError(data["error"])

    return data.get("local_results", [])[:MAPS_MAX_BUSINESSES]


async def serpapi_search_web(query: str) -> list[str]:
    if not SERPAPI_KEY:
        raise ValueError("SERPAPI_KEY belum diatur di file .env")

    params = {
        "engine": "google",
        "q": query,
        "num": DORK_MAX_RESULTS,
        "api_key": SERPAPI_KEY,
    }

    try:
        async with aiohttp.ClientSession() as session:
            async with session.get("https://serpapi.com/search.json", params=params, timeout=REQUEST_TIMEOUT) as resp:
                if resp.status != 200:
                    return []
                data = await resp.json()

        return [
            result.get("link")
            for result in data.get("organic_results", [])
            if result.get("link")
        ]
    except Exception as e:
        logging.warning(f"SerpAPI web search error: {e}")
        return []


def split_keyword_location(text: str) -> tuple[str, str]:
    parts = text.strip().split()
    if len(parts) <= 1:
        return text.strip(), ""

    keyword = " ".join(parts[:-1])
    location = parts[-1]
    return keyword, location


def parse_target_from_input(text: str) -> tuple[str, int]:
    if "|" not in text:
        return text.strip(), DEFAULT_TARGET_EMAILS

    main_input, raw_target = text.rsplit("|", 1)
    try:
        target = int(raw_target.strip())
        target = max(1, min(target, 10000))
    except ValueError:
        target = DEFAULT_TARGET_EMAILS

    return main_input.strip(), target


# =====================================================
# CORE: EMAIL PROCESSING HELPER
# =====================================================

async def process_email(
    session: aiohttp.ClientSession,
    email: str,
    status: str,
    source_url: str,
    found_email_set: set,
    found_domain_set: set,
    counters: dict,
) -> bool:
    domain = get_email_domain(email)

    if status != "Valid":
        counters["invalid"] += 1
        return False
    if email in found_email_set:
        counters["dup_email"] += 1
        return False
    if should_skip_domain(domain, found_domain_set):
        counters["dup_domain"] += 1
        return False

    allowed, _ = await is_email_allowed_by_checker(session, email)
    if not allowed:
        counters["invalid"] += 1
        return False

    found_email_set.add(email)
    if not is_free_email_domain(domain):
        found_domain_set.add(domain)
    return True

# =====================================================
# CORE: SINGLE URL SCANNER
# =====================================================

async def scan_single_url(
    session: aiohttp.ClientSession,
    url: str,
    max_pages: int,
    start_url: str | None = None,
    use_playwright_fallback: bool = False,
) -> tuple[list, int]:
    root           = start_url or url
    queue          = [(normalize_crawl_url(url), 0)]
    queued         = {normalize_crawl_url(url)}
    visited        = set()
    found          = []
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
        if not html and use_playwright_fallback and PLAYWRIGHT_AVAILABLE:
            html, error = await fetch_html_playwright(current_url)

        if not html:
            if current_url.lower().endswith(".pdf") and PYMUPDF_AVAILABLE:
                pdf_b = await fetch_pdf_bytes(session, current_url)
                if pdf_b:
                    for e, s in extract_emails_from_pdf_bytes(pdf_b):
                        found.append((e, s, current_url))
            continue

        for e, s in extract_emails_from_text(html):
            found.append((e, s, current_url))

        if PYMUPDF_AVAILABLE:
            for a_tag in BeautifulSoup(html, "html.parser").find_all("a", href=True):
                href = a_tag["href"]
                if href.lower().endswith(".pdf"):
                    pdf_url = urljoin(current_url, href)
                    if pdf_url not in visited:
                        pdf_b = await fetch_pdf_bytes(session, pdf_url)
                        if pdf_b:
                            for e, s in extract_emails_from_pdf_bytes(pdf_b):
                                found.append((e, s, pdf_url))
                        visited.add(pdf_url)

        if depth < WEB_MAX_DEPTH:
            for link in sorted(extract_links(current_url, html), key=link_priority):
                c_link = normalize_crawl_url(link)
                if same_domain(root, c_link) and c_link not in visited and c_link not in queued:
                    queue.append((c_link, depth + 1))
                    queued.add(c_link)

    return found, pages_scanned

# =====================================================
# CORE: WEB SCRAPER
# =====================================================

async def scrape_website(
    start_url: str,
    progress_message,
    target_emails: int = DEFAULT_TARGET_EMAILS,
    output_filename: str | None = None,
) -> list:
    _, existing_emails, existing_domains, _ = load_existing_master("web", output_filename)
    found_email_set  = set(existing_emails)
    found_domain_set = set(existing_domains)
    session_start    = len(found_email_set)
    rows             = []
    counters         = {"invalid": 0, "dup_email": 0, "dup_domain": 0}

    queued_urls    = [(normalize_crawl_url(start_url), 0)]
    queued_url_set = {normalize_crawl_url(start_url)}
    visited_urls   = set()
    scanned_pages  = 0
    failed_pages   = 0
    latest_preview = []
    start_time     = time.time()
    stop_reason    = "Selesai"

    async with aiohttp.ClientSession() as session:
        while queued_urls:
            queued_urls.sort(key=lambda item: (item[1], link_priority(item[0])))
            current_url, depth = queued_urls.pop(0)
            queued_url_set.discard(current_url)

            if current_url in visited_urls:
                continue
            if scanned_pages >= WEB_MAX_PAGES:
                stop_reason = "Batas halaman tercapai"
                break
            new_count = len(found_email_set) - session_start
            if new_count >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            visited_urls.add(current_url)
            scanned_pages += 1

            html, error = await fetch_html(session, current_url)
            if not html and PLAYWRIGHT_AVAILABLE:
                html, error = await fetch_html_playwright(current_url)

            if html:
                for email, status in extract_emails_from_text(html):
                    accepted = await process_email(
                        session, email, status, current_url,
                        found_email_set, found_domain_set, counters
                    )
                    if accepted:
                        lbl = get_email_label(email)
                        rows.append([
                            len(rows) + 1, current_url, email, lbl,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Web Scraping",
                        ])
                        latest_preview.insert(0, f"✅ {email} [{lbl}]")
                        latest_preview = latest_preview[:10]

                        if len(rows) % PARTIAL_SAVE_EVERY == 0:
                            save_excel(rows, f"partial_web_{int(time.time())}.xlsx", "web")

                if depth < WEB_MAX_DEPTH:
                    for link in sorted(extract_links(current_url, html), key=link_priority):
                        c_link = normalize_crawl_url(link)
                        if (
                            same_domain(start_url, c_link)
                            and c_link not in visited_urls
                            and c_link not in queued_url_set
                        ):
                            queued_urls.append((c_link, depth + 1))
                            queued_url_set.add(c_link)
            else:
                failed_pages += 1

            new_count = len(found_email_set) - session_start
            pct = min(max(
                int((scanned_pages / WEB_MAX_PAGES) * 100),
                int((new_count / max(target_emails, 1)) * 100),
            ), 100)

            if scanned_pages == 1 or scanned_pages % PROGRESS_UPDATE_EVERY == 0 or not queued_urls:
                await safe_edit_message(progress_message,
                    f"🌐 *Progress Scraping Web*\n\n"
                    f"{make_progress_bar(pct)} {pct}%\n\n"
                    f"📄 Halaman: {current_url[:60]}\n"
                    f"📊 Discan: {scanned_pages}/{WEB_MAX_PAGES} | Kedalaman: {depth}/{WEB_MAX_DEPTH}\n"
                    f"✅ Valid: {len(rows)} | ❌ Invalid: {counters['invalid']}\n"
                    f"♻️ Dup Email: {counters['dup_email']} | 🏷️ Dup Domain: {counters['dup_domain']}\n"
                    f"📧 Baru: {new_count}/{target_emails}\n"
                    f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                    f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
                )
            await asyncio.sleep(REQUEST_DELAY)

    await safe_edit_message(
        progress_message,
        f"✅ *Scraping Web Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Halaman: {scanned_pages}\n"
        f"Email Valid: {len(rows)}\n"
        f"Baru: {len(found_email_set) - session_start}"
    )
    return rows

# =====================================================
# CORE: GOOGLE MAPS SCRAPER
# =====================================================

async def scrape_maps(
    keyword_location: str,
    progress_message,
    target_emails: int = DEFAULT_TARGET_EMAILS,
    output_filename: str | None = None,
) -> list:
    listings = await serpapi_search_maps(keyword_location)
    keyword, location = split_keyword_location(keyword_location)

    _, existing_emails, existing_domains, _ = load_existing_master("maps", output_filename)
    found_email_set  = set(existing_emails)
    found_domain_set = set(existing_domains)
    session_start    = len(found_email_set)
    rows             = []
    counters         = {"invalid": 0, "dup_email": 0, "dup_domain": 0}

    total_checked    = 0
    websites_scanned = 0
    latest_preview   = []
    start_time       = time.time()
    stop_reason      = "Selesai"

    async with aiohttp.ClientSession() as session:
        for business in listings:
            new_count = len(found_email_set) - session_start
            if new_count >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            total_checked += 1
            b_name   = business.get("title", "Unknown Business")
            website  = business.get("website")
            b_addr   = business.get("address", location)

            if website:
                websites_scanned += 1
                extracted, _ = await scan_single_url(
                    session, normalize_url(website), MAPS_WEBSITE_MAX_PAGES,
                    start_url=normalize_url(website)
                )
                for email, status, src_url in extracted:
                    accepted = await process_email(
                        session, email, status, src_url,
                        found_email_set, found_domain_set, counters
                    )
                    if accepted:
                        lbl = get_email_label(email)
                        rows.append([
                            len(rows) + 1, b_name, keyword, b_addr, src_url,
                            email, lbl,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Google Maps Scraping",
                        ])
                        latest_preview.insert(0, f"✅ {email} [{lbl}] | {b_name}")
                        latest_preview = latest_preview[:10]

                        if len(rows) % PARTIAL_SAVE_EVERY == 0:
                            save_excel(rows, f"partial_maps_{int(time.time())}.xlsx", "maps")

                        if (len(found_email_set) - session_start) >= target_emails:
                            stop_reason = "Target email baru tercapai"
                            break

            new_count = len(found_email_set) - session_start
            pct = min(max(
                int((total_checked / max(len(listings), 1)) * 100),
                int((new_count / max(target_emails, 1)) * 100),
            ), 100)

            await safe_edit_message(progress_message,
                f"🗺️ *Progress Scraping Google Maps*\n\n"
                f"{make_progress_bar(pct)} {pct}%\n\n"
                f"🏢 {b_name[:40]}\n"
                f"📊 Bisnis: {total_checked}/{len(listings)} | Web Discan: {websites_scanned}\n"
                f"✅ Valid: {len(rows)} | ❌ Invalid: {counters['invalid']}\n"
                f"♻️ Dup Email: {counters['dup_email']} | 🏷️ Dup Domain: {counters['dup_domain']}\n"
                f"📧 Baru: {new_count}/{target_emails}\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )
            await asyncio.sleep(REQUEST_DELAY)

    await safe_edit_message(
        progress_message,
        f"✅ *Scraping Google Maps Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Bisnis Dicek: {total_checked}/{len(listings)}\n"
        f"Email Valid: {len(rows)}\n"
        f"Baru: {len(found_email_set) - session_start}"
    )
    return rows

# =====================================================
# CORE: GOOGLE DORKING
# =====================================================

async def scrape_dork(
    keyword_location: str,
    progress_message,
    target_emails: int = DEFAULT_TARGET_EMAILS,
    output_filename: str | None = None,
) -> list:
    keyword, location = split_keyword_location(keyword_location)

    _, existing_emails, existing_domains, _ = load_existing_master("dork", output_filename)
    found_email_set  = set(existing_emails)
    found_domain_set = set(existing_domains)
    session_start    = len(found_email_set)
    rows             = []
    counters         = {"invalid": 0, "dup_email": 0, "dup_domain": 0}

    total_queries  = 0
    total_found    = 0
    total_scanned  = 0
    latest_preview = []
    start_time     = time.time()
    stop_reason    = "Selesai"

    async with aiohttp.ClientSession() as session:
        for template in DORK_TEMPLATES:
            new_count = len(found_email_set) - session_start
            if new_count >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            query = template.format(keyword=keyword, location=location or "Indonesia")
            total_queries += 1

            new_count = len(found_email_set) - session_start
            pct = min(int((new_count / max(target_emails, 1)) * 100), 100)

            await safe_edit_message(progress_message,
                f"🔍 *Progress Google Dorking*\n\n"
                f"{make_progress_bar(pct)} {pct}%\n\n"
                f"🔎 Query {total_queries}/{len(DORK_TEMPLATES)}:\n{query[:80]}\n\n"
                f"🌐 URL Ditemukan: {total_found} | Discan: {total_scanned}\n"
                f"✅ Valid: {len(rows)} | ❌ Invalid: {counters['invalid']}\n"
                f"♻️ Dup Email: {counters['dup_email']} | 🏷️ Dup Domain: {counters['dup_domain']}\n"
                f"📧 Baru: {new_count}/{target_emails}\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )

            urls = await serpapi_search_web(query)
            total_found += len(urls)

            for url in urls:
                new_count = len(found_email_set) - session_start
                if new_count >= target_emails:
                    stop_reason = "Target email baru tercapai"
                    break

                url = normalize_url(url)
                if not is_valid_url(url):
                    continue

                total_scanned += 1
                extracted, _ = await scan_single_url(
                    session, url, DORK_MAX_PAGES_PER_URL,
                    start_url=url, use_playwright_fallback=True
                )

                for email, status, src_url in extracted:
                    accepted = await process_email(
                        session, email, status, src_url,
                        found_email_set, found_domain_set, counters
                    )
                    if accepted:
                        lbl = get_email_label(email)
                        rows.append([
                            len(rows) + 1, query, src_url, email, lbl,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Google Dorking",
                        ])
                        latest_preview.insert(0, f"✅ {email} [{lbl}]")
                        latest_preview = latest_preview[:10]

                        if len(rows) % PARTIAL_SAVE_EVERY == 0:
                            save_excel(rows, f"partial_dork_{int(time.time())}.xlsx", "dork")

                await asyncio.sleep(REQUEST_DELAY)
            await asyncio.sleep(1.0)

    await safe_edit_message(
        progress_message,
        f"✅ *Google Dorking Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Query: {total_queries}/{len(DORK_TEMPLATES)}\n"
        f"URL Ditemukan: {total_found} | Discan: {total_scanned}\n"
        f"Email Valid: {len(rows)}\n"
        f"Baru: {len(found_email_set) - session_start}"
    )
    return rows

# =====================================================
# CORE: DIRECTORY SCRAPER
# =====================================================

async def scrape_directory(
    keyword_location: str,
    progress_message,
    target_emails: int = DEFAULT_TARGET_EMAILS,
    output_filename: str | None = None,
) -> list:
    keyword, location = split_keyword_location(keyword_location)

    _, existing_emails, existing_domains, _ = load_existing_master("directory", output_filename)
    found_email_set  = set(existing_emails)
    found_domain_set = set(existing_domains)
    session_start    = len(found_email_set)
    rows             = []
    counters         = {"invalid": 0, "dup_email": 0, "dup_domain": 0}

    total_listings   = 0
    total_web        = 0
    latest_preview   = []
    start_time       = time.time()
    stop_reason      = "Selesai"

    async with aiohttp.ClientSession() as session:
        for dir_key, dir_conf in DIRECTORIES.items():
            new_count = len(found_email_set) - session_start
            if new_count >= target_emails:
                stop_reason = "Target email baru tercapai"
                break

            dir_name = dir_conf["name"]

            for page in range(1, DIRECTORY_MAX_PAGES + 1):
                new_count = len(found_email_set) - session_start
                if new_count >= target_emails:
                    stop_reason = "Target email baru tercapai"
                    break

                l_url = dir_conf["search_url"].format(
                    keyword=keyword.replace(" ", "+"),
                    location=location.replace(" ", "+"),
                    page=page,
                )

                pct = min(int((new_count / max(target_emails, 1)) * 100), 100)
                await safe_edit_message(progress_message,
                    f"📋 *Progress Scraping Direktori*\n\n"
                    f"{make_progress_bar(pct)} {pct}%\n\n"
                    f"📁 Direktori: {dir_name} | Hal: {page}\n"
                    f"🏢 Listing: {total_listings} | Web Discan: {total_web}\n"
                    f"✅ Valid: {len(rows)} | ❌ Invalid: {counters['invalid']}\n"
                    f"♻️ Dup Email: {counters['dup_email']} | 🏷️ Dup Domain: {counters['dup_domain']}\n"
                    f"📧 Baru: {new_count}/{target_emails}\n\n"
                    f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
                )

                html, _ = await fetch_html(session, l_url)
                if not html:
                    break

                soup     = BeautifulSoup(html, "html.parser")
                listings = soup.select(dir_conf["listing_selector"])
                if not listings:
                    break

                for listing in listings:
                    total_listings += 1
                    n_tag  = listing.select_one(dir_conf["name_selector"])
                    b_name = n_tag.get_text(strip=True) if n_tag else "Unknown"

                    u_tag = listing.select_one(dir_conf["url_selector"])
                    if not u_tag:
                        continue
                    web = u_tag.get("href", "").strip()
                    if not web or not is_valid_url(web):
                        continue

                    total_web += 1
                    extracted, _ = await scan_single_url(
                        session, web, DIRECTORY_DETAIL_MAX, start_url=web
                    )

                    for email, status, src_url in extracted:
                        accepted = await process_email(
                            session, email, status, src_url,
                            found_email_set, found_domain_set, counters
                        )
                        if accepted:
                            lbl = get_email_label(email)
                            rows.append([
                                len(rows) + 1, dir_name, b_name, src_url, email, lbl,
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Directory Scraping",
                            ])
                            latest_preview.insert(0, f"✅ {email} [{lbl}] | {b_name}")
                            latest_preview = latest_preview[:10]

                            if len(rows) % PARTIAL_SAVE_EVERY == 0:
                                save_excel(rows, f"partial_dir_{int(time.time())}.xlsx", "directory")

                            if (len(found_email_set) - session_start) >= target_emails:
                                stop_reason = "Target email baru tercapai"
                                break

                    await asyncio.sleep(REQUEST_DELAY)
                await asyncio.sleep(1.0)

    await safe_edit_message(
        progress_message,
        f"✅ *Scraping Direktori Selesai*\n\n"
        f"Alasan: {stop_reason}\n"
        f"Listing: {total_listings} | Web Discan: {total_web}\n"
        f"Email Valid: {len(rows)}\n"
        f"Baru: {len(found_email_set) - session_start}"
    )
    return rows

# =====================================================
# CORE: BATCH URL SCRAPER
# =====================================================

async def scrape_batch(
    urls: list,
    progress_message,
    target_emails: int = DEFAULT_TARGET_EMAILS,
    output_filename: str | None = None,
) -> list:
    _, existing_emails, existing_domains, _ = load_existing_master("web", output_filename)
    found_email_set  = set(existing_emails)
    found_domain_set = set(existing_domains)
    session_start    = len(found_email_set)
    all_rows         = []
    counters         = {"invalid": 0, "dup_email": 0, "dup_domain": 0}

    total_urls     = len(urls)
    scanned_urls   = 0
    total_pages    = 0
    latest_preview = []
    start_time     = time.time()

    async with aiohttp.ClientSession() as session:
        for url in urls:
            url = normalize_url(url.strip())
            if not is_valid_url(url):
                continue

            scanned_urls += 1
            new_count = len(found_email_set) - session_start
            pct = min(max(
                int((scanned_urls / max(total_urls, 1)) * 100),
                int((new_count / max(target_emails, 1)) * 100),
            ), 100)

            await safe_edit_message(progress_message,
                f"📦 *Progress Batch URL Scraping*\n\n"
                f"{make_progress_bar(pct)} {pct}%\n\n"
                f"🌐 URL {scanned_urls}/{total_urls}:\n{url[:80]}\n\n"
                f"📄 Total Halaman: {total_pages}\n"
                f"✅ Valid: {len(all_rows)} | ❌ Invalid: {counters['invalid']}\n"
                f"♻️ Dup Email: {counters['dup_email']} | 🏷️ Dup Domain: {counters['dup_domain']}\n"
                f"📧 Baru: {new_count}/{target_emails}\n"
                f"📡 Verifalia: {get_verifalia_status_text()}\n\n"
                f"Preview:\n" + "\n".join(latest_preview[:5] or ["-"])
            )

            extracted, pg_scanned = await scan_single_url(
                session, url,
                WEB_MAX_PAGES // max(total_urls, 1) + 10,
                start_url=url, use_playwright_fallback=True,
            )
            total_pages += pg_scanned

            for email, status, src_url in extracted:
                accepted = await process_email(
                    session, email, status, src_url,
                    found_email_set, found_domain_set, counters
                )
                if accepted:
                    lbl = get_email_label(email)
                    all_rows.append([
                        len(all_rows) + 1, url, email, lbl,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Batch Web Scraping",
                    ])
                    latest_preview.insert(0, f"✅ {email} [{lbl}]")
                    latest_preview = latest_preview[:10]

            if (len(found_email_set) - session_start) >= target_emails:
                break
            await asyncio.sleep(REQUEST_DELAY)

    await safe_edit_message(
        progress_message,
        f"✅ *Batch Scraping Selesai*\n\n"
        f"URL Diproses: {scanned_urls}/{total_urls}\n"
        f"Halaman: {total_pages}\n"
        f"Email Valid: {len(all_rows)}\n"
        f"Baru: {len(found_email_set) - session_start}"
    )
    return all_rows

# =====================================================
# TELEGRAM HANDLERS
# =====================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global VERIFALIA_LIMIT_REACHED, VERIFALIA_CACHE, MX_CACHE
    VERIFALIA_LIMIT_REACHED = False
    VERIFALIA_CACHE, MX_CACHE = {}, {}
    context.user_data.clear()

    status_icon = lambda ok: "✅" if ok else "⚠️"
    features = [
        f"{status_icon(PLAYWRIGHT_AVAILABLE)} Playwright (JS rendering)",
        f"{status_icon(PYMUPDF_AVAILABLE)} PyMuPDF (PDF scraping)",
        f"{status_icon(DNS_AVAILABLE)} MX checker (dnspython)",
        f"📡 Verifalia: {get_verifalia_status_text()}",
        f"🇮🇩 Geo-filter: Indonesia only (aktif)",
    ]
    await update.message.reply_text(
        "🤖 *Email Scraper Bot (Ultimate B2B Edition)*\n\nFitur aktif:\n" + "\n".join(features) + "\n\nPilih mode scraping:",
        reply_markup=main_menu(),
        parse_mode="Markdown",
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "back_main":
        context.user_data.clear()
        await query.edit_message_text("Pilih mode scraping:", reply_markup=main_menu())
        return

    if data == "manage_xlsx":
        await query.edit_message_text("📁 Kelola file XLSX hasil scraping.", reply_markup=manage_xlsx_menu())
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
        for i, f in enumerate(files[:20], 1):
            lines.append(f"{i}. {f['filename']}\n   Email: {f['rows']} | {f['size_kb']:.1f} KB | {f['modified']}")
        await query.edit_message_text("\n".join(lines), reply_markup=manage_xlsx_menu())
        return

    if data.startswith("delete_xlsx:"):
        scraping_type = data.split(":", 1)[1]
        files = list_xlsx_files(None)
        context.user_data[f"delete_files_{scraping_type}"] = files
        if not files:
            await query.edit_message_text("Belum ada file XLSX.", reply_markup=manage_xlsx_menu())
            return
        await query.edit_message_text(
            "🗑️ Pilih file XLSX yang mau dihapus:",
            reply_markup=build_xlsx_file_menu(files, "confirm_delete_xlsx", scraping_type)
        )
        return

    if data.startswith("confirm_delete_xlsx:"):
        _, scraping_type, idx_text = data.split(":")
        files = context.user_data.get(f"delete_files_{scraping_type}", [])
        try:
            f_info = files[int(idx_text)]
        except Exception:
            await query.edit_message_text("File tidak ditemukan.", reply_markup=manage_xlsx_menu())
            return
        context.user_data["delete_target"] = f_info
        await query.edit_message_text(
            f"Yakin hapus?\n{f_info['filename']}\nEmail: {f_info['rows']}",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Ya, hapus", callback_data="do_delete_xlsx")],
                [InlineKeyboardButton("❌ Batal",     callback_data="manage_xlsx")],
            ])
        )
        return

    if data == "do_delete_xlsx":
        f_info = context.user_data.get("delete_target")
        if not f_info:
            await query.edit_message_text("Tidak ada file yang dipilih.", reply_markup=manage_xlsx_menu())
            return
        try:
            path = f_info["path"]
            if (
                os.path.dirname(os.path.abspath(path)) == os.path.abspath(OUTPUT_DIR)
                and path.lower().endswith(".xlsx")
            ):
                os.remove(path)
                await query.edit_message_text(f"✅ File dihapus:\n{f_info['filename']}", reply_markup=manage_xlsx_menu())
            else:
                await query.edit_message_text("File tidak aman untuk dihapus.", reply_markup=manage_xlsx_menu())
        except Exception as e:
            await query.edit_message_text(f"Gagal hapus file:\n{e}", reply_markup=manage_xlsx_menu())
        return

    mode_map = [
        ("scrape_web",       "web",       "🌐 *Scraping Web*\n\nFormat: `https://example.com | 300`\n\nPilih output XLSX:"),
        ("scrape_maps",      "maps",      "🗺️ *Scraping Google Maps*\n\nFormat: `digital agency Jakarta | 200`\n\nPilih output XLSX:"),
        ("scrape_dork",      "dork",      "🔍 *Google Dorking*\n\nFormat: `digital agency Jakarta | 500`\n\nPilih output XLSX:"),
        ("scrape_directory", "directory", "📋 *Scraping Direktori Bisnis*\n\nFormat: `digital agency Jakarta | 300`\n\nPilih output XLSX:"),
        ("scrape_batch",     "web",       "📦 *Batch URL Scraping*\n\nPilih output XLSX dulu:"),
    ]
    for mode_key, label, prompt in mode_map:
        if data == mode_key:
            context.user_data.clear()
            context.user_data["pending_mode"] = label
            context.user_data["actual_mode"]  = mode_key.replace("scrape_", "")
            await query.edit_message_text(prompt, reply_markup=output_choice_menu(label), parse_mode="Markdown")
            return

    if data.startswith("xlsx_new:"):
        scraping_type = data.split(":", 1)[1]
        context.user_data["pending_mode"]          = scraping_type
        context.user_data["awaiting_new_xlsx_name"] = True
        await query.edit_message_text("🆕 Kirim nama file XLSX baru.\nContoh: `leads_jakarta.xlsx`")
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
            reply_markup=build_xlsx_file_menu(files, "select_xlsx", scraping_type)
        )
        return

    if data.startswith("select_xlsx:"):
        _, scraping_type, idx_text = data.split(":")
        files = context.user_data.get(f"select_files_{scraping_type}", [])
        try:
            f_info = files[int(idx_text)]
        except Exception:
            await query.edit_message_text("File tidak ditemukan.", reply_markup=output_choice_menu(scraping_type))
            return

        context.user_data["mode"]                              = scraping_type
        context.user_data[f"selected_{scraping_type}_xlsx"]   = f_info["filename"]
        actual_mode = context.user_data.get("actual_mode", scraping_type)

        input_prompts = {
            "web":       "🌐 Kirim URL website.\nContoh: `https://example.com | 300`",
            "maps":      "🗺️ Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 200`",
            "dork":      "🔍 Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 500`",
            "directory": "📋 Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 300`",
            "batch":     "📦 Kirim file .txt atau ketik URL langsung (1 per baris).",
        }
        await query.edit_message_text(
            f"✅ Output: {f_info['filename']}\n\n" + input_prompts.get(actual_mode, "Kirim input:")
        )
        return


async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    mode        = context.user_data.get("mode")
    actual_mode = context.user_data.get("actual_mode", mode)
    user_input  = update.message.text.strip() if update.message.text else ""

    if context.user_data.get("awaiting_new_xlsx_name"):
        scraping_type = context.user_data.get("pending_mode")
        filename      = safe_xlsx_name(user_input, scraping_type)
        filepath      = os.path.join(OUTPUT_DIR, filename)

        if os.path.exists(filepath):
            await update.message.reply_text(
                "Nama file sudah ada. Kirim nama berbeda.",
                reply_markup=output_choice_menu(scraping_type)
            )
            context.user_data.pop("awaiting_new_xlsx_name", None)
            return

        pd.DataFrame(columns=get_columns(scraping_type)).to_excel(filepath, index=False)
        context.user_data[f"select_files_{scraping_type}"] = list_xlsx_files(scraping_type)
        context.user_data["awaiting_new_xlsx_name"]         = False
        context.user_data["mode"]                           = scraping_type
        context.user_data[f"selected_{scraping_type}_xlsx"] = filename

        am = context.user_data.get("actual_mode", scraping_type)
        input_prompts = {
            "web":       "🌐 Kirim URL website.\nContoh: `https://example.com | 300`",
            "maps":      "🗺️ Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 200`",
            "dork":      "🔍 Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 500`",
            "directory": "📋 Kirim keyword + lokasi.\nContoh: `digital agency Jakarta | 300`",
            "batch":     "📦 Kirim file .txt atau ketik URL langsung (1 per baris).",
        }
        await update.message.reply_text(f"✅ File baru: {filename}\n\n" + input_prompts.get(am, "Kirim input:"))
        return

    if update.message.document:
        doc = update.message.document
        if (doc.file_name or "").lower().endswith(".txt") and actual_mode == "batch":
            content = (await (await doc.get_file()).download_as_bytearray()).decode("utf-8", errors="ignore")
            lines   = [l.strip() for l in content.splitlines() if l.strip()]

            target, urls = DEFAULT_TARGET_EMAILS, []
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

            p_msg  = await update.message.reply_text(f"📦 Memulai batch scraping {len(urls)} URL... Target: {target} email")
            output_filename = get_selected_output_filename(context, "web")
            rows   = await scrape_batch(urls, p_msg, target_emails=target, output_filename=output_filename)

            if not rows:
                await p_msg.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
                context.user_data.clear()
                return

            filepath, added, total = append_to_master_excel(rows, "web", output_filename)
            await update.message.reply_document(
                document=open(filepath, "rb"),
                filename=os.path.basename(filepath),
                caption=f"✅ Batch selesai.\nBaru: {added} | Total: {total}",
            )
            await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
            context.user_data.clear()
            return

        await update.message.reply_text("File tidak dikenali. Kirim .txt untuk batch URL.", reply_markup=main_menu())
        return

    if not mode:
        await update.message.reply_text("Pilih menu terlebih dahulu.", reply_markup=main_menu())
        return

    async def _run_and_send(scrape_fn, scraping_type: str, p_msg_text: str, *args, **kwargs):
        p_msg = await update.message.reply_text(p_msg_text)
        output_filename = get_selected_output_filename(context, scraping_type)
        try:
            rows = await scrape_fn(*args, p_msg, target_emails=kwargs.get("target", DEFAULT_TARGET_EMAILS), output_filename=output_filename)
            if not rows:
                await p_msg.edit_text("Selesai, tidak ada email valid baru.", reply_markup=main_menu())
            else:
                filepath, added, total = append_to_master_excel(rows, scraping_type, output_filename)
                await update.message.reply_document(
                    document=open(filepath, "rb"),
                    filename=os.path.basename(filepath),
                    caption=f"✅ Selesai.\nBaru: {added} | Total: {total}",
                )
                await update.message.reply_text("Mau scraping lagi?", reply_markup=main_menu())
        except Exception as e:
            logging.exception(e)
            await update.message.reply_text(f"Error: {e}", reply_markup=main_menu())
        context.user_data.clear()

    if actual_mode == "batch":
        lines  = [l.strip() for l in user_input.splitlines() if l.strip()]
        target, urls = DEFAULT_TARGET_EMAILS, []
        for line in lines:
            if line.upper().startswith("TARGET:"):
                try:
                    target = int(line.split(":", 1)[1].strip())
                except ValueError:
                    pass
            elif is_valid_url(normalize_url(line)):
                urls.append(line)

        if not urls:
            await update.message.reply_text("Tidak ada URL valid. Kirim 1 per baris.")
            return
        await _run_and_send(scrape_batch, "web", f"📦 Memulai batch... Target: {target} email", urls, target=target)

    elif actual_mode == "web":
        main_input, target = parse_target_from_input(user_input)
        url = normalize_url(main_input)
        if not is_valid_url(url):
            await update.message.reply_text("URL tidak valid. Contoh: https://example.com | 300", reply_markup=main_menu())
            return
        await _run_and_send(scrape_website, "web", f"🌐 Memulai scraping web... Target: {target} email", url, target=target)

    elif actual_mode == "maps":
        if not SERPAPI_KEY:
            await update.message.reply_text("SERPAPI_KEY belum diatur di .env.", reply_markup=main_menu())
            context.user_data.clear()
            return
        main_input, target = parse_target_from_input(user_input)
        await _run_and_send(scrape_maps, "maps", f"🗺️ Memulai Maps scraping... Target: {target} email", main_input, target=target)

    elif actual_mode == "dork":
        if not SERPAPI_KEY:
            await update.message.reply_text("SERPAPI_KEY belum diatur di .env.", reply_markup=main_menu())
            context.user_data.clear()
            return
        main_input, target = parse_target_from_input(user_input)
        await _run_and_send(scrape_dork, "dork", f"🔍 Memulai Google Dorking... Target: {target} email", main_input, target=target)

    elif actual_mode == "directory":
        main_input, target = parse_target_from_input(user_input)
        await _run_and_send(scrape_directory, "directory", f"📋 Memulai Direktori scraping... Target: {target} email", main_input, target=target)

# =====================================================
# MAIN
# =====================================================

def main() -> None:
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
    app.run_polling()


if __name__ == "__main__":
    main()