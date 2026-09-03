import argparse
import re
from dataclasses import dataclass
from urllib.parse import urljoin, urlparse, unquote
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm


# ============================================================
# CONFIG
# ============================================================

YC_API = "https://api.ycombinator.com/v0.1/companies"

DEFAULT_BATCHES = [
    "F2024",
    "W2025",
    "S2025",
    "F2025",
    "W2026",
    "S2026",
    "F2026",
]

OUTPUT_XLSX = "yc_outreach_database.xlsx"
OUTPUT_CSV = "yc_outreach_database.csv"

TIMEOUT = 20
WORKERS = 8

# NEVER collect more than 2 founders
MAX_FOUNDERS = 2

# Maximum number of company pages to inspect
MAX_PAGES_PER_COMPANY = 15

CONTACT_PATHS = [
    "/contact",
    "/about",
    "/team",
    "/founders",
    "/leadership",
    "/partners",
    "/partnerships",
    "/business",
    "/company",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


# ============================================================
# REGEX
# ============================================================

EMAIL_RE = re.compile(
    r"\b[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+"
    r"@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)+\b"
)

LINKEDIN_RE = re.compile(
    r"https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9_%\-./]+",
    re.IGNORECASE,
)


# ============================================================
# DATA CLASSES
# ============================================================

@dataclass
class Founder:
    name: str = ""
    linkedin: str = ""
    email: str = ""


@dataclass
class CompanyResult:
    company: str = ""
    yc_batch: str = ""
    headquarters_email: str = ""

    founder_1_name: str = ""
    founder_1_linkedin: str = ""
    founder_1_email: str = ""

    founder_2_name: str = ""
    founder_2_linkedin: str = ""
    founder_2_email: str = ""


# ============================================================
# HTTP
# ============================================================

def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def get_json(session, url, params=None):
    try:
        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT,
        )

        response.raise_for_status()
        return response.json()

    except Exception:
        return {}


def get_html(session, url):
    try:
        response = session.get(
            url,
            timeout=TIMEOUT,
            allow_redirects=True,
        )

        response.raise_for_status()

        content_type = response.headers.get(
            "content-type",
            "",
        ).lower()

        if "text/html" not in content_type:
            return ""

        return response.text

    except Exception:
        return ""


# ============================================================
# URL HELPERS
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    parsed = urlparse(url)

    if not parsed.netloc:
        return ""

    return url


def domain_of(url):
    try:
        domain = urlparse(url).netloc.lower()

        if domain.startswith("www."):
            domain = domain[4:]

        return domain

    except Exception:
        return ""


def normalize_linkedin(url):
    if not url:
        return ""

    url = unquote(str(url).strip())

    # Sometimes LinkedIn URLs have tracking parameters
    url = url.split("?", 1)[0]
    url = url.split("#", 1)[0]

    match = re.search(
        r"https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9_%\-./]+",
        url,
        re.IGNORECASE,
    )

    if not match:
        return ""

    result = match.group(0)

    result = result.rstrip(
        ".,;:)]}>\"'"
    )

    if not result.endswith("/"):
        result += "/"

    return result


# ============================================================
# YC API
# ============================================================

def get_companies_for_batch(session, batch):
    companies = []

    next_url = YC_API
    params = {"batch": batch}

    seen = set()

    while next_url:

        if next_url in seen:
            break

        seen.add(next_url)

        data = get_json(
            session,
            next_url,
            params=params,
        )

        if not data:
            break

        items = (
            data.get("companies")
            or data.get("items")
            or data.get("data")
            or []
        )

        if isinstance(items, list):
            companies.extend(items)

        next_page = data.get("nextPage")

        if next_page:
            next_url = normalize_url(next_page)
            params = None
        else:
            break

    return companies


# ============================================================
# YC FIELD HELPERS
# ============================================================

def get_company_name(row):
    for key in [
        "name",
        "company_name",
        "title",
    ]:
        value = row.get(key)

        if value:
            return str(value).strip()

    return ""


def get_batch(row, fallback=""):
    for key in [
        "batch",
        "yc_batch",
        "batch_name",
    ]:
        value = row.get(key)

        if value:
            return str(value).strip()

    return fallback


def get_website(row):
    for key in [
        "website",
        "website_url",
        "company_website",
        "companyWebsite",
        "company_url",
    ]:
        value = row.get(key)

        if isinstance(value, str) and value.strip():

            url = normalize_url(value)

            if url and "ycombinator.com" not in domain_of(url):
                return url

    return ""


def get_yc_url(row):
    # Direct YC URL fields
    for key in [
        "yc_url",
        "yc_company_url",
        "profile_url",
        "company_profile_url",
    ]:
        value = row.get(key)

        if isinstance(value, str) and "ycombinator.com" in value:
            return normalize_url(value)

    # Sometimes URL itself is YC
    for key in [
        "url",
        "company_url",
    ]:
        value = row.get(key)

        if isinstance(value, str) and "ycombinator.com" in value:
            return normalize_url(value)

    # Slug
    slug = row.get("slug")

    if slug:
        return (
            "https://www.ycombinator.com/companies/"
            + str(slug).strip()
        )

    # Some API versions may expose an ID
    company_id = row.get("id")

    if company_id:
        # Do not assume an ID is a valid YC slug.
        return ""

    return ""


# ============================================================
# FOUNDER EXTRACTION FROM API
# ============================================================

def extract_founders_from_api(row):
    founders = []

    possible = [
        row.get("founders"),
        row.get("people"),
        row.get("team"),
    ]

    founder_data = None

    for candidate in possible:

        if isinstance(candidate, list) and candidate:
            founder_data = candidate
            break

    if not founder_data:
        return founders

    for person in founder_data:

        if isinstance(person, str):

            name = person.strip()

            if name:
                founders.append(
                    Founder(name=name)
                )

        elif isinstance(person, dict):

            name = ""

            for key in [
                "name",
                "full_name",
                "display_name",
            ]:
                value = person.get(key)

                if value:
                    name = str(value).strip()
                    break

            linkedin = ""

            for key in [
                "linkedin",
                "linkedin_url",
                "linkedinUrl",
                "linkedIn",
            ]:
                value = person.get(key)

                if value:
                    linkedin = normalize_linkedin(
                        value
                    )

                    if linkedin:
                        break

            if name or linkedin:

                founders.append(
                    Founder(
                        name=name,
                        linkedin=linkedin,
                    )
                )

    # Remove duplicates
    unique = []
    seen = set()

    for founder in founders:

        key = (
            founder.name.lower().strip(),
            founder.linkedin.lower().strip(),
        )

        if key in seen:
            continue

        seen.add(key)
        unique.append(founder)

    return unique[:MAX_FOUNDERS]


# ============================================================
# YC PROFILE FOUNDER EXTRACTION
# ============================================================

def looks_like_person_name(text):
    if not text:
        return False

    text = " ".join(text.split()).strip()

    if len(text) < 3 or len(text) > 100:
        return False

    lower = text.lower()

    bad_phrases = [
        "linkedin",
        "twitter",
        "founder",
        "co-founder",
        "cofounder",
        "about",
        "contact",
        "company",
        "jobs",
        "careers",
        "website",
        "product",
        "yc",
        "batch",
        "view profile",
        "follow",
    ]

    if any(
        phrase in lower
        for phrase in bad_phrases
    ):
        return False

    # Avoid obvious sentences
    if any(
        char in text
        for char in [
            "@",
            "{",
            "}",
            "<",
            ">",
        ]
    ):
        return False

    words = text.split()

    if len(words) < 2 or len(words) > 5:
        return False

    # Names normally have alphabetic characters
    alpha_words = 0

    for word in words:

        cleaned = re.sub(
            r"[^A-Za-zÀ-ÖØ-öø-ÿ'-]",
            "",
            word,
        )

        if len(cleaned) >= 2:
            alpha_words += 1

    return alpha_words >= 2


def extract_founders_from_yc_html(
    html,
    existing_founders=None,
):
    """
    Extract founder names and LinkedIn URLs from
    the publicly accessible YC company profile.

    This intentionally only keeps the first 2 founders.
    """

    founders = []

    if existing_founders:
        founders.extend(
            existing_founders[:MAX_FOUNDERS]
        )

    if not html:
        return founders[:MAX_FOUNDERS]

    soup = BeautifulSoup(
        html,
        "lxml",
    )

    # --------------------------------------------------------
    # First: identify LinkedIn links and their nearby text.
    # This is the strongest signal.
    # --------------------------------------------------------

    linkedin_candidates = []

    for tag in soup.find_all(
        "a",
        href=True,
    ):

        linkedin = normalize_linkedin(
            tag.get("href", "")
        )

        if not linkedin:
            continue

        text_parts = []

        # Link's own text
        own_text = tag.get_text(
            " ",
            strip=True,
        )

        if own_text:
            text_parts.append(own_text)

        # Parent text
        if tag.parent:
            parent_text = tag.parent.get_text(
                " ",
                strip=True,
            )

            if parent_text:
                text_parts.append(parent_text)

        # Grandparent text
        if tag.parent and tag.parent.parent:
            grandparent_text = tag.parent.parent.get_text(
                " ",
                strip=True,
            )

            if grandparent_text:
                text_parts.append(
                    grandparent_text
                )

        context = " ".join(
            text_parts
        )

        linkedin_candidates.append(
            (
                linkedin,
                context,
            )
        )

    # --------------------------------------------------------
    # Match LinkedIn context to existing founder names
    # --------------------------------------------------------

    for founder in founders:

        if founder.linkedin:
            continue

        if not founder.name:
            continue

        name_lower = founder.name.lower()

        parts = name_lower.split()

        for linkedin, context in linkedin_candidates:

            context_lower = context.lower()

            if name_lower in context_lower:

                founder.linkedin = linkedin
                break

            if len(parts) >= 2:

                first = parts[0]
                last = parts[-1]

                if (
                    first in context_lower
                    and last in context_lower
                ):

                    founder.linkedin = linkedin
                    break

    # --------------------------------------------------------
    # If API gave no founder names, derive names from
    # the text around LinkedIn links.
    # --------------------------------------------------------

    if len(founders) < MAX_FOUNDERS:

        for linkedin, context in linkedin_candidates:

            if len(founders) >= MAX_FOUNDERS:
                break

            candidate_names = []

            # Break context into lines / chunks
            chunks = re.split(
                r"[\n|•·]+",
                context,
            )

            for chunk in chunks:

                chunk = " ".join(
                    chunk.split()
                ).strip()

                if looks_like_person_name(chunk):
                    candidate_names.append(
                        chunk
                    )

            # Prefer the shortest plausible name
            candidate_names.sort(
                key=len
            )

            selected_name = ""

            for candidate in candidate_names:

                duplicate = False

                for founder in founders:

                    if (
                        founder.name.lower()
                        == candidate.lower()
                    ):
                        duplicate = True
                        break

                if not duplicate:
                    selected_name = candidate
                    break

            if selected_name:

                founders.append(
                    Founder(
                        name=selected_name,
                        linkedin=linkedin,
                    )
                )

    # --------------------------------------------------------
    # Search structured data / JSON-LD
    # --------------------------------------------------------

    if len(founders) < MAX_FOUNDERS:

        for script in soup.find_all(
            "script",
            type="application/ld+json",
        ):

            text = script.get_text(
                " ",
                strip=True,
            )

            if not text:
                continue

            # Find LinkedIn URLs inside JSON-LD
            urls = LINKEDIN_RE.findall(
                text
            )

            for url in urls:

                if len(founders) >= MAX_FOUNDERS:
                    break

                linkedin = normalize_linkedin(
                    url
                )

                if not linkedin:
                    continue

                already = any(
                    f.linkedin == linkedin
                    for f in founders
                )

                if already:
                    continue

                # Look for a nearby "name"
                match = re.search(
                    r'"name"\s*:\s*"([^"]+)"',
                    text,
                    re.IGNORECASE,
                )

                if match:

                    name = match.group(1).strip()

                    if looks_like_person_name(name):

                        founders.append(
                            Founder(
                                name=name,
                                linkedin=linkedin,
                            )
                        )

    # Final dedupe
    unique = []
    seen = set()

    for founder in founders:

        key = (
            founder.name.lower().strip(),
            founder.linkedin.lower().strip(),
        )

        if key in seen:
            continue

        seen.add(key)
        unique.append(founder)

    return unique[:MAX_FOUNDERS]


# ============================================================
# EMAIL EXTRACTION
# ============================================================

def clean_email(email):
    if not email:
        return ""

    email = unquote(
        str(email).strip().lower()
    )

    if email.startswith("mailto:"):
        email = email[7:]

    email = email.split("?", 1)[0]

    email = email.rstrip(
        ".,;:)]}>\"'"
    )

    return email


def valid_email(email):
    email = clean_email(email)

    if not email:
        return False

    if not EMAIL_RE.fullmatch(email):
        return False

    local, domain = email.split("@", 1)

    if local in [
        "noreply",
        "no-reply",
        "donotreply",
        "do-not-reply",
    ]:
        return False

    if domain in [
        "example.com",
        "example.org",
        "example.net",
        "localhost",
    ]:
        return False

    return True


def extract_emails(soup):
    emails = set()

    # mailto
    for tag in soup.find_all(
        "a",
        href=True,
    ):

        href = tag.get(
            "href",
            "",
        )

        if href.lower().startswith("mailto:"):

            email = clean_email(
                href
            )

            if valid_email(email):
                emails.add(email)

    # Visible text
    text = soup.get_text(
        " ",
        strip=True,
    )

    for match in EMAIL_RE.findall(
        text
    ):

        email = clean_email(
            match
        )

        if valid_email(email):
            emails.add(email)

    return emails


# ============================================================
# HEADQUARTERS EMAIL
# ============================================================

GENERIC_PREFIXES = [
    "partnerships",
    "partnership",
    "business",
    "sales",
    "contact",
    "hello",
    "info",
    "team",
    "support",
]


def choose_hq_email(
    emails,
    company_domain,
):
    if not emails:
        return ""

    company_domain = (
        company_domain.lower().strip()
    )

    # First preference: company-domain emails
    company_emails = []

    for email in emails:

        domain = email.split(
            "@",
            1,
        )[1].lower()

        if (
            company_domain
            and domain == company_domain
        ):
            company_emails.append(
                email
            )

    # Prefer generic company mailbox
    for prefix in GENERIC_PREFIXES:

        for email in company_emails:

            local = email.split(
                "@",
                1,
            )[0].lower()

            if local == prefix:
                return email

    if company_emails:
        return sorted(
            company_emails
        )[0]

    # If no company-domain email,
    # use a generic public business email.
    for prefix in GENERIC_PREFIXES:

        for email in emails:

            local = email.split(
                "@",
                1,
            )[0].lower()

            if local == prefix:
                return email

    return sorted(
        emails
    )[0] if emails else ""


# ============================================================
# FOUNDER EMAIL MATCHING
# ============================================================

def find_founder_email(
    founder,
    emails,
    soup,
    company_domain,
):
    if not founder.name:
        return ""

    name = founder.name.lower().strip()

    parts = name.split()

    if len(parts) < 2:
        return ""

    first = parts[0]
    last = parts[-1]

    # Look for email close to founder's name
    for tag in soup.find_all(
        ["div", "section", "article", "li", "p"],
    ):

        text = tag.get_text(
            " ",
            strip=True,
        )

        text_lower = text.lower()

        if name not in text_lower:
            continue

        found = EMAIL_RE.findall(
            text
        )

        # Prefer company domain
        for raw in found:

            email = clean_email(
                raw
            )

            if not valid_email(email):
                continue

            domain = email.split(
                "@",
                1,
            )[1]

            if (
                company_domain
                and domain == company_domain
            ):
                return email

        # Then any actual public email
        for raw in found:

            email = clean_email(
                raw
            )

            if valid_email(email):
                return email

    # Check mailto links and surrounding text
    for tag in soup.find_all(
        "a",
        href=True,
    ):

        href = tag.get(
            "href",
            "",
        )

        if not href.lower().startswith(
            "mailto:"
        ):
            continue

        email = clean_email(
            href
        )

        if not valid_email(email):
            continue

        context = (
            tag.get_text(
                " ",
                strip=True,
            )
        )

        if tag.parent:
            context += " " + tag.parent.get_text(
                " ",
                strip=True,
            )

        context = context.lower()

        if (
            name in context
            or (
                first in context
                and last in context
            )
        ):
            return email

    return ""


# ============================================================
# DISCOVER INTERNAL LINKS
# ============================================================

def discover_useful_links(
    soup,
    base_url,
    company_domain,
):
    links = []

    keywords = [
        "contact",
        "team",
        "founder",
        "leadership",
        "about",
        "partner",
        "partnership",
        "business",
        "company",
        "sales",
    ]

    for tag in soup.find_all(
        "a",
        href=True,
    ):

        href = tag.get(
            "href",
            "",
        ).strip()

        if not href:
            continue

        if href.startswith(
            (
                "#",
                "mailto:",
                "tel:",
                "javascript:",
            )
        ):
            continue

        absolute = urljoin(
            base_url,
            href,
        )

        if domain_of(
            absolute
        ) != company_domain:
            continue

        text = tag.get_text(
            " ",
            strip=True,
        ).lower()

        path = urlparse(
            absolute
        ).path.lower()

        combined = (
            text
            + " "
            + path
        )

        if any(
            keyword in combined
            for keyword in keywords
        ):
            links.append(
                absolute
            )

    return links


# ============================================================
# COUNTRY FILTER
# ============================================================

def matches_country(
    row,
    country,
):
    if not country:
        return True

    target = country.lower().strip()

    values = []

    for key in [
        "country",
        "location",
        "headquarters",
        "hq",
        "city",
        "region",
    ]:

        value = row.get(key)

        if value:
            values.append(
                str(value).lower()
            )

    location = row.get(
        "location"
    )

    if isinstance(
        location,
        dict,
    ):

        for value in location.values():

            if value:
                values.append(
                    str(value).lower()
                )

    combined = " ".join(
        values
    )

    return target in combined


# ============================================================
# MAIN COMPANY CRAWLER
# ============================================================

def crawl_company(
    row,
    fallback_batch="",
):
    session = make_session()

    company = get_company_name(
        row
    )

    batch = get_batch(
        row,
        fallback_batch,
    )

    result = CompanyResult(
        company=company,
        yc_batch=batch,
    )

    try:

        # ----------------------------------------------------
        # API founder data
        # ----------------------------------------------------

        founders = extract_founders_from_api(
            row
        )

        # ----------------------------------------------------
        # URLs
        # ----------------------------------------------------

        yc_url = get_yc_url(
            row
        )

        website = get_website(
            row
        )

        # ----------------------------------------------------
        # Fetch YC profile FIRST
        # ----------------------------------------------------

        yc_soup = None

        if yc_url:

            yc_html = get_html(
                session,
                yc_url,
            )

            if yc_html:

                yc_soup = BeautifulSoup(
                    yc_html,
                    "lxml",
                )

                founders = extract_founders_from_yc_html(
                    yc_html,
                    founders,
                )

        founders = founders[
            :MAX_FOUNDERS
        ]

        # ----------------------------------------------------
        # Website crawling
        # ----------------------------------------------------

        all_emails = set()

        all_linkedins = set()

        visited = set()

        queue = []

        if website:
            queue.append(
                website
            )

            website_domain = domain_of(
                website
            )

            for path in CONTACT_PATHS:

                queue.append(
                    urljoin(
                        website,
                        path,
                    )
                )

        # ----------------------------------------------------
        # Crawl website
        # ----------------------------------------------------

        while (
            queue
            and len(visited)
            < MAX_PAGES_PER_COMPANY
        ):

            url = queue.pop(0)

            url = normalize_url(
                url
            )

            if not url:
                continue

            if url in visited:
                continue

            visited.add(url)

            html = get_html(
                session,
                url,
            )

            if not html:
                continue

            soup = BeautifulSoup(
                html,
                "lxml",
            )

            # Public emails
            all_emails.update(
                extract_emails(
                    soup
                )
            )

            # Public LinkedIn links
            for linkedin in extract_linkedins_from_soup(
                soup
            ):
                all_linkedins.add(
                    linkedin
                )

            # More internal useful pages
            discovered = discover_useful_links(
                soup,
                url,
                website_domain,
            )

            for link in discovered:

                if (
                    link not in visited
                    and link not in queue
                ):
                    queue.append(
                        link
                    )

        # ----------------------------------------------------
        # Build combined soup from website pages
        # ----------------------------------------------------

        combined_parts = []

        for url in visited:

            html = get_html(
                session,
                url,
            )

            if html:
                combined_parts.append(
                    html
                )

        combined_html = "\n".join(
            combined_parts
        )

        combined_soup = BeautifulSoup(
            combined_html,
            "lxml",
        )

        # ----------------------------------------------------
        # Also try to get founder names/LinkedIns from
        # the company's own website
        # ----------------------------------------------------

        founders = extract_founders_from_yc_html(
            combined_html,
            founders,
        )

        founders = founders[
            :MAX_FOUNDERS
        ]

        # ----------------------------------------------------
        # Match existing LinkedIns to founder names
        # ----------------------------------------------------

        for founder in founders:

            if founder.linkedin:
                continue

            if not founder.name:
                continue

            name = founder.name.lower()

            parts = name.split()

            if len(parts) < 2:
                continue

            first = parts[0]
            last = parts[-1]

            for linkedin in all_linkedins:

                slug = linkedin.lower()

                if (
                    first in slug
                    and last in slug
                ):
                    founder.linkedin = linkedin
                    break

        # ----------------------------------------------------
        # HQ email
        # ----------------------------------------------------

        result.headquarters_email = choose_hq_email(
            all_emails,
            website_domain if website else "",
        )

        # ----------------------------------------------------
        # Founder emails
        # ----------------------------------------------------

        for founder in founders:

            founder.email = find_founder_email(
                founder,
                all_emails,
                combined_soup,
                website_domain if website else "",
            )

        # ----------------------------------------------------
        # Save first founder
        # ----------------------------------------------------

        if len(founders) >= 1:

            result.founder_1_name = (
                founders[0].name
            )

            result.founder_1_linkedin = (
                founders[0].linkedin
            )

            result.founder_1_email = (
                founders[0].email
            )

        # ----------------------------------------------------
        # Save second founder
        # ----------------------------------------------------

        if len(founders) >= 2:

            result.founder_2_name = (
                founders[1].name
            )

            result.founder_2_linkedin = (
                founders[1].linkedin
            )

            result.founder_2_email = (
                founders[1].email
            )

        return result

    except Exception as e:

        print(
            f"[ERROR] {company}: {e}"
        )

        # IMPORTANT:
        # Never lose the company row because scraping
        # contact information failed.
        return result

    finally:

        session.close()


# ============================================================
# LINKEDIN EXTRACTION HELPER
# ============================================================

def extract_linkedins_from_soup(
    soup,
):
    linkedins = set()

    for tag in soup.find_all(
        "a",
        href=True,
    ):

        linkedin = normalize_linkedin(
            tag.get(
                "href",
                "",
            )
        )

        if linkedin:
            linkedins.add(
                linkedin
            )

    text = soup.get_text(
        " ",
        strip=True,
    )

    for match in LINKEDIN_RE.findall(
        text
    ):

        linkedin = normalize_linkedin(
            match
        )

        if linkedin:
            linkedins.add(
                linkedin
            )

    return linkedins


# ============================================================
# DATAFRAME
# ============================================================

def build_dataframe(
    results,
):
    columns = [
        "Company",
        "YC Batch",
        "Headquarters Email",
        "Founder 1 Name",
        "Founder 1 LinkedIn",
        "Founder 1 Email",
        "Founder 2 Name",
        "Founder 2 LinkedIn",
        "Founder 2 Email",
    ]

    rows = []

    for r in results:

        rows.append(
            {
                "Company": r.company,
                "YC Batch": r.yc_batch,
                "Headquarters Email": r.headquarters_email,

                "Founder 1 Name": r.founder_1_name,
                "Founder 1 LinkedIn": r.founder_1_linkedin,
                "Founder 1 Email": r.founder_1_email,

                "Founder 2 Name": r.founder_2_name,
                "Founder 2 LinkedIn": r.founder_2_linkedin,
                "Founder 2 Email": r.founder_2_email,
            }
        )

    return pd.DataFrame(
        rows,
        columns=columns,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--batches",
        nargs="+",
        default=DEFAULT_BATCHES,
    )

    parser.add_argument(
        "--country",
        default="",
    )

    parser.add_argument(
        "--max-companies",
        type=int,
        default=0,
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=WORKERS,
    )

    args = parser.parse_args()

    print()
    print("=" * 70)
    print("YC STARTUP OUTREACH DATABASE")
    print("=" * 70)
    print()

    print(
        "Batches:",
        ", ".join(args.batches),
    )

    if args.country:
        print(
            "Country:",
            args.country,
        )

    # --------------------------------------------------------
    # Fetch companies
    # --------------------------------------------------------

    session = make_session()

    all_rows = []

    for batch in args.batches:

        print(
            f"Fetching YC batch {batch}..."
        )

        rows = get_companies_for_batch(
            session,
            batch,
        )

        print(
            f"  Found {len(rows)} companies"
        )

        for row in rows:

            if not matches_country(
                row,
                args.country,
            ):
                continue

            row["_fallback_batch"] = batch

            all_rows.append(
                row
            )

    session.close()

    # --------------------------------------------------------
    # Deduplicate
    # --------------------------------------------------------

    unique = []

    seen = set()

    for row in all_rows:

        name = get_company_name(
            row
        )

        if not name:
            continue

        key = name.lower().strip()

        if key in seen:
            continue

        seen.add(key)

        unique.append(
            row
        )

    all_rows = unique

    # --------------------------------------------------------
    # Limit for testing
    # --------------------------------------------------------

    if (
        args.max_companies
        and args.max_companies > 0
    ):

        all_rows = all_rows[
            :args.max_companies
        ]

    print()
    print(
        f"Companies queued: {len(all_rows)}"
    )
    print()

    # --------------------------------------------------------
    # No companies
    # --------------------------------------------------------

    if not all_rows:

        df = pd.DataFrame(
            columns=[
                "Company",
                "YC Batch",
                "Headquarters Email",
                "Founder 1 Name",
                "Founder 1 LinkedIn",
                "Founder 1 Email",
                "Founder 2 Name",
                "Founder 2 LinkedIn",
                "Founder 2 Email",
            ]
        )

        df.to_excel(
            OUTPUT_XLSX,
            index=False,
        )

        df.to_csv(
            OUTPUT_CSV,
            index=False,
        )

        print(
            "No companies found."
        )

        return

    # --------------------------------------------------------
    # Crawl concurrently
    # --------------------------------------------------------

    results = []

    with ThreadPoolExecutor(
        max_workers=args.workers
    ) as executor:

        futures = {
            executor.submit(
                crawl_company,
                row,
                row.get(
                    "_fallback_batch",
                    "",
                ),
            ): row
            for row in all_rows
        }

        for future in tqdm(
            as_completed(futures),
            total=len(futures),
            desc="Extracting",
        ):

            row = futures[future]

            try:

                result = future.result()

                if result.company:
                    results.append(
                        result
                    )

            except Exception as e:

                company = get_company_name(
                    row
                )

                print(
                    f"[ERROR] {company}: {e}"
                )

                # Preserve company even on failure
                results.append(
                    CompanyResult(
                        company=company,
                        yc_batch=get_batch(
                            row,
                            row.get(
                                "_fallback_batch",
                                "",
                            ),
                        ),
                    )
                )

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    results.sort(
        key=lambda r: (
            r.yc_batch,
            r.company.lower(),
        )
    )

    # --------------------------------------------------------
    # Export
    # --------------------------------------------------------

    df = build_dataframe(
        results
    )

    df.to_excel(
        OUTPUT_XLSX,
        index=False,
    )

    df.to_csv(
        OUTPUT_CSV,
        index=False,
    )

    # --------------------------------------------------------
    # Stats
    # --------------------------------------------------------

    hq_count = sum(
        bool(
            r.headquarters_email
        )
        for r in results
    )

    founder_linkedin_count = sum(
        bool(r.founder_1_linkedin)
        + bool(r.founder_2_linkedin)
        for r in results
    )

    founder_email_count = sum(
        bool(r.founder_1_email)
        + bool(r.founder_2_email)
        for r in results
    )

    founder_name_count = sum(
        bool(r.founder_1_name)
        + bool(r.founder_2_name)
        for r in results
    )

    # --------------------------------------------------------
    # Final
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("DONE")
    print("=" * 70)

    print(
        f"Companies: {len(results)}"
    )

    print(
        f"Founder names: {founder_name_count}"
    )

    print(
        f"HQ emails: {hq_count}"
    )

    print(
        f"Founder LinkedIns: {founder_linkedin_count}"
    )

    print(
        f"Founder emails: {founder_email_count}"
    )

    print()

    print(
        f"Excel: {OUTPUT_XLSX}"
    )

    print(
        f"CSV:   {OUTPUT_CSV}"
    )

    print()


if __name__ == "__main__":
    main()