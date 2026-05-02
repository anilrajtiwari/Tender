#!/usr/bin/env python3
"""
============================================================
POWERGRID India International Business Department
Tender Intelligence Analyzer & Report Generator  v3.0
============================================================
Streamlit Web App — Cloud Ready
Sources  : www.globaltransmission.info  (requests session)
           www.tendersinfo.com          (requests session)
           Open web search via DuckDuckGo Lite
           fragilestatesindex.org       (live FSI scores)
Output   : Excel (.xlsx) with clickable hyperlinks + PDF
           (downloadable via Streamlit buttons)
============================================================
"""

import streamlit as st
import sys, time, re, json, logging, datetime, textwrap, io
from urllib.parse import urlparse, parse_qs, unquote, urljoin

import requests
from bs4 import BeautifulSoup
import pandas as pd

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
)
from reportlab.lib.enums import TA_CENTER

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="POWERGRID IB — Tender Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  CREDENTIALS — from Streamlit Secrets or defaults
# ─────────────────────────────────────────────
def get_credentials():
    try:
        return {
            "tendersinfo": {
                "login_url": "https://www.tendersinfo.com/login",
                "username":  st.secrets["tendersinfo"]["username"],
                "password":  st.secrets["tendersinfo"]["password"],
            },
            "globaltransmission": {
                "login_url": "https://globaltransmission.info/login/",
                "username":  st.secrets["globaltransmission"]["username"],
                "password":  st.secrets["globaltransmission"]["password"],
            },
        }
    except Exception:
        # Fallback — will attempt unauthenticated scraping
        return {
            "tendersinfo":       {"login_url": "", "username": "", "password": ""},
            "globaltransmission":{"login_url": "", "username": "", "password": ""},
        }

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# ─────────────────────────────────────────────
#  KEYWORD FILTERS
# ─────────────────────────────────────────────
POSITIVE_KEYWORDS = [
    "transmission line", "substation", "switchyard", "grid",
    "hvdc", "uhvdc", "interconnection", "power evacuation",
    "renewable integration", "consultancy", "pmc", "feasibility",
    "ppp power", "ehv", "uhv", "400 kv", "765 kv", "220 kv",
    "132 kv", "500 kv", "230 kv", "overhead line", "ohl",
    "power transformer", "gis", "ais", "smart grid",
    "grid modernisation", "grid modernization", "power sector",
    "transmission system", "electricity transmission",
]

NEGATIVE_KEYWORDS = [
    "road", "highway", "bridge", "school", "hospital",
    "pipeline", "oil", "gas pipeline", "building", "housing",
    "water supply", "sewerage", "railway", "metro", "defence",
    "weapon", "military",
]

CONSULTANCY_KEYWORDS = [
    "consultancy", "consulting", "consultant", "pmc",
    "project management consultant", "project management consultancy",
    "feasibility study", "feasibility", "pre-feasibility",
    "advisory", "adviser", "advisor",
    "detailed design", "conceptual design", "design review",
    "dpr", "detailed project report",
    "supervision", "supervisory", "site supervision",
    "owner's engineer", "owners engineer", "independent engineer",
    "technical assistance", "technical advisory",
    "study", "assessment", "audit",
    "planning", "master plan", "due diligence",
    "environmental impact", "esia", "eia",
    "transaction advisory", "bid advisory",
    "inspection", "third-party inspection",
    "preparation of bid", "preparation of tender",
    "technical consultant",
]

CONSTRUCTION_SIGNAL_KEYWORDS = [
    "epc contract", "epc turnkey", "turnkey epc",
    "supply installation commissioning",
    "erection and commissioning", "erection commissioning",
    "civil and structural works",
    "invitation for bids",
    "international competitive bidding",
]

EXCLUDE_COUNTRIES = ["india"]

STRATEGIC_COUNTRIES = [
    "egypt", "nigeria", "uganda", "tanzania", "kenya",
    "mozambique", "zambia", "jordan", "uae", "united arab emirates",
    "thailand", "nepal", "mauritania", "mali", "guinea",
]

FSI_FALLBACK = {
    "australia": ("Low",  21.6),  "austria": ("Low", 18.2),
    "canada":    ("Low",  22.4),  "denmark": ("Low", 16.1),
    "finland":   ("Low",  14.8),  "france":  ("Low", 38.1),
    "germany":   ("Low",  21.3),  "greece":  ("Low", 55.4),
    "israel":    ("Low",  51.6),  "jordan":  ("Low", 72.1),
    "kazakhstan":("Low",  65.8),  "lithuania":("Low", 35.2),
    "morocco":   ("Medium", 71.4),"namibia": ("Low", 57.2),
    "poland":    ("Low",  39.6),  "slovakia":("Low", 38.7),
    "thailand":  ("Low",  65.3),
    "trinidad and tobago": ("Low", 56.2),
    "turkey":    ("Medium", 75.2),"uae":     ("Low", 39.4),
    "united arab emirates": ("Low", 39.4),
    "united kingdom": ("Low", 30.8), "uruguay": ("Low", 45.1),
    "angola":    ("Medium", 82.3),"congo":   ("High", 88.1),
    "ecuador":   ("Medium", 71.8),"egypt":   ("Medium", 78.4),
    "kenya":     ("Medium", 74.3),"mauritania": ("Medium", 79.1),
    "mozambique":("Medium", 81.2),"nepal":   ("Medium", 68.7),
    "nigeria":   ("Medium", 84.6),"paraguay":("Medium", 65.4),
    "peru":      ("Medium", 62.1),"senegal": ("Medium", 72.8),
    "tanzania":  ("Medium", 76.2),"uganda":  ("Medium", 78.3),
    "zambia":    ("Medium", 76.9),"zimbabwe":("High",   86.4),
    "cameroon":  ("High",  87.3),
    "central african republic": ("High", 106.9),
    "democratic republic of congo": ("High", 107.2),
    "drc":       ("High",  107.2),"ethiopia":("High",  90.6),
    "guinea":    ("High",  87.4), "iraq":    ("High",  86.4),
    "libya":     ("High",  93.7), "mali":    ("High",  94.2),
    "myanmar":   ("High",  92.1), "south sudan": ("High", 111.1),
    "sudan":     ("High",  108.4),"syria":   ("High",  110.8),
    "venezuela": ("High",  88.2), "yemen":   ("High",  112.4),
}

RISK_LABEL = {
    "Low":    "Proceed",
    "Medium": "Proceed with Caution",
    "High":   "Management Decision Required",
}

COLUMNS = [
    "S.No.", "Name of Work", "Scope of Work", "Organisation",
    "Source", "Country", "Deadline", "Est. Cost",
    "Score", "FSI Score", "Risk Category", "Risk Label",
    "Recommendation", "Remarks", "Tender Link",
]

FSI_SOURCE_NOTE = "Source: fragilestatesindex.org/global-data/"


# ═══════════════════════════════════════════════════════════
#  SECTION 0 – LIVE FSI SCRAPER
# ═══════════════════════════════════════════════════════════

def fetch_live_fsi_scores():
    FSI_URL = "https://fragilestatesindex.org/global-data/"

    def _band(score):
        if score < 60:
            return "Low"
        elif score < 80:
            return "Medium"
        else:
            return "High"

    try:
        resp = requests.get(FSI_URL, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        fsi_data = {}

        table = soup.find("table")
        if table:
            rows = table.find_all("tr")
            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if len(cells) >= 3:
                    country_cell = cells[1].get_text(strip=True).lower()
                    score_cell   = cells[2].get_text(strip=True)
                    try:
                        score = float(score_cell.replace(",", "."))
                        fsi_data[country_cell] = (_band(score), round(score, 1))
                    except ValueError:
                        pass

        if not fsi_data:
            scripts = soup.find_all("script")
            for sc in scripts:
                txt = sc.string or ""
                if "\"country\"" in txt.lower() and "\"total\"" in txt.lower():
                    m = re.search(r'\[(\{.*?"country".*?\})\]', txt, re.DOTALL)
                    if m:
                        try:
                            arr = json.loads("[" + m.group(1) + "]")
                            for entry in arr:
                                c = entry.get("country", "").lower().strip()
                                s = entry.get("total", entry.get("score", None))
                                if c and s is not None:
                                    score = float(s)
                                    fsi_data[c] = (_band(score), round(score, 1))
                        except Exception:
                            pass

        if fsi_data:
            return fsi_data, f"Live FSI data loaded ({len(fsi_data)} countries)"
        else:
            return FSI_FALLBACK, "Live FSI parse failed — using fallback table"

    except Exception as e:
        return FSI_FALLBACK, f"Live FSI fetch failed ({e}) — using fallback table"


# ═══════════════════════════════════════════════════════════
#  SECTION 1 – SESSION LOGIN (requests-based, no Selenium)
# ═══════════════════════════════════════════════════════════

def login_globaltransmission(username, password):
    """Login to globaltransmission.info using requests (WordPress form)."""
    session = requests.Session()
    session.headers.update(HEADERS)
    if not username or not password:
        return session  # unauthenticated
    try:
        login_url = "https://globaltransmission.info/login/"
        r = session.get(login_url, timeout=15)
        soup = BeautifulSoup(r.text, "lxml")
        nonce = ""
        nonce_input = soup.find("input", {"name": "redirect_to"})
        redirect = nonce_input["value"] if nonce_input else "/"

        data = {
            "log":         username,
            "pwd":         password,
            "wp-submit":   "Log In",
            "redirect_to": redirect,
            "testcookie":  "1",
        }
        session.cookies.set("wordpress_test_cookie", "WP Cookie check")
        session.post(login_url, data=data, timeout=15)
    except Exception:
        pass
    return session


def login_tendersinfo(username, password):
    """Login to tendersinfo.com using requests."""
    session = requests.Session()
    session.headers.update(HEADERS)
    if not username or not password:
        return session
    try:
        login_url = "https://www.tendersinfo.com/login"
        r = session.get(login_url, timeout=15)
        data = {"email": username, "password": password}
        session.post(login_url, data=data, timeout=15)
    except Exception:
        pass
    return session


# ═══════════════════════════════════════════════════════════
#  SECTION 2 – SCRAPERS
# ═══════════════════════════════════════════════════════════

class SessionBase:
    def __init__(self, session=None):
        if session:
            self.session = session
        else:
            self.session = requests.Session()
            self.session.headers.update(HEADERS)

    def _get(self, url, **kw):
        try:
            r = self.session.get(url, timeout=25, **kw)
            r.raise_for_status()
            return r
        except Exception:
            return None


class GlobalTransmissionScraper(SessionBase):
    BASE = "https://globaltransmission.info"
    CATEGORIES = [
        "/category/tenders-contracts/middle-east-africa-tenders/",
        "/category/tenders-contracts/asia-pacific-tenders/",
        "/category/tenders-contracts/latin-america-tenders/",
        "/category/tenders-contracts/europe-tenders/",
        "/category/tenders-contracts/north-america-tenders/",
    ]

    def scrape_category(self, path):
        tenders = []
        url = self.BASE + path
        r = self._get(url)
        if not r:
            return tenders
        soup  = BeautifulSoup(r.text, "lxml")
        links = soup.select("h2 a") or soup.select("h3 a")
        for link in links:
            title = link.get_text(strip=True)
            href  = link.get("href", "")
            if not href:
                continue
            detail = self._scrape_detail(href, title)
            if detail:
                tenders.append(detail)
            time.sleep(0.4)
        return tenders

    def _scrape_detail(self, url, title):
        r = self._get(url)
        if not r:
            return None
        soup = BeautifulSoup(r.text, "lxml")
        body = soup.select_one("div.entry-content") or soup.select_one("article")
        text = body.get_text(" ", strip=True) if body else ""

        country  = _extract_field(text, "Country")
        org      = (_extract_field(text, "Organisation") or _extract_field(text, "Organization"))
        scope    = (_extract_field(text, r"Description/Scope of work") or _extract_field(text, "Description"))
        deadline = (_extract_field(text, r"Closing [Dd]ate") or _extract_field(text, r"Submission [Dd]ate"))

        published = ""
        time_el = soup.find("time", class_=re.compile(r"entry|published|post"))
        if not time_el:
            time_el = soup.find("time")
        if time_el:
            published = (time_el.get("datetime", "") or time_el.get_text(strip=True))[:20]
        if not published:
            meta_date = soup.find("meta", {"property": "article:published_time"})
            if meta_date:
                published = meta_date.get("content", "")[:20]

        if not country:
            return None
        return {
            "name":      title,
            "country":   country,
            "org":       org or "",
            "scope":     scope or text[:300],
            "deadline":  deadline or "See tender doc",
            "cost":      _extract_cost(text),
            "source":    "globaltransmission.info",
            "url":       url,
            "published": published,
        }

    def run(self, progress_cb=None):
        results = []
        for i, cat in enumerate(self.CATEGORIES):
            if progress_cb:
                progress_cb(f"globaltransmission.info — {cat.split('/')[-2]}")
            results.extend(self.scrape_category(cat))
            time.sleep(0.8)
        return results


class TendersInfoScraper(SessionBase):
    BASE = "https://www.tendersinfo.com"
    SEARCH_URLS = [
        "/global-consulting-services-tenders.php",
        "/global-consultancy-services-tenders.php",
        "/global-project-management-tenders.php",
        "/global-feasibility-study-tenders.php",
        "/global-transmission-lines-tenders.php",
        "/global-energy-and-power-renewable-energy-tenders.php",
    ]

    def scrape_listing(self, path):
        tenders = []
        url = self.BASE + path
        r = self._get(url)
        if not r:
            return tenders
        soup = BeautifulSoup(r.text, "lxml")
        rows = soup.select("table tr") or soup.select("div.tender-item")
        for row in rows:
            cells = row.find_all("td")
            if len(cells) < 3:
                continue
            name      = cells[0].get_text(" ", strip=True)
            country   = cells[1].get_text(strip=True) if len(cells) > 1 else ""
            deadline  = cells[2].get_text(strip=True) if len(cells) > 2 else ""
            cost      = cells[3].get_text(strip=True) if len(cells) > 3 else "N/A"
            scope     = cells[4].get_text(" ", strip=True) if len(cells) > 4 else ""
            published = cells[5].get_text(strip=True) if len(cells) > 5 else ""
            link_tag  = row.find("a")
            href      = (link_tag.get("href", "") if link_tag else "")
            if href and not href.startswith("http"):
                href = self.BASE + ("" if href.startswith("/") else "/") + href
            if name and country:
                tenders.append({
                    "name":      name,
                    "country":   country,
                    "org":       "",
                    "scope":     scope or name,
                    "deadline":  deadline or "See tender doc",
                    "cost":      cost,
                    "source":    "tendersinfo.com",
                    "url":       href,
                    "published": published,
                })
        return tenders

    def run(self, progress_cb=None):
        results = []
        for path in self.SEARCH_URLS:
            if progress_cb:
                progress_cb(f"tendersinfo.com — {path.split('/')[-1][:40]}")
            results.extend(self.scrape_listing(path))
            time.sleep(0.8)
        return results


class WebSearchScraper:
    QUERIES = [
        "power transmission consultancy RFP international tender 2025 2026",
        "substation PMC supervision consultancy tender international 2026",
        "HVDC feasibility study RFP Africa Asia 2026",
        "grid modernization advisory consultancy tender 2026",
        "EHV transmission design consultancy bid 2025 2026",
        "power sector advisory feasibility tender Africa Middle East 2026",
        "owner's engineer transmission project RFP 2026",
        "PPP power sector consultancy bid developing countries 2026",
        "renewable energy grid integration consultancy tender 2026",
        "transmission line feasibility study bid Africa Asia 2025 2026",
        "substation design consultancy international competitive 2026",
        "project management consultancy power sector Africa 2026",
    ]
    DDG_URL = "https://lite.duckduckgo.com/lite/"

    EXTRA_PORTALS = [
        {
            "name": "UNGM",
            "url": "https://www.ungm.org/Public/Notice?noticeTypeId=0&categoryId=all&keyword=transmission+consultancy",
            "link_selector": "a.title",
            "base": "https://www.ungm.org",
        },
        {
            "name": "ADB Tenders",
            "url": "https://www.adb.org/projects/tenders/active?keywords=transmission+consultancy&sector=Energy",
            "link_selector": "td.views-field-title a",
            "base": "https://www.adb.org",
        },
        {
            "name": "AfDB",
            "url": "https://www.afdb.org/en/projects-and-operations/procurement?keywords=transmission+consultancy",
            "link_selector": "h3.field-content a",
            "base": "https://www.afdb.org",
        },
        {
            "name": "EBRD",
            "url": "https://ecepp.ebrd.com/delta/viewActive.html?keywords=transmission+consultancy",
            "link_selector": "a.notice-link",
            "base": "https://ecepp.ebrd.com",
        },
        {
            "name": "World Bank",
            "url": "https://projects.worldbank.org/en/projects-operations/procurement/procurementsearch?searchTerm=transmission+consultancy&lang=en",
            "link_selector": "a.search-result-item",
            "base": "https://projects.worldbank.org",
        },
    ]

    @staticmethod
    def _resolve_ddg_url(href):
        if not href:
            return href
        if href.startswith("//"):
            href = "https:" + href
        try:
            parsed = urlparse(href)
            if "duckduckgo.com" in parsed.netloc:
                qs = parse_qs(parsed.query)
                if "uddg" in qs:
                    return unquote(qs["uddg"][0])
        except Exception:
            pass
        if href and not href.startswith("http"):
            return ""
        return href

    def _search(self, query):
        results = []
        try:
            r = requests.post(
                self.DDG_URL, data={"q": query},
                headers=HEADERS, timeout=15,
            )
            soup = BeautifulSoup(r.text, "lxml")
            for a in soup.select("a.result-link")[:8]:
                title   = a.get_text(strip=True)
                href    = self._resolve_ddg_url(a.get("href", ""))
                snippet = ""
                nxt = a.find_parent("tr")
                if nxt:
                    sib = nxt.find_next_sibling("tr")
                    if sib:
                        snippet = sib.get_text(" ", strip=True)
                if title and href:
                    results.append({"title": title, "url": href, "snippet": snippet})
        except Exception:
            pass
        return results

    def _scrape_portal(self, portal):
        results = []
        try:
            r = requests.get(portal["url"], headers=HEADERS, timeout=20)
            if not r or r.status_code != 200:
                return results
            soup = BeautifulSoup(r.text, "lxml")
            links = soup.select(portal["link_selector"])[:10]
            for a in links:
                title = a.get_text(strip=True)
                href  = a.get("href", "")
                if href and not href.startswith("http"):
                    href = urljoin(portal["base"], href)
                if title and href:
                    results.append({
                        "title":   title,
                        "url":     href,
                        "snippet": title,
                        "source":  portal["name"],
                    })
        except Exception:
            pass
        return results

    def run(self, progress_cb=None):
        raw = []
        for q in self.QUERIES:
            if progress_cb:
                progress_cb(f"Web search: {q[:60]}…")
            for h in self._search(q):
                country = _guess_country_from_text(h["title"] + " " + h["snippet"])
                raw.append({
                    "name":      h["title"],
                    "country":   country,
                    "org":       "",
                    "scope":     h["snippet"][:400],
                    "deadline":  "See source",
                    "cost":      _extract_cost(h["snippet"]),
                    "source":    "Web Search",
                    "url":       h["url"],
                    "published": "",
                })
            time.sleep(1.2)

        for portal in self.EXTRA_PORTALS:
            if progress_cb:
                progress_cb(f"Portal: {portal['name']}")
            for h in self._scrape_portal(portal):
                country = _guess_country_from_text(h["title"] + " " + h.get("snippet", ""))
                raw.append({
                    "name":      h["title"],
                    "country":   country,
                    "org":       "",
                    "scope":     h.get("snippet", h["title"])[:400],
                    "deadline":  "See source",
                    "cost":      "Not disclosed",
                    "source":    portal["name"],
                    "url":       h["url"],
                    "published": "",
                })
            time.sleep(1.0)
        return raw


# ═══════════════════════════════════════════════════════════
#  SECTION 3 – HELPERS
# ═══════════════════════════════════════════════════════════

def _extract_field(text, label):
    m = re.search(rf"{label}\s*[:\-–]\s*([^\n\|\.]{3,120})", text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def _extract_cost(text):
    for p in [
        r"(USD|EUR|GBP|US\$|€|£)\s*([\d,\.]+)\s*(billion|million|mn|bn|M|B)",
        r"([\d,\.]+)\s*(billion|million|mn|bn)\s*(USD|EUR|US\$)?",
        r"cost[:\s]*(USD|EUR|US\$)?\s*([\d,\.]+)\s*(million|billion|mn|bn|M|B)?",
    ]:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(0).strip()[:40]
    return "Not disclosed"

def _guess_country_from_text(text):
    for c in [
        "Nigeria","Uganda","Tanzania","Kenya","Egypt","Morocco",
        "Ethiopia","Zambia","Zimbabwe","Namibia","Mozambique",
        "Mauritania","Mali","Guinea","Ghana","Senegal","Congo",
        "Jordan","UAE","Saudi Arabia","Iraq","Syria","Israel",
        "Thailand","Nepal","Philippines","Vietnam","Indonesia",
        "Bangladesh","Sri Lanka","Myanmar","Kazakhstan",
        "Poland","Lithuania","Finland","Slovakia","Uruguay",
        "Paraguay","Ecuador","Brazil","Peru","Panama",
        "Colombia","United Kingdom","France","Germany",
        "Trinidad and Tobago",
    ]:
        if re.search(rf"\b{c}\b", text, re.IGNORECASE):
            return c
    return ""

def is_relevant(tender):
    text    = (tender.get("name","") + " " + tender.get("scope","")).lower()
    country = tender.get("country","").lower().strip()

    for excl in EXCLUDE_COUNTRIES:
        if excl in country:
            return False
    if not country or len(country) < 2:
        return False
    if not any(pos in text for pos in POSITIVE_KEYWORDS):
        return False
    if not any(kw in text for kw in CONSULTANCY_KEYWORDS):
        return False
    for neg in NEGATIVE_KEYWORDS:
        if neg in text:
            if sum(1 for p in POSITIVE_KEYWORDS if p in text) < 2:
                return False
    return True

def score_tender(tender):
    text    = (tender.get("name","") + " " + tender.get("scope","")).lower()
    country = tender.get("country","").lower()
    cost    = tender.get("cost","").lower()

    domain_hits  = sum(1 for p in POSITIVE_KEYWORDS if p in text)
    domain_score = min(50, domain_hits * 6)
    for kv in ["400 kv","500 kv","765 kv","hvdc","uhvdc","225 kv","330 kv"]:
        if kv in text:
            domain_score = min(50, domain_score + 8)
            break
    if any(k in text for k in ["consultancy","pmc","feasibility","advisory"]):
        domain_score = min(50, domain_score + 5)

    size_score = 10
    if any(x in cost for x in ["billion","bn"]):
        size_score = 20
    elif any(x in cost for x in ["million","mn","m"]):
        nums = re.findall(r"[\d]+", cost)
        if nums:
            val = int(nums[0])
            size_score = 20 if val >= 200 else (15 if val >= 50 else 10)
    elif cost in ["not disclosed","n/a",""]:
        size_score = 8

    strat_score   = 10 if any(sc in country for sc in STRATEGIC_COUNTRIES) else 0
    clarity_score = min(20, len(tender.get("scope","")) // 25)
    return min(100, domain_score + size_score + strat_score + clarity_score)

def get_risk(country, fsi_data):
    c = country.lower().strip()
    for key, val in fsi_data.items():
        if key in c or c in key:
            return val
    return ("Medium", None)

def get_recommendation(score, risk_category):
    if score > 75:
        return "Strongly Recommended"
    elif score >= 50:
        return "Consider"
    else:
        return "Not Recommended"

def deduplicate(tenders):
    seen, out = set(), []
    for t in tenders:
        key = (t["name"][:40].lower().strip(), t["country"].lower().strip())
        if key not in seen:
            seen.add(key)
            out.append(t)
    return out

def _parse_publish_date(date_str):
    if not date_str:
        return None
    date_str = date_str.strip()[:20]
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y",
                "%m/%d/%Y", "%d %b %Y", "%d %B %Y", "%B %d, %Y"):
        try:
            return datetime.datetime.strptime(date_str[:len(fmt)+2], fmt).date()
        except ValueError:
            continue
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", date_str)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass
    return None

def process_tenders(raw, fsi_data, cutoff_date=None):
    filtered = [t for t in raw if is_relevant(t)]

    if cutoff_date:
        date_filtered = []
        for t in filtered:
            pub = _parse_publish_date(t.get("published", ""))
            if pub is None or pub >= cutoff_date:
                date_filtered.append(t)
        filtered = date_filtered

    for t in filtered:
        risk_cat, fsi_score = get_risk(t["country"], fsi_data)
        t["score"]          = score_tender(t)
        t["risk_category"]  = risk_cat
        t["fsi_score"]      = fsi_score if fsi_score is not None else "N/A"
        t["risk_label"]     = RISK_LABEL[risk_cat]
        t["recommendation"] = get_recommendation(t["score"], risk_cat)

    deduped = deduplicate(filtered)
    risk_order = {"Low": 0, "Medium": 1, "High": 2}
    deduped.sort(key=lambda x: (-x["score"], risk_order.get(x["risk_category"], 1)))
    return deduped


# ═══════════════════════════════════════════════════════════
#  SECTION 4 – SEED / FALLBACK DATA
# ═══════════════════════════════════════════════════════════

def get_demo_tenders():
    return [
        {
            "name":      "Implementation Consultancy for 500 kV & 220 kV Double Circuit Transmission Works",
            "country":   "Egypt",
            "org":       "Egyptian Electricity Transmission Company (EETC)",
            "scope":     "PMC/supervision consultancy for 500 kV and 220 kV double-circuit transmission line EPC works. EBRD-facilitated procurement.",
            "deadline":  "Apr–May 2026",
            "cost":      "USD 10–30 M (est.)",
            "source":    "globaltransmission.info / ebrd.com",
            "url":       "https://globaltransmission.info/implementation-consultancy-for-high-voltage-transmission-works/",
            "published": "2026-01-15",
        },
        {
            "name":      "Feasibility Study – UAE–India Undersea HVDC Power Interconnector",
            "country":   "UAE",
            "org":       "Etihad Water and Electricity (EtihadWE)",
            "scope":     "Techno-economic feasibility study for proposed UAE–India HVDC undersea power interconnection. International consulting firms invited.",
            "deadline":  "Feb–Mar 2026",
            "cost":      "USD 3–8 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/feasibility-study-for-uae-india-undersea-power-interconnector/",
            "published": "2025-12-20",
        },
        {
            "name":      "Owner's Engineer / PMC – 400 kV Wobulenzi–Masaka–Mutukula Transmission Lines",
            "country":   "Uganda",
            "org":       "Uganda Electricity Transmission Company Limited (UETCL)",
            "scope":     "Project Management Consultancy (Owner's Engineer) for supervision of design, supply and installation of 400 kV Wobulenzi–Masaka 165 km and Masaka–Mutukula 92 km lines and associated substations.",
            "deadline":  "TBD (issued Jan 2026)",
            "cost":      "USD 5–15 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/plant-design-supply-and-installation-of-400-kv-transmission-lines-and-substations/",
            "published": "2026-01-10",
        },
        {
            "name":      "Feasibility Study & Detailed Design – 132 kV KETRACO Uprating",
            "country":   "Kenya",
            "org":       "Kenya Electricity Transmission Company Limited (KETRACO)",
            "scope":     "Feasibility study and detailed engineering design for uprating of existing 132 kV transmission lines to high-ampacity conductors (ACCC).",
            "deadline":  "Verify (posted Feb 2026)",
            "cost":      "USD 2–5 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/design-supply-and-installation-of-132-kv-line-3/",
            "published": "2026-02-05",
        },
        {
            "name":      "Technical Assistance – Grid Modernisation Advisory, EGAT Thailand",
            "country":   "Thailand",
            "org":       "Electricity Generating Authority of Thailand (EGAT)",
            "scope":     "Advisory and technical assistance for smart grid modernisation planning, 500 kV grid reinforcement study, and renewable energy integration feasibility for national grid.",
            "deadline":  "Jun 2026",
            "cost":      "USD 3–10 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/supply-of-500-kv-power-transformer-3/",
            "published": "2026-01-28",
        },
        {
            "name":      "PMC / Supervision Consultancy – 225 kV El Ghaira–Aleg Line, Mauritania",
            "country":   "Mauritania",
            "org":       "Societe Mauritanienne D'electricite (SOMELEC)",
            "scope":     "Project management consultancy and supervision of 225 kV El Ghaira–Aleg overhead transmission line construction and associated substation commissioning.",
            "deadline":  "Verify (posted Feb 2026)",
            "cost":      "USD 2–5 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/construction-of-225-kv-line-and-associated-substation/",
            "published": "2026-02-01",
        },
        {
            "name":      "Feasibility Study – 500/150 kV Substation, Uruguay (UTE)",
            "country":   "Uruguay",
            "org":       "National Administration of Electrical Power Plants and Transmissions (UTE)",
            "scope":     "Feasibility study, conceptual design and detailed design advisory for 500/150 kV substation facilities including equipment specification and bid preparation support.",
            "deadline":  "Verify (posted Mar 2026)",
            "cost":      "USD 2–6 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/construction-of-500-150-kv-substation-facilities/",
            "published": "2026-03-10",
        },
        {
            "name":      "Technical Consultant – 400/110 kV Substation Expansion, Poland (PSE)",
            "country":   "Poland",
            "org":       "Polskie Sieci Elektroenergetyczne SA (PSE)",
            "scope":     "Independent technical consultant / Owner's Engineer for expansion works at 400/220/110 kV Mikułowa station; grid connection advisory and supervision for renewable integration.",
            "deadline":  "Verify (posted May 2026)",
            "cost":      "Not disclosed",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/expansion-of-400-110-kv-substation/",
            "published": "2026-04-20",
        },
        {
            "name":      "Owner's Engineer – 132 kV Transmission Line, ZESCO Zambia",
            "country":   "Zambia",
            "org":       "Zambia Electricity Supply Corporation Limited (ZESCO)",
            "scope":     "Owner's Engineer / independent supervision consultancy for design, supply and construction of 132 kV overhead transmission line and associated substation works. World Bank / AfDB co-financed.",
            "deadline":  "May–Jun 2026",
            "cost":      "USD 3–8 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/construction-of-132-kv-transmission-line-11/",
            "published": "2026-02-18",
        },
        {
            "name":      "Pre-Feasibility & DPR – Regional Power Interconnection, West Africa (WAPP)",
            "country":   "Nigeria",
            "org":       "West African Power Pool (WAPP) / ECOWAS",
            "scope":     "Pre-feasibility and Detailed Project Report (DPR) for regional high-voltage power interconnection across West Africa countries. Funded by World Bank/AfDB.",
            "deadline":  "Verify (posted Jan 2026)",
            "cost":      "USD 5–12 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/",
            "published": "2026-01-05",
        },
        {
            "name":      "Transaction Advisory – PPP Power Transmission, Jordan",
            "country":   "Jordan",
            "org":       "National Electric Power Company (NEPCO)",
            "scope":     "Transaction advisory for structuring a PPP/private participation arrangement for grid expansion and 400 kV interconnection with neighbouring countries. IFC/World Bank facilitated.",
            "deadline":  "Verify (posted Mar 2026)",
            "cost":      "USD 2–6 M (est., consultancy)",
            "source":    "globaltransmission.info",
            "url":       "https://globaltransmission.info/",
            "published": "2026-03-01",
        },
    ]


# ═══════════════════════════════════════════════════════════
#  SECTION 5 – DATAFRAME & EXPORT
# ═══════════════════════════════════════════════════════════

def build_dataframe(tenders):
    rows = []
    for i, t in enumerate(tenders, 1):
        rec = t["recommendation"]
        if rec == "Strongly Recommended":
            remark = "High priority – pursue immediately."
        elif rec == "Consider":
            remark = "Evaluate bid capacity and local partner requirements."
        else:
            remark = "Low value or high-risk; review if pipeline thin."

        rows.append({
            "S.No.":          i,
            "Name of Work":   t["name"][:120],
            "Scope of Work":  t["scope"][:300],
            "Organisation":   t.get("org", "")[:100],
            "Source":         t["source"],
            "Country":        t["country"],
            "Deadline":       t["deadline"],
            "Est. Cost":      t["cost"],
            "Score":          t["score"],
            "FSI Score":      t.get("fsi_score", "N/A"),
            "Risk Category":  t["risk_category"],
            "Risk Label":     t["risk_label"],
            "Recommendation": rec,
            "Remarks":        remark,
            "Tender Link":    t.get("url", ""),
        })
    return pd.DataFrame(rows, columns=COLUMNS)


def export_excel_bytes(df):
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tender Intelligence"
    today = datetime.date.today().strftime("%d-%b-%Y")

    ncols   = len(COLUMNS)
    last_col = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = (
        f"POWERGRID IB — International Tender Intelligence Report  |  "
        f"Generated: {today}  |  FSI Ratings: fragilestatesindex.org"
    )
    ws["A1"].font      = Font(bold=True, size=13, color="FFFFFF", name="Arial")
    ws["A1"].fill      = PatternFill("solid", fgColor="003366")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = (
        "Sources: globaltransmission.info | tendersinfo.com | Open Web Search    |    "
        f"Filters: International only (ex-India)    |    FSI Data: {FSI_SOURCE_NOTE}"
    )
    ws["A2"].font      = Font(italic=True, size=9, color="FFFFFF", name="Arial")
    ws["A2"].fill      = PatternFill("solid", fgColor="1F4E79")
    ws["A2"].alignment = Alignment(horizontal="center")

    hdr_fill = PatternFill("solid", fgColor="2E75B6")
    hdr_font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    thin     = Side(style="thin", color="CCCCCC")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=3, column=col_idx, value=col_name)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[3].height = 30

    rec_colors  = {
        "Strongly Recommended": "E2EFDA",
        "Consider":             "FFF2CC",
        "Not Recommended":      "FCE4D6",
    }
    risk_colors = {
        "Low":    "C6EFCE",
        "Medium": "FFEB9C",
        "High":   "FFC7CE",
    }
    col_map = {name: idx+1 for idx, name in enumerate(COLUMNS)}

    for row_idx, row in df.iterrows():
        excel_row = row_idx + 4
        rec       = row["Recommendation"]
        row_fill  = PatternFill("solid", fgColor=rec_colors.get(rec, "FFFFFF"))

        for col_name in COLUMNS:
            col_idx = col_map[col_name]
            val     = row[col_name]
            cell    = ws.cell(row=excel_row, column=col_idx)
            cell.border    = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font      = Font(size=9, name="Arial")
            cell.fill      = row_fill

            if col_name == "Tender Link":
                url = str(val).strip()
                if url and url.startswith("http"):
                    cell.value     = "View Tender ↗"
                    cell.hyperlink = url
                    cell.font      = Font(size=9, name="Arial", color="0563C1", underline="single", bold=True)
                else:
                    cell.value = url
                continue
            if col_name == "FSI Score":
                try:
                    cell.value = float(val) if val != "N/A" else "N/A"
                except (ValueError, TypeError):
                    cell.value = val
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font      = Font(size=9, name="Arial", bold=True)
                continue
            if col_name == "Risk Category":
                cell.value = val
                cell.fill  = PatternFill("solid", fgColor=risk_colors.get(str(val), "FFFFFF"))
                cell.font  = Font(size=9, name="Arial", bold=True)
                continue
            if col_name == "Score":
                cell.value     = val
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font      = Font(bold=True, size=9, name="Arial")
                continue
            if col_name == "S.No.":
                cell.value     = val
                cell.alignment = Alignment(horizontal="center", vertical="top")
                continue
            cell.value = val

    widths = [5, 32, 45, 28, 18, 12, 14, 18, 7, 9, 12, 22, 22, 30, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"A3:{last_col}{3 + len(df)}"

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "POWERGRID IB — Tender Intelligence Summary"
    ws2["A1"].font = Font(bold=True, size=13, name="Arial", color="003366")
    ws2["A2"] = f"Generated: {today}  |  FSI Source: {FSI_SOURCE_NOTE}"
    ws2.append([])
    ws2.append(["Recommendation", "Count"])
    for rec, cnt in df["Recommendation"].value_counts().items():
        ws2.append([rec, cnt])
    ws2.append([])
    ws2.append(["Risk Category", "Count", "FSI Score (avg)"])
    for risk in ["Low", "Medium", "High"]:
        subset = df[df["Risk Category"] == risk]
        if len(subset):
            fsi_vals = pd.to_numeric(subset["FSI Score"], errors="coerce")
            avg_fsi  = round(fsi_vals.mean(), 1) if not fsi_vals.isna().all() else "N/A"
            ws2.append([risk, len(subset), avg_fsi])
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_pdf_bytes(df):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=15*mm, bottomMargin=10*mm,
    )
    styles    = getSampleStyleSheet()
    title_sty = ParagraphStyle("T", parent=styles["Normal"],
                                fontSize=16, fontName="Helvetica-Bold",
                                textColor=colors.HexColor("#003366"), spaceAfter=4)
    sub_sty   = ParagraphStyle("S", parent=styles["Normal"],
                                fontSize=9, textColor=colors.HexColor("#555555"), spaceAfter=12)
    cell_sty  = ParagraphStyle("C", parent=styles["Normal"], fontSize=7.5, leading=10)
    bold_cell = ParagraphStyle("B", parent=styles["Normal"],
                                fontSize=7.5, fontName="Helvetica-Bold", leading=10)

    story = []
    today = datetime.date.today().strftime("%d %B %Y")
    story.append(Paragraph("POWERGRID India — International Business Department", title_sty))
    story.append(Paragraph(
        f"Tender Intelligence Report  |  Generated: {today}  |  "
        f"Sources: globaltransmission.info | tendersinfo.com | Open Web Search  |  "
        f"FSI: {FSI_SOURCE_NOTE}", sub_sty
    ))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#003366")))
    story.append(Spacer(1, 6*mm))

    total = len(df)
    sr    = (df["Recommendation"] == "Strongly Recommended").sum()
    co    = (df["Recommendation"] == "Consider").sum()
    nr    = (df["Recommendation"] == "Not Recommended").sum()
    stats = Table(
        [["Total Tenders", "Strongly Recommended", "Consider", "Not Recommended"],
         [str(total), str(sr), str(co), str(nr)]],
        colWidths=[60*mm]*4,
    )
    stats.setStyle(TableStyle([
        ("BACKGROUND",  (0,0),(-1,0), colors.HexColor("#003366")),
        ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
        ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0),(-1,0), 9),
        ("FONTSIZE",    (0,1),(-1,1), 18),
        ("FONTNAME",    (0,1),(-1,1), "Helvetica-Bold"),
        ("ALIGN",       (0,0),(-1,-1),"CENTER"),
        ("BOX",         (0,0),(-1,-1),0.5,colors.HexColor("#AAAAAA")),
        ("INNERGRID",   (0,0),(-1,-1),0.5,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",  (0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(stats)
    story.append(Spacer(1, 6*mm))

    pdf_cols = [
        "S.No.", "Name of Work", "Scope of Work", "Source",
        "Country", "Deadline", "Est. Cost", "Score",
        "FSI Score", "Risk Category", "Recommendation", "Remarks",
    ]
    col_widths = [8*mm,45*mm,52*mm,20*mm,16*mm,16*mm,20*mm,
                  10*mm,12*mm,18*mm,28*mm,32*mm]

    header_row = [Paragraph(f"<b>{c}</b>", cell_sty) for c in pdf_cols]
    table_data = [header_row]

    rec_pdf  = {"Strongly Recommended": colors.HexColor("#E2EFDA"),
                "Consider":             colors.HexColor("#FFF2CC"),
                "Not Recommended":      colors.HexColor("#FCE4D6")}
    risk_pdf = {"Low":    colors.HexColor("#C6EFCE"),
                "Medium": colors.HexColor("#FFEB9C"),
                "High":   colors.HexColor("#FFC7CE")}

    style_cmds = [
        ("BACKGROUND",  (0,0),(-1,0), colors.HexColor("#2E75B6")),
        ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
        ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
        ("ALIGN",       (0,0),(0,-1), "CENTER"),
        ("VALIGN",      (0,0),(-1,-1),"TOP"),
        ("FONTSIZE",    (0,0),(-1,-1),7.5),
        ("BOX",         (0,0),(-1,-1),0.5,colors.HexColor("#AAAAAA")),
        ("INNERGRID",   (0,0),(-1,-1),0.25,colors.HexColor("#DDDDDD")),
        ("TOPPADDING",  (0,0),(-1,-1),3),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING", (0,0),(-1,-1),3),
        ("RIGHTPADDING",(0,0),(-1,-1),3),
    ]

    for row_idx, row in df.iterrows():
        rec  = row["Recommendation"]
        data_row = []
        for col in pdf_cols:
            val = str(row[col]) if row[col] is not None else ""
            if col == "Name of Work":
                data_row.append(Paragraph(val[:100], bold_cell))
            else:
                data_row.append(Paragraph(val[:250], cell_sty))
        table_data.append(data_row)

        er = row_idx + 1
        style_cmds.append(("BACKGROUND", (0,er),(-1,er), rec_pdf.get(rec, colors.white)))
        ri = pdf_cols.index("Risk Category")
        style_cmds.append(("BACKGROUND", (ri,er),(ri,er), risk_pdf.get(row["Risk Category"], colors.white)))

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#AAAAAA")))
    footer_sty = ParagraphStyle("F", parent=styles["Normal"],
                                 fontSize=7, textColor=colors.HexColor("#888888"), spaceBefore=4)
    story.append(Paragraph(
        "CONFIDENTIAL — For internal use of POWERGRID IB Department only.  "
        "Scores and risk ratings are indicative; verify all deadlines with issuing authorities.  "
        f"Country risk based on Fragile States Index (fragilestatesindex.org).  "
        "Tender hyperlinks available in the companion Excel workbook.",
        footer_sty,
    ))
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════
#  SECTION 6 – STREAMLIT UI
# ═══════════════════════════════════════════════════════════

def color_recommendation(val):
    colors_map = {
        "Strongly Recommended": "background-color: #E2EFDA; color: #375623; font-weight: bold",
        "Consider":             "background-color: #FFF2CC; color: #7D6608; font-weight: bold",
        "Not Recommended":      "background-color: #FCE4D6; color: #922B21; font-weight: bold",
    }
    return colors_map.get(val, "")

def color_risk(val):
    colors_map = {
        "Low":    "background-color: #C6EFCE; color: #375623; font-weight: bold",
        "Medium": "background-color: #FFEB9C; color: #7D6608; font-weight: bold",
        "High":   "background-color: #FFC7CE; color: #922B21; font-weight: bold",
    }
    return colors_map.get(val, "")


def main():
    # ── CSS ─────────────────────────────────────────────────
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #003366 0%, #1F4E79 100%);
        padding: 20px 30px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .main-header h1 { color: #FFFFFF; margin: 0; font-size: 1.8em; }
    .main-header p  { color: #BDD7EE; margin: 4px 0 0 0; font-size: 0.95em; }
    .metric-card {
        background: #F0F7FF;
        border-left: 4px solid #2E75B6;
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    .metric-card.green { border-left-color: #70AD47; background: #F0FFF0; }
    .metric-card.orange{ border-left-color: #FFC000; background: #FFFBF0; }
    .metric-card.red   { border-left-color: #FF0000; background: #FFF0F0; }
    .stProgress > div > div { background-color: #2E75B6; }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ──────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <h1>⚡ POWERGRID IB — Tender Intelligence Analyzer</h1>
        <p>International Consultancy / PMC / Feasibility / Advisory Tenders · v3.0 (Streamlit Cloud)</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ─────────────────────────────────────────────
    with st.sidebar:
        st.image("https://upload.wikimedia.org/wikipedia/en/thumb/b/b6/Power_Grid_Corporation_of_India.svg/200px-Power_Grid_Corporation_of_India.svg.png", width=140)
        st.markdown("### ⚙️ Analysis Settings")

        cutoff_date = st.date_input(
            "Publish Date Cutoff",
            value=None,
            help="Tenders published before this date will be excluded. Leave blank to include all.",
        )

        st.markdown("---")
        st.markdown("### 🔍 Data Sources")
        use_gt   = st.checkbox("globaltransmission.info", value=True)
        use_ti   = st.checkbox("tendersinfo.com",         value=True)
        use_web  = st.checkbox("Web Search (DuckDuckGo + Portals)", value=True)
        use_seed = st.checkbox("Seed / Verified Data",   value=True)

        st.markdown("---")
        st.markdown("### 🔐 Credentials")
        st.info("Add credentials in Streamlit Secrets for authenticated scraping. See deployment guide.", icon="ℹ️")
        creds = get_credentials()
        if creds["globaltransmission"]["username"]:
            st.success("globaltransmission.info ✓")
        else:
            st.warning("globaltransmission.info — unauthenticated")
        if creds["tendersinfo"]["username"]:
            st.success("tendersinfo.com ✓")
        else:
            st.warning("tendersinfo.com — unauthenticated")

        st.markdown("---")
        run_btn = st.button("🚀 Run Analysis", type="primary", use_container_width=True)

    # ── Main content ─────────────────────────────────────────
    if not run_btn:
        st.markdown("""
        ### Welcome
        This tool scrapes and analyzes international power transmission **consultancy / PMC / feasibility** tenders
        for POWERGRID India's International Business Department.

        **How to use:**
        1. Set a publish date cutoff in the sidebar (optional)
        2. Choose your data sources
        3. Click **Run Analysis**
        4. Download the colour-coded Excel or PDF report

        **Sources scraped:**
        - globaltransmission.info (authenticated)
        - tendersinfo.com (authenticated)
        - DuckDuckGo web search
        - UNGM, ADB, AfDB, EBRD, World Bank portals
        - Verified seed tenders

        > ⚠️ Only **consultancy / PMC / design / feasibility** tenders are included.
        > POWERGRID IB does not bid on EPC/construction contracts.
        """)
        return

    # ── Run pipeline ─────────────────────────────────────────
    status_box = st.empty()
    progress   = st.progress(0)
    log_box    = st.expander("📋 Scraping Log", expanded=False)
    log_lines  = []

    def update_status(msg):
        status_box.info(f"⏳ {msg}")
        log_lines.append(msg)
        with log_box:
            st.text("\n".join(log_lines[-30:]))

    raw_tenders = []

    # FSI
    update_status("Fetching live FSI scores…")
    progress.progress(5)
    fsi_data, fsi_msg = fetch_live_fsi_scores()
    log_lines.append(fsi_msg)

    # Logins
    gt_session = None
    ti_session = None
    if use_gt or use_ti:
        update_status("Logging in to data sources…")
        progress.progress(10)
        if use_gt:
            gt_session = login_globaltransmission(
                creds["globaltransmission"]["username"],
                creds["globaltransmission"]["password"],
            )
        if use_ti:
            ti_session = login_tendersinfo(
                creds["tendersinfo"]["username"],
                creds["tendersinfo"]["password"],
            )

    # globaltransmission.info
    if use_gt:
        update_status("Scraping globaltransmission.info…")
        progress.progress(20)
        try:
            gt_scraper = GlobalTransmissionScraper(session=gt_session)
            gt_results = gt_scraper.run(progress_cb=update_status)
            raw_tenders.extend(gt_results)
            log_lines.append(f"  globaltransmission.info: {len(gt_results)} raw entries")
        except Exception as e:
            log_lines.append(f"  globaltransmission.info error: {e}")

    # tendersinfo.com
    if use_ti:
        update_status("Scraping tendersinfo.com…")
        progress.progress(45)
        try:
            ti_scraper = TendersInfoScraper(session=ti_session)
            ti_results = ti_scraper.run(progress_cb=update_status)
            raw_tenders.extend(ti_results)
            log_lines.append(f"  tendersinfo.com: {len(ti_results)} raw entries")
        except Exception as e:
            log_lines.append(f"  tendersinfo.com error: {e}")

    # Web search
    if use_web:
        update_status("Running web searches and portal scrapes…")
        progress.progress(60)
        try:
            ws_scraper = WebSearchScraper()
            ws_results = ws_scraper.run(progress_cb=update_status)
            raw_tenders.extend(ws_results)
            log_lines.append(f"  Web search + portals: {len(ws_results)} raw entries")
        except Exception as e:
            log_lines.append(f"  Web search error: {e}")

    # Seed data
    if use_seed:
        seed = get_demo_tenders()
        raw_tenders.extend(seed)
        log_lines.append(f"  Seed data: {len(seed)} entries")

    log_lines.append(f"Total raw pool: {len(raw_tenders)}")

    # Process
    update_status("Filtering, scoring and deduplicating…")
    progress.progress(85)
    tenders = process_tenders(raw_tenders, fsi_data, cutoff_date=cutoff_date)

    progress.progress(95)

    if not tenders:
        status_box.error("No relevant consultancy tenders found after filtering. Try adjusting filters.")
        st.stop()

    df = build_dataframe(tenders)
    progress.progress(100)
    status_box.success(f"✅ Analysis complete — {len(df)} tenders found")

    # ── Summary metrics ──────────────────────────────────────
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Tenders", len(df))
    with col2:
        sr = (df["Recommendation"] == "Strongly Recommended").sum()
        st.metric("Strongly Recommended", sr, delta=None)
    with col3:
        co = (df["Recommendation"] == "Consider").sum()
        st.metric("Consider", co)
    with col4:
        nr = (df["Recommendation"] == "Not Recommended").sum()
        st.metric("Not Recommended", nr)

    # ── Filters ─────────────────────────────────────────────
    st.markdown("### 🔎 Filter Results")
    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        rec_filter = st.multiselect(
            "Recommendation",
            options=["Strongly Recommended", "Consider", "Not Recommended"],
            default=["Strongly Recommended", "Consider"],
        )
    with fcol2:
        risk_filter = st.multiselect(
            "Risk Category",
            options=["Low", "Medium", "High"],
            default=["Low", "Medium", "High"],
        )
    with fcol3:
        countries = sorted(df["Country"].unique().tolist())
        country_filter = st.multiselect("Country", options=countries, default=countries)

    mask = (
        df["Recommendation"].isin(rec_filter) &
        df["Risk Category"].isin(risk_filter) &
        df["Country"].isin(country_filter)
    )
    df_filtered = df[mask].reset_index(drop=True)
    df_filtered["S.No."] = range(1, len(df_filtered) + 1)

    st.markdown(f"**Showing {len(df_filtered)} of {len(df)} tenders**")

    # ── Table display ────────────────────────────────────────
    display_cols = ["S.No.", "Name of Work", "Country", "Organisation",
                    "Deadline", "Est. Cost", "Score", "FSI Score",
                    "Risk Category", "Recommendation", "Tender Link"]

    def make_clickable(url):
        if url and str(url).startswith("http"):
            return f'<a href="{url}" target="_blank">View ↗</a>'
        return url

    df_display = df_filtered[display_cols].copy()
    df_display["Tender Link"] = df_display["Tender Link"].apply(make_clickable)

    styled = (
        df_filtered[display_cols[:-1]]
        .style
        .applymap(color_recommendation, subset=["Recommendation"])
        .applymap(color_risk, subset=["Risk Category"])
        .set_properties(**{"font-size": "12px"})
        .format({"Score": "{:.0f}", "FSI Score": lambda x: f"{x:.1f}" if isinstance(x, float) else x})
    )
    st.write(styled.to_html(escape=False), unsafe_allow_html=True)

    # Tender links as separate expander
    with st.expander("🔗 Tender Links"):
        for _, row in df_filtered.iterrows():
            url = row["Tender Link"]
            if url and str(url).startswith("http"):
                st.markdown(f"**{row['S.No.']}. {row['Name of Work'][:80]}** — [{url}]({url})")
            else:
                st.markdown(f"**{row['S.No.']}. {row['Name of Work'][:80]}** — No link")

    # ── Downloads ────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 Download Reports")
    dcol1, dcol2 = st.columns(2)

    with dcol1:
        with st.spinner("Generating Excel…"):
            xlsx_bytes = export_excel_bytes(df_filtered)
        fname_xlsx = f"POWERGRID_IB_Tender_Report_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️ Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=fname_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dcol2:
        with st.spinner("Generating PDF…"):
            pdf_bytes = export_pdf_bytes(df_filtered)
        fname_pdf = f"POWERGRID_IB_Tender_Report_{datetime.date.today().strftime('%Y%m%d')}.pdf"
        st.download_button(
            label="⬇️ Download PDF",
            data=pdf_bytes,
            file_name=fname_pdf,
            mime="application/pdf",
            use_container_width=True,
        )

    st.caption(
        f"Report generated: {datetime.date.today().strftime('%d %B %Y')}  |  "
        f"FSI Data: fragilestatesindex.org  |  "
        "CONFIDENTIAL — Internal use only"
    )


if __name__ == "__main__":
    main()
