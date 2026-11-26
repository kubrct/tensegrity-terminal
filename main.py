# -*- coding: utf-8 -*-
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import asynccontextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, parseaddr
import asyncio
import base64
import hashlib
import html
import io
import json
import logging
import math
import os
import random
import re
import smtplib
import ssl
from urllib.parse import quote, urlparse
import urllib.request
import xml.etree.ElementTree as ET

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as PdfImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    import resend  # type: ignore
except ImportError:
    resend = None

import yfinance as yf

try:
    from zoneinfo import ZoneInfo
except ImportError:
    ZoneInfo = None

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("tensegrity_terminal")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
CACHE_DIR = os.path.join(BASE_DIR, "data_cache")
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(CACHE_DIR, exist_ok=True)
PARENT_DIR = os.path.dirname(BASE_DIR)

BLACK_LOGO_CANDIDATES = [
    os.path.join(STATIC_DIR, "logo-tt-black_2.png"),
    os.path.join(BASE_DIR, "logo-tt-black_2.png"),
    os.path.join(STATIC_DIR, "logo-tt-black.png"),
    os.path.join(BASE_DIR, "logo-tt-black.png"),
    os.path.join(PARENT_DIR, "static", "logo-tt-black.png"),
    os.path.join(PARENT_DIR, "logo-tt-black.png"),
    os.path.join(STATIC_DIR, "logo-tt-nb.png"),
    os.path.join(BASE_DIR, "logo-tt-nb.png"),
    os.path.join(STATIC_DIR, "logo.png"),
]
WHITE_LOGO_CANDIDATES = [
    os.path.join(STATIC_DIR, "logo-tt-white.png"),
    os.path.join(BASE_DIR, "logo-tt-white.png"),
    os.path.join(PARENT_DIR, "static", "logo-tt-white.png"),
    os.path.join(PARENT_DIR, "logo-tt-white.png"),
]

SUBSCRIBERS_FILE = os.path.join(BASE_DIR, "subscribers.json")
SCHEDULER_STATE_FILE = os.path.join(BASE_DIR, "scheduler_state.json")
ENV_FILE_PATH = os.path.join(BASE_DIR, ".env")

SENDER_EMAIL = os.getenv("SENDER_EMAIL", "kubenkamatyas@gmail.com").strip()
RAW_SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "yokf zefs doxl ujci")
SENDER_PASSWORD = "".join(RAW_SENDER_PASSWORD.split())
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "").strip()
RESEND_FROM_EMAIL = os.getenv("RESEND_FROM_EMAIL", "onboarding@resend.dev").strip()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
FMP_API_KEY = os.getenv("FMP_API_KEY", "").strip()
SEC_USER_AGENT = os.getenv("SEC_USER_AGENT", "TensegrityTerminal Research@tensegrity-terminal.local").strip()
CRON_SECRET = os.getenv("CRON_SECRET", "").strip()
ENV_SUBSCRIBERS = os.getenv("SUBSCRIBERS_LIST", "").strip()

FOREX_PAIRS = {
    "EUR/CZK": "EURCZK=X",
    "USD/CZK": "USDCZK=X",
    "GBP/CZK": "GBPCZK=X",
}
FOREX_FALLBACK_RATES = {"EUR/CZK": 24.45, "USD/CZK": 22.30, "GBP/CZK": 29.18}


# ==============================================================================
# 1. POMOCNÉ FUNKCE, LOGA, PÍSMA A FORMÁTOVÁNÍ
# ==============================================================================

def resolve_brand_logo_path(prefer_white: bool = False) -> str | None:
    search_order = (WHITE_LOGO_CANDIDATES + BLACK_LOGO_CANDIDATES) if prefer_white else BLACK_LOGO_CANDIDATES
    for candidate in search_order:
        if os.path.exists(candidate):
            return candidate
    return None


def register_unicode_pdf_fonts() -> tuple[str, str]:
    font_candidates = [
        ("DejaVuSans", "DejaVuSans-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("LiberationSans", "LiberationSans-Bold", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        ("ArialUnicode", "ArialUnicode-Bold", "C:\\Windows\\Fonts\\arial.ttf", "C:\\Windows\\Fonts\\arialbd.ttf"),
        ("SegoeUI", "SegoeUI-Bold", "C:\\Windows\\Fonts\\segoeui.ttf", "C:\\Windows\\Fonts\\segoeuib.ttf"),
        ("HelveticaNeue", "HelveticaNeue-Bold", "/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ]
    for reg_name, bold_name, reg_path, bold_path in font_candidates:
        if os.path.exists(reg_path):
            try:
                pdfmetrics.registerFont(TTFont(reg_name, reg_path))
                if os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                    return reg_name, bold_name
                return reg_name, reg_name
            except Exception as err:
                logger.warning("Nelze zaregistrovat písmo %s: %s", reg_path, err)
    return "Helvetica", "Helvetica-Bold"


PDF_FONT_REGULAR, PDF_FONT_BOLD = register_unicode_pdf_fonts()


def get_prague_now() -> datetime:
    if ZoneInfo is not None:
        try:
            return datetime.now(ZoneInfo("Europe/Prague"))
        except Exception:
            pass
    return datetime.now(timezone(timedelta(hours=1)))


def safe_float(value, default: float | None = 0.0) -> float | None:
    try:
        if value is None or isinstance(value, bool):
            return default
        num = float(value)
        if math.isnan(num) or math.isinf(num):
            return default
        return num
    except (TypeError, ValueError):
        return default


def extract_statement_series(df, row_names: list[str], max_years: int = 4, fx_mult: float = 1.0) -> list[float]:
    if df is None or getattr(df, "empty", True):
        return []
    for name in row_names:
        if name in df.index:
            series = df.loc[name].dropna()
            values = []
            for v in series.iloc[:max_years].tolist():
                val = safe_float(v, None)
                if val is not None:
                    values.append(val * fx_mult)
            if values:
                return values
    return []


def _format_cz_percent(val: float, decimals: int = 2) -> str:
    return f"{val:.{decimals}f}".replace(".", ",") + " %"


def save_gemini_key_to_env(new_key: str):
    global GEMINI_API_KEY
    clean_key = new_key.strip()
    GEMINI_API_KEY = clean_key
    os.environ["GEMINI_API_KEY"] = clean_key
    try:
        lines: list[str] = []
        if os.path.exists(ENV_FILE_PATH):
            with open(ENV_FILE_PATH, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
        updated = False
        new_lines = []
        for line in lines:
            if line.strip().startswith("GEMINI_API_KEY="):
                new_lines.append(f"GEMINI_API_KEY={clean_key}")
                updated = True
            else:
                new_lines.append(line)
        if not updated:
            new_lines.append(f"GEMINI_API_KEY={clean_key}")
        with open(ENV_FILE_PATH, "w", encoding="utf-8") as f:
            f.write("\n".join(new_lines) + "\n")
    except Exception as err:
        logger.warning("Klíč byl aktivován v paměti, ale zápis do .env selhal: %s", err)


# ==============================================================================
# 2. VYHLEDÁVÁNÍ SPOLEČNOSTÍ (VČETNĚ NOVÝCH IPO JAKO SPACEX / ANTHROPIC), FX A LOGA
# ==============================================================================

KNOWN_COMPANY_ALIASES = {
    "čez": "CEZ.PR",
    "cez": "CEZ.PR",
    "skupina čez": "CEZ.PR",
    "kofola": "KOFOL.PR",
    "komerční banka": "KOMB.PR",
    "komercni banka": "KOMB.PR",
    "moneta": "MONET.PR",
    "erste": "ERBAG.PR",
    "colt": "CZG.PR",
    "siemens": "SIE.DE",
    "nvidia": "NVDA",
    "apple": "AAPL",
    "microsoft": "MSFT",
    "google": "GOOGL",
    "alphabet": "GOOGL",
    "amazon": "AMZN",
    "meta": "META",
    "tesla": "TSLA",
    "netflix": "NFLX",
    "berkshire": "BRK-B",
    "berkshire hathaway": "BRK-B",
    "jpmorgan": "JPM",
    "visa": "V",
    "mastercard": "MA",
    "coca-cola": "KO",
    "coca cola": "KO",
    "pepsi": "PEP",
    "pepsico": "PEP",
    "mcdonald": "MCD",
    "mcdonalds": "MCD",
    "intel": "INTC",
    "amd": "AMD",
    "broadcom": "AVGO",
    "tsmc": "TSM",
    "asml": "ASML",
    "sap": "SAP",
    "volkswagen": "VOW3.DE",
    "allianz": "ALV.DE",
    "deutsche bank": "DBK.DE",
    "shell": "SHEL",
    "exxon": "XOM",
    "chevron": "CVX",
}

SPELLING_NORMALIZATIONS = {
    "entropic": "Anthropic",
    "antropic": "Anthropic",
    "space x": "SpaceX",
    "spacex": "SpaceX",
}


def _yahoo_search_equity_symbol(query_str: str) -> str | None:
    try:
        search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={quote(query_str)}&quotesCount=10&newsCount=0"
        req = urllib.request.Request(
            search_url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=7) as resp:
            payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
        quotes = payload.get("quotes") or []
        q_norm = query_str.lower().strip()

        # 1. Hledáme přímou shodu v názvu společnosti (EQUITY), nikoli ETF fondy typu SPCX
        for q in quotes:
            if q.get("quoteType") == "EQUITY" and q.get("symbol"):
                long_n = str(q.get("longname") or q.get("shortname") or "").lower()
                if q_norm in long_n or long_n.startswith(q_norm):
                    return str(q["symbol"]).upper()

        # 2. Jakákoliv veřejně obchodovaná akcie (EQUITY)
        for q in quotes:
            if q.get("quoteType") == "EQUITY" and q.get("symbol"):
                return str(q["symbol"]).upper()

        # 3. Teprve pokud neexistuje EQUITY, zkusíme ETF
        for q in quotes:
            if q.get("quoteType") == "ETF" and q.get("symbol"):
                return str(q["symbol"]).upper()
    except Exception as err:
        logger.warning("Vyhledání společnosti '%s' na Yahoo Finance selhalo: %s", query_str, err)
    return None


def _sec_search_ticker_by_name(query_str: str) -> str | None:
    """Vyhledá oficiální ticker přímo v registru SEC EDGAR podle názvu firmy (např. Space Exploration Technologies / Anthropic)."""
    q_low = query_str.lower().strip()
    search_terms = [q_low]
    if "spacex" in q_low:
        search_terms.append("space exploration technologies")
    try:
        cik_map = _load_sec_ticker_to_cik_map()
        for tkr, (_, title) in cik_map.items():
            t_low = title.lower()
            if any(term in t_low for term in search_terms):
                return tkr
    except Exception:
        pass
    return None


def resolve_company_ticker(query: str) -> str:
    raw = (query or "").strip()
    if not raw:
        return "NVDA"

    lower_q = raw.lower()
    if lower_q in KNOWN_COMPANY_ALIASES:
        return KNOWN_COMPANY_ALIASES[lower_q]

    normalized_name = SPELLING_NORMALIZATIONS.get(lower_q, raw)

    # Pokud uživatel zadal přímo krátký burzovní ticker (např. NVDA, CEZ.PR, BRK-B) a není to název firmy jako SpaceX/Anthropic
    if (
        lower_q not in SPELLING_NORMALIZATIONS
        and normalized_name.lower() not in ("anthropic", "spacex", "openai", "stripe", "databricks")
        and re.fullmatch(r"[A-Za-z0-9\.\-\^=]{1,12}", raw)
        and " " not in raw
    ):
        return raw.upper()

    # Dynamické vyhledání přes Yahoo Finance (řeší i čerstvá IPO jako SpaceX nebo Anthropic)
    found_yf = _yahoo_search_equity_symbol(normalized_name)
    if found_yf:
        return found_yf

    # Záložní vyhledání přímo v SEC EDGAR podle názvu emitenta
    found_sec = _sec_search_ticker_by_name(normalized_name)
    if found_sec:
        return found_sec

    if normalized_name.lower() in ("spacex", "anthropic"):
        raise ValueError(
            f"Společnost '{normalized_name}' nebyla v aktuálním feedu nalezena jako veřejně obchodovaný titul s dostupnými výkazy "
            "(může jít o soukromou firmu před IPO nebo pod jiným burzovním symbolem — zadejte prosím její přesný burzovní ticker)."
        )

    return raw.upper()


_FX_PAIR_CACHE: dict[str, tuple[datetime, float]] = {}


def get_fx_conversion_rate(from_curr: str, to_curr: str) -> float:
    f_c = (from_curr or "USD").strip().upper()
    t_c = (to_curr or "USD").strip().upper()
    if not f_c or not t_c or f_c == t_c:
        return 1.0

    scale = 1.0
    if f_c in ("GBP", "GBX") and from_curr == "GBp":
        f_c = "GBP"
        scale *= 0.01
    if t_c in ("GBP", "GBX") and to_curr == "GBp":
        t_c = "GBP"
        scale *= 100.0
    if f_c == t_c:
        return scale

    cache_key = f"{f_c}->{t_c}"
    now = datetime.now(timezone.utc)
    if cache_key in _FX_PAIR_CACHE and now < _FX_PAIR_CACHE[cache_key][0]:
        return _FX_PAIR_CACHE[cache_key][1] * scale

    pair_symbol = f"{f_c}{t_c}=X"
    try:
        hist = yf.Ticker(pair_symbol).history(period="5d", interval="1d")["Close"].dropna()
        if not hist.empty:
            rate = float(hist.iloc[-1])
            if rate > 0:
                _FX_PAIR_CACHE[cache_key] = (now + timedelta(minutes=30), rate)
                return rate * scale
    except Exception:
        pass

    inv_symbol = f"{t_c}{f_c}=X"
    try:
        hist_inv = yf.Ticker(inv_symbol).history(period="5d", interval="1d")["Close"].dropna()
        if not hist_inv.empty:
            inv_rate = float(hist_inv.iloc[-1])
            if inv_rate > 0:
                rate = 1.0 / inv_rate
                _FX_PAIR_CACHE[cache_key] = (now + timedelta(minutes=30), rate)
                return rate * scale
    except Exception:
        pass

    static_usd = {
        "USD": 1.0, "EUR": 1.08, "GBP": 1.27, "CZK": 0.044, "CHF": 1.13,
        "JPY": 0.0066, "CAD": 0.72, "AUD": 0.65, "SEK": 0.094, "NOK": 0.092,
        "DKK": 0.145, "PLN": 0.25, "TWD": 0.031, "HKD": 0.128, "KRW": 0.00072,
    }
    if f_c in static_usd and t_c in static_usd:
        fallback = static_usd[f_c] / static_usd[t_c]
        return fallback * scale

    return 1.0 * scale


TICKER_DOMAIN_FALLBACKS = {
    "NVDA": "nvidia.com",
    "AAPL": "apple.com",
    "MSFT": "microsoft.com",
    "GOOGL": "abc.xyz",
    "GOOG": "abc.xyz",
    "AMZN": "amazon.com",
    "META": "meta.com",
    "TSLA": "tesla.com",
    "NFLX": "netflix.com",
    "AMD": "amd.com",
    "INTC": "intel.com",
    "AVGO": "broadcom.com",
    "TSM": "tsmc.com",
    "ASML": "asml.com",
    "SAP": "sap.com",
    "SIE.DE": "siemens.com",
    "CEZ.PR": "cez.cz",
    "KOFOL.PR": "kofola.cz",
    "KOMB.PR": "kb.cz",
    "MONET.PR": "moneta.cz",
    "ERBAG.PR": "erstegroup.com",
    "CZG.PR": "coltczgroup.com",
    "JPM": "jpmorganchase.com",
    "V": "visa.com",
    "MA": "mastercard.com",
    "KO": "coca-colacompany.com",
    "PEP": "pepsico.com",
    "MCD": "mcdonalds.com",
    "BRK-B": "berkshirehathaway.com",
}


def extract_company_logo_urls(info: dict, ticker: str = "") -> list[str]:
    domain = ""
    website = str(info.get("website") or "").strip()
    if website:
        if not website.startswith(("http://", "https://")):
            website = "https://" + website
        try:
            parsed_host = urlparse(website).netloc.lower()
            if parsed_host.startswith("www."):
                parsed_host = parsed_host[4:]
            if parsed_host:
                domain = parsed_host
        except Exception:
            pass

    if not domain and ticker:
        domain = TICKER_DOMAIN_FALLBACKS.get(ticker.upper(), "")

    if not domain:
        return []

    return [
        f"https://www.google.com/s2/favicons?domain={quote(domain)}&sz=128",
        f"https://icons.duckduckgo.com/ip3/{quote(domain)}.ico",
        f"https://icon.horse/icon/{quote(domain)}",
    ]


# ==============================================================================
# 3. MODULÁRNÍ DATOVÁ VRSTVA (SEC EDGAR 10-K/S-1/10-Q + FMP + DUAL-CLASS + IPO GATE)
# ==============================================================================

@dataclass
class DataPoint:
    value: float | None
    status: str  # "REPORTED" | "ESTIMATED" | "MISSING"
    source: str
    note: str = ""

    def to_dict(self) -> dict:
        return {
            "value": round(self.value, 2) if self.value is not None else None,
            "status": self.status,
            "source": self.source,
            "note": self.note,
        }


_SEC_CIK_MAP_CACHE: dict[str, tuple[str, str]] | None = None
_SEC_FACTS_MEMORY_CACHE: dict[str, tuple[datetime, dict]] = {}


def _load_sec_ticker_to_cik_map() -> dict[str, tuple[str, str]]:
    global _SEC_CIK_MAP_CACHE
    if _SEC_CIK_MAP_CACHE is not None:
        return _SEC_CIK_MAP_CACHE

    cache_file = os.path.join(CACHE_DIR, "sec_company_tickers.json")
    raw_data = None

    if os.path.exists(cache_file):
        age_sec = datetime.now().timestamp() - os.path.getmtime(cache_file)
        if age_sec < 7 * 86400:
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    raw_data = json.load(f)
            except Exception:
                raw_data = None

    if raw_data is None:
        try:
            req = urllib.request.Request(
                "https://www.sec.gov/files/company_tickers.json",
                headers={"User-Agent": SEC_USER_AGENT, "Accept": "application/json"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                raw_data = json.loads(resp.read().decode("utf-8"))
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(raw_data, f)
        except Exception as err:
            logger.warning("Nelze stáhnout mapu CIK ze SEC EDGAR: %s", err)
            _SEC_CIK_MAP_CACHE = {}
            return _SEC_CIK_MAP_CACHE

    mapping: dict[str, tuple[str, str]] = {}
    if isinstance(raw_data, dict):
        for _, entry in raw_data.items():
            if not isinstance(entry, dict):
                continue
            tkr = str(entry.get("ticker") or "").upper().strip()
            cik_num = entry.get("cik_str")
            title = str(entry.get("title") or "").strip()
            if tkr and cik_num is not None:
                cik_padded = str(int(cik_num)).zfill(10)
                mapping[tkr] = (cik_padded, title)
                if "-" in tkr:
                    mapping[tkr.replace("-", ".")] = (cik_padded, title)
                if "." in tkr:
                    mapping[tkr.replace(".", "-")] = (cik_padded, title)

    _SEC_CIK_MAP_CACHE = mapping
    return mapping


def _extract_sec_annual_series(
    us_gaap: dict,
    dei: dict,
    tag_candidates: list[tuple[str, str]],
    max_years: int = 4,
    is_balance_sheet: bool = False,
) -> tuple[list[float], str | None, str | None, str | None]:
    annual_forms = ("10-K", "10-K/A", "20-F", "40-F", "S-1", "S-1/A")

    for taxonomy, tag in tag_candidates:
        root_dict = us_gaap if taxonomy == "us-gaap" else dei
        tag_obj = root_dict.get(tag)
        if not isinstance(tag_obj, dict):
            continue
        units = tag_obj.get("units") or {}
        unit_items = None
        for u_key in ("USD", "shares"):
            if u_key in units and isinstance(units[u_key], list):
                unit_items = units[u_key]
                break
        if unit_items is None and units:
            first_key = next(iter(units.keys()))
            if isinstance(units[first_key], list):
                unit_items = units[first_key]
        if not unit_items:
            continue

        by_fiscal_year: dict[int, tuple[str, float, str, str]] = {}
        latest_10q: tuple[str, float, str] | None = None

        for item in unit_items:
            if not isinstance(item, dict):
                continue
            form = str(item.get("form") or "").upper()
            val = safe_float(item.get("val"), None)
            if val is None:
                continue
            fy = item.get("fy")
            end_date = str(item.get("end") or "")
            start_date = str(item.get("start") or "")
            filed = str(item.get("filed") or "")
            frame = str(item.get("frame") or "").upper()

            days_span: int | None = None
            if start_date and end_date:
                try:
                    d0 = datetime.strptime(start_date[:10], "%Y-%m-%d")
                    d1 = datetime.strptime(end_date[:10], "%Y-%m-%d")
                    days_span = (d1 - d0).days
                except Exception:
                    days_span = None

            if form in annual_forms:
                if not is_balance_sheet:
                    if frame and not re.fullmatch(r"CY\d{4}", frame):
                        continue
                    if days_span is not None and days_span < 330:
                        continue
                    if not frame and days_span is None:
                        continue

                yr_key = int(fy) if isinstance(fy, int) else (int(end_date[:4]) if len(end_date) >= 4 and end_date[:4].isdigit() else None)
                if yr_key is None:
                    continue

                prev = by_fiscal_year.get(yr_key)
                if prev is None or (end_date, filed) >= (prev[0], prev[2]):
                    by_fiscal_year[yr_key] = (end_date, val, filed, form)

            elif form in ("10-Q", "10-Q/A"):
                if not is_balance_sheet:
                    if days_span is not None and not (75 <= days_span <= 105):
                        continue
                    if frame and not re.fullmatch(r"CY\d{4}Q[1-4]", frame):
                        continue
                    ann_val = val * 4.0
                else:
                    ann_val = val

                if latest_10q is None or (end_date, filed) >= (latest_10q[0], latest_10q[2]):
                    latest_10q = (end_date, ann_val, filed)

        if by_fiscal_year:
            sorted_years = sorted(by_fiscal_year.keys(), reverse=True)[:max_years]
            vals = [by_fiscal_year[y][1] for y in sorted_years]
            latest_end = by_fiscal_year[sorted_years[0]][0]
            latest_form = by_fiscal_year[sorted_years[0]][3]
            return vals, f"{taxonomy}:{tag}", latest_end, latest_form

        if latest_10q is not None:
            return [latest_10q[1]], f"{taxonomy}:{tag}", latest_10q[0], "10-Q TTM"

    return [], None, None, None


def fetch_sec_edgar_fundamentals(ticker: str) -> dict | None:
    clean_tkr = ticker.upper().strip()
    if "." in clean_tkr and not clean_tkr.endswith((".A", ".B", ".C")):
        return None

    cik_map = _load_sec_ticker_to_cik_map()
    sec_entry = cik_map.get(clean_tkr)
    if not sec_entry:
        return None

    cik_padded, sec_title = sec_entry
    now = datetime.now(timezone.utc)
    facts_json = None

    if cik_padded in _SEC_FACTS_MEMORY_CACHE and now < _SEC_FACTS_MEMORY_CACHE[cik_padded][0]:
        facts_json = _SEC_FACTS_MEMORY_CACHE[cik_padded][1]
    else:
        cache_path = os.path.join(CACHE_DIR, f"sec_facts_v2_{cik_padded}.json")
        if os.path.exists(cache_path):
            age_sec = datetime.now().timestamp() - os.path.getmtime(cache_path)
            if age_sec < 12 * 3600:
                try:
                    with open(cache_path, "r", encoding="utf-8") as f:
                        facts_json = json.load(f)
                except Exception:
                    facts_json = None

        if facts_json is None:
            try:
                url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik_padded}.json"
                req = urllib.request.Request(
                    url,
                    headers={"User-Agent": SEC_USER_AGENT, "Accept": "application/json"},
                )
                with urllib.request.urlopen(req, timeout=12) as resp:
                    facts_json = json.loads(resp.read().decode("utf-8"))
                with open(cache_path, "w", encoding="utf-8") as f:
                    json.dump(facts_json, f)
            except Exception as err:
                logger.warning("SEC EDGAR API selhalo pro %s (CIK %s): %s", clean_tkr, cik_padded, err)
                return None

        if facts_json is not None:
            _SEC_FACTS_MEMORY_CACHE[cik_padded] = (now + timedelta(minutes=45), facts_json)

    facts = (facts_json or {}).get("facts") or {}
    us_gaap = facts.get("us-gaap") or {}
    dei = facts.get("dei") or {}
    if not us_gaap:
        return None

    balance_sheet_keys = {
        "equity", "long_term_debt", "short_term_debt", "cash",
        "short_term_investments", "long_term_investments", "deferred_tax_liab",
        "insurance_float", "current_assets", "current_liabilities",
        "total_assets", "retained_earnings", "minority_interest", "shares",
    }

    tag_specs = {
        "revenue": [
            ("us-gaap", "RevenueFromContractWithCustomerExcludingAssessedTax"),
            ("us-gaap", "Revenues"),
            ("us-gaap", "SalesRevenueNet"),
            ("us-gaap", "RevenueFromContractWithCustomerIncludingAssessedTax"),
        ],
        "ebit": [("us-gaap", "OperatingIncomeLoss")],
        "da": [
            ("us-gaap", "DepreciationDepletionAndAmortization"),
            ("us-gaap", "DepreciationAndAmortization"),
            ("us-gaap", "Depreciation"),
        ],
        "net_income": [("us-gaap", "NetIncomeLoss"), ("us-gaap", "ProfitLoss")],
        "pretax_income": [
            ("us-gaap", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
            ("us-gaap", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"),
        ],
        "tax_provision": [("us-gaap", "IncomeTaxExpenseBenefit")],
        "interest_expense": [("us-gaap", "InterestExpense"), ("us-gaap", "InterestExpenseDebt")],
        "ocf": [
            ("us-gaap", "NetCashProvidedByUsedInOperatingActivities"),
            ("us-gaap", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"),
        ],
        "capex": [
            ("us-gaap", "PaymentsToAcquirePropertyPlantAndEquipment"),
            ("us-gaap", "PaymentsToAcquireProductiveAssets"),
        ],
        "sbc": [
            ("us-gaap", "ShareBasedCompensation"),
            ("us-gaap", "AllocatedShareBasedCompensationExpense"),
        ],
        "equity": [
            ("us-gaap", "StockholdersEquity"),
            ("us-gaap", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"),
        ],
        "long_term_debt": [("us-gaap", "LongTermDebtNoncurrent"), ("us-gaap", "LongTermDebt")],
        "short_term_debt": [
            ("us-gaap", "LongTermDebtCurrent"),
            ("us-gaap", "ShortTermBorrowings"),
            ("us-gaap", "DebtCurrent"),
        ],
        "cash": [
            ("us-gaap", "CashAndCashEquivalentsAtCarryingValue"),
            ("us-gaap", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"),
        ],
        "short_term_investments": [
            ("us-gaap", "MarketableSecuritiesCurrent"),
            ("us-gaap", "ShortTermInvestments"),
            ("us-gaap", "AvailableForSaleSecuritiesDebtSecuritiesCurrent"),
        ],
        "long_term_investments": [
            ("us-gaap", "MarketableSecuritiesNoncurrent"),
            ("us-gaap", "LongTermInvestments"),
            ("us-gaap", "EquityMethodInvestments"),
        ],
        "deferred_tax_liab": [
            ("us-gaap", "DeferredIncomeTaxLiabilitiesNet"),
            ("us-gaap", "DeferredTaxLiabilitiesNoncurrent"),
        ],
        "insurance_float": [
            ("us-gaap", "LiabilityForClaimsAndClaimsAdjustmentExpense"),
            ("us-gaap", "UnearnedPremiums"),
        ],
        "current_assets": [("us-gaap", "AssetsCurrent")],
        "current_liabilities": [("us-gaap", "LiabilitiesCurrent")],
        "total_assets": [("us-gaap", "Assets")],
        "retained_earnings": [("us-gaap", "RetainedEarningsAccumulatedDeficit")],
        "minority_interest": [("us-gaap", "MinorityInterest")],
        "shares": [
            ("dei", "EntityCommonStockSharesOutstanding"),
            ("us-gaap", "CommonStockSharesOutstanding"),
            ("us-gaap", "WeightedAverageNumberOfDilutedSharesOutstanding"),
        ],
    }

    series_map: dict[str, list[float]] = {}
    tags_used: dict[str, str] = {}
    latest_period = None
    detected_form = "10-K"

    for key, candidates in tag_specs.items():
        vals, used_tag, period_end, form_used = _extract_sec_annual_series(
            us_gaap,
            dei,
            candidates,
            max_years=4,
            is_balance_sheet=(key in balance_sheet_keys),
        )
        series_map[key] = vals
        if used_tag:
            tags_used[key] = used_tag
        if period_end and (latest_period is None or period_end > latest_period):
            latest_period = period_end
            if form_used:
                detected_form = form_used

    if not series_map["revenue"] and not series_map["total_assets"]:
        return None

    return {
        "provider": f"SEC EDGAR XBRL ({detected_form})",
        "cik": cik_padded,
        "entity_name": facts_json.get("entityName") or sec_title,
        "latest_period": latest_period or detected_form,
        "currency": "USD",
        "series": series_map,
        "tags_used": tags_used,
    }


def fetch_fmp_fundamentals(ticker: str) -> dict | None:
    if not FMP_API_KEY:
        return None
    clean_tkr = ticker.upper().strip()
    try:
        base = "https://financialmodelingprep.com/api/v3"
        urls = {
            "inc": f"{base}/income-statement/{quote(clean_tkr)}?limit=4&apikey={quote(FMP_API_KEY)}",
            "bs": f"{base}/balance-sheet-statement/{quote(clean_tkr)}?limit=4&apikey={quote(FMP_API_KEY)}",
            "cf": f"{base}/cash-flow-statement/{quote(clean_tkr)}?limit=4&apikey={quote(FMP_API_KEY)}",
        }
        fetched = {}
        for k, u in urls.items():
            req = urllib.request.Request(u, headers={"User-Agent": "TensegrityTerminal/5.2"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            if isinstance(data, list) and data:
                fetched[k] = data
        if "inc" not in fetched or "bs" not in fetched:
            return None

        inc_list = fetched["inc"]
        bs_list = fetched["bs"]
        cf_list = fetched.get("cf") or []

        def col(lst: list[dict], field: str) -> list[float]:
            out = []
            for row in lst:
                v = safe_float(row.get(field), None)
                if v is not None:
                    out.append(v)
            return out

        series = {
            "revenue": col(inc_list, "revenue"),
            "ebit": col(inc_list, "operatingIncome"),
            "ebitda": col(inc_list, "ebitda"),
            "da": col(inc_list, "depreciationAndAmortization"),
            "net_income": col(inc_list, "netIncome"),
            "pretax_income": col(inc_list, "incomeBeforeTax"),
            "tax_provision": col(inc_list, "incomeTaxExpense"),
            "interest_expense": col(inc_list, "interestExpense"),
            "ocf": col(cf_list, "operatingCashFlow"),
            "capex": [abs(x) for x in col(cf_list, "capitalExpenditure")],
            "sbc": [abs(x) for x in col(cf_list, "stockBasedCompensation")],
            "equity": col(bs_list, "totalStockholdersEquity"),
            "long_term_debt": col(bs_list, "longTermDebt"),
            "short_term_debt": col(bs_list, "shortTermDebt"),
            "cash": col(bs_list, "cashAndCashEquivalents"),
            "short_term_investments": col(bs_list, "shortTermInvestments"),
            "long_term_investments": col(bs_list, "longTermInvestments"),
            "deferred_tax_liab": col(bs_list, "deferredTaxLiabilitiesNonCurrent"),
            "insurance_float": [],
            "current_assets": col(bs_list, "totalCurrentAssets"),
            "current_liabilities": col(bs_list, "totalCurrentLiabilities"),
            "total_assets": col(bs_list, "totalAssets"),
            "retained_earnings": col(bs_list, "retainedEarnings"),
            "minority_interest": col(bs_list, "minorityInterest"),
            "shares": col(inc_list, "weightedAverageShsOutDil"),
        }
        rep_curr = str(inc_list[0].get("reportedCurrency") or "USD").upper()
        latest_date = str(inc_list[0].get("date") or "FY")
        return {
            "provider": "Financial Modeling Prep (Standardized)",
            "latest_period": latest_date,
            "currency": rep_curr,
            "series": series,
        }
    except Exception as err:
        logger.warning("FMP API selhalo pro %s: %s", clean_tkr, err)
        return None


def fetch_unified_financial_dataset(clean_ticker: str) -> dict:
    info: dict = {}
    fin_df = None
    bs_df = None
    cf_df = None

    try:
        stock = yf.Ticker(clean_ticker)
        info = stock.info or {}
        fin_df = stock.financials
        bs_df = stock.balance_sheet
        cf_df = stock.cashflow
    except Exception as err:
        logger.warning("Chyba při stahování tržního profilu pro %s: %s", clean_ticker, err)

    quote_type = str(info.get("quoteType") or "EQUITY").upper()
    if quote_type in ("INDEX", "MUTUALFUND", "CRYPTOCURRENCY", "CURRENCY", "FUTURE", "ETF"):
        raise ValueError(
            f"Instrument '{clean_ticker}' je typu {quote_type} (nikoli veřejně obchodovaná společnost s účetními výkazy)."
        )

    raw_currency = str(info.get("currency") or ("CZK" if clean_ticker.endswith(".PR") else "USD"))
    market_currency = "GBP" if raw_currency == "GBp" else raw_currency.upper()
    price_scale = 0.01 if raw_currency == "GBp" else 1.0

    current_price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"), None)
    if current_price is not None:
        current_price *= price_scale
    if current_price is None or current_price <= 0:
        try:
            hist_px = yf.Ticker(clean_ticker).history(period="5d")["Close"].dropna()
            if not hist_px.empty:
                current_price = float(hist_px.iloc[-1]) * price_scale
        except Exception:
            pass
    if current_price is None or current_price <= 0:
        raise ValueError(f"Nepodařilo se zjistit aktuální tržní cenu pro titul '{clean_ticker}'.")

    sec_data = fetch_sec_edgar_fundamentals(clean_ticker)
    fmp_data = None if sec_data is not None else fetch_fmp_fundamentals(clean_ticker)

    if sec_data is not None:
        stmt_provider = f"{sec_data['provider']} ({sec_data['latest_period']})"
        stmt_currency = sec_data["currency"]
        stmt_fx = get_fx_conversion_rate(stmt_currency, market_currency)
        raw_series = sec_data["series"]
        scaled_series = {k: [v * stmt_fx for v in vals] for k, vals in raw_series.items()}
        entity_name = str(info.get("longName") or info.get("shortName") or sec_data.get("entity_name") or clean_ticker)
    elif fmp_data is not None:
        stmt_provider = f"{fmp_data['provider']} ({fmp_data['latest_period']})"
        stmt_currency = fmp_data["currency"]
        stmt_fx = get_fx_conversion_rate(stmt_currency, market_currency)
        raw_series = fmp_data["series"]
        scaled_series = {k: [v * stmt_fx for v in vals] for k, vals in raw_series.items()}
        entity_name = str(info.get("longName") or info.get("shortName") or clean_ticker)
    else:
        stmt_provider = "Yahoo Finance (Neověřená struktura výkazů)"
        stmt_currency = str(info.get("financialCurrency") or market_currency).upper()
        stmt_fx = get_fx_conversion_rate(stmt_currency, market_currency)
        entity_name = str(info.get("longName") or info.get("shortName") or clean_ticker)
        scaled_series = {
            "revenue": extract_statement_series(fin_df, ["Total Revenue", "Operating Revenue"], fx_mult=stmt_fx),
            "ebit": extract_statement_series(fin_df, ["EBIT", "Operating Income"], fx_mult=stmt_fx),
            "ebitda": extract_statement_series(fin_df, ["EBITDA", "Normalized EBITDA"], fx_mult=stmt_fx),
            "da": extract_statement_series(cf_df, ["Depreciation And Amortization", "Depreciation Amortization Depletion"], fx_mult=stmt_fx),
            "net_income": extract_statement_series(fin_df, ["Net Income", "Net Income Common Stockholders"], fx_mult=stmt_fx),
            "pretax_income": extract_statement_series(fin_df, ["Pretax Income"], fx_mult=stmt_fx),
            "tax_provision": extract_statement_series(fin_df, ["Tax Provision"], fx_mult=stmt_fx),
            "interest_expense": extract_statement_series(fin_df, ["Interest Expense", "Interest Expense Non Operating"], fx_mult=stmt_fx),
            "ocf": extract_statement_series(cf_df, ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"], fx_mult=stmt_fx),
            "capex": [abs(x) for x in extract_statement_series(cf_df, ["Capital Expenditure", "Purchase Of PPE"], fx_mult=stmt_fx)],
            "sbc": [abs(x) for x in extract_statement_series(cf_df, ["Stock Based Compensation"], fx_mult=stmt_fx)],
            "equity": extract_statement_series(bs_df, ["Stockholders Equity", "Total Equity Gross Minority Interest"], fx_mult=stmt_fx),
            "long_term_debt": extract_statement_series(bs_df, ["Long Term Debt", "Total Debt"], fx_mult=stmt_fx),
            "short_term_debt": extract_statement_series(bs_df, ["Current Debt", "Current Debt And Capital Lease Obligation"], fx_mult=stmt_fx),
            "cash": extract_statement_series(bs_df, ["Cash And Cash Equivalents"], fx_mult=stmt_fx),
            "short_term_investments": extract_statement_series(bs_df, ["Other Short Term Investments"], fx_mult=stmt_fx),
            "long_term_investments": extract_statement_series(bs_df, ["Investments And Advances", "Long Term Equity Investment"], fx_mult=stmt_fx),
            "deferred_tax_liab": extract_statement_series(bs_df, ["Non Current Deferred Taxes Liabilities"], fx_mult=stmt_fx),
            "insurance_float": [],
            "current_assets": extract_statement_series(bs_df, ["Current Assets"], fx_mult=stmt_fx),
            "current_liabilities": extract_statement_series(bs_df, ["Current Liabilities"], fx_mult=stmt_fx),
            "total_assets": extract_statement_series(bs_df, ["Total Assets"], fx_mult=stmt_fx),
            "retained_earnings": extract_statement_series(bs_df, ["Retained Earnings"], fx_mult=stmt_fx),
            "minority_interest": extract_statement_series(bs_df, ["Minority Interest"], fx_mult=stmt_fx),
            "shares": [],
        }

    info_tot_rev = safe_float(info.get("totalRevenue"), None)
    if info_tot_rev is not None and info_tot_rev > 0 and scaled_series.get("revenue"):
        info_rev_scaled = info_tot_rev * stmt_fx
        sec_rev_latest = scaled_series["revenue"][0]
        if sec_rev_latest > 0 and (info_rev_scaled / sec_rev_latest) > 2.2:
            scaled_series["revenue"][0] = info_rev_scaled

    def dp_from_series(key: str, info_fallback_key: str | None = None, scale_fx: bool = True, abs_val: bool = False) -> DataPoint:
        vals = scaled_series.get(key) or []
        if vals:
            v = abs(vals[0]) if abs_val else vals[0]
            return DataPoint(value=v, status="REPORTED", source=stmt_provider)
        if info_fallback_key:
            raw_v = safe_float(info.get(info_fallback_key), None)
            if raw_v is not None:
                v = raw_v * (stmt_fx if scale_fx else 1.0)
                if abs_val:
                    v = abs(v)
                return DataPoint(value=v, status="REPORTED", source="Market Feed (Summary)")
        return DataPoint(value=None, status="MISSING", source=stmt_provider, note=f"Chybí položka '{key}' ve výkazu.")

    rev_dp = dp_from_series("revenue", "totalRevenue")
    ebit_dp = dp_from_series("ebit")
    da_dp = dp_from_series("da", abs_val=True)
    ni_dp = dp_from_series("net_income", "netIncomeToCommon")
    pretax_dp = dp_from_series("pretax_income")
    tax_dp = dp_from_series("tax_provision")
    int_exp_dp = dp_from_series("interest_expense", abs_val=True)
    ocf_dp = dp_from_series("ocf", "operatingCashflow")
    capex_dp = dp_from_series("capex", abs_val=True)
    sbc_dp = dp_from_series("sbc", abs_val=True)
    equity_dp = dp_from_series("equity")
    cash_dp = dp_from_series("cash", "totalCash")
    st_inv_dp = dp_from_series("short_term_investments")
    lt_inv_dp = dp_from_series("long_term_investments")
    dtl_dp = dp_from_series("deferred_tax_liab", abs_val=True)
    float_dp = dp_from_series("insurance_float", abs_val=True)
    min_int_dp = dp_from_series("minority_interest", "minorityInterest")

    if scaled_series.get("ebitda"):
        ebitda_dp = DataPoint(value=scaled_series["ebitda"][0], status="REPORTED", source=stmt_provider)
    elif ebit_dp.value is not None and da_dp.value is not None:
        ebitda_dp = DataPoint(value=ebit_dp.value + da_dp.value, status="REPORTED", source=f"{stmt_provider} (EBIT + D&A)")
    else:
        raw_eb_info = safe_float(info.get("ebitda"), None)
        if raw_eb_info is not None:
            ebitda_dp = DataPoint(value=raw_eb_info * stmt_fx, status="REPORTED", source="Market Feed (Summary)")
        else:
            ebitda_dp = DataPoint(value=None, status="MISSING", source=stmt_provider, note="Chybí EBITDA i kombinace EBIT + odpisy.")

    lt_debt_vals = scaled_series.get("long_term_debt") or []
    st_debt_vals = scaled_series.get("short_term_debt") or []
    if lt_debt_vals or st_debt_vals:
        tot_d = (lt_debt_vals[0] if lt_debt_vals else 0.0) + (st_debt_vals[0] if st_debt_vals else 0.0)
        debt_dp = DataPoint(value=tot_d, status="REPORTED", source=stmt_provider)
    else:
        raw_d_info = safe_float(info.get("totalDebt"), None)
        if raw_d_info is not None:
            debt_dp = DataPoint(value=raw_d_info * stmt_fx, status="REPORTED", source="Market Feed (Summary)")
        else:
            debt_dp = DataPoint(value=None, status="MISSING", source=stmt_provider, note="Chybí údaj o celkovém dluhu v rozvaze.")

    sec_shares_vals = raw_series.get("shares") if (sec_data or fmp_data) else []
    info_shares = safe_float(info.get("sharesOutstanding"), None)
    info_mcap = safe_float(info.get("marketCap"), None)

    shares_dp: DataPoint | None = None
    if sec_shares_vals and sec_shares_vals[0] > 0:
        sec_sh = sec_shares_vals[0]
        if info_mcap and info_mcap > 0 and current_price > 0:
            implied_class_shares = info_mcap / current_price
            ratio_sh = max(sec_sh, implied_class_shares) / max(1.0, min(sec_sh, implied_class_shares))
            if ratio_sh > 1.25:
                shares_dp = DataPoint(
                    value=implied_class_shares,
                    status="REPORTED",
                    source="MarketCap / Price (Dual-Class / Post-IPO Share Class)",
                    note=f"Počet akcií upraven pro obchodovanou třídu '{clean_ticker}'.",
                )
        if shares_dp is None:
            shares_dp = DataPoint(value=sec_sh, status="REPORTED", source=stmt_provider)
    elif info_mcap and info_mcap > 0 and current_price > 0:
        implied_sh = info_mcap / current_price
        if info_shares and info_shares > 0:
            ratio_sh = max(info_shares, implied_sh) / max(1.0, min(info_shares, implied_sh))
            if ratio_sh > 1.25:
                shares_dp = DataPoint(
                    value=implied_sh,
                    status="ESTIMATED",
                    source="MarketCap / Price (Post-IPO Normalization)",
                    note="Počet akcií sjednocen s aktuální tržní kapitalizací po vstupu na burzu.",
                )
            else:
                shares_dp = DataPoint(value=info_shares, status="REPORTED", source="Market Feed")
        else:
            shares_dp = DataPoint(value=implied_sh, status="ESTIMATED", source="MarketCap / Price")
    elif info_shares and info_shares > 0:
        shares_dp = DataPoint(value=info_shares, status="REPORTED", source="Market Feed")
    else:
        raise ValueError(f"Nelze ověřit počet vydaných akcií ani tržní kapitalizaci pro '{clean_ticker}'.")

    if info_mcap and info_mcap > 0:
        mcap_dp = DataPoint(value=info_mcap, status="REPORTED", source="Market Feed")
    else:
        mcap_dp = DataPoint(value=current_price * shares_dp.value, status="ESTIMATED", source="Price * Shares")

    return {
        "ticker": clean_ticker,
        "company_name": entity_name,
        "info": info,
        "market_currency": market_currency,
        "stmt_currency": stmt_currency,
        "stmt_fx": stmt_fx,
        "stmt_provider": stmt_provider,
        "current_price": current_price,
        "series": scaled_series,
        "points": {
            "shares": shares_dp,
            "market_cap": mcap_dp,
            "revenue": rev_dp,
            "ebit": ebit_dp,
            "ebitda": ebitda_dp,
            "da": da_dp,
            "net_income": ni_dp,
            "pretax_income": pretax_dp,
            "tax_provision": tax_dp,
            "interest_expense": int_exp_dp,
            "ocf": ocf_dp,
            "capex": capex_dp,
            "sbc": sbc_dp,
            "equity": equity_dp,
            "debt": debt_dp,
            "cash": cash_dp,
            "short_term_investments": st_inv_dp,
            "long_term_investments": lt_inv_dp,
            "deferred_tax_liab": dtl_dp,
            "insurance_float": float_dp,
            "minority_interest": min_int_dp,
        },
    }


# ==============================================================================
# 4. ŽIVÁ MAKROEKONOMICKÁ DATA, MĚNOVÉ KURZY A ČNB
# ==============================================================================

_MACRO_CACHE: dict = {"expires_at": None, "data": None}
_CNB_CACHE: dict = {"expires_at": None, "data": None}


def get_live_market_benchmarks() -> dict:
    now = datetime.now(timezone.utc)
    if _MACRO_CACHE["data"] is not None and _MACRO_CACHE["expires_at"] is not None and now < _MACRO_CACHE["expires_at"]:
        return _MACRO_CACHE["data"]

    us10y = None
    sp500_pe = None

    try:
        req = urllib.request.Request(
            "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            lines = resp.read().decode("utf-8", errors="ignore").strip().splitlines()
        for line in reversed(lines[1:]):
            parts = line.split(",")
            if len(parts) == 2 and parts[1].strip() != ".":
                candidate = safe_float(parts[1].strip(), 0.0)
                if 0.5 < candidate < 15.0:
                    us10y = round(candidate, 2)
                    break
    except Exception:
        pass

    if us10y is None:
        try:
            tnx_hist = yf.Ticker("^TNX").history(period="5d", interval="1d")
            if not tnx_hist.empty:
                last_tnx = safe_float(tnx_hist["Close"].dropna().iloc[-1], 0.0)
                if 0.5 < last_tnx < 15.0:
                    us10y = round(last_tnx, 2)
        except Exception:
            pass

    try:
        spy_info = yf.Ticker("SPY").info or {}
        trailing_pe = safe_float(spy_info.get("trailingPE"), 0.0)
        if 8.0 < trailing_pe < 60.0:
            sp500_pe = round(trailing_pe, 1)
    except Exception:
        pass

    result = {
        "us10y_yield": us10y if us10y is not None else 4.25,
        "sp500_pe": sp500_pe if sp500_pe is not None else 24.5,
    }
    _MACRO_CACHE["data"] = result
    _MACRO_CACHE["expires_at"] = now + timedelta(minutes=20)
    return result


def _extract_html_tables(raw_html: str) -> list[list[list[str]]]:
    from html.parser import HTMLParser

    class TableParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.tables: list[list[list[str]]] = []
            self._in_table = False
            self._in_row = False
            self._in_cell = False
            self._current_table: list[list[str]] = []
            self._current_row: list[str] = []
            self._current_cell: list[str] = []

        def handle_starttag(self, tag: str, attrs):
            if tag == "table":
                self._in_table = True
                self._current_table = []
            elif tag == "tr" and self._in_table:
                self._in_row = True
                self._current_row = []
            elif tag in ("td", "th") and self._in_row:
                self._in_cell = True
                self._current_cell = []

        def handle_data(self, data: str):
            if self._in_cell:
                self._current_cell.append(data)

        def handle_endtag(self, tag: str):
            if tag in ("td", "th") and self._in_cell:
                self._in_cell = False
                cell_text = re.sub(r"\s+", " ", "".join(self._current_cell)).strip()
                self._current_row.append(cell_text)
            elif tag == "tr" and self._in_row:
                self._in_row = False
                if self._current_row:
                    self._current_table.append(self._current_row)
            elif tag == "table" and self._in_table:
                self._in_table = False
                if self._current_table:
                    self.tables.append(self._current_table)

    parser = TableParser()
    parser.feed(raw_html)
    return parser.tables


def get_cnb_macro_indicators() -> dict:
    now = datetime.now(timezone.utc)
    if _CNB_CACHE["data"] is not None and _CNB_CACHE["expires_at"] is not None and now < _CNB_CACHE["expires_at"]:
        return _CNB_CACHE["data"]

    repo_val = 4.00
    discount_val = 3.00
    lombard_val = 5.00
    inflation_val = 2.4

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept-Language": "cs-CZ,cs;q=0.9",
    }

    try:
        req = urllib.request.Request(
            "https://www.cnb.cz/cs/casto-kladene-otazky/Jak-se-vyvijela-dvoutydenni-repo-sazba-CNB/",
            headers=headers,
        )
        with urllib.request.urlopen(req, timeout=7) as resp:
            history_html = resp.read().decode("utf-8", errors="ignore")

        tables = _extract_html_tables(history_html)
        best_date = (0, 0, 0)
        for table in tables:
            if len(table) < 2:
                continue
            header_cells = [c.lower() for c in table[0]]
            repo_idx = next((i for i, c in enumerate(header_cells) if "repo" in c), None)
            disc_idx = next((i for i, c in enumerate(header_cells) if "diskont" in c), None)
            lomb_idx = next((i for i, c in enumerate(header_cells) if "lombard" in c), None)
            if repo_idx is None:
                continue

            for row in table[1:]:
                if len(row) <= repo_idx:
                    continue
                date_match = re.search(r"(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})", row[0])
                if not date_match:
                    continue
                row_date = (int(date_match.group(3)), int(date_match.group(2)), int(date_match.group(1)))
                if row_date >= best_date:
                    num_match = re.search(r"(\d+[,.]\d+)", row[repo_idx])
                    if num_match:
                        best_date = row_date
                        repo_val = float(num_match.group(1).replace(",", "."))
                        if disc_idx is not None and len(row) > disc_idx:
                            d_match = re.search(r"(\d+[,.]\d+)", row[disc_idx])
                            if d_match:
                                discount_val = float(d_match.group(1).replace(",", "."))
                        else:
                            discount_val = max(0.0, round(repo_val - 1.0, 2))
                        if lomb_idx is not None and len(row) > lomb_idx:
                            l_match = re.search(r"(\d+[,.]\d+)", row[lomb_idx])
                            if l_match:
                                lombard_val = float(l_match.group(1).replace(",", "."))
                        else:
                            lombard_val = round(repo_val + 1.0, 2)
    except Exception:
        pass

    try:
        req_home = urllib.request.Request("https://www.cnb.cz/cs/", headers=headers)
        with urllib.request.urlopen(req_home, timeout=7) as resp_home:
            home_html = resp_home.read().decode("utf-8", errors="ignore")
        plain_home = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(home_html)))
        infl_match = re.search(
            r"(?:(\d+[,.]\d+)\s*%\s*(?:Míra\s+inflace|Inflace)|(?:Míra\s+inflace|Inflace)\s*(\d+[,.]\d+)\s*%)",
            plain_home,
            flags=re.IGNORECASE,
        )
        if infl_match:
            raw_infl = infl_match.group(1) or infl_match.group(2)
            parsed_infl = float(raw_infl.replace(",", "."))
            if -5.0 <= parsed_infl <= 30.0:
                inflation_val = parsed_infl
    except Exception:
        pass

    result = {
        "repo_rate": _format_cz_percent(repo_val, 2),
        "discount_rate": _format_cz_percent(discount_val, 2),
        "lombard_rate": _format_cz_percent(lombard_val, 2),
        "latest_inflation": _format_cz_percent(inflation_val, 1),
        "inflation_target": "2,0 % (toleranční pásmo 1,0 % až 3,0 %)",
    }
    _CNB_CACHE["data"] = result
    _CNB_CACHE["expires_at"] = now + timedelta(minutes=30)
    return result


def get_live_forex_rates() -> dict:
    cnb_daily_rates = {}
    try:
        req = urllib.request.Request(
            "https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            lines = resp.read().decode("utf-8", errors="ignore").splitlines()
        for line in lines[2:]:
            parts = line.split("|")
            if len(parts) == 5:
                currency_code = parts[3].strip().upper()
                amount = safe_float(parts[2].strip().replace(",", "."), 0.0)
                czk_value = safe_float(parts[4].strip().replace(",", "."), 0.0)
                if amount > 0 and czk_value > 0:
                    cnb_daily_rates[f"{currency_code}/CZK"] = round(czk_value / amount, 4)
    except Exception:
        pass

    rates = {}
    for pair, symbol in FOREX_PAIRS.items():
        base_currency = pair.replace(" ", "").replace("/", "").replace("_", "").upper()[:3]
        cnb_key = f"{base_currency}/CZK"
        if cnb_key in cnb_daily_rates:
            rate = cnb_daily_rates[cnb_key]
            live = True
        else:
            rate = FOREX_FALLBACK_RATES[pair]
            live = False

        if not live:
            try:
                history = yf.Ticker(symbol).history(period="5d", interval="1d")
                closes = history["Close"].dropna()
                if not closes.empty:
                    candidate = safe_float(closes.iloc[-1], 0.0)
                    if candidate > 0:
                        rate = round(candidate, 4)
                        live = True
            except Exception:
                pass

        rates[pair] = {"rate": rate, "live": live}

    return {
        "rates": rates,
        "updated_at": get_prague_now().isoformat(timespec="minutes"),
    }


# ==============================================================================
# 5. SPRÁVA ODBĚRATELŮ, PŘEKLAD ZPRÁV DO ČEŠTINY A E-MAILOVÝ NEWSLETTER
# ==============================================================================

def is_valid_email_address(email: str) -> bool:
    if not isinstance(email, str) or len(email) > 254:
        return False
    parsed = parseaddr(email)
    return parsed[1] == email and bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email))


def get_subscribers_internal() -> list[str]:
    merged: list[str] = []
    seen_lower: set[str] = set()

    if ENV_SUBSCRIBERS:
        for raw_item in re.split(r"[,;\s]+", ENV_SUBSCRIBERS):
            clean_item = raw_item.strip()
            if is_valid_email_address(clean_item) and clean_item.casefold() not in seen_lower:
                merged.append(clean_item)
                seen_lower.add(clean_item.casefold())

    if os.path.exists(SUBSCRIBERS_FILE):
        try:
            with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as file:
                data = json.load(file)
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, str):
                        clean_item = item.strip()
                        if is_valid_email_address(clean_item) and clean_item.casefold() not in seen_lower:
                            merged.append(clean_item)
                            seen_lower.add(clean_item.casefold())
        except (OSError, json.JSONDecodeError):
            pass

    return merged


def save_subscriber_private(email: str) -> bool:
    normalized = email.strip()
    subscribers = get_subscribers_internal()
    if any(existing.casefold() == normalized.casefold() for existing in subscribers):
        return False
    subscribers.append(normalized)
    temporary_path = SUBSCRIBERS_FILE + ".tmp"
    with open(temporary_path, "w", encoding="utf-8") as file:
        json.dump(subscribers, file, ensure_ascii=False, indent=2)
        file.flush()
        os.fsync(file.fileno())
    os.replace(temporary_path, SUBSCRIBERS_FILE)
    return True


def load_sent_slots() -> set[str]:
    if not os.path.exists(SCHEDULER_STATE_FILE):
        return set()
    try:
        with open(SCHEDULER_STATE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if isinstance(data, list):
            return set(str(x) for x in data[-150:])
    except Exception:
        pass
    return set()


def save_sent_slots(slots: set[str]):
    try:
        tmp = SCHEDULER_STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as file:
            json.dump(sorted(list(slots))[-150:], file, ensure_ascii=False, indent=2)
            file.flush()
            os.fsync(file.fileno())
        os.replace(tmp, SCHEDULER_STATE_FILE)
    except Exception as err:
        logger.warning("Nelze uložit stav plánovače: %s", err)


def _decode_bytes_smart(raw_bytes: bytes, content_type: str = "") -> str:
    ct_lower = (content_type or "").lower()
    if "windows-1250" in ct_lower or "cp1250" in ct_lower:
        return raw_bytes.decode("cp1250", errors="replace")
    if "iso-8859-2" in ct_lower:
        return raw_bytes.decode("iso-8859-2", errors="replace")
    try:
        return raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        for enc in ("cp1250", "iso-8859-2", "latin-1"):
            try:
                return raw_bytes.decode(enc)
            except Exception:
                continue
        return raw_bytes.decode("utf-8", errors="replace")


def _looks_english(text: str) -> bool:
    if not text:
        return False
    en_markers = re.findall(
        r"\b(the|and|with|from|for|that|this|are|was|were|have|has|will|would|could|should|market|stocks|shares|fed|rate|rates|inflation|economy|quarter|earnings|guidance|investors|trading|growth|debt|says|after|over|new|us|china|europe|bank|price|prices|oil|tech|tariff|threats|pile|pressure|partners|sign|trade|deals)\b",
        text.lower(),
    )
    cz_chars = re.findall(r"[ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]", text)
    return len(en_markers) >= 1 and len(cz_chars) == 0


def _fallback_translate_en_to_cs(text: str) -> str:
    clean = re.sub(r"\s+", " ", (text or "")).strip()
    if not clean or not _looks_english(clean):
        return clean

    # 1. Primární překlad přes Google Translate endpoint (přeloží celé věty i titulky do přirozené češtiny)
    try:
        gt_url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=cs&dt=t&q={quote(clean[:2200])}"
        req_gt = urllib.request.Request(
            gt_url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
        )
        with urllib.request.urlopen(req_gt, timeout=8) as resp_gt:
            gt_payload = json.loads(resp_gt.read().decode("utf-8", errors="ignore"))
        if isinstance(gt_payload, list) and gt_payload and isinstance(gt_payload[0], list):
            gt_translated = "".join(
                str(seg[0]) for seg in gt_payload[0] if isinstance(seg, list) and seg and seg[0]
            ).strip()
            if gt_translated and not _looks_english(gt_translated):
                return gt_translated
    except Exception:
        pass

    # 2. Záložní překlad po blocích přes MyMemory API
    chunks = [clean[i:i + 450] for i in range(0, min(len(clean), 1800), 450)]
    translated_parts: list[str] = []
    for chunk in chunks:
        try:
            api_url = f"https://api.mymemory.translated.net/get?q={quote(chunk)}&langpair=en|cs"
            req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=6) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
            translated = str((payload.get("responseData") or {}).get("translatedText") or "").strip()
            if translated and "MYMEMORY WARNING" not in translated.upper():
                translated_parts.append(translated)
            else:
                translated_parts.append(chunk)
        except Exception:
            translated_parts.append(chunk)
    return " ".join(translated_parts).strip()


def _fetch_article_lead_paragraphs(article_url: str) -> str:
    """Stáhne úvodní odstavce přímo z článku, pokud má RSS jen krátký popisek."""
    if not article_url or not article_url.startswith(("http://", "https://")):
        return ""
    try:
        req = urllib.request.Request(
            article_url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "text/html,application/xhtml+xml",
            },
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            ctype = resp.headers.get("Content-Type", "")
            page_html = _decode_bytes_smart(resp.read(160_000), ctype)
        paragraphs = re.findall(r"<p[^>]*>(.*?)</p>", page_html, flags=re.IGNORECASE | re.DOTALL)
        clean_paras: list[str] = []
        for p in paragraphs:
            txt = re.sub(r"<[^>]+>", " ", html.unescape(p))
            txt = re.sub(r"\s+", " ", txt).strip()
            if len(txt) >= 55 and not any(
                skip in txt.lower()
                for skip in ("cookie", "subscribe", "sign in", "prihlasit", "přihlásit", "newsletter", "all rights reserved")
            ):
                clean_paras.append(txt)
            if len(" ".join(clean_paras)) >= 1100:
                break
        return " ".join(clean_paras)[:1400]
    except Exception:
        return ""


def _clean_czech_news_summary(text: str, title_cs: str) -> str:
    """Zajistí 100% český text v rozsahu 4 až 5 věcných vět bez anglických zbytků."""
    cleaned = re.sub(r"\[?\s*Zdroj\s*:[^\]]*\]?", "", text or "", flags=re.IGNORECASE)
    cleaned = re.sub(r"https?://\S+", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    if _looks_english(cleaned):
        cleaned = _fallback_translate_en_to_cs(cleaned)

    raw_sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", cleaned) if len(s.strip()) > 10]
    sentences: list[str] = []
    for s in raw_sentences:
        if _looks_english(s):
            s = _fallback_translate_en_to_cs(s)
        if _looks_english(s):
            continue
        s_clean = s.rstrip(".!?").strip() + "."
        if s_clean not in sentences:
            sentences.append(s_clean)
        if len(sentences) >= 5:
            break

    clean_subject = re.sub(r"^\d+\.\s*", "", title_cs or "").rstrip(".!?").strip()
    if _looks_english(clean_subject):
        clean_subject = _fallback_translate_en_to_cs(clean_subject).rstrip(".!?").strip()

    # Doplnění do minimálně 4 vět (cíl 4–5 vět) výhradně čistou češtinou
    fallback_pool = [
        f"Klíčovým bodem aktuálního vývoje jsou události kolem tématu {clean_subject}." if clean_subject and not _looks_english(clean_subject) else "Zpráva přináší zásadní impulsy pro vývoj na finančních a kapitálových trzích.",
        "Tržní účastníci vyhodnocují bezprostřední dopad těchto informací na ocenění aktiv a ziskovost firem.",
        "Další směřování situace může ovlivnit měnovou politiku centrálních bank i regulatorní prostředí.",
        "Institucionální investoři proto pečlivě sledují navazující makroekonomická data a oficiální stanoviska.",
    ]
    idx_fb = 0
    while len(sentences) < 4 and idx_fb < len(fallback_pool):
        cand = fallback_pool[idx_fb]
        idx_fb += 1
        if cand not in sentences:
            sentences.append(cand)

    return " ".join(sentences[:5])


def _translate_and_summarize_batch_czech(articles: list[dict]) -> list[dict]:
    if not articles:
        return articles

    if GEMINI_API_KEY:
        try:
            items_payload = [
                {
                    "index": idx,
                    "title": art["raw_title"],
                    "snippet": art["raw_text"][:2200],
                }
                for idx, art in enumerate(articles)
            ]
            prompt = (
                "Jsi špičkový editor a překladatel českého institucionálního zpravodaje Tensegrity Terminal zaměřeného na "
                "makroekonomiku, finanční trhy, akvizice (M&A), centrální banky, hospodářskou politiku, investice a legislativu.\n"
                "Pro KAŽDOU položku ze seznamu splň přísně tato pravidla:\n"
                "1. Přelož nadpis ('title_cs') do naprosto plynulé, spisovné ČEŠTINY (nikdy nenechávej anglická slova v titulku).\n"
                "2. Vytvoř věcné, hutné a konkrétní analytické shrnutí ('summary_cs') v ČEŠTINĚ o rozsahu 4 AŽ 5 VĚT na základě textu článku.\n"
                "3. Do textu shrnutí NIKDY nepiš název média ani zdroj článku.\n"
                "4. Vrať výhradně čistý JSON ve formátu pole objektů: [{\"index\": 0, \"title_cs\": \"...\", \"summary_cs\": \"...\"}].\n\n"
                f"Vstupní data: {json.dumps(items_payload, ensure_ascii=False)}"
            )
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            body = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json"},
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(body).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=22) as resp:
                raw_resp = json.loads(resp.read().decode("utf-8"))
            candidates = raw_resp.get("candidates") or []
            if candidates:
                text_out = candidates[0]["content"]["parts"][0]["text"].strip()
                text_out = re.sub(r"^```(?:json)?|```$", "", text_out, flags=re.MULTILINE).strip()
                parsed_list = json.loads(text_out)
                by_idx = {int(item["index"]): item for item in parsed_list if "index" in item}
                for idx, art in enumerate(articles):
                    if idx in by_idx:
                        t_cs = str(by_idx[idx].get("title_cs") or art["raw_title"]).strip()
                        if _looks_english(t_cs):
                            t_cs = _fallback_translate_en_to_cs(t_cs)
                        s_cs = str(by_idx[idx].get("summary_cs") or art["raw_text"]).strip()
                        art["title_cs"] = t_cs
                        art["summary_cs"] = _clean_czech_news_summary(s_cs, t_cs)
                return articles
        except Exception as err:
            logger.warning("Překlad zpráv přes Gemini selhal, spouštím záložní překlad: %s", err)

    for art in articles:
        t_cs = _fallback_translate_en_to_cs(art["raw_title"])
        art["title_cs"] = t_cs
        art["summary_cs"] = _clean_czech_news_summary(art["raw_text"], t_cs)
    return articles


def _is_relevant_business_article(title: str, text: str, czech_domestic_only: bool = False) -> bool:
    combined = f" {title} {text} ".lower()

    banned_keywords = (
        "fotbal", "slavia", "sparta", "hokej", "liga mistrů", "penalta", "zápas", "tenis", "olympiád",
        "nehoda", "obětí", "káhiře", "vražda", "počasí", "horoskop", "celebrit", "showbyznys", "film", "seriál",
        "married", "wife", "husband", "401(k)", "důchodový fond", "rodinný rozpočet", "kroužky", "hypotéka",
        "jak investovat, když", "proč tolik párů", "ženatý už", "pokrytec", "zamilov", "svatb", "recept",
        "discount rate meetings", "board of governors", "zasedání rady k diskontní",
    )
    if any(b in combined for b in banned_keywords):
        return False

    required_domain_terms = (
        "financ", "akvizic", "fúze", "m&a", "merger", "acquisition", "buyout",
        "politik", "vlád", "parlament", "senát", "sněmovn", "prezident", "ministers", "volb", "geopolit", "sankc", "clo", "tariff",
        "investic", "akci", "burz", "dluhopis", "fond", "kapitál", "ipo", "trh", "market", "stock", "equity", "yield", "fed", "ecb", "čnb", "inflac", "úrok",
        "zákoník", "zákon", "legislativ", "novela", "nařízení", "směrnic", "regulac", "antimonopol", "komise", "law", "bill", "regulation", "antitrust",
        "soud", "žalob", "rozsudek", "ústavní soud", "nejvyšší soud", "arbitráž", "pokut", "court", "ruling", "lawsuit", "sec", "doj",
    )
    if not any(term in combined for term in required_domain_terms):
        return False

    if czech_domestic_only:
        foreign_markers = ("washington -", "new york -", "peking -", "moskva -", "londýn -", "brusel -", "spacex", "nasa", "pentagon", "wall street")
        if any(fm in combined for fm in foreign_markers):
            return False
        cz_markers = (
            "česk", " čr ", "praha", "brno", "ostrav", "čnb", "korun", "vlád", "senát",
            "sněmovn", "čez", "škoda", "kofola", "moneta", "komerční banka", "orlen",
            "tuzemsk", "domácí", "ministerstv", "rozpoč", "zákoník", "zákon", "soud", "ústavní", "burza", "px",
        )
        return any(cm in combined for cm in cz_markers)

    return True


def _fetch_single_rss_source(source_tuple: tuple[str, str, str], badge_color: str, czech_domestic_only: bool = False) -> list[dict]:
    source_name, feed_url, fallback_site = source_tuple
    items_out: list[dict] = []
    try:
        req = urllib.request.Request(
            feed_url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/rss+xml, application/xml, text/xml;q=0.9, */*;q=0.8",
            },
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            ctype = resp.headers.get("Content-Type", "")
            raw_xml_str = _decode_bytes_smart(resp.read(), ctype)
        root = ET.fromstring(raw_xml_str.encode("utf-8"))

        entries = root.findall(".//item") or root.findall(".//{http://www.w3.org/2005/Atom}entry")
        for entry in entries[:8]:
            title_el = entry.find("title")
            if title_el is None:
                title_el = entry.find("{http://www.w3.org/2005/Atom}title")
            raw_title = html.unescape(title_el.text or "").strip() if title_el is not None else ""
            if not raw_title:
                continue

            link_el = entry.find("link")
            if link_el is None:
                link_el = entry.find("{http://www.w3.org/2005/Atom}link")
            article_url = ""
            if link_el is not None:
                article_url = (link_el.text or link_el.attrib.get("href") or "").strip()
            if not article_url.startswith(("http://", "https://")):
                article_url = fallback_site

            desc_el = entry.find("description")
            if desc_el is None:
                desc_el = entry.find("{http://www.w3.org/2005/Atom}summary")
            raw_desc = html.unescape(desc_el.text or "") if desc_el is not None else ""
            clean_desc = re.sub(r"<[^>]+>", " ", raw_desc)
            clean_desc = re.sub(r"\s+", " ", clean_desc).strip()

            if not _is_relevant_business_article(raw_title, clean_desc, czech_domestic_only=czech_domestic_only):
                continue

            items_out.append({
                "source_name": source_name,
                "badge": f"{source_name.upper()} | ŽIVÉ ZPRÁVY",
                "badge_color": badge_color,
                "raw_title": raw_title,
                "raw_text": clean_desc or raw_title,
                "url": article_url,
            })
            if len(items_out) >= 3:
                break
    except Exception as err:
        logger.warning("Nelze načíst RSS kanál %s: %s", source_name, err)
    return items_out


def _fetch_live_rss_articles_parallel(
    feed_sources: list[tuple[str, str, str]],
    badge_color: str,
    max_items: int = 5,
    czech_domestic_only: bool = False,
) -> list[dict]:
    collected: list[dict] = []
    seen_titles: set[str] = set()

    with ThreadPoolExecutor(max_workers=min(6, max(1, len(feed_sources)))) as pool:
        futures = [pool.submit(_fetch_single_rss_source, src, badge_color, czech_domestic_only) for src in feed_sources]
        for fut in as_completed(futures):
            try:
                batch = fut.result()
            except Exception:
                batch = []
            per_source = 0
            for item in batch:
                if len(collected) >= max_items or per_source >= 2:
                    break
                norm_key = re.sub(r"\W+", "", item["raw_title"].lower())[:60]
                if norm_key in seen_titles:
                    continue
                seen_titles.add(norm_key)
                per_source += 1
                collected.append(item)

    return collected[:max_items]


def build_curated_czech_news_sections(forex_data: dict, benchmarks: dict, cnb: dict) -> tuple[list[dict], list[dict]]:
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]
    us10y = benchmarks["us10y_yield"]
    sp500_pe = benchmarks["sp500_pe"]

    world_feeds = [
        ("CNBC Finance", "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=10000664", "https://www.cnbc.com/finance/"),
        ("CNBC Ekonomika", "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=20910258", "https://www.cnbc.com/economy/"),
        ("MarketWatch", "https://feeds.content.dowjones.io/public/rss/mw_topstories", "https://www.marketwatch.com/latest-news"),
        ("Yahoo Finance", "https://finance.yahoo.com/news/rssindex", "https://finance.yahoo.com/topic/stock-market-news/"),
    ]
    raw_world = _fetch_live_rss_articles_parallel(world_feeds, badge_color="#38bdf8", max_items=5, czech_domestic_only=False)
    processed_world = _translate_and_summarize_batch_czech(raw_world)

    world_sections: list[dict] = []
    for idx, item in enumerate(processed_world):
        t_clean = item.get("title_cs") or item["raw_title"]
        if _looks_english(t_clean):
            t_clean = _fallback_translate_en_to_cs(t_clean)
        world_sections.append({
            "badge": item["badge"],
            "badge_color": item["badge_color"],
            "title": f"{idx + 1}. {t_clean}",
            "body": item.get("summary_cs") or _clean_czech_news_summary(item["raw_text"], t_clean),
            "url": item["url"],
        })

    if len(world_sections) < 5:
        live_tickers = [
            ("INDEX S&P 500", "^GSPC", "https://finance.yahoo.com/quote/%5EGSPC"),
            ("INDEX NASDAQ COMPOSITE", "^IXIC", "https://finance.yahoo.com/quote/%5EIXIC"),
            ("ZLATO (TERMÍNOVÝ KONTRAKT)", "GC=F", "https://finance.yahoo.com/quote/GC%3DF"),
            ("ROPA WTI", "CL=F", "https://finance.yahoo.com/quote/CL%3DF"),
            ("MĚNOVÝ PÁR EUR / USD", "EURUSD=X", "https://finance.yahoo.com/quote/EURUSD%3DX"),
        ]
        for label, symbol, link in live_tickers:
            if len(world_sections) >= 5:
                break
            try:
                hist = yf.Ticker(symbol).history(period="5d", interval="1d")["Close"].dropna()
                if len(hist) >= 2:
                    last_px = float(hist.iloc[-1])
                    prev_px = float(hist.iloc[-2])
                    chg_pct = ((last_px - prev_px) / prev_px) * 100.0 if prev_px else 0.0
                    world_sections.append({
                        "badge": f"{label} | ŽIVÁ TRŽNÍ DATA",
                        "badge_color": "#38bdf8",
                        "title": f"{len(world_sections) + 1}. {label}: aktuální hodnota {last_px:,.2f} ({chg_pct:+.2f} %)",
                        "body": _clean_czech_news_summary(
                            f"Aktuální burzovní hodnota instrumentu {label} dosahuje {last_px:,.2f} při denní změně {chg_pct:+.2f} %. "
                            f"Bezrizikový výnos desetiletých amerických státních dluhopisů činí {us10y:.2f} %. "
                            f"Průměrné ocenění indexu S&P 500 se pohybuje kolem {sp500_pe:.1f}násobku zisku.",
                            f"{label}: aktuální hodnota {last_px:,.2f}",
                        ),
                        "url": link,
                    })
            except Exception:
                pass

    czech_feeds = [
        ("Česká tisková kancelář", "https://www.ceskenoviny.cz/sluzby/rss/ekonomika.php", "https://www.ceskenoviny.cz/ekonomika/"),
        ("Deník E15", "https://www.e15.cz/rss", "https://www.e15.cz/"),
        ("Kurzy.cz Ekonomika", "https://www.kurzy.cz/zpravy/util/rss.asp", "https://www.kurzy.cz/zpravy/"),
    ]
    raw_cz = _fetch_live_rss_articles_parallel(czech_feeds, badge_color="#f59e0b", max_items=5, czech_domestic_only=True)
    processed_cz = _translate_and_summarize_batch_czech(raw_cz)

    czech_sections: list[dict] = []
    seen_cz: set[str] = set()
    for item in processed_cz:
        if len(czech_sections) >= 5:
            break
        raw_heading = re.sub(r"^\d+\.\s*", "", item.get("title_cs") or item["raw_title"]).strip()
        dedup_key = re.sub(r"\W+", "", raw_heading.lower())[:55]
        if dedup_key in seen_cz:
            continue
        seen_cz.add(dedup_key)
        czech_sections.append({
            "badge": item["badge"],
            "badge_color": item["badge_color"],
            "title": f"{len(czech_sections) + 1}. {raw_heading}",
            "body": _clean_czech_news_summary(item.get("summary_cs") or item["raw_text"], raw_heading),
            "url": item["url"],
        })

    if len(czech_sections) < 5:
        live_cz_cards = [
            {
                "badge": "ČESKÁ NÁRODNÍ BANKA | ŽIVÉ SAZBY",
                "badge_color": "#f59e0b",
                "title": f"Česká národní banka drží základní repo sazbu na {cnb['repo_rate']} při inflaci {cnb['latest_inflation']}",
                "body": _clean_czech_news_summary(
                    f"Aktuální základní dvoutýdenní repo sazba činí {cnb['repo_rate']}. "
                    f"Diskontní sazba dosahuje {cnb['discount_rate']} a lombardní sazba {cnb['lombard_rate']}. "
                    f"Meziroční míra inflace v České republice se nachází na úrovni {cnb['latest_inflation']}.",
                    "Základní repo sazba ČNB a inflace",
                ),
                "url": "https://www.cnb.cz/cs/menova-politika/",
            },
            {
                "badge": "KURZY.CZ | DEVIZOVÝ TRH",
                "badge_color": "#f59e0b",
                "title": f"Aktuální kurz české koruny: {eur_czk:.2f} za euro a {usd_czk:.2f} za americký dolar",
                "body": _clean_czech_news_summary(
                    f"Na devizovém trhu se česká koruna obchoduje za {eur_czk:.2f} CZK/EUR. "
                    f"Kurz vůči americkému dolaru činí {usd_czk:.2f} CZK/USD a vůči libře {gbp_czk:.2f} CZK/GBP. "
                    f"Kurzový vývoj ovlivňuje úrokový diferenciál i globální tržní nálada.",
                    "Aktuální kurz české koruny",
                ),
                "url": "https://www.kurzy.cz/kurzy-men/",
            },
        ]
        for fallback_card in live_cz_cards:
            if len(czech_sections) >= 5:
                break
            fallback_card["title"] = f"{len(czech_sections) + 1}. {fallback_card['title']}"
            czech_sections.append(fallback_card)

    for sec in world_sections:
        sec["body"] = _clean_czech_news_summary(sec["body"], sec["title"])
    for sec in czech_sections:
        sec["body"] = _clean_czech_news_summary(sec["body"], sec["title"])

    return world_sections, czech_sections


def build_newsletter_pdf_bytes(forex_data: dict, benchmarks: dict, cnb: dict, world_sections: list[dict], czech_sections: list[dict]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("TTitle", parent=styles["Heading1"], fontName=PDF_FONT_BOLD, fontSize=16, textColor=colors.HexColor("#0f172a"), spaceAfter=4)
    subtitle_style = ParagraphStyle("TSub", parent=styles["Normal"], fontName=PDF_FONT_REGULAR, fontSize=9.5, textColor=colors.HexColor("#475569"), spaceAfter=14)
    section_style = ParagraphStyle("SHead", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=11.5, textColor=colors.HexColor("#0284c7"), spaceBefore=12, spaceAfter=6)
    item_title_style = ParagraphStyle("ITitle", parent=styles["Heading3"], fontName=PDF_FONT_BOLD, fontSize=10.5, textColor=colors.HexColor("#0f172a"), spaceBefore=8, spaceAfter=3)
    body_style = ParagraphStyle("BText", parent=styles["Normal"], fontName=PDF_FONT_REGULAR, fontSize=9.2, leading=13.5, textColor=colors.HexColor("#1e293b"), spaceAfter=8)

    now_str = get_prague_now().strftime("%d. %m. %Y | %H:%M:%S CET")
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]

    elements = []
    logo_path = resolve_brand_logo_path(prefer_white=False)
    if logo_path:
        try:
            logo_img = PdfImage(logo_path, width=38, height=26, kind="proportional")
            header_tbl = Table([[logo_img, Paragraph("TENSEGRITY TERMINAL | INTELLIGENCE REPORT", title_style)]], colWidths=[48, 470])
            header_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
            elements.append(header_tbl)
        except Exception:
            elements.append(Paragraph("TENSEGRITY TERMINAL | INTELLIGENCE REPORT", title_style))
    else:
        elements.append(Paragraph("TENSEGRITY TERMINAL | INTELLIGENCE REPORT", title_style))

    elements.extend([
        Paragraph(f"Datum a čas vydání: {now_str} | Kompletní český makroekonomický přehled", subtitle_style),
        Paragraph("1. GLOBÁLNÍ MAKROEKONOMICKÝ RÁMEC & MĚNOVÁ POLITIKA", section_style),
        Paragraph(
            f"Výnos desetiletých amerických státních dluhopisů dosahuje {benchmarks['us10y_yield']:.2f} % a průměrný poměr ceny "
            f"a zisku indexu S&P 500 činí {benchmarks['sp500_pe']:.1f}x.",
            body_style,
        ),
        Paragraph("2. ČESKÁ REPUBLIKA — SAZBY ČNB, INFLACE A MĚNOVÉ KURZY", section_style),
    ])

    fx_table = Table(
        [
            ["Repo sazba ČNB", "Inflace ČR", "EUR / CZK", "USD / CZK", "GBP / CZK"],
            [cnb["repo_rate"], cnb["latest_inflation"], f"{eur_czk:.2f} CZK", f"{usd_czk:.2f} CZK", f"{gbp_czk:.2f} CZK"],
        ],
        colWidths=[102, 102, 105, 105, 105],
    )
    fx_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), PDF_FONT_BOLD),
        ("FONTNAME", (0, 1), (-1, -1), PDF_FONT_REGULAR),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(fx_table)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("3. PĚT KLÍČOVÝCH ZPRÁV ZE SVĚTA (ČESKÝ PŘEKLAD A SHRNUTÍ)", section_style))
    for item in world_sections:
        elements.append(Paragraph(f"{item['title']} [{item['badge']}]", item_title_style))
        elements.append(Paragraph(item["body"], body_style))

    elements.append(Paragraph("4. PĚT KLÍČOVÝCH ZPRÁV Z ČESKÉ REPUBLIKY", section_style))
    for item in czech_sections:
        elements.append(Paragraph(f"{item['title']} [{item['badge']}]", item_title_style))
        elements.append(Paragraph(item["body"], body_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer.read()


def generate_html_newsletter_content(
    forex_data: dict,
    benchmarks: dict,
    cnb: dict,
    world_sections: list[dict],
    czech_sections: list[dict],
) -> str:
    now_str = get_prague_now().strftime("%d. %m. %Y | %H:%M CET")
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]
    us10y = benchmarks["us10y_yield"]
    sp500_pe = benchmarks["sp500_pe"]

    macro_summary_3to4 = _clean_czech_news_summary(
        f"Americká centrální banka (Fed) i Evropská centrální banka (ECB) udržují obezřetný postoj vůči vývoji inflace. "
        f"Výnos desetiletých amerických státních dluhopisů se pohybuje kolem {us10y:.2f} %. "
        f"Průměrný poměr ceny a zisku indexu S&P 500 činí {sp500_pe:.1f}násobku, zatímco základní repo sazba ČNB je na {cnb['repo_rate']}.",
        "Globální makro, centrální banky a tržní ukazatele",
    )
    cz_macro_summary_3to4 = _clean_czech_news_summary(
        f"Meziroční inflace v České republice dosahuje {cnb['latest_inflation']} a základní dvoutýdenní repo sazba ČNB činí {cnb['repo_rate']}. "
        f"Na devizovém trhu se koruna obchoduje za {eur_czk:.2f} CZK/EUR a {usd_czk:.2f} CZK/USD. "
        f"Kurz libry vůči koruně se pohybuje na úrovni {gbp_czk:.2f} CZK/GBP.",
        "Česká republika — inflace, sazby ČNB a měnové kurzy",
    )

    def render_cards_table(sections: list[dict]) -> str:
        rows_html = ""
        for item in sections:
            is_blue = item["badge_color"] == "#38bdf8"
            badge_bg = "#132840" if is_blue else "#332410"
            badge_fg = "#38bdf8" if is_blue else "#f59e0b"
            safe_title = html.escape(item["title"])
            safe_body = html.escape(_clean_czech_news_summary(item["body"], item["title"]))
            safe_badge = html.escape(item["badge"])
            safe_url = html.escape(item["url"], quote=True)
            rows_html += f"""
            <tr>
                <td style="padding: 0 0 12px 0;">
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px;">
                        <tr>
                            <td style="padding: 16px 18px;">
                                <span style="display: inline-block; font-size: 10px; font-weight: 700; letter-spacing: 0.6px; background-color: {badge_bg}; color: {badge_fg}; padding: 3px 8px; border-radius: 4px; font-family: Consolas, monospace;">{safe_badge}</span>
                                <div style="color: #f8fafc; font-size: 15px; font-weight: 700; margin: 8px 0 6px 0; line-height: 1.4;">{safe_title}</div>
                                <div style="color: #cbd5e1; font-size: 13px; line-height: 1.6; margin: 0 0 10px 0;">{safe_body}</div>
                                <a href="{safe_url}" target="_blank" rel="noopener" style="color: #38bdf8; font-size: 12px; font-weight: 600; text-decoration: none;">Oficiální zdroj &#8599;</a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            """
        return f'<table role="presentation" width="100%" cellspacing="0" cellpadding="0">{rows_html}</table>'

    world_cards_html = render_cards_table(world_sections)
    czech_cards_html = render_cards_table(czech_sections)

    return f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tensegrity Terminal | Intelligence</title>
</head>
<body style="margin: 0; padding: 0; background-color: #04070d; font-family: Arial, Helvetica, sans-serif; color: #f8fafc;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color: #04070d; padding: 16px 8px;">
        <tr>
            <td align="center">
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width: 720px; background-color: #060911; border: 1px solid #1e293b; border-radius: 12px;">
                    <tr>
                        <td style="padding: 24px 22px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-bottom: 1px solid #1e293b; padding-bottom: 14px; margin-bottom: 18px;">
                                <tr>
                                    <td align="left" style="font-size: 16px; font-weight: 800; letter-spacing: 2px; color: #ffffff;">
                                        TENSEGRITY<span style="color: #38bdf8;">TERMINAL</span>
                                    </td>
                                    <td align="right" style="font-size: 11px; font-family: Consolas, monospace; color: #94a3b8;">
                                        {now_str}
                                    </td>
                                </tr>
                            </table>

                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color: #0b1322; border: 1px solid #1d4ed8; border-radius: 8px; margin-bottom: 20px;">
                                <tr>
                                    <td style="padding: 10px 14px; text-align: center; font-size: 12px; color: #38bdf8; font-weight: 600;">
                                        Kompletní přehled je zobrazen níže a zároveň přiložen v PDF (<span style="color: #ffffff;">Tensegrity_Terminal_Newsletter.pdf</span>).
                                    </td>
                                </tr>
                            </table>

                            <div style="color: #38bdf8; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #1e293b; padding-bottom: 6px; margin-bottom: 12px;">
                                1. GLOBÁLNÍ MAKRO, CENTRÁLNÍ BANKY A TRŽNÍ UKAZATELE
                            </div>
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px; margin-bottom: 20px;">
                                <tr>
                                    <td style="padding: 16px;">
                                        <div style="color: #cbd5e1; font-size: 13px; line-height: 1.55; margin-bottom: 12px;">
                                            {html.escape(macro_summary_3to4)}
                                        </div>
                                        <table role="presentation" width="100%" cellspacing="6" cellpadding="0">
                                            <tr>
                                                <td width="33%" style="background-color: #080d18; border: 1px solid #1e293b; border-radius: 6px; padding: 10px; text-align: center;">
                                                    <div style="font-size: 10px; color: #94a3b8;">VÝNOS US 10Y</div>
                                                    <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: Consolas, monospace; margin-top: 3px;">{us10y:.2f} %</div>
                                                </td>
                                                <td width="33%" style="background-color: #080d18; border: 1px solid #1e293b; border-radius: 6px; padding: 10px; text-align: center;">
                                                    <div style="font-size: 10px; color: #94a3b8;">P/E S&amp;P 500</div>
                                                    <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: Consolas, monospace; margin-top: 3px;">{sp500_pe:.1f}x</div>
                                                </td>
                                                <td width="33%" style="background-color: #080d18; border: 1px solid #1e293b; border-radius: 6px; padding: 10px; text-align: center;">
                                                    <div style="font-size: 10px; color: #94a3b8;">REPO SAZBA ČNB</div>
                                                    <div style="font-size: 15px; font-weight: 700; color: #f59e0b; font-family: Consolas, monospace; margin-top: 3px;">{cnb['repo_rate']}</div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>

                            <div style="color: #f59e0b; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #1e293b; padding-bottom: 6px; margin-bottom: 12px;">
                                2. ČESKÁ REPUBLIKA — INFLACE, SAZBY ČNB A MĚNOVÉ KURZY
                            </div>
                            <table role="presentation" width="100%" cellspacing="6" cellpadding="0" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px; padding: 14px; margin-bottom: 8px;">
                                <tr>
                                    <td style="color: #cbd5e1; font-size: 13px; line-height: 1.55; padding-bottom: 10px;">
                                        {html.escape(cz_macro_summary_3to4)}
                                    </td>
                                </tr>
                            </table>
                            <table role="presentation" width="100%" cellspacing="6" cellpadding="0" style="margin-bottom: 20px;">
                                <tr>
                                    <td width="33%" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px; padding: 12px; text-align: center;">
                                        <div style="font-size: 10px; color: #94a3b8;">EUR / CZK</div>
                                        <div style="font-size: 16px; font-weight: 800; color: #ffffff; font-family: Consolas, monospace; margin-top: 4px;">{eur_czk:.2f} CZK</div>
                                    </td>
                                    <td width="33%" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px; padding: 12px; text-align: center;">
                                        <div style="font-size: 10px; color: #94a3b8;">USD / CZK</div>
                                        <div style="font-size: 16px; font-weight: 800; color: #ffffff; font-family: Consolas, monospace; margin-top: 4px;">{usd_czk:.2f} CZK</div>
                                    </td>
                                    <td width="33%" style="background-color: #0f1625; border: 1px solid #1e293b; border-radius: 8px; padding: 12px; text-align: center;">
                                        <div style="font-size: 10px; color: #94a3b8;">GBP / CZK</div>
                                        <div style="font-size: 16px; font-weight: 800; color: #ffffff; font-family: Consolas, monospace; margin-top: 4px;">{gbp_czk:.2f} CZK</div>
                                    </td>
                                </tr>
                            </table>

                            <div style="color: #38bdf8; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #1e293b; padding-bottom: 6px; margin-bottom: 12px;">
                                3. GLOBÁLNÍ NOVINKY &amp; ZPRÁVY ZE SVĚTA (PŘELOŽENO DO ČEŠTINY)
                            </div>
                            {world_cards_html}

                            <div style="color: #f59e0b; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #1e293b; padding-bottom: 6px; margin: 16px 0 12px 0;">
                                4. AKTUÁLNÍ ZPRÁVY Z ČESKÉ REPUBLIKY
                            </div>
                            {czech_cards_html}

                            <div style="text-align: center; font-size: 11px; color: #64748b; border-top: 1px solid #1e293b; padding-top: 14px; margin-top: 16px;">
                                Tensegrity Terminal | Intelligence Platform &bull; Automatický institucionální přehled
                            </div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""


def _send_via_gmail_smtp_sync(recipient: str, subject: str, html_content: str, plain_text: str, pdf_bytes: bytes | None):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise RuntimeError("Chybí SENDER_EMAIL nebo SENDER_PASSWORD pro Gmail SMTP.")

    msg = MIMEMultipart("mixed")
    msg["From"] = formataddr(("Tensegrity Terminal", SENDER_EMAIL))
    msg["To"] = recipient
    msg["Subject"] = subject

    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(plain_text, "plain", "utf-8"))
    alt_part.attach(MIMEText(html_content, "html", "utf-8"))
    msg.attach(alt_part)

    if pdf_bytes:
        pdf_part = MIMEBase("application", "pdf")
        pdf_part.set_payload(pdf_bytes)
        encoders.encode_base64(pdf_part)
        pdf_part.add_header("Content-Disposition", 'attachment; filename="Tensegrity_Terminal_Newsletter.pdf"')
        msg.attach(pdf_part)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=20) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, [recipient], msg.as_string())


async def send_email_report(recipient: str, subject: str, html_content: str, pdf_bytes: bytes | None = None):
    plain_text = (
        "Tensegrity Terminal | Intelligence Report.\n"
        "Kompletní přehled trhů, měnových kurzů a zpráv v češtině je přiložen v HTML těle zprávy i v PDF."
    )

    smtp_error = None
    if SENDER_EMAIL and SENDER_PASSWORD:
        try:
            await asyncio.to_thread(_send_via_gmail_smtp_sync, recipient, subject, html_content, plain_text, pdf_bytes)
            logger.info("E-mail úspěšně odeslán přes Gmail SMTP na adresu %s.", recipient)
            return {"provider": "gmail_smtp", "recipient": recipient}
        except Exception as err:
            smtp_error = err
            logger.warning("Gmail SMTP odeslání selhalo (%s), zkouším Resend API...", err)

    if resend is not None and RESEND_API_KEY:
        resend.api_key = RESEND_API_KEY
        params = {
            "from": RESEND_FROM_EMAIL,
            "to": [recipient],
            "subject": subject,
            "html": html_content,
            "text": plain_text,
        }
        if pdf_bytes:
            params["attachments"] = [
                {
                    "filename": "Tensegrity_Terminal_Newsletter.pdf",
                    "content": base64.b64encode(pdf_bytes).decode("utf-8"),
                }
            ]
        resp = await asyncio.to_thread(resend.Emails.send, params)
        logger.info("E-mail úspěšně odeslán přes Resend API na adresu %s.", recipient)
        return resp

    raise RuntimeError(f"Odeslání e-mailu selhalo. Gmail SMTP chyba: {smtp_error}")


async def dispatch_newsletter_to_recipients(recipients: list[str]) -> dict:
    if not recipients:
        return {"sent": 0, "failed": 0, "errors": []}

    forex_data = await asyncio.to_thread(get_live_forex_rates)
    benchmarks = await asyncio.to_thread(get_live_market_benchmarks)
    cnb = await asyncio.to_thread(get_cnb_macro_indicators)
    world_sections, czech_sections = await asyncio.to_thread(build_curated_czech_news_sections, forex_data, benchmarks, cnb)

    html_body = generate_html_newsletter_content(forex_data, benchmarks, cnb, world_sections, czech_sections)
    pdf_bytes = await asyncio.to_thread(build_newsletter_pdf_bytes, forex_data, benchmarks, cnb, world_sections, czech_sections)
    subject = f"Tensegrity Terminal | Tržní přehled ({get_prague_now().strftime('%d. %m. %Y')})"

    sent_count = 0
    errors: list[str] = []
    for recipient in recipients:
        try:
            await send_email_report(recipient, subject, html_body, pdf_bytes)
            sent_count += 1
        except Exception as err:
            logger.error("Odeslání zpravodaje na %s selhalo: %s", recipient, err)
            errors.append(f"{recipient}: {err}")

    return {"sent": sent_count, "failed": len(errors), "errors": errors}


async def scheduled_newsletter_worker():
    sent_slots = load_sent_slots()
    schedule_rules = {(2, 7), (4, 7), (6, 19)}
    while True:
        try:
            now = get_prague_now()
            slot_key = f"{now.strftime('%Y-%m-%d')}-{now.hour}"
            if (now.weekday(), now.hour) in schedule_rules and slot_key not in sent_slots:
                subscribers = get_subscribers_internal()
                if subscribers:
                    logger.info("Spouštím pravidelné odeslání zpravodaje pro %d odběratelů.", len(subscribers))
                    res = await dispatch_newsletter_to_recipients(subscribers)
                    if res["sent"] > 0:
                        sent_slots.add(slot_key)
                        save_sent_slots(sent_slots)
                else:
                    sent_slots.add(slot_key)
                    save_sent_slots(sent_slots)
        except Exception as err:
            logger.error("Chyba v plánovači zpravodaje: %s", err)
        await asyncio.sleep(45)


@asynccontextmanager
async def lifespan(app_instance: FastAPI):
    task = asyncio.create_task(scheduled_newsletter_worker())
    yield
    task.cancel()


app = FastAPI(title="Tensegrity Terminal | Institutional Valuation Platform", version="5.2", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ==============================================================================
# 6. VALUAČNÍ ENGINE BEZ TICHÉHO DOMÝŠLENÍ + OWNER EARNINGS + LOOK-THROUGH + SOTP
# ==============================================================================

ALL_SECTOR_BENCHMARKS = {
    "Technology": {"ev_ebitda": 20.5, "pe": 27.0, "ev_revenue": 8.0, "pb": 9.0, "growth": 0.12},
    "Communication Services": {"ev_ebitda": 13.5, "pe": 20.0, "ev_revenue": 3.8, "pb": 4.2, "growth": 0.095},
    "Healthcare": {"ev_ebitda": 15.5, "pe": 21.0, "ev_revenue": 4.2, "pb": 4.0, "growth": 0.085},
    "Consumer Cyclical": {"ev_ebitda": 13.0, "pe": 19.5, "ev_revenue": 1.9, "pb": 3.8, "growth": 0.08},
    "Industrials": {"ev_ebitda": 12.8, "pe": 18.5, "ev_revenue": 2.1, "pb": 3.3, "growth": 0.07},
    "Financial Services": {"ev_ebitda": 10.0, "pe": 12.0, "ev_revenue": 3.0, "pb": 1.30, "growth": 0.055},
    "Consumer Defensive": {"ev_ebitda": 12.8, "pe": 19.5, "ev_revenue": 1.9, "pb": 4.0, "growth": 0.05},
    "Utilities": {"ev_ebitda": 9.8, "pe": 13.8, "ev_revenue": 2.2, "pb": 1.50, "growth": 0.04},
    "Energy": {"ev_ebitda": 6.8, "pe": 10.2, "ev_revenue": 1.20, "pb": 1.60, "growth": 0.04},
    "Real Estate": {"ev_ebitda": 14.2, "pe": 21.5, "ev_revenue": 5.8, "pb": 1.70, "growth": 0.045},
    "Basic Materials": {"ev_ebitda": 8.0, "pe": 12.8, "ev_revenue": 1.50, "pb": 1.80, "growth": 0.05},
}


def get_currency_risk_free_rate(currency: str, us10y_pct: float) -> float:
    curr = (currency or "USD").upper()
    if curr == "CZK":
        try:
            cnb = get_cnb_macro_indicators()
            repo_num = float(cnb["repo_rate"].replace("%", "").replace(",", ".").strip())
            if 0.5 < repo_num < 15.0:
                return repo_num / 100.0
        except Exception:
            pass
        return 0.040
    if curr == "EUR":
        return max(0.022, (us10y_pct / 100.0) - 0.014)
    if curr == "GBP":
        return max(0.030, (us10y_pct / 100.0) - 0.002)
    return us10y_pct / 100.0


def compute_dynamic_scenario_defaults(
    info: dict,
    series: dict[str, list[float]],
    points: dict[str, DataPoint],
    sector: str,
    market_currency: str,
) -> tuple[float, float, float]:
    bench = ALL_SECTOR_BENCHMARKS.get(sector, {"ev_ebitda": 13.0, "pe": 19.0, "ev_revenue": 3.0, "pb": 3.0, "growth": 0.075})
    current_price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"), 0.0) or 0.0
    mcap_val = points["market_cap"].value or 0.0
    debt_val = points["debt"].value if points["debt"].value is not None else 0.0

    total_cap = mcap_val + debt_val
    if total_cap > 0 and points["debt"].value is not None:
        auto_leverage = round(max(0.0, min(0.85, debt_val / total_cap)), 2)
    else:
        auto_leverage = 0.15

    sector_base_g = bench["growth"]
    rev_growth = safe_float(info.get("revenueGrowth"), None)
    earn_growth = safe_float(info.get("earningsGrowth"), None)
    earn_q_growth = safe_float(info.get("earningsQuarterlyGrowth"), None)

    trailing_pe = safe_float(info.get("trailingPE"), 0.0) or 0.0
    forward_pe = safe_float(info.get("forwardPE"), 0.0) or 0.0
    forward_implied_g = None
    if 3.0 < forward_pe < 120.0 and 3.0 < trailing_pe < 150.0:
        raw_fwd_g = (trailing_pe / forward_pe) - 1.0
        if -0.35 < raw_fwd_g < 1.20:
            forward_implied_g = raw_fwd_g

    rev_hist = series.get("revenue") or []
    hist_cagr = None
    if len(rev_hist) >= 2 and rev_hist[-1] > 0 and rev_hist[0] > 0:
        years = len(rev_hist) - 1
        hist_cagr = (rev_hist[0] / rev_hist[-1]) ** (1.0 / years) - 1.0

    backward_signals = []
    if rev_growth is not None and -0.45 < rev_growth < 2.5:
        backward_signals.append(rev_growth)
    if earn_growth is not None and -0.45 < earn_growth < 2.5:
        backward_signals.append(earn_growth * 0.70)
    elif earn_q_growth is not None and -0.45 < earn_q_growth < 2.5:
        backward_signals.append(earn_q_growth * 0.65)
    if hist_cagr is not None and -0.45 < hist_cagr < 2.5:
        backward_signals.append(hist_cagr)

    usd_rate = get_fx_conversion_rate(market_currency, "USD")
    mcap_usd = mcap_val * usd_rate

    if mcap_usd >= 1_000_000_000_000:
        size_dampener = 0.42
        max_cap = 0.175
    elif mcap_usd >= 250_000_000_000:
        size_dampener = 0.55
        max_cap = 0.195
    elif mcap_usd >= 30_000_000_000:
        size_dampener = 0.72
        max_cap = 0.22
    else:
        size_dampener = 0.88
        max_cap = 0.26

    high_52w = safe_float(info.get("fiftyTwoWeekHigh"), 0.0) or 0.0
    drawdown_factor = 1.0
    if high_52w > 0 and current_price > 0 and current_price < high_52w:
        drawdown_pct = (high_52w - current_price) / high_52w
        if drawdown_pct > 0.25:
            drawdown_factor = max(0.75, 1.0 - (drawdown_pct - 0.25) * 0.5)

    if backward_signals or forward_implied_g is not None:
        back_avg = (sum(backward_signals) / len(backward_signals)) if backward_signals else sector_base_g
        capped_back = max(-0.20, min(0.35, back_avg))
        if forward_implied_g is not None:
            combined_signal = 0.50 * forward_implied_g + 0.50 * capped_back
        else:
            combined_signal = capped_back

        raw_blended = (0.32 * size_dampener) * combined_signal + (1.0 - 0.32 * size_dampener) * sector_base_g
        raw_blended *= drawdown_factor
        auto_growth = round(max(-0.08, min(min(max_cap, 0.18), raw_blended)), 3)
    else:
        auto_growth = sector_base_g

    sector_base_mult = bench["ev_ebitda"]
    ev_to_ebitda = safe_float(info.get("enterpriseToEbitda"), 0.0) or 0.0
    profit_margin = safe_float(info.get("profitMargins"), 0.0) or 0.0
    margin_adj = max(0.85, min(1.15, 1.0 + (profit_margin - 0.15) * 0.45))

    if 3.0 < ev_to_ebitda < 55.0:
        blended_mult = (0.22 * ev_to_ebitda + 0.78 * sector_base_mult * margin_adj) * drawdown_factor
        auto_exit_mult = round(max(5.0, min(22.0, blended_mult)), 1)
    else:
        auto_exit_mult = round(max(5.0, min(22.0, sector_base_mult * margin_adj * drawdown_factor)), 1)

    return auto_leverage, auto_growth, auto_exit_mult


def run_valuation_logic(
    query_or_ticker: str,
    leverage: float | None = None,
    growth: float | None = None,
    exit_mult: float | None = None,
) -> dict:
    clean_ticker = resolve_company_ticker(query_or_ticker)
    dataset = fetch_unified_financial_dataset(clean_ticker)

    info = dataset["info"]
    currency = dataset["market_currency"]
    fin_currency = dataset["stmt_currency"]
    stmt_fx = dataset["stmt_fx"]
    stmt_provider = dataset["stmt_provider"]
    current_price = dataset["current_price"]
    series = dataset["series"]
    pts: dict[str, DataPoint] = dataset["points"]

    sector = str(info.get("sector") or "Technology")
    industry = str(info.get("industry") or "General")
    is_financial = sector == "Financial Services" or any(
        k in industry.lower() for k in ("bank", "insurance", "financial", "asset management", "capital markets")
    )
    is_real_estate = sector == "Real Estate" or "reit" in industry.lower()

    auto_lev, auto_grw, auto_mult = compute_dynamic_scenario_defaults(info, series, pts, sector, currency)
    eff_leverage = leverage if leverage is not None else auto_lev
    eff_growth = growth if growth is not None else auto_grw
    eff_exit_mult = exit_mult if exit_mult is not None else auto_mult

    benchmarks = get_live_market_benchmarks()
    rf_rate = get_currency_risk_free_rate(currency, benchmarks["us10y_yield"])
    erp_rate = 0.052 if currency == "CZK" else 0.046

    exchange = str(info.get("exchange") or "NMS")
    company_name = dataset["company_name"]
    company_logo_urls = extract_company_logo_urls(info, clean_ticker)
    company_logo_url = company_logo_urls[0] if company_logo_urls else ""
    country = str(info.get("country") or "United States")
    employees = int(safe_float(info.get("fullTimeEmployees"), 0) or 0)
    business_summary = str(info.get("longBusinessSummary") or "")

    shares_outstanding = pts["shares"].value
    market_cap = pts["market_cap"].value

    data_warnings: list[str] = []
    missing_fields: list[str] = []
    estimated_fields: list[str] = []

    field_labels_cs = {
        "revenue": "Tržby",
        "ebit": "Provozní zisk (EBIT)",
        "ebitda": "EBITDA",
        "da": "Odpisy a amortizace (D&A)",
        "net_income": "Čistý zisk",
        "ocf": "Provozní cash flow (OCF)",
        "capex": "Kapitálové výdaje (CapEx)",
        "sbc": "Akciové odměny (SBC)",
        "equity": "Vlastní kapitál",
        "debt": "Celkový dluh",
        "cash": "Hotovost a ekvivalenty",
        "interest_expense": "Úrokové náklady",
    }
    for f_key, f_label in field_labels_cs.items():
        dp_item = pts[f_key]
        if dp_item.status == "MISSING":
            missing_fields.append(f_label)
        elif dp_item.status == "ESTIMATED":
            estimated_fields.append(f_label)

    if pts["shares"].note:
        data_warnings.append(pts["shares"].note)

    if "SEC EDGAR" not in stmt_provider and "Financial Modeling Prep" not in stmt_provider:
        data_warnings.append(
            "Účetní výkazy pocházejí ze záložního zdroje (Yahoo Finance), nikoli přímo z regulatorního podání."
        )
    elif "S-1" in stmt_provider or "10-Q" in stmt_provider:
        data_warnings.append(
            f"Čerstvě kotovaná společnost: data vycházejí z regulatorního podání {stmt_provider} (před prvním výročním 10-K)."
        )

    if missing_fields:
        data_warnings.append(
            "Chybějící položky výkazů (nebyly uměle dopočítány): " + ", ".join(missing_fields) + "."
        )

    revenue = pts["revenue"].value
    reported_ebit = pts["ebit"].value
    reported_ebitda = pts["ebitda"].value
    da_val = pts["da"].value
    net_income = pts["net_income"].value
    raw_ocf = pts["ocf"].value
    raw_capex = pts["capex"].value
    sbc_expense = pts["sbc"].value
    book_equity = pts["equity"].value
    total_debt = pts["debt"].value
    total_cash = pts["cash"].value
    minority_interest = pts["minority_interest"].value or 0.0
    st_investments = pts["short_term_investments"].value or 0.0
    lt_investments = pts["long_term_investments"].value or 0.0
    deferred_tax_liab = pts["deferred_tax_liab"].value or 0.0
    insurance_float = pts["insurance_float"].value or 0.0

    if total_debt is not None and total_cash is not None:
        net_debt = total_debt - total_cash - st_investments
        enterprise_value = market_cap + total_debt + minority_interest - total_cash - st_investments
    else:
        net_debt = None
        enterprise_value = None

    pretax_val = pts["pretax_income"].value
    tax_val = pts["tax_provision"].value
    if pretax_val is not None and tax_val is not None and pretax_val > 0 and 0 <= tax_val <= pretax_val * 0.50:
        effective_tax_rate = max(0.05, min(0.35, tax_val / pretax_val))
        tax_rate_status = "REPORTED"
    else:
        effective_tax_rate = 0.21
        tax_rate_status = "STATUTORY_21PCT"

    maintenance_capex_dp: DataPoint
    growth_capex_dp: DataPoint
    owner_earnings_dp: DataPoint

    if raw_capex is not None:
        if da_val is not None and da_val > 0:
            maint_capex = min(raw_capex, da_val * 1.05)
            growth_capex = max(0.0, raw_capex - maint_capex)
            maintenance_capex_dp = DataPoint(
                value=maint_capex,
                status="ESTIMATED",
                source="Pravidlo min(CapEx, 1,05 × D&A)",
                note="Odhad kapitálu nutného k udržení stávající konkurenční pozice a objemu.",
            )
            growth_capex_dp = DataPoint(
                value=growth_capex,
                status="ESTIMATED",
                source="CapEx - Maintenance CapEx",
                note="Expanzní kapitálové výdaje nad rámec obnovy.",
            )
        else:
            maint_capex = raw_capex
            growth_capex = 0.0
            maintenance_capex_dp = DataPoint(
                value=maint_capex,
                status="ESTIMATED",
                source="100 % reportovaného CapEx (chybí D&A)",
                note="Bez údajů o odpisech se celý CapEx konzervativně považuje za udržovací.",
            )
            growth_capex_dp = DataPoint(value=0.0, status="ESTIMATED", source="Konzervativní předpoklad")
    else:
        maint_capex = None
        growth_capex = None
        maintenance_capex_dp = DataPoint(value=None, status="MISSING", source=stmt_provider, note="Chybí reportovaný CapEx.")
        growth_capex_dp = DataPoint(value=None, status="MISSING", source=stmt_provider, note="Chybí reportovaný CapEx.")

    if net_income is not None and maint_capex is not None:
        da_addback = da_val if da_val is not None else 0.0
        sbc_deduct = sbc_expense if sbc_expense is not None else 0.0
        oe_val = net_income + da_addback - sbc_deduct - maint_capex
        owner_earnings_dp = DataPoint(
            value=oe_val,
            status="ESTIMATED",
            source="Net Income + D&A - SBC - Maintenance CapEx",
            note="Skutečný peněžní tok dostupný vlastníkům po odečtení obnovy aktiv a SBC.",
        )
    elif raw_ocf is not None and maint_capex is not None:
        sbc_deduct = sbc_expense if sbc_expense is not None else 0.0
        oe_val = raw_ocf - maint_capex - sbc_deduct
        owner_earnings_dp = DataPoint(
            value=oe_val,
            status="ESTIMATED",
            source="OCF - Maintenance CapEx - SBC",
            note="Odvozeno z provozního cash flow po odečtení udržovacího CapEx.",
        )
    else:
        owner_earnings_dp = DataPoint(
            value=None,
            status="MISSING",
            source=stmt_provider,
            note="Nelze spočítat Owner Earnings (chybí čistý zisk/OCF nebo CapEx).",
        )

    if net_income is not None:
        if lt_investments > 0:
            undistributed_investee_earnings = lt_investments * 0.075 * (1.0 - 0.15)
            look_through_val = net_income + undistributed_investee_earnings
            look_through_dp = DataPoint(
                value=look_through_val,
                status="ESTIMATED",
                source="Net Income + Look-Through z dlouhodobých investic",
                note="Zahrnuje odhadovaný nerozdělený zisk z majetkových účastí po zdanění.",
            )
        else:
            look_through_dp = DataPoint(
                value=net_income,
                status="REPORTED",
                source=stmt_provider,
                note="Firma nevykazuje významné neprovozní majetkové účasti; shodné s čistým ziskem.",
            )
    else:
        look_through_dp = DataPoint(value=None, status="MISSING", source=stmt_provider)

    if raw_ocf is not None and raw_capex is not None:
        if is_real_estate and net_income is not None and da_val is not None:
            reported_fcf = net_income + da_val - maint_capex
        else:
            reported_fcf = raw_ocf - raw_capex
        fcf_dp = DataPoint(value=reported_fcf, status="REPORTED", source=f"{stmt_provider} (OCF - CapEx)")
    else:
        raw_fcf_info = safe_float(info.get("freeCashflow"), None)
        if raw_fcf_info is not None:
            reported_fcf = raw_fcf_info * stmt_fx
            fcf_dp = DataPoint(value=reported_fcf, status="REPORTED", source="Market Feed (FCF)")
        else:
            reported_fcf = None
            fcf_dp = DataPoint(value=None, status="MISSING", source=stmt_provider, note="Chybí OCF nebo CapEx pro výpočet FCF.")

    sbc_known = sbc_expense if sbc_expense is not None else 0.0
    normalized_ebitda = (reported_ebitda - sbc_known) if reported_ebitda is not None else None
    normalized_fcf = (reported_fcf - sbc_known) if reported_fcf is not None else None

    raw_beta = safe_float(info.get("beta"), None)
    if raw_beta is not None and 0.20 <= raw_beta <= 3.0:
        beta = (2.0 / 3.0) * raw_beta + (1.0 / 3.0) * 1.0
        beta_status = "REPORTED_BLENDED"
    else:
        beta = 1.0
        beta_status = "NEUTRAL_1.0"
        data_warnings.append("Beta nebyla v tržním feedu dostupná; použita neutrální tržní hodnota 1,00.")

    cost_of_equity = rf_rate + beta * erp_rate

    int_exp_val = pts["interest_expense"].value
    if total_debt is not None and total_debt > 0 and int_exp_val is not None and int_exp_val > 0:
        pre_tax_cost_of_debt = max(rf_rate + 0.004, min(0.12, int_exp_val / total_debt))
        kd_status = "REPORTED"
    else:
        pre_tax_cost_of_debt = rf_rate + 0.015
        kd_status = "ESTIMATED_SPREAD"

    debt_for_wacc = total_debt if total_debt is not None else 0.0
    total_capital = market_cap + debt_for_wacc
    weight_equity = (market_cap / total_capital) if total_capital > 0 else 1.0
    weight_debt = (debt_for_wacc / total_capital) if total_capital > 0 else 0.0
    wacc = (weight_equity * cost_of_equity) + (weight_debt * pre_tax_cost_of_debt * (1.0 - effective_tax_rate))
    wacc = max(0.062, min(0.155, wacc))

    if book_equity is not None and total_debt is not None and total_cash is not None and reported_ebit is not None:
        invested_capital = max(1.0, book_equity + total_debt - total_cash - st_investments)
        nopat = reported_ebit * (1.0 - effective_tax_rate)
        roic = nopat / invested_capital
        economic_profit = (roic - wacc) * invested_capital
    else:
        invested_capital = None
        nopat = None
        roic = None
        economic_profit = None

    terminal_growth = min(0.03, max(0.02, wacc - 0.055))
    if terminal_growth >= wacc:
        terminal_growth = wacc * 0.45

    def simulate_fcff_dcf(initial_g: float, custom_wacc: float, custom_exit_mult: float, custom_tg: float) -> dict | None:
        if normalized_fcf is None or normalized_fcf <= 0 or net_debt is None:
            return None
        cf_rows = []
        curr_fcf = normalized_fcf
        curr_ebitda = normalized_ebitda if (normalized_ebitda is not None and normalized_ebitda > 0) else None
        pv_sum = 0.0
        for yr in range(1, 6):
            fade_weight = (yr - 1) / 4.0
            yr_growth = initial_g * (1.0 - 0.35 * fade_weight) + custom_tg * (0.35 * fade_weight)
            curr_fcf *= (1.0 + yr_growth)
            if curr_ebitda is not None:
                curr_ebitda *= (1.0 + yr_growth)
            disc_factor = 1.0 / ((1.0 + custom_wacc) ** yr)
            pv_cf = curr_fcf * disc_factor
            pv_sum += pv_cf
            cf_rows.append({
                "rok": yr,
                "rust_procent": round(yr_growth * 100.0, 2),
                "volny_penezni_tok": round(curr_fcf, 2),
                "diskontni_faktor": round(disc_factor, 6),
                "soucasna_hodnota_fcf": round(pv_cf, 2),
            })

        gordon_tv = (curr_fcf * (1.0 + custom_tg)) / max(0.015, custom_wacc - custom_tg)
        if curr_ebitda is not None:
            multiple_tv = curr_ebitda * custom_exit_mult
            blended_tv = 0.5 * gordon_tv + 0.5 * multiple_tv
        else:
            blended_tv = gordon_tv

        pv_tv = blended_tv / ((1.0 + custom_wacc) ** 5)
        ev_val = pv_sum + pv_tv
        eq_val = max(0.0, ev_val - net_debt - minority_interest)
        per_share = eq_val / shares_outstanding
        tv_share_pct = (pv_tv / ev_val * 100.0) if ev_val > 0 else 0.0
        return {
            "enterprise_value": ev_val,
            "equity_value": eq_val,
            "intrinsic_value_per_share": per_share,
            "terminal_value": blended_tv,
            "tv_share_percent": tv_share_pct,
            "cash_flows": cf_rows,
        }

    dcf_res = simulate_fcff_dcf(eff_growth, wacc, eff_exit_mult, terminal_growth)
    if dcf_res is not None:
        dcf_intrinsic_share = dcf_res["intrinsic_value_per_share"]
        dcf_upside = ((dcf_intrinsic_share / current_price) - 1.0) * 100.0
    else:
        dcf_intrinsic_share = None
        dcf_upside = None

    bench = ALL_SECTOR_BENCHMARKS.get(sector, {"ev_ebitda": 13.0, "pe": 19.0, "ev_revenue": 3.0, "pb": 3.0, "growth": 0.075})
    if owner_earnings_dp.value is not None and owner_earnings_dp.value > 0:
        base_fcfe = owner_earnings_dp.value
    elif is_financial and net_income is not None and net_income > 0:
        base_fcfe = net_income
    else:
        base_fcfe = None

    if base_fcfe is not None:
        curr_fcfe = base_fcfe
        fcfe_pv = 0.0
        for yr in range(1, 6):
            fade_w = (yr - 1) / 4.0
            yr_g = eff_growth * (1.0 - 0.35 * fade_w) + terminal_growth * (0.35 * fade_w)
            curr_fcfe *= (1.0 + yr_g)
            fcfe_pv += curr_fcfe / ((1.0 + cost_of_equity) ** yr)

        fcfe_gordon_tv = (curr_fcfe * (1.0 + terminal_growth)) / max(0.02, cost_of_equity - terminal_growth)
        fcfe_mult_tv = curr_fcfe * min(24.0, bench["pe"] * 0.85)
        fcfe_tv = fcfe_gordon_tv if is_financial else (0.5 * fcfe_gordon_tv + 0.5 * fcfe_mult_tv)
        fcfe_equity_val = fcfe_pv + (fcfe_tv / ((1.0 + cost_of_equity) ** 5))
        fcfe_per_share = max(0.0, fcfe_equity_val / shares_outstanding)
    else:
        fcfe_equity_val = None
        fcfe_per_share = None

    if dcf_res is not None:
        low_g, high_g = -0.35, 1.50
        for _ in range(60):
            mid_g = (low_g + high_g) / 2.0
            sim_mid = simulate_fcff_dcf(mid_g, wacc, eff_exit_mult, terminal_growth)
            if sim_mid is not None and sim_mid["intrinsic_value_per_share"] < current_price:
                low_g = mid_g
            else:
                high_g = mid_g
        implied_growth_pct = ((low_g + high_g) / 2.0) * 100.0
    else:
        implied_growth_pct = None

    if book_equity is not None and book_equity > 0 and net_income is not None:
        roe_est = min(0.55, max(-0.20, net_income / book_equity))
        pv_ri = 0.0
        curr_bv = book_equity
        faded_roe = roe_est
        for yr in range(1, 6):
            faded_roe = roe_est * (1.0 - 0.12 * (yr - 1)) + cost_of_equity * (0.12 * (yr - 1))
            ri_yr = (faded_roe - cost_of_equity) * curr_bv
            pv_ri += ri_yr / ((1.0 + cost_of_equity) ** yr)
            curr_bv += max(0.0, faded_roe * curr_bv * 0.45)

        terminal_ri = ((faded_roe - cost_of_equity) * curr_bv) / max(0.02, cost_of_equity - terminal_growth)
        ri_equity_val = max(0.0, book_equity + pv_ri + (terminal_ri / ((1.0 + cost_of_equity) ** 5)))
        ri_per_share = ri_equity_val / shares_outstanding
    else:
        roe_est = None
        ri_equity_val = None
        ri_per_share = None

    fwd_factor = (1.0 + eff_growth) ** 2
    pv_discount_2y_eq = 1.0 / ((1.0 + cost_of_equity) ** 2)
    pv_discount_2y_ev = 1.0 / ((1.0 + wacc) ** 2)

    quality_signal = (roe_est - cost_of_equity) if (is_financial and roe_est is not None) else ((roic - wacc) if roic is not None else 0.0)
    quality_mult_adj = max(0.85, min(1.12, 1.0 + quality_signal * 0.35))

    peer_rows = []
    if not is_financial and normalized_ebitda is not None and normalized_ebitda > 0 and net_debt is not None:
        mature_ev_ebitda = 0.62 * bench["ev_ebitda"] + 0.38 * 14.5
        adj_ev_ebitda = round(mature_ev_ebitda * quality_mult_adj, 1)
        fwd_ebitda_pv = normalized_ebitda * fwd_factor * pv_discount_2y_ev
        ev_ebitda_price = max(0.0, (fwd_ebitda_pv * adj_ev_ebitda - net_debt - minority_interest) / shares_outstanding)
        peer_rows.append({
            "method": "Hodnota podniku / Provozní zisk (Odiskont. Forward 2Y EV/EBITDA)",
            "multiple": adj_ev_ebitda,
            "forward_metric": round(fwd_ebitda_pv, 2),
            "implied_share_price": round(ev_ebitda_price, 2),
        })

    if net_income is not None and net_income > 0:
        mature_pe = (0.65 * bench["pe"] + 0.35 * 11.5) if is_financial else (0.62 * bench["pe"] + 0.38 * 19.5)
        adj_pe = round(mature_pe * quality_mult_adj, 1)
        fwd_ni_pv = net_income * fwd_factor * pv_discount_2y_eq
        pe_price = max(0.0, (fwd_ni_pv * adj_pe) / shares_outstanding)
        peer_rows.append({
            "method": "Cena / Čistý zisk (Odiskont. Forward 2Y P/E)",
            "multiple": adj_pe,
            "forward_metric": round(fwd_ni_pv, 2),
            "implied_share_price": round(pe_price, 2),
        })

    if revenue is not None and revenue > 0:
        if is_financial:
            fwd_rev_pv = revenue * fwd_factor * pv_discount_2y_eq
            ps_price = max(0.0, (fwd_rev_pv * bench["ev_revenue"]) / shares_outstanding)
            peer_rows.append({
                "method": "Tržní kapitalizace / Výnosy (Odiskont. Forward 2Y P/S)",
                "multiple": bench["ev_revenue"],
                "forward_metric": round(fwd_rev_pv, 2),
                "implied_share_price": round(ps_price, 2),
            })
        elif net_debt is not None:
            mature_ev_rev = 0.62 * bench["ev_revenue"] + 0.38 * 4.5
            adj_ev_rev = round(mature_ev_rev * quality_mult_adj, 1)
            fwd_rev_pv = revenue * fwd_factor * pv_discount_2y_ev
            ev_sales_price = max(0.0, (fwd_rev_pv * adj_ev_rev - net_debt - minority_interest) / shares_outstanding)
            peer_rows.append({
                "method": "Hodnota podniku / Tržby (Odiskont. Forward 2Y EV/Sales)",
                "multiple": adj_ev_rev,
                "forward_metric": round(fwd_rev_pv, 2),
                "implied_share_price": round(ev_sales_price, 2),
            })

    if book_equity is not None and book_equity > 0:
        if is_financial and roe_est is not None:
            pb_mult = round(max(0.65, min(3.2, bench["pb"] * (max(0.04, roe_est) / 0.12))), 2)
        else:
            pb_mult = bench["pb"]
        pb_price = max(0.0, (book_equity * pb_mult) / shares_outstanding)
        peer_rows.append({
            "method": "Cena / Účetní hodnota vlastního kapitálu (P/B)",
            "multiple": pb_mult,
            "forward_metric": round(book_equity, 2),
            "implied_share_price": round(pb_price, 2),
        })

    comps_prices = [r["implied_share_price"] for r in peer_rows if r["implied_share_price"] > 0]
    blended_comps_price = (sum(comps_prices) / len(comps_prices)) if comps_prices else None

    non_operating_investments = st_investments + lt_investments
    economic_float_benefit = (insurance_float * 0.35) + (deferred_tax_liab * 0.25)

    if is_financial and book_equity is not None and book_equity > 0:
        operating_core_val = book_equity * 1.05
        sotp_ev = operating_core_val + non_operating_investments + economic_float_benefit
        sotp_equity = max(0.0, sotp_ev - minority_interest)
        sotp_per_share = sotp_equity / shares_outstanding
    elif normalized_ebitda is not None and normalized_ebitda > 0 and total_debt is not None and total_cash is not None:
        operating_core_val = normalized_ebitda * eff_exit_mult
        sotp_ev = operating_core_val + non_operating_investments + economic_float_benefit
        sotp_equity = max(0.0, sotp_ev + total_cash - total_debt - minority_interest)
        sotp_per_share = sotp_equity / shares_outstanding
    else:
        operating_core_val = None
        sotp_ev = None
        sotp_equity = None
        sotp_per_share = None

    if book_equity is not None and total_debt is not None and total_cash is not None:
        nav_equity = book_equity + non_operating_investments + (0.0 if is_financial else (total_cash - total_debt) * 0.25)
        nav_per_share = max(0.0, nav_equity / shares_outstanding)
    else:
        nav_per_share = None

    if normalized_ebitda is not None and normalized_ebitda > 0 and normalized_fcf is not None and normalized_fcf > 0:
        lbo_base_ev = max(market_cap, enterprise_value or market_cap)
        lbo_entry_debt = lbo_base_ev * eff_leverage
        lbo_entry_equity = max(1000.0, lbo_base_ev - lbo_entry_debt)

        def simulate_lbo(g_rate: float) -> list[dict]:
            rem_debt = lbo_entry_debt
            curr_eb = normalized_ebitda
            curr_fcf_lbo = normalized_fcf
            rows = []
            for yr in range(1, 6):
                curr_eb *= (1.0 + g_rate)
                curr_fcf_lbo *= (1.0 + g_rate)
                int_cost = rem_debt * pre_tax_cost_of_debt
                cash_avail = max(0.0, curr_fcf_lbo - int_cost * (1.0 - effective_tax_rate))
                principal_paid = min(rem_debt, cash_avail * 0.75)
                rem_debt -= principal_paid
                rows.append({
                    "rok": yr,
                    "projektovany_provozni_zisk_pred_odpisy": round(curr_eb, 2),
                    "volny_tok_pro_splaceni": round(cash_avail, 2),
                    "mor_jistiny": round(principal_paid, 2),
                    "zbyvajici_cizi_kapital": round(rem_debt, 2),
                    "fcf_lbo": round(curr_fcf_lbo, 2),
                })
            return rows

        lbo_years = simulate_lbo(eff_growth)
        lbo_exit_ev = lbo_years[-1]["projektovany_provozni_zisk_pred_odpisy"] * eff_exit_mult
        lbo_exit_equity = max(0.0, lbo_exit_ev - lbo_years[-1]["zbyvajici_cizi_kapital"])
        lbo_moic = (lbo_exit_equity / lbo_entry_equity) if lbo_entry_equity > 0 else 0.0
        lbo_irr = ((lbo_moic ** (1.0 / 5.0)) - 1.0) * 100.0 if lbo_moic > 0 else -100.0

        growth_grid = [max(-0.50, eff_growth + d) for d in [-0.04, -0.02, 0.0, 0.02, 0.04]]
        mult_grid = [max(2.0, eff_exit_mult + d) for d in [-5.0, -2.5, 0.0, 2.5, 5.0]]
        lbo_sensitivity = []
        for g_val in growth_grid:
            sim_y = simulate_lbo(g_val)
            end_eb = sim_y[-1]["projektovany_provozni_zisk_pred_odpisy"]
            end_debt = sim_y[-1]["zbyvajici_cizi_kapital"]
            row_cells = []
            for m_val in mult_grid:
                eq_out = max(0.0, end_eb * m_val - end_debt)
                m_out = eq_out / lbo_entry_equity if lbo_entry_equity > 0 else 0.0
                irr_out = ((m_out ** 0.2) - 1.0) * 100.0 if m_out > 0 else -100.0
                row_cells.append({"vystupni_nasobek": round(m_val, 1), "irr_procent": round(irr_out, 2)})
            lbo_sensitivity.append({"tempo_rustu_ebitda_procent": round(g_val * 100.0, 1), "vysledky_irr_dle_nasobku": row_cells})
    else:
        lbo_entry_equity = None
        lbo_entry_debt = None
        lbo_exit_equity = None
        lbo_moic = None
        lbo_irr = None
        lbo_years = []
        lbo_sensitivity = []

    if dcf_res is not None:
        wacc_grid = [max(0.05, wacc + d) for d in [-0.015, -0.0075, 0.0, 0.0075, 0.015]]
        tg_grid = [0.015, 0.020, 0.025, 0.030, 0.035]
        dcf_sensitivity = []
        for w_val in wacc_grid:
            cells = []
            for tg_val in tg_grid:
                safe_tg = min(tg_val, w_val - 0.015)
                sim_c = simulate_fcff_dcf(eff_growth, w_val, eff_exit_mult, safe_tg)
                p_val = sim_c["intrinsic_value_per_share"] if sim_c else 0.0
                cells.append({"tg": round(tg_val * 100.0, 2), "price": round(p_val, 2)})
            dcf_sensitivity.append({"wacc": round(w_val * 100.0, 2), "cells": cells})

        bear_val = simulate_fcff_dcf(eff_growth * 0.55, wacc + 0.015, eff_exit_mult * 0.78, max(0.015, terminal_growth - 0.005))["intrinsic_value_per_share"]
        base_val = dcf_intrinsic_share
        bull_val = simulate_fcff_dcf(eff_growth * 1.30, max(0.065, wacc - 0.010), eff_exit_mult * 1.18, min(0.035, terminal_growth + 0.005))["intrinsic_value_per_share"]
        scenario_expected = 0.25 * bear_val + 0.50 * base_val + 0.25 * bull_val
    else:
        dcf_sensitivity = []
        bear_val = None
        base_val = None
        bull_val = None
        scenario_expected = None

    mc_base_ready = (is_financial and base_fcfe is not None) or (not is_financial and dcf_res is not None)
    if mc_base_ready:
        seed_val = int(hashlib.sha256(clean_ticker.encode("utf-8")).hexdigest()[:8], 16)
        rng = random.Random(seed_val)
        mc_results = []
        for _ in range(2500):
            sim_g = rng.gauss(eff_growth, max(0.03, abs(eff_growth) * 0.22))
            sim_wacc = max(0.06, min(0.18, rng.gauss(wacc, 0.011)))
            sim_mult = max(5.0, rng.gauss(eff_exit_mult, eff_exit_mult * 0.14))
            sim_margin_shock = rng.gauss(1.0, 0.08)
            if is_financial:
                c_fe = base_fcfe * sim_margin_shock
                pv_fe = 0.0
                for yr in range(1, 6):
                    yg = sim_g * (1.0 - 0.35 * ((yr - 1) / 4.0))
                    c_fe *= (1.0 + yg)
                    pv_fe += c_fe / ((1.0 + sim_wacc) ** yr)
                tv_fe = (c_fe * (1.0 + terminal_growth)) / max(0.015, sim_wacc - terminal_growth)
                eq_m = max(0.0, pv_fe + (tv_fe / ((1.0 + sim_wacc) ** 5)))
            else:
                c_fcf = normalized_fcf * sim_margin_shock
                c_eb = (normalized_ebitda * sim_margin_shock) if normalized_ebitda is not None else None
                pv_cf = 0.0
                for yr in range(1, 6):
                    fade = (yr - 1) / 4.0
                    yg = sim_g * (1.0 - 0.35 * fade) + terminal_growth * (0.35 * fade)
                    c_fcf *= (1.0 + yg)
                    if c_eb is not None:
                        c_eb *= (1.0 + yg)
                    pv_cf += c_fcf / ((1.0 + sim_wacc) ** yr)
                tv_g = (c_fcf * (1.0 + terminal_growth)) / max(0.015, sim_wacc - terminal_growth)
                if c_eb is not None:
                    tv_m = c_eb * sim_mult
                    blended_tv_m = 0.5 * tv_g + 0.5 * tv_m
                else:
                    blended_tv_m = tv_g
                ev_m = pv_cf + (blended_tv_m / ((1.0 + sim_wacc) ** 5))
                eq_m = max(0.0, ev_m - net_debt - minority_interest)
            mc_results.append(eq_m / shares_outstanding)

        mc_results.sort()
        n_mc = len(mc_results)
        mc_p5 = mc_results[int(0.05 * n_mc)]
        mc_p10 = mc_results[int(0.10 * n_mc)]
        mc_p25 = mc_results[int(0.25 * n_mc)]
        mc_p50 = mc_results[int(0.50 * n_mc)]
        mc_p75 = mc_results[int(0.75 * n_mc)]
        mc_p90 = mc_results[int(0.90 * n_mc)]
        mc_p95 = mc_results[int(0.95 * n_mc)]
        prob_above_market = sum(1 for x in mc_results if x > current_price) / n_mc * 100.0

        hist_low = mc_results[int(0.02 * n_mc)]
        hist_high = max(hist_low + 1.0, mc_results[int(0.98 * n_mc)])
        num_bins = 14
        bin_width = (hist_high - hist_low) / num_bins
        mc_histogram = []
        for i in range(num_bins):
            b_start = hist_low + i * bin_width
            b_end = b_start + bin_width
            b_mid = (b_start + b_end) / 2.0
            count = sum(1 for v in mc_results if (b_start <= v < b_end) or (i == num_bins - 1 and v >= b_start and v <= b_end))
            mc_histogram.append({"label": f"{b_mid:.1f}", "midpoint": round(b_mid, 2), "count": count})
    else:
        mc_p5 = mc_p10 = mc_p25 = mc_p50 = mc_p75 = mc_p90 = mc_p95 = None
        prob_above_market = None
        mc_histogram = []

    if is_financial:
        base_weights = {
            "Vícefázový DCF pro podnik": 0.00,
            "Owner Earnings / FCFE DCF": 0.30,
            "Srovnatelné tržní násobky": 0.30,
            "Model reziduálního zisku": 0.30,
            "Strukturovaný součet částí (SOTP)": 0.00,
            "Medián Monte Carlo (P50)": 0.10,
        }
    else:
        base_weights = {
            "Vícefázový DCF pro podnik": 0.35,
            "Owner Earnings / FCFE DCF": 0.15,
            "Srovnatelné tržní násobky": 0.20,
            "Model reziduálního zisku": 0.05,
            "Strukturovaný součet částí (SOTP)": 0.10,
            "Medián Monte Carlo (P50)": 0.15,
        }

    method_values: dict[str, float | None] = {
        "Vícefázový DCF pro podnik": dcf_intrinsic_share,
        "Owner Earnings / FCFE DCF": fcfe_per_share,
        "Srovnatelné tržní násobky": blended_comps_price,
        "Model reziduálního zisku": ri_per_share,
        "Strukturovaný součet částí (SOTP)": sotp_per_share,
        "Medián Monte Carlo (P50)": mc_p50,
    }

    available_weight_sum = sum(
        w for m, w in base_weights.items() if method_values[m] is not None and method_values[m] > 0
    )

    if available_weight_sum <= 0:
        raise ValueError(
            f"Pro titul '{clean_ticker}' chybí klíčové položky účetních výkazů ({', '.join(missing_fields)}), "
            "proto nelze spočítat spolehlivou vnitřní hodnotu bez umělého domýšlení dat."
        )

    effective_weights: dict[str, float] = {}
    for m_name, raw_w in base_weights.items():
        val = method_values[m_name]
        if val is not None and val > 0 and raw_w > 0:
            effective_weights[m_name] = raw_w / available_weight_sum
        else:
            effective_weights[m_name] = 0.0

    weighted_fair_value = sum(
        effective_weights[m] * (method_values[m] or 0.0) for m in effective_weights
    )

    synthesis_rows = []
    for m_name in base_weights:
        w_val = effective_weights[m_name]
        v_price = method_values[m_name]
        if v_price is not None and v_price > 0:
            diff_pct = ((v_price / current_price) - 1.0) * 100.0
            synthesis_rows.append({
                "method": m_name,
                "status": "AVAILABLE",
                "weight_percent": round(w_val * 100.0, 1),
                "fair_value": round(v_price, 2),
                "contribution": round(w_val * v_price, 2),
                "vs_market_percent": round(diff_pct, 2),
            })
        else:
            synthesis_rows.append({
                "method": m_name,
                "status": "N/A (chybí vstupní data výkazu)",
                "weight_percent": 0.0,
                "fair_value": None,
                "contribution": 0.0,
                "vs_market_percent": None,
            })

    low_beta_ok = beta <= 1.05
    quality_ok = (roic is not None and roic > wacc) or (is_financial and roe_est is not None and roe_est > cost_of_equity)
    oe_conversion_ok = (
        owner_earnings_dp.value is not None
        and net_income is not None
        and net_income > 0
        and owner_earnings_dp.value >= net_income * 0.75
    )
    cheap_float_present = (insurance_float > 0) or (deferred_tax_liab > (market_cap * 0.03))
    high_52w = safe_float(info.get("fiftyTwoWeekHigh"), None)
    current_drawdown_pct = ((high_52w - current_price) / high_52w * 100.0) if (high_52w and high_52w > current_price) else 0.0

    factor_profile = {
        "source_provider": stmt_provider,
        "low_beta_bab": {
            "label": "Nízká beta / stabilita (BAB)",
            "passed": low_beta_ok,
            "detail": f"Upravená Beta {beta:.2f} ({'nízká tržní citlivost' if low_beta_ok else 'vyšší cyklická citlivost'}).",
        },
        "quality_qmj": {
            "label": "Kvalita kapitálu & Moat (QMJ)",
            "passed": quality_ok,
            "detail": (
                f"ROIC {roic * 100:.2f} % vs. WACC {wacc * 100:.2f} %."
                if roic is not None
                else (f"ROE {roe_est * 100:.2f} % vs. CoE {cost_of_equity * 100:.2f} %." if roe_est is not None else "Nedostatek dat pro ROIC.")
            ),
        },
        "owner_earnings_conversion": {
            "label": "Konverze zisku na Owner Earnings",
            "passed": oe_conversion_ok,
            "detail": (
                f"Owner Earnings pokrývají {(owner_earnings_dp.value / net_income) * 100:.1f} % účetního čistého zisku."
                if (owner_earnings_dp.value is not None and net_income and net_income > 0)
                else "Nelze ověřit bez kompletního CapEx a čistého zisku."
            ),
        },
        "cheap_float_financing": {
            "label": "Nízkonákladové zdroje (Float / Odložené daně)",
            "passed": cheap_float_present,
            "detail": f"Pojistný float: {insurance_float:,.0f} {currency} | Odložené daňové závazky: {deferred_tax_liab:,.0f} {currency}.",
        },
        "drawdown_and_size": {
            "label": "Aktuální drawdown & velikostní kotva",
            "passed": current_drawdown_pct < 25.0,
            "detail": f"Propad od 52týdenního maxima: -{current_drawdown_pct:.1f} % | Tržní kapitalizace: {market_cap:,.0f} {currency}.",
        },
        "model_limits": [
            "Udržovací CapEx je odhadován pravidlem min(CapEx, 1,05 × odpisy), protože účetní standardy nerozlišují expanzní investice v rozvaze.",
            "Kvantitativní metriky (ROIC, marže, Owner Earnings) potvrzují historickou kvalitu, ale nenahrazují posouzení konkurenční výhody (moat) a regulace.",
            "U velkých společností (nad 250 mld. USD) působí efekt velikosti ('kotva'), který přirozeně brzdí budoucí procentuální růst.",
        ],
    }

    integrity_score = 45
    if "SEC EDGAR" in stmt_provider:
        integrity_score += 25
    elif "Financial Modeling Prep" in stmt_provider:
        integrity_score += 20
    else:
        integrity_score += 10

    if len(series.get("revenue") or []) >= 3:
        integrity_score += 10
    if not missing_fields:
        integrity_score += 10
    else:
        integrity_score -= min(20, len(missing_fields) * 4)
    if quality_ok:
        integrity_score += 5
    integrity_score = max(25, min(98, integrity_score))

    primary_intrinsic = fcfe_per_share if is_financial else (dcf_intrinsic_share or fcfe_per_share)
    if primary_intrinsic is not None and blended_comps_price is not None:
        spread_ratio = abs(primary_intrinsic - blended_comps_price) / max(1.0, weighted_fair_value)
        if spread_ratio > 0.25:
            disagreement_note = (
                f"Rozptyl mezi výnosovým modelem ({primary_intrinsic:,.2f} {currency}) a tržními násobky "
                f"({blended_comps_price:,.2f} {currency}) dosahuje {spread_ratio * 100:.1f} %."
            )
        else:
            disagreement_note = (
                f"Vysoká konvergence oceňovacích přístupů: výnosový model ({primary_intrinsic:,.2f} {currency}) "
                f"a relativní násobky ({blended_comps_price:,.2f} {currency}) potvrzují konzistentní odhad vnitřní hodnoty."
            )
    else:
        disagreement_note = (
            "Vážená vnitřní hodnota byla sestavena výhradně z modelů, pro které existují ověřená data výkazů."
        )

    if data_warnings:
        disagreement_note += " Kontrola kvality dat: " + " ".join(data_warnings)

    green_upper = weighted_fair_value * 0.80
    yellow_upper = weighted_fair_value * 0.90
    orange_upper = weighted_fair_value * 1.00

    mos_percent = ((weighted_fair_value - current_price) / weighted_fair_value) * 100.0 if weighted_fair_value > 0 else 0.0
    upside_to_fair = ((weighted_fair_value / current_price) - 1.0) * 100.0 if current_price > 0 else 0.0

    if current_price < green_upper:
        zone = "GREEN"
        tranche = 55
        recommendation = "ZÓNA GREEN — Výrazná bezpečnostní rezerva (nad 20 %)"
    elif current_price < yellow_upper:
        zone = "YELLOW"
        tranche = 30
        recommendation = "ZÓNA YELLOW — Atraktivní diskont vůči vnitřní hodnotě (10–20 %)"
    elif current_price <= orange_upper:
        zone = "ORANGE"
        tranche = 15
        recommendation = "ZÓNA ORANGE — Mírné podhodnocení, vhodná úvodní tranše (0–10 %)"
    else:
        zone = "RED"
        tranche = 0
        recommendation = "ZÓNA RED — Tržní cena převyšuje váženou vnitřní hodnotu (vyčkat)"

    rev_hist = series.get("revenue") or []
    ebit_hist = series.get("ebit") or []
    cur_assets_hist = series.get("current_assets") or []
    cur_liab_hist = series.get("current_liabilities") or []
    total_assets_hist = series.get("total_assets") or []
    retained_earn_hist = series.get("retained_earnings") or []

    f_score = 0
    if net_income is not None and net_income > 0:
        f_score += 1
    if raw_ocf is not None and raw_ocf > 0:
        f_score += 1
    if raw_ocf is not None and net_income is not None and raw_ocf > net_income:
        f_score += 1
    if quality_ok:
        f_score += 1
    if is_financial or (net_debt is not None and (net_debt <= 0 or (normalized_ebitda and total_debt / normalized_ebitda < 2.2))):
        f_score += 1
    cur_ratio = (cur_assets_hist[0] / cur_liab_hist[0]) if (cur_assets_hist and cur_liab_hist and cur_liab_hist[0] > 0) else None
    if is_financial or (cur_ratio is not None and cur_ratio > 1.2):
        f_score += 1
    if len(rev_hist) >= 2 and rev_hist[0] >= rev_hist[1]:
        f_score += 1
    if len(ebit_hist) >= 2 and len(rev_hist) >= 2 and rev_hist[0] > 0 and rev_hist[1] > 0:
        if (ebit_hist[0] / rev_hist[0]) >= (ebit_hist[1] / rev_hist[1]):
            f_score += 1
    if is_financial or (revenue is not None and total_assets_hist and total_assets_hist[0] > 0 and (revenue / total_assets_hist[0]) > 0.45):
        f_score += 1

    if is_financial:
        altman_z = None
        altman_status = f"Finanční instituce (ROE {roe_est * 100:.1f} %)" if roe_est is not None else "Finanční instituce"
    elif (
        total_assets_hist and total_assets_hist[0] > 0
        and cur_assets_hist and cur_liab_hist
        and retained_earn_hist
        and reported_ebit is not None
        and total_debt is not None and total_debt > 0
        and revenue is not None
    ):
        tot_assets = total_assets_hist[0]
        wc = cur_assets_hist[0] - cur_liab_hist[0]
        ret_earn = retained_earn_hist[0]
        altman_z = (
            1.2 * (wc / tot_assets)
            + 1.4 * (ret_earn / tot_assets)
            + 3.3 * (reported_ebit / tot_assets)
            + 0.6 * (market_cap / total_debt)
            + 1.0 * (revenue / tot_assets)
        )
        altman_status = "Bezpečné pásmo (nad 2,99)" if altman_z > 2.99 else ("Šedá zóna (1,81–2,99)" if altman_z >= 1.81 else "Zvýšené riziko (pod 1,81)")
    else:
        altman_z = None
        altman_status = "N/A (chybí položky rozvahy nebo nulový dluh)"

    interest_coverage = (reported_ebit / int_exp_val) if (reported_ebit is not None and int_exp_val is not None and int_exp_val > 0) else None

    return {
        "spolecnost": clean_ticker,
        "company_name": company_name,
        "company_logo_url": company_logo_url,
        "company_logo_urls": company_logo_urls,
        "ticker": clean_ticker,
        "currency": currency,
        "exchange": exchange,
        "sector": sector,
        "industry": industry,
        "country": country,
        "employees": employees,
        "business_summary": business_summary,
        "data_quality": {
            "primary_statement_source": stmt_provider,
            "missing_fields": missing_fields,
            "estimated_fields": estimated_fields,
            "warnings": data_warnings,
            "points": {k: v.to_dict() for k, v in pts.items()},
        },
        "factor_and_quality_diagnostics": factor_profile,
        "dynamic_defaults": {
            "leverage_percent": round(auto_lev * 100.0, 1),
            "growth_percent": round(auto_grw * 100.0, 1),
            "exit_multiple": round(auto_mult, 1),
        },
        "trzni_kapitalizace": round(market_cap, 2),
        "celkovy_dluh": round(total_debt, 2) if total_debt is not None else None,
        "hotovost": round(total_cash, 2) if total_cash is not None else None,
        "cisty_dluh": round(net_debt, 2) if net_debt is not None else None,
        "celkova_hodnota_podniku": round(enterprise_value, 2) if enterprise_value is not None else None,
        "fundamentals": {
            "revenue": round(revenue, 2) if revenue is not None else None,
            "reported_ebitda": round(reported_ebitda, 2) if reported_ebitda is not None else None,
            "normalized_ebitda": round(normalized_ebitda, 2) if normalized_ebitda is not None else None,
            "reported_fcf": round(reported_fcf, 2) if reported_fcf is not None else None,
            "normalized_fcf": round(normalized_fcf, 2) if normalized_fcf is not None else None,
            "owner_earnings": round(owner_earnings_dp.value, 2) if owner_earnings_dp.value is not None else None,
            "owner_earnings_per_share": round(owner_earnings_dp.value / shares_outstanding, 2) if owner_earnings_dp.value is not None else None,
            "look_through_earnings": round(look_through_dp.value, 2) if look_through_dp.value is not None else None,
            "maintenance_capex": round(maintenance_capex_dp.value, 2) if maintenance_capex_dp.value is not None else None,
            "growth_capex": round(growth_capex_dp.value, 2) if growth_capex_dp.value is not None else None,
            "sbc_expense": round(sbc_expense, 2) if sbc_expense is not None else None,
            "net_income": round(net_income, 2) if net_income is not None else None,
            "roic_percent": round((roe_est if is_financial else roic) * 100.0, 2) if ((roe_est if is_financial else roic) is not None) else None,
            "economic_profit": round(economic_profit, 2) if economic_profit is not None else None,
            "fcf_conversion_percent": round((normalized_fcf / normalized_ebitda) * 100.0, 1) if (normalized_fcf is not None and normalized_ebitda and normalized_ebitda > 0) else None,
            "net_debt_to_ebitda": round(net_debt / normalized_ebitda, 2) if (net_debt is not None and normalized_ebitda and normalized_ebitda > 0) else None,
            "interest_coverage": round(interest_coverage, 1) if interest_coverage is not None else None,
            "current_ratio": round(cur_ratio, 2) if cur_ratio is not None else None,
            "operating_margin_percent": round((reported_ebit / revenue) * 100.0, 2) if (reported_ebit is not None and revenue and revenue > 0) else None,
        },
        "wacc_engine": {
            "risk_free_rate_percent": round(rf_rate * 100.0, 2),
            "erp_percent": round(erp_rate * 100.0, 2),
            "beta": round(beta, 2),
            "beta_status": beta_status,
            "cost_of_equity_percent": round(cost_of_equity * 100.0, 2),
            "pre_tax_cost_of_debt_percent": round(pre_tax_cost_of_debt * 100.0, 2),
            "kd_status": kd_status,
            "after_tax_cost_of_debt_percent": round(pre_tax_cost_of_debt * (1.0 - effective_tax_rate) * 100.0, 2),
            "tax_rate_status": tax_rate_status,
            "weight_equity_percent": round(weight_equity * 100.0, 1),
            "weight_debt_percent": round(weight_debt * 100.0, 1),
            "wacc_percent": round(wacc * 100.0, 2),
        },
        "dcf_valuation": {
            "available": dcf_res is not None,
            "wacc": round(wacc, 6),
            "terminal_growth": round(terminal_growth, 4),
            "terminal_value": round(dcf_res["terminal_value"], 2) if dcf_res else None,
            "tv_share_percent": round(dcf_res["tv_share_percent"], 1) if dcf_res else None,
            "enterprise_value": round(dcf_res["enterprise_value"], 2) if dcf_res else None,
            "equity_value": round(dcf_res["equity_value"], 2) if dcf_res else None,
            "intrinsic_value_per_share": round(dcf_intrinsic_share, 2) if dcf_intrinsic_share is not None else None,
            "current_share_price": round(current_price, 2),
            "upside_downside_percent": round(dcf_upside, 2) if dcf_upside is not None else None,
            "cash_flows": dcf_res["cash_flows"] if dcf_res else [],
            "sensitivity_matrix": dcf_sensitivity,
        },
        "fcfe_valuation": {
            "available": fcfe_per_share is not None,
            "cost_of_equity": round(cost_of_equity, 4),
            "equity_value": round(fcfe_equity_val, 2) if fcfe_equity_val is not None else None,
            "intrinsic_value_per_share": round(fcfe_per_share, 2) if fcfe_per_share is not None else None,
        },
        "reverse_dcf": {
            "available": implied_growth_pct is not None,
            "implied_fcf_growth_percent": round(implied_growth_pct, 2) if implied_growth_pct is not None else None,
            "assumed_growth_percent": round(eff_growth * 100.0, 2),
            "expectation_gap_percent": round((eff_growth * 100.0) - implied_growth_pct, 2) if implied_growth_pct is not None else None,
        },
        "residual_income": {
            "available": ri_per_share is not None,
            "book_value_equity": round(book_equity, 2) if book_equity is not None else None,
            "roe_percent": round(roe_est * 100.0, 2) if roe_est is not None else None,
            "equity_value": round(ri_equity_val, 2) if ri_equity_val is not None else None,
            "intrinsic_value_per_share": round(ri_per_share, 2) if ri_per_share is not None else None,
        },
        "peer_comps": {
            "sector": sector,
            "blended_implied_price": round(blended_comps_price, 2) if blended_comps_price is not None else None,
            "multiples": peer_rows,
        },
        "metoda_souctu_casti": {
            "available": sotp_per_share is not None,
            "provozni_jadro_hodnota": round(operating_core_val, 2) if operating_core_val is not None else None,
            "neprovozni_investice": round(non_operating_investments, 2),
            "ekonomicka_vyhoda_float_a_dani": round(economic_float_benefit, 2),
            "vypoctena_hodnota_podniku": round(sotp_ev, 2) if sotp_ev is not None else None,
            "vypoctena_hodnota_vlastniho_kapitalu": round(sotp_equity, 2) if sotp_equity is not None else None,
            "sotp_per_share": round(sotp_per_share, 2) if sotp_per_share is not None else None,
            "nav_per_share": round(nav_per_share, 2) if nav_per_share is not None else None,
        },
        "model_odkupu_za_cizi_kapital_base_case": {
            "available": lbo_moic is not None,
            "vstupni_vlastni_kapital": round(lbo_entry_equity, 2) if lbo_entry_equity is not None else None,
            "vstupni_cizi_kapital": round(lbo_entry_debt, 2) if lbo_entry_debt is not None else None,
            "vystupni_hodnota_vlastniho_kapitalu": round(lbo_exit_equity, 2) if lbo_exit_equity is not None else None,
            "moic": round(lbo_moic, 2) if lbo_moic is not None else None,
            "irr_procent": round(lbo_irr, 2) if lbo_irr is not None else None,
        },
        "projektovane_roky": lbo_years,
        "citlivostni_matice_lbo_irr": lbo_sensitivity,
        "scenarios": {
            "available": scenario_expected is not None,
            "bear": {"probability": 25, "fair_value": round(bear_val, 2) if bear_val is not None else None},
            "base": {"probability": 50, "fair_value": round(base_val, 2) if base_val is not None else None},
            "bull": {"probability": 25, "fair_value": round(bull_val, 2) if bull_val is not None else None},
            "expected_value": round(scenario_expected, 2) if scenario_expected is not None else None,
        },
        "monte_carlo": {
            "available": mc_p50 is not None,
            "iterations": 2500 if mc_p50 is not None else 0,
            "p5": round(mc_p5, 2) if mc_p5 is not None else None,
            "p10": round(mc_p10, 2) if mc_p10 is not None else None,
            "p25": round(mc_p25, 2) if mc_p25 is not None else None,
            "p50": round(mc_p50, 2) if mc_p50 is not None else None,
            "p75": round(mc_p75, 2) if mc_p75 is not None else None,
            "p90": round(mc_p90, 2) if mc_p90 is not None else None,
            "p95": round(mc_p95, 2) if mc_p95 is not None else None,
            "prob_above_market_percent": round(prob_above_market, 1) if prob_above_market is not None else None,
            "histogram": mc_histogram,
        },
        "valuation_synthesis": {
            "weighted_fair_value": round(weighted_fair_value, 2),
            "methods": synthesis_rows,
            "integrity_score": integrity_score,
            "disagreement_note": disagreement_note,
        },
        "decision_engine": {
            "margin_of_safety_percent": round(mos_percent, 2),
            "upside_to_fair_value_percent": round(upside_to_fair, 2),
            "zone": zone,
            "tranche_allocation_percent": tranche,
            "recommendation": recommendation,
            "bands": {
                "green_max": round(green_upper, 2),
                "yellow_min": round(green_upper, 2),
                "yellow_max": round(yellow_upper, 2),
                "orange_min": round(yellow_upper, 2),
                "orange_max": round(orange_upper, 2),
                "red_min": round(orange_upper, 2),
            },
        },
        "risk_analysis": {
            "piotroski_f_score": f_score,
            "altman_z_score": round(altman_z, 2) if altman_z is not None else None,
            "altman_status": altman_status,
        },
        "audit_trail": {
            "model_version": "Tensegrity Institutional Engine v5.2 (SEC + Dynamic IPO + Resilient AI)",
            "timestamp": get_prague_now().strftime("%Y-%m-%d %H:%M:%S %Z"),
            "data_source": f"{stmt_provider} + Live FX & Treasury",
            "currency": currency,
            "financial_currency": fin_currency,
            "fx_applied": round(stmt_fx, 4),
            "assumptions": {
                "leverage_percent": round(eff_leverage * 100.0, 1),
                "growth_percent": round(eff_growth * 100.0, 1),
                "exit_multiple": round(eff_exit_mult, 1),
                "terminal_growth_percent": round(terminal_growth * 100.0, 2),
            },
        },
    }


# ==============================================================================
# 7. HISTORICKÁ DATA PRO INTERAKTIVNÍ GRAF (TRADINGVIEW / TRADING 212 STYL)
# ==============================================================================

CHART_TIMEFRAME_MAP = {
    "1D": ("1d", "5m"),
    "1T": ("5d", "30m"),
    "1M": ("1mo", "1d"),
    "6M": ("6mo", "1d"),
    "YTD": ("ytd", "1d"),
    "1R": ("1y", "1d"),
    "5R": ("5y", "1wk"),
    "MAX": ("max", "1mo"),
}


def fetch_interactive_chart_series(query_or_ticker: str, timeframe: str = "1R") -> dict:
    clean_ticker = resolve_company_ticker(query_or_ticker)
    tf = (timeframe or "1R").upper()
    period, interval = CHART_TIMEFRAME_MAP.get(tf, ("1y", "1d"))
    intraday = tf in ("1D", "1T")

    stock = yf.Ticker(clean_ticker)
    hist = stock.history(period=period, interval=interval)
    if hist is None or hist.empty:
        hist = stock.history(period="1y", interval="1d")
        intraday = False

    candles = []
    volumes = []
    seen_ts = set()
    for idx, row in hist.iterrows():
        c_val = safe_float(row.get("Close"), None)
        o_val = safe_float(row.get("Open"), c_val)
        h_val = safe_float(row.get("High"), c_val)
        l_val = safe_float(row.get("Low"), c_val)
        v_val = safe_float(row.get("Volume"), 0.0) or 0.0
        if c_val is None or c_val <= 0:
            continue

        if intraday:
            ts_val = int(idx.timestamp())
            if ts_val in seen_ts:
                continue
            seen_ts.add(ts_val)
            time_key = ts_val
        else:
            date_str = idx.strftime("%Y-%m-%d")
            if date_str in seen_ts:
                continue
            seen_ts.add(date_str)
            time_key = date_str

        candles.append({
            "time": time_key,
            "open": round(o_val, 4),
            "high": round(h_val, 4),
            "low": round(l_val, 4),
            "close": round(c_val, 4),
        })
        volumes.append({
            "time": time_key,
            "value": round(v_val, 0),
            "color": "rgba(16, 185, 129, 0.35)" if c_val >= o_val else "rgba(239, 68, 68, 0.35)",
        })

    first_close = candles[0]["close"] if candles else 0.0
    last_close = candles[-1]["close"] if candles else 0.0
    change_abs = last_close - first_close
    change_pct = ((last_close / first_close) - 1.0) * 100.0 if first_close > 0 else 0.0

    return {
        "ticker": clean_ticker,
        "timeframe": tf,
        "intraday": intraday,
        "candles": candles,
        "volumes": volumes,
        "first_close": round(first_close, 2),
        "last_close": round(last_close, 2),
        "change_abs": round(change_abs, 2),
        "change_pct": round(change_pct, 2),
    }


# ==============================================================================
# 8. AI ANALYTICKÝ PORADCE (ODOLNÝ PROTI VÝPADKŮM + ČISTÁ ČEŠTINA BEZ LATEXU)
# ==============================================================================

def _sanitize_advisor_output(text: str) -> str:
    """Převede surové LaTeX příkazy (\text, \frac, \times, \beta) na čistý čitelný text."""
    s = str(text or "")
    s = s.replace("\\[", "").replace("\\]", "").replace("\\(", "").replace("\\)", "")
    s = s.replace("$$", "")
    # \frac{a}{b} -> (a / b)
    for _ in range(3):
        s = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"(\1 / \2)", s)
    # \text{něco} -> něco
    s = re.sub(r"\\text\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\mathrm\{([^{}]*)\}", r"\1", s)
    s = s.replace("\\times", "×").replace("\\cdot", "·").replace("\\beta", "Beta").replace("\\approx", "≈")
    s = s.replace("_{net}", " (po zdanění)")
    return s.strip()


def _build_ai_advisor_system_prompt(d: dict | None) -> str:
    context_str = ""
    if d:
        compact_context = {
            "spolecnost": d["company_name"],
            "ticker": d["ticker"],
            "zdroj_vykazu": d["data_quality"]["primary_statement_source"],
            "chybejici_polozky": d["data_quality"]["missing_fields"],
            "sektor": d["sector"],
            "odvetvi": d["industry"],
            "mena": d["currency"],
            "trzni_cena": d["dcf_valuation"]["current_share_price"],
            "vazena_vnitrni_hodnota": d["valuation_synthesis"]["weighted_fair_value"],
            "zona": d["decision_engine"]["zone"],
            "bezpecnostni_rezerva_procent": d["decision_engine"]["margin_of_safety_percent"],
            "dcf_podnik": d["dcf_valuation"]["intrinsic_value_per_share"],
            "owner_earnings_fcfe_akcie": d["fcfe_valuation"]["intrinsic_value_per_share"],
            "owner_earnings_celkem": d["fundamentals"]["owner_earnings"],
            "look_through_earnings": d["fundamentals"]["look_through_earnings"],
            "maintenance_capex": d["fundamentals"]["maintenance_capex"],
            "growth_capex": d["fundamentals"]["growth_capex"],
            "sotp_akcie": d["metoda_souctu_casti"]["sotp_per_share"],
            "nasobky_cena": d["peer_comps"]["blended_implied_price"],
            "monte_carlo_p50": d["monte_carlo"]["p50"],
            "obraceny_dcf_rust_procent": d["reverse_dcf"]["implied_fcf_growth_percent"],
            "wacc_procent": d["wacc_engine"]["wacc_percent"],
            "roic_procent": d["fundamentals"]["roic_percent"],
            "trzni_kapitalizace": d["trzni_kapitalizace"],
            "dluh": d["celkovy_dluh"],
            "hotovost": d["hotovost"],
            "piotroski": d["risk_analysis"]["piotroski_f_score"],
        }
        context_str = f"\nKontext aktuálně otevřené firmy v terminálu: {json.dumps(compact_context, ensure_ascii=False)}"

    return (
        "Jsi špičkový finanční, investiční, právní a všeobecný AI poradce platformy Tensegrity Terminal. "
        "Odpovídej v perfektní češtině na JAKÝKOLIV dotaz uživatele.\n"
        "KLÍČOVÁ PRAVIDLA:\n"
        "1. Přizpůsob styl a jednoduchost vysvětlení přesně přání uživatele (pokud chce vysvětlení 'jako dítěti', použij přirovnání ke stánku s limonádou, hračkám nebo kapesnému).\n"
        "2. NIKDY nepoužívej LaTeX značky jako \\text{}, \\frac{}, \\[, \\] nebo \\( \\) — piš vzorce obyčejným českým textem (např. 'WACC = podíl vlastního kapitálu × jeho cena + podíl dluhu × úrok po zdanění').\n"
        "3. Pokud v kontextu firmy chybí některé položky ('chybejici_polozky'), otevřeně na to upozorni.\n"
        f"{context_str}"
    )


def _offline_smart_czech_knowledge_fallback(question: str, d: dict | None) -> str:
    """Chytrá lokální znalostní báze pro případ, že externí bezplatné API dočasně omezí požadavky."""
    q = question.lower().strip()
    as_kid = any(k in q for k in ("ditet", "dítět", "deck", "děck", "jednodus", "jednoduš", "polopat", "laick"))

    if "wacc" in q:
        wacc_note = ""
        if d and d.get("wacc_engine"):
            w_val = d["wacc_engine"]["wacc_percent"]
            wacc_note = f"\n\nU aktuálně otevřené firmy **{d['company_name']} ({d['ticker']})** vychází WACC na **{w_val} %**."
        if as_kid:
            return (
                "Představ si, že si chceš otevřít **stánek s limonádou** a potřebuješ na něj 100 korun:\n"
                "1. **50 korun** ti půjčí kamarád a chce za to vrátit o **5 %** víc (to je cena dluhu).\n"
                "2. **50 korun** ti dá brácha jako spolumajitel a chce z tvého zisku podíl odpovídající **15 %** (to je cena vlastního kapitálu, protože nese větší riziko).\n\n"
                "**WACC** je prostě průměrná cena všech peněz, které sis na svůj stánek sehnal — tady přesně **10 %**. "
                "Aby se ti stánek vyplatil, musí vydělat víc než těch 10 %."
                f"{wacc_note}"
            )
        return (
            "**WACC (vážené průměrné náklady kapitálu)** udává, kolik procent ročně v průměru stojí firmu peníze, kterými financuje své podnikání:\n"
            "- **Vlastní kapitál (akcie):** výnos, který očekávají akcionáři za podstoupené riziko.\n"
            "- **Cizí kapitál (dluh):** úrok z úvěrů a dluhopisů snížený o daňovou úsporu (úroky snižují základ daně).\n\n"
            "**Vzorec:** `WACC = (podíl vlastního kapitálu × cena vlastního kapitálu) + (podíl dluhu × úrok × (1 - daňová sazba))`.\n"
            "Slouží jako minimální laťka výnosnosti (diskontní sazba) při výpočtu vnitřní hodnoty firmy."
            f"{wacc_note}"
        )

    if "dcf" in q:
        if as_kid:
            return (
                "**DCF** je způsob, jak zjistit, kolik dnes stojí kouzelná pokladnička, která ti každý rok vyrobí peníze. "
                "Protože 100 korun dnes má větší hodnotu než 100 korun za pět let, všechny budoucí peníze z pokladničky se postupně zmenší (odiskontují) a sečtou — a to je dnešní férová cena firmy."
            )
        return (
            "**DCF (Discounted Cash Flow)** převádí odhadované budoucí volné peněžní toky firmy na jejich dnešní hodnotu pomocí diskontní sazby (**WACC**). "
            "Součet těchto dnešních hodnot plus hodnota firmy po pátém roce (terminální hodnota) minus čistý dluh dává vnitřní hodnotu pro akcionáře."
        )

    if "owner earnings" in q:
        if as_kid:
            return (
                "**Owner Earnings** jsou peníze, které majiteli firmy skutečně zůstanou v kapse na konci roku poté, co zaplatí všechny opravy strojů, aby firma dál fungovala stejně dobře jako dřív."
            )
        return (
            "**Owner Earnings** (vlastnické zisky podle Warrena Buffetta) se počítají jako: "
            "`Čistý zisk + Odpisy - Akciové odměny (SBC) - Udržovací kapitálové výdaje (Maintenance CapEx)`. "
            "Na rozdíl od běžného FCF nepenalizují firmu za investice do dalšího růstu."
        )

    if "roic" in q:
        if as_kid:
            return (
                "**ROIC** ukazuje, kolik korun zisku dokáže firma vyrobit z každé stokoruny, kterou do ní lidé vložili. Čím vyšší číslo než WACC, tím šikovnější firma."
            )
        return (
            "**ROIC (Return on Invested Capital)** měří návratnost veškerého investovaného kapitálu (`NOPAT / Investovaný kapitál`). "
            "Pokud je ROIC dlouhodobě vyšší než WACC, firma vytváří ekonomickou hodnotu."
        )

    return (
        "Externí bezplatný AI server je momentálně přetížený (limit požadavků), takže nedokázal vygenerovat volnou odpověď. "
        "Pokud si vložíš vlastní bezplatný **Gemini API klíč** (z Google AI Studio) do souboru `.env` jako `GEMINI_API_KEY=...`, "
        "bude poradce odpovídat okamžitě a bez omezení na jakoukoliv otázku."
    )


def _call_pollinations_get_or_post(messages: list[dict], question: str, sys_prompt: str) -> str | None:
    # 1. POST na OpenAI-kompatibilní endpoint
    for model_name in ("openai", "mistral", "llama"):
        try:
            payload = {
                "model": model_name,
                "messages": messages,
                "temperature": 0.65,
                "seed": random.randint(1, 999999),
            }
            req = urllib.request.Request(
                "https://text.pollinations.ai/openai",
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "User-Agent": f"Mozilla/5.0 Tensegrity/{random.randint(100,999)}",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=12) as resp:
                raw_json = json.loads(resp.read().decode("utf-8", errors="ignore"))
            choices = raw_json.get("choices") or []
            if choices:
                content = str((choices[0].get("message") or {}).get("content") or "").strip()
                if content and len(content) > 5 and "rate limit" not in content.lower():
                    return content
        except Exception:
            continue

    # 2. Přímý GET fallback (často projde i tehdy, když POST endpoint hlásí rate-limit)
    try:
        compact_prompt = f"{sys_prompt[:900]}\n\nDotaz uživatele: {question}"
        get_url = f"https://text.pollinations.ai/{quote(compact_prompt)}?model=openai&seed={random.randint(1,999999)}"
        req_get = urllib.request.Request(
            get_url,
            headers={"User-Agent": f"Mozilla/5.0 TensegrityGET/{random.randint(100,999)}"},
        )
        with urllib.request.urlopen(req_get, timeout=12) as resp_get:
            txt = resp_get.read().decode("utf-8", errors="ignore").strip()
        if txt and len(txt) > 5 and "rate limit" not in txt.lower() and not txt.startswith("<!DOCTYPE"):
            return txt
    except Exception:
        pass

    return None


def generate_pure_ai_advisor_answer(question: str, history: list[dict], d: dict | None) -> str:
    sys_prompt = _build_ai_advisor_system_prompt(d)
    history_lines = []
    for m in history[-8:]:
        role_label = "Uživatel" if m.get("role") == "user" else "Poradce"
        history_lines.append(f"{role_label}: {str(m.get('text', ''))[:600]}")
    history_block = "\n".join(history_lines)

    if GEMINI_API_KEY:
        try:
            full_prompt = f"{sys_prompt}\n\nHistorie konverzace:\n{history_block}\n\nAktuální dotaz: {question}"
            prompt_body = {
                "contents": [{"parts": [{"text": full_prompt}]}],
                "generationConfig": {"temperature": 0.65},
            }
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            req = urllib.request.Request(
                url,
                data=json.dumps(prompt_body).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=20) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            candidates = raw.get("candidates") or []
            if candidates:
                parts = candidates[0].get("content", {}).get("parts") or []
                if parts and parts[0].get("text"):
                    return _sanitize_advisor_output(parts[0]["text"])
        except Exception as err:
            logger.warning("Gemini API selhalo (%s), zkouším záložní AI...", err)

    messages = [{"role": "system", "content": sys_prompt}]
    for m in history[-8:]:
        messages.append({
            "role": "user" if m.get("role") == "user" else "assistant",
            "content": str(m.get("text", ""))[:600],
        })
    messages.append({"role": "user", "content": question})

    free_reply = _call_pollinations_get_or_post(messages, question, sys_prompt)
    if free_reply:
        return _sanitize_advisor_output(free_reply)

    return _sanitize_advisor_output(_offline_smart_czech_knowledge_fallback(question, d))


# ==============================================================================
# 9. API ENDPOINTY, GRAF, EXTERNÍ CRON A EXPORTY
# ==============================================================================

@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    logo_path = resolve_brand_logo_path(prefer_white=False)
    if logo_path:
        return FileResponse(logo_path)
    raise HTTPException(status_code=404, detail="Ikona nenalezena.")


@app.get("/brand-logo.png", include_in_schema=False)
async def brand_logo():
    logo_path = resolve_brand_logo_path(prefer_white=True) or resolve_brand_logo_path(prefer_white=False)
    if logo_path:
        return FileResponse(logo_path)
    raise HTTPException(status_code=404, detail="Logo nenalezeno.")


@app.post("/api/settings/gemini-key")
async def update_gemini_key_endpoint(request: Request):
    try:
        payload = await request.json()
        key_val = str(payload.get("api_key", "")).strip()
        if not key_val or len(key_val) < 15:
            raise HTTPException(status_code=422, detail="Zadejte platný Gemini API klíč.")
        save_gemini_key_to_env(key_val)
        return {"status": "ok", "message": "Gemini API klíč byl aktivován a uložen."}
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/chart/{ticker}")
async def get_interactive_chart_endpoint(ticker: str, timeframe: str = Query("1R")):
    try:
        return await asyncio.to_thread(fetch_interactive_chart_series, ticker, timeframe)
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.api_route("/trigger-newsletter", methods=["GET", "POST"])
async def trigger_newsletter_endpoint(key: str = Query("")):
    if CRON_SECRET and key != CRON_SECRET:
        raise HTTPException(status_code=403, detail="Neplatný bezpečnostní klíč plánovače (CRON_SECRET).")
    try:
        now = get_prague_now()
        subscribers = get_subscribers_internal()
        if not subscribers:
            return {"status": "ok", "message": "Žádní odběratelé k rozeslání.", "time": now.isoformat()}

        res = await dispatch_newsletter_to_recipients(subscribers)
        if res["sent"] > 0:
            slot_key = f"{now.strftime('%Y-%m-%d')}-{now.hour}"
            sent_slots = load_sent_slots()
            sent_slots.add(slot_key)
            save_sent_slots(sent_slots)

        if res["failed"] > 0 and res["sent"] == 0:
            raise HTTPException(status_code=500, detail=f"Odeslání selhalo: {'; '.join(res['errors'])}")

        return {
            "status": "ok",
            "message": f"Zpravodaj odeslán: {res['sent']} úspěšně, {res['failed']} chyb.",
            "errors": res["errors"],
            "time": now.isoformat(),
        }
    except HTTPException:
        raise
    except Exception as err:
        logger.error("Chyba při externím spuštění zpravodaje: %s", err)
        raise HTTPException(status_code=500, detail=str(err))


@app.api_route("/subscribe", methods=["GET", "POST"])
async def handle_subscribe(request: Request, email: str = Query("")):
    candidate = email.strip()
    if request.method == "POST" and not candidate:
        try:
            payload = await request.json()
            candidate = str(payload.get("email", "")).strip()
        except Exception:
            candidate = ""
    if not is_valid_email_address(candidate):
        raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")
    created = save_subscriber_private(candidate)
    return {
        "status": "ok",
        "message": (
            "Pravidelný odběr byl bezpečně zaevidován."
            if created
            else "Tato e-mailová adresa je již zaevidována."
        ),
    }


@app.post("/api/newsletter/send-now")
async def send_newsletter_now(request: Request):
    try:
        payload = await request.json()
        email = str(payload.get("email", "")).strip()
        if not is_valid_email_address(email):
            raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")

        res = await dispatch_newsletter_to_recipients([email])
        if res["sent"] == 0:
            raise HTTPException(status_code=500, detail=f"Odeslání selhalo: {'; '.join(res['errors'])}")

        return {
            "status": "ok",
            "message": f"Zpravodaj v češtině včetně PDF byl úspěšně odeslán na {email}.",
        }
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/valuation/{ticker}")
def get_institutional_terminal_valuation(
    ticker: str,
    leverage: float | None = Query(None, ge=0.0, le=1.0),
    growth: float | None = Query(None, gt=-0.95, le=5.0),
    exit_mult: float | None = Query(None, gt=0.5, le=150.0),
):
    try:
        return run_valuation_logic(ticker, leverage, growth, exit_mult)
    except ValueError as err:
        raise HTTPException(status_code=422, detail=str(err))
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.post("/api/advisor/{ticker}")
async def ask_terminal_advisor(ticker: str, request: Request):
    try:
        payload = await request.json()
        question = str(payload.get("question", "")).strip()[:2000]
        history = payload.get("history") or []
        scenario = payload.get("scenario") or {}
        leverage = safe_float(scenario.get("leverage"), None) if scenario.get("leverage") is not None else None
        growth = safe_float(scenario.get("growth"), None) if scenario.get("growth") is not None else None
        exit_mult = safe_float(scenario.get("exit_mult"), None) if scenario.get("exit_mult") is not None else None
        if not question:
            raise HTTPException(status_code=422, detail="Zadejte prosím svůj dotaz.")

        d = None
        try:
            d = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)
        except Exception:
            d = None

        answer = await asyncio.to_thread(generate_pure_ai_advisor_answer, question, history, d)
        return {"answer": answer}
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=400, detail=str(err))


@app.get("/api/news/{ticker}")
def get_ticker_news(ticker: str):
    try:
        resolved = resolve_company_ticker(ticker)
        raw_news = yf.Ticker(resolved).news or []
        articles = []
        for item in raw_news[:6]:
            content = item.get("content", item) if isinstance(item, dict) else {}
            if not isinstance(content, dict):
                continue
            provider = content.get("provider", {})
            clickthrough = content.get("clickThroughUrl") or content.get("canonicalUrl") or {}
            raw_title = content.get("title") or item.get("title") or "Bez názvu"
            articles.append({
                "raw_title": raw_title,
                "raw_text": content.get("summary") or raw_title,
                "source": provider.get("displayName") or item.get("publisher") or "Finanční zpravodajství",
                "url": clickthrough.get("url") or item.get("link") or "#",
            })
        translated = _translate_and_summarize_batch_czech(articles)
        out = [
            {
                "title": a.get("title_cs") or a["raw_title"],
                "source": a["source"],
                "url": a["url"],
            }
            for a in translated
        ]
        return {"ticker": resolved, "articles": out}
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/forex-rates")
async def get_forex_rates():
    return await asyncio.to_thread(get_live_forex_rates)


@app.get("/api/export/excel/{ticker}")
def export_valuation_excel(
    ticker: str,
    leverage: float | None = Query(None, ge=0.0, le=1.0),
    growth: float | None = Query(None, gt=-0.95, le=5.0),
    exit_mult: float | None = Query(None, gt=0.5, le=150.0),
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        wb = openpyxl.Workbook()
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        ws_sum = wb.active
        ws_sum.title = "Syntéza ocenění"

        logo_path = resolve_brand_logo_path(prefer_white=False)
        if logo_path:
            try:
                xl_logo = ExcelImage(logo_path)
                xl_logo.width = 48
                xl_logo.height = 34
                ws_sum.add_image(xl_logo, "A1")
            except Exception:
                pass

        ws_sum["B1"] = "TENSEGRITY TERMINAL | INSTITUCIONÁLNÍ VALUAČNÍ MODEL"
        ws_sum["B1"].font = Font(size=14, bold=True)
        ws_sum["B2"] = f"Společnost: {data['company_name']} ({data['ticker']}) | Zdroj: {data['data_quality']['primary_statement_source']}"
        ws_sum.append([])
        ws_sum.append(["Metoda ocenění", "Váha (%)", f"Vnitřní hodnota ({data['currency']})", "Vážený příspěvek", "Rozdíl vs. trh (%)"])
        for cell in ws_sum[4]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row in data["valuation_synthesis"]["methods"]:
            ws_sum.append([
                row["method"],
                row["weight_percent"],
                row["fair_value"] if row["fair_value"] is not None else "N/A",
                row["contribution"],
                row["vs_market_percent"] if row["vs_market_percent"] is not None else "N/A",
            ])

        ws_sum.append([])
        ws_sum.append(["Vážená vnitřní hodnota", data["valuation_synthesis"]["weighted_fair_value"]])
        ws_sum.append(["Aktuální tržní cena akcie", data["dcf_valuation"]["current_share_price"]])
        ws_sum.append(["Owner Earnings (celkem)", data["fundamentals"]["owner_earnings"] if data["fundamentals"]["owner_earnings"] is not None else "N/A"])
        ws_sum.append(["Look-Through Earnings", data["fundamentals"]["look_through_earnings"] if data["fundamentals"]["look_through_earnings"] is not None else "N/A"])
        ws_sum.append(["Udržovací CapEx (odhad)", data["fundamentals"]["maintenance_capex"] if data["fundamentals"]["maintenance_capex"] is not None else "N/A"])
        ws_sum.append(["Bezpečnostní rezerva (%)", data["decision_engine"]["margin_of_safety_percent"]])
        ws_sum.append(["Aktivní cenová zóna", data["decision_engine"]["zone"]])
        ws_sum.append(["Doporučená alokace tranše (%)", data["decision_engine"]["tranche_allocation_percent"]])

        ws_dcf = wb.create_sheet("Peněžní toky a diskont")
        ws_dcf.append(["Rok", "Růst peněžních toků (%)", "Volný peněžní tok", "Diskontní faktor", "Současná hodnota toku"])
        for cell in ws_dcf[1]:
            cell.fill = header_fill
            cell.font = header_font
        for cf in data["dcf_valuation"]["cash_flows"]:
            ws_dcf.append([cf["rok"], cf["rust_procent"], cf["volny_penezni_tok"], cf["diskontni_faktor"], cf["soucasna_hodnota_fcf"]])

        ws_mc = wb.create_sheet("Scénáře a Monte Carlo")
        ws_mc.append(["Percentil / Scénář", f"Hodnota na akcii ({data['currency']})"])
        for cell in ws_mc[1]:
            cell.fill = header_fill
            cell.font = header_font
        mc = data["monte_carlo"]
        for label, val in [
            ("Pesimistický scénář (25 %)", data["scenarios"]["bear"]["fair_value"]),
            ("Základní scénář (50 %)", data["scenarios"]["base"]["fair_value"]),
            ("Optimistický scénář (25 %)", data["scenarios"]["bull"]["fair_value"]),
            ("Monte Carlo P5", mc["p5"]),
            ("Monte Carlo P10", mc["p10"]),
            ("Monte Carlo P25", mc["p25"]),
            ("Monte Carlo P50 (Medián)", mc["p50"]),
            ("Monte Carlo P75", mc["p75"]),
            ("Monte Carlo P90", mc["p90"]),
            ("Monte Carlo P95", mc["p95"]),
        ]:
            ws_mc.append([label, val if val is not None else "N/A"])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"Tensegrity_Valuation_{data['spolecnost']}.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/export/pdf/{ticker}")
def export_valuation_pdf(
    ticker: str,
    leverage: float | None = Query(None, ge=0.0, le=1.0),
    growth: float | None = Query(None, gt=-0.95, le=5.0),
    exit_mult: float | None = Query(None, gt=0.5, le=150.0),
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        title_s = ParagraphStyle("T1", parent=styles["Heading1"], fontName=PDF_FONT_BOLD, fontSize=15, textColor=colors.HexColor("#0f172a"), spaceAfter=4)
        sub_s = ParagraphStyle("S1", parent=styles["Normal"], fontName=PDF_FONT_REGULAR, fontSize=9.5, textColor=colors.HexColor("#475569"), spaceAfter=12)
        sec_s = ParagraphStyle("H2", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=11.5, textColor=colors.HexColor("#0284c7"), spaceBefore=10, spaceAfter=6)
        note_s = ParagraphStyle("N1", parent=styles["Normal"], fontName=PDF_FONT_REGULAR, fontSize=9, leading=13, textColor=colors.HexColor("#334155"), spaceBefore=8)

        synth = data["valuation_synthesis"]
        dec = data["decision_engine"]
        dcf = data["dcf_valuation"]
        mc = data["monte_carlo"]
        fund = data["fundamentals"]
        curr = data["currency"]

        def fmt_or_na(v):
            return f"{v:,.2f} {curr}" if v is not None else "N/A"

        elements = []
        logo_path = resolve_brand_logo_path(prefer_white=False)
        if logo_path:
            try:
                logo_img = PdfImage(logo_path, width=38, height=26, kind="proportional")
                hdr_tbl = Table([[logo_img, Paragraph(f"TENSEGRITY TERMINAL — VALUAČNÍ ZPRÁVA ({data['company_name']})", title_s)]], colWidths=[48, 470])
                hdr_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
                elements.append(hdr_tbl)
            except Exception:
                elements.append(Paragraph(f"TENSEGRITY TERMINAL — VALUAČNÍ ZPRÁVA ({data['company_name']})", title_s))
        else:
            elements.append(Paragraph(f"TENSEGRITY TERMINAL — VALUAČNÍ ZPRÁVA ({data['company_name']})", title_s))

        elements.extend([
            Paragraph(f"Ticker: {data['ticker']} | Zdroj výkazů: {data['data_quality']['primary_statement_source']} | Měna: {curr}", sub_s),
            Paragraph("1. SYNTÉZA VNITŘNÍ HODNOTY, OWNER EARNINGS A NÁKUPNÍ PÁSMA", sec_s),
        ])

        mc_str = f"{mc['p50']:,.2f} ({mc['p10']:,.2f} – {mc['p90']:,.2f}) {curr}" if mc["p50"] is not None else "N/A"
        tbl = Table(
            [
                ["Ukazatel / Oceňovací přístup", "Výstup modelu"],
                ["Aktuální tržní cena akcie", fmt_or_na(dcf["current_share_price"])],
                ["Vážená vnitřní hodnota", fmt_or_na(synth["weighted_fair_value"])],
                ["Vícefázový výnosový model pro podnik (FCFF)", fmt_or_na(dcf["intrinsic_value_per_share"])],
                ["Owner Earnings / FCFE model", fmt_or_na(data["fcfe_valuation"]["intrinsic_value_per_share"])],
                ["Owner Earnings / akcii (po udržovacím CapEx)", fmt_or_na(fund["owner_earnings_per_share"])],
                ["Model reziduálního zisku (Residual Income)", fmt_or_na(data["residual_income"]["intrinsic_value_per_share"])],
                ["Strukturovaný součet částí (SOTP)", fmt_or_na(data["metoda_souctu_casti"]["sotp_per_share"])],
                ["Srovnatelné tržní násobky", fmt_or_na(data["peer_comps"]["blended_implied_price"])],
                ["Medián Monte Carlo (P50) / P10–P90", mc_str],
                ["Bezpečnostní rezerva", f"{dec['margin_of_safety_percent']:.2f} %"],
                ["Aktivní cenová zóna a alokace", f"{dec['zone']} (Tranše {dec['tranche_allocation_percent']} %)"],
                ["Index integrity dat / Piotroski skóre", f"{synth['integrity_score']} / 100 | {data['risk_analysis']['piotroski_f_score']} / 9"],
            ],
            colWidths=[295, 220],
        )
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), PDF_FONT_BOLD),
            ("FONTNAME", (0, 1), (-1, -1), PDF_FONT_REGULAR),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("PADDING", (0, 0), (-1, -1), 6),
        ]))
        elements.append(tbl)
        elements.append(Paragraph(synth["disagreement_note"], note_s))

        doc.build(elements)
        buffer.seek(0)
        filename = f"Tensegrity_Terminal_Report_{data['spolecnost']}.pdf"
        return StreamingResponse(
            buffer,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


# ==============================================================================
# 10. KOMPLETNÍ WEBOVÉ ROZHRANÍ (ZVĚTŠENÉ KPI BOXY + OKAMŽITÝ OHLC PANEL + AI KLÍČ)
# ==============================================================================

@app.get("/", response_class=HTMLResponse)
def read_root():
    return r"""<!DOCTYPE html>
<html lang="cs">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tensegrity Terminal | Institucionální Valuační Platforma</title>
    <link rel="icon" type="image/png" href="/favicon.ico?v=10">
    <link rel="shortcut icon" type="image/png" href="/favicon.ico?v=10">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/lightweight-charts@4.1.3/dist/lightweight-charts.standalone.production.js"></script>
    <style>
        :root {
            --bg-main: #060911;
            --bg-card: rgba(15, 22, 37, 0.88);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --accent-blue: #38bdf8;
            --accent-blue-glow: rgba(56, 189, 248, 0.2);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --success: #10b981;
            --warning: #f59e0b;
            --yellow: #eab308;
            --danger: #ef4444;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: var(--bg-main);
            color: var(--text-primary);
            min-height: 100vh;
            padding: 24px 32px;
        }
        .container { max-width: 1440px; margin: 0 auto; }
        header {
            display: flex; justify-content: space-between; align-items: center;
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 16px 28px; margin-bottom: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.5); flex-wrap: wrap; gap: 16px;
        }
        .brand-wrap { display: flex; align-items: center; gap: 14px; }
        .brand-logo-box { display: flex; align-items: center; justify-content: center; }
        .brand-logo-box img { height: 32px; width: auto; display: block; }
        .brand-title { font-size: 16px; font-weight: 800; letter-spacing: 4px; color: #fff; text-transform: uppercase; }
        .brand-title span { color: var(--accent-blue); }
        .brand-subtitle { font-size: 10px; letter-spacing: 4.5px; color: var(--text-secondary); text-transform: uppercase; margin-top: 2px; }
        .controls { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
        .forex-strip {
            display: flex; align-items: center; gap: 12px; padding: 8px 14px;
            border: 1px solid var(--border-subtle); border-radius: 8px;
            background: rgba(6, 9, 17, 0.65); color: var(--text-secondary);
            font: 11px 'JetBrains Mono', monospace;
        }
        .forex-strip strong { color: var(--text-primary); }
        .search-wrapper input {
            background: rgba(6, 9, 17, 0.85); border: 1px solid var(--border-subtle);
            color: var(--text-primary); font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 13px; font-weight: 600; padding: 10px 14px; border-radius: 10px;
            outline: none; width: 230px; letter-spacing: 0.3px;
        }
        .search-wrapper input:focus { border-color: var(--accent-blue); box-shadow: 0 0 12px var(--accent-blue-glow); }
        .btn {
            font-family: 'Plus Jakarta Sans', sans-serif; font-size: 12px; font-weight: 700;
            padding: 10px 16px; border-radius: 10px; border: 1px solid transparent; cursor: pointer;
            display: inline-flex; align-items: center; gap: 6px; transition: all 0.2s ease;
        }
        .btn-primary { background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%); color: #040914; }
        .btn-primary:hover { opacity: 0.94; transform: translateY(-1px); }
        .btn-secondary { background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.35); color: #34d399; }
        .btn-excel { background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.3); color: #34d399; }
        .btn-pdf { background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); color: #f87171; }
        .btn-news { background: rgba(148, 163, 184, 0.12); border: 1px solid var(--border-subtle); color: var(--text-primary); }
        .btn-news:hover { border-color: var(--accent-blue); color: var(--accent-blue); }

        .scenario-bar {
            display: flex; align-items: center; flex-wrap: wrap; gap: 14px 22px;
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 12px; padding: 12px 22px; margin-bottom: 22px;
        }
        .scenario-heading { color: var(--accent-blue); font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; }
        .scenario-field { display: flex; align-items: center; gap: 10px; color: var(--text-secondary); font-size: 12px; }
        .scenario-input input {
            width: 72px; padding: 6px 8px; border: 1px solid var(--border-subtle); border-radius: 6px;
            background: rgba(6, 9, 17, 0.75); color: var(--text-primary); font: 600 13px 'JetBrains Mono', monospace; text-align: right;
        }
        .audit-pill { margin-left: auto; font: 11px 'JetBrains Mono', monospace; color: var(--text-muted); }

        .data-warning-banner {
            display: none; margin-bottom: 16px; padding: 12px 16px; border-radius: 10px;
            background: rgba(245, 158, 11, 0.12); border: 1px solid rgba(245, 158, 11, 0.38);
            color: #fbbf24; font-size: 12.5px; line-height: 1.5;
        }

        .company-identity-banner {
            display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 18px;
            margin-top: 18px; padding: 18px 22px; border-radius: 14px;
            background: rgba(6, 9, 17, 0.8); border: 1px solid rgba(56, 189, 248, 0.25);
        }
        .company-identity-left { display: flex; align-items: center; gap: 16px; }
        .company-logo-frame {
            width: 50px; height: 50px; border-radius: 12px; background: rgba(255, 255, 255, 0.95);
            display: flex; align-items: center; justify-content: center; overflow: hidden;
            border: 1px solid rgba(255,255,255,0.25); flex-shrink: 0;
        }
        .company-logo-frame img { width: 36px; height: 36px; object-fit: contain; display: block; }
        .company-logo-fallback {
            width: 50px; height: 50px; border-radius: 12px;
            background: linear-gradient(135deg, rgba(56,189,248,0.25), rgba(2,132,199,0.45));
            color: #38bdf8; font-weight: 800; font-size: 18px;
            display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        }
        .company-title-main { font-size: 18px; font-weight: 800; color: #f8fafc; letter-spacing: 0.2px; }
        .company-ticker-pill {
            display: inline-block; margin-left: 8px; padding: 3px 9px; border-radius: 6px;
            background: rgba(56, 189, 248, 0.16); color: #38bdf8;
            font: 700 12.5px 'JetBrains Mono', monospace;
        }
        .company-meta-sub { font-size: 12.5px; color: var(--text-secondary); margin-top: 4px; }

        .kpi-hero-group {
            display: flex; align-items: stretch; gap: 16px; flex-wrap: wrap;
        }
        .kpi-hero-box {
            min-width: 225px; padding: 14px 20px; border-radius: 12px;
            display: flex; flex-direction: column; justify-content: center; gap: 4px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.35);
        }
        .kpi-hero-box.market {
            background: linear-gradient(145deg, rgba(56, 189, 248, 0.14), rgba(15, 23, 42, 0.92));
            border: 1.5px solid rgba(56, 189, 248, 0.42);
        }
        .kpi-hero-box.fair {
            background: linear-gradient(145deg, rgba(16, 185, 129, 0.16), rgba(15, 23, 42, 0.92));
            border: 1.5px solid rgba(16, 185, 129, 0.45);
        }
        .kpi-hero-label {
            font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px;
        }
        .kpi-hero-box.market .kpi-hero-label { color: #7dd3fc; }
        .kpi-hero-box.fair .kpi-hero-label { color: #6ee7b7; }
        .kpi-hero-value {
            font: 800 22px 'JetBrains Mono', monospace; color: #ffffff; letter-spacing: -0.3px;
        }
        .kpi-hero-sub {
            font: 600 11.5px 'JetBrains Mono', monospace; color: var(--text-secondary);
        }

        .tv-toolbar {
            display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;
            margin-bottom: 12px;
        }
        .tv-tf-group { display: flex; gap: 6px; flex-wrap: wrap; }
        .tv-tf-btn {
            background: rgba(6, 9, 17, 0.75); border: 1px solid var(--border-subtle);
            color: var(--text-secondary); font: 700 11px 'JetBrains Mono', monospace;
            padding: 6px 11px; border-radius: 8px; cursor: pointer; transition: all 0.15s ease;
        }
        .tv-tf-btn:hover { border-color: var(--accent-blue); color: #fff; }
        .tv-tf-btn.active { background: rgba(56, 189, 248, 0.18); border-color: var(--accent-blue); color: #38bdf8; }
        .tv-readout {
            display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
            font: 600 12px 'JetBrains Mono', monospace; color: var(--text-secondary);
        }
        .tv-readout strong { color: #fff; }
        #tvChartContainer {
            width: 100%; height: 360px; border-radius: 12px; overflow: hidden;
            border: 1px solid var(--border-subtle); background: #070b14;
        }

        .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(460px, 1fr)); gap: 22px; margin-bottom: 22px; }
        .grid-3 { display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 22px; margin-bottom: 22px; }
        .grid-4 { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-top: 16px; }
        .card {
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 22px 26px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .matrix-card { margin-bottom: 22px; }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid var(--border-subtle); flex-wrap: wrap; gap: 8px; }
        .card-title { font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.4px; color: var(--text-secondary); }
        .badge { font-size: 11px; font-family: 'JetBrains Mono', monospace; padding: 3px 8px; border-radius: 6px; background: rgba(56, 189, 248, 0.1); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.25); }
        .badge-verified { background: rgba(16, 185, 129, 0.14); color: #34d399; border-color: rgba(16, 185, 129, 0.35); }
        .badge-warn { background: rgba(245, 158, 11, 0.14); color: #fbbf24; border-color: rgba(245, 158, 11, 0.35); }
        .metric-row { display: flex; justify-content: space-between; align-items: center; margin: 10px 0; font-size: 13.5px; gap: 10px; }
        .metric-label { color: var(--text-secondary); }
        .metric-value { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: var(--text-primary); font-size: 13.5px; text-align: right; }
        .metric-value.highlight { color: var(--accent-blue); font-size: 14.5px; }

        .chart-container { position: relative; height: 280px; width: 100%; margin-top: 8px; }

        .thesis-layout { display: grid; grid-template-columns: auto 1fr; gap: 22px; align-items: center; }
        .traffic-lights { display: flex; gap: 10px; padding: 12px 14px; border: 1px solid var(--border-subtle); border-radius: 999px; background: rgba(6, 9, 17, 0.7); }
        .traffic-light { width: 18px; height: 18px; border-radius: 50%; opacity: 0.2; }
        .traffic-light.active { opacity: 1; box-shadow: 0 0 16px currentColor; }
        .traffic-light.green { color: #10b981; background: #10b981; }
        .traffic-light.yellow { color: #eab308; background: #eab308; }
        .traffic-light.orange { color: #f59e0b; background: #f59e0b; }
        .traffic-light.red { color: #ef4444; background: #ef4444; }

        .zone-box { background: rgba(6, 9, 17, 0.65); border: 1px solid var(--border-subtle); border-radius: 12px; padding: 16px; }
        .zone-box.green { border-color: rgba(16, 185, 129, 0.35); }
        .zone-box.yellow { border-color: rgba(234, 179, 8, 0.35); }
        .zone-box.orange { border-color: rgba(245, 158, 11, 0.35); }
        .zone-box.red { border-color: rgba(239, 68, 68, 0.35); }
        .zone-title { font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px; }
        .zone-box.green .zone-title { color: #34d399; }
        .zone-box.yellow .zone-title { color: #facc15; }
        .zone-box.orange .zone-title { color: #fb923c; }
        .zone-box.red .zone-title { color: #f87171; }
        .zone-val { font-family: 'JetBrains Mono', monospace; font-size: 15px; font-weight: 700; color: #fff; margin-bottom: 4px; }

        .factor-pill-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; margin-top: 10px; }
        .factor-item {
            background: rgba(6, 9, 17, 0.65); border: 1px solid var(--border-subtle);
            border-radius: 10px; padding: 12px 14px;
        }
        .factor-item-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; font-size: 12px; font-weight: 700; }
        .factor-status-ok { color: #34d399; font-family: 'JetBrains Mono', monospace; font-size: 11px; }
        .factor-status-warn { color: #fbbf24; font-family: 'JetBrains Mono', monospace; font-size: 11px; }
        .factor-item-desc { font-size: 11.5px; color: var(--text-secondary); line-height: 1.45; }

        .table-responsive { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 8px; }
        th, td { padding: 10px 14px; text-align: center; font-size: 12.5px; border-bottom: 1px solid var(--border-subtle); }
        th { background: rgba(6, 9, 17, 0.8); color: var(--text-secondary); font-weight: 700; text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.6px; }
        th:first-child, td:first-child { text-align: left; }
        td { font-family: 'JetBrains Mono', monospace; color: var(--text-primary); }

        .plain-guide { margin-top: 10px; font-size: 12px; color: var(--text-secondary); line-height: 1.55; }
        .advisor-thread {
            display: flex; flex-direction: column; gap: 14px; margin: 16px 0;
            max-height: 460px; overflow-y: auto; padding: 14px;
            border: 1px solid var(--border-subtle); border-radius: 12px; background: rgba(6, 9, 17, 0.72);
        }
        .advisor-message { max-width: 88%; padding: 12px 16px; border-radius: 12px; font-size: 13.5px; line-height: 1.65; }
        .advisor-message.user { align-self: flex-end; background: rgba(56, 189, 248, 0.14); border: 1px solid rgba(56, 189, 248, 0.35); color: #e0f2fe; }
        .advisor-message.assistant { align-self: flex-start; background: rgba(15, 22, 37, 0.96); border: 1px solid var(--border-subtle); color: #f8fafc; }
        .advisor-form { display: flex; gap: 10px; flex-wrap: wrap; }
        .advisor-input {
            flex: 1; min-width: 240px; padding: 12px 16px; border: 1px solid var(--border-subtle);
            border-radius: 10px; background: rgba(6, 9, 17, 0.85); color: var(--text-primary);
            font: 13px 'Plus Jakarta Sans', sans-serif;
        }
        .advisor-input:focus { outline: none; border-color: var(--accent-blue); }

        .modal-backdrop {
            display: none; position: fixed; inset: 0; background: rgba(2, 6, 17, 0.82);
            backdrop-filter: blur(8px); z-index: 1000; align-items: center; justify-content: center; padding: 20px;
        }
        .modal-backdrop.open { display: flex; }
        .modal-card {
            background: #0d1322; border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 18px;
            max-width: 560px; width: 100%; padding: 26px 28px; box-shadow: 0 25px 60px rgba(0,0,0,0.8);
        }
        .modal-title { font-size: 17px; font-weight: 800; color: #fff; margin-bottom: 8px; }
        .modal-sub { font-size: 13px; color: var(--text-secondary); line-height: 1.6; margin-bottom: 18px; }
        .modal-input {
            width: 100%; background: rgba(6, 9, 17, 0.9); border: 1px solid var(--border-subtle);
            color: #fff; padding: 12px 14px; border-radius: 10px; font-size: 14px; margin-bottom: 14px; outline: none;
        }
        .modal-input:focus { border-color: var(--accent-blue); }
        .modal-actions { display: flex; gap: 10px; flex-wrap: wrap; }
        .modal-status { margin-top: 12px; font-size: 12px; font-family: 'JetBrains Mono', monospace; color: #34d399; }

        .loading-overlay {
            display: none; position: fixed; inset: 0; background: rgba(6, 9, 17, 0.86);
            backdrop-filter: blur(10px); z-index: 999; flex-direction: column; align-items: center; justify-content: center;
        }
        .spinner { width: 44px; height: 44px; border: 3px solid rgba(56, 189, 248, 0.15); border-top-color: var(--accent-blue); border-radius: 50%; animation: spin 0.9s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
<div class="loading-overlay" id="loader">
    <div class="spinner"></div>
    <p style="margin-top: 14px; font-family: 'JetBrains Mono', monospace; font-size: 12px; color: var(--accent-blue); letter-spacing: 1px;">OVĚŘUJI REGULATORNÍ VÝKAZY &amp; SPOUŠTÍM ENGINE...</p>
</div>

<div class="modal-backdrop" id="newsletterModal">
    <div class="modal-card">
        <div class="modal-title">Zpravodaj Tensegrity Terminal</div>
        <div class="modal-sub">
            Automatické rozesílání českého makroekonomického a zpravodajského přehledu včetně PDF přílohy probíhá každou
            <strong>středu a pátek v 7:00</strong> a v <strong>neděli v 19:00</strong>.
        </div>
        <input type="email" id="newsletterEmailInput" class="modal-input" placeholder="vas.email@domena.cz">
        <div class="modal-actions">
            <button type="button" class="btn btn-primary" onclick="subscribeNewsletter()">Přidat do pravidelného odběru</button>
            <button type="button" class="btn btn-secondary" onclick="sendNewsletterNow()">Odeslat aktuální vydání ihned (s PDF)</button>
            <button type="button" class="btn btn-news" onclick="closeNewsletterModal()">Zavřít</button>
        </div>
        <div id="newsletterStatus" class="modal-status"></div>
    </div>
</div>

<div class="modal-backdrop" id="aiKeyModal">
    <div class="modal-card">
        <div class="modal-title">Aktivace vlastního Gemini API klíče</div>
        <div class="modal-sub">
            Vložte bezplatný klíč z <strong>Google AI Studio</strong> (začíná znaky <code>AIza...</code>).
            Klíč se uloží lokálně do souboru <code>.env</code> a AI poradce pak bude odpovídat okamžitě na jakoukoliv otázku bez omezení.
        </div>
        <input type="password" id="geminiKeyInput" class="modal-input" placeholder="AIzaSy...">
        <div class="modal-actions">
            <button type="button" class="btn btn-primary" onclick="saveGeminiKey()">Uložit a aktivovat AI klíč</button>
            <button type="button" class="btn btn-news" onclick="closeAiKeyModal()">Zavřít</button>
        </div>
        <div id="aiKeyStatus" class="modal-status"></div>
    </div>
</div>

<div class="container">
    <header>
        <div class="brand-wrap">
            <div class="brand-logo-box">
                <img src="/brand-logo.png?v=10" alt="Tensegrity Terminal Logo" onerror="this.style.display='none'">
            </div>
            <div>
                <div class="brand-title">TENSEGRITY<span>TERMINAL</span></div>
                <div class="brand-subtitle">INSTITUCIONÁLNÍ VALUACE &bull; SEC EDGAR &amp; DATA QUALITY GATE</div>
            </div>
        </div>
        <div class="controls">
            <div class="search-wrapper">
                <input type="text" id="tickerInput" value="NVDA" placeholder="Ticker nebo název (např. ČEZ, SpaceX)" onkeydown="if(event.key==='Enter') runValuation(true)">
            </div>
            <div class="forex-strip" id="forexRatesBanner">
                <span>EUR/CZK <strong>24,45</strong></span>
                <span>USD/CZK <strong>22,30</strong></span>
                <span>GBP/CZK <strong>29,18</strong></span>
            </div>
            <button type="button" class="btn btn-primary" onclick="runValuation(true)">Analyzovat</button>
            <button type="button" class="btn btn-excel" onclick="downloadExcel()">Excel Model</button>
            <button type="button" class="btn btn-pdf" onclick="downloadPdf()">PDF Zpráva</button>
            <button type="button" class="btn btn-news" onclick="openNewsletterModal()">Newsletter</button>
        </div>
    </header>

    <div class="scenario-bar">
        <span class="scenario-heading">Parametry scénáře</span>
        <label class="scenario-field scenario-input">
            Cizí kapitál (% Dluhu)
            <input type="number" id="leverageInput" value="15" min="0" max="100" step="1"> %
        </label>
        <label class="scenario-field scenario-input">
            Růst peněžních toků
            <input type="number" id="growthInput" value="10" min="-90" max="200" step="0.5"> %
        </label>
        <label class="scenario-field scenario-input">
            Výstupní násobek
            <input type="number" id="exitMultInput" value="14" min="1" max="120" step="0.5"> x
        </label>
        <button type="button" class="btn btn-primary" onclick="runValuation(false)">Přepočítat model</button>
        <span class="audit-pill" id="auditStamp">Načítám auditní stopu...</span>
    </div>

    <div class="data-warning-banner" id="dataWarningBanner"></div>

    <!-- 1. ROZHODOVACÍ RÁMEC & 4 CENOVÉ ZÓNY -->
    <div class="card matrix-card">
        <div class="card-header">
            <span class="card-title">Rozhodovací rámec — 4 nákupní zóny &amp; alokace tranší</span>
            <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                <span class="badge badge-verified" id="dataSourceBadge">Zdroj výkazů: --</span>
                <span class="badge" id="integrityBadge">Index integrity: --</span>
            </div>
        </div>
        <div class="thesis-layout">
            <div class="traffic-lights">
                <span class="traffic-light green" id="light-GREEN"></span>
                <span class="traffic-light yellow" id="light-YELLOW"></span>
                <span class="traffic-light orange" id="light-ORANGE"></span>
                <span class="traffic-light red" id="light-RED"></span>
            </div>
            <div>
                <div style="font-size: 16px; font-weight: 800; color: #fff; margin-bottom: 4px;" id="thesisTitle">Načítám institucionální syntézu...</div>
                <div style="font-size: 13px; color: var(--text-secondary);" id="thesisSub"></div>
            </div>
        </div>

        <div class="company-identity-banner">
            <div class="company-identity-left">
                <div id="companyLogoContainer" class="company-logo-fallback">TT</div>
                <div>
                    <div>
                        <span class="company-title-main" id="companyFullName">NVIDIA Corporation</span>
                        <span class="company-ticker-pill" id="companyTickerBadge">NVDA</span>
                    </div>
                    <div class="company-meta-sub" id="companySectorMeta">Sektor: Technology &bull; Odvětví: Semiconductors &bull; Burza: NMS</div>
                </div>
            </div>
            <div class="kpi-hero-group">
                <div class="kpi-hero-box market">
                    <div class="kpi-hero-label">Aktuální tržní cena</div>
                    <div class="kpi-hero-value" id="companyPriceBig">--</div>
                    <div class="kpi-hero-sub" id="companyMcapSub">Tržní kap.: --</div>
                </div>
                <div class="kpi-hero-box fair">
                    <div class="kpi-hero-label">Vážená vnitřní hodnota</div>
                    <div class="kpi-hero-value" id="companyFairBig">--</div>
                    <div class="kpi-hero-sub" id="companyUpsideSub">Potenciál: --</div>
                </div>
            </div>
        </div>

        <div class="grid-4">
            <div class="zone-box green">
                <div class="zone-title">ZÓNA GREEN (Tranše 55 %)</div>
                <div class="zone-val" id="zoneGreenBand">--</div>
                <div class="plain-guide">Bezpečnostní rezerva nad 20 % pod váženou vnitřní hodnotou.</div>
            </div>
            <div class="zone-box yellow">
                <div class="zone-title">ZÓNA YELLOW (Tranše 30 %)</div>
                <div class="zone-val" id="zoneYellowBand">--</div>
                <div class="plain-guide">Atraktivní diskont 10 % až 20 % pod váženou vnitřní hodnotou.</div>
            </div>
            <div class="zone-box orange">
                <div class="zone-title">ZÓNA ORANGE (Tranše 15 %)</div>
                <div class="zone-val" id="zoneOrangeBand">--</div>
                <div class="plain-guide">Úvodní akumulační pásmo do 10 % pod vnitřní hodnotou.</div>
            </div>
            <div class="zone-box red">
                <div class="zone-title">ZÓNA RED (Tranše 0 %)</div>
                <div class="zone-val" id="zoneRedBand">--</div>
                <div class="plain-guide">Cena nad odhadovanou vnitřní hodnotou; vyčkat.</div>
            </div>
        </div>
        <p class="plain-guide" id="disagreementNote" style="margin-top: 14px;"></p>
    </div>

    <!-- 2. INTERAKTIVNÍ BURZOVNÍ GRAF -->
    <div class="card matrix-card">
        <div class="card-header">
            <span class="card-title">Interaktivní cenový graf &amp; historický vývoj</span>
            <span class="badge" id="chartChangeBadge">Vývoj období: --</span>
        </div>
        <div class="tv-toolbar">
            <div class="tv-tf-group" id="tvTimeframeGroup">
                <button type="button" class="tv-tf-btn" data-tf="1D" onclick="switchChartTimeframe('1D')">1D</button>
                <button type="button" class="tv-tf-btn" data-tf="1T" onclick="switchChartTimeframe('1T')">1T</button>
                <button type="button" class="tv-tf-btn" data-tf="1M" onclick="switchChartTimeframe('1M')">1M</button>
                <button type="button" class="tv-tf-btn" data-tf="6M" onclick="switchChartTimeframe('6M')">6M</button>
                <button type="button" class="tv-tf-btn" data-tf="YTD" onclick="switchChartTimeframe('YTD')">YTD</button>
                <button type="button" class="tv-tf-btn active" data-tf="1R" onclick="switchChartTimeframe('1R')">1R</button>
                <button type="button" class="tv-tf-btn" data-tf="5R" onclick="switchChartTimeframe('5R')">5R</button>
                <button type="button" class="tv-tf-btn" data-tf="MAX" onclick="switchChartTimeframe('MAX')">MAX</button>
            </div>
            <div class="tv-readout" id="tvCrosshairReadout">
                <span>Načítám OHLC přehled posledního obchodního dne...</span>
            </div>
        </div>
        <div id="tvChartContainer"></div>
    </div>

    <!-- 3. SYNTÉZA OCEŇOVACÍCH METOD -->
    <div class="card matrix-card">
        <div class="card-header">
            <span class="card-title">Syntéza ocenění — adaptivní spojení oceňovacích metod</span>
            <span class="badge" id="weightedFairBadge">Vážená vnitřní hodnota: --</span>
        </div>
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Oceňovací metoda</th>
                        <th>Váha</th>
                        <th>Vnitřní hodnota / akcii</th>
                        <th>Vážený příspěvek</th>
                        <th>Rozdíl vs. tržní cena</th>
                    </tr>
                </thead>
                <tbody id="synthesisTableBody"></tbody>
            </table>
        </div>
    </div>

    <!-- 4. INTERAKTIVNÍ GRAFY -->
    <div class="grid-3">
        <div class="card">
            <div class="card-header"><span class="card-title">Srovnání metod ocenění</span></div>
            <div class="chart-container"><canvas id="methodsChart"></canvas></div>
        </div>
        <div class="card">
            <div class="card-header"><span class="card-title">Rozdělení Monte Carlo (2 500 běhů)</span></div>
            <div class="chart-container"><canvas id="mcChart"></canvas></div>
        </div>
        <div class="card">
            <div class="card-header"><span class="card-title">Projekce toků &amp; dluhu (5 let)</span></div>
            <div class="chart-container"><canvas id="cashflowChart"></canvas></div>
        </div>
    </div>

    <!-- 5. FUNDAMENTÁLNÍ UKAZATELE, OWNER EARNINGS, VÝNOSOVÉ MODELY A SCÉNÁŘE -->
    <div class="grid-3">
        <div class="card">
            <div class="card-header"><span class="card-title">Fundamenty, Owner Earnings &amp; kvalita</span></div>
            <div class="metric-row"><span class="metric-label">Tržní kapitalizace</span><span class="metric-value" id="mcap">--</span></div>
            <div class="metric-row"><span class="metric-label">Hodnota podniku (EV)</span><span class="metric-value" id="ev">--</span></div>
            <div class="metric-row"><span class="metric-label">Owner Earnings (po udržov. CapEx)</span><span class="metric-value highlight" id="ownerEarnings">--</span></div>
            <div class="metric-row"><span class="metric-label">Udržovací / Růstový CapEx</span><span class="metric-value" id="capexSplit">--</span></div>
            <div class="metric-row"><span class="metric-label">Look-Through Earnings</span><span class="metric-value" id="lookThrough">--</span></div>
            <div class="metric-row"><span class="metric-label">Normalizovaný volný tok (FCF)</span><span class="metric-value" id="normFcf">--</span></div>
            <div class="metric-row"><span class="metric-label">ROIC vs. WACC</span><span class="metric-value" id="roicWacc">--</span></div>
            <div class="metric-row"><span class="metric-label">Piotroski / Altman</span><span class="metric-value" id="riskScores">--</span></div>
        </div>

        <div class="card">
            <div class="card-header"><span class="card-title">Výnosové modely &amp; tržní očekávání</span></div>
            <div class="metric-row"><span class="metric-label">DCF pro podnik / akcii</span><span class="metric-value highlight" id="dcfShare">--</span></div>
            <div class="metric-row"><span class="metric-label">Owner Earnings / Reziduální zisk</span><span class="metric-value" id="fcfeRi">--</span></div>
            <div class="metric-row"><span class="metric-label">Náklady kapitálu (WACC / Beta)</span><span class="metric-value" id="waccBeta">--</span></div>
            <div class="metric-row"><span class="metric-label">Podíl terminální hodnoty</span><span class="metric-value" id="tvShare">--</span></div>
            <div class="metric-row"><span class="metric-label">Implikovaný růst z ceny</span><span class="metric-value" id="impliedGrowth">--</span></div>
            <div class="metric-row"><span class="metric-label">Strukturovaný SOTP / akcii</span><span class="metric-value" id="sotpShare">--</span></div>
        </div>

        <div class="card">
            <div class="card-header"><span class="card-title">Scénáře &amp; pravděpodobnost</span></div>
            <div class="metric-row"><span class="metric-label">Pesimistický / Základní / Optimistický</span><span class="metric-value" id="scenariosTriple">--</span></div>
            <div class="metric-row"><span class="metric-label">Očekávaná hodnota scénářů</span><span class="metric-value" id="scenarioExpected">--</span></div>
            <div class="metric-row"><span class="metric-label">Monte Carlo Medián (P50)</span><span class="metric-value highlight" id="mcMedian">--</span></div>
            <div class="metric-row"><span class="metric-label">Interval P10 – P90</span><span class="metric-value" id="mc1090">--</span></div>
            <div class="metric-row"><span class="metric-label">Rozpětí P5 – P95</span><span class="metric-value" id="mc595">--</span></div>
            <div class="metric-row"><span class="metric-label">Pravděpodobnost nad trhem</span><span class="metric-value" id="mcProb">--</span></div>
        </div>
    </div>

    <!-- 6. KARTA EKONOMICKÉ KVALITY, FAKTOROVÉHO PROFILU A MEZÍ MODELU -->
    <div class="card matrix-card">
        <div class="card-header">
            <span class="card-title">Ekonomická kvalita, faktorový profil (BAB / QMJ) &amp; meze modelu</span>
            <span class="badge">Kvantitativní i kvalitativní diagnostika</span>
        </div>
        <div class="factor-pill-grid" id="factorProfileGrid"></div>
        <div class="plain-guide" id="modelLimitsList" style="margin-top: 12px;"></div>
    </div>

    <!-- 7. RELATIVNÍ NÁSOBKY A CITLIVOSTNÍ MATICE -->
    <div class="grid-2">
        <div class="card">
            <div class="card-header"><span class="card-title">Relativní ocenění (srovnatelné násobky)</span></div>
            <div class="table-responsive">
                <table>
                    <thead><tr><th>Metoda</th><th>Násobek</th><th>Metrika</th><th>Implikovaná cena</th></tr></thead>
                    <tbody id="peerTableBody"></tbody>
                </table>
            </div>
        </div>
        <div class="card">
            <div class="card-header"><span class="card-title">Citlivostní matice (WACC vs. terminální růst)</span></div>
            <div class="table-responsive">
                <table>
                    <thead><tr id="dcfSensHeader"></tr></thead>
                    <tbody id="dcfSensBody"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- 8. AI ANALYTICKÝ PORADCE -->
    <div class="card">
        <div class="card-header">
            <span class="card-title">Tensegrity AI analytický poradce</span>
            <div style="display:flex;gap:8px;align-items:center;">
                <span class="badge">Kontext Owner Earnings, SOTP &amp; kvalita dat</span>
                <button type="button" class="btn btn-news" style="padding:5px 10px;font-size:11px;" onclick="openAiKeyModal()">Nastavit Gemini klíč</button>
            </div>
        </div>
        <div class="advisor-thread" id="advisorMessages">
            <div class="advisor-message assistant">
                Dobrý den. Zeptejte se mě na cokoliv — na principy valuací (Owner Earnings, Look-Through, WACC, DCF — klidně i jednoduchým vysvětlením jako pro dítě), konkrétní společnosti, kvalitu jejich účetních výkazů nebo obecná témata.
            </div>
        </div>
        <form class="advisor-form" onsubmit="event.preventDefault(); askAdvisor();">
            <input type="text" class="advisor-input" id="advisorInput" placeholder="Zeptejte se například: Vysvětli WACC jako dítěti nebo Jak vychází tato firma?">
            <button type="submit" class="btn btn-primary" id="advisorSendBtn">Odeslat</button>
        </form>
    </div>
</div>

<script>
    let currentTicker = 'NVDA';
    let currentTimeframe = '1R';
    let lastFairValue = 0;
    let lastCurrency = 'USD';
    let lastBarSnapshot = null;
    let advisorHistory = [];
    let chartInstances = {};
    let tvChart = null;
    let tvCandleSeries = null;
    let tvVolumeSeries = null;
    let tvFairValueLine = null;

    function fmtNum(val, dec = 2) {
        if (val === null || val === undefined || isNaN(val)) return 'N/A';
        return Number(val).toLocaleString('cs-CZ', { minimumFractionDigits: dec, maximumFractionDigits: dec });
    }
    function fmtMoney(val, curr, dec = 2) {
        if (val === null || val === undefined || isNaN(val)) return 'N/A';
        return `${fmtNum(val, dec)} ${curr}`;
    }
    function fmtCompact(val, curr) {
        if (val === null || val === undefined || isNaN(val)) return 'N/A';
        const abs = Math.abs(val);
        if (abs >= 1e12) return `${fmtNum(val / 1e12, 2)} bil. ${curr}`;
        if (abs >= 1e9) return `${fmtNum(val / 1e9, 2)} mld. ${curr}`;
        if (abs >= 1e6) return `${fmtNum(val / 1e6, 1)} mil. ${curr}`;
        return `${fmtNum(val, 0)} ${curr}`;
    }
    function escapeHtml(text) {
        return String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function openNewsletterModal() {
        document.getElementById('newsletterStatus').innerText = '';
        document.getElementById('newsletterModal').classList.add('open');
    }
    function closeNewsletterModal() {
        document.getElementById('newsletterModal').classList.remove('open');
    }

    function openAiKeyModal() {
        document.getElementById('aiKeyStatus').innerText = '';
        document.getElementById('aiKeyModal').classList.add('open');
    }
    function closeAiKeyModal() {
        document.getElementById('aiKeyModal').classList.remove('open');
    }

    async function saveGeminiKey() {
        const apiKey = document.getElementById('geminiKeyInput').value.trim();
        const statusEl = document.getElementById('aiKeyStatus');
        if (!apiKey || apiKey.length < 15) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = 'Zadejte platný Gemini API klíč.';
            return;
        }
        statusEl.style.color = '#38bdf8';
        statusEl.innerText = 'Ukládám a aktivuji klíč...';
        try {
            const resp = await fetch('/api/settings/gemini-key', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ api_key: apiKey })
            });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.detail || 'Chyba uložení klíče.');
            statusEl.style.color = '#34d399';
            statusEl.innerText = data.message;
        } catch (err) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = err.message;
        }
    }

    async function subscribeNewsletter() {
        const email = document.getElementById('newsletterEmailInput').value.trim();
        const statusEl = document.getElementById('newsletterStatus');
        if (!email || !email.includes('@')) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = 'Zadejte platnou e-mailovou adresu.';
            return;
        }
        statusEl.style.color = '#38bdf8';
        statusEl.innerText = 'Ukládám odběr...';
        try {
            const resp = await fetch('/subscribe', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email })
            });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.detail || 'Chyba registrace.');
            statusEl.style.color = '#34d399';
            statusEl.innerText = data.message;
        } catch (err) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = err.message;
        }
    }

    async function sendNewsletterNow() {
        const email = document.getElementById('newsletterEmailInput').value.trim();
        const statusEl = document.getElementById('newsletterStatus');
        if (!email || !email.includes('@')) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = 'Zadejte platnou e-mailovou adresu příjemce.';
            return;
        }
        statusEl.style.color = '#38bdf8';
        statusEl.innerText = 'Překládám aktuální zprávy do češtiny, generuji PDF a odesílám e-mail...';
        try {
            const resp = await fetch('/api/newsletter/send-now', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email })
            });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.detail || 'Odeslání zpravodaje selhalo.');
            statusEl.style.color = '#34d399';
            statusEl.innerText = data.message;
        } catch (err) {
            statusEl.style.color = '#f87171';
            statusEl.innerText = err.message;
        }
    }

    async function loadForexBanner() {
        try {
            const resp = await fetch('/api/forex-rates');
            const data = await resp.json();
            const r = data.rates || {};
            document.getElementById('forexRatesBanner').innerHTML = `
                <span>EUR/CZK <strong>${fmtNum(r['EUR/CZK']?.rate, 2)}</strong></span>
                <span>USD/CZK <strong>${fmtNum(r['USD/CZK']?.rate, 2)}</strong></span>
                <span>GBP/CZK <strong>${fmtNum(r['GBP/CZK']?.rate, 2)}</strong></span>
            `;
        } catch (e) {}
    }

    function renderOhlcReadout(bar) {
        if (!bar) return;
        const readout = document.getElementById('tvCrosshairReadout');
        readout.innerHTML = `
            <span>O: <strong>${fmtNum(bar.open, 2)}</strong></span>
            <span>H: <strong>${fmtNum(bar.high, 2)}</strong></span>
            <span>L: <strong>${fmtNum(bar.low, 2)}</strong></span>
            <span>C: <strong>${fmtNum(bar.close, 2)} ${lastCurrency}</strong></span>
            <span style="color:#38bdf8;">Vnitřní hodnota: <strong>${fmtNum(lastFairValue, 2)} ${lastCurrency}</strong></span>
        `;
    }

    function initTvChartIfNeeded() {
        if (tvChart) return;
        const container = document.getElementById('tvChartContainer');
        tvChart = LightweightCharts.createChart(container, {
            layout: {
                background: { type: 'solid', color: '#070b14' },
                textColor: '#94a3b8',
                fontFamily: "'JetBrains Mono', monospace",
            },
            grid: {
                vertLines: { color: 'rgba(148, 163, 184, 0.06)' },
                horzLines: { color: 'rgba(148, 163, 184, 0.06)' },
            },
            crosshair: { mode: LightweightCharts.CrosshairMode.Normal },
            rightPriceScale: { borderColor: 'rgba(255, 255, 255, 0.1)' },
            timeScale: { borderColor: 'rgba(255, 255, 255, 0.1)', timeVisible: true, secondsVisible: false },
        });

        tvCandleSeries = tvChart.addCandlestickSeries({
            upColor: '#10b981',
            downColor: '#ef4444',
            borderVisible: false,
            wickUpColor: '#10b981',
            wickDownColor: '#ef4444',
        });

        tvVolumeSeries = tvChart.addHistogramSeries({
            priceFormat: { type: 'volume' },
            priceScaleId: '',
        });
        tvVolumeSeries.priceScale().applyOptions({ scaleMargins: { top: 0.82, bottom: 0 } });

        tvChart.subscribeCrosshairMove((param) => {
            if (!param || !param.time || !param.seriesData) {
                if (lastBarSnapshot) renderOhlcReadout(lastBarSnapshot);
                return;
            }
            const bar = param.seriesData.get(tvCandleSeries);
            if (!bar) {
                if (lastBarSnapshot) renderOhlcReadout(lastBarSnapshot);
                return;
            }
            renderOhlcReadout(bar);
        });

        window.addEventListener('resize', () => {
            if (tvChart && container) {
                tvChart.applyOptions({ width: container.clientWidth, height: container.clientHeight });
            }
        });
    }

    async function loadInteractiveChart(ticker, timeframe) {
        currentTimeframe = timeframe;
        document.querySelectorAll('.tv-tf-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tf === timeframe);
        });
        try {
            initTvChartIfNeeded();
            const resp = await fetch(`/api/chart/${encodeURIComponent(ticker)}?timeframe=${encodeURIComponent(timeframe)}`);
            const data = await resp.json();
            if (!resp.ok) return;

            const candles = data.candles || [];
            tvChart.applyOptions({ timeScale: { timeVisible: Boolean(data.intraday) } });
            tvCandleSeries.setData(candles);
            tvVolumeSeries.setData(data.volumes || []);

            if (candles.length > 0) {
                lastBarSnapshot = candles[candles.length - 1];
                renderOhlcReadout(lastBarSnapshot);
            }

            if (tvFairValueLine) {
                try { tvCandleSeries.removePriceLine(tvFairValueLine); } catch (e) {}
                tvFairValueLine = null;
            }
            if (lastFairValue > 0) {
                tvFairValueLine = tvCandleSeries.createPriceLine({
                    price: lastFairValue,
                    color: '#38bdf8',
                    lineWidth: 2,
                    lineStyle: LightweightCharts.LineStyle.Dashed,
                    axisLabelVisible: true,
                    title: 'Vnitřní hodnota',
                });
            }

            tvChart.timeScale().fitContent();
            const chg = data.change_pct || 0;
            const badge = document.getElementById('chartChangeBadge');
            badge.innerText = `Vývoj (${timeframe}): ${chg >= 0 ? '+' : ''}${fmtNum(chg, 2)} % (${fmtMoney(data.last_close, lastCurrency)})`;
            badge.style.color = chg >= 0 ? '#34d399' : '#f87171';
        } catch (err) {}
    }

    function switchChartTimeframe(tf) {
        loadInteractiveChart(currentTicker, tf);
    }

    function renderCompanyIdentity(d) {
        const curr = d.currency;
        const price = d.dcf_valuation.current_share_price;
        const fairVal = d.valuation_synthesis.weighted_fair_value;
        const upside = d.decision_engine.upside_to_fair_value_percent;

        document.getElementById('companyFullName').innerText = d.company_name || d.ticker;
        document.getElementById('companyTickerBadge').innerText = d.ticker;
        document.getElementById('companySectorMeta').innerText =
            `Sektor: ${d.sector} • Odvětví: ${d.industry} • Burza: ${d.exchange}`;

        document.getElementById('companyPriceBig').innerText = fmtMoney(price, curr);
        document.getElementById('companyMcapSub').innerText = `Tržní kap.: ${fmtCompact(d.trzni_kapitalizace, curr)}`;

        document.getElementById('companyFairBig').innerText = fmtMoney(fairVal, curr);
        const upsideEl = document.getElementById('companyUpsideSub');
        upsideEl.innerText = `Potenciál vůči trhu: ${upside >= 0 ? '+' : ''}${fmtNum(upside, 1)} %`;
        upsideEl.style.color = upside >= 0 ? '#34d399' : '#f87171';

        const logoContainer = document.getElementById('companyLogoContainer');
        const initials = (d.ticker || 'TT').replace(/[^A-Z0-9]/g, '').slice(0, 3) || 'TT';
        const candidates = Array.isArray(d.company_logo_urls) && d.company_logo_urls.length
            ? d.company_logo_urls
            : (d.company_logo_url ? [d.company_logo_url] : []);

        if (candidates.length > 0) {
            logoContainer.className = 'company-logo-frame';
            logoContainer.innerHTML = '';
            const img = document.createElement('img');
            img.alt = d.company_name || d.ticker;
            let idx = 0;
            img.onerror = function() {
                idx += 1;
                if (idx < candidates.length) {
                    img.src = candidates[idx];
                } else {
                    logoContainer.className = 'company-logo-fallback';
                    logoContainer.textContent = initials;
                }
            };
            img.src = candidates[0];
            logoContainer.appendChild(img);
        } else {
            logoContainer.className = 'company-logo-fallback';
            logoContainer.textContent = initials;
        }
    }

    function renderCharts(d) {
        Object.values(chartInstances).forEach(ch => { if (ch) ch.destroy(); });

        const curr = d.currency;
        const ctxM = document.getElementById('methodsChart').getContext('2d');
        const methodItems = [
            { label: 'Cena', val: d.dcf_valuation.current_share_price, color: '#64748b' },
            { label: 'Vážená', val: d.valuation_synthesis.weighted_fair_value, color: '#38bdf8' },
            { label: 'DCF', val: d.dcf_valuation.intrinsic_value_per_share, color: '#0284c7' },
            { label: 'Owner E.', val: d.fcfe_valuation.intrinsic_value_per_share, color: '#10b981' },
            { label: 'Násobky', val: d.peer_comps.blended_implied_price, color: '#8b5cf6' },
            { label: 'MC P50', val: d.monte_carlo.p50, color: '#f59e0b' },
        ].filter(x => x.val !== null && x.val !== undefined);

        chartInstances.methods = new Chart(ctxM, {
            type: 'bar',
            data: {
                labels: methodItems.map(x => x.label),
                datasets: [{
                    label: `Hodnota (${curr})`,
                    data: methodItems.map(x => x.val),
                    backgroundColor: methodItems.map(x => x.color),
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { display: false } },
                    y: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.06)' } }
                }
            }
        });

        const hist = d.monte_carlo.histogram || [];
        const ctxMc = document.getElementById('mcChart').getContext('2d');
        chartInstances.mc = new Chart(ctxMc, {
            type: 'bar',
            data: {
                labels: hist.map(b => b.label),
                datasets: [{
                    label: 'Počet simulací',
                    data: hist.map(b => b.count),
                    backgroundColor: hist.map(b => b.midpoint >= d.dcf_valuation.current_share_price ? 'rgba(16, 185, 129, 0.7)' : 'rgba(56, 189, 248, 0.45)'),
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { ticks: { color: '#94a3b8', font: { size: 10 }, maxRotation: 45 }, grid: { display: false } },
                    y: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.06)' } }
                }
            }
        });

        const cfRows = d.dcf_valuation.cash_flows || [];
        const lboRows = d.projektovane_roky || [];
        const ctxCf = document.getElementById('cashflowChart').getContext('2d');
        chartInstances.cf = new Chart(ctxCf, {
            type: 'bar',
            data: {
                labels: cfRows.map(r => `Rok ${r.rok}`),
                datasets: [
                    {
                        label: 'Volný peněžní tok (FCF)',
                        data: cfRows.map(r => r.volny_penezni_tok / 1e9),
                        backgroundColor: 'rgba(56, 189, 248, 0.75)',
                        borderRadius: 5
                    },
                    {
                        label: 'Zbývající dluh v LBO (mld.)',
                        data: lboRows.map(r => r.zbyvajici_cizi_kapital / 1e9),
                        backgroundColor: 'rgba(245, 158, 11, 0.75)',
                        borderRadius: 5
                    }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { labels: { color: '#94a3b8', font: { size: 11 } } } },
                scales: {
                    x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { display: false } },
                    y: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.06)' } }
                }
            }
        });
    }

    async function runValuation(useDynamicDefaults = false) {
        const rawQuery = document.getElementById('tickerInput').value.trim() || 'NVDA';
        document.getElementById('loader').style.display = 'flex';
        try {
            let url = `/api/valuation/${encodeURIComponent(rawQuery)}`;
            if (!useDynamicDefaults) {
                const lev = parseFloat(document.getElementById('leverageInput').value) / 100.0;
                const grw = parseFloat(document.getElementById('growthInput').value) / 100.0;
                const mult = parseFloat(document.getElementById('exitMultInput').value);
                url += `?leverage=${lev}&growth=${grw}&exit_mult=${mult}`;
            }

            const resp = await fetch(url);
            const d = await resp.json();
            if (!resp.ok) throw new Error(d.detail || 'Chyba výpočtu.');

            currentTicker = d.ticker;
            document.getElementById('tickerInput').value = d.ticker;

            if (d.dynamic_defaults && useDynamicDefaults) {
                document.getElementById('leverageInput').value = d.dynamic_defaults.leverage_percent;
                document.getElementById('growthInput').value = d.dynamic_defaults.growth_percent;
                document.getElementById('exitMultInput').value = d.dynamic_defaults.exit_multiple;
            }

            const curr = d.currency;
            const synth = d.valuation_synthesis;
            const dec = d.decision_engine;
            const dcf = d.dcf_valuation;
            const fund = d.fundamentals;
            const waccE = d.wacc_engine;
            const sc = d.scenarios;
            const mc = d.monte_carlo;
            const dq = d.data_quality;

            lastFairValue = synth.weighted_fair_value;
            lastCurrency = curr;

            renderCompanyIdentity(d);

            const srcBadge = document.getElementById('dataSourceBadge');
            srcBadge.innerText = `Výkazy: ${dq.primary_statement_source}`;
            if (dq.primary_statement_source.includes('SEC EDGAR') || dq.primary_statement_source.includes('Financial Modeling Prep')) {
                srcBadge.className = 'badge badge-verified';
            } else {
                srcBadge.className = 'badge badge-warn';
            }

            const warnBanner = document.getElementById('dataWarningBanner');
            if (dq.warnings && dq.warnings.length > 0) {
                warnBanner.style.display = 'block';
                warnBanner.innerHTML = `<strong>Kontrola kvality účetních dat:</strong> ${escapeHtml(dq.warnings.join(' '))}`;
            } else {
                warnBanner.style.display = 'none';
            }

            document.getElementById('auditStamp').innerText = `${d.audit_trail.model_version} | ${d.audit_trail.timestamp}`;
            document.getElementById('integrityBadge').innerText = `Index integrity: ${synth.integrity_score} / 100`;
            document.getElementById('weightedFairBadge').innerText = `Vážená vnitřní hodnota: ${fmtMoney(synth.weighted_fair_value, curr)}`;

            ['GREEN', 'YELLOW', 'ORANGE', 'RED'].forEach(z => {
                document.getElementById(`light-${z}`).classList.toggle('active', dec.zone === z);
            });
            document.getElementById('thesisTitle').innerText = `${dec.recommendation} (Doporučená tranše: ${dec.tranche_allocation_percent} %)`;
            document.getElementById('thesisSub').innerText =
                `Aktuální tržní cena: ${fmtMoney(dcf.current_share_price, curr)} | Vážená vnitřní hodnota: ${fmtMoney(synth.weighted_fair_value, curr)} | Bezpečnostní rezerva: ${fmtNum(dec.margin_of_safety_percent)} %.`;

            document.getElementById('zoneGreenBand').innerText = `0,00 – ${fmtMoney(dec.bands.green_max, curr)}`;
            document.getElementById('zoneYellowBand').innerText = `${fmtMoney(dec.bands.yellow_min, curr)} – ${fmtMoney(dec.bands.yellow_max, curr)}`;
            document.getElementById('zoneOrangeBand').innerText = `${fmtMoney(dec.bands.orange_min, curr)} – ${fmtMoney(dec.bands.orange_max, curr)}`;
            document.getElementById('zoneRedBand').innerText = `nad ${fmtMoney(dec.bands.red_min, curr)}`;
            document.getElementById('disagreementNote').innerText = synth.disagreement_note;

            document.getElementById('synthesisTableBody').innerHTML = synth.methods.map(m => `
                <tr>
                    <td>${escapeHtml(m.method)}</td>
                    <td>${fmtNum(m.weight_percent, 1)} %</td>
                    <td>${m.fair_value !== null ? fmtMoney(m.fair_value, curr) : '<span style="color:#fbbf24;">N/A (chybí data)</span>'}</td>
                    <td>${fmtMoney(m.contribution, curr)}</td>
                    <td style="color: ${m.vs_market_percent === null ? '#94a3b8' : (m.vs_market_percent >= 0 ? '#34d399' : '#f87171')};">${m.vs_market_percent !== null ? `${m.vs_market_percent >= 0 ? '+' : ''}${fmtNum(m.vs_market_percent)} %` : 'N/A'}</td>
                </tr>
            `).join('');

            document.getElementById('mcap').innerText = fmtCompact(d.trzni_kapitalizace, curr);
            document.getElementById('ev').innerText = fmtCompact(d.celkova_hodnota_podniku, curr);
            document.getElementById('ownerEarnings').innerText = fund.owner_earnings !== null
                ? `${fmtCompact(fund.owner_earnings, curr)} (${fmtMoney(fund.owner_earnings_per_share, curr)}/akcie)`
                : 'N/A';
            document.getElementById('capexSplit').innerText = (fund.maintenance_capex !== null && fund.growth_capex !== null)
                ? `${fmtCompact(fund.maintenance_capex, curr)} / ${fmtCompact(fund.growth_capex, curr)}`
                : 'N/A';
            document.getElementById('lookThrough').innerText = fmtCompact(fund.look_through_earnings, curr);
            document.getElementById('normFcf').innerText = fmtCompact(fund.normalized_fcf, curr);
            document.getElementById('roicWacc').innerText = fund.roic_percent !== null
                ? `${fmtNum(fund.roic_percent)} % vs. ${fmtNum(waccE.wacc_percent)} %`
                : `N/A vs. ${fmtNum(waccE.wacc_percent)} %`;
            document.getElementById('riskScores').innerText = `${d.risk_analysis.piotroski_f_score}/9 • ${d.risk_analysis.altman_status}`;

            document.getElementById('dcfShare').innerText = fmtMoney(dcf.intrinsic_value_per_share, curr);
            document.getElementById('fcfeRi').innerText = `${fmtMoney(d.fcfe_valuation.intrinsic_value_per_share, curr)} / ${fmtMoney(d.residual_income.intrinsic_value_per_share, curr)}`;
            document.getElementById('waccBeta').innerText = `${fmtNum(waccE.wacc_percent)} % (${fmtNum(waccE.beta)})`;
            document.getElementById('tvShare').innerText = dcf.tv_share_percent !== null ? `${fmtNum(dcf.tv_share_percent, 1)} % z EV` : 'N/A';
            document.getElementById('impliedGrowth').innerText = d.reverse_dcf.implied_fcf_growth_percent !== null ? `${fmtNum(d.reverse_dcf.implied_fcf_growth_percent)} % ročně` : 'N/A';
            document.getElementById('sotpShare').innerText = fmtMoney(d.metoda_souctu_casti.sotp_per_share, curr);

            document.getElementById('scenariosTriple').innerText = sc.available
                ? `${fmtNum(sc.bear.fair_value, 1)} / ${fmtNum(sc.base.fair_value, 1)} / ${fmtNum(sc.bull.fair_value, 1)} ${curr}`
                : 'N/A';
            document.getElementById('scenarioExpected').innerText = fmtMoney(sc.expected_value, curr);
            document.getElementById('mcMedian').innerText = fmtMoney(mc.p50, curr);
            document.getElementById('mc1090').innerText = mc.p10 !== null ? `${fmtMoney(mc.p10, curr)} – ${fmtMoney(mc.p90, curr)}` : 'N/A';
            document.getElementById('mc595').innerText = mc.p5 !== null ? `${fmtMoney(mc.p5, curr)} – ${fmtMoney(mc.p95, curr)}` : 'N/A';
            document.getElementById('mcProb').innerText = mc.prob_above_market_percent !== null ? `${fmtNum(mc.prob_above_market_percent, 1)} %` : 'N/A';

            const fp = d.factor_and_quality_diagnostics || {};
            const factorKeys = ['low_beta_bab', 'quality_qmj', 'owner_earnings_conversion', 'cheap_float_financing', 'drawdown_and_size'];
            document.getElementById('factorProfileGrid').innerHTML = factorKeys.map(k => {
                const item = fp[k];
                if (!item) return '';
                return `
                    <div class="factor-item">
                        <div class="factor-item-head">
                            <span>${escapeHtml(item.label)}</span>
                            <span class="${item.passed ? 'factor-status-ok' : 'factor-status-warn'}">${item.passed ? 'SPLNĚNO' : 'UPOZORNĚNÍ'}</span>
                        </div>
                        <div class="factor-item-desc">${escapeHtml(item.detail)}</div>
                    </div>
                `;
            }).join('');

            const limits = fp.model_limits || [];
            document.getElementById('modelLimitsList').innerHTML =
                `<strong>Metodické limity a rizika odhadu:</strong> ` + limits.map(l => escapeHtml(l)).join(' &bull; ');

            document.getElementById('peerTableBody').innerHTML = (d.peer_comps.multiples || []).map(p => `
                <tr>
                    <td>${escapeHtml(p.method)}</td>
                    <td>${fmtNum(p.multiple, 1)}x</td>
                    <td>${fmtCompact(p.forward_metric, curr)}</td>
                    <td>${fmtMoney(p.implied_share_price, curr)}</td>
                </tr>
            `).join('');

            const sens = dcf.sensitivity_matrix || [];
            if (sens.length > 0) {
                document.getElementById('dcfSensHeader').innerHTML = '<th>WACC \\ Růst</th>' + sens[0].cells.map(c => `<th>${fmtNum(c.tg, 2)} %</th>`).join('');
                document.getElementById('dcfSensBody').innerHTML = sens.map(row => `
                    <tr>
                        <td><strong>${fmtNum(row.wacc, 2)} %</strong></td>
                        ${row.cells.map(c => `<td style="color:${c.price >= dcf.current_share_price ? '#34d399' : '#f87171'};">${fmtNum(c.price, 2)}</td>`).join('')}
                    </tr>
                `).join('');
            } else {
                document.getElementById('dcfSensHeader').innerHTML = '<th>Citlivostní matice DCF</th>';
                document.getElementById('dcfSensBody').innerHTML = '<tr><td>N/A — pro výpočet matice chybí reportovaný FCF nebo dluh.</td></tr>';
            }

            renderCharts(d);
            await loadInteractiveChart(d.ticker, currentTimeframe);
        } catch (err) {
            const warnBanner = document.getElementById('dataWarningBanner');
            warnBanner.style.display = 'block';
            warnBanner.innerHTML = `<strong>Chyba při zpracování titulu:</strong> ${escapeHtml(err.message)}`;
        } finally {
            document.getElementById('loader').style.display = 'none';
        }
    }

    function downloadExcel() {
        const lev = parseFloat(document.getElementById('leverageInput').value) / 100.0;
        const grw = parseFloat(document.getElementById('growthInput').value) / 100.0;
        const mult = parseFloat(document.getElementById('exitMultInput').value);
        window.location.href = `/api/export/excel/${encodeURIComponent(currentTicker)}?leverage=${lev}&growth=${grw}&exit_mult=${mult}`;
    }

    function downloadPdf() {
        const lev = parseFloat(document.getElementById('leverageInput').value) / 100.0;
        const grw = parseFloat(document.getElementById('growthInput').value) / 100.0;
        const mult = parseFloat(document.getElementById('exitMultInput').value);
        window.location.href = `/api/export/pdf/${encodeURIComponent(currentTicker)}?leverage=${lev}&growth=${grw}&exit_mult=${mult}`;
    }

    function formatAdvisorReply(rawText) {
        let cleaned = String(rawText || '')
            .replace(/\\\[|\\\]|\\\(|\\\)/g, '')
            .replace(/\$\$/g, '')
            .replace(/\\text\{([^{}]*)\}/g, '$1')
            .replace(/\\frac\{([^{}]+)\}\{([^{}]+)\}/g, '($1 / $2)')
            .replace(/\\times/g, '×')
            .replace(/\\beta/g, 'Beta');
        let htmlOut = escapeHtml(cleaned);
        htmlOut = htmlOut.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
        htmlOut = htmlOut.replace(/`([^`]+)`/g, '<code style="background:rgba(56,189,248,0.14);padding:2px 5px;border-radius:4px;font-family:\'JetBrains Mono\',monospace;">$1</code>');
        return htmlOut.replace(/\n/g, '<br>');
    }

    async function askAdvisor() {
        const inputEl = document.getElementById('advisorInput');
        const btnEl = document.getElementById('advisorSendBtn');
        const msgBox = document.getElementById('advisorMessages');
        const question = inputEl.value.trim();
        if (!question) return;

        inputEl.value = '';
        btnEl.disabled = true;

        const userDiv = document.createElement('div');
        userDiv.className = 'advisor-message user';
        userDiv.innerText = question;
        msgBox.appendChild(userDiv);

        const aiDiv = document.createElement('div');
        aiDiv.className = 'advisor-message assistant';
        aiDiv.innerHTML = '<span style="color:#38bdf8;">Zpracovávám odpověď...</span>';
        msgBox.appendChild(aiDiv);
        msgBox.scrollTop = msgBox.scrollHeight;

        try {
            const resp = await fetch(`/api/advisor/${encodeURIComponent(currentTicker)}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    question,
                    history: advisorHistory,
                    scenario: {
                        leverage: parseFloat(document.getElementById('leverageInput').value) / 100.0,
                        growth: parseFloat(document.getElementById('growthInput').value) / 100.0,
                        exit_mult: parseFloat(document.getElementById('exitMultInput').value)
                    }
                })
            });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.detail || 'Chyba AI poradce.');

            aiDiv.innerHTML = formatAdvisorReply(data.answer);
            advisorHistory.push({ role: 'user', text: question });
            advisorHistory.push({ role: 'assistant', text: data.answer });
        } catch (err) {
            aiDiv.innerHTML = `<span style="color:#f87171;">${escapeHtml(err.message)}</span>`;
        } finally {
            btnEl.disabled = false;
            msgBox.scrollTop = msgBox.scrollHeight;
        }
    }

    window.addEventListener('DOMContentLoaded', () => {
        loadForexBanner();
        runValuation(true);
    });
</script>
</body>
</html>"""


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))