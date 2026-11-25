# -*- coding: utf-8 -*-
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, parseaddr
import asyncio
import base64
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
os.makedirs(STATIC_DIR, exist_ok=True)
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
        if value is None:
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
# 2. VYHLEDÁVÁNÍ SPOLEČNOSTÍ, MĚNOVÝ PŘEVODNÍK VÝKAZŮ A LOGO PROFILU
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
    "spacex": "SPCX",
}


def resolve_company_ticker(query: str) -> str:
    raw = (query or "").strip()
    if not raw:
        return "NVDA"

    lower_q = raw.lower()
    if lower_q in KNOWN_COMPANY_ALIASES:
        return KNOWN_COMPANY_ALIASES[lower_q]

    if re.fullmatch(r"[A-Za-z0-9\.\-\^=]{1,14}", raw) and " " not in raw:
        upper_candidate = raw.upper()
        try:
            test_info = yf.Ticker(upper_candidate).info or {}
            if (
                test_info.get("regularMarketPrice")
                or test_info.get("currentPrice")
                or test_info.get("shortName")
                or test_info.get("longName")
            ):
                return upper_candidate
        except Exception:
            pass

    try:
        search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={quote(raw)}&quotesCount=6&newsCount=0"
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
        for q in quotes:
            if q.get("quoteType") in ("EQUITY", "ETF") and q.get("symbol"):
                return str(q["symbol"]).upper()
        if quotes and quotes[0].get("symbol"):
            return str(quotes[0]["symbol"]).upper()
    except Exception as err:
        logger.warning("Vyhledání názvu společnosti '%s' selhalo: %s", raw, err)

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
    "SPCX": "spacex.com",
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
# 3. ŽIVÁ MAKROEKONOMICKÁ DATA, MĚNOVÉ KURZY A ČNB
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
        tnx_hist = yf.Ticker("^TNX").history(period="5d", interval="1d")
        if not tnx_hist.empty:
            last_tnx = safe_float(tnx_hist["Close"].dropna().iloc[-1], 0.0)
            if 0.5 < last_tnx < 15.0:
                us10y = round(last_tnx, 2)
    except Exception:
        pass

    if us10y is None:
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
# 4. SPRÁVA ODBĚRATELŮ, PŘEKLAD ZPRÁV DO ČEŠTINY A LEHKÝ E-MAILOVÝ NEWSLETTER
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
        r"\b(the|and|with|from|for|that|this|are|was|were|have|has|will|would|could|should|market|stocks|shares|fed|rate|rates|inflation|economy|quarter|earnings|guidance|investors|trading|growth|debt|says|after|over|new|us|china|europe|bank|price|prices|oil|tech)\b",
        text.lower(),
    )
    cz_chars = re.findall(r"[ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]", text)
    return len(en_markers) >= 1 and len(cz_chars) == 0


def _fallback_translate_en_to_cs(text: str) -> str:
    clean = re.sub(r"\s+", " ", (text or "")).strip()
    if not clean or not _looks_english(clean):
        return clean
    chunks = [clean[i:i + 450] for i in range(0, min(len(clean), 1350), 450)]
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


def _clean_czech_news_summary(text: str, title_cs: str) -> str:
    cleaned = re.sub(r"\[?\s*Zdroj\s*:[^\]]*\]?", "", text or "", flags=re.IGNORECASE)
    cleaned = re.sub(r"https?://\S+", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    if _looks_english(cleaned):
        cleaned = _fallback_translate_en_to_cs(cleaned)

    raw_sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", cleaned) if len(s.strip()) > 8]
    sentences: list[str] = []
    for s in raw_sentences:
        if _looks_english(s):
            s = _fallback_translate_en_to_cs(s)
        s_clean = s.rstrip(".!?").strip() + "."
        if s_clean not in sentences and not _looks_english(s_clean):
            sentences.append(s_clean)
        # Každé shrnutí má mít vždy přesně 3 věty.
        if len(sentences) >= 3:
            break

    base_t = title_cs.rstrip(".!?").strip()
    fallback_pool = [
        f"Téma se zaměřuje na klíčový vývoj v oblasti {base_t.lower()}.",
        "Událost má přímý dopad na tržní sentiment, regulaci nebo hospodaření dotčených subjektů.",
        "Další vývoj bude sledovat trh, investoři i institucionální analytici.",
    ]
    idx_fb = 0
    while len(sentences) < 3 and idx_fb < len(fallback_pool):
        cand = fallback_pool[idx_fb]
        idx_fb += 1
        if cand not in sentences:
            sentences.append(cand)

    return " ".join(sentences[:3])


def _translate_and_summarize_batch_czech(articles: list[dict]) -> list[dict]:
    if not articles:
        return articles

    if GEMINI_API_KEY:
        try:
            items_payload = [
                {
                    "index": idx,
                    "title": art["raw_title"],
                    "snippet": art["raw_text"][:1600],
                }
                for idx, art in enumerate(articles)
            ]
            prompt = (
                "Jsi špičkový editor a překladatel českého zpravodaje Tensegrity Terminal zaměřeného na "
                "finance, akvizice, politiku, investice, občanský a obchodní zákoník, zákony, legislativu a soudy.\n"
                "Pro KAŽDOU položku ze seznamu splň přísně tato pravidla:\n"
                "1. Přelož nadpis ('title_cs') do naprosto plynulé, spisovné a přirozené ČEŠTINY (nikdy nenechávej angličtinu).\n"
                "2. Vytvoř věcné, hutné a srozumitelné shrnutí nejdůležitějšího ('summary_cs') v ČEŠTINĚ přesně na 3 věty.\n"
                "3. Do textu shrnutí NIKDY nepiš název média, zdroj článku ani citace typu 'říká', 'napsal' bez kontextu – piš věcně jako analytické shrnutí.\n"
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
                            <table role="presentation" width="100%" cellspacing="6" cellpadding="0" style="margin-bottom: 8px;">
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
                            <table role="presentation" width="100%" cellspacing="6" cellpadding="0" style="margin-bottom: 20px;">
                                <tr>
                                    <td width="50%" style="background-color: #0f1625; border: 1px solid #78350f; border-radius: 8px; padding: 12px; text-align: center;">
                                        <div style="font-size: 10px; color: #f59e0b;">INFLACE V ČR</div>
                                        <div style="font-size: 16px; font-weight: 800; color: #ffffff; font-family: Consolas, monospace; margin-top: 4px;">{cnb['latest_inflation']}</div>
                                    </td>
                                    <td width="50%" style="background-color: #0f1625; border: 1px solid #78350f; border-radius: 8px; padding: 12px; text-align: center;">
                                        <div style="font-size: 10px; color: #f59e0b;">REPO SAZBA ČNB</div>
                                        <div style="font-size: 16px; font-weight: 800; color: #ffffff; font-family: Consolas, monospace; margin-top: 4px;">{cnb['repo_rate']}</div>
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


app = FastAPI(title="Tensegrity Terminal | Institutional Valuation Platform", version="4.1", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ==============================================================================
# 5. NEPRŮSTŘELNÝ ADAPTIVNÍ VALUAČNÍ ENGINE
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
    fin_df,
    cf_df,
    sector: str,
    stmt_fx: float = 1.0,
) -> tuple[float, float, float]:
    bench = ALL_SECTOR_BENCHMARKS.get(sector, {"ev_ebitda": 13.0, "pe": 19.0, "ev_revenue": 3.0, "pb": 3.0, "growth": 0.075})
    current_price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"), 0.0) or 0.0
    shares = safe_float(info.get("sharesOutstanding"), 0.0) or 0.0
    market_cap = safe_float(info.get("marketCap"), current_price * shares) or 0.0
    total_debt = (safe_float(info.get("totalDebt"), 0.0) or 0.0) * stmt_fx
    currency = str(info.get("currency") or "USD").upper()

    total_cap = market_cap + total_debt
    if total_cap > 0 and total_debt >= 0:
        auto_leverage = round(max(0.0, min(0.85, total_debt / total_cap)), 2)
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

    rev_hist = extract_statement_series(fin_df, ["Total Revenue", "Operating Revenue"], max_years=4, fx_mult=stmt_fx)
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

    usd_rate = get_fx_conversion_rate(currency, "USD")
    mcap_usd = market_cap * usd_rate

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
    try:
        stock = yf.Ticker(clean_ticker)
        info = stock.info or {}
        fin_df = stock.financials
        bs_df = stock.balance_sheet
        cf_df = stock.cashflow
    except Exception as err:
        logger.warning("Chyba při stahování dat z Yahoo Finance pro %s: %s", clean_ticker, err)
        info = {}
        fin_df = None
        bs_df = None
        cf_df = None

    quote_type = str(info.get("quoteType") or "EQUITY").upper()
    if quote_type in ("INDEX", "MUTUALFUND", "CRYPTOCURRENCY", "CURRENCY", "FUTURE"):
        raise ValueError(f"Instrument '{clean_ticker}' je typu {quote_type} (nikoli veřejně obchodovaná společnost s účetními výkazy).")

    raw_currency = str(info.get("currency") or "USD")
    currency = "GBP" if raw_currency == "GBp" else raw_currency.upper()
    fin_currency = str(info.get("financialCurrency") or currency).upper()

    stmt_fx = get_fx_conversion_rate(fin_currency, currency)
    price_scale = 0.01 if raw_currency == "GBp" else 1.0

    sector = str(info.get("sector") or "Technology")
    industry = str(info.get("industry") or "General")
    is_financial = sector == "Financial Services" or any(
        k in industry.lower() for k in ("bank", "insurance", "financial", "asset management", "capital markets")
    )
    is_real_estate = sector == "Real Estate" or "reit" in industry.lower()

    auto_lev, auto_grw, auto_mult = compute_dynamic_scenario_defaults(info, fin_df, cf_df, sector, stmt_fx)
    eff_leverage = leverage if leverage is not None else auto_lev
    eff_growth = growth if growth is not None else auto_grw
    eff_exit_mult = exit_mult if exit_mult is not None else auto_mult

    benchmarks = get_live_market_benchmarks()
    rf_rate = get_currency_risk_free_rate(currency, benchmarks["us10y_yield"])
    erp_rate = 0.052 if currency == "CZK" else 0.046

    exchange = str(info.get("exchange") or "NMS")
    company_name = str(info.get("longName") or info.get("shortName") or clean_ticker)
    company_logo_urls = extract_company_logo_urls(info, clean_ticker)
    company_logo_url = company_logo_urls[0] if company_logo_urls else ""
    country = str(info.get("country") or "United States")
    employees = int(safe_float(info.get("fullTimeEmployees"), 0) or 0)
    business_summary = str(info.get("longBusinessSummary") or "")

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

    raw_shares = safe_float(info.get("sharesOutstanding"), None)
    raw_mcap = safe_float(info.get("marketCap"), None)

    if raw_shares and raw_shares > 0 and raw_mcap and raw_mcap > 0:
        shares_outstanding = raw_shares
        market_cap = raw_mcap
    elif raw_mcap and raw_mcap > 0:
        market_cap = raw_mcap
        shares_outstanding = market_cap / current_price
    elif raw_shares and raw_shares > 0:
        shares_outstanding = raw_shares
        market_cap = current_price * shares_outstanding
    else:
        shares_outstanding = 1_000_000_000.0
        market_cap = current_price * shares_outstanding

    rev_hist = extract_statement_series(fin_df, ["Total Revenue", "Operating Revenue"], fx_mult=stmt_fx)
    ebitda_hist = extract_statement_series(fin_df, ["EBITDA", "Normalized EBITDA"], fx_mult=stmt_fx)
    ebit_hist = extract_statement_series(fin_df, ["EBIT", "Operating Income"], fx_mult=stmt_fx)
    da_hist = extract_statement_series(cf_df, ["Depreciation And Amortization", "Depreciation Amortization Depletion"], fx_mult=stmt_fx)
    net_inc_hist = extract_statement_series(fin_df, ["Net Income", "Net Income Common Stockholders"], fx_mult=stmt_fx)
    pretax_hist = extract_statement_series(fin_df, ["Pretax Income"], fx_mult=stmt_fx)
    tax_hist = extract_statement_series(fin_df, ["Tax Provision"], fx_mult=stmt_fx)
    int_exp_hist = extract_statement_series(fin_df, ["Interest Expense", "Interest Expense Non Operating"], fx_mult=stmt_fx)
    ocf_hist = extract_statement_series(cf_df, ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"], fx_mult=stmt_fx)
    capex_hist = extract_statement_series(cf_df, ["Capital Expenditure", "Purchase Of PPE"], fx_mult=stmt_fx)
    sbc_hist = extract_statement_series(cf_df, ["Stock Based Compensation"], fx_mult=stmt_fx)
    equity_hist = extract_statement_series(bs_df, ["Stockholders Equity", "Total Equity Gross Minority Interest"], fx_mult=stmt_fx)
    debt_hist = extract_statement_series(bs_df, ["Total Debt", "Long Term Debt"], fx_mult=stmt_fx)
    cash_hist = extract_statement_series(bs_df, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"], fx_mult=stmt_fx)
    cur_assets_hist = extract_statement_series(bs_df, ["Current Assets"], fx_mult=stmt_fx)
    cur_liab_hist = extract_statement_series(bs_df, ["Current Liabilities"], fx_mult=stmt_fx)
    total_assets_hist = extract_statement_series(bs_df, ["Total Assets"], fx_mult=stmt_fx)
    retained_earn_hist = extract_statement_series(bs_df, ["Retained Earnings"], fx_mult=stmt_fx)

    raw_debt_info = safe_float(info.get("totalDebt"), None)
    total_debt = (raw_debt_info * stmt_fx) if raw_debt_info is not None else (debt_hist[0] if debt_hist else 0.0)
    raw_cash_info = safe_float(info.get("totalCash"), None)
    total_cash = (raw_cash_info * stmt_fx) if raw_cash_info is not None else (cash_hist[0] if cash_hist else 0.0)
    raw_min_info = safe_float(info.get("minorityInterest"), 0.0) or 0.0
    minority_interest = raw_min_info * stmt_fx
    enterprise_value = market_cap + total_debt + minority_interest - total_cash

    raw_rev_info = safe_float(info.get("totalRevenue"), None)
    revenue = (raw_rev_info * stmt_fx) if raw_rev_info is not None else (rev_hist[0] if rev_hist else market_cap * 0.25)

    raw_ni_info = safe_float(info.get("netIncomeToCommon"), None)
    net_income = (raw_ni_info * stmt_fx) if raw_ni_info is not None else (net_inc_hist[0] if net_inc_hist else revenue * 0.12)

    da_val = abs(da_hist[0]) if da_hist else revenue * 0.04
    raw_ebitda_info = safe_float(info.get("ebitda"), None)
    reported_ebitda = (raw_ebitda_info * stmt_fx) if raw_ebitda_info is not None else (ebitda_hist[0] if ebitda_hist else None)
    if reported_ebitda is None:
        if ebit_hist:
            reported_ebitda = ebit_hist[0] + da_val
        else:
            reported_ebitda = max(net_income * 1.35 + da_val, revenue * 0.18)

    reported_ebit = ebit_hist[0] if ebit_hist else max(net_income * 1.25, reported_ebitda - da_val)
    book_val_per_share = safe_float(info.get("bookValue"), 0.0)
    if book_val_per_share and book_val_per_share > 0:
        book_equity = book_val_per_share * stmt_fx * shares_outstanding
    elif equity_hist and equity_hist[0] > 0:
        book_equity = equity_hist[0]
    else:
        book_equity = max(market_cap * 0.22, 1.0)

    # Automatická kontrola kvality vstupních dat (data sanity check) pro jednotný chod napříč firmami
    data_warnings: list[str] = []
    if raw_mcap and raw_mcap > 0 and raw_shares and raw_shares > 0:
        implied_mcap = current_price * raw_shares
        mcap_ratio = max(raw_mcap, implied_mcap) / max(1.0, min(raw_mcap, implied_mcap))
        if mcap_ratio > 1.35:
            mcap_rounded = round(raw_mcap, 0)
            implied_rounded = round(implied_mcap, 0)
            data_warnings.append(
                f"Rozdíl mezi tržní kapitalizací ({mcap_rounded}) a součinem ceny × počet akcií ({implied_rounded})."
            )

    sbc_expense = abs(sbc_hist[0]) if sbc_hist else max(0.0, revenue * 0.02)

    raw_ocf_info = safe_float(info.get("operatingCashflow"), None)
    raw_ocf = ocf_hist[0] if ocf_hist else ((raw_ocf_info * stmt_fx) if raw_ocf_info is not None else reported_ebitda * 0.78)
    raw_capex = abs(capex_hist[0]) if capex_hist else (da_val * 1.1)
    raw_fcf_info = safe_float(info.get("freeCashflow"), None)
    reported_fcf = (raw_fcf_info * stmt_fx) if raw_fcf_info is not None else (raw_ocf - raw_capex)

    # Jednotná normalizace volného peněžního toku (FCF) pro všechny společnosti bez ad-hoc úprav podle trhu
    is_capex_heavy_growth = False
    if is_real_estate:
        reported_fcf = max(net_income + da_val * 0.85, raw_ocf * 0.85)
    elif reported_fcf <= 0 and reported_ebitda > 0:
        is_capex_heavy_growth = True
        reported_fcf = max(reported_ebitda * 0.50, revenue * 0.06)

    normalized_ebitda = max(reported_ebitda * 0.5, reported_ebitda - (sbc_expense * 0.5))
    sbc_fcf_factor = 0.25 if (raw_ocf > 0 and reported_fcf > raw_ocf * 0.45) else 0.50
    normalized_fcf = max(reported_fcf * 0.65, reported_fcf - (sbc_expense * sbc_fcf_factor))

    if pretax_hist and tax_hist and pretax_hist[0] > 0 and 0 <= tax_hist[0] <= pretax_hist[0] * 0.5:
        effective_tax_rate = max(0.05, min(0.35, tax_hist[0] / pretax_hist[0]))
    else:
        effective_tax_rate = 0.21

    raw_beta = safe_float(info.get("beta"), 1.15) or 1.15
    raw_beta = max(0.50, min(2.40, raw_beta))
    beta = (2.0 / 3.0) * raw_beta + (1.0 / 3.0) * 1.0
    cost_of_equity = rf_rate + beta * erp_rate

    interest_expense = abs(int_exp_hist[0]) if int_exp_hist else total_debt * 0.048
    pre_tax_cost_of_debt = (interest_expense / total_debt) if total_debt > 0 else (rf_rate + 0.012)
    pre_tax_cost_of_debt = max(rf_rate + 0.004, min(0.11, pre_tax_cost_of_debt))

    total_capital = market_cap + total_debt
    weight_equity = (market_cap / total_capital) if total_capital > 0 else 0.90
    weight_debt = (total_debt / total_capital) if total_capital > 0 else 0.10
    wacc = (weight_equity * cost_of_equity) + (weight_debt * pre_tax_cost_of_debt * (1.0 - effective_tax_rate))
    wacc = max(0.068, min(0.145, wacc))

    invested_capital = max(1.0, book_equity + total_debt - total_cash)
    nopat = reported_ebit * (1.0 - effective_tax_rate)
    roic = nopat / invested_capital
    economic_profit = (roic - wacc) * invested_capital

    terminal_growth = min(0.03, max(0.02, wacc - 0.055))
    if terminal_growth >= wacc:
        terminal_growth = wacc * 0.45

    def simulate_fcff_dcf(initial_g: float, custom_wacc: float, custom_exit_mult: float, custom_tg: float) -> dict:
        cf_rows = []
        curr_fcf = normalized_fcf
        curr_ebitda = normalized_ebitda
        pv_sum = 0.0
        for yr in range(1, 6):
            fade_weight = (yr - 1) / 4.0
            yr_growth = initial_g * (1.0 - 0.35 * fade_weight) + custom_tg * (0.35 * fade_weight)
            curr_fcf *= (1.0 + yr_growth)
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
        multiple_tv = curr_ebitda * custom_exit_mult
        blended_tv = 0.5 * gordon_tv + 0.5 * multiple_tv
        pv_tv = blended_tv / ((1.0 + custom_wacc) ** 5)

        ev_val = pv_sum + pv_tv
        eq_val = max(0.0, ev_val + total_cash - total_debt - minority_interest)
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
    dcf_intrinsic_share = dcf_res["intrinsic_value_per_share"]
    dcf_upside = ((dcf_intrinsic_share / current_price) - 1.0) * 100.0 if current_price > 0 else 0.0

    bench = ALL_SECTOR_BENCHMARKS.get(sector, {"ev_ebitda": 13.0, "pe": 19.0, "ev_revenue": 3.0, "pb": 3.0, "growth": 0.075})
    if is_financial:
        base_fcfe = max(1.0, net_income * 0.85)
    else:
        # Použijeme vyšší ze dvou konzistentních základů: očištěný FCFF po úrocích nebo hotovostní zisk po úrocích
        cash_fcfe_base = normalized_fcf - interest_expense * (1.0 - effective_tax_rate)
        ni_fcfe_base = max(1.0, (net_income - sbc_expense * 0.25)) * 0.90
        base_fcfe = max(1.0, 0.65 * max(cash_fcfe_base, ni_fcfe_base * 0.75) + 0.35 * ni_fcfe_base)

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

    low_g, high_g = -0.35, 1.50
    for _ in range(60):
        mid_g = (low_g + high_g) / 2.0
        test_price = simulate_fcff_dcf(mid_g, wacc, eff_exit_mult, terminal_growth)["intrinsic_value_per_share"]
        if test_price < current_price:
            low_g = mid_g
        else:
            high_g = mid_g
    implied_growth_pct = ((low_g + high_g) / 2.0) * 100.0

    roe_est = min(0.55, max(0.04, net_income / max(1.0, book_equity)))
    pv_ri = 0.0
    curr_bv = book_equity
    faded_roe = roe_est
    for yr in range(1, 6):
        faded_roe = roe_est * (1.0 - 0.12 * (yr - 1)) + cost_of_equity * (0.12 * (yr - 1))
        ri_yr = (faded_roe - cost_of_equity) * curr_bv
        pv_ri += ri_yr / ((1.0 + cost_of_equity) ** yr)
        curr_bv += faded_roe * curr_bv * 0.45

    terminal_ri = ((faded_roe - cost_of_equity) * curr_bv) / max(0.02, cost_of_equity - terminal_growth)
    ri_equity_val = max(0.0, book_equity + pv_ri + (terminal_ri / ((1.0 + cost_of_equity) ** 5)))
    ri_per_share = ri_equity_val / shares_outstanding

    fwd_factor = (1.0 + eff_growth) ** 2
    pv_discount_2y_eq = 1.0 / ((1.0 + cost_of_equity) ** 2)
    pv_discount_2y_ev = 1.0 / ((1.0 + wacc) ** 2)

    company_roic_or_roe = roe_est if is_financial else roic
    quality_mult_adj = max(0.85, min(1.12, 1.0 + (company_roic_or_roe - wacc) * 0.35))

    if is_financial:
        roe_adjusted_pb = max(0.65, min(3.2, bench["pb"] * (roe_est / 0.12)))
        adj_pe = round((0.65 * bench["pe"] + 0.35 * 11.5) * quality_mult_adj, 1)
        fwd_ni_pv = max(1.0, net_income) * fwd_factor * pv_discount_2y_eq
        fwd_rev_pv = revenue * fwd_factor * pv_discount_2y_eq
        peer_rows = [
            {
                "method": "Cena / Čistý zisk (Odiskontovaný Forward 2Y P/E)",
                "multiple": adj_pe,
                "forward_metric": round(fwd_ni_pv, 2),
                "implied_share_price": round(max(0.0, (fwd_ni_pv * adj_pe) / shares_outstanding), 2),
            },
            {
                "method": "Cena / Účetní hodnota upravená o ROE (P/B)",
                "multiple": round(roe_adjusted_pb, 2),
                "forward_metric": round(book_equity, 2),
                "implied_share_price": round(max(0.0, (book_equity * roe_adjusted_pb) / shares_outstanding), 2),
            },
            {
                "method": "Tržní kapitalizace / Výnosy (Odiskontovaný Forward 2Y P/S)",
                "multiple": bench["ev_revenue"],
                "forward_metric": round(fwd_rev_pv, 2),
                "implied_share_price": round(max(0.0, (fwd_rev_pv * bench["ev_revenue"]) / shares_outstanding), 2),
            },
        ]
        comps_prices = [r["implied_share_price"] for r in peer_rows if r["implied_share_price"] > 0]
        blended_comps_price = sum(comps_prices) / len(comps_prices) if comps_prices else fcfe_per_share
    else:
        mature_ev_ebitda = 0.62 * bench["ev_ebitda"] + 0.38 * 14.5
        mature_pe = 0.62 * bench["pe"] + 0.38 * 19.5
        mature_ev_rev = 0.62 * bench["ev_revenue"] + 0.38 * 4.5

        adj_ev_ebitda = round(mature_ev_ebitda * quality_mult_adj, 1)
        adj_pe = round(mature_pe * quality_mult_adj, 1)
        adj_ev_rev = round(mature_ev_rev * quality_mult_adj, 1)

        fwd_ebitda_pv = normalized_ebitda * fwd_factor * pv_discount_2y_ev
        fwd_ni_pv = max(1.0, net_income) * fwd_factor * pv_discount_2y_eq
        fwd_rev_pv = revenue * fwd_factor * pv_discount_2y_ev

        peer_rows = [
            {
                "method": "Hodnota podniku / Provozní zisk (Odiskont. Forward 2Y EV/EBITDA)",
                "multiple": adj_ev_ebitda,
                "forward_metric": round(fwd_ebitda_pv, 2),
                "implied_share_price": round(max(0.0, (fwd_ebitda_pv * adj_ev_ebitda + total_cash - total_debt) / shares_outstanding), 2),
            },
            {
                "method": "Cena / Čistý zisk (Odiskont. Forward 2Y P/E)",
                "multiple": adj_pe,
                "forward_metric": round(fwd_ni_pv, 2),
                "implied_share_price": round(max(0.0, (fwd_ni_pv * adj_pe) / shares_outstanding), 2),
            },
            {
                "method": "Hodnota podniku / Tržby (Odiskont. Forward 2Y EV/Sales)",
                "multiple": adj_ev_rev,
                "forward_metric": round(fwd_rev_pv, 2),
                "implied_share_price": round(max(0.0, (fwd_rev_pv * adj_ev_rev + total_cash - total_debt) / shares_outstanding), 2),
            },
            {
                "method": "Cena / Účetní hodnota vlastního kapitálu (P/B)",
                "multiple": bench["pb"],
                "forward_metric": round(book_equity, 2),
                "implied_share_price": round(max(0.0, (book_equity * bench["pb"]) / shares_outstanding), 2),
            },
        ]
        if is_capex_heavy_growth:
            comps_prices = [peer_rows[0]["implied_share_price"], peer_rows[2]["implied_share_price"]]
            comps_prices = [p for p in comps_prices if p > 0]
        else:
            comps_prices = [r["implied_share_price"] for r in peer_rows[:3] if r["implied_share_price"] > 0]
        blended_comps_price = sum(comps_prices) / len(comps_prices) if comps_prices else dcf_intrinsic_share

    core_segment_ev = normalized_ebitda * eff_exit_mult * 0.80
    growth_segment_ev = normalized_ebitda * (eff_exit_mult * 1.15) * 0.20
    sotp_ev = core_segment_ev + growth_segment_ev
    sotp_equity = max(0.0, sotp_ev + total_cash - total_debt - minority_interest)
    sotp_per_share = (book_equity * 1.05 / shares_outstanding) if is_financial else (sotp_equity / shares_outstanding)
    nav_per_share = max(0.0, (book_equity + (0.0 if is_financial else (total_cash - total_debt) * 0.25)) / shares_outstanding)

    lbo_base_ev = max(market_cap, enterprise_value)
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

    wacc_grid = [max(0.05, wacc + d) for d in [-0.015, -0.0075, 0.0, 0.0075, 0.015]]
    tg_grid = [0.015, 0.020, 0.025, 0.030, 0.035]
    dcf_sensitivity = []
    for w_val in wacc_grid:
        cells = []
        for tg_val in tg_grid:
            safe_tg = min(tg_val, w_val - 0.015)
            p_val = simulate_fcff_dcf(eff_growth, w_val, eff_exit_mult, safe_tg)["intrinsic_value_per_share"]
            cells.append({"tg": round(tg_val * 100.0, 2), "price": round(p_val, 2)})
        dcf_sensitivity.append({"wacc": round(w_val * 100.0, 2), "cells": cells})

    bear_val = simulate_fcff_dcf(eff_growth * 0.55, wacc + 0.015, eff_exit_mult * 0.78, max(0.015, terminal_growth - 0.005))["intrinsic_value_per_share"]
    base_val = dcf_intrinsic_share
    bull_val = simulate_fcff_dcf(eff_growth * 1.30, max(0.065, wacc - 0.010), eff_exit_mult * 1.18, min(0.035, terminal_growth + 0.005))["intrinsic_value_per_share"]
    scenario_expected = 0.25 * bear_val + 0.50 * base_val + 0.25 * bull_val

    rng = random.Random(42)
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
            c_eb = normalized_ebitda * sim_margin_shock
            pv_cf = 0.0
            for yr in range(1, 6):
                fade = (yr - 1) / 4.0
                yg = sim_g * (1.0 - 0.35 * fade) + terminal_growth * (0.35 * fade)
                c_fcf *= (1.0 + yg)
                c_eb *= (1.0 + yg)
                pv_cf += c_fcf / ((1.0 + sim_wacc) ** yr)
            tv_g = (c_fcf * (1.0 + terminal_growth)) / max(0.015, sim_wacc - terminal_growth)
            tv_m = c_eb * sim_mult
            ev_m = pv_cf + ((0.5 * tv_g + 0.5 * tv_m) / ((1.0 + sim_wacc) ** 5))
            eq_m = max(0.0, ev_m + total_cash - total_debt - minority_interest)
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

    # Jednotný a konzistentní váhový rámec napříč všemi společnostmi (s úpravou pouze pro finanční instituce)
    if is_financial:
        weights = {
            "Vícefázový DCF pro podnik": 0.00,
            "DCF pro vlastní kapitál": 0.30,
            "Srovnatelné tržní násobky": 0.30,
            "Model reziduálního zisku": 0.30,
            "Metoda součtu částí": 0.00,
            "Medián Monte Carlo (P50)": 0.10,
        }
    else:
        weights = {
            "Vícefázový DCF pro podnik": 0.35,
            "DCF pro vlastní kapitál": 0.15,
            "Srovnatelné tržní násobky": 0.20,
            "Model reziduálního zisku": 0.05,
            "Metoda součtu částí": 0.10,
            "Medián Monte Carlo (P50)": 0.15,
        }

    method_values = {
        "Vícefázový DCF pro podnik": dcf_intrinsic_share,
        "DCF pro vlastní kapitál": fcfe_per_share,
        "Srovnatelné tržní násobky": blended_comps_price,
        "Model reziduálního zisku": ri_per_share,
        "Metoda součtu částí": sotp_per_share,
        "Medián Monte Carlo (P50)": mc_p50,
    }

    weighted_fair_value = sum(weights[m] * method_values[m] for m in weights)
    synthesis_rows = []
    for m_name, w_val in weights.items():
        v_price = method_values[m_name]
        diff_pct = ((v_price / current_price) - 1.0) * 100.0 if current_price > 0 else 0.0
        synthesis_rows.append({
            "method": m_name,
            "weight_percent": round(w_val * 100.0, 1),
            "fair_value": round(v_price, 2),
            "contribution": round(w_val * v_price, 2),
            "vs_market_percent": round(diff_pct, 2),
        })

    integrity_score = 65
    if len(rev_hist) >= 3:
        integrity_score += 10
    if not is_capex_heavy_growth:
        integrity_score += 10
    if roic > wacc or (is_financial and roe_est > cost_of_equity):
        integrity_score += 10
    primary_intrinsic = fcfe_per_share if is_financial else dcf_intrinsic_share
    spread_ratio = abs(primary_intrinsic - blended_comps_price) / max(1.0, weighted_fair_value)
    if spread_ratio < 0.20:
        integrity_score += 5
    elif spread_ratio > 0.45:
        integrity_score -= 10
    integrity_score = max(35, min(98, integrity_score))

    if is_financial:
        disagreement_note = (
            f"Aktivován bankovní/finanční režim (sektor {sector}): váha byla přesměrována na Model reziduálního zisku "
            f"({ri_per_share:,.2f} {currency}), DCF pro vlastní kapitál ({fcfe_per_share:,.2f} {currency}) a násobky P/E a P/B."
        )
    elif is_capex_heavy_growth:
        disagreement_note = (
            f"Aktivován růstový režim s normalizací expanzního CapEx: firma masivně investuje do infrastruktury, "
            f"proto byla posílena váha výhledových tržních násobků ({blended_comps_price:,.2f} {currency}) a očištěného cash flow."
        )
    elif spread_ratio > 0.25:
        disagreement_note = (
            f"Rozptyl mezi výnosovým modelem ({primary_intrinsic:,.2f} {currency}) a tržními násobky "
            f"({blended_comps_price:,.2f} {currency}) dosahuje {spread_ratio * 100:.1f} %."
        )
    else:
        disagreement_note = (
            f"Vysoká konvergence oceňovacích přístupů: výnosový model ({primary_intrinsic:,.2f} {currency}) "
            f"a relativní násobky ({blended_comps_price:,.2f} {currency}) potvrzují konzistentní odhad vnitřní hodnoty."
        )

    if data_warnings:
        disagreement_note += " Upozornění na vstupní data: " + " ".join(data_warnings)

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

    f_score = 0
    if net_income > 0:
        f_score += 1
    if (raw_ocf or 0) > 0:
        f_score += 1
    if (raw_ocf or 0) > net_income:
        f_score += 1
    if roic > wacc or (is_financial and roe_est > cost_of_equity):
        f_score += 1
    net_debt = total_debt - total_cash
    if is_financial or net_debt <= 0 or (total_debt / max(1.0, reported_ebitda)) < 2.2:
        f_score += 1
    cur_ratio = (cur_assets_hist[0] / cur_liab_hist[0]) if (cur_assets_hist and cur_liab_hist and cur_liab_hist[0] > 0) else 1.8
    if is_financial or cur_ratio > 1.2:
        f_score += 1
    if len(rev_hist) < 2 or rev_hist[0] >= rev_hist[1]:
        f_score += 1
    if len(ebit_hist) < 2 or (ebit_hist[0] / max(1.0, rev_hist[0])) >= (ebit_hist[1] / max(1.0, rev_hist[1])):
        f_score += 1
    if is_financial or (revenue / max(1.0, total_assets_hist[0] if total_assets_hist else invested_capital)) > 0.45:
        f_score += 1

    if is_financial:
        altman_z = 0.0
        altman_status = f"Finanční instituce (ROE {roe_est * 100:.1f} %)"
    else:
        tot_assets = total_assets_hist[0] if total_assets_hist else (invested_capital + total_cash)
        wc = (cur_assets_hist[0] - cur_liab_hist[0]) if (cur_assets_hist and cur_liab_hist) else (total_cash * 0.6)
        ret_earn = retained_earn_hist[0] if retained_earn_hist else (book_equity * 0.65)
        altman_z = (
            1.2 * (wc / max(1.0, tot_assets))
            + 1.4 * (ret_earn / max(1.0, tot_assets))
            + 3.3 * (reported_ebit / max(1.0, tot_assets))
            + 0.6 * (market_cap / max(1.0, total_debt if total_debt > 0 else tot_assets * 0.25))
            + 1.0 * (revenue / max(1.0, tot_assets))
        )
        altman_status = "Bezpečné pásmo (nad 2,99)" if altman_z > 2.99 else ("Šedá zóna (1,81–2,99)" if altman_z >= 1.81 else "Zvýšené riziko (pod 1,81)")

    interest_coverage = (reported_ebit / interest_expense) if interest_expense > 0 else 99.0

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
        "dynamic_defaults": {
            "leverage_percent": round(auto_lev * 100.0, 1),
            "growth_percent": round(auto_grw * 100.0, 1),
            "exit_multiple": round(auto_mult, 1),
        },
        "trzni_kapitalizace": round(market_cap, 2),
        "celkovy_dluh": round(total_debt, 2),
        "hotovost": round(total_cash, 2),
        "cisty_dluh": round(net_debt, 2),
        "celkova_hodnota_podniku": round(enterprise_value, 2),
        "fundamentals": {
            "revenue": round(revenue, 2),
            "reported_ebitda": round(reported_ebitda, 2),
            "normalized_ebitda": round(normalized_ebitda, 2),
            "reported_fcf": round(reported_fcf, 2),
            "normalized_fcf": round(normalized_fcf, 2),
            "sbc_expense": round(sbc_expense, 2),
            "net_income": round(net_income, 2),
            "roic_percent": round((roe_est if is_financial else roic) * 100.0, 2),
            "economic_profit": round(economic_profit, 2),
            "fcf_conversion_percent": round((normalized_fcf / max(1.0, normalized_ebitda)) * 100.0, 1),
            "net_debt_to_ebitda": round(net_debt / max(1.0, normalized_ebitda), 2),
            "interest_coverage": round(interest_coverage, 1),
            "current_ratio": round(cur_ratio, 2),
            "operating_margin_percent": round((reported_ebit / max(1.0, revenue)) * 100.0, 2),
        },
        "wacc_engine": {
            "risk_free_rate_percent": round(rf_rate * 100.0, 2),
            "erp_percent": round(erp_rate * 100.0, 2),
            "beta": round(beta, 2),
            "cost_of_equity_percent": round(cost_of_equity * 100.0, 2),
            "pre_tax_cost_of_debt_percent": round(pre_tax_cost_of_debt * 100.0, 2),
            "after_tax_cost_of_debt_percent": round(pre_tax_cost_of_debt * (1.0 - effective_tax_rate) * 100.0, 2),
            "weight_equity_percent": round(weight_equity * 100.0, 1),
            "weight_debt_percent": round(weight_debt * 100.0, 1),
            "wacc_percent": round(wacc * 100.0, 2),
        },
        "dcf_valuation": {
            "wacc": round(wacc, 6),
            "terminal_growth": round(terminal_growth, 4),
            "terminal_value": round(dcf_res["terminal_value"], 2),
            "tv_share_percent": round(dcf_res["tv_share_percent"], 1),
            "enterprise_value": round(dcf_res["enterprise_value"], 2),
            "equity_value": round(dcf_res["equity_value"], 2),
            "intrinsic_value_per_share": round(dcf_intrinsic_share, 2),
            "current_share_price": round(current_price, 2),
            "upside_downside_percent": round(dcf_upside, 2),
            "cash_flows": dcf_res["cash_flows"],
            "sensitivity_matrix": dcf_sensitivity,
        },
        "fcfe_valuation": {
            "cost_of_equity": round(cost_of_equity, 4),
            "equity_value": round(fcfe_equity_val, 2),
            "intrinsic_value_per_share": round(fcfe_per_share, 2),
        },
        "reverse_dcf": {
            "implied_fcf_growth_percent": round(implied_growth_pct, 2),
            "assumed_growth_percent": round(eff_growth * 100.0, 2),
            "expectation_gap_percent": round((eff_growth * 100.0) - implied_growth_pct, 2),
        },
        "residual_income": {
            "book_value_equity": round(book_equity, 2),
            "roe_percent": round(roe_est * 100.0, 2),
            "equity_value": round(ri_equity_val, 2),
            "intrinsic_value_per_share": round(ri_per_share, 2),
        },
        "peer_comps": {
            "sector": sector,
            "blended_implied_price": round(blended_comps_price, 2),
            "multiples": peer_rows,
        },
        "metoda_souctu_casti": {
            "vypoctena_hodnota_podniku": round(sotp_ev, 2),
            "vypoctena_hodnota_vlastniho_kapitalu": round(sotp_equity, 2),
            "sotp_per_share": round(sotp_per_share, 2),
            "nav_per_share": round(nav_per_share, 2),
        },
        "model_odkupu_za_cizi_kapital_base_case": {
            "vstupni_vlastni_kapital": round(lbo_entry_equity, 2),
            "vstupni_cizi_kapital": round(lbo_entry_debt, 2),
            "vystupni_hodnota_vlastniho_kapitalu": round(lbo_exit_equity, 2),
            "moic": round(lbo_moic, 2),
            "irr_procent": round(lbo_irr, 2),
        },
        "projektovane_roky": lbo_years,
        "citlivostni_matice_lbo_irr": lbo_sensitivity,
        "scenarios": {
            "bear": {"probability": 25, "fair_value": round(bear_val, 2)},
            "base": {"probability": 50, "fair_value": round(base_val, 2)},
            "bull": {"probability": 25, "fair_value": round(bull_val, 2)},
            "expected_value": round(scenario_expected, 2),
        },
        "monte_carlo": {
            "iterations": 2500,
            "p5": round(mc_p5, 2),
            "p10": round(mc_p10, 2),
            "p25": round(mc_p25, 2),
            "p50": round(mc_p50, 2),
            "p75": round(mc_p75, 2),
            "p90": round(mc_p90, 2),
            "p95": round(mc_p95, 2),
            "prob_above_market_percent": round(prob_above_market, 1),
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
            "altman_z_score": round(altman_z, 2),
            "altman_status": altman_status,
        },
        "audit_trail": {
            "model_version": "Tensegrity Institutional Engine v4.1",
            "timestamp": get_prague_now().strftime("%Y-%m-%d %H:%M:%S %Z"),
            "data_source": "Yahoo Finance Institutional Feed + Live FX & Treasury",
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
# 6. HISTORICKÁ DATA PRO INTERAKTIVNÍ GRAF (TRADINGVIEW / TRADING 212 STYL)
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
# 7. SKUTEČNÝ VOLNÝ AI ANALYTICKÝ PORADCE (BEZ PŘEDNASTAVENÝCH ŠABLON)
# ==============================================================================

def _build_ai_advisor_system_prompt(d: dict) -> str:
    compact_context = {
        "spolecnost": d["company_name"],
        "ticker": d["ticker"],
        "sektor": d["sector"],
        "odvetvi": d["industry"],
        "mena": d["currency"],
        "trzni_cena": d["dcf_valuation"]["current_share_price"],
        "vazena_vnitrni_hodnota": d["valuation_synthesis"]["weighted_fair_value"],
        "zona": d["decision_engine"]["zone"],
        "bezpecnostni_rezerva_procent": d["decision_engine"]["margin_of_safety_percent"],
        "dcf_podnik": d["dcf_valuation"]["intrinsic_value_per_share"],
        "dcf_vlastni_kapital": d["fcfe_valuation"]["intrinsic_value_per_share"],
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
    return (
        "Jsi špičkový, neomezený finanční, právní, investiční a všeobecný AI poradce platformy Tensegrity Terminal. "
        "Odpovídej v perfektní češtině na JAKÝKOLIV dotaz uživatele — ať jde o jednoduchou definici, vysvětlení pojmu, "
        "hlubokou institucionální analýzu, právo/legislativu, makroekonomiku nebo libovolnou obecnou otázku v dlouhém "
        "vícekrokovém rozhovoru.\n"
        "KLÍČOVÉ PRAVIDLO ODBORNOSTI A FORMÁTOVÁNÍ:\n"
        "- Přizpůsob hloubku, odbornost a délku odpovědi přesně povaze a odbornosti otázky!\n"
        "- Nikdy nepoužívej surové LaTeX značky jako \\[ ... \\] nebo \\( ... \\) — vzorce a pojmy piš obyčejným českým textem.\n"
        "- Na jednoduché nebo definitorické otázky (např. 'co je nvda', 'co znamená dcf', 'co je wacc') odpovídej srozumitelně, "
        "přirozeně a stručně bez zbytečného zahlcování tabulkami nebo výpisem všech metrik terminálu, pokud si je uživatel výslovně nevyžádá.\n"
        "- Tabulky a hluboké finanční rozbory použij pouze tehdy, když se dotaz skutečně týká srovnání, dat nebo pokročilé analýzy.\n"
        "- Pečlivě sleduj předchozí zprávy v historii konverzace, aby navazující otázky dávaly plný kontext.\n"
        f"Volitelný kontext aktuálně otevřené firmy v terminálu: {json.dumps(compact_context, ensure_ascii=False)}"
    )


def generate_pure_ai_advisor_answer(question: str, history: list[dict], d: dict) -> str:
    sys_prompt = _build_ai_advisor_system_prompt(d)
    history_lines = []
    for m in history[-10:]:
        role_label = "Uživatel" if m.get("role") == "user" else "Poradce"
        history_lines.append(f"{role_label}: {str(m.get('text', ''))[:700]}")
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
                    return parts[0]["text"].strip()
        except Exception as err:
            logger.warning("Gemini API selhalo (%s), zkouším záložní AI...", err)

    for model_name in ("openai", "mistral", "llama"):
        try:
            messages = [{"role": "system", "content": sys_prompt}]
            for m in history[-10:]:
                messages.append({
                    "role": "user" if m.get("role") == "user" else "assistant",
                    "content": str(m.get("text", ""))[:700],
                })
            messages.append({"role": "user", "content": question})

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
                    "User-Agent": f"TensegrityTerminal/{random.randint(100,999)}",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=18) as resp:
                raw_json = json.loads(resp.read().decode("utf-8", errors="ignore"))
            choices = raw_json.get("choices") or []
            if choices:
                content = str((choices[0].get("message") or {}).get("content") or "").strip()
                if content and len(content) > 5 and "rate limit" not in content.lower():
                    return content
        except Exception:
            continue

    try:
        short_prompt = (
            f"Odpověz česky, srozumitelně a přímo na otázku bez tabulek, pokud nejsou vyžádány. "
            f"Otázka: {question}"
        )
        get_url = f"https://text.pollinations.ai/{quote(short_prompt[:1200])}?model=openai&seed={random.randint(1, 999999)}"
        req_get = urllib.request.Request(get_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req_get, timeout=16) as resp_get:
            txt_out = resp_get.read().decode("utf-8", errors="ignore").strip()
        if txt_out and len(txt_out) > 5 and not txt_out.startswith("<!DOCTYPE") and "rate limit" not in txt_out.lower():
            return txt_out
    except Exception:
        pass

    q_low = question.lower().strip()
    if "apple" in q_low or "aapl" in q_low:
        return (
            "**Apple (`AAPL`)** je americká technologická společnost známá především chytrými telefony iPhone, "
            "počítači Mac, tablety iPad, nositelnou elektronikou (Apple Watch, AirPods), operačními systémy "
            "a rostoucím segmentem služeb (App Store, iCloud, Apple Music, Apple Pay)."
        )
    if "dcf" in q_low and ("co je" in q_low or "co znamená" in q_low or "co to je" in q_low):
        return (
            "**DCF (Discounted Cash Flow)** znamená diskontování budoucích volných peněžních toků. "
            "Zjednodušeně jde o výpočet toho, jakou hodnotu mají peníze, které firma nebo projekt vydělá "
            "v budoucnu, přepočtené na dnešní hodnotu pomocí diskontní sazby (která zohledňuje čas a riziko)."
        )
    if "nvda" in q_low or "nvidia" in q_low:
        return (
            "**NVIDIA (`NVDA`)** je technologická společnost zaměřená hlavně na návrh grafických čipů (GPU), "
            "akcelerátorů pro umělou inteligenci, datových center a softwarového ekosystému CUDA."
        )

    return (
        "Váš dotaz jsem zaznamenal, ale externí AI poskytovatel právě neodpověděl. "
        "Zkuste prosím dotaz zopakovat nebo upřesnit, co přesně vás zajímá."
    )


# ==============================================================================
# 8. API ENDPOINTY, GRAF, EXTERNÍ CRON A EXPORTY
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
        raise HTTPException(status_code=404, detail=str(err))
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

        d = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)
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
        ws_sum["B2"] = f"Společnost: {data['company_name']} ({data['ticker']}) | Sektor: {data['sector']} | Datum: {data['audit_trail']['timestamp']}"
        ws_sum.append([])
        ws_sum.append(["Metoda ocenění", "Váha (%)", f"Vnitřní hodnota ({data['currency']})", "Vážený příspěvek", "Rozdíl vs. trh (%)"])
        for cell in ws_sum[4]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row in data["valuation_synthesis"]["methods"]:
            ws_sum.append([row["method"], row["weight_percent"], row["fair_value"], row["contribution"], row["vs_market_percent"]])

        ws_sum.append([])
        ws_sum.append(["Vážená vnitřní hodnota", data["valuation_synthesis"]["weighted_fair_value"]])
        ws_sum.append(["Aktuální tržní cena akcie", data["dcf_valuation"]["current_share_price"]])
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
            ws_mc.append([label, val])

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
        curr = data["currency"]

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
            Paragraph(f"Ticker: {data['ticker']} | Sektor: {data['sector']} | Datum: {data['audit_trail']['timestamp']} | Měna: {curr}", sub_s),
            Paragraph("1. SYNTÉZA VNITŘNÍ HODNOTY A NÁKUPNÍ PÁSMA", sec_s),
        ])

        tbl = Table(
            [
                ["Ukazatel / Oceňovací přístup", "Výstup modelu"],
                ["Aktuální tržní cena akcie", f"{dcf['current_share_price']:,.2f} {curr}"],
                ["Vážená vnitřní hodnota", f"{synth['weighted_fair_value']:,.2f} {curr}"],
                ["Vícefázový výnosový model pro podnik (FCFF)", f"{dcf['intrinsic_value_per_share']:,.2f} {curr}"],
                ["Výnosový model pro vlastní kapitál (FCFE)", f"{data['fcfe_valuation']['intrinsic_value_per_share']:,.2f} {curr}"],
                ["Model reziduálního zisku (Residual Income)", f"{data['residual_income']['intrinsic_value_per_share']:,.2f} {curr}"],
                ["Srovnatelné tržní násobky", f"{data['peer_comps']['blended_implied_price']:,.2f} {curr}"],
                ["Medián Monte Carlo (P50) / P10–P90", f"{mc['p50']:,.2f} ({mc['p10']:,.2f} – {mc['p90']:,.2f}) {curr}"],
                ["Bezpečnostní rezerva", f"{dec['margin_of_safety_percent']:.2f} %"],
                ["Aktivní cenová zóna a alokace", f"{dec['zone']} (Tranše {dec['tranche_allocation_percent']} %)"],
                ["Index integrity valuace / Piotroski skóre", f"{synth['integrity_score']} / 100 | {data['risk_analysis']['piotroski_f_score']} / 9"],
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
# 9. KOMPLETNÍ WEBOVÉ ROZHRANÍ (TRADINGVIEW GRAF + VOLNÝ AI PORADCE + VALUACE)
# ==============================================================================

@app.get("/", response_class=HTMLResponse)
def read_root():
    return r"""<!DOCTYPE html>
<html lang="cs">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tensegrity Terminal | Institucionální Valuační Platforma</title>
    <link rel="icon" type="image/png" href="/favicon.ico?v=7">
    <link rel="shortcut icon" type="image/png" href="/favicon.ico?v=7">
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
            outline: none; width: 220px; letter-spacing: 0.3px;
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

        .company-identity-banner {
            display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 16px;
            margin-top: 18px; padding: 14px 18px; border-radius: 12px;
            background: rgba(6, 9, 17, 0.75); border: 1px solid rgba(56, 189, 248, 0.22);
        }
        .company-identity-left { display: flex; align-items: center; gap: 14px; }
        .company-logo-frame {
            width: 44px; height: 44px; border-radius: 10px; background: rgba(255, 255, 255, 0.94);
            display: flex; align-items: center; justify-content: center; overflow: hidden;
            border: 1px solid rgba(255,255,255,0.2); flex-shrink: 0;
        }
        .company-logo-frame img { width: 32px; height: 32px; object-fit: contain; display: block; }
        .company-logo-fallback {
            width: 44px; height: 44px; border-radius: 10px;
            background: linear-gradient(135deg, rgba(56,189,248,0.25), rgba(2,132,199,0.45));
            color: #38bdf8; font-weight: 800; font-size: 16px;
            display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        }
        .company-title-main { font-size: 16px; font-weight: 800; color: #f8fafc; letter-spacing: 0.2px; }
        .company-ticker-pill {
            display: inline-block; margin-left: 8px; padding: 2px 8px; border-radius: 6px;
            background: rgba(56, 189, 248, 0.14); color: #38bdf8;
            font: 700 12px 'JetBrains Mono', monospace;
        }
        .company-meta-sub { font-size: 12px; color: var(--text-secondary); margin-top: 3px; }

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
        .grid-4 { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-top: 14px; }
        .card {
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 22px 26px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .matrix-card { margin-bottom: 22px; }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid var(--border-subtle); flex-wrap: wrap; gap: 8px; }
        .card-title { font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.4px; color: var(--text-secondary); }
        .badge { font-size: 11px; font-family: 'JetBrains Mono', monospace; padding: 3px 8px; border-radius: 6px; background: rgba(56, 189, 248, 0.1); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.25); }
        .metric-row { display: flex; justify-content: space-between; align-items: center; margin: 10px 0; font-size: 13.5px; }
        .metric-label { color: var(--text-secondary); }
        .metric-value { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: var(--text-primary); font-size: 14px; }
        .metric-value.highlight { color: var(--accent-blue); font-size: 15px; }

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

        .table-responsive { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 8px; }
        th, td { padding: 10px 14px; text-align: center; font-size: 12.5px; border-bottom: 1px solid var(--border-subtle); }
        th { background: rgba(6, 9, 17, 0.8); color: var(--text-secondary); font-weight: 700; text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.6px; }
        th:first-child, td:first-child { text-align: left; }
        td { font-family: 'JetBrains Mono', monospace; color: var(--text-primary); }

        .plain-guide { margin-top: 10px; font-size: 12px; color: var(--text-secondary); line-height: 1.55; }
        .advisor-thread {
            display: flex; flex-direction: column; gap: 14px; margin: 16px 0;
            max-height: 480px; overflow-y: auto; padding: 8px 10px 8px 2px;
            scrollbar-width: thin; scrollbar-color: rgba(56, 189, 248, 0.35) rgba(6, 9, 17, 0.4);
        }
        .advisor-thread::-webkit-scrollbar { width: 6px; }
        .advisor-thread::-webkit-scrollbar-track { background: rgba(6, 9, 17, 0.4); border-radius: 6px; }
        .advisor-thread::-webkit-scrollbar-thumb { background: rgba(56, 189, 248, 0.35); border-radius: 6px; }
        .advisor-msg-row { display: flex; gap: 12px; align-items: flex-start; max-width: 92%; }
        .advisor-msg-row.user-row { align-self: flex-end; flex-direction: row-reverse; }
        .advisor-avatar {
            width: 32px; height: 32px; border-radius: 10px; flex-shrink: 0;
            display: flex; align-items: center; justify-content: center;
            font: 800 11px 'JetBrains Mono', monospace;
            background: linear-gradient(135deg, rgba(56, 189, 248, 0.25), rgba(2, 132, 199, 0.45));
            color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.35);
        }
        .advisor-msg-row.user-row .advisor-avatar {
            background: rgba(16, 185, 129, 0.2); color: #34d399; border-color: rgba(16, 185, 129, 0.35);
        }
        .advisor-bubble {
            padding: 14px 18px; border-radius: 14px; font-size: 13.5px; line-height: 1.65;
            background: rgba(10, 16, 28, 0.92); border: 1px solid rgba(56, 189, 248, 0.18);
            color: var(--text-primary); box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
        }
        .advisor-msg-row.user-row .advisor-bubble {
            background: rgba(16, 185, 129, 0.12); border-color: rgba(16, 185, 129, 0.3);
        }
        .advisor-bubble h3, .advisor-bubble h4 {
            color: #38bdf8; font-size: 14px; font-weight: 800; margin: 10px 0 6px 0; letter-spacing: 0.2px;
        }
        .advisor-bubble h3:first-child, .advisor-bubble h4:first-child { margin-top: 0; }
        .advisor-bubble p { margin: 0 0 8px 0; }
        .advisor-bubble p:last-child { margin-bottom: 0; }
        .advisor-bubble ul { margin: 6px 0 10px 18px; }
        .advisor-bubble li { margin-bottom: 4px; }
        .advisor-bubble strong { color: #fff; font-weight: 700; }
        .advisor-bubble table {
            width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 12px;
            background: rgba(6, 9, 17, 0.7); border-radius: 8px; overflow: hidden;
        }
        .advisor-bubble th {
            background: rgba(56, 189, 248, 0.14); color: #38bdf8; padding: 7px 10px; text-align: left; font-size: 11px;
        }
        .advisor-bubble td {
            padding: 7px 10px; border-bottom: 1px solid var(--border-subtle); font-family: 'Plus Jakarta Sans', sans-serif; text-align: left;
        }
        .advisor-form { display: flex; gap: 10px; margin-top: 12px; }
        .advisor-form input { flex: 1; padding: 12px 16px; border: 1px solid var(--border-subtle); border-radius: 10px; background: rgba(6, 9, 17, 0.85); color: var(--text-primary); font-size: 13.5px; outline: none; }
        .advisor-form input:focus { border-color: var(--accent-blue); }

        .modal-backdrop {
            display: none; position: fixed; inset: 0; background: rgba(4, 7, 15, 0.84);
            backdrop-filter: blur(6px); z-index: 999; align-items: center; justify-content: center; padding: 20px;
        }
        .modal-backdrop.open { display: flex; }
        .modal-box {
            background: #0f1625; border: 1px solid var(--border-subtle); border-radius: 16px;
            max-width: 720px; width: 100%; padding: 28px; box-shadow: 0 20px 50px rgba(0,0,0,0.75);
            max-height: 86vh; overflow-y: auto;
        }
        .news-item { padding: 13px 0; border-bottom: 1px solid var(--border-subtle); }
        .news-item:last-child { border-bottom: none; }
        .news-item a { color: var(--accent-blue); text-decoration: none; font-weight: 600; font-size: 13.5px; line-height: 1.4; display: block; }
        .news-item a:hover { text-decoration: underline; }
        .news-meta { font-size: 11px; color: var(--text-muted); margin-top: 4px; font-family: 'JetBrains Mono', monospace; }
        .loader { width: 22px; height: 22px; border: 3px solid rgba(255,255,255,0.2); border-radius: 50%; border-top-color: var(--accent-blue); animation: spin 0.8s linear infinite; margin: auto; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="brand-wrap">
                <div class="brand-logo-box">
                    <img src="/brand-logo.png?v=7" alt="Tensegrity Terminal Logo" onerror="this.parentElement.style.display='none'">
                </div>
                <div>
                    <div class="brand-title">Tensegrity<span>Terminal</span></div>
                    <div class="brand-subtitle">Intelligence &bull; Adaptivní Institucionální Valuace</div>
                </div>
            </div>
            <div class="controls">
                <div class="search-wrapper">
                    <input type="text" id="tickerInput" value="NVDA" placeholder="Název firmy nebo ticker..." aria-label="Název společnosti nebo ticker" onkeydown="if(event.key==='Enter') handleSearchTrigger()">
                </div>
                <div class="forex-strip" id="forexRates">
                    <span>EUR/CZK <strong>--</strong></span>
                    <span>USD/CZK <strong>--</strong></span>
                    <span>GBP/CZK <strong>--</strong></span>
                </div>
                <button class="btn btn-primary" onclick="handleSearchTrigger()">Analyzovat</button>
                <button class="btn btn-excel" onclick="downloadExcel()">Excel Model</button>
                <button class="btn btn-pdf" onclick="downloadPDF()">PDF Zpráva</button>
                <button class="btn btn-news" onclick="openNewsletterModal()">Newsletter &amp; Zprávy</button>
            </div>
        </header>

        <section class="scenario-bar" aria-label="Parametry scénáře">
            <span class="scenario-heading">Parametry Scénáře</span>
            <label class="scenario-field" for="leverageInput">
                Cizí kapitál (% Dluhu)
                <span class="scenario-input"><input id="leverageInput" type="number" min="0" max="100" step="0.5" value="15"><span>%</span></span>
            </label>
            <label class="scenario-field" for="growthInput">
                Růst peněžních toků
                <span class="scenario-input"><input id="growthInput" type="number" min="-50" max="150" step="0.5" value="14"><span>%</span></span>
            </label>
            <label class="scenario-field" for="exitMultipleInput">
                Výstupní násobek
                <span class="scenario-input"><input id="exitMultipleInput" type="number" min="1" max="150" step="0.5" value="21.0"><span>x</span></span>
            </label>
            <button class="btn btn-primary" style="padding: 6px 14px;" onclick="loadValuation(false)">Přepočítat model</button>
            <span class="audit-pill" id="auditTimestamp">Načítám auditní stopu...</span>
        </section>

        <section class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Rozhodovací Rámec — 4 Nákupní Zóny &amp; Alokace Tranší</span>
                <span class="badge" id="integrityBadge">Index integrity: --</span>
            </div>
            <div class="thesis-layout">
                <div class="traffic-lights" aria-label="Semafor zón">
                    <span class="traffic-light green" id="lightGreen"></span>
                    <span class="traffic-light yellow" id="lightYellow"></span>
                    <span class="traffic-light orange" id="lightOrange"></span>
                    <span class="traffic-light red" id="lightRed"></span>
                </div>
                <div>
                    <p class="metric-value" id="thesisRecommendation">Načítám institucionální syntézu...</p>
                    <p class="plain-guide" id="thesisSummary">Porovnání vážené vnitřní hodnoty se současnou tržní cenou a výpočet bezpečnostní rezervy.</p>
                </div>
            </div>

            <div class="company-identity-banner">
                <div class="company-identity-left">
                    <div id="companyLogoContainer" class="company-logo-fallback">NV</div>
                    <div>
                        <div>
                            <span class="company-title-main" id="companyFullName">NVIDIA Corporation</span>
                            <span class="company-ticker-pill" id="companyTickerBadge">NVDA</span>
                        </div>
                        <div class="company-meta-sub" id="companySectorMeta">Sektor: Technology &bull; Odvětví: Semiconductors &bull; Burza: NMS</div>
                    </div>
                </div>
                <div class="badge" id="companyPricePill">Tržní cena: --</div>
            </div>

            <div class="grid-4">
                <div class="zone-box green">
                    <div class="zone-title">Zóna GREEN (Tranše 55 %)</div>
                    <div class="zone-val" id="zoneGreenRange">-</div>
                    <p class="plain-guide" style="margin:0;">Bezpečnostní rezerva nad 20 % pod váženou vnitřní hodnotou.</p>
                </div>
                <div class="zone-box yellow">
                    <div class="zone-title">Zóna YELLOW (Tranše 30 %)</div>
                    <div class="zone-val" id="zoneYellowRange">-</div>
                    <p class="plain-guide" style="margin:0;">Atraktivní diskont 10 % až 20 % pod vnitřní hodnotou.</p>
                </div>
                <div class="zone-box orange">
                    <div class="zone-title">Zóna ORANGE (Tranše 15 %)</div>
                    <div class="zone-val" id="zoneOrangeRange">-</div>
                    <p class="plain-guide" style="margin:0;">Úvodní akumulační pásmo do 10 % pod vnitřní hodnotou.</p>
                </div>
                <div class="zone-box red">
                    <div class="zone-title">Zóna RED (Tranše 0 %)</div>
                    <div class="zone-val" id="zoneRedRange">-</div>
                    <p class="plain-guide" style="margin:0;">Cena nad odhadovanou vnitřní hodnotou; vyčkat.</p>
                </div>
            </div>
        </section>

        <section class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Interaktivní Cenový Graf &amp; Historický Vývoj (TradingView / Trading 212)</span>
                <div style="display:flex; gap:8px; align-items:center;">
                    <span class="badge" id="tvPeriodChangeBadge">Změna období: --</span>
                    <button class="tv-tf-btn active" id="btnChartArea" onclick="setTvChartStyle('area')">Plošný (T212)</button>
                    <button class="tv-tf-btn" id="btnChartCandle" onclick="setTvChartStyle('candle')">Svíčky (TV)</button>
                </div>
            </div>
            <div class="tv-toolbar">
                <div class="tv-tf-group" id="tvTimeframeGroup">
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('1D', this)">1D</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('1T', this)">1T</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('1M', this)">1M</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('6M', this)">6M</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('YTD', this)">YTD</button>
                    <button class="tv-tf-btn active" onclick="changeTvTimeframe('1R', this)">1R</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('5R', this)">5R</button>
                    <button class="tv-tf-btn" onclick="changeTvTimeframe('MAX', this)">MAX</button>
                </div>
                <div class="tv-readout" id="tvCrosshairReadout">
                    <span>Najeďte myší na graf pro detailní cenu, čas a objem obchodů (kolečkem myši přiblížíte osu).</span>
                </div>
            </div>
            <div id="tvChartContainer"></div>
        </section>

        <section class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Syntéza Ocenění — Adaptivní Spojení Oceňovacích Metod</span>
                <span class="badge" id="weightedFairBadge">Vážená vnitřní hodnota: --</span>
            </div>
            <div class="table-responsive" id="synthesisTableContainer">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
            <p class="plain-guide" id="disagreementNote"></p>
        </section>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Živý Graf: Srovnání Metod Ocenění</span>
                    <span class="badge">Cena na akcii</span>
                </div>
                <div class="chart-container">
                    <canvas id="methodsChartCanvas"></canvas>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Živý Graf: Rozdělení Monte Carlo</span>
                    <span class="badge">2 500 simulací</span>
                </div>
                <div class="chart-container">
                    <canvas id="monteCarloChartCanvas"></canvas>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Živý Graf: Projekce Toků &amp; Dluhu</span>
                    <span class="badge">Horizont 5 let</span>
                </div>
                <div class="chart-container">
                    <canvas id="cashFlowChartCanvas"></canvas>
                </div>
            </div>
        </div>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Fundamenty &amp; Normalizace</span>
                    <span class="badge" id="fundTicker">NVDA</span>
                </div>
                <div class="metric-row"><span class="metric-label">Tržní kapitalizace:</span><span class="metric-value" id="valCap">-</span></div>
                <div class="metric-row"><span class="metric-label">Hodnota podniku (EV):</span><span class="metric-value highlight" id="valEV">-</span></div>
                <div class="metric-row"><span class="metric-label">Normalizovaný volný tok:</span><span class="metric-value" id="normFcf">-</span></div>
                <div class="metric-row"><span class="metric-label">Návratnost vs. náklady kapitálu:</span><span class="metric-value highlight" id="roicVsWacc">-</span></div>
                <div class="metric-row"><span class="metric-label">Ekonomický zisk:</span><span class="metric-value" id="econProfit">-</span></div>
                <div class="metric-row"><span class="metric-label">Piotroski / Altman Z-skóre:</span><span class="metric-value" id="riskScores">-</span></div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Výnosové Modely &amp; Tržní Očekávání</span>
                    <span class="badge">Diskontované toky</span>
                </div>
                <div class="metric-row"><span class="metric-label">Vícefázový DCF (podnik) / akcii:</span><span class="metric-value highlight" id="dcfIntrinsicValue">-</span></div>
                <div class="metric-row"><span class="metric-label">DCF (vlastní kapitál) / Reziduální zisk:</span><span class="metric-value" id="fcfeAndRi">-</span></div>
                <div class="metric-row"><span class="metric-label">Náklady kapitálu (Beta):</span><span class="metric-value" id="dcfWacc">-</span></div>
                <div class="metric-row"><span class="metric-label">Podíl terminální hodnoty:</span><span class="metric-value" id="tvShare">-</span></div>
                <div class="metric-row"><span class="metric-label">Obrácený DCF implikovaný růst:</span><span class="metric-value highlight" id="reverseDcfGrowth">-</span></div>
                <div class="metric-row"><span class="metric-label">Součet částí / Čistá aktiva:</span><span class="metric-value" id="sotpNav">-</span></div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Scénáře &amp; Pravděpodobnost</span>
                    <span class="badge">Distribuce hodnot</span>
                </div>
                <div class="metric-row"><span class="metric-label">Pesimistický / Základní / Optimistický:</span><span class="metric-value" id="bearBaseBull">-</span></div>
                <div class="metric-row"><span class="metric-label">Očekávaná hodnota scénářů:</span><span class="metric-value highlight" id="scenExpected">-</span></div>
                <div class="metric-row"><span class="metric-label">Monte Carlo Medián (P50):</span><span class="metric-value highlight" id="mcP50">-</span></div>
                <div class="metric-row"><span class="metric-label">Interval P10 – P90:</span><span class="metric-value" id="mcP10P90">-</span></div>
                <div class="metric-row"><span class="metric-label">Extrémní rozpětí P5 – P95:</span><span class="metric-value" id="mcP5P95">-</span></div>
                <div class="metric-row"><span class="metric-label">Pravděpodobnost hodnoty nad trhem:</span><span class="metric-value highlight" id="mcProbUpside">-</span></div>
            </div>
        </div>

        <div class="grid-2">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Relativní Ocenění (Srovnatelné Násobky)</span>
                    <span class="badge" id="peerSector">Sektor</span>
                </div>
                <div class="table-responsive" id="peerCompsContainer">
                    <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Pákový Odkup &amp; Splácení Dluhu</span>
                    <span class="badge" id="lboSummaryBadge">Násobek / Výnos</span>
                </div>
                <div class="table-responsive" id="lboScheduleContainer">
                    <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
                </div>
            </div>
        </div>

        <div class="grid-2">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Citlivostní Matice (Náklady kapitálu vs. Terminální růst)</span>
                    <span class="badge">Vnitřní hodnota / akcii</span>
                </div>
                <div class="table-responsive" id="dcfSensitivityContainer"></div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Citlivostní Matice Pákového Odkupu (Růst vs. Násobek)</span>
                    <span class="badge">Roční výnosnost (%)</span>
                </div>
                <div class="table-responsive" id="sensitivityMatrix"></div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <span class="card-title">Tensegrity AI Analytický Poradce</span>
                <span class="badge">Neomezená umělá inteligence</span>
            </div>
            <div class="advisor-thread" id="advisorThread">
                <div class="advisor-msg-row">
                    <div class="advisor-avatar">AI</div>
                    <div class="advisor-bubble">
                        <p>Dobrý den. Jsem plnohodnotná umělá inteligence napojená přímo na živé výkazy, valuace i tržní kontext vybrané společnosti. Zeptejte se mě úplně na cokoliv — na detailní rozbor ocenění, srovnání s konkurencí, rizika, makroekonomiku, investiční strategii nebo jakoukoliv obecnou otázku.</p>
                    </div>
                </div>
            </div>
            <form class="advisor-form" onsubmit="askAdvisor(event)">
                <input id="advisorQuestion" type="text" placeholder="Zeptejte se umělé inteligence úplně na cokoliv..." required>
                <button class="btn btn-primary" id="advisorSubmitBtn" type="submit">Odeslat</button>
            </form>
        </div>
    </div>

    <div class="modal-backdrop" id="newsletterModal" onclick="if(event.target===this) closeNewsletterModal()">
        <div class="modal-box">
            <div class="card-header">
                <span class="card-title">Tensegrity Terminal | Intelligence Newsletter (<span id="newsModalTicker">NVDA</span>)</span>
                <button class="btn btn-news" style="padding: 4px 10px;" onclick="closeNewsletterModal()">&#10005;</button>
            </div>
            <div style="margin-bottom: 14px;">
                <input id="subscribeEmailInput" type="email" placeholder="Zadejte svou e-mailovou adresu..." style="width: 100%; padding: 12px 16px; border: 1px solid var(--border-subtle); border-radius: 8px; background: rgba(6, 9, 17, 0.85); color: var(--text-primary); font-size: 13px; outline: none; margin-bottom: 10px;">
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <button class="btn btn-primary" style="flex: 1; justify-content: center;" onclick="sendNewsletterNow()">Odeslat jednorázově nyní</button>
                    <button class="btn btn-secondary" style="flex: 1; justify-content: center;" onclick="subscribeNewsletter(event)">Zaevidovat pravidelný odběr</button>
                </div>
            </div>
            <div id="subscribeStatus" style="font-size: 12.5px; margin-bottom: 12px; font-weight: 600;"></div>
            <div style="font-size: 11.5px; color: var(--text-muted); margin-bottom: 18px; background: rgba(6,9,17,0.6); padding: 10px 14px; border-radius: 8px; border: 1px solid var(--border-subtle);">
                Pravidelný odběr je automaticky odesílán každou středu v 7:00, každý pátek v 7:00 a každou neděli v 19:00 v kompletním českém překladu.
            </div>

            <div style="border-top: 1px solid var(--border-subtle); padding-top: 14px;">
                <div class="card-title" style="margin-bottom: 8px;">Aktuální tržní zprávy pro vybraný titul (přeloženo do češtiny)</div>
                <div id="newsFeedList">Načítám zprávy...</div>
            </div>
        </div>
    </div>

    <script>
        let advisorHistory = [];
        let methodsChartInstance = null;
        let monteCarloChartInstance = null;
        let cashFlowChartInstance = null;
        let currentResolvedTicker = 'NVDA';
        let currentCurrency = 'USD';
        let currentWeightedFairValue = null;

        let tvChart = null;
        let tvMainSeries = null;
        let tvVolumeSeries = null;
        let tvFairPriceLine = null;
        let tvChartStyle = 'area';
        let tvCurrentTimeframe = '1R';
        let tvLastPayload = null;

        function fmtMoney(val, curr = 'USD') {
            if (val === null || val === undefined || isNaN(val)) return 'N/A';
            return Number(val).toLocaleString('cs-CZ', {maximumFractionDigits: 2}) + ' ' + curr;
        }

        function fmtLarge(val, curr = 'USD') {
            if (val === null || val === undefined || isNaN(val)) return 'N/A';
            const abs = Math.abs(val);
            if (abs >= 1e12) return (val / 1e12).toFixed(2) + ' bil. ' + curr;
            if (abs >= 1e9) return (val / 1e9).toFixed(2) + ' mld. ' + curr;
            if (abs >= 1e6) return (val / 1e6).toFixed(2) + ' mil. ' + curr;
            return Number(val).toLocaleString('cs-CZ', {maximumFractionDigits: 0}) + ' ' + curr;
        }

        function handleSearchTrigger() {
            const typed = (document.getElementById('tickerInput').value || '').trim().toUpperCase();
            const isSameTicker = typed === currentResolvedTicker.toUpperCase();
            loadValuation(!isSameTicker);
        }

        function setTvChartStyle(style) {
            tvChartStyle = style;
            document.getElementById('btnChartArea').classList.toggle('active', style === 'area');
            document.getElementById('btnChartCandle').classList.toggle('active', style === 'candle');
            if (tvLastPayload) renderTvChartData(tvLastPayload);
        }

        function changeTvTimeframe(tf, btnEl) {
            tvCurrentTimeframe = tf;
            document.querySelectorAll('#tvTimeframeGroup .tv-tf-btn').forEach(b => b.classList.remove('active'));
            if (btnEl) btnEl.classList.add('active');
            loadInteractivePriceChart(currentResolvedTicker, tf);
        }

        async function loadInteractivePriceChart(ticker, timeframe = tvCurrentTimeframe) {
            try {
                const res = await fetch('/api/chart/' + encodeURIComponent(ticker) + '?timeframe=' + encodeURIComponent(timeframe));
                if (!res.ok) return;
                const payload = await res.json();
                tvLastPayload = payload;
                renderTvChartData(payload);
            } catch (err) {
                console.error('Chyba při načítání cenového grafu:', err);
            }
        }

        function renderTvChartData(payload) {
            const container = document.getElementById('tvChartContainer');
            if (!container || typeof LightweightCharts === 'undefined') return;

            const badge = document.getElementById('tvPeriodChangeBadge');
            const isUp = payload.change_pct >= 0;
            badge.textContent = `Změna (${payload.timeframe}): ${isUp ? '+' : ''}${payload.change_pct.toFixed(2)} % (${isUp ? '+' : ''}${payload.change_abs.toFixed(2)} ${currentCurrency})`;
            badge.style.color = isUp ? '#34d399' : '#f87171';
            badge.style.borderColor = isUp ? 'rgba(16,185,129,0.35)' : 'rgba(239,68,68,0.35)';

            if (tvChart) {
                tvChart.remove();
                tvChart = null;
            }

            tvChart = LightweightCharts.createChart(container, {
                width: container.clientWidth,
                height: 360,
                layout: {
                    background: { type: 'solid', color: '#070b14' },
                    textColor: '#94a3b8',
                    fontFamily: "'JetBrains Mono', monospace",
                },
                grid: {
                    vertLines: { color: 'rgba(255, 255, 255, 0.04)' },
                    horzLines: { color: 'rgba(255, 255, 255, 0.04)' },
                },
                crosshair: {
                    mode: LightweightCharts.CrosshairMode.Normal,
                    vertLine: { color: 'rgba(56, 189, 248, 0.45)', labelBackgroundColor: '#0284c7' },
                    horzLine: { color: 'rgba(56, 189, 248, 0.45)', labelBackgroundColor: '#0284c7' },
                },
                rightPriceScale: { borderColor: 'rgba(255, 255, 255, 0.08)' },
                timeScale: {
                    borderColor: 'rgba(255, 255, 255, 0.08)',
                    timeVisible: payload.intraday,
                    secondsVisible: false,
                },
            });

            if (tvChartStyle === 'candle') {
                tvMainSeries = tvChart.addCandlestickSeries({
                    upColor: '#10b981',
                    downColor: '#ef4444',
                    borderVisible: false,
                    wickUpColor: '#10b981',
                    wickDownColor: '#ef4444',
                });
                tvMainSeries.setData(payload.candles);
            } else {
                const lineCol = isUp ? '#10b981' : '#ef4444';
                const topFill = isUp ? 'rgba(16, 185, 129, 0.28)' : 'rgba(239, 68, 68, 0.28)';
                tvMainSeries = tvChart.addAreaSeries({
                    lineColor: lineCol,
                    topColor: topFill,
                    bottomColor: 'rgba(7, 11, 20, 0.0)',
                    lineWidth: 2,
                });
                tvMainSeries.setData(payload.candles.map(c => ({ time: c.time, value: c.close })));
            }

            tvVolumeSeries = tvChart.addHistogramSeries({
                priceFormat: { type: 'volume' },
                priceScaleId: '',
            });
            tvVolumeSeries.priceScale().applyOptions({ scaleMargins: { top: 0.82, bottom: 0 } });
            tvVolumeSeries.setData(payload.volumes);

            if (currentWeightedFairValue && currentWeightedFairValue > 0) {
                tvFairPriceLine = tvMainSeries.createPriceLine({
                    price: currentWeightedFairValue,
                    color: '#38bdf8',
                    lineWidth: 1,
                    lineStyle: LightweightCharts.LineStyle.Dashed,
                    axisLabelVisible: true,
                    title: 'Vnitřní hodnota',
                });
            }

            const readout = document.getElementById('tvCrosshairReadout');
            const candleByTime = new Map(payload.candles.map(c => [String(c.time), c]));
            tvChart.subscribeCrosshairMove(param => {
                if (!readout) return;
                if (!param || !param.time) {
                    readout.innerHTML = `<span>Poslední cena: <strong>${fmtMoney(payload.last_close, currentCurrency)}</strong> | Najeďte myší na graf pro detailní cenu a čas.</span>`;
                    return;
                }
                const key = String(param.time);
                const c = candleByTime.get(key);
                let timeLabel = key;
                if (typeof param.time === 'number') {
                    const dt = new Date(param.time * 1000);
                    timeLabel = dt.toLocaleString('cs-CZ', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
                }
                if (c) {
                    readout.innerHTML =
                        `<span>Čas: <strong>${timeLabel}</strong></span>` +
                        `<span>O: <strong>${c.open.toFixed(2)}</strong></span>` +
                        `<span>H: <strong>${c.high.toFixed(2)}</strong></span>` +
                        `<span>L: <strong>${c.low.toFixed(2)}</strong></span>` +
                        `<span>C: <strong>${c.close.toFixed(2)} ${currentCurrency}</strong></span>`;
                }
            });

            tvChart.timeScale().fitContent();
        }

        window.addEventListener('resize', () => {
            const container = document.getElementById('tvChartContainer');
            if (tvChart && container) {
                tvChart.applyOptions({ width: container.clientWidth });
            }
        });

        function updateCompanyIdentityBanner(d) {
            const curr = d.currency || 'USD';
            document.getElementById('companyFullName').textContent = d.company_name;
            document.getElementById('companyTickerBadge').textContent = d.ticker;
            document.getElementById('companySectorMeta').textContent =
                'Sektor: ' + d.sector + ' • Odvětví: ' + d.industry + ' • Burza: ' + d.exchange;
            document.getElementById('companyPricePill').textContent =
                'Tržní cena: ' + fmtMoney(d.dcf_valuation.current_share_price, curr);

            const logoContainer = document.getElementById('companyLogoContainer');
            const initials = (d.company_name || d.ticker || 'TT').substring(0, 2).toUpperCase();
            const candidates = Array.isArray(d.company_logo_urls) && d.company_logo_urls.length > 0
                ? [...d.company_logo_urls]
                : (d.company_logo_url ? [d.company_logo_url] : []);

            function tryNextLogo(idx) {
                if (idx >= candidates.length) {
                    logoContainer.className = 'company-logo-fallback';
                    logoContainer.textContent = initials;
                    return;
                }
                logoContainer.className = 'company-logo-frame';
                const img = document.createElement('img');
                img.src = candidates[idx];
                img.alt = d.company_name;
                img.onload = function() {
                    if (img.naturalWidth < 12 || img.naturalHeight < 12) tryNextLogo(idx + 1);
                };
                img.onerror = function() { tryNextLogo(idx + 1); };
                logoContainer.replaceChildren(img);
            }
            tryNextLogo(0);
        }

        function renderCharts(d) {
            if (typeof Chart === 'undefined') return;
            const curr = d.currency || 'USD';
            const gridColor = 'rgba(255, 255, 255, 0.06)';
            const tickColor = '#94a3b8';

            const ctx1 = document.getElementById('methodsChartCanvas').getContext('2d');
            if (methodsChartInstance) methodsChartInstance.destroy();
            const activeMethods = d.valuation_synthesis.methods.filter(m => m.weight_percent > 0);
            const methodLabels = ['Tržní cena', 'Vážená hodnota', ...activeMethods.map(m => m.method)];
            const methodValues = [
                d.dcf_valuation.current_share_price,
                d.valuation_synthesis.weighted_fair_value,
                ...activeMethods.map(m => m.fair_value)
            ];
            const barColors = methodValues.map((val, idx) => {
                if (idx === 0) return '#f59e0b';
                if (idx === 1) return '#38bdf8';
                return val >= d.dcf_valuation.current_share_price ? 'rgba(16, 185, 129, 0.75)' : 'rgba(239, 68, 68, 0.75)';
            });

            methodsChartInstance = new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: methodLabels,
                    datasets: [{ label: 'Hodnota (' + curr + ')', data: methodValues, backgroundColor: barColors, borderRadius: 6 }]
                },
                options: {
                    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { grid: { color: gridColor }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } }
                    }
                }
            });

            const ctx2 = document.getElementById('monteCarloChartCanvas').getContext('2d');
            if (monteCarloChartInstance) monteCarloChartInstance.destroy();
            const hist = d.monte_carlo.histogram || [];
            const currentP = d.dcf_valuation.current_share_price;
            monteCarloChartInstance = new Chart(ctx2, {
                type: 'bar',
                data: {
                    labels: hist.map(h => h.label),
                    datasets: [{
                        label: 'Počet simulací',
                        data: hist.map(h => h.count),
                        backgroundColor: hist.map(h => h.midpoint >= currentP ? 'rgba(56, 189, 248, 0.75)' : 'rgba(148, 163, 184, 0.45)'),
                        borderRadius: 4
                    }]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: { grid: { color: gridColor }, ticks: { color: tickColor, font: { size: 10 } } }
                    }
                }
            });

            const ctx3 = document.getElementById('cashFlowChartCanvas').getContext('2d');
            if (cashFlowChartInstance) cashFlowChartInstance.destroy();
            const cfRows = d.dcf_valuation.cash_flows || [];
            const lboRows = d.projektovane_roky || [];
            cashFlowChartInstance = new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: cfRows.map(r => 'Rok ' + r.rok),
                    datasets: [
                        { label: 'Volný peněžní tok', data: cfRows.map(r => r.volny_penezni_tok / 1e9), backgroundColor: 'rgba(16, 185, 129, 0.75)', borderRadius: 6 },
                        { label: 'Zůstatek dluhu (LBO)', data: lboRows.map(r => r.zbyvajici_cizi_kapital / 1e9), type: 'line', borderColor: '#38bdf8', backgroundColor: '#38bdf8', tension: 0.3 }
                    ]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: tickColor, font: { size: 10 } } } },
                    scales: {
                        x: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: { grid: { color: gridColor }, ticks: { color: tickColor, font: { size: 10 }, callback: v => v.toFixed(1) + ' mld.' } }
                    }
                }
            });
        }

        async function loadForexRates() {
            try {
                const res = await fetch('/api/forex-rates');
                const data = await res.json();
                const container = document.getElementById('forexRates');
                container.replaceChildren();
                Object.entries(data.rates).forEach(([pair, item]) => {
                    const span = document.createElement('span');
                    span.innerHTML = pair + ' <strong>' + Number(item.rate).toFixed(2) + '</strong>';
                    container.appendChild(span);
                });
            } catch (err) { console.error(err); }
        }

        async function loadValuation(useCompanyDefaults = true) {
            const rawQuery = (document.getElementById('tickerInput').value || 'NVDA').trim();
            const params = new URLSearchParams();

            if (!useCompanyDefaults) {
                const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
                const growth = document.getElementById('growthInput').valueAsNumber / 100;
                const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
                if (!isNaN(leverage)) params.set('leverage', leverage);
                if (!isNaN(growth)) params.set('growth', growth);
                if (!isNaN(exitMult)) params.set('exit_mult', exitMult);
            }

            const queryString = params.toString();
            const endpoint = '/api/valuation/' + encodeURIComponent(rawQuery) + (queryString ? '?' + queryString : '');

            try {
                const res = await fetch(endpoint);
                const d = await res.json();
                if (!res.ok) throw new Error(d.detail || 'Valuaci se nepodařilo načíst.');

                const curr = d.currency || 'USD';
                const prevTicker = currentResolvedTicker;
                currentResolvedTicker = d.ticker;
                currentCurrency = curr;
                currentWeightedFairValue = d.valuation_synthesis.weighted_fair_value;

                document.getElementById('tickerInput').value = d.ticker;
                document.getElementById('fundTicker').textContent = d.ticker;
                document.getElementById('auditTimestamp').textContent = d.audit_trail.model_version + ' | ' + d.audit_trail.timestamp;

                if (useCompanyDefaults && d.dynamic_defaults) {
                    document.getElementById('leverageInput').value = d.dynamic_defaults.leverage_percent;
                    document.getElementById('growthInput').value = d.dynamic_defaults.growth_percent;
                    document.getElementById('exitMultipleInput').value = d.dynamic_defaults.exit_multiple;
                }

                updateCompanyIdentityBanner(d);

                if (useCompanyDefaults || prevTicker !== d.ticker || !tvLastPayload) {
                    loadInteractivePriceChart(d.ticker, tvCurrentTimeframe);
                } else if (tvLastPayload) {
                    renderTvChartData(tvLastPayload);
                }

                const dec = d.decision_engine;
                const synth = d.valuation_synthesis;
                document.getElementById('integrityBadge').textContent = 'Index integrity: ' + synth.integrity_score + ' / 100';
                document.getElementById('weightedFairBadge').textContent = 'Vážená vnitřní hodnota: ' + fmtMoney(synth.weighted_fair_value, curr);

                ['lightGreen', 'lightYellow', 'lightOrange', 'lightRed'].forEach(id => document.getElementById(id).classList.remove('active'));
                const lightMap = { GREEN: 'lightGreen', YELLOW: 'lightYellow', ORANGE: 'lightOrange', RED: 'lightRed' };
                if (lightMap[dec.zone]) document.getElementById(lightMap[dec.zone]).classList.add('active');

                document.getElementById('thesisRecommendation').textContent = dec.recommendation + ' (Doporučená tranše: ' + dec.tranche_allocation_percent + ' %)';
                document.getElementById('thesisSummary').textContent =
                    'Aktuální tržní cena: ' + fmtMoney(d.dcf_valuation.current_share_price, curr) +
                    ' | Vážená vnitřní hodnota: ' + fmtMoney(synth.weighted_fair_value, curr) +
                    ' | Bezpečnostní rezerva: ' + dec.margin_of_safety_percent.toFixed(2) + ' %.';

                const b = dec.bands;
                document.getElementById('zoneGreenRange').textContent = '0,00 – ' + fmtMoney(b.green_max, curr);
                document.getElementById('zoneYellowRange').textContent = fmtMoney(b.yellow_min, curr) + ' – ' + fmtMoney(b.yellow_max, curr);
                document.getElementById('zoneOrangeRange').textContent = fmtMoney(b.orange_min, curr) + ' – ' + fmtMoney(b.orange_max, curr);
                document.getElementById('zoneRedRange').textContent = 'nad ' + fmtMoney(b.red_min, curr);

                let synHtml = '<table><thead><tr><th>Oceňovací metoda</th><th>Váha</th><th>Vnitřní hodnota / akcii</th><th>Vážený příspěvek</th><th>Rozdíl vs. tržní cena</th></tr></thead><tbody>';
                synth.methods.forEach(m => {
                    const col = m.vs_market_percent >= 0 ? '#34d399' : '#f87171';
                    synHtml += `<tr><td>${m.method}</td><td>${m.weight_percent.toFixed(1)} %</td><td>${fmtMoney(m.fair_value, curr)}</td><td>${fmtMoney(m.contribution, curr)}</td><td style="color:${col};">${m.vs_market_percent >= 0 ? '+' : ''}${m.vs_market_percent.toFixed(2)} %</td></tr>`;
                });
                synHtml += '</tbody></table>';
                document.getElementById('synthesisTableContainer').innerHTML = synHtml;
                document.getElementById('disagreementNote').textContent = synth.disagreement_note;

                renderCharts(d);

                const f = d.fundamentals;
                document.getElementById('valCap').textContent = fmtLarge(d.trzni_kapitalizace, curr);
                document.getElementById('valEV').textContent = fmtLarge(d.celkova_hodnota_podniku, curr);
                document.getElementById('normFcf').textContent = fmtLarge(f.normalized_fcf, curr);
                document.getElementById('roicVsWacc').textContent = f.roic_percent.toFixed(2) + ' % vs. ' + d.wacc_engine.wacc_percent.toFixed(2) + ' %';
                document.getElementById('econProfit').textContent = fmtLarge(f.economic_profit, curr);
                document.getElementById('riskScores').textContent = 'Piotroski ' + d.risk_analysis.piotroski_f_score + '/9 | ' + d.risk_analysis.altman_status;

                const dcf = d.dcf_valuation;
                document.getElementById('dcfIntrinsicValue').textContent = fmtMoney(dcf.intrinsic_value_per_share, curr);
                document.getElementById('fcfeAndRi').textContent = fmtMoney(d.fcfe_valuation.intrinsic_value_per_share, curr) + ' / ' + fmtMoney(d.residual_income.intrinsic_value_per_share, curr);
                document.getElementById('dcfWacc').textContent = d.wacc_engine.wacc_percent.toFixed(2) + ' % (Beta ' + d.wacc_engine.beta.toFixed(2) + ')';
                document.getElementById('tvShare').textContent = dcf.tv_share_percent.toFixed(1) + ' % z EV';
                document.getElementById('reverseDcfGrowth').textContent = d.reverse_dcf.implied_fcf_growth_percent.toFixed(2) + ' % ročně';
                document.getElementById('sotpNav').textContent = fmtMoney(d.metoda_souctu_casti.sotp_per_share, curr) + ' / ' + fmtMoney(d.metoda_souctu_casti.nav_per_share, curr);

                const sc = d.scenarios;
                const mc = d.monte_carlo;
                document.getElementById('bearBaseBull').textContent = sc.bear.fair_value.toFixed(1) + ' / ' + sc.base.fair_value.toFixed(1) + ' / ' + sc.bull.fair_value.toFixed(1);
                document.getElementById('scenExpected').textContent = fmtMoney(sc.expected_value, curr);
                document.getElementById('mcP50').textContent = fmtMoney(mc.p50, curr);
                document.getElementById('mcP10P90').textContent = fmtMoney(mc.p10, curr) + ' – ' + fmtMoney(mc.p90, curr);
                document.getElementById('mcP5P95').textContent = fmtMoney(mc.p5, curr) + ' – ' + fmtMoney(mc.p95, curr);
                document.getElementById('mcProbUpside').textContent = mc.prob_above_market_percent.toFixed(1) + ' %';

                document.getElementById('peerSector').textContent = d.peer_comps.sector;
                let peerHtml = '<table><thead><tr><th>Metoda</th><th>Násobek</th><th>Metrika (PV)</th><th>Implikovaná cena</th></tr></thead><tbody>';
                d.peer_comps.multiples.forEach(p => {
                    peerHtml += `<tr><td>${p.method}</td><td>${p.multiple}x</td><td>${fmtLarge(p.forward_metric, curr)}</td><td>${fmtMoney(p.implied_share_price, curr)}</td></tr>`;
                });
                peerHtml += '</tbody></table>';
                document.getElementById('peerCompsContainer').innerHTML = peerHtml;

                const lbo = d.model_odkupu_za_cizi_kapital_base_case;
                document.getElementById('lboSummaryBadge').textContent = 'Násobek ' + lbo.moic + 'x | Výnos ' + lbo.irr_procent + ' %';
                let schedHtml = '<table><thead><tr><th>Rok</th><th>Provozní zisk</th><th>Volný tok</th><th>Splátka dluhu</th><th>Zůstatek dluhu</th></tr></thead><tbody>';
                d.projektovane_roky.forEach(y => {
                    schedHtml += `<tr><td>Rok ${y.rok}</td><td>${fmtLarge(y.projektovany_provozni_zisk_pred_odpisy, curr)}</td><td>${fmtLarge(y.volny_tok_pro_splaceni, curr)}</td><td>${fmtLarge(y.mor_jistiny, curr)}</td><td>${fmtLarge(y.zbyvajici_cizi_kapital, curr)}</td></tr>`;
                });
                schedHtml += '</tbody></table>';
                document.getElementById('lboScheduleContainer').innerHTML = schedHtml;

                const dcfMat = dcf.sensitivity_matrix || [];
                if (dcfMat.length > 0) {
                    let dMatHtml = '<table><thead><tr><th>Náklady kapitálu \\ Terminální růst</th>' + dcfMat[0].cells.map(c => `<th>${c.tg.toFixed(2)} %</th>`).join('') + '</tr></thead><tbody>';
                    dcfMat.forEach(r => {
                        dMatHtml += `<tr><td><strong>${r.wacc.toFixed(2)} %</strong></td>`;
                        r.cells.forEach(c => {
                            const col = c.price >= dcf.current_share_price ? '#34d399' : '#f87171';
                            dMatHtml += `<td style="color:${col};">${c.price.toFixed(2)}</td>`;
                        });
                        dMatHtml += '</tr>';
                    });
                    dMatHtml += '</tbody></table>';
                    document.getElementById('dcfSensitivityContainer').innerHTML = dMatHtml;
                }

                const lboMat = d.citlivostni_matice_lbo_irr || [];
                if (lboMat.length > 0) {
                    let lMatHtml = '<table><thead><tr><th>Růst \\ Výstupní násobek</th>' + lboMat[0].vysledky_irr_dle_nasobku.map(m => `<th>${m.vystupni_nasobek.toFixed(1)}x</th>`).join('') + '</tr></thead><tbody>';
                    lboMat.forEach(r => {
                        lMatHtml += `<tr><td><strong>${r.tempo_rustu_ebitda_procent.toFixed(1)} %</strong></td>`;
                        r.vysledky_irr_dle_nasobku.forEach(c => {
                            const col = c.irr_procent >= 0 ? '#34d399' : '#f87171';
                            lMatHtml += `<td style="color:${col};">${c.irr_procent.toFixed(2)} %</td>`;
                        });
                        lMatHtml += '</tr>';
                    });
                    lMatHtml += '</tbody></table>';
                    document.getElementById('sensitivityMatrix').innerHTML = lMatHtml;
                }
            } catch (err) {
                document.getElementById('thesisRecommendation').textContent = 'Chyba při načítání valuace: ' + err.message;
                console.error(err);
            }
        }

        function openNewsletterModal() {
            const ticker = currentResolvedTicker || (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            document.getElementById('newsModalTicker').textContent = ticker;
            document.getElementById('subscribeStatus').textContent = '';
            document.getElementById('newsletterModal').classList.add('open');
            loadTickerNews(ticker);
        }

        function closeNewsletterModal() {
            document.getElementById('newsletterModal').classList.remove('open');
        }

        async function loadTickerNews(ticker) {
            const listEl = document.getElementById('newsFeedList');
            listEl.innerHTML = '<div style="text-align:center; padding:12px;"><div class="loader"></div></div>';
            try {
                const res = await fetch('/api/news/' + encodeURIComponent(ticker));
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Zprávy se nepodařilo načíst.');
                const articles = data.articles || [];
                if (articles.length === 0) {
                    listEl.innerHTML = '<p style="color:var(--text-muted); font-size:12.5px;">Žádné aktuální zprávy nenalezeny.</p>';
                    return;
                }
                listEl.innerHTML = articles.map(a => `
                    <div class="news-item">
                        <a href="${a.url}" target="_blank" rel="noopener">${a.title} &#8599;</a>
                        <div class="news-meta">${a.source}</div>
                    </div>
                `).join('');
            } catch (err) {
                listEl.innerHTML = `<p style="color:#f87171; font-size:12.5px;">${err.message}</p>`;
            }
        }

        async function sendNewsletterNow() {
            const email = document.getElementById('subscribeEmailInput').value.trim();
            const st = document.getElementById('subscribeStatus');
            if (!email) {
                st.style.color = '#f87171';
                st.textContent = 'Zadejte prosím platnou e-mailovou adresu.';
                return;
            }
            st.style.color = '#38bdf8';
            st.textContent = 'Generuji kompletní český přehled a odesílám e-mail...';
            try {
                const res = await fetch('/api/newsletter/send-now', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Odeslání selhalo.');
                st.style.color = '#34d399';
                st.textContent = data.message;
            } catch (err) {
                st.style.color = '#f87171';
                st.textContent = err.message;
            }
        }

        async function subscribeNewsletter(e) {
            if (e) e.preventDefault();
            const email = document.getElementById('subscribeEmailInput').value.trim();
            const st = document.getElementById('subscribeStatus');
            if (!email) {
                st.style.color = '#f87171';
                st.textContent = 'Zadejte prosím platnou e-mailovou adresu.';
                return;
            }
            try {
                const res = await fetch('/subscribe', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Registrace odběru selhala.');
                st.style.color = '#34d399';
                st.textContent = data.message;
            } catch (err) {
                st.style.color = '#f87171';
                st.textContent = err.message;
            }
        }

        function renderMarkdownToHtml(mdText) {
            var rawStr = String(mdText || '');
            var lines = rawStr.split('\n');
            var htmlOut = '';
            var inList = false;
            var inTable = false;

            function escapeHtmlText(txt) {
                return String(txt || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;');
            }

            function formatInline(str) {
                var safe = escapeHtmlText(str);
                safe = safe.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
                safe = safe.replace(/`([^`]+)`/g, '<code style="background:rgba(56,189,248,0.12);padding:1px 5px;border-radius:4px;font-size:12px;">$1</code>');
                return safe;
            }

            function isTableDivider(lineStr) {
                var trimmed = lineStr.replace(/[|\-:\s]/g, '');
                return trimmed.length === 0;
            }

            for (var i = 0; i < lines.length; i++) {
                var rawLine = lines[i].trim();

                if (rawLine.length > 1 && rawLine.charAt(0) === '|' && rawLine.charAt(rawLine.length - 1) === '|') {
                    if (inList) {
                        htmlOut += '</ul>';
                        inList = false;
                    }
                    if (isTableDivider(rawLine)) {
                        continue;
                    }
                    var innerLine = rawLine.substring(1, rawLine.length - 1);
                    var cells = innerLine.split('|');
                    if (!inTable) {
                        inTable = true;
                        htmlOut += '<table><thead><tr>';
                        for (var cIdx = 0; cIdx < cells.length; cIdx++) {
                            htmlOut += '<th>' + formatInline(cells[cIdx].trim()) + '</th>';
                        }
                        htmlOut += '</tr></thead><tbody>';
                    } else {
                        htmlOut += '<tr>';
                        for (var rIdx = 0; rIdx < cells.length; rIdx++) {
                            htmlOut += '<td>' + formatInline(cells[rIdx].trim()) + '</td>';
                        }
                        htmlOut += '</tr>';
                    }
                    continue;
                } else if (inTable) {
                    htmlOut += '</tbody></table>';
                    inTable = false;
                }

                if (rawLine.indexOf('### ') === 0) {
                    if (inList) {
                        htmlOut += '</ul>';
                        inList = false;
                    }
                    htmlOut += '<h4>' + formatInline(rawLine.substring(4)) + '</h4>';
                } else if (rawLine.indexOf('## ') === 0) {
                    if (inList) {
                        htmlOut += '</ul>';
                        inList = false;
                    }
                    htmlOut += '<h3>' + formatInline(rawLine.substring(3)) + '</h3>';
                } else if (rawLine.indexOf('- ') === 0 || rawLine.indexOf('* ') === 0) {
                    if (!inList) {
                        htmlOut += '<ul>';
                        inList = true;
                    }
                    htmlOut += '<li>' + formatInline(rawLine.substring(2)) + '</li>';
                } else if (rawLine === '') {
                    if (inList) {
                        htmlOut += '</ul>';
                        inList = false;
                    }
                } else {
                    if (inList) {
                        htmlOut += '</ul>';
                        inList = false;
                    }
                    htmlOut += '<p>' + formatInline(rawLine) + '</p>';
                }
            }
            if (inList) {
                htmlOut += '</ul>';
            }
            if (inTable) {
                htmlOut += '</tbody></table>';
            }
            return htmlOut || ('<p>' + escapeHtmlText(rawStr) + '</p>');
        }

        async function askAdvisor(e) {
            e.preventDefault();
            const input = document.getElementById('advisorQuestion');
            const submitBtn = document.getElementById('advisorSubmitBtn');
            const q = input.value.trim();
            if (!q) return;

            const thread = document.getElementById('advisorThread');
            const userRow = document.createElement('div');
            userRow.className = 'advisor-msg-row user-row';
            userRow.innerHTML = `<div class="advisor-avatar">VY</div><div class="advisor-bubble"><p>${q.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</p></div>`;
            thread.appendChild(userRow);
            input.value = '';
            submitBtn.disabled = true;

            const aiRow = document.createElement('div');
            aiRow.className = 'advisor-msg-row';
            aiRow.innerHTML = `<div class="advisor-avatar">AI</div><div class="advisor-bubble"><p style="color:#94a3b8;">Analyzuji data a připravuji odpověď...</p></div>`;
            thread.appendChild(aiRow);
            thread.scrollTop = thread.scrollHeight;

            const bubbleEl = aiRow.querySelector('.advisor-bubble');
            const ticker = currentResolvedTicker || document.getElementById('tickerInput').value.trim() || 'NVDA';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;

            try {
                const res = await fetch('/api/advisor/' + encodeURIComponent(ticker), {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        question: q,
                        history: advisorHistory,
                        scenario: { leverage, growth, exit_mult: exitMult }
                    })
                });
                const data = await res.json();
                const replyText = res.ok ? data.answer : (data.detail || 'Nepodařilo se zpracovat dotaz.');
                bubbleEl.innerHTML = renderMarkdownToHtml(replyText);
                thread.scrollTop = thread.scrollHeight;

                if (res.ok) {
                    advisorHistory.push({ role: 'user', text: q });
                    advisorHistory.push({ role: 'advisor', text: replyText });
                    if (advisorHistory.length > 40) advisorHistory = advisorHistory.slice(-40);
                }
            } catch (err) {
                bubbleEl.innerHTML = `<p>Nepodařilo se spojit s AI poradcem: ${err.message}</p>`;
            } finally {
                submitBtn.disabled = false;
            }
        }

        function downloadExcel() {
            const ticker = currentResolvedTicker || document.getElementById('tickerInput').value.trim() || 'NVDA';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/excel/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        function downloadPDF() {
            const ticker = currentResolvedTicker || document.getElementById('tickerInput').value.trim() || 'NVDA';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/pdf/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        document.addEventListener('DOMContentLoaded', () => {
            try { loadForexRates(); } catch (e) { console.error(e); }
            try { loadValuation(true); } catch (e) { console.error(e); }
        });
    </script>
</body>
</html>
"""