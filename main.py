# -*- coding: utf-8 -*-
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr
import asyncio
import io
import json
import logging
import math
import os
import random
import re
import smtplib
import socket
import ssl
import urllib.request

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import yfinance as yf

try:
    from zoneinfo import ZoneInfo
except ImportError:
    ZoneInfo = None

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tensegrity_terminal")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)

LOGO_PATH = os.path.join(STATIC_DIR, "logo.png")
SUBSCRIBERS_FILE = os.path.join(BASE_DIR, "subscribers.json")
SCHEDULER_STATE_FILE = os.path.join(BASE_DIR, "scheduler_state.json")

SENDER_EMAIL = os.getenv("SENDER_EMAIL", "kubenkamatyas@gmail.com").strip()
RAW_SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "yokf zefs doxl ujci")
SENDER_PASSWORD = "".join(RAW_SENDER_PASSWORD.split())
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

FOREX_PAIRS = {
    "EUR/CZK": "EURCZK=X",
    "USD/CZK": "USDCZK=X",
    "GBP/CZK": "GBPCZK=X",
}
FOREX_FALLBACK_RATES = {"EUR/CZK": 25.24, "USD/CZK": 23.15, "GBP/CZK": 30.12}


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


def safe_float(value, default: float = 0.0) -> float:
    try:
        if value is None:
            return default
        num = float(value)
        if math.isnan(num) or math.isinf(num):
            return default
        return num
    except (TypeError, ValueError):
        return default


def extract_statement_series(df, row_names: list[str], max_years: int = 4) -> list[float]:
    if df is None or getattr(df, "empty", True):
        return []
    for name in row_names:
        if name in df.index:
            series = df.loc[name].dropna()
            values = [safe_float(v) for v in series.iloc[:max_years].tolist()]
            if values:
                return values
    return []


def get_live_market_benchmarks() -> dict:
    import re
    import urllib.request

    us10y = None
    sp500_pe = None

    try:
        tnx_hist = yf.Ticker("^TNX").history(period="5d", interval="1d")
        if not tnx_hist.empty:
            last_tnx = safe_float(tnx_hist["Close"].dropna().iloc[-1], 0.0)
            if 0.5 < last_tnx < 15.0:
                us10y = round(last_tnx, 2)
    except Exception as err:
        logger.warning("Nelze načíst burzovní výnos desetiletých amerických dluhopisů: %s", err)

    if us10y is None:
        try:
            req = urllib.request.Request(
                "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10",
                headers={"User-Agent": "Mozilla/5.0"},
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                lines = resp.read().decode("utf-8", errors="ignore").strip().splitlines()
            for line in reversed(lines[1:]):
                parts = line.split(",")
                if len(parts) == 2 and parts[1].strip() != ".":
                    candidate = safe_float(parts[1].strip(), 0.0)
                    if 0.5 < candidate < 15.0:
                        us10y = round(candidate, 2)
                        break
        except Exception as err:
            logger.warning("Nelze načíst záložní živý výnos z Federální rezervní banky: %s", err)

    try:
        spy_info = yf.Ticker("SPY").info or {}
        trailing_pe = safe_float(spy_info.get("trailingPE"), 0.0)
        if 8.0 < trailing_pe < 60.0:
            sp500_pe = round(trailing_pe, 1)
    except Exception as err:
        logger.warning("Nelze načíst burzovní poměr ceny a zisku indexu Standard and Poor's 500: %s", err)

    if sp500_pe is None:
        try:
            req_pe = urllib.request.Request(
                "https://www.multpl.com/s-p-500-pe-ratio",
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
            )
            with urllib.request.urlopen(req_pe, timeout=8) as resp_pe:
                page_text = re.sub(r"<[^>]+>", " ", resp_pe.read().decode("utf-8", errors="ignore"))
                page_text = re.sub(r"\s+", " ", page_text)
            pe_match = re.search(r"Current\s+S&P\s+500\s+PE\s+Ratio\s+is\s+(\d+\.\d+)", page_text, flags=re.IGNORECASE)
            if pe_match:
                candidate_pe = safe_float(pe_match.group(1), 0.0)
                if 8.0 < candidate_pe < 60.0:
                    sp500_pe = round(candidate_pe, 1)
        except Exception as err:
            logger.warning("Nelze načíst záložní živý poměr ceny a zisku: %s", err)

    return {
        "us10y_yield": us10y if us10y is not None else 0.0,
        "sp500_pe": sp500_pe if sp500_pe is not None else 0.0,
    }


_CNB_CACHE: dict = {"expires_at": None, "data": None}


def _format_cz_percent(val: float, decimals: int = 2) -> str:
    return f"{val:.{decimals}f}".replace(".", ",") + " %"


def _extract_html_tables(raw_html: str) -> list[list[list[str]]]:
    from html.parser import HTMLParser
    import re

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
    from datetime import datetime, timedelta, timezone
    import html
    import re
    import urllib.request

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
        with urllib.request.urlopen(req, timeout=8) as resp:
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
        with urllib.request.urlopen(req_home, timeout=8) as resp_home:
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
        "outlook": (
            "Bankovní rada České národní banky průběžně vyhodnocuje proinflační rizika ve službách, mzdové dynamice "
            "a vývoji cen nemovitostí při nastavování měnových podmínek. "
            "Další kroky centrální banky se odvíjejí od nových makroekonomických prognóz a vývoje jádrové inflace."
        ),
    }
    _CNB_CACHE["data"] = result
    _CNB_CACHE["expires_at"] = now + timedelta(minutes=30)
    return result


def get_live_forex_rates() -> dict:
    import urllib.request

    cnb_daily_rates = {}
    try:
        req = urllib.request.Request(
            "https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            lines = resp.read().decode("utf-8", errors="ignore").splitlines()
        for line in lines[2:]:
            parts = line.split("|")
            if len(parts) == 5:
                currency_code = parts[3].strip().upper()
                amount = safe_float(parts[2].strip().replace(",", "."), 0.0)
                czk_value = safe_float(parts[4].strip().replace(",", "."), 0.0)
                if amount > 0 and czk_value > 0:
                    cnb_daily_rates[f"{currency_code}/CZK"] = round(czk_value / amount, 4)
    except Exception as error:
        logger.warning("Nelze načíst denní kurzovní lístek České národní banky: %s", error)

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
        except Exception as error:
            logger.warning("Nelze načíst živý burzovní kurz pro %s: %s", pair, error)

        rates[pair] = {"rate": rate, "live": live}

    return {
        "rates": rates,
        "updated_at": get_prague_now().isoformat(timespec="minutes"),
    }


def get_subscribers_internal() -> list[str]:
    if not os.path.exists(SUBSCRIBERS_FILE):
        return []
    try:
        with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, list):
            return []
        return [item.strip() for item in data if isinstance(item, str) and is_valid_email_address(item.strip())]
    except (OSError, json.JSONDecodeError):
        return []


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


def is_valid_email_address(email: str) -> bool:
    if not isinstance(email, str) or len(email) > 254:
        return False
    parsed = parseaddr(email)
    return parsed[1] == email and bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email))


def load_sent_slots() -> set[str]:
    if not os.path.exists(SCHEDULER_STATE_FILE):
        return set()
    try:
        with open(SCHEDULER_STATE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if isinstance(data, list):
            return set(str(x) for x in data[-100:])
    except Exception:
        pass
    return set()


def save_sent_slots(slots: set[str]):
    try:
        tmp = SCHEDULER_STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as file:
            json.dump(sorted(list(slots))[-100:], file, ensure_ascii=False, indent=2)
            file.flush()
            os.fsync(file.fileno())
        os.replace(tmp, SCHEDULER_STATE_FILE)
    except Exception as err:
        logger.warning("Nelze uložit stav plánovače: %s", err)


def _fetch_live_rss_articles(
    feed_sources: list[tuple[str, str]],
    badge_color: str,
    max_items: int = 5,
) -> list[dict]:
    from email.utils import parsedate_to_datetime
    import html
    import re
    import urllib.request
    import xml.etree.ElementTree as ET

    collected: list[dict] = []
    seen_titles: set[str] = set()

    for source_name, feed_url in feed_sources:
        if len(collected) >= max_items:
            break
        try:
            req = urllib.request.Request(
                feed_url,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "application/rss+xml, application/xml, text/xml;q=0.9, */*;q=0.8",
                },
            )
            with urllib.request.urlopen(req, timeout=7) as resp:
                raw_xml = resp.read()
            root = ET.fromstring(raw_xml)

            items = root.findall(".//item")
            if not items:
                items = root.findall(".//{http://www.w3.org/2005/Atom}entry")

            per_source_count = 0
            for entry in items:
                if len(collected) >= max_items or per_source_count >= 2:
                    break

                title_el = entry.find("title")
                if title_el is None:
                    title_el = entry.find("{http://www.w3.org/2005/Atom}title")
                raw_title = html.unescape(title_el.text or "").strip() if title_el is not None else ""
                if not raw_title:
                    continue

                norm_key = re.sub(r"\W+", "", raw_title.lower())[:60]
                if norm_key in seen_titles:
                    continue

                link_el = entry.find("link")
                if link_el is None:
                    link_el = entry.find("{http://www.w3.org/2005/Atom}link")
                article_url = ""
                if link_el is not None:
                    article_url = (link_el.text or link_el.attrib.get("href") or "").strip()
                if not article_url:
                    article_url = feed_url

                desc_el = entry.find("description")
                if desc_el is None:
                    desc_el = entry.find("{http://www.w3.org/2005/Atom}summary")
                raw_desc = html.unescape(desc_el.text or "") if desc_el is not None else ""
                clean_desc = re.sub(r"<[^>]+>", " ", raw_desc)
                clean_desc = re.sub(r"\s+", " ", clean_desc).strip()

                pub_el = entry.find("pubDate")
                if pub_el is None:
                    pub_el = entry.find("{http://www.w3.org/2005/Atom}updated")
                pub_label = ""
                if pub_el is not None and pub_el.text:
                    try:
                        dt = parsedate_to_datetime(pub_el.text.strip())
                        pub_label = dt.strftime("%d. %m. %Y %H:%M")
                    except Exception:
                        pub_label = pub_el.text.strip()[:25]

                if len(clean_desc) > 460:
                    clean_desc = clean_desc[:457].rsplit(" ", 1)[0] + "..."

                meta_prefix = f"Zdroj: {source_name}" + (f" ({pub_label})" if pub_label else "")
                if clean_desc and clean_desc.lower() not in raw_title.lower():
                    body_text = f"{clean_desc} [{meta_prefix}]"
                else:
                    body_text = f"{meta_prefix}. Aktuální tržní zpráva převzatá přímo z oficiálního zpravodajského kanálu."

                seen_titles.add(norm_key)
                per_source_count += 1
                collected.append(
                    {
                        "badge": f"{source_name.upper()} | ŽIVÉ ZPRÁVY",
                        "badge_color": badge_color,
                        "title": f"{len(collected) + 1}. {raw_title}",
                        "body": body_text,
                        "url": article_url,
                    }
                )
        except Exception as err:
            logger.warning("Nelze načíst živý zpravodajský kanál %s: %s", source_name, err)

    return collected


def build_curated_czech_news_sections(forex_data: dict, benchmarks: dict, cnb: dict) -> tuple[list[dict], list[dict]]:
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]
    us10y = benchmarks["us10y_yield"]
    sp500_pe = benchmarks["sp500_pe"]

    world_feeds = [
        ("Americká centrální banka", "https://www.federalreserve.gov/feeds/press_monetary.xml"),
        ("CNBC Finance", "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=10000664"),
        ("MarketWatch", "https://feeds.content.dowjones.io/public/rss/mw_topstories"),
        ("Yahoo Finance", "https://finance.yahoo.com/news/rssindex"),
    ]
    world_sections = _fetch_live_rss_articles(world_feeds, badge_color="#38bdf8", max_items=5)

    if len(world_sections) < 5:
        live_tickers = [
            ("INDEX STANDARD AND POOR'S 500", "^GSPC", "https://finance.yahoo.com/quote/%5EGSPC"),
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
                    world_sections.append(
                        {
                            "badge": f"{label} | ŽIVÁ TRŽNÍ DATA",
                            "badge_color": "#38bdf8",
                            "title": f"{len(world_sections) + 1}. {label}: aktuální hodnota {last_px:,.2f} ({chg_pct:+.2f} %)",
                            "body": (
                                f"Živý burzovní údaj pro instrument {label} činí {last_px:,.2f} při denní změně {chg_pct:+.2f} %. "
                                f"Bezrizikový výnos desetiletých amerických státních dluhopisů dosahuje {us10y:.2f} % "
                                f"a ocenění indexu Standard and Poor's 500 činí {sp500_pe:.1f}násobek ročních zisků."
                            ),
                            "url": link,
                        }
                    )
            except Exception:
                pass

    def _fetch_direct_portal_articles() -> list[dict]:
        import html as html_lib
        import re
        from urllib.parse import urljoin
        import urllib.request

        portals = [
            (
                "Česká národní banka",
                "https://www.cnb.cz/cs/cnb-news/tiskove-zpravy/",
                r"/(?:cnb-news|verejnost/servis-pro-media)/(?:tiskove-zpravy|aktuality)/[^\"'#?]+",
            ),
            (
                "Kurzy.cz",
                "https://www.kurzy.cz/zpravy/",
                r"/zpravy/\d+-[^\"'#?]+",
            ),
        ]
        portal_items: list[dict] = []
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept-Language": "cs-CZ,cs;q=0.9",
        }

        for source_name, page_url, link_pattern in portals:
            try:
                req = urllib.request.Request(page_url, headers=headers)
                with urllib.request.urlopen(req, timeout=7) as resp:
                    raw_html = resp.read().decode("utf-8", errors="ignore")

                for match in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", raw_html, flags=re.IGNORECASE | re.DOTALL):
                    href = match.group(1).strip()
                    if not re.search(link_pattern, href, flags=re.IGNORECASE):
                        continue
                    clean_title = re.sub(r"<[^>]+>", " ", html_lib.unescape(match.group(2)))
                    clean_title = re.sub(r"\s+", " ", clean_title).strip()
                    if len(clean_title) < 28 or clean_title.lower().startswith(("více", "archiv", "všechny")):
                        continue

                    full_url = urljoin(page_url, href)
                    portal_items.append(
                        {
                            "badge": f"{source_name.upper()} | ŽIVÝ ZDROJ",
                            "badge_color": "#f59e0b",
                            "title": clean_title,
                            "body": (
                                f"Aktuální zpráva převzatá přímo z oficiálního portálu {source_name}. "
                                f"Kliknutím na tlačítko níže otevřete celé původní znění na zdrojovém webu."
                            ),
                            "url": full_url,
                        }
                    )
                    break
            except Exception as err:
                logger.warning("Nelze načíst přímý odkaz z portálu %s: %s", source_name, err)

        return portal_items

    direct_cz_articles = _fetch_direct_portal_articles()

    czech_feeds = [
        ("Česká tisková kancelář", "https://www.ceskenoviny.cz/sluzby/rss/ekonomika.php"),
        ("Deník E15", "https://www.e15.cz/rss"),
        ("Seznam Zprávy", "https://www.seznamzpravy.cz/rss"),
    ]
    rss_cz_articles = _fetch_live_rss_articles(czech_feeds, badge_color="#f59e0b", max_items=5)

    czech_sections: list[dict] = []
    seen_cz: set[str] = set()
    for item in direct_cz_articles + rss_cz_articles:
        if len(czech_sections) >= 5:
            break
        raw_heading = re.sub(r"^\d+\.\s*", "", item["title"]).strip()
        dedup_key = re.sub(r"\W+", "", raw_heading.lower())[:55]
        if dedup_key in seen_cz:
            continue
        seen_cz.add(dedup_key)
        item["title"] = f"{len(czech_sections) + 1}. {raw_heading}"
        czech_sections.append(item)

    if len(czech_sections) < 5:
        live_cz_cards = [
            {
                "badge": "ČESKÁ NÁRODNÍ BANKA | ŽIVÉ SAZBY",
                "badge_color": "#f59e0b",
                "title": f"Česká národní banka drží základní repo sazbu na {cnb['repo_rate']} při inflaci {cnb['latest_inflation']}",
                "body": (
                    f"Aktuální základní dvoutýdenní repo sazba činí {cnb['repo_rate']}, diskontní sazba {cnb['discount_rate']} "
                    f"a lombardní sazba {cnb['lombard_rate']}. Meziroční míra inflace v České republice dosahuje {cnb['latest_inflation']}."
                ),
                "url": "https://www.cnb.cz/cs/menova-politika/",
            },
            {
                "badge": "KURZY.CZ | DEVIZOVÝ TRH",
                "badge_color": "#f59e0b",
                "title": f"Aktuální kurz české koruny: {eur_czk:.2f} za euro a {usd_czk:.2f} za americký dolar",
                "body": (
                    f"Na devizovém trhu se česká koruna obchoduje za {eur_czk:.2f} koruny za jedno euro, "
                    f"{usd_czk:.2f} koruny za jeden americký dolar a {gbp_czk:.2f} koruny za jednu britskou libru."
                ),
                "url": "https://www.kurzy.cz/kurzy-men/",
            },
        ]
        for fallback_card in live_cz_cards:
            if len(czech_sections) >= 5:
                break
            fallback_card["title"] = f"{len(czech_sections) + 1}. {fallback_card['title']}"
            czech_sections.append(fallback_card)

    return world_sections, czech_sections


def build_newsletter_pdf_bytes(forex_data: dict, benchmarks: dict, cnb: dict) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TerminalTitle",
        parent=styles["Heading1"],
        fontName=PDF_FONT_BOLD,
        fontSize=17,
        textColor=colors.HexColor("#0f172a"),
        spaceAfter=6,
    )
    subtitle_style = ParagraphStyle(
        "TerminalSub",
        parent=styles["Normal"],
        fontName=PDF_FONT_REGULAR,
        fontSize=9.5,
        textColor=colors.HexColor("#475569"),
        spaceAfter=14,
    )
    section_style = ParagraphStyle(
        "SectionHeader",
        parent=styles["Heading2"],
        fontName=PDF_FONT_BOLD,
        fontSize=12,
        textColor=colors.HexColor("#0284c7"),
        spaceBefore=12,
        spaceAfter=6,
    )
    item_title_style = ParagraphStyle(
        "ItemTitle",
        parent=styles["Heading3"],
        fontName=PDF_FONT_BOLD,
        fontSize=10.5,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=8,
        spaceAfter=3,
    )
    body_style = ParagraphStyle(
        "BodyTextClean",
        parent=styles["Normal"],
        fontName=PDF_FONT_REGULAR,
        fontSize=9.2,
        leading=13.5,
        textColor=colors.HexColor("#1e293b"),
        spaceAfter=8,
    )

    now_str = get_prague_now().strftime("%d. %m. %Y | %H:%M:%S CET")
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]
    world_sections, czech_sections = build_curated_czech_news_sections(forex_data, benchmarks, cnb)

    elements = [
        Paragraph("TENSEGRITY TERMINAL | INTELLIGENCE REPORT", title_style),
        Paragraph(f"Datum a čas vydání: {now_str} | Globální a český makroekonomický přehled", subtitle_style),
        Paragraph("1. GLOBÁLNÍ MAKROEKONOMICKÝ RÁMEC & MĚNOVÁ POLITIKA", section_style),
        Paragraph(
            f"Výnos desetiletých amerických dluhopisů dosahuje {benchmarks['us10y_yield']:.2f} % a průměrný poměr ceny "
            f"a zisku indexu S&P 500 činí {benchmarks['sp500_pe']:.1f}x. Americká centrální banka i Evropská centrální "
            f"banka udržují obezřetný postoj vůči inflaci ve službách.",
            body_style,
        ),
        Paragraph("2. ČESKÁ REPUBLIKA — SAZBY ČESKÉ NÁRODNÍ BANKY, INFLACE A MĚNOVÉ KURZY", section_style),
    ]

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

    elements.append(Paragraph("3. PĚT KLÍČOVÝCH ZPRÁV ZE SVĚTA (EKONOMIKA, TRHY A POLITIKA)", section_style))
    for item in world_sections:
        elements.append(Paragraph(f"{item['title']} [{item['badge']}]", item_title_style))
        elements.append(Paragraph(item["body"], body_style))

    elements.append(Paragraph("4. PĚT KLÍČOVÝCH ZPRÁV Z ČESKÉ REPUBLIKY (ČNB, TRHY A POLITIKA)", section_style))
    for item in czech_sections:
        elements.append(Paragraph(f"{item['title']} [{item['badge']}]", item_title_style))
        elements.append(Paragraph(item["body"], body_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer.read()


def generate_html_newsletter_content(forex_data: dict, benchmarks: dict, cnb: dict) -> str:
    now_str = get_prague_now().strftime("%d. %m. %Y | %H:%M:%S CET")
    eur_czk = forex_data["rates"]["EUR/CZK"]["rate"]
    usd_czk = forex_data["rates"]["USD/CZK"]["rate"]
    gbp_czk = forex_data["rates"]["GBP/CZK"]["rate"]

    us10y = benchmarks["us10y_yield"]
    sp500_pe = benchmarks["sp500_pe"]
    world_sections, czech_sections = build_curated_czech_news_sections(forex_data, benchmarks, cnb)

    def render_cards(sections: list[dict]) -> str:
        html_out = ""
        for item in sections:
            badge_bg = "rgba(56, 189, 248, 0.14)" if item["badge_color"] == "#38bdf8" else "rgba(245, 158, 11, 0.14)"
            html_out += f"""
            <div style="background-color: #0f1625; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 20px; margin-bottom: 14px;">
                <span style="display: inline-block; font-size: 10px; font-weight: 700; letter-spacing: 0.8px; background-color: {badge_bg}; color: {item['badge_color']}; padding: 4px 9px; border-radius: 5px; text-transform: uppercase; font-family: monospace;">{item['badge']}</span>
                <h4 style="color: #f8fafc; font-size: 15px; font-weight: 700; margin: 10px 0 8px 0; line-height: 1.4;">{item['title']}</h4>
                <p style="color: #94a3b8; font-size: 13px; line-height: 1.65; margin: 0 0 12px 0;">{item['body']}</p>
                <a href="{item['url']}" target="_blank" rel="noopener" style="color: #38bdf8; font-size: 12px; font-weight: 600; text-decoration: none;">&#128279; Oficiální zdroj &#8599;</a>
            </div>
            """
        return html_out

    world_cards_html = render_cards(world_sections)
    czech_cards_html = render_cards(czech_sections)

    return f"""<!DOCTYPE html>
<html lang="cs">
<head><meta charset="UTF-8"></head>
<body style="margin: 0; padding: 20px; background-color: #04070d; font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color: #f8fafc;">
    <div style="max-width: 820px; margin: 0 auto; background-color: #060911; border: 1px solid rgba(255,255,255,0.08); border-radius: 16px; padding: 32px;">
        
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 16px; margin-bottom: 24px;">
            <tr>
                <td align="left" style="font-size: 16px; font-weight: 800; letter-spacing: 3px; color: #ffffff; text-transform: uppercase;">
                    TENSEGRITY<span style="color: #38bdf8;">TERMINAL</span> <span style="font-size: 11px; font-weight: 500; color: #94a3b8; letter-spacing: 2px;">| INTELLIGENCE</span>
                </td>
                <td align="right" style="font-size: 12px; font-family: monospace; color: #94a3b8;">
                    {now_str}
                </td>
            </tr>
        </table>

        <div style="background-color: #0b1322; border: 1px solid rgba(56, 189, 248, 0.35); border-radius: 10px; padding: 14px 20px; text-align: center; margin-bottom: 28px; font-size: 13px; color: #38bdf8; font-weight: 600;">
            &#128206; Součástí tohoto e-mailu je kompletní přehled v přiloženém souboru (<span style="color: #ffffff;">Tensegrity_Terminal_Newsletter.pdf</span>).
        </div>

        <h3 style="color: #38bdf8; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.2px; border-bottom: 1px solid rgba(56, 189, 248, 0.25); padding-bottom: 8px; margin: 0 0 16px 0;">
            &#127757; 1. GLOBÁLNÍ SVĚT — AMERICKÁ CENTRÁLNÍ BANKA, EVROPSKÁ CENTRÁLNÍ BANKA A NÁLADA NA TRZÍCH
        </h3>

        <div style="background-color: #0f1625; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 22px; margin-bottom: 26px;">
            <h4 style="color: #f8fafc; font-size: 15px; font-weight: 700; margin: 0 0 10px 0;">Měnová politika, inflační setrvačnost a úrokové sazby ve světě</h4>
            <p style="color: #94a3b8; font-size: 13px; line-height: 1.65; margin: 0 0 14px 0;">
                Centrální bankéři ve Spojených státech amerických i v eurozóně nadále uplatňují obezřetný měnový kurz a vyhodnocují setrvačnost inflace v sektoru služeb. Přestože celková inflace oproti předchozím rokům výrazně ustoupila, napjatý trh práce a mzdová dynamika brání guvernérům v unáhleném uvolňování měnových podmínek.
            </p>
            <ul style="color: #94a3b8; font-size: 13px; line-height: 1.65; padding-left: 20px; margin: 0 0 18px 0;">
                <li style="margin-bottom: 6px;"><strong style="color: #f8fafc;">Dopad na dluhopisy a diskontní sazby:</strong> Výnos desetiletých amerických státních dluhopisů na úrovni <strong style="color: #38bdf8;">{us10y:.2f} %</strong> slouží jako hlavní globální bezriziková sazba pro oceňování budoucích peněžních toků.</li>
                <li><strong style="color: #f8fafc;">Rozdílná dynamika ekonomik:</strong> Americká ekonomika vykazuje silnější růstovou odolnost, zatímco Evropská centrální banka čelí strukturálnímu zpomalení průmyslového jádra v Německu.</li>
            </ul>

            <table role="presentation" width="100%" cellspacing="8" cellpadding="0" style="margin-bottom: 14px;">
                <tr>
                    <td width="33.3%" style="background-color: #080d18; border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; padding: 12px; text-align: center;">
                        <div style="font-size: 10px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.8px;">VÝNOS US 10Y DLUHOPISŮ</div>
                        <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: monospace; margin-top: 4px;">{us10y:.2f} %</div>
                    </td>
                    <td width="33.3%" style="background-color: #080d18; border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; padding: 12px; text-align: center;">
                        <div style="font-size: 10px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.8px;">P/E INDEXU S&amp;P 500</div>
                        <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: monospace; margin-top: 4px;">{sp500_pe:.1f}x</div>
                    </td>
                    <td width="33.3%" style="background-color: #080d18; border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; padding: 12px; text-align: center;">
                        <div style="font-size: 10px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.8px;">REPO SAZBA ČNB</div>
                        <div style="font-size: 15px; font-weight: 700; color: #f59e0b; font-family: monospace; margin-top: 4px;">{cnb['repo_rate']}</div>
                    </td>
                </tr>
            </table>

            <a href="https://www.federalreserve.gov/monetarypolicy.htm" target="_blank" rel="noopener" style="display: inline-block; padding: 7px 13px; border: 1px solid rgba(56,189,248,0.4); border-radius: 6px; color: #38bdf8; font-size: 11px; font-weight: 700; text-decoration: none; letter-spacing: 0.5px;">
                &#128279; OFICIÁLNÍ ZDROJ: AMERICKÁ A EVROPSKÁ CENTRÁLNÍ BANKA &#8599;
            </a>
        </div>

        <h3 style="color: #38bdf8; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.2px; border-bottom: 1px solid rgba(56, 189, 248, 0.25); padding-bottom: 8px; margin: 0 0 16px 0;">
            &#127464;&#127487; 2. ČESKÁ REPUBLIKA — ČESKÁ NÁRODNÍ BANKA, INFLACE, SAZBY A MĚNOVÉ KURZY
        </h3>

        <table role="presentation" width="100%" cellspacing="10" cellpadding="0" style="margin-bottom: 14px;">
            <tr>
                <td width="33.3%" style="background-color: #0f1625; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 16px; text-align: center;">
                    <div style="font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;">EUR / CZK</div>
                    <div style="font-size: 18px; font-weight: 800; color: #ffffff; font-family: monospace; margin-top: 6px;">{eur_czk:.2f} CZK</div>
                </td>
                <td width="33.3%" style="background-color: #0f1625; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 16px; text-align: center;">
                    <div style="font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;">USD / CZK</div>
                    <div style="font-size: 18px; font-weight: 800; color: #ffffff; font-family: monospace; margin-top: 6px;">{usd_czk:.2f} CZK</div>
                </td>
                <td width="33.3%" style="background-color: #0f1625; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 16px; text-align: center;">
                    <div style="font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;">GBP / CZK</div>
                    <div style="font-size: 18px; font-weight: 800; color: #ffffff; font-family: monospace; margin-top: 6px;">{gbp_czk:.2f} CZK</div>
                </td>
            </tr>
        </table>

        <table role="presentation" width="100%" cellspacing="10" cellpadding="0" style="margin-bottom: 26px;">
            <tr>
                <td width="50%" style="background-color: #0f1625; border: 1px solid rgba(245, 158, 11, 0.25); border-radius: 12px; padding: 16px; text-align: center;">
                    <div style="font-size: 11px; color: #f59e0b; text-transform: uppercase; letter-spacing: 1px;">NEJNOVĚJŠÍ INFLACE V ČR</div>
                    <div style="font-size: 18px; font-weight: 800; color: #ffffff; font-family: monospace; margin-top: 6px;">{cnb['latest_inflation']}</div>
                </td>
                <td width="50%" style="background-color: #0f1625; border: 1px solid rgba(245, 158, 11, 0.25); border-radius: 12px; padding: 16px; text-align: center;">
                    <div style="font-size: 11px; color: #f59e0b; text-transform: uppercase; letter-spacing: 1px;">ZÁKLADNÍ SAZBA ČNB (2T REPO)</div>
                    <div style="font-size: 18px; font-weight: 800; color: #ffffff; font-family: monospace; margin-top: 6px;">{cnb['repo_rate']}</div>
                </td>
            </tr>
        </table>

        <h3 style="color: #38bdf8; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.2px; border-bottom: 1px solid rgba(56, 189, 248, 0.25); padding-bottom: 8px; margin: 0 0 16px 0;">
            &#127760; 3. AKTUÁLNÍ ČERSTVÉ ZPRÁVY ZE SVĚTA (5 ODDÍLŮ)
        </h3>

        {world_cards_html}

        <h3 style="color: #f59e0b; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.2px; border-bottom: 1px solid rgba(245, 158, 11, 0.25); padding-bottom: 8px; margin: 24px 0 16px 0;">
            &#127464;&#127487; 4. AKTUÁLNÍ ČERSTVÉ ZPRÁVY Z ČESKÉ REPUBLIKY (5 ODDÍLŮ)
        </h3>

        {czech_cards_html}

        <div style="text-align: center; font-size: 11px; color: #64748b; border-top: 1px solid rgba(255,255,255,0.08); padding-top: 16px; margin-top: 24px;">
            Tensegrity Terminal | Intelligence Platform &bull; Automatický institucionální přehled
        </div>
    </div>
</body>
</html>"""


def _connect_ipv4_smtp_ssl(host: str, port: int, timeout: float = 15.0):
    addr_infos = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)
    last_exc = None
    for family, socktype, proto, _, sockaddr in addr_infos:
        sock = None
        try:
            sock = socket.socket(family, socktype, proto)
            sock.settimeout(timeout)
            sock.connect(sockaddr)
            context = ssl.create_default_context()
            tls_sock = context.wrap_socket(sock, server_hostname=host)
            server = smtplib.SMTP_SSL(host=host, port=port, timeout=timeout)
            server.sock = tls_sock
            server.file = tls_sock.makefile("rb")
            code, msg = server.getreply()
            if code != 220:
                raise smtplib.SMTPConnectError(code, msg)
            server.ehlo()
            return server
        except Exception as exc:
            last_exc = exc
            if sock is not None:
                try:
                    sock.close()
                except Exception:
                    pass
    if last_exc:
        raise last_exc
    raise OSError("Nepodařilo se navázat spojení se serverem elektronické pošty.")


def _connect_ipv4_smtp_starttls(host: str, port: int, timeout: float = 15.0):
    addr_infos = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)
    last_exc = None
    for family, socktype, proto, _, sockaddr in addr_infos:
        sock = None
        try:
            sock = socket.socket(family, socktype, proto)
            sock.settimeout(timeout)
            sock.connect(sockaddr)
            server = smtplib.SMTP(host=host, port=port, timeout=timeout)
            server.sock = sock
            server.file = sock.makefile("rb")
            code, msg = server.getreply()
            if code != 220:
                raise smtplib.SMTPConnectError(code, msg)
            server.ehlo()
            context = ssl.create_default_context()
            server.starttls(context=context)
            server.ehlo()
            return server
        except Exception as exc:
            last_exc = exc
            if sock is not None:
                try:
                    sock.close()
                except Exception:
                    pass
    if last_exc:
        raise last_exc
    raise OSError("Nepodařilo se navázat zabezpečené spojení se serverem elektronické pošty.")


def send_email_report(recipient: str, subject: str, html_content: str, pdf_bytes: bytes | None = None):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise RuntimeError("Chybí přihlašovací údaje odesílatele pro elektronickou poštu.")

    root_msg = MIMEMultipart("mixed")
    root_msg["Subject"] = subject
    root_msg["From"] = f"Tensegrity Terminal | Intelligence <{SENDER_EMAIL}>"
    root_msg["To"] = recipient

    alt_part = MIMEMultipart("alternative")
    plain_text = (
        "Tensegrity Terminal | Intelligence Report.\n"
        "Pro zobrazení kompletního vizuálního přehledu otevřete zprávu v režimu formátovaného textu "
        "nebo si otevřete přiložený soubor Tensegrity_Terminal_Newsletter.pdf."
    )
    alt_part.attach(MIMEText(plain_text, "plain", "utf-8"))
    alt_part.attach(MIMEText(html_content, "html", "utf-8"))
    root_msg.attach(alt_part)

    if pdf_bytes:
        pdf_attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
        pdf_attachment.add_header("Content-Disposition", "attachment", filename="Tensegrity_Terminal_Newsletter.pdf")
        root_msg.attach(pdf_attachment)

    errors = []
    for connector_name, connector_fn in [
        ("SSL_465", lambda: _connect_ipv4_smtp_ssl("smtp.gmail.com", 465)),
        ("STARTTLS_587", lambda: _connect_ipv4_smtp_starttls("smtp.gmail.com", 587)),
    ]:
        server = None
        try:
            server = connector_fn()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, [recipient], root_msg.as_string())
            server.quit()
            logger.info("Zpravodaj úspěšně odeslán přes %s.", connector_name)
            return True
        except smtplib.SMTPAuthenticationError as auth_err:
            raise RuntimeError(
                "Autentizace vůči poštovnímu účtu selhala. Ověřte platnost šestnáctiznakového hesla aplikace."
            ) from auth_err
        except Exception as conn_err:
            errors.append(f"{connector_name}: {conn_err}")
        finally:
            if server is not None:
                try:
                    server.close()
                except Exception:
                    pass

    raise RuntimeError(
        "Odchozí spojení na poštovní server (porty 465 i 587) je blokováno lokální sítí nebo firewallem. "
        f"Technický detail: {' | '.join(errors)}"
    )


async def dispatch_newsletter_to_recipients(recipients: list[str]):
    if not recipients:
        return
    forex_data = await asyncio.to_thread(get_live_forex_rates)
    benchmarks = await asyncio.to_thread(get_live_market_benchmarks)
    cnb = get_cnb_macro_indicators()
    html_body = await asyncio.to_thread(generate_html_newsletter_content, forex_data, benchmarks, cnb)
    pdf_bytes = await asyncio.to_thread(build_newsletter_pdf_bytes, forex_data, benchmarks, cnb)
    subject = "Tensegrity Terminal | Intelligence — Globální a Český Tržní Přehled"
    for recipient in recipients:
        try:
            await asyncio.to_thread(send_email_report, recipient, subject, html_body, pdf_bytes)
        except Exception as err:
            logger.error("Plánované odeslání zpravodaje selhalo: %s", err)


async def scheduled_newsletter_worker():
    sent_slots = load_sent_slots()
    # Rozvrh: Středa (2) v 7:00, Pátek (4) v 7:00, Neděle (6) v 19:00
    schedule_rules = {(2, 7), (4, 7), (6, 19)}
    while True:
        try:
            now = get_prague_now()
            slot_key = f"{now.strftime('%Y-%m-%d')}-{now.hour}"
            if (now.weekday(), now.hour) in schedule_rules and 0 <= now.minute < 15 and slot_key not in sent_slots:
                subscribers = get_subscribers_internal()
                if subscribers:
                    logger.info("Spouštím pravidelné plánované odeslání zpravodaje pro %d odběratelů.", len(subscribers))
                    await dispatch_newsletter_to_recipients(subscribers)
                sent_slots.add(slot_key)
                save_sent_slots(sent_slots)
        except Exception as err:
            logger.error("Chyba v plánovači zpravodaje: %s", err)
        await asyncio.sleep(30)


@asynccontextmanager
async def lifespan(app_instance: FastAPI):
    task = asyncio.create_task(scheduled_newsletter_worker())
    yield
    task.cancel()


app = FastAPI(title="Tensegrity Terminal | Institutional Valuation Platform", version="2.1", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def run_valuation_logic(ticker: str, leverage: float = 0.15, growth: float = 0.22, exit_mult: float = 24.0) -> dict:
    clean_ticker = ticker.strip().upper() or "NVDA"
    try:
        stock = yf.Ticker(clean_ticker)
        info = stock.info or {}
        fin_df = stock.financials
        bs_df = stock.balance_sheet
        cf_df = stock.cashflow
    except Exception:
        info = {}
        fin_df = None
        bs_df = None
        cf_df = None

    benchmarks = get_live_market_benchmarks()
    rf_rate = benchmarks["us10y_yield"] / 100.0
    erp_rate = 0.052

    currency = str(info.get("currency") or "USD")
    exchange = str(info.get("exchange") or "NMS")
    company_name = str(info.get("longName") or info.get("shortName") or clean_ticker)
    sector = str(info.get("sector") or "Technology")
    industry = str(info.get("industry") or "Semiconductors")
    country = str(info.get("country") or "United States")
    employees = int(safe_float(info.get("fullTimeEmployees"), 0))
    business_summary = str(info.get("longBusinessSummary") or "")

    current_price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"), 162.85)
    shares_outstanding = safe_float(info.get("sharesOutstanding"), 24_150_000_000.0)
    if shares_outstanding <= 0:
        shares_outstanding = 1_000_000_000.0

    market_cap = safe_float(info.get("marketCap"), current_price * shares_outstanding)
    total_debt = safe_float(info.get("totalDebt"), 11_000_000_000.0)
    total_cash = safe_float(info.get("totalCash"), 38_000_000_000.0)
    minority_interest = safe_float(info.get("minorityInterest"), 0.0)
    enterprise_value = market_cap + total_debt + minority_interest - total_cash

    # Vrstva 01: Historická data a normalizace výkazů
    rev_hist = extract_statement_series(fin_df, ["Total Revenue", "Operating Revenue"])
    ebitda_hist = extract_statement_series(fin_df, ["EBITDA", "Normalized EBITDA"])
    ebit_hist = extract_statement_series(fin_df, ["EBIT", "Operating Income"])
    net_inc_hist = extract_statement_series(fin_df, ["Net Income", "Net Income Common Stockholders"])
    int_exp_hist = extract_statement_series(fin_df, ["Interest Expense", "Interest Expense Non Operating"])
    ocf_hist = extract_statement_series(cf_df, ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"])
    capex_hist = extract_statement_series(cf_df, ["Capital Expenditure", "Purchase Of PPE"])
    sbc_hist = extract_statement_series(cf_df, ["Stock Based Compensation"])
    equity_hist = extract_statement_series(bs_df, ["Stockholders Equity", "Total Equity Gross Minority Interest"])
    cur_assets_hist = extract_statement_series(bs_df, ["Current Assets"])
    cur_liab_hist = extract_statement_series(bs_df, ["Current Liabilities"])
    total_assets_hist = extract_statement_series(bs_df, ["Total Assets"])
    retained_earn_hist = extract_statement_series(bs_df, ["Retained Earnings"])

    revenue = safe_float(info.get("totalRevenue"), rev_hist[0] if rev_hist else 130_000_000_000.0)
    reported_ebitda = safe_float(info.get("ebitda"), ebitda_hist[0] if ebitda_hist else revenue * 0.62)
    reported_ebit = ebit_hist[0] if ebit_hist else reported_ebitda * 0.94
    net_income = safe_float(info.get("netIncomeToCommon"), net_inc_hist[0] if net_inc_hist else revenue * 0.52)
    book_equity = equity_hist[0] if equity_hist else max(market_cap * 0.12, 50_000_000_000.0)
    sbc_expense = abs(sbc_hist[0]) if sbc_hist else revenue * 0.035

    raw_ocf = ocf_hist[0] if ocf_hist else safe_float(info.get("operatingCashflow"), reported_ebitda * 0.82)
    raw_capex = abs(capex_hist[0]) if capex_hist else revenue * 0.035
    reported_fcf = safe_float(info.get("freeCashflow"), raw_ocf - raw_capex)
    if reported_fcf <= 0:
        reported_fcf = max(reported_ebitda * 0.65, net_income * 0.85)

    # Ekonomická normalizace: započtení ekonomického nákladu akciových odměn zaměstnancům
    normalized_ebitda = reported_ebitda - (sbc_expense * 0.5)
    normalized_fcf = max(reported_fcf * 0.5, reported_fcf - (sbc_expense * 0.65))

    # Vrstva 02: Vážené průměrné náklady kapitálu a model oceňování kapitálových aktiv
    beta = safe_float(info.get("beta"), 1.28)
    beta = max(0.55, min(2.35, beta))
    cost_of_equity = rf_rate + beta * erp_rate

    interest_expense = abs(int_exp_hist[0]) if int_exp_hist else total_debt * 0.045
    pre_tax_cost_of_debt = (interest_expense / total_debt) if total_debt > 0 else (rf_rate + 0.011)
    pre_tax_cost_of_debt = max(rf_rate + 0.004, min(0.11, pre_tax_cost_of_debt))
    effective_tax_rate = 0.165

    total_capital = market_cap + total_debt
    weight_equity = (market_cap / total_capital) if total_capital > 0 else 0.95
    weight_debt = (total_debt / total_capital) if total_capital > 0 else 0.05
    wacc = (weight_equity * cost_of_equity) + (weight_debt * pre_tax_cost_of_debt * (1.0 - effective_tax_rate))
    wacc = max(0.075, min(0.15, wacc))

    invested_capital = max(1.0, book_equity + total_debt - total_cash)
    nopat = reported_ebit * (1.0 - effective_tax_rate)
    roic = nopat / invested_capital
    economic_profit = (roic - wacc) * invested_capital

    # Vícefázový model diskontovaných volných peněžních toků pro podnik (5 let)
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
            "gordon_tv": gordon_tv,
            "multiple_tv": multiple_tv,
            "tv_share_percent": tv_share_pct,
            "cash_flows": cf_rows,
        }

    dcf_res = simulate_fcff_dcf(growth, wacc, exit_mult, terminal_growth)
    dcf_intrinsic_share = dcf_res["intrinsic_value_per_share"]
    dcf_upside = ((dcf_intrinsic_share / current_price) - 1.0) * 100.0 if current_price > 0 else 0.0

    # Model diskontovaných peněžních toků pro akcionáře
    base_fcfe = max(1.0, normalized_fcf - interest_expense * (1.0 - effective_tax_rate))
    fcfe_pv = 0.0
    curr_fcfe = base_fcfe
    for yr in range(1, 6):
        yr_g = growth * (1.0 - 0.35 * ((yr - 1) / 4.0))
        curr_fcfe *= (1.0 + yr_g)
        fcfe_pv += curr_fcfe / ((1.0 + cost_of_equity) ** yr)
    fcfe_tv = (curr_fcfe * (1.0 + terminal_growth)) / max(0.02, cost_of_equity - terminal_growth)
    fcfe_equity_val = fcfe_pv + (fcfe_tv / ((1.0 + cost_of_equity) ** 5))
    fcfe_per_share = max(0.0, fcfe_equity_val / shares_outstanding)

    # Obrácený model diskontovaných peněžních toků
    low_g, high_g = -0.30, 1.20
    for _ in range(60):
        mid_g = (low_g + high_g) / 2.0
        test_price = simulate_fcff_dcf(mid_g, wacc, exit_mult, terminal_growth)["intrinsic_value_per_share"]
        if test_price < current_price:
            low_g = mid_g
        else:
            high_g = mid_g
    implied_growth_pct = ((low_g + high_g) / 2.0) * 100.0

    # Model reziduálního zisku a ekonomické přidané hodnoty
    roe_est = min(0.65, max(0.05, net_income / max(1.0, book_equity)))
    pv_ri = 0.0
    curr_bv = book_equity
    for yr in range(1, 6):
        faded_roe = roe_est * (1.0 - 0.12 * (yr - 1)) + cost_of_equity * (0.12 * (yr - 1))
        ri_yr = (faded_roe - cost_of_equity) * curr_bv
        pv_ri += ri_yr / ((1.0 + cost_of_equity) ** yr)
        curr_bv *= (1.0 + min(0.25, growth * 0.6))
    term_ri = ((cost_of_equity + 0.02 - cost_of_equity) * curr_bv) / max(0.02, cost_of_equity - terminal_growth)
    ri_equity_val = max(0.0, book_equity + pv_ri + (term_ri / ((1.0 + cost_of_equity) ** 5)))
    ri_per_share = ri_equity_val / shares_outstanding

    # Vrstva 03: Relativní ocenění pomocí srovnatelných tržních násobků
    sector_benchmarks = {
        "Technology": {"ev_ebitda": 24.0, "pe": 32.0, "ev_revenue": 11.5, "pb": 12.0},
        "Financial Services": {"ev_ebitda": 11.0, "pe": 13.5, "ev_revenue": 3.5, "pb": 1.45},
        "Healthcare": {"ev_ebitda": 16.5, "pe": 22.0, "ev_revenue": 4.8, "pb": 4.5},
        "Industrials": {"ev_ebitda": 14.0, "pe": 20.0, "ev_revenue": 2.4, "pb": 3.8},
        "Energy": {"ev_ebitda": 7.5, "pe": 11.5, "ev_revenue": 1.4, "pb": 1.8},
    }
    bench = sector_benchmarks.get(sector, {"ev_ebitda": 18.0, "pe": 24.0, "ev_revenue": 5.5, "pb": 5.0})
    fwd_factor = (1.0 + growth) ** 2

    peer_rows = [
        {
            "method": "Hodnota podniku / Provozní zisk (Forward 2Y)",
            "multiple": bench["ev_ebitda"],
            "forward_metric": round(normalized_ebitda * fwd_factor, 2),
            "implied_share_price": round(max(0.0, (normalized_ebitda * fwd_factor * bench["ev_ebitda"] + total_cash - total_debt) / shares_outstanding), 2),
        },
        {
            "method": "Cena / Čistý zisk (Forward 2Y)",
            "multiple": bench["pe"],
            "forward_metric": round(net_income * fwd_factor, 2),
            "implied_share_price": round(max(0.0, (net_income * fwd_factor * bench["pe"]) / shares_outstanding), 2),
        },
        {
            "method": "Hodnota podniku / Tržby (Forward 2Y)",
            "multiple": bench["ev_revenue"],
            "forward_metric": round(revenue * fwd_factor, 2),
            "implied_share_price": round(max(0.0, (revenue * fwd_factor * bench["ev_revenue"] + total_cash - total_debt) / shares_outstanding), 2),
        },
        {
            "method": "Cena / Účetní hodnota vlastního kapitálu",
            "multiple": bench["pb"],
            "forward_metric": round(book_equity, 2),
            "implied_share_price": round(max(0.0, (book_equity * bench["pb"]) / shares_outstanding), 2),
        },
    ]
    comps_prices = [r["implied_share_price"] for r in peer_rows[:3] if r["implied_share_price"] > 0]
    blended_comps_price = sum(comps_prices) / len(comps_prices) if comps_prices else dcf_intrinsic_share

    # Vrstva 04: Majetkové ocenění (součet částí a čistá hodnota aktiv)
    core_segment_ev = normalized_ebitda * exit_mult * 0.82
    growth_segment_ev = normalized_ebitda * (exit_mult * 1.15) * 0.18
    sotp_ev = core_segment_ev + growth_segment_ev
    sotp_equity = max(0.0, sotp_ev + total_cash - total_debt - minority_interest)
    sotp_per_share = sotp_equity / shares_outstanding
    nav_per_share = max(0.0, (book_equity + (total_cash - total_debt) * 0.25) / shares_outstanding)

    # Vrstva 05: Model pákového odkupu a citlivostní matice
    lbo_entry_debt = enterprise_value * leverage
    lbo_entry_equity = max(1000.0, enterprise_value - lbo_entry_debt)

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

    lbo_years = simulate_lbo(growth)
    lbo_exit_ev = lbo_years[-1]["projektovany_provozni_zisk_pred_odpisy"] * exit_mult
    lbo_exit_equity = max(0.0, lbo_exit_ev - lbo_years[-1]["zbyvajici_cizi_kapital"])
    lbo_moic = (lbo_exit_equity / lbo_entry_equity) if lbo_entry_equity > 0 else 0.0
    lbo_irr = ((lbo_moic ** (1.0 / 5.0)) - 1.0) * 100.0 if lbo_moic > 0 else -100.0

    growth_grid = [max(-0.50, growth + d) for d in [-0.04, -0.02, 0.0, 0.02, 0.04]]
    mult_grid = [max(2.0, exit_mult + d) for d in [-5.0, -2.5, 0.0, 2.5, 5.0]]
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
        lbo_sensitivity.append({
            "tempo_rustu_ebitda_procent": round(g_val * 100.0, 1),
            "vysledky_irr_dle_nasobku": row_cells,
        })

    wacc_grid = [max(0.05, wacc + d) for d in [-0.015, -0.0075, 0.0, 0.0075, 0.015]]
    tg_grid = [0.015, 0.020, 0.025, 0.030, 0.035]
    dcf_sensitivity = []
    for w_val in wacc_grid:
        cells = []
        for tg_val in tg_grid:
            safe_tg = min(tg_val, w_val - 0.015)
            p_val = simulate_fcff_dcf(growth, w_val, exit_mult, safe_tg)["intrinsic_value_per_share"]
            cells.append({"tg": round(tg_val * 100.0, 2), "price": round(p_val, 2)})
        dcf_sensitivity.append({"wacc": round(w_val * 100.0, 2), "cells": cells})

    # Vrstva 06: Scénáře a simulace Monte Carlo (2 500 běhů) + histogram pro živý graf
    bear_val = simulate_fcff_dcf(growth * 0.55, wacc + 0.015, exit_mult * 0.78, max(0.015, terminal_growth - 0.005))["intrinsic_value_per_share"]
    base_val = dcf_intrinsic_share
    bull_val = simulate_fcff_dcf(growth * 1.30, max(0.065, wacc - 0.010), exit_mult * 1.18, min(0.035, terminal_growth + 0.005))["intrinsic_value_per_share"]
    scenario_expected = 0.25 * bear_val + 0.50 * base_val + 0.25 * bull_val

    rng = random.Random(42)
    mc_results = []
    for _ in range(2500):
        sim_g = rng.gauss(growth, max(0.03, abs(growth) * 0.22))
        sim_wacc = max(0.06, min(0.18, rng.gauss(wacc, 0.011)))
        sim_mult = max(5.0, rng.gauss(exit_mult, exit_mult * 0.14))
        sim_margin_shock = rng.gauss(1.0, 0.08)
        sim_fcf = normalized_fcf * sim_margin_shock
        pv_cf = 0.0
        c_fcf = sim_fcf
        c_eb = normalized_ebitda * sim_margin_shock
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

    # Výpočet 14 intervalů histogramu mezi P2 a P98 pro graf Monte Carlo
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
        mc_histogram.append({
            "label": f"{b_mid:.1f}",
            "midpoint": round(b_mid, 2),
            "count": count,
        })

    # Vrstva 07: Syntéza ocenění
    if sector == "Financial Services":
        weights = {
            "Vícefázový DCF pro podnik": 0.15,
            "DCF pro vlastní kapitál": 0.25,
            "Srovnatelné tržní násobky": 0.25,
            "Model reziduálního zisku": 0.20,
            "Metoda součtu částí": 0.05,
            "Medián Monte Carlo (P50)": 0.10,
        }
    else:
        weights = {
            "Vícefázový DCF pro podnik": 0.35,
            "DCF pro vlastní kapitál": 0.10,
            "Srovnatelné tržní násobky": 0.20,
            "Model reziduálního zisku": 0.10,
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
    if reported_fcf > 0:
        integrity_score += 10
    if roic > wacc:
        integrity_score += 10
    spread_ratio = abs(dcf_intrinsic_share - blended_comps_price) / max(1.0, weighted_fair_value)
    if spread_ratio < 0.20:
        integrity_score += 5
    elif spread_ratio > 0.45:
        integrity_score -= 10
    integrity_score = max(35, min(98, integrity_score))

    if spread_ratio > 0.25:
        disagreement_note = (
            f"Rozptyl mezi fundamentálním výnosovým modelem ({dcf_intrinsic_share:,.2f} {currency}) a tržními násobky "
            f"({blended_comps_price:,.2f} {currency}) dosahuje {spread_ratio * 100:.1f} %. Rozdíl je způsoben projekcí "
            f"růstu volných peněžních toků ({growth * 100:.1f} %) vůči průměrným sektorovým násobkům."
        )
    else:
        disagreement_note = (
            f"Vysoká konvergence oceňovacích přístupů: fundamentální výnosový model ({dcf_intrinsic_share:,.2f} {currency}) "
            f"a relativní násobky ({blended_comps_price:,.2f} {currency}) potvrzují konzistentní odhad vnitřní hodnoty."
        )

    # Vrstva 08: Rozhodovací rámec a analýza rizik
    mos_percent = ((weighted_fair_value - current_price) / weighted_fair_value) * 100.0 if weighted_fair_value > 0 else 0.0
    upside_to_fair = ((weighted_fair_value / current_price) - 1.0) * 100.0 if current_price > 0 else 0.0

    green_upper = weighted_fair_value * 0.80
    yellow_upper = weighted_fair_value * 0.90
    orange_upper = weighted_fair_value * 1.00

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
    if raw_ocf > 0:
        f_score += 1
    if raw_ocf > net_income:
        f_score += 1
    if roic > wacc:
        f_score += 1
    net_debt = total_debt - total_cash
    if net_debt <= 0 or (total_debt / max(1.0, reported_ebitda)) < 2.0:
        f_score += 1
    cur_ratio = (cur_assets_hist[0] / cur_liab_hist[0]) if (cur_assets_hist and cur_liab_hist and cur_liab_hist[0] > 0) else 2.2
    if cur_ratio > 1.2:
        f_score += 1
    if len(rev_hist) < 2 or rev_hist[0] >= rev_hist[1]:
        f_score += 1
    if len(ebit_hist) < 2 or (ebit_hist[0] / max(1.0, rev_hist[0])) >= (ebit_hist[1] / max(1.0, rev_hist[1])):
        f_score += 1
    if revenue / max(1.0, total_assets_hist[0] if total_assets_hist else invested_capital) > 0.5:
        f_score += 1

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
        "ticker": clean_ticker,
        "currency": currency,
        "exchange": exchange,
        "sector": sector,
        "industry": industry,
        "country": country,
        "employees": employees,
        "business_summary": business_summary,
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
            "roic_percent": round(roic * 100.0, 2),
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
            "assumed_growth_percent": round(growth * 100.0, 2),
            "expectation_gap_percent": round((growth * 100.0) - implied_growth_pct, 2),
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
            "model_version": "Tensegrity Institutional Engine v2.1",
            "timestamp": get_prague_now().strftime("%Y-%m-%d %H:%M:%S %Z"),
            "data_source": "Yahoo Finance Institutional Feed + Live Treasury Benchmark",
            "currency": currency,
            "units": "Absolutní hodnoty (akcie v jednotkách)",
            "assumptions": {
                "leverage_percent": round(leverage * 100.0, 1),
                "growth_percent": round(growth * 100.0, 1),
                "exit_multiple": round(exit_mult, 1),
                "terminal_growth_percent": round(terminal_growth * 100.0, 2),
            },
        },
    }


def _fmt_cz_large(val: float, curr: str) -> str:
    abs_v = abs(val)
    if abs_v >= 1e12:
        return f"{val / 1e12:.2f} bil. {curr}"
    if abs_v >= 1e9:
        return f"{val / 1e9:.2f} mld. {curr}"
    if abs_v >= 1e6:
        return f"{val / 1e6:.2f} mil. {curr}"
    return f"{val:,.0f} {curr}"


def call_external_llm_if_configured(question: str, history: list[dict], d: dict) -> str | None:
    if not GEMINI_API_KEY:
        return None
    try:
        compact_context = {
            "company": d["company_name"],
            "ticker": d["ticker"],
            "sector": d["sector"],
            "industry": d["industry"],
            "business_summary": d["business_summary"][:1200],
            "currency": d["currency"],
            "market_cap": d["trzni_kapitalizace"],
            "enterprise_value": d["celkova_hodnota_podniku"],
            "total_debt": d["celkovy_dluh"],
            "total_cash": d["hotovost"],
            "net_debt": d["cisty_dluh"],
            "fundamentals": d["fundamentals"],
            "wacc_engine": d["wacc_engine"],
            "dcf_intrinsic_per_share": d["dcf_valuation"]["intrinsic_value_per_share"],
            "current_share_price": d["dcf_valuation"]["current_share_price"],
            "tv_share_percent": d["dcf_valuation"]["tv_share_percent"],
            "fcfe_per_share": d["fcfe_valuation"]["intrinsic_value_per_share"],
            "residual_income_per_share": d["residual_income"]["intrinsic_value_per_share"],
            "peer_comps_price": d["peer_comps"]["blended_implied_price"],
            "sotp_per_share": d["metoda_souctu_casti"]["sotp_per_share"],
            "reverse_dcf": d["reverse_dcf"],
            "scenarios": d["scenarios"],
            "monte_carlo": {k: v for k, v in d["monte_carlo"].items() if k != "histogram"},
            "valuation_synthesis": d["valuation_synthesis"],
            "decision_engine": d["decision_engine"],
            "risk_analysis": d["risk_analysis"],
        }
        system_instruction = (
            "Jsi institucionální finanční a investiční analytik platformy Tensegrity Terminal. "
            "Odpovídej vždy česky, věcně, přesně, kriticky a přímo na otázku uživatele. "
            "Nepoužívej zbytečné zkratky, vysvětluj souvislosti s konkrétními čísly z přiloženého modelu společnosti. "
            "Pokud se uživatel ptá na obecné finanční, ekonomické nebo investiční téma, odpověz naplno a propojíš to s kontextem.\n"
            f"Data aktuálně analyzované společnosti: {json.dumps(compact_context, ensure_ascii=False)}"
        )
        history_text = ""
        for msg in history[-6:]:
            role = "Uživatel" if msg.get("role") == "user" else "Poradce"
            history_text += f"{role}: {msg.get('text', '')}\n"

        prompt_body = {
            "contents": [
                {
                    "parts": [
                        {"text": f"{system_instruction}\n\nHistorie konverzace:\n{history_text}\nAktuální dotaz uživatele: {question}"}
                    ]
                }
            ]
        }
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
        req = urllib.request.Request(
            url,
            data=json.dumps(prompt_body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        candidates = raw.get("candidates") or []
        if candidates:
            parts = candidates[0].get("content", {}).get("parts") or []
            if parts and parts[0].get("text"):
                return parts[0]["text"].strip()
    except Exception as err:
        logger.warning("Externí jazykový model nebyl dostupný, přepínám na lokální analytické jádro: %s", err)
    return None


def generate_smart_advisor_answer(question: str, history: list[dict], d: dict) -> str:
    llm_reply = call_external_llm_if_configured(question, history, d)
    if llm_reply:
        return llm_reply

    q = question.lower().strip()
    curr = d["currency"]
    name = d["company_name"]
    tk = d["ticker"]
    p_now = d["dcf_valuation"]["current_share_price"]
    p_fair = d["valuation_synthesis"]["weighted_fair_value"]
    dcf_p = d["dcf_valuation"]["intrinsic_value_per_share"]
    fcfe_p = d["fcfe_valuation"]["intrinsic_value_per_share"]
    ri_p = d["residual_income"]["intrinsic_value_per_share"]
    comps_p = d["peer_comps"]["blended_implied_price"]
    sotp_p = d["metoda_souctu_casti"]["sotp_per_share"]
    mos = d["decision_engine"]["margin_of_safety_percent"]
    zone = d["decision_engine"]["zone"]
    tranche = d["decision_engine"]["tranche_allocation_percent"]
    f = d["fundamentals"]
    w = d["wacc_engine"]
    rev = d["reverse_dcf"]
    sc = d["scenarios"]
    mc = d["monte_carlo"]
    rsk = d["risk_analysis"]
    lbo = d["model_odkupu_za_cizi_kapital_base_case"]

    # 1. Vysvětlení konkrétních pojmů ("co znamená", "co je", "vysvětli")
    if any(k in q for k in ["co znamená", "co je ", "vysvětli", "definice", "jak se počítá"]):
        if "wacc" in q or "náklady kapitálu" in q or "diskont" in q:
            return (
                f"Vážené průměrné náklady kapitálu představují minimální roční výnosnost, kterou musí {name} ({tk}) dosáhnout, "
                f"aby uspokojila jak své akcionáře, tak věřitele. V aktuálním modelu vycházejí na {w['wacc_percent']:.2f} %.\n\n"
                f"Jak jsme k tomuto číslu došli:\n"
                f"• Náklady vlastního kapitálu činí {w['cost_of_equity_percent']:.2f} % (součet bezrizikové sazby desetiletých amerických dluhopisů "
                f"{w['risk_free_rate_percent']:.2f} % a součinu citlivosti akcie {w['beta']:.2f} s tržní rizikovou prémií {w['erp_percent']:.2f} %).\n"
                f"• Náklady cizího kapitálu po zdanění činí {w['after_tax_cost_of_debt_percent']:.2f} %.\n"
                f"• Vzhledem ke kapitálové struktuře ({w['weight_equity_percent']:.1f} % vlastní kapitál a {w['weight_debt_percent']:.1f} % dluh) "
                f"vzniká výsledná diskontní sazba {w['wacc_percent']:.2f} %. Čím vyšší tato sazba je, tím nižší je současná hodnota budoucích peněžních toků."
            )
        if "reverse" in q or "obrácen" in q or "implikovan" in q:
            return (
                f"Obrácený model diskontovaných peněžních toků otáčí běžný výpočet naruby: místo hledání vnitřní hodnoty zjišťuje, "
                f"jaké roční tempo růstu volných peněžních toků už je započítáno v dnešní tržní ceně akcie ({p_now:,.2f} {curr}).\n\n"
                f"U společnosti {name} ({tk}) současná cena na burze předpokládá roční růst volných peněžních toků o {rev['implied_fcf_growth_percent']:.2f} % "
                f"v příštích pěti letech. V našem scénáři uvažujeme růst {rev['assumed_growth_percent']:.2f} %, což vytváří rozdíl "
                f"{rev['expectation_gap_percent']:+.2f} procentního bodu mezi naším odhadem a očekáváním trhu."
            )
        if "monte carlo" in q or "percentil" in q or "p50" in q:
            return (
                f"Simulace Monte Carlo spouští 2 500 nezávislých výpočtů vnitřní hodnoty, ve kterých se současně náhodně mění tempo růstu, "
                f"diskontní sazba, výstupní násobek i provozní marže podle pravděpodobnostního rozdělení.\n\n"
                f"Pro {name} ({tk}) ukazuje prostřední hodnota (medián P50) úroveň {mc['p50']:,.2f} {curr}. "
                f"Osmdesát procent všech scénářů se nachází mezi {mc['p10']:,.2f} {curr} a {mc['p90']:,.2f} {curr}, "
                f"přičemž pravděpodobnost, že vnitřní hodnota převyšuje dnešní tržní cenu, dosahuje {mc['prob_above_market_percent']:.1f} %."
            )
        if "roic" in q or "návratnost" in q or "ekonomický zisk" in q or "eva" in q:
            return (
                f"Návratnost investovaného kapitálu měří, kolik čistého provozního zisku po zdanění dokáže {name} ({tk}) vygenerovat "
                f"z každé koruny či dolaru vloženého do podnikání. Aktuálně dosahuje {f['roic_percent']:.2f} %.\n\n"
                f"Protože tato návratnost převyšuje vážené průměrné náklady kapitálu ({w['wacc_percent']:.2f} %), firma vytváří skutečnou "
                f"ekonomickou hodnotu ve výši {_fmt_cz_large(f['economic_profit'], curr)} ročně. Pokud by návratnost klesla pod náklady kapitálu, "
                f"růst firmy by hodnotu pro akcionáře naopak ničil."
            )
        if "piotroski" in q or "altman" in q or "skóre" in q:
            return (
                f"Piotroski skóre (0 až 9 bodů) hodnotí devět účetních kritérií ziskovosti, tvorby hotovosti, zadlužení a provozní efektivity. "
                f"{name} ({tk}) dosahuje {rsk['piotroski_f_score']} z 9 bodů.\n\n"
                f"Altmanovo Z-skóre měří finanční stabilitu a odolnost vůči úpadku kombinací pracovního kapitálu, nerozděleného zisku, "
                f"provozního zisku, tržní hodnoty vlastního kapitálu a obratu aktiv. Hodnota {rsk['altman_z_score']:.2f} znamená: {rsk['altman_status']}."
            )
        if "margin of safety" in q or "bezpečnostní rezerva" in q or "zón" in q:
            return (
                f"Bezpečnostní rezerva vyjadřuje procentuální rozdíl mezi odhadovanou váženou vnitřní hodnotou ({p_fair:,.2f} {curr}) "
                f"a aktuální tržní cenou ({p_now:,.2f} {curr}). U {name} ({tk}) nyní činí {mos:.2f} %.\n\n"
                f"Model rozděluje cenové pásmo do čtyř zón: zelená (rezerva nad 20 %, alokace 55 % plánované pozice), "
                f"žlutá (rezerva 10 až 20 %, alokace 30 %), oranžová (rezerva 0 až 10 %, alokace 15 %) a červená (cena nad vnitřní hodnotou, alokace 0 %). "
                f"Aktuálně se titul nachází v zóně {zone}."
            )

    # 2. Zadluženost, hotovost, likvidita, rozvaha
    if any(k in q for k in ["dluh", "zadlužen", "hotovost", "likvidit", "solvent", "úrok", "rozvah", "přežije"]):
        net_debt_txt = (
            f"čistou hotovost ve výši {_fmt_cz_large(abs(d['cisty_dluh']), curr)} (hotovost převyšuje celkový dluh)"
            if d["cisty_dluh"] <= 0
            else f"čistý dluh ve výši {_fmt_cz_large(d['cisty_dluh'], curr)}, což odpovídá {f['net_debt_to_ebitda']:.2f}násobku normalizovaného provozního zisku"
        )
        return (
            f"Rozvahová a dluhová analýza společnosti {name} ({tk}):\n"
            f"• Celkový úročený dluh činí {_fmt_cz_large(d['celkovy_dluh'], curr)}, zatímco disponibilní hotovost a krátkodobé investice dosahují {_fmt_cz_large(d['hotovost'], curr)}.\n"
            f"• Firma tedy vykazuje {net_debt_txt}.\n"
            f"• Běžná likvidita (poměr oběžných aktiv ke krátkodobým závazkům) je {f['current_ratio']:.2f} a úrokové krytí provozním ziskem dosahuje {f['interest_coverage']:.1f}násobku.\n"
            f"• Altmanovo Z-skóre stability je {rsk['altman_z_score']:.2f} ({rsk['altman_status']}). Z pohledu rozvahy je finanční riziko {'velmi nízké' if d['cisty_dluh'] <= 0 or f['net_debt_to_ebitda'] < 1.5 else 'zvýšené a vyžaduje sledování úrokových nákladů'}."
        )

    # 3. Negativa, skryté chyby, problémy firmy, rizika
    if any(k in q for k in ["negativ", "skryt", "chyb", "problém", "rizik", "slab", "pozor", "hrozb", "špatn"]):
        red_flags = []
        sbc_share = (f["sbc_expense"] / max(1.0, f["reported_fcf"])) * 100.0
        red_flags.append(
            f"1. Ředění prostřednictvím akciových odměn zaměstnancům: Vykazovaný volný peněžní tok ({_fmt_cz_large(f['reported_fcf'], curr)}) "
            f"je nadhodnocen nepeněžními akciovými odměnami ve výši {_fmt_cz_large(f['sbc_expense'], curr)} ({sbc_share:.1f} % reportovaného toku). "
            f"Po jejich ekonomickém očištění klesá reálný peněžní tok pro akcionáře na {_fmt_cz_large(f['normalized_fcf'], curr)}."
        )
        if d["dcf_valuation"]["tv_share_percent"] > 65.0:
            red_flags.append(
                f"2. Vysoká citlivost na vzdálenou budoucnost: Plných {d['dcf_valuation']['tv_share_percent']:.1f} % hodnoty podniku ve výnosovém modelu "
                f"pochází až z terminální hodnoty po pátém roce. I malá změna úrokových sazeb nebo zpomalení růstu po pátém roce výrazně zahýbe s cenou akcie."
            )
        if rev["implied_fcf_growth_percent"] > 15.0:
            red_flags.append(
                f"3. Vysoko nastavená laťka tržních očekávání: Současná burzovní cena už v sobě započítává roční růst peněžních toků o "
                f"{rev['implied_fcf_growth_percent']:.2f} % po dobu pěti let. Jakékoliv čtvrtletní zaváhání v maržích může vyvolat prudkou korekci násobků."
            )
        if mos < 0:
            red_flags.append(
                f"4. Absence bezpečnostní rezervy: Akcie se obchoduje o {abs(mos):.2f} % dráže, než činí naše vážená vnitřní hodnota ({p_fair:,.2f} {curr}), "
                f"což neposkytuje žádný polštář proti makroekonomickému zpomalení."
            )
        if d["cisty_dluh"] > 0 and f["net_debt_to_ebitda"] > 2.0:
            red_flags.append(
                f"5. Vyšší dluhová páka: Čistý dluh dosahuje {f['net_debt_to_ebitda']:.2f}násobku provozního zisku, což zvyšuje citlivost na refinanční sazby."
            )
        if len(red_flags) < 3:
            red_flags.append(
                f"3. Cyklicita a koncentrace odvětví ({d['sector']} / {d['industry']}): Provozní marže ({f['operating_margin_percent']:.1f} %) "
                f"jsou vystaveny investičním cyklům odběratelů, konkurenčnímu tlaku a regulačním zásahům."
            )
        return (
            f"Hloubkový audit rizik, slabin a skrytých úskalí pro {name} ({tk}):\n\n"
            + "\n\n".join(red_flags)
        )

    # 4. Pozitiva, silné stránky, proč koupit
    if any(k in q for k in ["pozitiv", "siln", "výhod", "proč koupit", "atraktiv", "kvalit", "potenciál"]):
        positives = [
            f"1. Tvorba ekonomické hodnoty: Návratnost investovaného kapitálu ({f['roic_percent']:.2f} %) výrazně překonává náklady kapitálu ({w['wacc_percent']:.2f} %), "
            f"čímž firma generuje čistý ekonomický zisk {_fmt_cz_large(f['economic_profit'], curr)} ročně.",
            f"2. Silná přeměna zisku na hotovost: Konverze normalizovaného provozního zisku do volného peněžního toku dosahuje {f['fcf_conversion_percent']:.1f} %, "
            f"což umožňuje financovat další expanzi bez nutnosti externího zadlužování.",
            f"3. Fundamentální zdraví účetních výkazů: Piotroski skóre dosahuje {rsk['piotroski_f_score']} z 9 bodů při provozní marži {f['operating_margin_percent']:.1f} %.",
        ]
        if d["cisty_dluh"] <= 0:
            positives.append(
                f"4. Pevnostní rozvaha s čistou hotovostí: Hotovost ({_fmt_cz_large(d['hotovost'], curr)}) převyšuje celkový dluh ({_fmt_cz_large(d['celkovy_dluh'], curr)}), "
                f"což dává vedení prostor pro zpětné odkupy akcií i odolnost v recesi."
            )
        if mos > 0:
            positives.append(
                f"5. Pozitivní bezpečnostní rezerva: Akcie se obchoduje se slevou {mos:.2f} % pod váženou vnitřní hodnotou {p_fair:,.2f} {curr} (zóna {zone})."
            )
        return f"Hlavní fundamentální pozitiva a konkurenční přednosti {name} ({tk}):\n\n" + "\n\n".join(positives)

    # 5. Proč valuace vyšla takhle / vysvětlení výpočtu
    if any(k in q for k in ["proč", "vyšla", "vysvětli valuaci", "jak vyšla", "rozdíl", "metod"]):
        return (
            f"Rozklad toho, proč vážená vnitřní hodnota {name} ({tk}) vyšla na {p_fair:,.2f} {curr} oproti tržní ceně {p_now:,.2f} {curr}:\n\n"
            f"• Výchozí základna: Vycházíme z ročních tržeb {_fmt_cz_large(f['revenue'], curr)} a normalizovaného volného peněžního toku "
            f"{_fmt_cz_large(f['normalized_fcf'], curr)} (po odečtení ekonomického nákladu akciových odměn).\n"
            f"• Jednotlivé oceňovací přístupy dávají tyto výsledky:\n"
            f"  – Vícefázový výnosový model pro podnik (váha 35 %): {dcf_p:,.2f} {curr} při diskontní sazbě {w['wacc_percent']:.2f} %.\n"
            f"  – Srovnatelné tržní násobky sektoru (váha 20 %): {comps_p:,.2f} {curr}.\n"
            f"  – Medián simulace Monte Carlo (váha 15 %): {mc['p50']:,.2f} {curr}.\n"
            f"  – Výnosový model pro vlastní kapitál (váha 10 %): {fcfe_p:,.2f} {curr}.\n"
            f"  – Model reziduálního zisku (váha 10 %): {ri_p:,.2f} {curr}.\n"
            f"  – Ocenění součtem částí (váha 10 %): {sotp_p:,.2f} {curr}.\n\n"
            f"• Diagnostika shody: {d['valuation_synthesis']['disagreement_note']}"
        )

    # 6. Scénáře, výhled, očekávání
    if any(k in q for k in ["scénář", "očekávat", "výhled", "budoucnost", "bear", "bull", "pravděpodobnost"]):
        return (
            f"Scénářová a pravděpodobnostní analýza vývoje hodnoty pro {name} ({tk}):\n\n"
            f"• Pesimistický scénář (pravděpodobnost 25 %): Při zpomalení růstu a vyšších úrokových sazbách vychází hodnota akcie na {sc['bear']['fair_value']:,.2f} {curr}.\n"
            f"• Základní scénář (pravděpodobnost 50 %): Při naplnění modelových předpokladů činí hodnota {sc['base']['fair_value']:,.2f} {curr}.\n"
            f"• Optimistický scénář (pravděpodobnost 25 %): Při akceleraci tržeb a expanzi násobků dosahuje hodnota {sc['bull']['fair_value']:,.2f} {curr}.\n"
            f"• Pravděpodobnostně vážená očekávaná hodnota scénářů činí {sc['expected_value']:,.2f} {curr}.\n"
            f"• Z 2 500 simulací Monte Carlo končí {mc['prob_above_market_percent']:.1f} % případů nad dnešní tržní cenou ({p_now:,.2f} {curr}), "
            f"přičemž střední hodnota leží na {mc['p50']:,.2f} {curr}."
        )

    # 7. Profil společnosti, odvětví, čím se zabývá
    if any(k in q for k in ["čím se zabývá", "co dělá", "profil", "byznys", "odvětví", "sektor", "zaměstnanc", "společnost"]):
        summary_snippet = d["business_summary"][:450] + ("..." if len(d["business_summary"]) > 450 else "")
        return (
            f"Profil společnosti {name} ({tk}):\n"
            f"• Zařazení: Sektor {d['sector']}, odvětví {d['industry']}, sídlo {d['country']}, počet zaměstnanců přibližně {d['employees']:,}.\n"
            f"• Velikost a výkonnost: Tržní kapitalizace {_fmt_cz_large(d['trzni_kapitalizace'], curr)}, roční tržby {_fmt_cz_large(f['revenue'], curr)} "
            f"a provozní marže {f['operating_margin_percent']:.1f} %.\n"
            f"• Popis činnosti z burzovního prospektu: {summary_snippet if summary_snippet else 'Podrobný popis činnosti není u tohoto titulu v datovém zdroji uveden.'}"
        )

    # 8. Pákový odkup
    if any(k in q for k in ["pákov", "odkup", "lbo", "irr", "moic", "akvizic"]):
        return (
            f"Analýza modelu pákového odkupu pro {name} ({tk}) při nastavení {d['audit_trail']['assumptions']['leverage_percent']:.1f} % cizího kapitálu:\n"
            f"• Vstupní vlastní kapitál investora činí {_fmt_cz_large(lbo['vstupni_vlastni_kapital'], curr)} a vstupní akviziční dluh {_fmt_cz_large(lbo['vstupni_cizi_kapital'], curr)}.\n"
            f"• Za pět let splácení dluhu z volných peněžních toků dosahuje výstupní hodnota vlastního kapitálu {_fmt_cz_large(lbo['vystupni_hodnota_vlastniho_kapitalu'], curr)}.\n"
            f"• Výsledný násobek investovaného kapitálu činí {lbo['moic']}x a vnitřní míra výnosnosti dosahuje {lbo['irr_procent']:.2f} % ročně."
        )

    # 9. Obecný inteligentní rozbor na jakýkoliv jiný dotaz
    return (
        f"Odpověď k vašemu dotazu „{question}“ v kontextu společnosti {name} ({tk}):\n\n"
        f"• Aktuální tržní cena činí {p_now:,.2f} {curr}, zatímco syntéza šesti nezávislých oceňovacích metod stanovuje váženou vnitřní hodnotu "
        f"na {p_fair:,.2f} {curr} (bezpečnostní rezerva {mos:.2f} %, zóna {zone} s doporučenou alokací {tranche} %).\n"
        f"• Klíčový fundamentální závěr: Firma generuje návratnost investovaného kapitálu {f['roic_percent']:.2f} % proti nákladům kapitálu "
        f"{w['wacc_percent']:.2f} %, což znamená roční tvorbu ekonomické hodnoty {_fmt_cz_large(f['economic_profit'], curr)}.\n"
        f"• Tržní očekávání vs. rizika: Cena na burze předpokládá roční růst peněžních toků o {rev['implied_fcf_growth_percent']:.2f} %. "
        f"Pokud se chcete zaměřit na konkrétní oblast, zeptejte se například na zadluženost, skryté účetní chyby, vysvětlení libovolného ukazatele "
        f"nebo pravděpodobnost jednotlivých scénářů."
    )


@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    if os.path.exists(LOGO_PATH):
        return FileResponse(LOGO_PATH)
    raise HTTPException(status_code=404, detail="Ikona nenalezena.")


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
            "Pravidelný odběr byl bezpečně zaevidován. Seznam odběratelů je neveřejný a chráněný."
            if created
            else "Tato e-mailová adresa je již bezpečně zaevidována k pravidelnému odběru."
        ),
    }


@app.post("/api/newsletter/send-now")
async def send_newsletter_now(request: Request):
    try:
        payload = await request.json()
        email = str(payload.get("email", "")).strip()
        if not is_valid_email_address(email):
            raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")

        forex_data = await asyncio.to_thread(get_live_forex_rates)
        benchmarks = await asyncio.to_thread(get_live_market_benchmarks)
        cnb = get_cnb_macro_indicators()
        html_body = await asyncio.to_thread(generate_html_newsletter_content, forex_data, benchmarks, cnb)
        pdf_bytes = await asyncio.to_thread(build_newsletter_pdf_bytes, forex_data, benchmarks, cnb)

        await asyncio.to_thread(
            send_email_report,
            email,
            "Tensegrity Terminal | Intelligence — Globální a Český Tržní Přehled",
            html_body,
            pdf_bytes,
        )
        return {
            "status": "ok",
            "message": f"Zpravodaj Tensegrity Terminal | Intelligence včetně přílohy byl úspěšně odeslán na adresu {email}.",
        }
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/valuation/{ticker}")
def get_institutional_terminal_valuation(
    ticker: str,
    leverage: float = Query(0.15, ge=0.0, le=1.0),
    growth: float = Query(0.22, gt=-0.95, le=5.0),
    exit_mult: float = Query(24.0, gt=0.5, le=150.0),
):
    try:
        return run_valuation_logic(ticker, leverage, growth, exit_mult)
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.post("/api/advisor/{ticker}")
async def ask_terminal_advisor(ticker: str, request: Request):
    try:
        payload = await request.json()
        question = str(payload.get("question", "")).strip()[:1200]
        history = payload.get("history") or []
        scenario = payload.get("scenario") or {}
        leverage = safe_float(scenario.get("leverage"), 0.15)
        growth = safe_float(scenario.get("growth"), 0.22)
        exit_mult = safe_float(scenario.get("exit_mult"), 24.0)
        if not question:
            raise HTTPException(status_code=422, detail="Zadejte prosím svůj dotaz.")

        d = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)
        answer = await asyncio.to_thread(generate_smart_advisor_answer, question, history, d)
        return {"answer": answer}
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=400, detail=str(err))


@app.get("/api/news/{ticker}")
def get_ticker_news(ticker: str):
    try:
        raw_news = yf.Ticker(ticker.strip().upper() or "NVDA").news or []
        articles = []
        for item in raw_news[:8]:
            content = item.get("content", item) if isinstance(item, dict) else {}
            if not isinstance(content, dict):
                continue
            provider = content.get("provider", {})
            clickthrough = content.get("clickThroughUrl") or content.get("canonicalUrl") or {}
            articles.append({
                "title": content.get("title") or item.get("title") or "Bez názvu",
                "source": provider.get("displayName") or item.get("publisher") or "Finanční zpravodajství",
                "url": clickthrough.get("url") or item.get("link") or "#",
            })
        return {"ticker": ticker.upper(), "articles": articles}
    except Exception as err:
        raise HTTPException(status_code=500, detail=str(err))


@app.get("/api/forex-rates")
async def get_forex_rates():
    return await asyncio.to_thread(get_live_forex_rates)


@app.get("/api/export/excel/{ticker}")
def export_valuation_excel(
    ticker: str,
    leverage: float = Query(0.15, ge=0.0, le=1.0),
    growth: float = Query(0.22, gt=-0.95, le=5.0),
    exit_mult: float = Query(24.0, gt=0.5, le=150.0),
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        wb = openpyxl.Workbook()
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        ws_sum = wb.active
        ws_sum.title = "Syntéza ocenění"
        ws_sum.append(["TENSEGRITY TERMINAL | INSTITUCIONÁLNÍ VALUAČNÍ MODEL"])
        ws_sum["A1"].font = Font(size=14, bold=True)
        ws_sum.append([f"Společnost: {data['company_name']} ({data['ticker']}) | Datum: {data['audit_trail']['timestamp']}"])
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
    leverage: float = Query(0.15, ge=0.0, le=1.0),
    growth: float = Query(0.22, gt=-0.95, le=5.0),
    exit_mult: float = Query(24.0, gt=0.5, le=150.0),
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        title_s = ParagraphStyle("T1", parent=styles["Heading1"], fontName=PDF_FONT_BOLD, fontSize=16, textColor=colors.HexColor("#0f172a"), spaceAfter=6)
        sub_s = ParagraphStyle("S1", parent=styles["Normal"], fontName=PDF_FONT_REGULAR, fontSize=9.5, textColor=colors.HexColor("#475569"), spaceAfter=12)
        sec_s = ParagraphStyle("H2", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=11.5, textColor=colors.HexColor("#0284c7"), spaceBefore=10, spaceAfter=6)

        synth = data["valuation_synthesis"]
        dec = data["decision_engine"]
        dcf = data["dcf_valuation"]
        mc = data["monte_carlo"]
        curr = data["currency"]

        elements = [
            Paragraph(f"TENSEGRITY TERMINAL — VALUAČNÍ ZPRÁVA ({data['company_name']})", title_s),
            Paragraph(f"Ticker: {data['ticker']} | Datum: {data['audit_trail']['timestamp']} | Měna: {curr}", sub_s),
            Paragraph("1. SYNTÉZA VNITŘNÍ HODNOTY A NÁKUPNÍ PÁSMA", sec_s),
        ]

        tbl = Table(
            [
                ["Ukazatel / Oceňovací přístup", "Výstup modelu"],
                ["Aktuální tržní cena akcie", f"{dcf['current_share_price']:,.2f} {curr}"],
                ["Vážená vnitřní hodnota", f"{synth['weighted_fair_value']:,.2f} {curr}"],
                ["Vícefázový výnosový model pro podnik", f"{dcf['intrinsic_value_per_share']:,.2f} {curr}"],
                ["Výnosový model pro vlastní kapitál", f"{data['fcfe_valuation']['intrinsic_value_per_share']:,.2f} {curr}"],
                ["Model reziduálního zisku", f"{data['residual_income']['intrinsic_value_per_share']:,.2f} {curr}"],
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


@app.get("/", response_class=HTMLResponse)
def read_root():
    return """<!DOCTYPE html>
<html lang="cs">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tensegrity Terminal | Institucionální Valuační Platforma</title>
    <link rel="icon" type="image/png" href="/favicon.ico">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
        .brand-title { font-size: 16px; font-weight: 800; letter-spacing: 4px; color: #fff; text-transform: uppercase; }
        .brand-title span { color: var(--accent-blue); }
        .brand-subtitle { font-size: 10px; letter-spacing: 5px; color: var(--text-secondary); text-transform: uppercase; margin-top: 2px; }
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
            color: var(--text-primary); font-family: 'JetBrains Mono', monospace;
            font-size: 14px; font-weight: 700; padding: 10px 14px; border-radius: 10px;
            outline: none; width: 125px; text-transform: uppercase; letter-spacing: 1px;
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

        .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(460px, 1fr)); gap: 22px; margin-bottom: 22px; }
        .grid-3 { display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 22px; margin-bottom: 22px; }
        .grid-4 { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-top: 14px; }
        .card {
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 22px 26px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .matrix-card { margin-bottom: 22px; }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid var(--border-subtle); }
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
        .advisor-thread { display: grid; gap: 12px; margin: 14px 0; max-height: 360px; overflow-y: auto; padding-right: 6px; }
        .advisor-message { padding: 12px 16px; border-left: 3px solid var(--accent-blue); background: rgba(56, 189, 248, 0.06); color: var(--text-secondary); font-size: 13px; line-height: 1.65; white-space: pre-wrap; border-radius: 0 8px 8px 0; }
        .advisor-message.user-message { border-color: var(--success); background: rgba(16, 185, 129, 0.06); color: var(--text-primary); }
        .advisor-form { display: flex; gap: 10px; margin-top: 12px; }
        .advisor-form input { flex: 1; padding: 11px 14px; border: 1px solid var(--border-subtle); border-radius: 8px; background: rgba(6, 9, 17, 0.85); color: var(--text-primary); font-size: 13px; outline: none; }

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
            <div>
                <div class="brand-title">Tensegrity<span>Terminal</span></div>
                <div class="brand-subtitle">Intelligence &bull; Institucionální Valuační Platforma</div>
            </div>
            <div class="controls">
                <div class="search-wrapper">
                    <input type="text" id="tickerInput" value="NVDA" placeholder="NVDA" aria-label="Ticker společnosti">
                </div>
                <div class="forex-strip" id="forexRates">
                    <span>EUR/CZK <strong>--</strong></span>
                    <span>USD/CZK <strong>--</strong></span>
                    <span>GBP/CZK <strong>--</strong></span>
                </div>
                <button class="btn btn-primary" onclick="loadValuation()">Analyzovat</button>
                <button class="btn btn-excel" onclick="downloadExcel()">Excel Model</button>
                <button class="btn btn-pdf" onclick="downloadPDF()">PDF Zpráva</button>
                <button class="btn btn-news" onclick="openNewsletterModal()">Newsletter &amp; Tržní zprávy</button>
            </div>
        </header>

        <section class="scenario-bar" aria-label="Parametry scénáře">
            <span class="scenario-heading">Parametry Scénáře</span>
            <label class="scenario-field" for="leverageInput">
                Cizí kapitál (% Dluhu)
                <span class="scenario-input"><input id="leverageInput" type="number" min="0" max="100" step="1" value="15"><span>%</span></span>
            </label>
            <label class="scenario-field" for="growthInput">
                Růst peněžních toků
                <span class="scenario-input"><input id="growthInput" type="number" min="-50" max="150" step="0.5" value="22"><span>%</span></span>
            </label>
            <label class="scenario-field" for="exitMultipleInput">
                Výstupní násobek
                <span class="scenario-input"><input id="exitMultipleInput" type="number" min="1" max="150" step="0.5" value="24.0"><span>x</span></span>
            </label>
            <button class="btn btn-primary" style="padding: 6px 14px;" onclick="loadValuation()">Přepočítat model</button>
            <span class="audit-pill" id="auditTimestamp">Načítám auditní stopu...</span>
        </section>

        <!-- Rozhodovací rámec a 4 nákupní zóny -->
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

        <!-- Syntéza ocenění -->
        <section class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Syntéza Ocenění — Vážené Spojení Oceňovacích Metod</span>
                <span class="badge" id="weightedFairBadge">Vážená vnitřní hodnota: --</span>
            </div>
            <div class="table-responsive" id="synthesisTableContainer">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
            <p class="plain-guide" id="disagreementNote"></p>
        </section>

        <!-- Živé interaktivní grafy -->
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

        <!-- Fundamenty, Výnosové modely a Scénáře -->
        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Fundamenty &amp; Normalizace</span>
                    <span class="badge" id="fundTicker">NVDA</span>
                </div>
                <div class="metric-row"><span class="metric-label">Tržní kapitalizace:</span><span class="metric-value" id="valCap">-</span></div>
                <div class="metric-row"><span class="metric-label">Hodnota podniku:</span><span class="metric-value highlight" id="valEV">-</span></div>
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

        <!-- Relativní ocenění a Pákový odkup -->
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

        <!-- Citlivostní matice -->
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

        <!-- Analytický poradce -->
        <div class="card">
            <div class="card-header">
                <span class="card-title">Tensegrity Analytický Poradce</span>
                <span class="badge">Kontextová diagnostika a volné dotazy</span>
            </div>
            <div class="advisor-thread" id="advisorThread">
                <div class="advisor-message">Dobrý den. Mám načtena kompletní finanční data, výkazy i oceňovací modely vybrané společnosti. Zeptejte se na cokoliv — například proč valuace vyšla na tuto úroveň, jaká je zadluženost, jaké jsou skryté chyby a rizika firmy, co znamenají jednotlivé ukazatele nebo jaké scénáře můžeme očekávat.</div>
            </div>
            <form class="advisor-form" onsubmit="askAdvisor(event)">
                <input id="advisorQuestion" type="text" placeholder="Napište libovolnou otázku ke společnosti nebo finančním ukazatelům..." required>
                <button class="btn btn-primary" type="submit">Odeslat</button>
            </form>
        </div>
    </div>

    <!-- Modální okno pro Newsletter a Tržní zprávy -->
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
                Pravidelný odběr je automaticky odesílán každou středu v 7:00, každý pátek v 7:00 a každou neděli v 19:00. Evidované adresy jsou uloženy neveřejně.
            </div>

            <div style="border-top: 1px solid var(--border-subtle); padding-top: 14px;">
                <div class="card-title" style="margin-bottom: 8px;">Aktuální tržní zprávy pro vybraný titul</div>
                <div id="newsFeedList">Načítám zprávy...</div>
            </div>
        </div>
    </div>

    <script>
        let advisorHistory = [];
        let methodsChartInstance = null;
        let monteCarloChartInstance = null;
        let cashFlowChartInstance = null;

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

        function renderCharts(d) {
            if (typeof Chart === 'undefined') return;
            const curr = d.currency || 'USD';
            const gridColor = 'rgba(255, 255, 255, 0.06)';
            const tickColor = '#94a3b8';

            // 1. Graf srovnání metod ocenění vs. tržní cena
            const ctx1 = document.getElementById('methodsChartCanvas').getContext('2d');
            if (methodsChartInstance) methodsChartInstance.destroy();
            const methodLabels = ['Tržní cena', 'Vážená hodnota', ...d.valuation_synthesis.methods.map(m => m.method)];
            const methodValues = [
                d.dcf_valuation.current_share_price,
                d.valuation_synthesis.weighted_fair_value,
                ...d.valuation_synthesis.methods.map(m => m.fair_value)
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
                    datasets: [{
                        label: 'Hodnota (' + curr + ')',
                        data: methodValues,
                        backgroundColor: barColors,
                        borderRadius: 6
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { grid: { color: gridColor }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } }
                    }
                }
            });

            // 2. Graf rozdělení Monte Carlo
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
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: { grid: { color: gridColor }, ticks: { color: tickColor, font: { size: 10 } } }
                    }
                }
            });

            // 3. Graf pětileté projekce peněžních toků a splácení dluhu
            const ctx3 = document.getElementById('cashFlowChartCanvas').getContext('2d');
            if (cashFlowChartInstance) cashFlowChartInstance.destroy();
            const cfRows = d.dcf_valuation.cash_flows || [];
            const lboRows = d.projektovane_roky || [];
            cashFlowChartInstance = new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: cfRows.map(r => 'Rok ' + r.rok),
                    datasets: [
                        {
                            label: 'Volný peněžní tok',
                            data: cfRows.map(r => r.volny_penezni_tok / 1e9),
                            backgroundColor: 'rgba(16, 185, 129, 0.75)',
                            borderRadius: 6
                        },
                        {
                            label: 'Zůstatek dluhu (LBO)',
                            data: lboRows.map(r => r.zbyvajici_cizi_kapital / 1e9),
                            type: 'line',
                            borderColor: '#38bdf8',
                            backgroundColor: '#38bdf8',
                            tension: 0.3
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: tickColor, font: { size: 10 } } } },
                    scales: {
                        x: { grid: { display: false }, ticks: { color: tickColor, font: { size: 10 } } },
                        y: {
                            grid: { color: gridColor },
                            ticks: { color: tickColor, font: { size: 10 }, callback: v => v.toFixed(1) + ' mld.' }
                        }
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
            } catch (err) {
                console.error(err);
            }
        }

        async function loadValuation() {
            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            const params = new URLSearchParams({ leverage, growth, exit_mult: exitMult });

            try {
                const res = await fetch('/api/valuation/' + encodeURIComponent(ticker) + '?' + params);
                const d = await res.json();
                if (!res.ok) throw new Error(d.detail || 'Valuaci se nepodařilo načíst.');

                const curr = d.currency || 'USD';
                document.getElementById('tickerInput').value = d.spolecnost;
                document.getElementById('fundTicker').textContent = d.spolecnost;
                document.getElementById('auditTimestamp').textContent = d.audit_trail.model_version + ' | ' + d.audit_trail.timestamp;

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
                document.getElementById('riskScores').textContent = 'Piotroski ' + d.risk_analysis.piotroski_f_score + '/9 | Altman ' + d.risk_analysis.altman_z_score.toFixed(2);

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
                let peerHtml = '<table><thead><tr><th>Metoda</th><th>Násobek</th><th>Metrika</th><th>Implikovaná cena</th></tr></thead><tbody>';
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
                console.error(err);
                alert('Chyba při výpočtu valuace: ' + err.message);
            }
        }

        function openNewsletterModal() {
            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
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
            listEl.innerHTML = '<div style="text-align:center;padding:15px;"><div class="loader"></div></div>';
            try {
                const res = await fetch('/api/news/' + encodeURIComponent(ticker));
                const data = await res.json();
                if (!data.articles || data.articles.length === 0) {
                    listEl.innerHTML = '<p class="plain-guide">Pro tento titul nejsou momentálně dostupné žádné zprávy.</p>';
                } else {
                    listEl.innerHTML = data.articles.map(a =>
                        `<div class="news-item"><a href="${a.url}" target="_blank" rel="noopener">${a.title}</a><div class="news-meta">${a.source}</div></div>`
                    ).join('');
                }
            } catch (err) {
                listEl.innerHTML = '<p class="plain-guide">Nepodařilo se načíst tržní zprávy.</p>';
            }
        }

        async function sendNewsletterNow() {
            const email = document.getElementById('subscribeEmailInput').value.trim();
            const statusEl = document.getElementById('subscribeStatus');
            statusEl.style.color = '#38bdf8';
            statusEl.textContent = 'Generuji zpravodaj s aktuálními daty a odesílám...';
            try {
                const res = await fetch('/api/newsletter/send-now', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Chyba při odesílání.');
                statusEl.style.color = '#34d399';
                statusEl.textContent = data.message;
            } catch (err) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = err.message;
            }
        }

        async function subscribeNewsletter(e) {
            e.preventDefault();
            const emailInput = document.getElementById('subscribeEmailInput');
            const statusEl = document.getElementById('subscribeStatus');
            try {
                const res = await fetch('/subscribe', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email: emailInput.value.trim() })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Chyba při registraci.');
                statusEl.style.color = '#34d399';
                statusEl.textContent = data.message;
                emailInput.value = '';
            } catch (err) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = err.message;
            }
        }

        async function askAdvisor(e) {
            e.preventDefault();
            const input = document.getElementById('advisorQuestion');
            const q = input.value.trim();
            if (!q) return;

            const thread = document.getElementById('advisorThread');
            const userMsg = document.createElement('div');
            userMsg.className = 'advisor-message user-message';
            userMsg.textContent = q;
            thread.appendChild(userMsg);
            input.value = '';
            userMsg.scrollIntoView({ block: 'nearest' });

            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'NVDA';
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
                const replyText = res.ok ? data.answer : (data.detail || 'Chyba poradce.');
                const aiMsg = document.createElement('div');
                aiMsg.className = 'advisor-message';
                aiMsg.textContent = replyText;
                thread.appendChild(aiMsg);
                aiMsg.scrollIntoView({ block: 'nearest' });

                advisorHistory.push({ role: 'user', text: q });
                advisorHistory.push({ role: 'advisor', text: replyText });
                if (advisorHistory.length > 12) advisorHistory = advisorHistory.slice(-12);
            } catch (err) {
                const aiMsg = document.createElement('div');
                aiMsg.className = 'advisor-message';
                aiMsg.textContent = 'Chyba při komunikaci s analytickým poradcem.';
                thread.appendChild(aiMsg);
            }
        }

        function downloadExcel() {
            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'NVDA';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/excel/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        function downloadPDF() {
            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'NVDA';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/pdf/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        document.addEventListener('DOMContentLoaded', () => {
            loadForexRates();
            loadValuation();
        });
    </script>
</body>
</html>
"""