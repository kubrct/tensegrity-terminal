from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse
import asyncio
from concurrent.futures import ThreadPoolExecutor
from html import escape
import json
import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Path
from reportlab.pdfgen import canvas

import io
import os
import smtplib
import ssl
from contextlib import suppress
from datetime import datetime, time, timedelta
from email.message import EmailMessage
from email.utils import parseaddr
from urllib import request as urllib_request
from urllib.parse import quote, urlsplit
from xml.etree import ElementTree

app = FastAPI(title="Tensegrity Terminal API", version="12.1")

dispatch_subscriptions = {}
dispatch_task = None
FIXED_DISPATCH_SLOTS = ((2, time(7, 0)), (4, time(7, 0)), (6, time(19, 0)))
MACRO_NEWS_SOURCES = (
    ("Bloomberg", "global", "https://feeds.bloomberg.com/markets/news.rss", "site:bloomberg.com markets economy"),
    ("Reuters", "global", "https://www.reutersagency.com/feed/?best-topics=business-finance&post_type=best", "site:reuters.com markets economy"),
    ("Financial Times", "global", "https://www.ft.com/rss/home", "site:ft.com markets economy"),
    ("WSJ", "global", "https://feeds.a.dj.com/rss/RSSMarketsMain.xml", "site:wsj.com market economy"),
    ("ČNB", "czech", "https://www.cnb.cz/en/media_service/rss/", "site:cnb.cz ekonomika měnová politika"),
    ("Hospodářské noviny", "czech", "https://hn.cz/?m=rss", "site:hn.cz ekonomika finance"),
)


def get_smtp_config():
    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    try:
        port = int(os.getenv("SMTP_PORT", "587"))
    except ValueError:
        return None
    username = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    sender = os.getenv("SMTP_FROM", username or "")
    if not host or not username or not password or not sender:
        return None
    return host, port, username, password, sender


def next_dispatch_time(now=None):
    now = (now or datetime.now().astimezone()).astimezone()
    for day_offset in range(8):
        candidate_date = now.date() + timedelta(days=day_offset)
        for weekday, scheduled_time in FIXED_DISPATCH_SLOTS:
            if candidate_date.weekday() != weekday:
                continue
            candidate = datetime.combine(candidate_date, scheduled_time, tzinfo=now.tzinfo)
            if candidate > now:
                return candidate
    raise RuntimeError("Nelze určit další pevný čas rozesílky.")


def send_smtp_email(recipient, subject, body, html_body=None, pdf_bytes=None):
    config = get_smtp_config()
    if config is None:
        raise RuntimeError("SMTP není nakonfigurováno.")
    host, port, username, password, sender = config
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    if html_body:
        message.add_alternative(html_body, subtype="html")
    if pdf_bytes:
        message.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename="Tensegrity_Investor_Bulletin.pdf")

    if port == 465:
        with smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context()) as server:
            server.login(username, password)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(username, password)
            server.send_message(message)


def fetch_macro_feed(source_name, region, feed_url, search_query):
    urls = (
        feed_url,
        f"https://news.google.com/rss/search?q={quote(search_query)}&hl=en-US&gl=US&ceid=US:en"
    )
    for url in urls:
        try:
            feed_request = urllib_request.Request(url, headers={"User-Agent": "TensegrityTerminal/1.0"})
            with urllib_request.urlopen(feed_request, timeout=6) as response:
                root = ElementTree.fromstring(response.read(2_000_000))
            entries = root.findall(".//item")
            if not entries:
                entries = root.findall(".//{http://www.w3.org/2005/Atom}entry")
            articles = []
            for entry in entries[:5]:
                title_node = entry.find("title")
                if title_node is None:
                    title_node = entry.find("{http://www.w3.org/2005/Atom}title")
                title = " ".join(title_node.itertext()).strip() if title_node is not None else ""
                link_node = entry.find("link")
                if link_node is None:
                    link_node = entry.find("{http://www.w3.org/2005/Atom}link")
                link = (link_node.get("href") or "") if link_node is not None else ""
                if not link and link_node is not None:
                    link = (link_node.text or "").strip()
                published_node = entry.find("pubDate")
                if published_node is None:
                    published_node = entry.find("{http://www.w3.org/2005/Atom}published")
                if published_node is None:
                    published_node = entry.find("{http://www.w3.org/2005/Atom}updated")
                published = " ".join(published_node.itertext()).strip() if published_node is not None else "Čas neuveden"
                if title:
                    articles.append({"title": title, "source": source_name, "published": published, "url": link, "region": region})
            if articles:
                return articles
        except Exception:
            continue
    return []


def fetch_macro_news():
    with ThreadPoolExecutor(max_workers=len(MACRO_NEWS_SOURCES)) as executor:
        results = list(executor.map(lambda source: fetch_macro_feed(*source), MACRO_NEWS_SOURCES))
    global_news = []
    czech_news = []
    unavailable_sources = []
    for source, articles in zip(MACRO_NEWS_SOURCES, results):
        if not articles:
            unavailable_sources.append(source[0])
        (global_news if source[1] == "global" else czech_news).extend(articles)
    return {"global": global_news[:12], "czech": czech_news[:12], "unavailable_sources": unavailable_sources}


def build_newsletter_html(ticker, newsletter, articles):
    article_markup = []
    for article in articles:
        title = escape(str(article.get("title", "Bez názvu")))
        source = escape(str(article.get("source", "Neznámý zdroj")))
        region = escape(str(article.get("region", "Ticker")))
        published = escape(str(article.get("published", "Čas neuveden")))
        url = str(article.get("url", ""))
        parsed_url = urlsplit(url)
        link_open = f'<a href="{escape(url, quote=True)}" style="color:#075985;text-decoration:none">' if parsed_url.scheme in ("http", "https") else ""
        link_close = "</a>" if link_open else ""
        article_markup.append(
            f'<li style="margin:0 0 12px;padding:12px;border:1px solid #dbe3ec;border-radius:6px">'
            f'<strong>{link_open}{title}{link_close}</strong><br>'
            f'<small>{source} · {region} · {published}</small></li>'
        )
    articles_html = "".join(article_markup) or "<li>Aktuální titulky nejsou dostupné.</li>"
    safe_newsletter = escape(newsletter).replace("\n", "<br>")
    return (
        '<html><body style="margin:0;background:#f1f5f9;font-family:Arial,sans-serif;color:#0f172a">'
        '<main style="max-width:720px;margin:24px auto;padding:24px;background:#fff;border-radius:8px">'
        f'<h1 style="font-size:20px">Tensegrity Investor Bulletin · {escape(ticker)}</h1>'
        f'<section style="line-height:1.6">{safe_newsletter}</section>'
        '<h2 style="font-size:15px;margin-top:24px">Makro zprávy</h2>'
        f'<ul style="padding-left:18px">{articles_html}</ul>'
        '</main></body></html>'
    )


async def send_bulletin_email(recipient, ticker, leverage, growth, exit_mult, private_inputs=None, data=None):
    if data is None:
        data = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult, private_inputs)
    if data.get("private_mode_required"):
        raise RuntimeError(data.get("detail", "Private scénář není kompletní."))
    ticker_news = await asyncio.to_thread(fetch_market_news, ticker)
    macro = await asyncio.to_thread(fetch_macro_news)
    macro_news = macro["global"] + macro["czech"]
    articles = ticker_news + macro_news
    newsletter = await asyncio.to_thread(build_investor_newsletter, data, articles)
    pdf_response = await asyncio.to_thread(
        export_valuation_pdf,
        ticker,
        leverage,
        growth,
        exit_mult,
        private_inputs.get("revenue") if private_inputs else None,
        private_inputs.get("ebitda_margin") if private_inputs else None,
        private_inputs.get("debt") if private_inputs else None,
        private_inputs.get("last_round_valuation") if private_inputs else None
    )
    pdf_chunks = []
    async for chunk in pdf_response.body_iterator:
        pdf_chunks.append(chunk)
    pdf_bytes = b"".join(pdf_chunks)
    html_body = build_newsletter_html(ticker, newsletter, articles)
    await asyncio.to_thread(
        send_smtp_email,
        recipient,
        f"Tensegrity Investor Bulletin: {ticker}",
        newsletter,
        html_body,
        pdf_bytes
    )
    return {"articles": len(articles), "pdf_bytes": len(pdf_bytes)}


def fetch_market_news(ticker):
    try:
        raw_news = yf.Ticker(ticker).news or []
    except Exception:
        return []

    articles = []
    for item in raw_news[:10]:
        content = item.get("content", item) if isinstance(item, dict) else {}
        if not isinstance(content, dict):
            continue
        provider = content.get("provider", {})
        if not isinstance(provider, dict):
            provider = {}
        clickthrough = content.get("clickThroughUrl") or content.get("canonicalUrl") or {}
        if not isinstance(clickthrough, dict):
            clickthrough = {}
        published = content.get("pubDate") or item.get("providerPublishTime")
        if isinstance(published, (int, float)):
            published = datetime.fromtimestamp(published).astimezone().isoformat(timespec="minutes")
        articles.append({
            "title": content.get("title") or item.get("title") or "Bez názvu",
            "source": provider.get("displayName") or item.get("publisher") or "Neznámý zdroj",
            "published": published or "Čas neuveden",
            "url": clickthrough.get("url") or item.get("link") or ""
        })
    return articles


def request_ai_completion(system_prompt, user_prompt):
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None
    payload = {
        "model": os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.3
    }
    http_request = urllib_request.Request(
        os.getenv("OPENAI_API_URL", "https://api.openai.com/v1/chat/completions"),
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST"
    )
    try:
        with urllib_request.urlopen(http_request, timeout=20) as response:
            result = json.loads(response.read().decode("utf-8"))
            return result["choices"][0]["message"]["content"].strip()
    except Exception:
        return None


def build_investor_newsletter(data, news):
    dcf = data["dcf_valuation"]
    lbo = data["model_odkupu_za_cizi_kapital_base_case"]
    comps = data["peer_comps"]
    sotp = data["metoda_souctu_casti"]
    amortization = data["projektovane_roky"]
    valuation_line = (
        f"DCF indikace: {dcf['intrinsic_value_per_share']:.2f} USD/akcii, "
        f"WACC {dcf['wacc']:.2%}, odchylka vůči trhu "
        f"{dcf['upside_downside_percent']:.2f}%."
        if dcf["intrinsic_value_per_share"] is not None and dcf["upside_downside_percent"] is not None
        else "DCF indikace na akcii není dostupná kvůli chybějící tržní ceně nebo počtu akcií."
    )
    headlines = "\n".join(
        f"- {article['title']} ({article['source']}, {article['published']}) {article['url']}"
        for article in news
    ) or "V dostupném zdroji nejsou nové titulky."
    peer_prices = ", ".join(
        f"{peer['method']}: {peer['implied_share_price']:.2f} USD/akcii"
        for peer in comps["multiples"]
        if peer["implied_share_price"] is not None
    ) or "Peer indikace na akcii není dostupná."
    debt_schedule = ", ".join(
        f"Rok {year['rok']} {year['zbyvajici_cizi_kapital']:,.0f} USD"
        for year in amortization
    )
    fallback = (
        f"TENSEGRITY INVESTOR BULLETIN | {data['spolecnost']}\n\n"
        f"Trh: kapitalizace {data['trzni_kapitalizace']:,.0f} USD, dluh {data['celkovy_dluh']:,.0f} USD, EV {data['celkova_hodnota_podniku']:,.0f} USD.\n"
        f"SOTP: EV {sotp['vypoctena_hodnota_podniku']:,.0f} USD, equity {sotp['vypoctena_hodnota_vlastniho_kapitalu']:,.0f} USD.\n"
        f"Shrnutí DCF\n{valuation_line} Terminal Value {dcf['terminal_value']:,.0f} USD.\n"
        f"LBO scénář: MOIC {lbo['moic']:.2f}x, IRR {lbo['irr_procent']:.2f}%.\n"
        f"Dluh po letech: {debt_schedule}.\n"
        f"Sektorové benchmarky ({comps['sector']}): {peer_prices}.\n\n"
        f"Aktuální titulky\n{headlines}\n\n"
        "Tento materiál je informativní modelový výstup, nikoli investiční doporučení. "
        "Sektorové násobky jsou orientační benchmarky."
    )
    ai_context = {
        "ticker": data["spolecnost"],
        "scenario": {key: value for key, value in data["parametry_scenare"].items() if key != "private_inputs"},
        "market_data": {
            "market_cap": data["trzni_kapitalizace"],
            "debt": data["celkovy_dluh"],
            "cash": data["hotovost"],
            "enterprise_value": data["celkova_hodnota_podniku"]
        },
        "sotp": sotp,
        "dcf": dcf,
        "lbo": lbo,
        "five_year_amortization": amortization,
        "peer_comps": comps,
        "headlines": news
    }
    ai_newsletter = request_ai_completion(
        "Jsi finanční editor. Napiš profesionální a věcný český investorský bulletin založený pouze na dodaných údajích. Uveď nejistoty, nepředstírej ověřené zprávy ani jistotu výnosu a zakonči informací, že nejde o osobní investiční doporučení.",
        json.dumps(ai_context, ensure_ascii=False, default=str)
    )
    return ai_newsletter or fallback


async def dispatch_scheduler():
    while True:
        now = datetime.now().astimezone()
        for subscription in list(dispatch_subscriptions.values()):
            if subscription["next_send"] > now:
                continue
            try:
                data = await asyncio.to_thread(
                    run_valuation_logic,
                    subscription["ticker"],
                    subscription["leverage"],
                    subscription["growth"],
                    subscription["exit_mult"],
                    subscription.get("private_inputs")
                )
                if data.get("private_mode_required"):
                    raise RuntimeError("Soukromý scénář je nutné znovu zadat ve formuláři.")
                await send_bulletin_email(subscription["email"], subscription["ticker"], subscription["leverage"], subscription["growth"], subscription["exit_mult"], subscription.get("private_inputs"), data)
                subscription["last_error"] = None
            except Exception as error:
                subscription["last_error"] = str(error)
            subscription["next_send"] = next_dispatch_time()
        await asyncio.sleep(30)


@app.on_event("startup")
async def start_dispatch_scheduler():
    global dispatch_task
    if dispatch_task is None or dispatch_task.done():
        dispatch_task = asyncio.create_task(dispatch_scheduler())


@app.on_event("shutdown")
async def stop_dispatch_scheduler():
    global dispatch_task
    if dispatch_task is not None:
        dispatch_task.cancel()
        with suppress(asyncio.CancelledError):
            await dispatch_task
        dispatch_task = None

# --- REGISTRACE FONTY S ČESKOU DIAKRITIKOU ---
def setup_pdf_fonts():
    font_candidates = [
        ("C:\\Windows\\Fonts\\segoeui.ttf", "C:\\Windows\\Fonts\\segoeuib.ttf", "SegoeUI"),
        ("C:\\Windows\\Fonts\\arial.ttf", "C:\\Windows\\Fonts\\arialbd.ttf", "ArialFont"),
        ("C:\\Windows\\Fonts\\calibri.ttf", "C:\\Windows\\Fonts\\calibrib.ttf", "CalibriFont"),
    ]
    for font_reg, font_bold, font_name in font_candidates:
        if os.path.exists(font_reg) and os.path.exists(font_bold):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_reg))
                pdfmetrics.registerFont(TTFont(f"{font_name}-Bold", font_bold))
                return font_name, f"{font_name}-Bold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"

PDF_FONT_REG, PDF_FONT_BOLD = setup_pdf_fonts()

SECTOR_PEER_BENCHMARKS = {
    "Technology": {"ev_ebitda": 18.0, "pe": 28.0, "ev_revenue": 6.0},
    "Communication Services": {"ev_ebitda": 12.0, "pe": 18.0, "ev_revenue": 3.0},
    "Consumer Cyclical": {"ev_ebitda": 13.0, "pe": 20.0, "ev_revenue": 1.5},
    "Consumer Defensive": {"ev_ebitda": 14.0, "pe": 23.0, "ev_revenue": 1.8},
    "Healthcare": {"ev_ebitda": 16.0, "pe": 24.0, "ev_revenue": 4.0},
    "Financial Services": {"ev_ebitda": 10.0, "pe": 15.0, "ev_revenue": 3.0},
    "Industrials": {"ev_ebitda": 14.0, "pe": 22.0, "ev_revenue": 2.0},
    "Energy": {"ev_ebitda": 7.0, "pe": 12.0, "ev_revenue": 1.4},
    "Utilities": {"ev_ebitda": 11.0, "pe": 17.0, "ev_revenue": 2.0},
    "Real Estate": {"ev_ebitda": 15.0, "pe": 30.0, "ev_revenue": 7.0},
    "Basic Materials": {"ev_ebitda": 9.0, "pe": 15.0, "ev_revenue": 2.0},
}
DEFAULT_PEER_BENCHMARKS = {"ev_ebitda": 14.0, "pe": 22.0, "ev_revenue": 3.0}

# --- VEKTOROVÉ LOGO TENSEGRITY (OPRAVENÁ ORIENTACE Y) ---
def create_pdf_logo(width=45, height=35):
    d = Drawing(width, height)
    
    # P1: Hlavní horní příčka
    p1 = Path(fillColor=colors.HexColor('#0F172A'), strokeColor=None)
    p1.moveTo(3.8, 28.0); p1.lineTo(41.2, 28.0); p1.lineTo(37.5, 25.2); p1.lineTo(7.5, 25.2); p1.closePath()
    
    # P2: Akcentní sekundární příčka
    p2 = Path(fillColor=colors.HexColor('#0284C7'), strokeColor=None)
    p2.moveTo(9.4, 23.8); p2.lineTo(41.2, 23.8); p2.lineTo(39.4, 21.7); p2.lineTo(13.1, 21.7); p2.closePath()
    
    # P3: Hlavní svislá stojka
    p3 = Path(fillColor=colors.HexColor('#0F172A'), strokeColor=None)
    p3.moveTo(15.8, 25.2); p3.lineTo(15.8, 6.3); p3.lineTo(18.8, 3.5); p3.lineTo(18.8, 25.2); p3.closePath()
    
    # P4: Akcentní svislá stojka
    p4 = Path(fillColor=colors.HexColor('#0284C7'), strokeColor=None)
    p4.moveTo(21.8, 21.7); p4.lineTo(21.8, 8.8); p4.lineTo(24.8, 10.9); p4.lineTo(24.8, 21.7); p4.closePath()
    
    d.add(p1); d.add(p2); d.add(p3); d.add(p4)
    return d

# --- ČÍSLOVÁNÍ STRÁNEK A ZÁPATÍ ---
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont(PDF_FONT_REG, 8)
        self.setFillColor(colors.HexColor('#64748B'))
        
        self.setStrokeColor(colors.HexColor('#E2E8F0'))
        self.setLineWidth(0.5)
        self.line(35, 30, 560, 30)
        
        self.drawString(35, 18, "Tensegrity Terminal — Strictly Confidential & Proprietary Valuation")
        page_text = f"Strana {self._pageNumber} z {page_count}"
        self.drawRightString(560, 18, page_text)
        self.restoreState()


@app.get("/", response_class=HTMLResponse)
def read_root():
    return """
    <!DOCTYPE html>
    <html lang="cs">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Tensegrity Terminal | Institucionální Valuační Platforma</title>
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
            :root {
                --bg-main: #060911;
                --bg-card: rgba(15, 22, 37, 0.65);
                --bg-card-hover: rgba(21, 31, 52, 0.8);
                --border-subtle: rgba(255, 255, 255, 0.07);
                --accent-blue: #38bdf8;
                --accent-blue-glow: rgba(56, 189, 248, 0.15);
                --text-primary: #f8fafc;
                --text-secondary: #94a3b8;
                --text-muted: #64748b;
                --success: #10b981;
                --danger: #ef4444;
            }

            * {
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }

            body {
                font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
                background-color: var(--bg-main);
                background-image: 
                    radial-gradient(circle at 50% 0%, rgba(30, 58, 138, 0.18) 0%, transparent 60%),
                    radial-gradient(circle at 85% 30%, rgba(56, 189, 248, 0.05) 0%, transparent 40%);
                background-attachment: fixed;
                color: var(--text-primary);
                min-height: 100vh;
                padding: 32px 40px;
                -webkit-font-smoothing: antialiased;
            }

            .container {
                max-width: 1400px;
                margin: 0 auto;
            }

            header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                background: var(--bg-card);
                backdrop-filter: blur(20px);
                -webkit-backdrop-filter: blur(20px);
                border: 1px solid var(--border-subtle);
                border-radius: 16px;
                padding: 16px 28px;
                margin-bottom: 32px;
                box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
            }

            .scenario-bar {
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 16px 24px;
                background: var(--bg-card);
                border: 1px solid var(--border-subtle);
                border-radius: 12px;
                padding: 14px 22px;
                margin: -12px 0 28px;
            }

            .scenario-heading {
                color: var(--accent-blue);
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 1px;
                white-space: nowrap;
            }

            .scenario-field {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                color: var(--text-secondary);
                font-size: 12px;
            }

            .scenario-input {
                display: inline-flex;
                align-items: center;
                gap: 6px;
                color: var(--text-muted);
                font-family: 'JetBrains Mono', monospace;
            }

            .scenario-input input {
                width: 76px;
                padding: 7px 8px;
                border: 1px solid var(--border-subtle);
                border-radius: 6px;
                outline: none;
                background: rgba(6, 9, 17, 0.7);
                color: var(--text-primary);
                font: 600 13px 'JetBrains Mono', monospace;
                text-align: right;
            }

            .scenario-input input:focus {
                border-color: var(--accent-blue);
                box-shadow: 0 0 10px var(--accent-blue-glow);
            }

            @media (max-width: 760px) {
                .scenario-bar {
                    align-items: flex-start;
                    flex-direction: column;
                    gap: 12px;
                }

                .scenario-field {
                    width: 100%;
                }
            }

            .brand {
                display: flex;
                align-items: center;
                gap: 20px;
            }

            .logo-icon {
                width: 36px;
                height: 36px;
                display: flex;
                align-items: center;
                justify-content: center;
            }

            .brand-divider {
                width: 1px;
                height: 32px;
                background: linear-gradient(180deg, rgba(255,255,255,0.05) 0%, rgba(255,255,255,0.2) 50%, rgba(255,255,255,0.05) 100%);
            }

            .brand-text-block {
                display: flex;
                flex-direction: column;
                justify-content: center;
            }

            .brand-title {
                font-size: 16px;
                font-weight: 800;
                letter-spacing: 5px;
                color: #ffffff;
                line-height: 1.1;
                text-transform: uppercase;
            }

            .brand-subtitle {
                font-size: 10px;
                font-weight: 300;
                letter-spacing: 8px;
                color: var(--text-secondary);
                line-height: 1.2;
                margin-top: 2px;
                text-transform: uppercase;
            }

            .controls {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .search-wrapper input {
                background: rgba(6, 9, 17, 0.7);
                border: 1px solid var(--border-subtle);
                color: var(--text-primary);
                font-family: 'JetBrains Mono', monospace;
                font-size: 14px;
                font-weight: 600;
                padding: 10px 16px;
                border-radius: 10px;
                outline: none;
                width: 130px;
                transition: all 0.2s ease;
                text-transform: uppercase;
                letter-spacing: 1px;
            }

            .search-wrapper input:focus {
                border-color: var(--accent-blue);
                box-shadow: 0 0 12px var(--accent-blue-glow);
            }

            .btn {
                font-family: 'Plus Jakarta Sans', sans-serif;
                font-size: 13px;
                font-weight: 600;
                padding: 10px 20px;
                border-radius: 10px;
                border: 1px solid transparent;
                cursor: pointer;
                display: inline-flex;
                align-items: center;
                gap: 8px;
                transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            }

            .btn-primary {
                background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%);
                color: #040914;
                font-weight: 700;
                box-shadow: 0 4px 14px rgba(56, 189, 248, 0.25);
            }

            .btn-primary:hover {
                opacity: 0.95;
                transform: translateY(-1px);
                box-shadow: 0 6px 20px rgba(56, 189, 248, 0.35);
            }

            .btn-excel {
                background: rgba(16, 185, 129, 0.1);
                border: 1px solid rgba(16, 185, 129, 0.3);
                color: #34d399;
            }

            .btn-excel:hover {
                background: rgba(16, 185, 129, 0.2);
                border-color: #34d399;
                transform: translateY(-1px);
            }

            .btn-pdf {
                background: rgba(239, 68, 68, 0.1);
                border: 1px solid rgba(239, 68, 68, 0.3);
                color: #f87171;
            }

            .btn-pdf:hover {
                background: rgba(239, 68, 68, 0.2);
                border-color: #f87171;
                transform: translateY(-1px);
            }

            .grid-3 {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
                gap: 24px;
                margin-bottom: 28px;
            }

            .card {
                background: var(--bg-card);
                backdrop-filter: blur(16px);
                -webkit-backdrop-filter: blur(16px);
                border: 1px solid var(--border-subtle);
                border-radius: 16px;
                padding: 24px 28px;
                box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
                transition: all 0.25s ease;
            }

            .card:hover {
                border-color: rgba(255, 255, 255, 0.12);
                background: var(--bg-card-hover);
            }

            .card-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                padding-bottom: 12px;
                border-bottom: 1px solid var(--border-subtle);
            }

            .card-title {
                font-size: 13px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 1.5px;
                color: var(--text-secondary);
            }

            .badge {
                font-size: 11px;
                font-family: 'JetBrains Mono', monospace;
                padding: 3px 8px;
                border-radius: 6px;
                background: rgba(56, 189, 248, 0.1);
                color: var(--accent-blue);
                border: 1px solid rgba(56, 189, 248, 0.2);
            }

            .metric-row {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin: 14px 0;
                font-size: 14px;
            }

            .metric-label {
                color: var(--text-secondary);
                font-weight: 400;
            }

            .metric-value {
                font-family: 'JetBrains Mono', monospace;
                font-weight: 600;
                color: var(--text-primary);
                font-size: 15px;
            }

            .metric-value.highlight {
                color: var(--accent-blue);
            }

            .matrix-card {
                margin-bottom: 30px;
            }

            .table-responsive {
                width: 100%;
                overflow-x: auto;
            }

            table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                margin-top: 10px;
            }

            th, td {
                padding: 14px 18px;
                text-align: center;
                font-size: 13px;
                border-bottom: 1px solid var(--border-subtle);
            }

            th {
                background: rgba(6, 9, 17, 0.8);
                color: var(--text-secondary);
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 1px;
                font-size: 12px;
            }

            th:first-child, td:first-child {
                text-align: left;
            }

            .cash-flow-table th:not(:first-child),
            .cash-flow-table td:not(:first-child) {
                text-align: right;
            }

            td {
                font-family: 'JetBrains Mono', monospace;
                color: var(--text-primary);
            }

            tbody tr:hover td {
                background: rgba(56, 189, 248, 0.04);
            }

            .val-positive {
                color: #34d399;
            }

            .val-negative {
                color: #f87171;
            }

            .valuation-status {
                display: inline-flex;
                align-items: center;
                padding: 4px 9px;
                border-radius: 6px;
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
            }

            .valuation-status.undervalued {
                color: #34d399;
                background: rgba(16, 185, 129, 0.12);
                border: 1px solid rgba(16, 185, 129, 0.3);
            }

            .valuation-status.overvalued {
                color: #f87171;
                background: rgba(239, 68, 68, 0.12);
                border: 1px solid rgba(239, 68, 68, 0.3);
            }

            .valuation-status.neutral {
                color: var(--text-secondary);
                background: rgba(148, 163, 184, 0.1);
                border: 1px solid var(--border-subtle);
            }

            .waterfall-scroll {
                width: 100%;
                overflow-x: auto;
                margin: 4px 0 20px;
            }

            .waterfall-scroll svg {
                display: block;
                width: 100%;
                min-width: 700px;
                height: auto;
            }

            .waterfall-scroll text {
                fill: var(--text-secondary);
                font-family: 'JetBrains Mono', monospace;
                font-size: 10px;
            }

            .plain-guide {
                margin-top: 18px;
                padding: 12px 14px;
                border-left: 2px solid var(--accent-blue);
                background: rgba(56, 189, 248, 0.045);
                color: var(--text-secondary);
                font-size: 12px;
                line-height: 1.6;
            }

            .private-panel, .newsletter-output {
                display: none;
            }

            .private-panel {
                margin: -10px 0 28px;
                border-color: rgba(245, 158, 11, 0.35);
            }

            .private-fields {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
                gap: 14px;
            }

            .private-fields label, .dispatch-form label, .advisor-form label {
                display: grid;
                gap: 6px;
                color: var(--text-secondary);
                font-size: 12px;
            }

            .private-fields input, .dispatch-form input, .dispatch-form select, .advisor-form input, .newsletter-output {
                width: 100%;
                padding: 10px 12px;
                border: 1px solid var(--border-subtle);
                border-radius: 6px;
                background: rgba(6, 9, 17, 0.7);
                color: var(--text-primary);
                font: 13px 'JetBrains Mono', monospace;
            }

            .compact-button {
                margin-top: 14px;
            }

            .guide-grid, .news-grid, .utility-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap: 18px;
                margin-bottom: 28px;
            }

            .news-item {
                display: flex;
                flex-direction: column;
                min-height: 132px;
                padding: 16px;
                border: 1px solid var(--border-subtle);
                border-radius: 8px;
                background: rgba(6, 9, 17, 0.38);
            }

            .news-item a {
                color: var(--text-primary);
                font-size: 13px;
                font-weight: 600;
                line-height: 1.5;
                text-decoration: none;
            }

            .news-item a:hover { color: var(--accent-blue); }
            .news-meta { margin-top: auto; padding-top: 12px; color: var(--text-muted); font: 11px 'JetBrains Mono', monospace; }
            .empty-state { color: var(--text-muted); font-size: 13px; }

            .newsletter-output {
                min-height: 180px;
                margin-top: 14px;
                resize: vertical;
                display: block;
                font-family: 'Plus Jakarta Sans', sans-serif;
                line-height: 1.6;
            }

            .dispatch-form, .advisor-form {
                display: grid;
                gap: 12px;
            }

            .dispatch-fields {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 12px;
            }

            .advisor-thread {
                display: grid;
                gap: 10px;
                max-height: 280px;
                overflow-y: auto;
                margin: 14px 0;
            }

            .advisor-message {
                padding: 11px 13px;
                border-left: 2px solid var(--accent-blue);
                background: rgba(56, 189, 248, 0.05);
                color: var(--text-secondary);
                font-size: 13px;
                line-height: 1.55;
                white-space: pre-wrap;
            }

            .advisor-message.user-message { border-color: var(--success); }

            .newsletter-drawer {
                position: fixed;
                inset: 0 0 0 auto;
                z-index: 1002;
                display: flex;
                flex-direction: column;
                width: min(520px, 100vw);
                height: 100dvh;
                padding: 28px;
                overflow-y: auto;
                border-left: 1px solid var(--border-subtle);
                background: rgba(8, 13, 24, 0.96);
                backdrop-filter: blur(24px);
                -webkit-backdrop-filter: blur(24px);
                box-shadow: -24px 0 70px rgba(0, 0, 0, 0.45);
                transform: translateX(105%);
                visibility: hidden;
                transition: transform 280ms cubic-bezier(0.2, 0.8, 0.2, 1), visibility 280ms;
            }

            .newsletter-drawer.open {
                transform: translateX(0);
                visibility: visible;
            }

            .drawer-backdrop {
                position: fixed;
                inset: 0;
                z-index: 1001;
                border: 0;
                background: rgba(0, 0, 0, 0.56);
                opacity: 0;
                visibility: hidden;
                transition: opacity 280ms ease, visibility 280ms;
            }

            .drawer-backdrop.open {
                opacity: 1;
                visibility: visible;
            }

            .drawer-heading {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                gap: 16px;
                margin-bottom: 20px;
            }

            .drawer-heading h2 {
                color: var(--text-primary);
                font-size: 15px;
                font-weight: 700;
            }

            .drawer-close {
                border: 1px solid var(--border-subtle);
                border-radius: 6px;
                padding: 7px 10px;
                background: rgba(255,255,255,0.04);
                color: var(--text-secondary);
                cursor: pointer;
            }

            .macro-section { margin: 8px 0 22px; }
            .macro-section h3 { margin: 0 0 12px; color: var(--accent-blue); font-size: 11px; text-transform: uppercase; }
            .drawer-schedule { margin: 4px 0 14px; color: var(--text-secondary); font-size: 12px; line-height: 1.6; }
            .drawer-divider { height: 1px; margin: 8px 0 18px; border: 0; background: var(--border-subtle); }

            @media (max-width: 760px) {
                header { flex-wrap: wrap; gap: 14px; }
                .controls { width: 100%; flex-wrap: wrap; }
                .controls .btn-news { flex: 1 1 100%; justify-content: center; }
                .newsletter-drawer { padding: 20px; }
            }

            .thesis-layout {
                display: grid;
                grid-template-columns: auto 1fr;
                gap: 24px;
                align-items: center;
            }

            .traffic-lights {
                display: flex;
                gap: 10px;
                padding: 12px;
                border: 1px solid var(--border-subtle);
                border-radius: 999px;
                background: rgba(6, 9, 17, 0.65);
            }

            .traffic-light {
                width: 18px;
                aspect-ratio: 1;
                border-radius: 50%;
                opacity: 0.22;
                box-shadow: inset 0 0 0 1px rgba(255,255,255,0.2);
            }

            .traffic-light.active { opacity: 1; box-shadow: 0 0 16px currentColor; }
            .traffic-light.green { color: #10b981; background: #10b981; }
            .traffic-light.orange { color: #f59e0b; background: #f59e0b; }
            .traffic-light.red { color: #ef4444; background: #ef4444; }

            @media (max-width: 620px) {
                .thesis-layout { grid-template-columns: 1fr; }
                body { padding: 20px 14px; }
            }

            .loader {
                display: inline-block;
                width: 18px;
                height: 18px;
                border: 2px solid rgba(255,255,255,0.2);
                border-radius: 50%;
                border-top-color: var(--accent-blue);
                animation: spin 0.8s linear infinite;
            }

            @keyframes spin {
                to { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>

        <div class="container">
            <header>
                <div class="brand">
                    <div class="logo-icon">
                        <svg viewBox="0 0 120 100" fill="none" xmlns="http://www.w3.org/2000/svg" style="width: 100%; height: 100%;">
                            <path d="M10 20 H110 L100 28 H20 Z" fill="#FFFFFF"/>
                            <path d="M25 32 H110 L105 38 H35 Z" fill="#FFFFFF"/>
                            <path d="M42 28 V82 L50 90 V28 Z" fill="#FFFFFF"/>
                            <path d="M58 38 V75 L66 69 V38 Z" fill="#FFFFFF"/>
                        </svg>
                    </div>
                    <div class="brand-divider"></div>
                    <div class="brand-text-block">
                        <span class="brand-title">TENSEGRITY</span>
                        <span class="brand-subtitle">TERMINAL</span>
                    </div>
                </div>

                <div class="controls">
                    <div class="search-wrapper">
                        <input type="text" id="tickerInput" value="AAPL" placeholder="TICKER...">
                    </div>
                    <button class="btn btn-primary" onclick="loadValuation()">
                        <span>Analyzovat</span>
                    </button>
                    <button class="btn btn-excel" onclick="downloadExcel()">
                        <span>Excel Model</span>
                    </button>
                    <button class="btn btn-pdf" onclick="downloadPDF()">
                        <span>PDF Report</span>
                    </button>
                    <button class="btn btn-excel btn-news" id="newsletterDrawerToggle" aria-controls="newsletterDrawer" aria-expanded="false" onclick="toggleNewsletterDrawer(true)">
                        <span>📰 Newsletter &amp; News Feed</span>
                    </button>
                </div>
            </header>

            <button class="drawer-backdrop" id="newsletterDrawerBackdrop" aria-label="Zavřít news feed" onclick="toggleNewsletterDrawer(false)"></button>
            <aside class="newsletter-drawer" id="newsletterDrawer" aria-label="Newsletter & News Feed" aria-hidden="true">
                <div class="drawer-heading">
                    <div>
                        <h2>Newsletter &amp; News Feed</h2>
                        <p class="drawer-schedule">Makro výběr podle dostupných RSS zdrojů. Týdenní bulletiny: středa 07:00, pátek 07:00, neděle 19:00 (čas serveru).</p>
                        <span class="badge" id="newsTicker">-</span>
                    </div>
                    <button class="drawer-close" type="button" aria-label="Zavřít" onclick="toggleNewsletterDrawer(false)">Zavřít</button>
                </div>
                <section class="macro-section" aria-labelledby="tickerNewsTitle">
                    <h3 id="tickerNewsTitle">Zprávy k analyzovanému tickeru</h3>
                    <div class="news-grid" id="marketNewsContainer"><p class="empty-state">Načtěte analýzu tickeru.</p></div>
                </section>
                <section class="macro-section" aria-labelledby="globalNewsTitle">
                    <h3 id="globalNewsTitle">Globální trhy</h3>
                    <div class="news-grid" id="globalMacroNews"><p class="empty-state">Načítám globální zprávy…</p></div>
                </section>
                <section class="macro-section" aria-labelledby="czechNewsTitle">
                    <h3 id="czechNewsTitle">Česká republika</h3>
                    <div class="news-grid" id="czechMacroNews"><p class="empty-state">Načítám zprávy z ČR…</p></div>
                </section>
                <hr class="drawer-divider">
                <form class="dispatch-form" onsubmit="subscribeDispatch(event)">
                    <label for="dispatchEmail">E-mailová adresa pro bulletin<input id="dispatchEmail" type="email" autocomplete="email" required></label>
                    <button class="btn btn-excel" type="submit">Přihlásit k pevnému rozesílání</button>
                    <button class="btn btn-primary" type="button" onclick="sendCurrentBulletinNow()">Odeslat aktuální bulletin ihned</button>
                    <button class="btn btn-pdf" type="button" onclick="unsubscribeDispatch()">Zrušit odběr pro tento e-mail</button>
                    <p class="empty-state" id="dispatchStatus">Vyžaduje SMTP_HOST, SMTP_USER a SMTP_PASSWORD. Odběry jsou uloženy v paměti této instance.</p>
                </form>
                <hr class="drawer-divider">
                <button class="btn btn-primary" type="button" onclick="generateNewsletter()">Generovat Investor Newsletter</button>
                <textarea class="newsletter-output" id="newsletterOutput" readonly aria-label="Investor newsletter"></textarea>
            </aside>

            <section class="scenario-bar" aria-label="Scenario Parameters">
                <span class="scenario-heading">Scenario Parameters</span>
                <label class="scenario-field" for="leverageInput">
                    Transakční cizí kapitál (% Dluhu)
                    <span class="scenario-input"><input id="leverageInput" type="number" min="0" max="100" step="1" value="30" required><span>%</span></span>
                </label>
                <label class="scenario-field" for="growthInput">
                    Očekávané tempo růstu EBITDA
                    <span class="scenario-input"><input id="growthInput" type="number" min="-99" max="500" step="0.5" value="5" required><span>%</span></span>
                </label>
                <label class="scenario-field" for="exitMultipleInput">
                    Výstupní Exit EV/EBITDA násobek
                    <span class="scenario-input"><input id="exitMultipleInput" type="number" min="0.1" max="100" step="0.1" value="13.5" required><span>x</span></span>
                </label>
            </section>

            <section class="card matrix-card" id="investmentThesis">
                <div class="card-header">
                    <span class="card-title">AI Investment Thesis & Semafor</span>
                    <span class="badge" id="thesisTicker">-</span>
                </div>
                <div class="thesis-layout">
                    <div class="traffic-lights" aria-label="Investiční semafor">
                        <span class="traffic-light green" id="thesisGreen"></span>
                        <span class="traffic-light orange" id="thesisOrange"></span>
                        <span class="traffic-light red" id="thesisRed"></span>
                    </div>
                    <div>
                        <p class="metric-value" id="thesisRecommendation">Čeká na analýzu</p>
                        <p class="plain-guide" id="thesisSummary">Panel shrne ocenění vůči trhu a hlavní citlivost modelu. Nejde o osobní investiční doporučení.</p>
                    </div>
                </div>
            </section>

            <section class="card private-panel" id="privatePanel">
                <div class="card-header">
                    <span class="card-title">Soukromá společnost: ruční odhad</span>
                    <span class="badge">Private Mode</span>
                </div>
                <p class="empty-state" id="privateModeMessage">Veřejná data nejsou dostupná. Zadejte odhadované hodnoty; soukromá firma nemá ověřenou tržní cenu ani počet akcií.</p>
                <div class="private-fields">
                    <label for="privateRevenue">Roční tržby (USD)<input id="privateRevenue" type="number" min="0" step="1000000" required></label>
                    <label for="privateEbitdaMargin">EBITDA marže (%)<input id="privateEbitdaMargin" type="number" min="-100" max="100" step="0.5" required></label>
                    <label for="privateDebt">Celkový dluh (USD)<input id="privateDebt" type="number" min="0" step="1000000" required></label>
                    <label for="privateLastRound">Valuace posledního kola (USD)<input id="privateLastRound" type="number" min="1" step="1000000" required></label>
                </div>
                <button class="btn btn-primary compact-button" onclick="loadValuation()">Přepočítat soukromý scénář</button>
            </section>

            <div class="grid-3">
                <div class="card">
                    <div class="card-header">
                        <span class="card-title">Základní Tržní Data</span>
                        <span class="badge" id="valTicker">AAPL</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Tržní Kapitalizace:</span>
                        <span class="metric-value" id="valCap">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Celkový Dluh:</span>
                        <span class="metric-value" id="valDebt">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Celková Hodnota Podniku:</span>
                        <span class="metric-value highlight" id="valEV">-</span>
                    </div>
                </div>

                <div class="card">
                    <div class="card-header">
                        <span class="card-title">Metoda Součtu Částí (SOTP)</span>
                        <span class="badge">Ocenění</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Ocenění Podniku (EV):</span>
                        <span class="metric-value" id="sotpEV">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Ocenění Vlastního Kapitálu:</span>
                        <span class="metric-value highlight" id="sotpEquity">-</span>
                    </div>
                    <p class="plain-guide">SOTP skládá firmu z hodnoty jednotlivých částí. Berte ho jako odhad hodnoty provozů po odečtení dluhu, ne jako cenu, za kterou lze akcii určitě prodat.</p>
                </div>

                <div class="card">
                    <div class="card-header">
                        <span class="card-title">Model Odkupu za Cizí Kapitál (LBO)</span>
                        <span class="badge">Base Case</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Vstupní Vlastní Kapitál:</span>
                        <span class="metric-value" id="lboInEq">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Výstupní Vlastní Kapitál:</span>
                        <span class="metric-value" id="lboOutEq">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Násobek Kapitálu (MOIC):</span>
                        <span class="metric-value highlight" id="lboMoic">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Vnitřní Míra Výnosnosti (IRR):</span>
                        <span class="metric-value highlight" id="lboIrr">-</span>
                    </div>
                    <p class="plain-guide">LBO ukazuje výnos kupujícího při financování části nákupu dluhem. MOIC říká, kolikrát se vrátí vložený kapitál; IRR zohledňuje i čas. Obě čísla stojí na předpokladech scénáře.</p>
                </div>
            </div>

            <div class="card matrix-card">
                <div class="card-header">
                    <span class="card-title">LBO Odbourávání Cizího Kapitálu</span>
                    <span class="badge">Waterfall</span>
                </div>
                <div class="waterfall-scroll" id="debtWaterfallContainer"></div>
            </div>

            <div class="grid-3">
                <div class="card">
                    <div class="card-header">
                        <span class="card-title">Diskontované Cash Flow (DCF Ocenění)</span>
                        <span class="badge">Exit Multiple Method</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Vnitřní hodnota na akcii:</span>
                        <span class="metric-value highlight" id="dcfIntrinsicValue">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">WACC:</span>
                        <span class="metric-value" id="dcfWacc">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Terminální hodnota:</span>
                        <span class="metric-value" id="dcfTerminalValue">-</span>
                    </div>
                    <div class="metric-row">
                        <span class="metric-label">Ocenění vůči tržní ceně:</span>
                        <span class="valuation-status" id="dcfValuationStatus">-</span>
                    </div>
                    <p class="plain-guide">DCF převádí budoucí hotovost firmy na dnešní hodnotu. WACC je použitá požadovaná výnosnost; vyšší WACC obvykle snižuje ocenění. Výsledek je citlivý na dlouhodobé předpoklady.</p>
                </div>

                <div class="card">
                    <div class="card-header">
                        <span class="card-title">Srovnání Konkurence (Trading Comps)</span>
                        <span class="badge" id="peerSector">Sektorový benchmark</span>
                    </div>
                    <div class="table-responsive" id="peerCompsContainer"></div>
                    <p class="plain-guide">Trading Comps porovnává firmu s orientačními násobky sektoru. Je to rychlé srovnání, ne skutečný seznam ověřených konkurentů; rozdíly v růstu a riziku mohou výsledek výrazně měnit.</p>
                </div>
            </div>

            <div class="card matrix-card">
                <div class="card-header">
                    <span class="card-title">5-Letý LBO Amortizační Kalendář a Cash Flow</span>
                    <span class="badge">Rok 1 až Rok 5</span>
                </div>
                <div class="table-responsive" id="cashFlowContainer"></div>
            </div>

            <div class="card matrix-card">
                <div class="card-header">
                    <span class="card-title">Citlivostní Matice Vnitřní Míry Výnosnosti (IRR %)</span>
                    <span class="badge">EBITDA Růst vs. Výstupní Násobek</span>
                </div>
                <div class="table-responsive" id="matrixContainer">
                    <div style="text-align: center; padding: 30px; color: var(--text-muted);">
                        <div class="loader"></div>
                        <div style="margin-top: 10px;">Načítám finanční data...</div>
                    </div>
                </div>
            </div>

            <div class="utility-grid">
                <section class="card">
                    <div class="card-header">
                        <span class="card-title">Terminal AI Advisor</span>
                        <span class="badge">Aktuální model</span>
                    </div>
                    <div class="advisor-thread" id="advisorThread" aria-live="polite"></div>
                    <p class="empty-state">S OPENAI_API_KEY se odešle otázka a souhrn modelu nakonfigurovanému AI poskytovateli. Bez klíče odpovídá lokální analytická nápověda.</p>
                    <form class="advisor-form" onsubmit="askAdvisor(event)">
                        <label for="advisorQuestion">Otázka k aktuální valuaci</label>
                        <input id="advisorQuestion" type="text" maxlength="1000" placeholder="Např. Co nejvíce ovlivňuje DCF?" required>
                        <button class="btn btn-primary" type="submit">Zeptat se</button>
                    </form>
                </section>
            </div>

        </div>

        <script>
            let latestValuation = null;
            let latestNews = [];

            function formatCurrency(value) {
                if (value === null || value === undefined || !Number.isFinite(Number(value))) {
                    return 'N/A';
                }
                return Number(value).toLocaleString(undefined, { maximumFractionDigits: 2 }) + ' USD';
            }

            function getScenarioQuery() {
                const controls = [
                    document.getElementById('leverageInput'),
                    document.getElementById('growthInput'),
                    document.getElementById('exitMultipleInput')
                ];
                if (!controls.every(control => control.reportValidity())) {
                    throw new Error('Zkontrolujte zadané parametry scénáře.');
                }

                const leverage = controls[0].valueAsNumber / 100;
                const growth = controls[1].valueAsNumber / 100;
                const exitMultiple = controls[2].valueAsNumber;
                return new URLSearchParams({
                    leverage: leverage.toString(),
                    growth: growth.toString(),
                    exit_mult: exitMultiple.toString()
                }).toString();
            }

            function getPrivateQuery() {
                const controls = [
                    document.getElementById('privateRevenue'),
                    document.getElementById('privateEbitdaMargin'),
                    document.getElementById('privateDebt'),
                    document.getElementById('privateLastRound')
                ];
                if (controls.every(control => control.value.trim() === '')) return '';
                if (!controls.every(control => control.reportValidity())) {
                    throw new Error('Doplňte všechny čtyři údaje soukromé společnosti.');
                }
                return new URLSearchParams({
                    private_revenue: controls[0].valueAsNumber.toString(),
                    private_ebitda_margin: (controls[1].valueAsNumber / 100).toString(),
                    private_debt: controls[2].valueAsNumber.toString(),
                    private_last_round_valuation: controls[3].valueAsNumber.toString()
                }).toString();
            }

            function getFullScenarioQuery() {
                const privateQuery = getPrivateQuery();
                return getScenarioQuery() + (privateQuery ? '&' + privateQuery : '');
            }

            function appendAdvisorMessage(text, isUser = false) {
                const message = document.createElement('div');
                message.className = 'advisor-message' + (isUser ? ' user-message' : '');
                message.textContent = text;
                document.getElementById('advisorThread').appendChild(message);
                message.scrollIntoView({ block: 'nearest' });
            }

            function renderMarketNews(articles, ticker) {
                const container = document.getElementById('marketNewsContainer');
                container.replaceChildren();
                document.getElementById('newsTicker').textContent = ticker;
                if (!articles.length) {
                    const empty = document.createElement('p');
                    empty.className = 'empty-state';
                    empty.textContent = 'Pro tento ticker nejsou dostupné aktuální titulky.';
                    container.appendChild(empty);
                    return;
                }
                articles.forEach(article => {
                    const card = document.createElement('article');
                    card.className = 'news-item';
                    const title = document.createElement(article.url ? 'a' : 'p');
                    title.textContent = article.title;
                    if (article.url) {
                        try {
                            const articleUrl = new URL(article.url);
                            if (articleUrl.protocol === 'https:' || articleUrl.protocol === 'http:') {
                                title.href = articleUrl.href;
                                title.target = '_blank';
                                title.rel = 'noopener noreferrer';
                            }
                        } catch (error) {
                            title.removeAttribute('href');
                        }
                    }
                    const meta = document.createElement('p');
                    meta.className = 'news-meta';
                    meta.textContent = `${article.source} · ${article.published}`;
                    card.append(title, meta);
                    container.appendChild(card);
                });
            }

            function renderMacroNews(articles, containerId, unavailableSources) {
                const container = document.getElementById(containerId);
                container.replaceChildren();
                if (!articles.length) {
                    const empty = document.createElement('p');
                    empty.className = 'empty-state';
                    empty.textContent = 'Zdroje nyní neposkytly žádné dostupné titulky.';
                    container.appendChild(empty);
                }
                articles.forEach(article => {
                    const card = document.createElement('article');
                    card.className = 'news-item';
                    const title = document.createElement(article.url ? 'a' : 'p');
                    title.textContent = article.title;
                    if (article.url) {
                        try {
                            const articleUrl = new URL(article.url);
                            if (articleUrl.protocol === 'https:' || articleUrl.protocol === 'http:') {
                                title.href = articleUrl.href;
                                title.target = '_blank';
                                title.rel = 'noopener noreferrer';
                            }
                        } catch (error) {
                            title.removeAttribute('href');
                        }
                    }
                    const meta = document.createElement('p');
                    meta.className = 'news-meta';
                    meta.textContent = `${article.source} · ${article.published}`;
                    card.append(title, meta);
                    container.appendChild(card);
                });
                if (unavailableSources.length) {
                    const note = document.createElement('p');
                    note.className = 'empty-state';
                    note.textContent = 'Dočasně nedostupné zdroje: ' + unavailableSources.join(', ');
                    container.appendChild(note);
                }
            }

            function toggleNewsletterDrawer(open) {
                const drawer = document.getElementById('newsletterDrawer');
                const backdrop = document.getElementById('newsletterDrawerBackdrop');
                drawer.classList.toggle('open', open);
                backdrop.classList.toggle('open', open);
                drawer.setAttribute('aria-hidden', String(!open));
                document.getElementById('newsletterDrawerToggle').setAttribute('aria-expanded', String(open));
                document.body.style.overflow = open ? 'hidden' : '';
                if (open) {
                    loadMacroNews();
                    document.querySelector('.drawer-close').focus();
                }
            }

            async function loadMacroNews() {
                try {
                    const response = await fetch('/api/macro-news');
                    const macro = await response.json();
                    if (!response.ok) throw new Error(macro.detail || 'Makro zprávy nejsou dostupné.');
                    renderMacroNews(macro.global || [], 'globalMacroNews', macro.unavailable_sources || []);
                    renderMacroNews(macro.czech || [], 'czechMacroNews', macro.unavailable_sources || []);
                } catch (error) {
                    renderMacroNews([], 'globalMacroNews', []);
                    renderMacroNews([], 'czechMacroNews', []);
                }
            }

            function renderInvestmentThesis(data) {
                const dcf = data.dcf_valuation;
                const upside = dcf.upside_downside_percent;
                const lights = ['thesisGreen', 'thesisOrange', 'thesisRed'];
                lights.forEach(id => document.getElementById(id).classList.remove('active'));
                document.getElementById('thesisTicker').textContent = data.spolecnost;
                if (upside === null) {
                    document.getElementById('thesisOrange').classList.add('active');
                    document.getElementById('thesisRecommendation').textContent = 'Nedostatek tržních dat';
                    document.getElementById('thesisSummary').textContent = 'Model nemá ověřenou aktuální cenu ani počet akcií. Poslední investiční kolo soukromé firmy není živá tržní cena.';
                    return;
                }

                const color = upside >= 15 ? 'green' : upside <= -15 ? 'red' : 'orange';
                document.getElementById(`thesis${color[0].toUpperCase()}${color.slice(1)}`).classList.add('active');
                const recommendation = upside >= 15 ? 'Model indikuje podhodnocení' : upside <= -15 ? 'Model indikuje předražení' : 'Model indikuje pásmo fair value';
                document.getElementById('thesisRecommendation').textContent = recommendation;
                document.getElementById('thesisSummary').textContent = `DCF hodnota ${formatCurrency(dcf.intrinsic_value_per_share)} proti tržní ceně ${formatCurrency(dcf.current_share_price)} (${upside.toFixed(2)} %). Výsledek je citlivý na WACC ${ (dcf.wacc * 100).toFixed(2) } % a exit násobek ${data.parametry_scenare.exit_mult.toFixed(1)}x. Jde o modelový signál, ne osobní doporučení.`;
            }

            function renderDebtWaterfall(data) {
                const years = data.projektovane_roky;
                const compactCurrency = value => new Intl.NumberFormat(undefined, {
                    notation: 'compact',
                    maximumFractionDigits: 1
                }).format(value);
                const initialDebt = Number(data.model_odkupu_za_cizi_kapital_base_case.vstupni_cizi_kapital);
                const maximumDebt = Math.max(initialDebt, ...years.map(year => Number(year.zbyvajici_cizi_kapital)), 1);
                const baseline = 148;
                const scale = 108 / maximumDebt;
                const barWidth = 42;
                const positions = [30, 138, 246, 354, 462, 570];
                let svg = '<svg viewBox="0 0 690 190" role="img" aria-label="Waterfall vývoje zůstatku dluhu v USD">';
                svg += `<line x1="18" y1="${baseline}" x2="670" y2="${baseline}" stroke="rgba(148,163,184,0.35)"/>`;
                const initialHeight = Math.max(2, initialDebt * scale);
                svg += `<rect x="${positions[0]}" y="${baseline - initialHeight}" width="${barWidth}" height="${initialHeight}" rx="3" fill="#38bdf8" opacity="0.55"/>`;
                svg += `<text x="${positions[0] + barWidth / 2}" y="${baseline + 18}" text-anchor="middle">Vstup</text>`;
                svg += `<text x="${positions[0] + barWidth / 2}" y="${Math.max(14, baseline - initialHeight - 7)}" text-anchor="middle">${compactCurrency(initialDebt)}</text>`;

                let previousDebt = initialDebt;
                years.forEach((year, index) => {
                    const repayment = Number(year.mor_jistiny);
                    const remainingDebt = Number(year.zbyvajici_cizi_kapital);
                    const x = positions[index + 1];
                    const top = baseline - previousDebt * scale;
                    const bottom = baseline - remainingDebt * scale;
                    const height = Math.max(2, bottom - top);
                    svg += `<line x1="${positions[index] + barWidth}" y1="${baseline - previousDebt * scale}" x2="${x}" y2="${baseline - previousDebt * scale}" stroke="rgba(148,163,184,0.55)" stroke-dasharray="3 3"/>`;
                    svg += `<rect x="${x}" y="${top}" width="${barWidth}" height="${height}" rx="3" fill="#34d399"/>`;
                    svg += `<text x="${x + barWidth / 2}" y="${Math.max(14, top - 7)}" text-anchor="middle">-${compactCurrency(repayment)}</text>`;
                    svg += `<text x="${x + barWidth / 2}" y="${baseline + 18}" text-anchor="middle">Rok ${year.rok}</text>`;
                    svg += `<text x="${x + barWidth / 2}" y="${baseline + 34}" text-anchor="middle">${compactCurrency(remainingDebt)} zbývá</text>`;
                    previousDebt = remainingDebt;
                });
                svg += '</svg>';
                document.getElementById('debtWaterfallContainer').innerHTML = svg;
            }

            async function loadValuation() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                const container = document.getElementById('matrixContainer');
                container.innerHTML = '<div style="text-align: center; padding: 30px; color: var(--text-muted);"><div class="loader"></div><div style="margin-top: 10px;">Načítám data pro ' + ticker + '...</div></div>';

                try {
                    const response = await fetch(`/api/valuation/${encodeURIComponent(ticker)}?${getFullScenarioQuery()}`);
                    const data = await response.json();
                    if (data.private_mode_required) {
                        latestValuation = null;
                        document.getElementById('privatePanel').style.display = 'block';
                        document.getElementById('privateModeMessage').textContent = data.detail;
                        renderMarketNews(data.market_news || [], ticker);
                        return;
                    }
                    if (!response.ok) throw new Error(data.detail || 'Valuaci se nepodařilo načíst.');
                    latestValuation = data;
                    latestNews = data.market_news || [];
                    document.getElementById('privatePanel').style.display = 'none';

                    document.getElementById('valTicker').innerText = data.spolecnost;
                    document.getElementById('valCap').innerText = Number(data.trzni_kapitalizace).toLocaleString() + " USD";
                    document.getElementById('valDebt').innerText = Number(data.celkovy_dluh).toLocaleString() + " USD";
                    document.getElementById('valEV').innerText = Number(data.celkova_hodnota_podniku).toLocaleString() + " USD";

                    document.getElementById('sotpEV').innerText = Number(data.metoda_souctu_casti.vypoctena_hodnota_podniku).toLocaleString() + " USD";
                    document.getElementById('sotpEquity').innerText = Number(data.metoda_souctu_casti.vypoctena_hodnota_vlastniho_kapitalu).toLocaleString() + " USD";

                    document.getElementById('lboInEq').innerText = Number(data.model_odkupu_za_cizi_kapital_base_case.vstupni_vlastni_kapital).toLocaleString() + " USD";
                    document.getElementById('lboOutEq').innerText = Number(data.model_odkupu_za_cizi_kapital_base_case.vystupni_hodnota_vlastniho_kapitalu).toLocaleString() + " USD";
                    document.getElementById('lboMoic').innerText = data.model_odkupu_za_cizi_kapital_base_case.moic + "x";
                    document.getElementById('lboIrr').innerText = data.model_odkupu_za_cizi_kapital_base_case.irr_procent + " %";

                    const dcf = data.dcf_valuation;
                    document.getElementById('dcfIntrinsicValue').innerText = formatCurrency(dcf.intrinsic_value_per_share);
                    document.getElementById('dcfWacc').innerText = (dcf.wacc * 100).toFixed(2) + ' %';
                    document.getElementById('dcfTerminalValue').innerText = formatCurrency(dcf.terminal_value);
                    const status = document.getElementById('dcfValuationStatus');
                    const upside = dcf.upside_downside_percent;
                    if (upside === null) {
                        status.className = 'valuation-status neutral';
                        status.innerText = 'Tržní cena nedostupná';
                    } else {
                        status.className = 'valuation-status ' + (upside >= 0 ? 'undervalued' : 'overvalued');
                        status.innerText = (upside >= 0 ? 'Podhodnoceno ' : 'Předraženo ') + Math.abs(upside).toFixed(2) + ' %';
                    }

                    document.getElementById('peerSector').innerText = 'Orientační · ' + data.peer_comps.sector;
                    let peerHtml = '<table><thead><tr><th>Metoda</th><th>Sektorový násobek</th><th>Metrika (rok 5)</th><th>Implikovaná cena/akcii</th></tr></thead><tbody>';
                    data.peer_comps.multiples.forEach(peer => {
                        peerHtml += `<tr><td>${peer.method}</td><td>${peer.multiple.toFixed(1)}x</td><td>${formatCurrency(peer.forward_metric)}</td><td>${formatCurrency(peer.implied_share_price)}</td></tr>`;
                    });
                    peerHtml += '</tbody></table>';
                    document.getElementById('peerCompsContainer').innerHTML = peerHtml;
                    renderMarketNews(latestNews, ticker);

                    const cashFlowContainer = document.getElementById('cashFlowContainer');
                    let cashFlowHtml = '<table class="cash-flow-table"><thead><tr><th>Rok</th><th>Projektovaná EBITDA</th><th>Volné Cash Flow pro splácení (FCF)</th><th>Splátka jistiny dluhu</th><th>Zůstatek cizího kapitálu</th></tr></thead><tbody>';
                    data.projektovane_roky.forEach(row => {
                        cashFlowHtml += `<tr><td>Rok ${row.rok}</td><td>${Number(row.projektovany_provozni_zisk_pred_odpisy).toLocaleString()} USD</td><td>${Number(row.volny_tok_pro_splaceni).toLocaleString()} USD</td><td>${Number(row.mor_jistiny).toLocaleString()} USD</td><td>${Number(row.zbyvajici_cizi_kapital).toLocaleString()} USD</td></tr>`;
                    });
                    cashFlowHtml += '</tbody></table>';
                    cashFlowContainer.innerHTML = cashFlowHtml;
                    renderDebtWaterfall(data);
                    renderInvestmentThesis(data);

                    let html = '<table><thead><tr><th>Růst \\ Multiplikátor</th>';
                    const firstRow = data.citlivostni_matice_lbo_irr[0].vysledky_irr_dle_nasobku;
                    firstRow.forEach(item => {
                        html += `<th>${item.vystupni_nasobek}x</th>`;
                    });
                    html += '</tr></thead><tbody>';

                    data.citlivostni_matice_lbo_irr.forEach(row => {
                        html += `<tr><td style="font-weight: 700; color: var(--text-secondary);">${row.tempo_rustu_ebitda_procent}%</td>`;
                        row.vysledky_irr_dle_nasobku.forEach(cell => {
                            const valClass = cell.irr_procent >= 0 ? 'val-positive' : 'val-negative';
                            html += `<td class="${valClass}">${cell.irr_procent}%</td>`;
                        });
                        html += '</tr>';
                    });
                    html += '</tbody></table>';

                    container.innerHTML = html;

                } catch (error) {
                    container.textContent = 'Chyba při načítání dat: ' + error.message;
                }
            }

            async function generateNewsletter() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                const output = document.getElementById('newsletterOutput');
                output.style.display = 'block';
                output.value = 'Připravuji přehled z aktuálních dat a titulků…';
                try {
                    const response = await fetch(`/api/newsletter/${encodeURIComponent(ticker)}?${getFullScenarioQuery()}`);
                    const result = await response.json();
                    if (!response.ok) throw new Error(result.detail || 'Newsletter se nepodařilo vytvořit.');
                    output.value = result.newsletter;
                } catch (error) {
                    output.value = 'Chyba: ' + error.message;
                }
            }

            async function subscribeDispatch(event) {
                event.preventDefault();
                const status = document.getElementById('dispatchStatus');
                status.textContent = 'Odesílám plán do backendu…';
                try {
                    const response = await fetch('/api/dispatch/subscribe', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            email: document.getElementById('dispatchEmail').value,
                            ticker: document.getElementById('tickerInput').value.trim().toUpperCase(),
                            scenario: {
                                leverage: document.getElementById('leverageInput').valueAsNumber / 100,
                                growth: document.getElementById('growthInput').valueAsNumber / 100,
                                exit_mult: document.getElementById('exitMultipleInput').valueAsNumber
                            },
                            private_inputs: collectPrivateInputValues()
                        })
                    });
                    const result = await response.json();
                    if (!response.ok) throw new Error(result.detail || 'Plánování se nezdařilo.');
                    status.textContent = `Přihlášení je aktivní. Další bulletin: ${result.next_send} (čas serveru).`;
                } catch (error) {
                    status.textContent = 'Chyba: ' + error.message;
                }
            }

            async function sendCurrentBulletinNow() {
                const status = document.getElementById('dispatchStatus');
                status.textContent = 'Sestavuji HTML bulletin a PDF přílohu…';
                try {
                    const response = await fetch('/api/dispatch/send-now', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            email: document.getElementById('dispatchEmail').value,
                            ticker: document.getElementById('tickerInput').value.trim().toUpperCase(),
                            scenario: {
                                leverage: document.getElementById('leverageInput').valueAsNumber / 100,
                                growth: document.getElementById('growthInput').valueAsNumber / 100,
                                exit_mult: document.getElementById('exitMultipleInput').valueAsNumber
                            },
                            private_inputs: collectPrivateInputValues()
                        })
                    });
                    const result = await response.json();
                    if (!response.ok) throw new Error(result.detail || 'Odeslání se nezdařilo.');
                    status.textContent = `Bulletin včetně PDF byl odeslán (${result.articles} zpráv).`;
                } catch (error) {
                    status.textContent = 'Chyba: ' + error.message;
                }
            }

            async function unsubscribeDispatch() {
                const status = document.getElementById('dispatchStatus');
                const email = document.getElementById('dispatchEmail').value.trim();
                try {
                    const response = await fetch('/api/dispatch/unsubscribe', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ email })
                    });
                    const result = await response.json();
                    if (!response.ok) throw new Error(result.detail || 'Odběr se nepodařilo zrušit.');
                    status.textContent = result.removed ? 'Plánované odesílání bylo zrušeno.' : 'Pro tento e-mail nebyl aktivní plán.';
                } catch (error) {
                    status.textContent = 'Chyba: ' + error.message;
                }
            }

            function collectPrivateInputValues() {
                const controls = [
                    document.getElementById('privateRevenue'),
                    document.getElementById('privateEbitdaMargin'),
                    document.getElementById('privateDebt'),
                    document.getElementById('privateLastRound')
                ];
                if (controls.every(control => control.value.trim() === '')) return null;
                if (!controls.every(control => control.reportValidity())) {
                    throw new Error('Doplňte všechny čtyři údaje soukromé společnosti.');
                }
                return {
                    revenue: controls[0].valueAsNumber,
                    ebitda_margin: controls[1].valueAsNumber / 100,
                    debt: controls[2].valueAsNumber,
                    last_round_valuation: controls[3].valueAsNumber
                };
            }

            async function askAdvisor(event) {
                event.preventDefault();
                const question = document.getElementById('advisorQuestion');
                appendAdvisorMessage(question.value, true);
                const submittedQuestion = question.value;
                question.value = '';
                try {
                    const response = await fetch(`/api/advisor/${encodeURIComponent(document.getElementById('tickerInput').value.trim().toUpperCase())}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            question: submittedQuestion,
                            scenario: {
                                leverage: document.getElementById('leverageInput').valueAsNumber / 100,
                                growth: document.getElementById('growthInput').valueAsNumber / 100,
                                exit_mult: document.getElementById('exitMultipleInput').valueAsNumber
                            },
                            private_inputs: collectPrivateInputValues()
                        })
                    });
                    const result = await response.json();
                    if (!response.ok) throw new Error(result.detail || 'Asistent nedokázal odpovědět.');
                    appendAdvisorMessage(result.answer);
                } catch (error) {
                    appendAdvisorMessage('Chyba: ' + error.message);
                }
            }

            function downloadExcel() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                try {
                    window.location.href = `/api/export/excel/${encodeURIComponent(ticker)}?${getFullScenarioQuery()}`;
                } catch (error) {
                    window.alert(error.message);
                }
            }

            function downloadPDF() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                try {
                    window.location.href = `/api/export/pdf/${encodeURIComponent(ticker)}?${getFullScenarioQuery()}`;
                } catch (error) {
                    window.alert(error.message);
                }
            }

            document.addEventListener('keydown', event => {
                if (event.key === 'Escape' && document.getElementById('newsletterDrawer').classList.contains('open')) {
                    toggleNewsletterDrawer(false);
                }
            });

            window.onload = loadValuation;
        </script>
    </body>
    </html>
    """

def run_valuation_logic(
    ticker: str,
    leverage: float = 0.30,
    growth: float = 0.05,
    exit_mult: float = 13.5,
    private_inputs=None
):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info or {}
        financials = stock.financials
    except Exception:
        stock = None
        info = {}
        financials = None

    current_price = info.get("currentPrice")
    shares_outstanding = info.get("sharesOutstanding")
    celkovy_dluh = info.get("totalDebt") or 0
    hotovost = info.get("totalCash") or 0
    ebitda = info.get("ebitda") or 0
    trzby = info.get("totalRevenue") or 0
    cisty_zisk = info.get("netIncomeToCommon") or 0
    market_cap = info.get("marketCap") or 0
    public_data_available = (
        info.get("ebitda") is not None
        and info.get("totalRevenue") is not None
        and bool(shares_outstanding or (current_price and market_cap))
    )

    if not public_data_available:
        required_inputs = ("revenue", "ebitda_margin", "debt", "last_round_valuation")
        if not private_inputs or any(private_inputs.get(key) is None for key in required_inputs):
            return {
                "spolecnost": ticker.upper(),
                "private_mode_required": True,
                "detail": "Veřejná tržní data nejsou dostupná. Zadejte odhadované finanční údaje soukromé společnosti."
            }
        trzby = private_inputs["revenue"]
        ebitda = trzby * private_inputs["ebitda_margin"]
        celkovy_dluh = private_inputs["debt"]
        hotovost = 0
        market_cap = private_inputs["last_round_valuation"]
        current_price = None
        shares_outstanding = None
        cisty_zisk = max(0.0, ebitda * 0.75)
        info = {"beta": 1.0, "sector": "Ostatní"}
        financials = None
    else:
        if stock is not None:
            for history_period, history_interval in (("1d", "1m"), ("5d", "1d")):
                try:
                    recent_history = stock.history(period=history_period, interval=history_interval)
                    if not recent_history.empty:
                        current_price = float(recent_history["Close"].dropna().iloc[-1])
                        break
                except Exception:
                    continue

    trzni_kapitalizace = (current_price * shares_outstanding) if (current_price and shares_outstanding) else market_cap
    celkova_hodnota_podniku = trzni_kapitalizace + celkovy_dluh - hotovost

    try:
        pretax_income = financials.loc["Pretax Income"].iloc[0] if "Pretax Income" in financials.index else cisty_zisk * 1.25
        tax_provision = financials.loc["Tax Provision"].iloc[0] if "Tax Provision" in financials.index else cisty_zisk * 0.25
        efektivni_danova_sazba = max(0.0, min(0.40, tax_provision / pretax_income)) if pretax_income > 0 else 0.21
        
        raw_interest_rate = abs(financials.loc["Interest Expense"].iloc[0]) / celkovy_dluh if "Interest Expense" in financials.index and celkovy_dluh > 0 else 0.05
        implikovana_sazba_dluhu = max(0.03, min(0.07, raw_interest_rate))
    except Exception:
        efektivni_danova_sazba = 0.21
        implikovana_sazba_dluhu = 0.045

    vstupni_cizi_kapital = celkova_hodnota_podniku * leverage
    vstupni_vlastni_kapital = celkova_hodnota_podniku - vstupni_cizi_kapital

    def project_lbo_years(growth_rate: float):
        current_debt = vstupni_cizi_kapital
        projected_ebitda = ebitda
        projected_years = []

        for year in range(1, 6):
            projected_ebitda *= (1 + growth_rate)
            odpisy = trzby * 0.025
            provozni_zisk_ebit = projected_ebitda - odpisy
            urokove_naklady = current_debt * implikovana_sazba_dluhu
            zisk_pred_danemi = max(0.0, provozni_zisk_ebit - urokove_naklady)
            dan_z_prijmu = zisk_pred_danemi * efektivni_danova_sazba
            investicni_vydaje = trzby * 0.035
            zmena_pracovniho_kapitalu = trzby * 0.01
            cisty_penezni_tok = provozni_zisk_ebit - dan_z_prijmu + odpisy - investicni_vydaje - zmena_pracovniho_kapitalu
            volny_tok_pro_splaceni = max(0.0, cisty_penezni_tok - urokove_naklady)
            dcf_dan_z_prijmu = max(0.0, provozni_zisk_ebit) * efektivni_danova_sazba
            dcf_volny_penezni_tok = provozni_zisk_ebit - dcf_dan_z_prijmu + odpisy - investicni_vydaje - zmena_pracovniho_kapitalu
            splatka_jistiny = min(current_debt, volny_tok_pro_splaceni * 0.80)
            current_debt -= splatka_jistiny

            projected_years.append({
                "rok": year,
                "projektovany_provozni_zisk_pred_odpisy": round(projected_ebitda, 2),
                "dcf_volny_penezni_tok": round(dcf_volny_penezni_tok, 2),
                "volny_tok_pro_splaceni": round(volny_tok_pro_splaceni, 2),
                "zbyvajici_cizi_kapital": round(current_debt, 2),
                "mor_jistiny": round(splatka_jistiny, 2)
            })

        return projected_years

    projektovane_roky = project_lbo_years(growth)
    bezkusova_ebitda = projektovane_roky[-1]["projektovany_provozni_zisk_pred_odpisy"]
    aktualni_dluh = projektovane_roky[-1]["zbyvajici_cizi_kapital"]

    hodnota_prvniho = (bezkusova_ebitda * 0.80) * exit_mult
    hodnota_druheho = (bezkusova_ebitda * 0.20) * 9.0
    celkova_hodnota_podniku_sotp = hodnota_prvniho + hodnota_druheho
    hodnota_vlastniho_kapitalu_sotp = celkova_hodnota_podniku_sotp - celkovy_dluh + hotovost

    vystupni_nasobek = exit_mult
    vystupni_hodnota_podniku = bezkusova_ebitda * vystupni_nasobek
    vystupni_hodnota_vlastniho_kapitalu = vystupni_hodnota_podniku - aktualni_dluh

    moic = (vystupni_hodnota_vlastniho_kapitalu / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0
    irr = ((moic ** (1 / 5)) - 1) * 100 if moic > 0 else 0

    try:
        beta = float(info.get("beta") or 1.0)
    except (TypeError, ValueError):
        beta = 1.0
    cost_of_equity = 0.04 + beta * 0.055
    capital_base = max(0.0, vstupni_cizi_kapital) + max(0.0, vstupni_vlastni_kapital)
    debt_weight = max(0.0, vstupni_cizi_kapital) / capital_base if capital_base > 0 else 0.30
    equity_weight = max(0.0, vstupni_vlastni_kapital) / capital_base if capital_base > 0 else 0.70
    wacc = debt_weight * implikovana_sazba_dluhu * (1 - efektivni_danova_sazba) + equity_weight * cost_of_equity
    wacc = max(0.06, min(0.20, wacc))

    dcf_cash_flows = []
    for year in projektovane_roky:
        discount_factor = 1 / ((1 + wacc) ** year["rok"])
        free_cash_flow = year["dcf_volny_penezni_tok"]
        dcf_cash_flows.append({
            "rok": year["rok"],
            "volny_penezni_tok": free_cash_flow,
            "diskontni_faktor": round(discount_factor, 6),
            "soucasna_hodnota_fcf": round(free_cash_flow * discount_factor, 2)
        })

    terminal_value = bezkusova_ebitda * exit_mult
    present_value_terminal_value = terminal_value * dcf_cash_flows[-1]["diskontni_faktor"]
    dcf_enterprise_value = sum(year["soucasna_hodnota_fcf"] for year in dcf_cash_flows) + present_value_terminal_value
    dcf_equity_value = dcf_enterprise_value - celkovy_dluh + hotovost
    shares = shares_outstanding or 0
    market_price = current_price or info.get("regularMarketPrice") or 0
    intrinsic_value_per_share = dcf_equity_value / shares if shares > 0 else None
    upside_downside_percent = ((intrinsic_value_per_share / market_price) - 1) * 100 if intrinsic_value_per_share is not None and market_price > 0 else None

    sector = str(info.get("sector") or "Ostatní")
    peer_multiples = SECTOR_PEER_BENCHMARKS.get(sector, DEFAULT_PEER_BENCHMARKS)
    forward_factor = (1 + growth) ** 5
    peer_metrics = [
        ("EV/EBITDA", peer_multiples["ev_ebitda"], bezkusova_ebitda, True),
        ("P/E", peer_multiples["pe"], cisty_zisk * forward_factor, False),
        ("EV/Revenue", peer_multiples["ev_revenue"], trzby * forward_factor, True)
    ]
    peer_comps_rows = []
    for method, multiple, forward_metric, is_enterprise_multiple in peer_metrics:
        implied_enterprise_value = forward_metric * multiple if is_enterprise_multiple and forward_metric is not None else None
        implied_equity_value = (
            implied_enterprise_value - celkovy_dluh + hotovost
            if is_enterprise_multiple and implied_enterprise_value is not None
            else forward_metric * multiple if forward_metric is not None else None
        )
        implied_share_price = implied_equity_value / shares if implied_equity_value is not None and shares > 0 else None
        peer_comps_rows.append({
            "method": method,
            "multiple": multiple,
            "forward_metric": round(forward_metric, 2) if forward_metric is not None else None,
            "implied_enterprise_value": round(implied_enterprise_value, 2) if implied_enterprise_value is not None else None,
            "implied_equity_value": round(implied_equity_value, 2) if implied_equity_value is not None else None,
            "implied_share_price": round(implied_share_price, 2) if implied_share_price is not None else None
        })

    citlivostni_matice_irr = []
    testovane_rustove_sazby = [max(-0.99, growth + offset) for offset in [-0.02, -0.01, 0.0, 0.01, 0.02]]
    testovane_nasobky = [max(0.1, exit_mult + offset) for offset in [-2.5, -1.5, 0.0, 1.5, 3.0]]

    for g in testovane_rustove_sazby:
        radek_matice = {"tempo_rustu_ebitda_procent": round(g * 100, 1), "vysledky_irr_dle_nasobku": []}
        sensitivity_years = project_lbo_years(g)
        ebitda_t5_g = sensitivity_years[-1]["projektovany_provozni_zisk_pred_odpisy"]
        final_debt_sim = sensitivity_years[-1]["zbyvajici_cizi_kapital"]
        for mult in testovane_nasobky:
            ev_t5 = ebitda_t5_g * mult
            equity_t5 = ev_t5 - final_debt_sim
            sim_moic = (equity_t5 / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0
            sim_irr = ((sim_moic ** (1 / 5)) - 1) * 100 if sim_moic > 0 else -100
            
            radek_matice["vysledky_irr_dle_nasobku"].append({
                "vystupni_nasobek": mult,
                "irr_procent": round(sim_irr, 2)
            })
        citlivostni_matice_irr.append(radek_matice)

    return {
        "spolecnost": ticker.upper(),
        "trzni_kapitalizace": trzni_kapitalizace,
        "celkovy_dluh": celkovy_dluh,
        "hotovost": hotovost,
        "celkova_hodnota_podniku": celkova_hodnota_podniku,
        "metoda_souctu_casti": {
            "vypoctena_hodnota_podniku": celkova_hodnota_podniku_sotp,
            "vypoctena_hodnota_vlastniho_kapitalu": hodnota_vlastniho_kapitalu_sotp
        },
        "model_odkupu_za_cizi_kapital_base_case": {
            "vstupni_vlastni_kapital": round(vstupni_vlastni_kapital, 2),
            "vstupni_cizi_kapital": round(vstupni_cizi_kapital, 2),
            "vystupni_hodnota_vlastniho_kapitalu": round(vystupni_hodnota_vlastniho_kapitalu, 2),
            "moic": round(moic, 2),
            "irr_procent": round(irr, 2)
        },
        "dcf_valuation": {
            "wacc": round(wacc, 6),
            "terminal_value": round(terminal_value, 2),
            "present_value_terminal_value": round(present_value_terminal_value, 2),
            "enterprise_value": round(dcf_enterprise_value, 2),
            "equity_value": round(dcf_equity_value, 2),
            "intrinsic_value_per_share": round(intrinsic_value_per_share, 2) if intrinsic_value_per_share is not None else None,
            "current_share_price": market_price,
            "upside_downside_percent": round(upside_downside_percent, 2) if upside_downside_percent is not None else None,
            "terminal_value_method": "Exit EV/EBITDA",
            "cash_flows": dcf_cash_flows
        },
        "peer_comps": {
            "sector": sector,
            "benchmark_name": "Orientační sektorový benchmark",
            "multiples": peer_comps_rows
        },
        "parametry_scenare": {
            "leverage": leverage,
            "growth": growth,
            "exit_mult": exit_mult,
            "private_inputs": private_inputs
        },
        "projektovane_roky": projektovane_roky,
        "citlivostni_matice_lbo_irr": citlivostni_matice_irr
    }

def build_private_inputs(revenue, ebitda_margin, debt, last_round_valuation):
    values = (revenue, ebitda_margin, debt, last_round_valuation)
    if all(value is None for value in values):
        return None
    return {
        "revenue": revenue,
        "ebitda_margin": ebitda_margin,
        "debt": debt,
        "last_round_valuation": last_round_valuation
    }


@app.get("/api/valuation/{ticker}")
def get_institutional_terminal_valuation(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0),
    private_revenue: float = Query(None, ge=0.0),
    private_ebitda_margin: float = Query(None, ge=-1.0, le=1.0),
    private_debt: float = Query(None, ge=0.0),
    private_last_round_valuation: float = Query(None, gt=0.0)
):
    try:
        private_inputs = build_private_inputs(private_revenue, private_ebitda_margin, private_debt, private_last_round_valuation)
        data = run_valuation_logic(ticker, leverage, growth, exit_mult, private_inputs)
        if not data.get("private_mode_required"):
            data["market_news"] = fetch_market_news(ticker)
        return data
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


def build_advisor_answer(question, data, news):
    dcf = data["dcf_valuation"]
    lbo = data["model_odkupu_za_cizi_kapital_base_case"]
    context = {
        "ticker": data["spolecnost"],
        "question": question,
        "scenario": {key: value for key, value in data["parametry_scenare"].items() if key != "private_inputs"},
        "dcf": dcf,
        "lbo": lbo,
        "peer_comps": data["peer_comps"],
        "headlines": news
    }
    ai_answer = request_ai_completion(
        "Jsi opatrný finanční analytický asistent. Odpovídej česky a stručně pouze z poskytnutých dat. Nezaměňuj modelové odhady za jistotu, upozorni na klíčové předpoklady a neposkytuj osobní investiční doporučení.",
        json.dumps(context, ensure_ascii=False, default=str)
    )
    if ai_answer:
        return ai_answer

    normalized_question = question.lower()
    if "wacc" in normalized_question or "dcf" in normalized_question:
        value_text = (
            f"DCF ukazuje {dcf['intrinsic_value_per_share']:.2f} USD na akcii při WACC {dcf['wacc']:.2%}. "
            if dcf["intrinsic_value_per_share"] is not None
            else f"DCF WACC je {dcf['wacc']:.2%}, ale hodnota na akcii není dostupná bez počtu akcií. "
        )
        return value_text + "Výsledek je citlivý zejména na diskontní sazbu a terminální násobek."
    if "lbo" in normalized_question or "dluh" in normalized_question or "irr" in normalized_question:
        return f"Base-case LBO: MOIC {lbo['moic']:.2f}x a pětiletá IRR {lbo['irr_procent']:.2f}%. Zůstatek dluhu po pěti letech je {data['projektovane_roky'][-1]['zbyvajici_cizi_kapital']:,.0f} USD; výsledek závisí na růstu EBITDA a schopnosti splácet."
    if "peer" in normalized_question or "konkur" in normalized_question or "násob" in normalized_question:
        multiples = ", ".join(f"{peer['method']} {peer['multiple']:.1f}x" for peer in data["peer_comps"]["multiples"])
        return f"Použité orientační sektorové benchmarky pro {data['peer_comps']['sector']}: {multiples}. Nejde o živý vzorek konkrétních konkurentů."
    return f"Pro {data['spolecnost']} model uvádí WACC {dcf['wacc']:.2%}, LBO IRR {lbo['irr_procent']:.2f}% a {len(news)} aktuálních titulků. Zeptejte se konkrétně na DCF, WACC, LBO, dluh nebo sektorové násobky."


@app.get("/api/newsletter/{ticker}")
async def get_investor_newsletter(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0),
    private_revenue: float = Query(None, ge=0.0),
    private_ebitda_margin: float = Query(None, ge=-1.0, le=1.0),
    private_debt: float = Query(None, ge=0.0),
    private_last_round_valuation: float = Query(None, gt=0.0)
):
    private_inputs = build_private_inputs(private_revenue, private_ebitda_margin, private_debt, private_last_round_valuation)
    data = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult, private_inputs)
    if data.get("private_mode_required"):
        raise HTTPException(status_code=422, detail=data["detail"])
    news = await asyncio.to_thread(fetch_market_news, ticker)
    macro = await asyncio.to_thread(fetch_macro_news)
    combined_news = news + macro["global"] + macro["czech"]
    newsletter = await asyncio.to_thread(build_investor_newsletter, data, combined_news)
    return {"newsletter": newsletter, "used_llm": bool(os.getenv("OPENAI_API_KEY")), "articles": len(combined_news), "macro_news": macro}


@app.get("/api/macro-news")
async def get_macro_news():
    return await asyncio.to_thread(fetch_macro_news)


@app.post("/api/advisor/{ticker}")
async def ask_terminal_advisor(ticker: str, request: Request):
    try:
        payload = await request.json()
        question = str(payload.get("question", "")).strip()[:1000]
        scenario = payload.get("scenario") or {}
        leverage = float(scenario.get("leverage", 0.30))
        growth = float(scenario.get("growth", 0.05))
        exit_mult = float(scenario.get("exit_mult", 13.5))
        private_inputs = payload.get("private_inputs")
        if not question:
            raise HTTPException(status_code=422, detail="Zadejte otázku.")
        if not 0 <= leverage <= 1 or not -1 < growth <= 5 or not 0 < exit_mult <= 100:
            raise HTTPException(status_code=422, detail="Parametry scénáře jsou mimo povolený rozsah.")
        data = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult, private_inputs)
        if data.get("private_mode_required"):
            raise HTTPException(status_code=422, detail=data["detail"])
        news = await asyncio.to_thread(fetch_market_news, ticker)
        answer = await asyncio.to_thread(build_advisor_answer, question, data, news)
        return {"answer": answer, "used_llm": bool(os.getenv("OPENAI_API_KEY"))}
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=400, detail=f"Požadavek asistenta nelze zpracovat: {error}")


@app.post("/api/dispatch/subscribe")
async def subscribe_investor_dispatch(request: Request):
    if get_smtp_config() is None:
        raise HTTPException(status_code=503, detail="E-mailový dispatch není nakonfigurován. Nastavte SMTP_HOST, SMTP_USER a SMTP_PASSWORD.")
    try:
        payload = await request.json()
        email = str(payload.get("email", "")).strip()
        if parseaddr(email)[1] != email or "@" not in email:
            raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")
        ticker = str(payload.get("ticker", "")).strip().upper()
        if not ticker or len(ticker) > 40 or any(not (char.isalnum() or char in ".-^= ") for char in ticker):
            raise HTTPException(status_code=422, detail="Zadejte platný ticker nebo krátký název společnosti.")
        scheduled = next_dispatch_time()
        scenario = payload.get("scenario") or {}
        leverage = float(scenario.get("leverage", 0.30))
        growth = float(scenario.get("growth", 0.05))
        exit_mult = float(scenario.get("exit_mult", 13.5))
        if not 0 <= leverage <= 1 or not -1 < growth <= 5 or not 0 < exit_mult <= 100:
            raise HTTPException(status_code=422, detail="Parametry scénáře jsou mimo povolený rozsah.")
        private_inputs = payload.get("private_inputs")
        if private_inputs:
            required = ("revenue", "ebitda_margin", "debt", "last_round_valuation")
            if any(key not in private_inputs or private_inputs[key] is None for key in required):
                raise HTTPException(status_code=422, detail="Soukromý scénář vyžaduje všechny čtyři finanční údaje.")
        valuation_check = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult, private_inputs)
        if valuation_check.get("private_mode_required"):
            raise HTTPException(status_code=422, detail=valuation_check["detail"])
        key = email.casefold()
        if key not in dispatch_subscriptions and len(dispatch_subscriptions) >= 20:
            raise HTTPException(status_code=429, detail="Byl dosažen limit aktivních odběrů v této instanci.")
        dispatch_subscriptions[key] = {
            "email": email,
            "ticker": ticker,
            "next_send": scheduled,
            "leverage": leverage,
            "growth": growth,
            "exit_mult": exit_mult,
            "private_inputs": private_inputs,
            "last_error": None
        }
        return {"status": "scheduled", "next_send": scheduled.isoformat(timespec="minutes")}
    except HTTPException:
        raise
    except (TypeError, ValueError) as error:
        raise HTTPException(status_code=422, detail=f"Neplatné údaje plánu: {error}")


@app.post("/api/dispatch/send-now")
async def send_investor_bulletin_now(request: Request):
    if get_smtp_config() is None:
        raise HTTPException(status_code=503, detail="E-mailový dispatch není nakonfigurován. Nastavte SMTP_HOST, SMTP_USER a SMTP_PASSWORD.")
    try:
        payload = await request.json()
        email = str(payload.get("email", "")).strip()
        if parseaddr(email)[1] != email or "@" not in email:
            raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")
        ticker = str(payload.get("ticker", "")).strip().upper()
        if not ticker or len(ticker) > 40 or any(not (char.isalnum() or char in ".-^= ") for char in ticker):
            raise HTTPException(status_code=422, detail="Zadejte platný ticker nebo krátký název společnosti.")
        scenario = payload.get("scenario") or {}
        leverage = float(scenario.get("leverage", 0.30))
        growth = float(scenario.get("growth", 0.05))
        exit_mult = float(scenario.get("exit_mult", 13.5))
        if not 0 <= leverage <= 1 or not -1 < growth <= 5 or not 0 < exit_mult <= 100:
            raise HTTPException(status_code=422, detail="Parametry scénáře jsou mimo povolený rozsah.")
        private_inputs = payload.get("private_inputs")
        if private_inputs:
            required = ("revenue", "ebitda_margin", "debt", "last_round_valuation")
            if any(key not in private_inputs or private_inputs[key] is None for key in required):
                raise HTTPException(status_code=422, detail="Soukromý scénář vyžaduje všechny čtyři finanční údaje.")
        result = await send_bulletin_email(email, ticker, leverage, growth, exit_mult, private_inputs)
        return {"status": "sent", **result}
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=502, detail=f"Bulletin se nepodařilo odeslat: {error}")


@app.post("/api/dispatch/unsubscribe")
async def unsubscribe_investor_dispatch(request: Request):
    try:
        payload = await request.json()
        email = str(payload.get("email", "")).strip()
        if parseaddr(email)[1] != email or "@" not in email:
            raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")
        removed = dispatch_subscriptions.pop(email.casefold(), None) is not None
        return {"removed": removed}
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=400, detail=f"Požadavek na zrušení odběru nelze zpracovat: {error}")

@app.get("/api/export/excel/{ticker}")
def export_valuation_excel(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0),
    private_revenue: float = Query(None, ge=0.0),
    private_ebitda_margin: float = Query(None, ge=-1.0, le=1.0),
    private_debt: float = Query(None, ge=0.0),
    private_last_round_valuation: float = Query(None, gt=0.0)
):
    try:
        private_inputs = build_private_inputs(private_revenue, private_ebitda_margin, private_debt, private_last_round_valuation)
        data = run_valuation_logic(ticker, leverage, growth, exit_mult, private_inputs)
        if data.get("private_mode_required"):
            raise HTTPException(status_code=422, detail=data["detail"])
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Valuation - {data['spolecnost']}"
        ws.views.sheetView[0].showGridLines = True
        
        font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
        font_section = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_header = Font(name="Segoe UI", size=10, bold=True, color="000000")
        font_bold = Font(name="Segoe UI", size=10, bold=True)
        font_normal = Font(name="Segoe UI", size=10)
        
        fill_section = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        fill_header = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        
        border_thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        ws.append(["TENSEGRITY TERMINAL — INSTITUTIONAL VALUATION REPORT"])
        ws.cell(row=1, column=1).font = font_title
        ws.append([f"Společnost: {data['spolecnost']} | Datum: 2026-11-04"])
        ws.append([f"Parametry scénáře: dluh {leverage:.1%} | růst EBITDA {growth:.1%} | výstupní násobek {exit_mult:.1f}x"])
        ws.append([])

        def add_section_header(title):
            ws.append([title])
            cell = ws.cell(row=ws.max_row, column=1)
            cell.font = font_section
            cell.fill = fill_section
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=2)

        add_section_header("ZÁKLADNÍ TRŽNÍ DATA & METODA SOUČTU ČÁSTÍ")
        ws.append(["Ukazatel", "Hodnota (USD)"])
        for col in range(1, 3):
            ws.cell(row=ws.max_row, column=col).font = font_header
            ws.cell(row=ws.max_row, column=col).fill = fill_header

        rows_sotp = [
            ("Tržní kapitalizace", data["trzni_kapitalizace"]),
            ("Celkový dluh", data["celkovy_dluh"]),
            ("Hotovost", data["hotovost"]),
            ("Celková hodnota podniku", data["celkova_hodnota_podniku"]),
            ("Součet částí - Hodnota podniku", data["metoda_souctu_casti"]["vypoctena_hodnota_podniku"]),
            ("Součet částí - Hodnota vlastní kapitálu", data["metoda_souctu_casti"]["vypoctena_hodnota_vlastniho_kapitalu"])
        ]
        for item, val in rows_sotp:
            ws.append([item, val])
            r = ws.max_row
            ws.cell(row=r, column=1).font = font_normal
            ws.cell(row=r, column=2).font = font_bold
            ws.cell(row=r, column=2).number_format = '#,##0'
            ws.cell(row=r, column=1).border = border_thin
            ws.cell(row=r, column=2).border = border_thin

        ws.append([])

        add_section_header("MODEL ODKUPU ZA CIZÍ KAPITÁL")
        ws.append(["Parametr transakce", "Hodnota"])
        for col in range(1, 3):
            ws.cell(row=ws.max_row, column=col).font = font_header
            ws.cell(row=ws.max_row, column=col).fill = fill_header

        rows_lbo = [
            ("Vstupní vlastní kapitál sponzora", data["model_odkupu_za_cizi_kapital_base_case"]["vstupni_vlastni_kapital"]),
            ("Vstupní cizí kapitál", data["model_odkupu_za_cizi_kapital_base_case"]["vstupni_cizi_kapital"]),
            ("Výstupní hodnota vlastní kapitálu", data["model_odkupu_za_cizi_kapital_base_case"]["vystupni_hodnota_vlastniho_kapitalu"]),
            ("Násobek investovaného kapitálu", f"{data['model_odkupu_za_cizi_kapital_base_case']['moic']}x"),
            ("Vnitřní míra výnosnosti", f"{data['model_odkupu_za_cizi_kapital_base_case']['irr_procent']}%")
        ]
        for item, val in rows_lbo:
            ws.append([item, val])
            r = ws.max_row
            ws.cell(row=r, column=1).font = font_normal
            ws.cell(row=r, column=2).font = font_bold
            if isinstance(val, (int, float)):
                ws.cell(row=r, column=2).number_format = '#,##0.00'
            ws.cell(row=r, column=1).border = border_thin
            ws.cell(row=r, column=2).border = border_thin

        ws.append([])

        ws.append(["5-LETÝ LBO AMORTIZAČNÍ KALENDÁŘ A CASH FLOW"])
        section_row = ws.max_row
        ws.cell(row=section_row, column=1).font = font_section
        ws.cell(row=section_row, column=1).fill = fill_section
        ws.merge_cells(start_row=section_row, start_column=1, end_row=section_row, end_column=5)
        ws.append(["Rok", "Projektovaná EBITDA", "Volné Cash Flow (FCF)", "Splátka jistiny", "Zůstatek dluhu"])
        for col in range(1, 6):
            cell = ws.cell(row=ws.max_row, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = border_thin
        for year in data["projektovane_roky"]:
            ws.append([
                f"Rok {year['rok']}",
                year["projektovany_provozni_zisk_pred_odpisy"],
                year["volny_tok_pro_splaceni"],
                year["mor_jistiny"],
                year["zbyvajici_cizi_kapital"]
            ])
            row_number = ws.max_row
            for col in range(1, 6):
                cell = ws.cell(row=row_number, column=col)
                cell.font = font_normal
                cell.border = border_thin
                if col > 1:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right")

        ws.append([])

        add_section_header("CITLIVOSTNÍ MATICE VNITŘNÍ MÍRY VÝNOSNOSTI (%)")
        headers = ["Růst / Multiplikátor"] + [f"{m}x" for m in [11.0, 12.0, 13.5, 15.0, 16.5]]
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row, column=col).font = font_header
            ws.cell(row=ws.max_row, column=col).fill = fill_header

        for row in data["citlivostni_matice_lbo_irr"]:
            row_data = [f"{row['tempo_rustu_ebitda_procent']}%"]
            for cell in row["vysledky_irr_dle_nasobku"]:
                row_data.append(f"{cell['irr_procent']}%")
            ws.append(row_data)
            r = ws.max_row
            for col in range(1, len(row_data) + 1):
                ws.cell(row=r, column=col).font = font_normal
                ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=col).border = border_thin

        ws.column_dimensions['A'].width = 38
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15

        dcf_ws = wb.create_sheet("DCF Analysis")
        dcf_ws.append(["DISKONTOVANÉ CASH FLOW (DCF)"])
        dcf_ws.cell(row=1, column=1).font = font_title
        dcf_ws.append([f"WACC: {data['dcf_valuation']['wacc']:.2%} | Terminální metoda: {data['dcf_valuation']['terminal_value_method']} | Exit násobek: {exit_mult:.1f}x"])
        dcf_ws.append(["Rok", "Unlevered FCF (USD)", "Diskontní faktor", "Současná hodnota FCF (USD)"])
        for col in range(1, 5):
            cell = dcf_ws.cell(row=dcf_ws.max_row, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = border_thin
        for year in data["dcf_valuation"]["cash_flows"]:
            dcf_ws.append([
                f"Rok {year['rok']}",
                year["volny_penezni_tok"],
                year["diskontni_faktor"],
                year["soucasna_hodnota_fcf"]
            ])
            row_number = dcf_ws.max_row
            for col in range(1, 5):
                cell = dcf_ws.cell(row=row_number, column=col)
                cell.font = font_normal
                cell.border = border_thin
                if col in (2, 4):
                    cell.number_format = '#,##0.00'
                elif col == 3:
                    cell.number_format = '0.000000'
        dcf_summary = [
            ("Terminální hodnota (USD)", data["dcf_valuation"]["terminal_value"]),
            ("Současná hodnota terminální hodnoty (USD)", data["dcf_valuation"]["present_value_terminal_value"]),
            ("DCF hodnota podniku (USD)", data["dcf_valuation"]["enterprise_value"]),
            ("DCF hodnota vlastního kapitálu (USD)", data["dcf_valuation"]["equity_value"]),
            ("Vnitřní hodnota na akcii (USD)", data["dcf_valuation"]["intrinsic_value_per_share"]),
            ("Aktuální tržní cena (USD)", data["dcf_valuation"]["current_share_price"]),
            ("Odchylka vůči trhu (%)", data["dcf_valuation"]["upside_downside_percent"])
        ]
        dcf_ws.append([])
        for label, value in dcf_summary:
            dcf_ws.append([label, value if value is not None else "N/A"])
            row_number = dcf_ws.max_row
            dcf_ws.cell(row=row_number, column=1).font = font_normal
            dcf_ws.cell(row=row_number, column=2).font = font_bold
            dcf_ws.cell(row=row_number, column=1).border = border_thin
            dcf_ws.cell(row=row_number, column=2).border = border_thin
            if isinstance(value, (int, float)):
                dcf_ws.cell(row=row_number, column=2).number_format = '0.00"%"' if "Odchylka" in label else '#,##0.00'
        dcf_ws.column_dimensions['A'].width = 42
        dcf_ws.column_dimensions['B'].width = 25
        dcf_ws.column_dimensions['C'].width = 20
        dcf_ws.column_dimensions['D'].width = 32

        peer_ws = wb.create_sheet("Peer Comps")
        peer_ws.append(["SROVNÁNÍ KONKURENCE (TRADING COMPS)"])
        peer_ws.cell(row=1, column=1).font = font_title
        peer_ws.append([f"{data['peer_comps']['benchmark_name']}: {data['peer_comps']['sector']} | Forward metriky: rok 5"])
        peer_ws.append(["Metoda", "Benchmark násobek", "Metrika (rok 5)", "Implikovaná EV / Equity (USD)", "Implikovaná cena/akcii (USD)"])
        for col in range(1, 6):
            cell = peer_ws.cell(row=peer_ws.max_row, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = border_thin
        for peer in data["peer_comps"]["multiples"]:
            implied_value = peer["implied_enterprise_value"] if peer["implied_enterprise_value"] is not None else peer["implied_equity_value"]
            peer_ws.append([
                peer["method"],
                peer["multiple"],
                peer["forward_metric"] if peer["forward_metric"] is not None else "N/A",
                implied_value if implied_value is not None else "N/A",
                peer["implied_share_price"] if peer["implied_share_price"] is not None else "N/A"
            ])
            row_number = peer_ws.max_row
            for col in range(1, 6):
                cell = peer_ws.cell(row=row_number, column=col)
                cell.font = font_normal
                cell.border = border_thin
                if col == 2 and isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0"x"'
                elif col > 2 and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
        peer_ws.column_dimensions['A'].width = 20
        peer_ws.column_dimensions['B'].width = 20
        peer_ws.column_dimensions['C'].width = 25
        peer_ws.column_dimensions['D'].width = 34
        peer_ws.column_dimensions['E'].width = 32

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"Tensegrity_Valuation_{data['spolecnost']}.xlsx"
        return StreamingResponse(
            output, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/export/pdf/{ticker}")
def export_valuation_pdf(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0),
    private_revenue: float = Query(None, ge=0.0),
    private_ebitda_margin: float = Query(None, ge=-1.0, le=1.0),
    private_debt: float = Query(None, ge=0.0),
    private_last_round_valuation: float = Query(None, gt=0.0)
):
    try:
        private_inputs = build_private_inputs(private_revenue, private_ebitda_margin, private_debt, private_last_round_valuation)
        data = run_valuation_logic(ticker, leverage, growth, exit_mult, private_inputs)
        if data.get("private_mode_required"):
            raise HTTPException(status_code=422, detail=data["detail"])
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=35,
            leftMargin=35,
            topMargin=35,
            bottomMargin=45
        )
        elements = []
        
        styles = getSampleStyleSheet()
        
        title_brand_style = ParagraphStyle(
            'TitleBrandStyle',
            parent=styles['Normal'],
            fontName=PDF_FONT_BOLD,
            fontSize=15,
            leading=16,
            textColor=colors.HexColor('#0F172A')
        )
        subtitle_brand_style = ParagraphStyle(
            'SubtitleBrandStyle',
            parent=styles['Normal'],
            fontName=PDF_FONT_REG,
            fontSize=8,
            leading=10,
            textColor=colors.HexColor('#64748B')
        )
        doc_info_style = ParagraphStyle(
            'DocInfoStyle',
            parent=styles['Normal'],
            fontName=PDF_FONT_REG,
            fontSize=9,
            leading=12,
            alignment=2,
            textColor=colors.HexColor('#334155')
        )
        section_style = ParagraphStyle(
            'SectionStyle',
            parent=styles['Heading2'],
            fontName=PDF_FONT_BOLD,
            fontSize=11,
            leading=14,
            textColor=colors.HexColor('#0F172A'),
            spaceBefore=12,
            spaceAfter=6
        )
        
        # VYMEZENÍ UNIKÁTNÍCH STYLŮ PRO TABULKY (ZABRÁNI PŘEPISU BARVY)
        th_left = ParagraphStyle('THLeft', parent=styles['Normal'], fontName=PDF_FONT_BOLD, fontSize=9, leading=11, textColor=colors.white)
        th_right = ParagraphStyle('THRight', parent=styles['Normal'], fontName=PDF_FONT_BOLD, fontSize=9, leading=11, alignment=2, textColor=colors.white)
        
        body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName=PDF_FONT_REG, fontSize=9, leading=12, textColor=colors.HexColor('#0F172A'))
        body_bold = ParagraphStyle('BodyBold', parent=body_style, fontName=PDF_FONT_BOLD)
        body_right = ParagraphStyle('BodyRight', parent=body_style, alignment=2)
        body_right_bold = ParagraphStyle('BodyRightBold', parent=body_bold, alignment=2)

        def format_pdf_number(value, decimals=2):
            return f"{value:,.{decimals}f}" if value is not None else "N/A"
        
        kpi_title_style = ParagraphStyle('KPITitleStyle', parent=styles['Normal'], fontName=PDF_FONT_REG, fontSize=7.5, leading=9, textColor=colors.HexColor('#64748B'), alignment=1)
        kpi_value_style = ParagraphStyle('KPIValueStyle', parent=styles['Normal'], fontName=PDF_FONT_BOLD, fontSize=10, leading=12, textColor=colors.HexColor('#0F172A'), alignment=1)

        # 1. ZÁHLAVÍ REPORTU S PRAVÝM LOGEM
        logo_drawing = create_pdf_logo(45, 35)
        header_brand_table = Table([
            [logo_drawing, [Paragraph("TENSEGRITY TERMINAL", title_brand_style), Paragraph("INSTITUTIONAL VALUATION PLATFORM", subtitle_brand_style)]]
        ], colWidths=[50, 200])
        header_brand_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 0),
        ]))

        meta_info_p = Paragraph(
            f"<b>VALUATION MEMORANDUM</b><br/>"
            f"Target: <b>{data['spolecnost']}</b> | Datum: 2026-11-04<br/>"
            f"Scenario: Debt {leverage:.1%} | EBITDA growth {growth:.1%} | Exit {exit_mult:.1f}x<br/>"
            f"<font color='#0284C7'>Classification: Highly Confidential</font>",
            doc_info_style
        )

        header_master = Table([[header_brand_table, meta_info_p]], colWidths=[270, 255])
        header_master.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 0),
        ]))
        elements.append(header_master)
        elements.append(Spacer(1, 8))
        elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0284C7'), spaceBefore=2, spaceAfter=12))

        # 2. EXECUTIVE SUMMARY CARDS (KPI BLOK)
        kpi_data = [
            [
                Paragraph("CELKOVÁ HODNOTA PODNIKU", kpi_title_style),
                Paragraph("SOTP HODNOTA KAPITÁLU", kpi_title_style),
                Paragraph("LBO BASE CASE IRR", kpi_title_style),
                Paragraph("SPONZOR MOIC", kpi_title_style)
            ],
            [
                Paragraph(f"${data['celkova_hodnota_podniku']/1e9:,.2f} B", kpi_value_style),
                Paragraph(f"${data['metoda_souctu_casti']['vypoctena_hodnota_vlastniho_kapitalu']/1e9:,.2f} B", kpi_value_style),
                Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['irr_procent']}%", kpi_value_style),
                Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['moic']}x", kpi_value_style)
            ]
        ]
        kpi_table = Table(kpi_data, colWidths=[130, 130, 130, 135])
        kpi_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E2E8F0')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 6),
        ]))
        elements.append(kpi_table)
        elements.append(Spacer(1, 10))

        # 3. ZÁKLADNÍ TRŽNÍ DATA A SOTP
        elements.append(Paragraph("1. Základní Tržní Data & Metoda Součtu Částí (SOTP)", section_style))
        sotp_table_data = [
            [Paragraph("<b>Finanční Ukazatel</b>", th_left), Paragraph("<b>Hodnota (USD)</b>", th_right)],
            [Paragraph("Tržní kapitalizace", body_style), Paragraph(f"{data['trzni_kapitalizace']:,.0f}", body_right)],
            [Paragraph("Celkový dluh", body_style), Paragraph(f"{data['celkovy_dluh']:,.0f}", body_right)],
            [Paragraph("Hotovost a ekvivalenty", body_style), Paragraph(f"{data['hotovost']:,.0f}", body_right)],
            [Paragraph("Celková hodnota podniku (Enterprise Value)", body_bold), Paragraph(f"{data['celkova_hodnota_podniku']:,.0f}", body_right_bold)],
            [Paragraph("Součet částí - Ocenění podniku (SOTP EV)", body_style), Paragraph(f"{data['metoda_souctu_casti']['vypoctena_hodnota_podniku']:,.0f}", body_right)],
            [Paragraph("Součet částí - Hodnota vlastní kapitálu (SOTP Equity)", body_bold), Paragraph(f"{data['metoda_souctu_casti']['vypoctena_hodnota_vlastniho_kapitalu']:,.0f}", body_right_bold)]
        ]
        t1 = Table(sotp_table_data, colWidths=[315, 210])
        t1.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(t1)
        elements.append(Spacer(1, 10))

        # 4. MODEL ODKUPU ZA CIZÍ KAPITÁL (LBO)
        elements.append(Paragraph("2. Model Odkupu za Cizí Kapitál (LBO Base Case)", section_style))
        lbo_table_data = [
            [Paragraph("<b>Parametr Transakce LBO</b>", th_left), Paragraph("<b>Hodnota</b>", th_right)],
            [Paragraph("Vstupní vlastní kapitál sponzora", body_style), Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['vstupni_vlastni_kapital']:,.0f} USD", body_right)],
            [Paragraph("Vstupní cizí kapitál (Senior Debt)", body_style), Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['vstupni_cizi_kapital']:,.0f} USD", body_right)],
            [Paragraph("Výstupní hodnota vlastní kapitálu (Exit Equity)", body_style), Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['vystupni_hodnota_vlastniho_kapitalu']:,.0f} USD", body_right)],
            [Paragraph("Násobek investovaného kapitálu (MOIC)", body_bold), Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['moic']}x", body_right_bold)],
            [Paragraph("Vnitřní míra výnosnosti (Base Case IRR)", body_bold), Paragraph(f"{data['model_odkupu_za_cizi_kapital_base_case']['irr_procent']}%", body_right_bold)]
        ]
        t2 = Table(lbo_table_data, colWidths=[315, 210])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(t2)
        elements.append(Spacer(1, 10))

        # 5. DISKONTOVANÉ CASH FLOW
        elements.append(Paragraph("3. Diskontované Cash Flow (DCF Ocenění)", section_style))
        elements.append(Paragraph(
            f"WACC: {data['dcf_valuation']['wacc']:.2%} | Terminální metoda: {data['dcf_valuation']['terminal_value_method']} ({exit_mult:.1f}x)",
            subtitle_brand_style
        ))
        dcf_table_data = [[
            Paragraph("<b>Rok</b>", th_left),
            Paragraph("<b>Unlevered FCF (USD)</b>", th_right),
            Paragraph("<b>Diskontní faktor</b>", th_right),
            Paragraph("<b>Současná hodnota FCF (USD)</b>", th_right)
        ]]
        for year in data["dcf_valuation"]["cash_flows"]:
            dcf_table_data.append([
                Paragraph(f"Rok {year['rok']}", body_style),
                Paragraph(format_pdf_number(year["volny_penezni_tok"]), body_right),
                Paragraph(f"{year['diskontni_faktor']:.4f}", body_right),
                Paragraph(format_pdf_number(year["soucasna_hodnota_fcf"]), body_right)
            ])
        dcf_table = Table(dcf_table_data, colWidths=[50, 145, 110, 220])
        dcf_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(dcf_table)
        dcf_summary_data = [[Paragraph("<b>DCF Ukazatel</b>", th_left), Paragraph("<b>Hodnota</b>", th_right)]]
        for label, value in [
            ("Terminální hodnota", data["dcf_valuation"]["terminal_value"]),
            ("Současná hodnota terminální hodnoty", data["dcf_valuation"]["present_value_terminal_value"]),
            ("DCF hodnota podniku", data["dcf_valuation"]["enterprise_value"]),
            ("DCF hodnota vlastního kapitálu", data["dcf_valuation"]["equity_value"]),
            ("Vnitřní hodnota na akcii", data["dcf_valuation"]["intrinsic_value_per_share"]),
            ("Aktuální tržní cena na akcii", data["dcf_valuation"]["current_share_price"]),
            ("Odchylka vůči tržní ceně", data["dcf_valuation"]["upside_downside_percent"])
        ]:
            value_text = f"{value:.2f}%" if label == "Odchylka vůči tržní ceně" and value is not None else format_pdf_number(value)
            dcf_summary_data.append([Paragraph(label, body_style), Paragraph(value_text, body_right)])
        dcf_summary_table = Table(dcf_summary_data, colWidths=[315, 210])
        dcf_summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(dcf_summary_table)
        elements.append(Spacer(1, 10))

        # 6. TRADING COMPS
        elements.append(Paragraph("4. Srovnání Konkurence (Trading Comps)", section_style))
        elements.append(Paragraph(
            f"{data['peer_comps']['benchmark_name']}: {data['peer_comps']['sector']} | Forward metriky: rok 5",
            subtitle_brand_style
        ))
        peer_table_data = [[
            Paragraph("<b>Metoda</b>", th_left),
            Paragraph("<b>Benchmark</b>", th_right),
            Paragraph("<b>Metrika (rok 5)</b>", th_right),
            Paragraph("<b>Implikovaná EV / Equity</b>", th_right),
            Paragraph("<b>Cena/akcii</b>", th_right)
        ]]
        for peer in data["peer_comps"]["multiples"]:
            implied_value = peer["implied_enterprise_value"] if peer["implied_enterprise_value"] is not None else peer["implied_equity_value"]
            peer_table_data.append([
                Paragraph(peer["method"], body_style),
                Paragraph(f"{peer['multiple']:.1f}x", body_right),
                Paragraph(format_pdf_number(peer["forward_metric"]), body_right),
                Paragraph(format_pdf_number(implied_value), body_right),
                Paragraph(format_pdf_number(peer["implied_share_price"]), body_right)
            ])
        peer_table = Table(peer_table_data, colWidths=[85, 85, 115, 130, 110])
        peer_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(peer_table)
        elements.append(Spacer(1, 10))

        # 7. LBO AMORTIZAČNÍ KALENDÁŘ A CASH FLOW
        elements.append(Paragraph("5. 5-Letý LBO Amortizační Kalendář a Cash Flow", section_style))
        amortization_headers = [
            Paragraph("<b>Rok</b>", th_left),
            Paragraph("<b>Projektovaná EBITDA</b>", th_right),
            Paragraph("<b>Volné Cash Flow (FCF)</b>", th_right),
            Paragraph("<b>Splátka jistiny</b>", th_right),
            Paragraph("<b>Zůstatek dluhu</b>", th_right)
        ]
        amortization_rows = [amortization_headers]
        for year in data["projektovane_roky"]:
            amortization_rows.append([
                Paragraph(f"Rok {year['rok']}", body_style),
                Paragraph(f"{year['projektovany_provozni_zisk_pred_odpisy']:,.0f} USD", body_right),
                Paragraph(f"{year['volny_tok_pro_splaceni']:,.0f} USD", body_right),
                Paragraph(f"{year['mor_jistiny']:,.0f} USD", body_right),
                Paragraph(f"{year['zbyvajici_cizi_kapital']:,.0f} USD", body_right)
            ])

        amortization_table = Table(amortization_rows, colWidths=[55, 110, 120, 110, 130])
        amortization_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(amortization_table)
        elements.append(Spacer(1, 10))

        # 8. CITLIVOSTNÍ MATICE IRR
        elements.append(Paragraph("6. Citlivostní Matice Vnitřní Míry Výnosnosti (IRR %)", section_style))
        matrix_headers = [Paragraph("<b>Tempo růstu \\ Násobek</b>", th_left)] + [Paragraph(f"<b>{m}x</b>", th_right) for m in [11.0, 12.0, 13.5, 15.0, 16.5]]
        matrix_rows = [matrix_headers]
        
        for row in data["citlivostni_matice_lbo_irr"]:
            r_data = [Paragraph(f"<b>{row['tempo_rustu_ebitda_procent']}%</b>", body_style)]
            for cell in row["vysledky_irr_dle_nasobku"]:
                val = cell['irr_procent']
                color_hex = '#16A34A' if val >= 0 else '#DC2626'
                p_style = ParagraphStyle('CellIRR', parent=body_right, textColor=colors.HexColor(color_hex), fontName=PDF_FONT_BOLD)
                r_data.append(Paragraph(f"{val}%", p_style))
            matrix_rows.append(r_data)
            
        t3 = Table(matrix_rows, colWidths=[125, 80, 80, 80, 80, 80])
        t3.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(t3)

        doc.build(elements, canvasmaker=NumberedCanvas)
        buffer.seek(0)
        
        filename = f"Tensegrity_Valuation_{data['spolecnost']}.pdf"
        return StreamingResponse(
            buffer,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))