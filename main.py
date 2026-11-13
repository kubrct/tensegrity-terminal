# -*- coding: utf-8 -*-
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import asyncio
import io
import json
import logging
import os
import re
import yfinance as yf
import openpyxl

from datetime import datetime
from email.utils import parseaddr
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

app = FastAPI(title="Tensegrity Terminal API", version="12.5")
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
logger = logging.getLogger(__name__)

FOREX_PAIRS = {
    "EUR/CZK": "EURCZK=X",
    "USD/CZK": "USDCZK=X",
    "GBP/CZK": "GBPCZK=X"
}
FOREX_FALLBACK_RATES = {"EUR/CZK": 25.0, "USD/CZK": 23.0, "GBP/CZK": 29.0}
LOGO_PATH = os.path.join(STATIC_DIR, "logo.png")


def get_logo_bytes():
    if not os.path.exists(LOGO_PATH):
        return None
    with Image.open(LOGO_PATH) as image:
        image = image.convert("RGBA")
        mark = image.crop((150, 250, 600, 600))
        output = io.BytesIO()
        mark.save(output, format="PNG")
        return output.getvalue()


def get_live_forex_rates():
    rates = {}
    for pair, symbol in FOREX_PAIRS.items():
        rate = FOREX_FALLBACK_RATES[pair]
        live = False
        try:
            history = yf.Ticker(symbol).history(period="5d", interval="1d")
            closes = history["Close"].dropna()
            if not closes.empty:
                rate = float(closes.iloc[-1])
                live = True
        except Exception as error:
            logger.warning("Unable to fetch live forex rate for %s: %s", pair, error)
        rates[pair] = {"rate": rate, "live": live}
    return {
        "rates": rates,
        "updated_at": datetime.now().astimezone().isoformat(timespec="minutes")
    }


SENDER_EMAIL = os.getenv("SENDER_EMAIL", os.getenv("SMTP_USER", "kubenkamatyas@gmail.com"))
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", os.getenv("SMTP_PASSWORD", ""))
SUBSCRIBERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "subscribers.json")


def get_subscribers():
    if not os.path.exists(SUBSCRIBERS_FILE):
        return []
    try:
        with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, list):
            return []
        return [item.strip() for item in data if isinstance(item, str) and item.strip()]
    except (OSError, json.JSONDecodeError):
        return []


def save_subscriber(email: str):
    normalized = email.strip()
    subscribers = get_subscribers()
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


@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    logo_file = os.path.join(STATIC_DIR, "logo.png")
    if os.path.exists(logo_file):
        return FileResponse(logo_file)
    raise HTTPException(status_code=404, detail="Favicon nenalezen.")


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
    created = save_subscriber(candidate)
    return {
        "status": "ok",
        "message": "Odběr investičního bulletinu byl úspěšně uložen." if created else "Tato e-mailová adresa je již k odběru přihlášena.",
        "already_subscribed": not created,
    }


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
    <style>
        :root {
            --bg-main: #060911;
            --bg-card: rgba(15, 22, 37, 0.85);
            --bg-card-hover: rgba(21, 31, 52, 0.95);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --accent-blue: #38bdf8;
            --accent-blue-glow: rgba(56, 189, 248, 0.2);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --success: #10b981;
            --warning: #f59e0b;
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
        .container { max-width: 1400px; margin: 0 auto; }
        header {
            display: flex; justify-content: space-between; align-items: center;
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 16px 28px; margin-bottom: 24px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.5); flex-wrap: wrap; gap: 16px;
        }
        .brand { display: flex; align-items: center; gap: 16px; }
        .brand-title { font-size: 16px; font-weight: 800; letter-spacing: 4px; color: #fff; text-transform: uppercase; }
        .brand-subtitle { font-size: 10px; letter-spacing: 6px; color: var(--text-secondary); text-transform: uppercase; }
        .controls { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
        .forex-strip {
            display: flex; align-items: center; gap: 12px; padding: 8px 14px;
            border: 1px solid var(--border-subtle); border-radius: 8px;
            background: rgba(6, 9, 17, 0.6); color: var(--text-secondary);
            font: 11px 'JetBrains Mono', monospace;
        }
        .forex-strip strong { color: var(--text-primary); }
        .search-wrapper input {
            background: rgba(6, 9, 17, 0.8); border: 1px solid var(--border-subtle);
            color: var(--text-primary); font-family: 'JetBrains Mono', monospace;
            font-size: 14px; font-weight: 600; padding: 10px 16px; border-radius: 10px;
            outline: none; width: 140px; text-transform: uppercase; letter-spacing: 1px;
        }
        .search-wrapper input:focus { border-color: var(--accent-blue); box-shadow: 0 0 12px var(--accent-blue-glow); }
        .btn {
            font-family: 'Plus Jakarta Sans', sans-serif; font-size: 13px; font-weight: 600;
            padding: 10px 18px; border-radius: 10px; border: 1px solid transparent; cursor: pointer;
            display: inline-flex; align-items: center; gap: 8px; transition: all 0.2s ease;
        }
        .btn-primary { background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%); color: #040914; font-weight: 700; }
        .btn-primary:hover { opacity: 0.95; transform: translateY(-1px); }
        .btn-excel { background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.3); color: #34d399; }
        .btn-pdf { background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); color: #f87171; }
        .btn-news { background: rgba(148, 163, 184, 0.12); border: 1px solid var(--border-subtle); color: var(--text-primary); }
        .btn-news:hover { border-color: var(--accent-blue); color: var(--accent-blue); }

        .modal-backdrop {
            display: none; position: fixed; inset: 0; background: rgba(4, 7, 15, 0.82);
            backdrop-filter: blur(6px); z-index: 999; align-items: center; justify-content: center; padding: 20px;
        }
        .modal-backdrop.open { display: flex; }
        .modal-box {
            background: #0f1625; border: 1px solid var(--border-subtle); border-radius: 16px;
            max-width: 720px; width: 100%; padding: 28px; box-shadow: 0 20px 50px rgba(0,0,0,0.7);
            max-height: 85vh; overflow-y: auto;
        }
        .news-item { padding: 14px 0; border-bottom: 1px solid var(--border-subtle); }
        .news-item:last-child { border-bottom: none; }
        .news-item a { color: var(--accent-blue); text-decoration: none; font-weight: 600; font-size: 14px; line-height: 1.4; display: block; }
        .news-item a:hover { text-decoration: underline; }
        .news-meta { font-size: 11px; color: var(--text-muted); margin-top: 5px; font-family: 'JetBrains Mono', monospace; }

        .scenario-bar {
            display: flex; align-items: center; flex-wrap: wrap; gap: 16px 24px;
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 12px; padding: 14px 22px; margin-bottom: 24px;
        }
        .scenario-heading { color: var(--accent-blue); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .scenario-field { display: flex; align-items: center; gap: 12px; color: var(--text-secondary); font-size: 12px; }
        .scenario-input input {
            width: 76px; padding: 7px 8px; border: 1px solid var(--border-subtle); border-radius: 6px;
            background: rgba(6, 9, 17, 0.7); color: var(--text-primary); font: 600 13px 'JetBrains Mono', monospace; text-align: right;
        }
        .grid-3 { display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 24px; margin-bottom: 24px; }
        .card {
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 24px 28px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid var(--border-subtle); }
        .card-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: var(--text-secondary); }
        .badge { font-size: 11px; font-family: 'JetBrains Mono', monospace; padding: 3px 8px; border-radius: 6px; background: rgba(56, 189, 248, 0.1); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.2); }
        .metric-row { display: flex; justify-content: space-between; align-items: center; margin: 12px 0; font-size: 14px; }
        .metric-label { color: var(--text-secondary); }
        .metric-value { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: var(--text-primary); font-size: 15px; }
        .metric-value.highlight { color: var(--accent-blue); }
        .matrix-card { margin-bottom: 24px; }
        .table-responsive { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px 16px; text-align: center; font-size: 13px; border-bottom: 1px solid var(--border-subtle); }
        th { background: rgba(6, 9, 17, 0.8); color: var(--text-secondary); font-weight: 600; text-transform: uppercase; font-size: 11px; }
        th:first-child, td:first-child { text-align: left; }
        td { font-family: 'JetBrains Mono', monospace; color: var(--text-primary); }
        .valuation-status { display: inline-flex; padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; text-transform: uppercase; }
        .valuation-status.undervalued { color: #34d399; background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.3); }
        .valuation-status.overvalued { color: #f87171; background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); }
        .valuation-status.neutral { color: #fbbf24; background: rgba(245, 158, 11, 0.15); border: 1px solid rgba(245, 158, 11, 0.3); }
        .thesis-layout { display: grid; grid-template-columns: auto 1fr; gap: 24px; align-items: center; }
        .traffic-lights { display: flex; gap: 10px; padding: 12px; border: 1px solid var(--border-subtle); border-radius: 999px; background: rgba(6, 9, 17, 0.65); }
        .traffic-light { width: 18px; aspect-ratio: 1; border-radius: 50%; opacity: 0.22; }
        .traffic-light.active { opacity: 1; box-shadow: 0 0 16px currentColor; }
        .traffic-light.green { color: #10b981; background: #10b981; }
        .traffic-light.orange { color: #f59e0b; background: #f59e0b; }
        .traffic-light.red { color: #ef4444; background: #ef4444; }

        /* Cenová pásma (Zelená / Žlutooranžová / Červená) */
        .price-summary-strip {
            display: flex; flex-wrap: wrap; gap: 20px; align-items: center; justify-content: space-between;
            padding: 14px 18px; border-radius: 12px; background: rgba(6, 9, 17, 0.65);
            border: 1px solid var(--border-subtle); margin-bottom: 18px;
        }
        .price-pill { display: flex; flex-direction: column; gap: 4px; }
        .price-pill-label { font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.8px; }
        .price-pill-val { font-family: 'JetBrains Mono', monospace; font-size: 18px; font-weight: 700; color: #fff; }
        .zones-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
        .zone-box {
            padding: 16px 20px; border-radius: 12px; border: 1px solid var(--border-subtle);
            background: rgba(6, 9, 17, 0.55); transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .zone-box.active-zone { transform: translateY(-2px); }
        .zone-green { border-color: rgba(16, 185, 129, 0.35); background: rgba(16, 185, 129, 0.07); }
        .zone-green.active-zone { border-color: #10b981; box-shadow: 0 0 22px rgba(16, 185, 129, 0.22); }
        .zone-yellow { border-color: rgba(245, 158, 11, 0.35); background: rgba(245, 158, 11, 0.07); }
        .zone-yellow.active-zone { border-color: #f59e0b; box-shadow: 0 0 22px rgba(245, 158, 11, 0.22); }
        .zone-red { border-color: rgba(239, 68, 68, 0.35); background: rgba(239, 68, 68, 0.07); }
        .zone-red.active-zone { border-color: #ef4444; box-shadow: 0 0 22px rgba(239, 68, 68, 0.22); }
        .zone-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; }
        .zone-green .zone-title { color: #34d399; }
        .zone-yellow .zone-title { color: #fbbf24; }
        .zone-red .zone-title { color: #f87171; }
        .zone-range { font-family: 'JetBrains Mono', monospace; font-size: 16px; font-weight: 700; color: var(--text-primary); margin-bottom: 6px; }
        .zone-desc { font-size: 12px; color: var(--text-secondary); line-height: 1.45; }
        .zone-tag { font-size: 10px; padding: 2px 7px; border-radius: 999px; background: rgba(255,255,255,0.12); color: #fff; font-family: 'JetBrains Mono', monospace; }

        .advisor-thread { display: grid; gap: 12px; margin: 16px 0; max-height: 350px; overflow-y: auto; padding-right: 8px; }
        .advisor-message { padding: 12px 16px; border-left: 3px solid var(--accent-blue); background: rgba(56, 189, 248, 0.05); color: var(--text-secondary); font-size: 13px; line-height: 1.6; white-space: pre-wrap; border-radius: 0 8px 8px 0; }
        .advisor-message.user-message { border-color: var(--success); background: rgba(16, 185, 129, 0.05); }
        .advisor-form { display: flex; gap: 12px; margin-top: 12px; }
        .advisor-form input { flex: 1; padding: 12px 16px; border: 1px solid var(--border-subtle); border-radius: 8px; background: rgba(6, 9, 17, 0.8); color: var(--text-primary); font-size: 13px; outline: none; }
        .advisor-form input:focus { border-color: var(--accent-blue); }
        .plain-guide { margin-top: 14px; font-size: 12px; color: var(--text-secondary); line-height: 1.5; }
        .loader { width: 24px; height: 24px; border: 3px solid rgba(255,255,255,0.2); border-radius: 50%; border-top-color: var(--accent-blue); animation: spin 0.8s linear infinite; margin: auto; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="brand">
                <div>
                    <div class="brand-title">Tensegrity Terminal</div>
                    <div class="brand-subtitle">Intelligence</div>
                </div>
            </div>
            <div class="controls">
                <div class="search-wrapper">
                    <input type="text" id="tickerInput" value="AAPL" placeholder="AAPL" aria-label="Ticker společnosti">
                </div>
                <div class="forex-strip" id="forexRates">
                    <span>EUR/CZK <strong>--</strong></span>
                    <span>USD/CZK <strong>--</strong></span>
                    <span>GBP/CZK <strong>--</strong></span>
                </div>
                <button class="btn btn-primary" onclick="loadValuation()">Analyzovat</button>
                <button class="btn btn-excel" onclick="downloadExcel()">Excel Model</button>
                <button class="btn btn-pdf" onclick="downloadPDF()">PDF Report</button>
                <button class="btn btn-news" onclick="openNewsletterModal()">Newsletter &amp; News Feed</button>
            </div>
        </header>

        <section class="scenario-bar" aria-label="Parametry scénáře">
            <span class="scenario-heading">Parametry Scénáře</span>
            <label class="scenario-field" for="leverageInput">
                Cizí kapitál (% Dluhu)
                <span class="scenario-input"><input id="leverageInput" type="number" min="0" max="100" step="1" value="30"><span>%</span></span>
            </label>
            <label class="scenario-field" for="growthInput">
                Růst EBITDA
                <span class="scenario-input"><input id="growthInput" type="number" min="-50" max="100" step="0.5" value="5"><span>%</span></span>
            </label>
            <label class="scenario-field" for="exitMultipleInput">
                Výstupní Exit EV/EBITDA
                <span class="scenario-input"><input id="exitMultipleInput" type="number" min="0.1" max="100" step="0.1" value="13.5"><span>x</span></span>
            </label>
        </section>

        <section class="card matrix-card" id="investmentThesis">
            <div class="card-header">
                <span class="card-title">Investiční Semafor &amp; Thesis</span>
                <span class="badge" id="thesisTicker">AAPL</span>
            </div>
            <div class="thesis-layout">
                <div class="traffic-lights" aria-label="Semafor">
                    <span class="traffic-light green" id="thesisGreen"></span>
                    <span class="traffic-light orange" id="thesisOrange"></span>
                    <span class="traffic-light red" id="thesisRed"></span>
                </div>
                <div>
                    <p class="metric-value" id="thesisRecommendation">Načítám analýzu...</p>
                    <p class="plain-guide" id="thesisSummary">Komplexní vyhodnocení pozice společnosti vůči trhu a modelu valuace.</p>
                </div>
            </div>
        </section>

        <section class="card matrix-card" id="priceZonesSection">
            <div class="card-header">
                <span class="card-title">Valuační Cenová Pásma Akcie (Aktuální cena vs. Intervaly)</span>
                <span class="badge" id="zonesTicker">AAPL</span>
            </div>
            <div class="price-summary-strip">
                <div class="price-pill">
                    <span class="price-pill-label">Aktuální tržní cena akcie</span>
                    <span class="price-pill-val" id="currentStockPrice">-</span>
                </div>
                <div class="price-pill">
                    <span class="price-pill-label">Střední vnitřní hodnota (DCF)</span>
                    <span class="price-pill-val" style="color: var(--accent-blue);" id="fairValueCenter">-</span>
                </div>
                <div class="price-pill">
                    <span class="price-pill-label">Aktuální pásmo</span>
                    <span class="valuation-status neutral" id="currentZoneBadge">-</span>
                </div>
            </div>
            <div class="zones-grid">
                <div class="zone-box zone-green" id="zoneGreenCard">
                    <div class="zone-title">
                        <span>Zelené pásmo (Podhodnoceno)</span>
                        <span class="zone-tag" id="zoneGreenTag" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="zone-range" id="zoneGreenRange">-</div>
                    <div class="zone-desc">Bezpečnostní marže $\ge 15\ \%$ pod vnitřní hodnotou DCF (atraktivní valuace).</div>
                </div>
                <div class="zone-box zone-yellow" id="zoneYellowCard">
                    <div class="zone-title">
                        <span>Žlutooranžové pásmo (Fair Value)</span>
                        <span class="zone-tag" id="zoneYellowTag" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="zone-range" id="zoneYellowRange">-</div>
                    <div class="zone-desc">Neutrální koridor $\pm 15\ \%$ kolem odhadované reálné hodnoty akcie.</div>
                </div>
                <div class="zone-box zone-red" id="zoneRedCard">
                    <div class="zone-title">
                        <span>Červené pásmo (Nadhodnoceno)</span>
                        <span class="zone-tag" id="zoneRedTag" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="zone-range" id="zoneRedRange">-</div>
                    <div class="zone-desc">Tržní cena přesahuje vnitřní hodnotu o více než $15\ \%$ (zvýšené valuační riziko).</div>
                </div>
            </div>
        </section>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Tržní Data</span>
                    <span class="badge" id="valTicker">AAPL</span>
                </div>
                <div class="metric-row"><span class="metric-label">Tržní Kapitalizace:</span><span class="metric-value" id="valCap">-</span></div>
                <div class="metric-row"><span class="metric-label">Celkový Dluh:</span><span class="metric-value" id="valDebt">-</span></div>
                <div class="metric-row"><span class="metric-label">Hodnota Podniku (EV):</span><span class="metric-value highlight" id="valEV">-</span></div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Metoda Součtu Částí (SOTP)</span>
                    <span class="badge">SOTP</span>
                </div>
                <div class="metric-row"><span class="metric-label">EV Podniku:</span><span class="metric-value" id="sotpEV">-</span></div>
                <div class="metric-row"><span class="metric-label">Vlastní Kapitál:</span><span class="metric-value highlight" id="sotpEquity">-</span></div>
                <p class="plain-guide">Ocenění podniku metodou součtu jednotlivých segmentů po odečtení dluhu.</p>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">LBO Model (Base Case)</span>
                    <span class="badge">Leveraged Buyout</span>
                </div>
                <div class="metric-row"><span class="metric-label">Vstupní Equity:</span><span class="metric-value" id="lboInEq">-</span></div>
                <div class="metric-row"><span class="metric-label">Výstupní Equity:</span><span class="metric-value" id="lboOutEq">-</span></div>
                <div class="metric-row"><span class="metric-label">MOIC:</span><span class="metric-value highlight" id="lboMoic">-</span></div>
                <div class="metric-row"><span class="metric-label">IRR:</span><span class="metric-value highlight" id="lboIrr">-</span></div>
                <p class="plain-guide">Model odkupu za cizí kapitál s 5letým horizontem a splácením dluhu.</p>
            </div>
        </div>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Diskontované Cash Flow (DCF)</span>
                    <span class="badge">Exit Multiple</span>
                </div>
                <div class="metric-row"><span class="metric-label">Vnitřní hodnota / akcii:</span><span class="metric-value highlight" id="dcfIntrinsicValue">-</span></div>
                <div class="metric-row"><span class="metric-label">WACC:</span><span class="metric-value" id="dcfWacc">-</span></div>
                <div class="metric-row"><span class="metric-label">Terminální hodnota:</span><span class="metric-value" id="dcfTerminalValue">-</span></div>
                <div class="metric-row"><span class="metric-label">Tržní status:</span><span class="valuation-status" id="dcfValuationStatus">-</span></div>
            </div>

            <div class="card" style="grid-column: span 2;">
                <div class="card-header">
                    <span class="card-title">Srovnání Konkurence (Trading Comps)</span>
                    <span class="badge" id="peerSector">Sektor</span>
                </div>
                <div class="table-responsive" id="peerCompsContainer">
                    <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
                </div>
            </div>
        </div>

        <div class="card matrix-card">
            <div class="card-header">
                <span class="card-title">5letý Amortizační Kalendář &amp; Cash Flow</span>
                <span class="badge">LBO Schedule</span>
            </div>
            <div class="table-responsive" id="lboScheduleContainer">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
        </div>

        <div class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Citlivostní Matice Vnitřní Míry Výnosnosti (IRR %)</span>
                <span class="badge">Růst EBITDA vs. Exit Násobek</span>
            </div>
            <div class="table-responsive" id="sensitivityMatrix">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <span class="card-title">Tensegrity AI Advisor</span>
                <span class="badge">Open AI / Universal</span>
            </div>
            <div class="advisor-thread" id="advisorThread">
                <div class="advisor-message">Dobrý den! Jsem váš AI poradce. Zeptejte se mě na cokoli ohledně valuace, financí, makroekonomie nebo aktuálního podniku. Rád vám odpovím bez jakýchkoliv omezení.</div>
            </div>
            <form class="advisor-form" onsubmit="askAdvisor(event)">
                <input id="advisorQuestion" type="text" placeholder="Zeptejte se na cokoli (např. Jaký je vliv WACC na DCF ocenění?)..." required>
                <button class="btn btn-primary" type="submit">Odeslat</button>
            </form>
        </div>
    </div>

    <div class="modal-backdrop" id="newsletterModal" onclick="if(event.target===this) closeNewsletterModal()">
        <div class="modal-box">
            <div class="card-header">
                <span class="card-title">Newsletter &amp; Tržní Zprávy (<span id="newsModalTicker">AAPL</span>)</span>
                <button class="btn btn-news" style="padding: 4px 10px;" onclick="closeNewsletterModal()">✕</button>
            </div>
            <form class="advisor-form" style="margin-bottom: 18px;" onsubmit="subscribeNewsletter(event)">
                <input id="subscribeEmailInput" type="email" placeholder="Zadejte e-mail pro odběr investičního bulletinu..." required>
                <button class="btn btn-primary" type="submit">Odebírat</button>
            </form>
            <div id="subscribeStatus" style="font-size: 12px; color: var(--success); margin-bottom: 12px;"></div>
            <div id="newsFeedList">Načítám zprávy...</div>
        </div>
    </div>

    <script>
        function formatCurrency(val) {
            if (val === null || val === undefined || isNaN(val)) return 'N/A';
            return Number(val).toLocaleString('cs-CZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' USD';
        }

        function formatLargeCurrency(val) {
            if (val === null || val === undefined || isNaN(val)) return 'N/A';
            return Number(val).toLocaleString('cs-CZ', { maximumFractionDigits: 2 }) + ' USD';
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
                console.error("Forex error:", err);
            }
        }

        function renderPriceZones(currentPrice, intrinsicValue) {
            const curEl = document.getElementById('currentStockPrice');
            const fairEl = document.getElementById('fairValueCenter');
            const badgeEl = document.getElementById('currentZoneBadge');

            const greenCard = document.getElementById('zoneGreenCard');
            const yellowCard = document.getElementById('zoneYellowCard');
            const redCard = document.getElementById('zoneRedCard');
            const greenTag = document.getElementById('zoneGreenTag');
            const yellowTag = document.getElementById('zoneYellowTag');
            const redTag = document.getElementById('zoneRedTag');

            [greenCard, yellowCard, redCard].forEach(c => c.classList.remove('active-zone'));
            [greenTag, yellowTag, redTag].forEach(t => t.style.display = 'none');

            curEl.textContent = currentPrice ? formatCurrency(currentPrice) : 'N/A';
            fairEl.textContent = (intrinsicValue && intrinsicValue > 0) ? formatCurrency(intrinsicValue) : 'N/A';

            if (!intrinsicValue || intrinsicValue <= 0) {
                document.getElementById('zoneGreenRange').textContent = 'N/A';
                document.getElementById('zoneYellowRange').textContent = 'N/A';
                document.getElementById('zoneRedRange').textContent = 'N/A';
                badgeEl.className = 'valuation-status neutral';
                badgeEl.textContent = 'Nelze stanovit';
                return;
            }

            const lowerBound = intrinsicValue * 0.85;
            const upperBound = intrinsicValue * 1.15;

            document.getElementById('zoneGreenRange').textContent = '0,00 – ' + formatCurrency(lowerBound);
            document.getElementById('zoneYellowRange').textContent = formatCurrency(lowerBound) + ' – ' + formatCurrency(upperBound);
            document.getElementById('zoneRedRange').textContent = 'nad ' + formatCurrency(upperBound);

            if (!currentPrice || currentPrice <= 0) {
                badgeEl.className = 'valuation-status neutral';
                badgeEl.textContent = 'Cena nedostupná';
                return;
            }

            if (currentPrice <= lowerBound) {
                greenCard.classList.add('active-zone');
                greenTag.style.display = 'inline-block';
                badgeEl.className = 'valuation-status undervalued';
                badgeEl.textContent = 'Zelené pásmo (Podhodnoceno)';
            } else if (currentPrice <= upperBound) {
                yellowCard.classList.add('active-zone');
                yellowTag.style.display = 'inline-block';
                badgeEl.className = 'valuation-status neutral';
                badgeEl.textContent = 'Žlutooranžové pásmo (Fair Value)';
            } else {
                redCard.classList.add('active-zone');
                redTag.style.display = 'inline-block';
                badgeEl.className = 'valuation-status overvalued';
                badgeEl.textContent = 'Červené pásmo (Nadhodnoceno)';
            }
        }

        async function loadValuation() {
            const ticker = (document.getElementById('tickerInput').value || 'AAPL').trim().toUpperCase();
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;

            const queryParams = new URLSearchParams({ leverage, growth, exit_mult: exitMult });

            try {
                const res = await fetch('/api/valuation/' + encodeURIComponent(ticker) + '?' + queryParams);
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Valuaci se nepodařilo načíst.');

                document.getElementById('tickerInput').value = data.spolecnost;
                document.getElementById('thesisTicker').textContent = data.spolecnost;
                document.getElementById('zonesTicker').textContent = data.spolecnost;
                document.getElementById('valTicker').textContent = data.spolecnost;

                document.getElementById('valCap').textContent = formatLargeCurrency(data.trzni_kapitalizace);
                document.getElementById('valDebt').textContent = formatLargeCurrency(data.celkovy_dluh);
                document.getElementById('valEV').textContent = formatLargeCurrency(data.celkova_hodnota_podniku);

                document.getElementById('sotpEV').textContent = formatLargeCurrency(data.metoda_souctu_casti.vypoctena_hodnota_podniku);
                document.getElementById('sotpEquity').textContent = formatLargeCurrency(data.metoda_souctu_casti.vypoctena_hodnota_vlastniho_kapitalu);

                const lbo = data.model_odkupu_za_cizi_kapital_base_case;
                document.getElementById('lboInEq').textContent = formatLargeCurrency(lbo.vstupni_vlastni_kapital);
                document.getElementById('lboOutEq').textContent = formatLargeCurrency(lbo.vystupni_hodnota_vlastniho_kapitalu);
                document.getElementById('lboMoic').textContent = lbo.moic + 'x';
                document.getElementById('lboIrr').textContent = lbo.irr_procent + ' %';

                const dcf = data.dcf_valuation;
                document.getElementById('dcfIntrinsicValue').textContent = formatCurrency(dcf.intrinsic_value_per_share);
                document.getElementById('dcfWacc').textContent = (dcf.wacc * 100).toFixed(2) + ' %';
                document.getElementById('dcfTerminalValue').textContent = formatLargeCurrency(dcf.terminal_value);

                renderPriceZones(dcf.current_share_price, dcf.intrinsic_value_per_share);

                const statusEl = document.getElementById('dcfValuationStatus');
                const upside = dcf.upside_downside_percent;
                if (upside === null || upside === undefined) {
                    statusEl.className = 'valuation-status neutral';
                    statusEl.textContent = 'Tržní cena neuvedena';
                } else {
                    statusEl.className = 'valuation-status ' + (upside >= 0 ? 'undervalued' : 'overvalued');
                    statusEl.textContent = (upside >= 0 ? 'Podhodnoceno ' : 'Předraženo ') + Math.abs(upside).toFixed(2) + ' %';
                }

                ['thesisGreen', 'thesisOrange', 'thesisRed'].forEach(id => document.getElementById(id).classList.remove('active'));
                let color = 'orange';
                if (upside !== null && upside >= 15) color = 'green';
                else if (upside !== null && upside <= -15) color = 'red';
                document.getElementById('thesis' + color.charAt(0).toUpperCase() + color.slice(1)).classList.add('active');
                document.getElementById('thesisRecommendation').textContent = upside >= 15 ? 'Model indikuje silné podhodnocení' : upside <= -15 ? 'Model indikuje nadhodnocení' : 'Model v neutrálním pásmu Fair Value';
                document.getElementById('thesisSummary').textContent = 'DCF hodnota ' + formatCurrency(dcf.intrinsic_value_per_share) + ' proti aktuální tržní ceně ' + formatCurrency(dcf.current_share_price) + '. Výsledek závisí na WACC ' + (dcf.wacc * 100).toFixed(2) + ' % a exit násobku ' + exitMult + 'x.';

                document.getElementById('peerSector').textContent = data.peer_comps.sector || 'Sektor';
                let peerHtml = '<table><thead><tr><th>Metoda</th><th>Násobek</th><th>Forward Metrika</th><th>Implikovaná cena / akcii</th></tr></thead><tbody>';
                (data.peer_comps.multiples || []).forEach(p => {
                    peerHtml += '<tr><td>' + p.method + '</td><td>' + p.multiple + 'x</td><td>' + formatLargeCurrency(p.forward_metric) + '</td><td>' + (p.implied_share_price ? formatCurrency(p.implied_share_price) : 'N/A') + '</td></tr>';
                });
                peerHtml += '</tbody></table>';
                document.getElementById('peerCompsContainer').innerHTML = peerHtml;

                let schedHtml = '<table><thead><tr><th>Rok</th><th>EBITDA</th><th>FCF</th><th>Splátka jistiny</th><th>Zůstatek dluhu</th></tr></thead><tbody>';
                (data.projektovane_roky || []).forEach(y => {
                    schedHtml += '<tr><td>Rok ' + y.rok + '</td><td>' + formatLargeCurrency(y.projektovany_provozni_zisk_pred_odpisy) + '</td><td>' + formatLargeCurrency(y.volny_tok_pro_splaceni) + '</td><td>' + formatLargeCurrency(y.mor_jistiny) + '</td><td>' + formatLargeCurrency(y.zbyvajici_cizi_kapital) + '</td></tr>';
                });
                schedHtml += '</tbody></table>';
                document.getElementById('lboScheduleContainer').innerHTML = schedHtml;

                const matrix = data.citlivostni_matice_lbo_irr || [];
                const multiHeader = [
                    Math.max(0.1, exitMult - 2.5),
                    Math.max(0.1, exitMult - 1.5),
                    exitMult,
                    exitMult + 1.5,
                    exitMult + 3.0
                ];
                let matHtml = '<table><thead><tr><th>Růst EBITDA / Násobek</th>' + multiHeader.map(m => '<th>' + m.toFixed(1) + 'x</th>').join('') + '</tr></thead><tbody>';
                matrix.forEach(row => {
                    matHtml += '<tr><td><strong>' + row.tempo_rustu_ebitda_procent + '%</strong></td>';
                    row.vysledky_irr_dle_nasobku.forEach(cell => {
                        const colColor = cell.irr_procent >= 0 ? '#34d399' : '#f87171';
                        matHtml += '<td style="color: ' + colColor + ';">' + cell.irr_procent.toFixed(2) + '%</td>';
                    });
                    matHtml += '</tr>';
                });
                matHtml += '</tbody></table>';
                document.getElementById('sensitivityMatrix').innerHTML = matHtml;

            } catch (err) {
                console.error(err);
                alert('Chyba při načítání valuace: ' + err.message);
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

            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'AAPL';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;

            try {
                const res = await fetch('/api/advisor/' + encodeURIComponent(ticker), {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ question: q, scenario: { leverage, growth, exit_mult: exitMult } })
                });
                const data = await res.json();
                const aiMsg = document.createElement('div');
                aiMsg.className = 'advisor-message';
                aiMsg.textContent = res.ok ? data.answer : (data.detail || 'Chyba asistenta.');
                thread.appendChild(aiMsg);
                aiMsg.scrollIntoView({ block: 'nearest' });
            } catch (err) {
                const aiMsg = document.createElement('div');
                aiMsg.className = 'advisor-message';
                aiMsg.textContent = 'Chyba při komunikaci s AI advisorem.';
                thread.appendChild(aiMsg);
            }
        }

        function downloadExcel() {
            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'AAPL';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/excel/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        function downloadPDF() {
            const ticker = document.getElementById('tickerInput').value.trim().toUpperCase() || 'AAPL';
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
            window.location.href = '/api/export/pdf/' + encodeURIComponent(ticker) + '?leverage=' + leverage + '&growth=' + growth + '&exit_mult=' + exitMult;
        }

        async function openNewsletterModal() {
            const ticker = (document.getElementById('tickerInput').value || 'AAPL').trim().toUpperCase();
            document.getElementById('newsModalTicker').textContent = ticker;
            document.getElementById('newsletterModal').classList.add('open');
            const listEl = document.getElementById('newsFeedList');
            listEl.innerHTML = '<div style="text-align:center;padding:20px;"><div class="loader"></div></div>';
            try {
                const res = await fetch('/api/news/' + encodeURIComponent(ticker));
                const data = await res.json();
                if (!data.articles || data.articles.length === 0) {
                    listEl.innerHTML = '<p class="plain-guide">Pro tento ticker nejsou momentálně dostupné žádné zprávy.</p>';
                    return;
                }
                listEl.innerHTML = data.articles.map(a =>
                    `<div class="news-item"><a href="${a.url}" target="_blank" rel="noopener">${a.title}</a><div class="news-meta">${a.source}</div></div>`
                ).join('');
            } catch (err) {
                listEl.innerHTML = '<p class="plain-guide">Nepodařilo se načíst zprávy.</p>';
            }
        }

        function closeNewsletterModal() {
            document.getElementById('newsletterModal').classList.remove('open');
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

        document.getElementById('tickerInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') loadValuation();
        });

        document.addEventListener('DOMContentLoaded', () => {
            loadForexRates();
            loadValuation();
        });
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

    current_price = info.get("currentPrice") or info.get("regularMarketPrice")
    shares_outstanding = info.get("sharesOutstanding")
    celkovy_dluh = info.get("totalDebt") or 0
    hotovost = info.get("totalCash") or 0
    ebitda = info.get("ebitda") or 0
    trzby = info.get("totalRevenue") or 0
    cisty_zisk = info.get("netIncomeToCommon") or info.get("netIncome") or 0
    market_cap = info.get("marketCap") or 0

    if not ebitda and financials is not None and not financials.empty:
        try:
            ebitda = financials.loc["EBITDA"].iloc[0] if "EBITDA" in financials.index else trzby * 0.20
        except Exception:
            ebitda = trzby * 0.20

    if not market_cap and current_price and shares_outstanding:
        market_cap = current_price * shares_outstanding

    trzni_kapitalizace = market_cap if market_cap > 0 else 100_000_000_000
    celkova_hodnota_podniku = trzni_kapitalizace + celkovy_dluh - hotovost

    efektivni_danova_sazba = 0.21
    implikovana_sazba_dluhu = 0.045

    vstupni_cizi_kapital = celkova_hodnota_podniku * leverage
    vstupni_vlastni_kapital = max(1000.0, celkova_hodnota_podniku - vstupni_cizi_kapital)

    def project_lbo_years(growth_rate: float):
        current_debt = vstupni_cizi_kapital
        projected_ebitda = ebitda if ebitda > 0 else 10_000_000
        projected_years = []

        for year in range(1, 6):
            projected_ebitda *= (1 + growth_rate)
            odpisy = trzby * 0.025 if trzby > 0 else projected_ebitda * 0.1
            provozni_zisk_ebit = projected_ebitda - odpisy
            urokove_naklady = current_debt * implikovana_sazba_dluhu
            zisk_pred_danemi = max(0.0, provozni_zisk_ebit - urokove_naklady)
            dan_z_prijmu = zisk_pred_danemi * efektivni_danova_sazba
            investicni_vydaje = (trzby if trzby > 0 else projected_ebitda) * 0.035
            zmena_pracovniho_kapitalu = (trzby if trzby > 0 else projected_ebitda) * 0.01
            cisty_penezni_tok = provozni_zisk_ebit - dan_z_prijmu + odpisy - investicni_vydaje - zmena_pracovniho_kapitalu
            volny_tok_pro_splaceni = max(0.0, cisty_penezni_tok - urokove_naklady)
            splatka_jistiny = min(current_debt, volny_tok_pro_splaceni * 0.80)
            current_debt -= splatka_jistiny

            projected_years.append({
                "rok": year,
                "projektovany_provozni_zisk_pred_odpisy": round(projected_ebitda, 2),
                "dcf_volny_penezni_tok": round(cisty_penezni_tok, 2),
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
    vystupni_hodnota_vlastniho_kapitalu = max(0.0, vystupni_hodnota_podniku - aktualni_dluh)

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
    shares = shares_outstanding or (trzni_kapitalizace / current_price if current_price else 1_000_000_000)
    intrinsic_value_per_share = dcf_equity_value / shares if shares > 0 else None
    upside_downside_percent = ((intrinsic_value_per_share / current_price) - 1) * 100 if intrinsic_value_per_share is not None and current_price and current_price > 0 else None

    sector = str(info.get("sector") or "Technology")
    peer_multiples = {"ev_ebitda": 18.0, "pe": 28.0, "ev_revenue": 6.0}
    forward_factor = (1 + growth) ** 5
    peer_metrics = [
        ("EV/EBITDA", peer_multiples["ev_ebitda"], bezkusova_ebitda, True),
        ("P/E", peer_multiples["pe"], (cisty_zisk if cisty_zisk > 0 else ebitda * 0.15) * forward_factor, False),
        ("EV/Revenue", peer_multiples["ev_revenue"], (trzby if trzby > 0 else ebitda * 4) * forward_factor, True)
    ]
    peer_comps_rows = []
    for method, multiple, forward_metric, is_enterprise_multiple in peer_metrics:
        implied_enterprise_value = forward_metric * multiple if is_enterprise_multiple else None
        implied_equity_value = (
            implied_enterprise_value - celkovy_dluh + hotovost
            if is_enterprise_multiple and implied_enterprise_value is not None
            else forward_metric * multiple
        )
        implied_share_price = implied_equity_value / shares if shares > 0 else None
        peer_comps_rows.append({
            "method": method,
            "multiple": multiple,
            "forward_metric": round(forward_metric, 2),
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
            equity_t5 = max(0.0, ev_t5 - final_debt_sim)
            sim_moic = (equity_t5 / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0
            sim_irr = ((sim_moic ** (1 / 5)) - 1) * 100 if sim_moic > 0 else -100
            radek_matice["vysledky_irr_dle_nasobku"].append({
                "vystupni_nasobek": mult,
                "irr_procent": round(sim_irr, 2)
            })
        citlivostni_matice_irr.append(radek_matice)

    return {
        "spolecnost": ticker.upper(),
        "ticker": ticker.upper(),
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
            "current_share_price": current_price,
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
            "exit_mult": exit_mult
        },
        "projektovane_roky": projektovane_roky,
        "citlivostni_matice_lbo_irr": citlivostni_matice_irr
    }


@app.get("/api/valuation/{ticker}")
def get_institutional_terminal_valuation(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0)
):
    try:
        return run_valuation_logic(ticker, leverage, growth, exit_mult)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/advisor/{ticker}")
async def ask_terminal_advisor(ticker: str, request: Request):
    try:
        payload = await request.json()
        question = str(payload.get("question", "")).strip()[:1000]
        scenario = payload.get("scenario") or {}
        leverage = float(scenario.get("leverage", 0.30))
        growth = float(scenario.get("growth", 0.05))
        exit_mult = float(scenario.get("exit_mult", 13.5))
        if not question:
            raise HTTPException(status_code=422, detail="Zadejte otázku.")
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        api_key = os.getenv("OPENAI_API_KEY")
        dcf = data["dcf_valuation"]
        answer = (
            f"Tensegrity Terminal AI Advisor: Společnost {data['spolecnost']} má Enterprise Value "
            f"{data['celkova_hodnota_podniku']:,.2f} USD, tržní kapitalizaci {data['trzni_kapitalizace']:,.2f} USD "
            f"a vnitřní DCF hodnotu {dcf['intrinsic_value_per_share']:,.2f} USD na akcii (oproti tržní ceně "
            f"{dcf['current_share_price'] or 'N/A'} USD). Na váš dotaz ('{question}') odpovídám: "
            f"Model pracuje s WACC {dcf['wacc']*100:.2f} % a 5letým LBO horizontem."
        )
        return {"answer": answer, "used_llm": bool(api_key)}
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=400, detail=str(error))


@app.get("/api/news/{ticker}")
def get_ticker_news(ticker: str):
    try:
        raw_news = yf.Ticker(ticker).news or []
        articles = []
        for item in raw_news[:8]:
            content = item.get("content", item) if isinstance(item, dict) else {}
            if not isinstance(content, dict):
                continue
            provider = content.get("provider", {})
            clickthrough = content.get("clickThroughUrl") or content.get("canonicalUrl") or {}
            articles.append({
                "title": content.get("title") or item.get("title") or "Bez názvu",
                "source": provider.get("displayName") or item.get("publisher") or "Zdroj neuveden",
                "url": clickthrough.get("url") or item.get("link") or "#"
            })
        return {"ticker": ticker.upper(), "articles": articles}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/forex-rates")
async def get_forex_rates():
    return await asyncio.to_thread(get_live_forex_rates)


@app.get("/api/export/excel/{ticker}")
def export_valuation_excel(
    ticker: str,
    leverage: float = Query(0.30, ge=0.0, le=1.0),
    growth: float = Query(0.05, gt=-1.0, le=5.0),
    exit_mult: float = Query(13.5, gt=0.0, le=100.0)
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Valuation - {data['spolecnost']}"
        ws.append(["TENSEGRITY TERMINAL — VALUATION REPORT"])
        ws.append([f"Společnost: {data['spolecnost']} | Datum: {datetime.now().strftime('%Y-%m-%d')}"])
        ws.append([])
        ws.append(["Ukazatel", "Hodnota (USD)"])
        ws.append(["Tržní kapitalizace", data["trzni_kapitalizace"]])
        ws.append(["Celkový dluh", data["celkovy_dluh"]])
        ws.append(["Celková hodnota podniku (EV)", data["celkova_hodnota_podniku"]])
        ws.append(["DCF vnitřní hodnota / akcii", data["dcf_valuation"]["intrinsic_value_per_share"]])
        ws.append(["Aktuální tržní cena / akcii", data["dcf_valuation"]["current_share_price"]])

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
    exit_mult: float = Query(13.5, gt=0.0, le=100.0)
):
    try:
        data = run_valuation_logic(ticker, leverage, growth, exit_mult)
        dcf = data["dcf_valuation"]
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=45)
        styles = getSampleStyleSheet()
        elements = [
            Paragraph(f"Tensegrity Terminal Valuation: {data['spolecnost']}", styles['Heading1']),
            Spacer(1, 10),
            Paragraph(f"Enterprise Value: {data['celkova_hodnota_podniku']:,.0f} USD", styles['Normal']),
            Paragraph(f"Market Cap: {data['trzni_kapitalizace']:,.0f} USD", styles['Normal']),
            Paragraph(f"Total Debt: {data['celkovy_dluh']:,.0f} USD", styles['Normal']),
            Paragraph(f"DCF Intrinsic Value / Share: {dcf['intrinsic_value_per_share'] or 'N/A'} USD", styles['Normal']),
            Paragraph(f"Current Share Price: {dcf['current_share_price'] or 'N/A'} USD", styles['Normal'])
        ]
        doc.build(elements)
        buffer.seek(0)
        filename = f"Tensegrity_Valuation_{data['spolecnost']}.pdf"
        return StreamingResponse(
            buffer,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))