import os
import json
import math
import time
import smtplib
import threading
import datetime
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Dict, Any, List, Optional

import yfinance as yf
import pandas as pd
import numpy as np
import requests
from fastapi import FastAPI, HTTPException, Response, Query
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Volitelné knihovny pro export do Excelu a PDF
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

# ==============================================================================
# KONFIGURACE APLIKACE A PROSTŘEDÍ
# ==============================================================================
app = FastAPI(title="Tensegrity Terminal Intelligence", version="12.5")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SUBSCRIBERS_FILE = "subscribers.json"
CACHE_STORE: Dict[str, Any] = {}
CACHE_TTL = 180  # 3 minuty mezipaměti pro rychlou odezvu

# Přihlašovací údaje pro odesílání e-mailů (nastavují se v prostředí serveru / Renderu)
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")


# ==============================================================================
# DATOVÉ MODELY PRO POŽADAVKY
# ==============================================================================
class OneOffNewsletterRequest(BaseModel):
    email: str
    ticker: str = "NVDA"
    leverage: float = 30.0
    growth: float = 5.0
    exit_mult: float = 13.5


class RegularSubscriptionRequest(BaseModel):
    email: str
    ticker: str = "NVDA"
    schedule_day: str = "MONDAY"     # DAILY, WORKDAYS, MONDAY, FRIDAY
    schedule_time: str = "08:00"     # HH:MM
    leverage: float = 30.0
    growth: float = 5.0
    exit_mult: float = 13.5


# ==============================================================================
# POMOCNÉ FUNKCE PRO CACHE, FORMÁTOVÁNÍ A SPRÁVU ODBĚRATELŮ
# ==============================================================================
def get_cached(key: str) -> Optional[Any]:
    entry = CACHE_STORE.get(key)
    if entry and (time.time() - entry["ts"] < CACHE_TTL):
        return entry["data"]
    return None


def set_cached(key: str, data: Any) -> None:
    CACHE_STORE[key] = {"ts": time.time(), "data": data}


def load_subscribers() -> List[Dict[str, Any]]:
    if not os.path.exists(SUBSCRIBERS_FILE):
        return []
    try:
        with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                # Migrace případných starších textových záznamů na strukturované objekty
                normalized = []
                for item in data:
                    if isinstance(item, str):
                        normalized.append({
                            "email": item,
                            "ticker": "NVDA",
                            "schedule_day": "MONDAY",
                            "schedule_time": "08:00",
                            "leverage": 30.0,
                            "growth": 5.0,
                            "exit_mult": 13.5,
                            "last_sent_date": ""
                        })
                    elif isinstance(item, dict):
                        normalized.append(item)
                return normalized
    except Exception:
        pass
    return []


def save_subscribers(subscribers: List[Dict[str, Any]]) -> None:
    with open(SUBSCRIBERS_FILE, "w", encoding="utf-8") as f:
        json.dump(subscribers, f, ensure_ascii=False, indent=2)


def safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None or (isinstance(val, float) and math.isnan(val)):
            return default
        return float(val)
    except Exception:
        return default


# ==============================================================================
# FINANČNÍ VÝPOČTY A VALUAČNÍ MODELY (DCF, LBO, SOTP, COMPS, CENOVÁ PÁSMA)
# ==============================================================================
def get_live_forex() -> Dict[str, float]:
    cached = get_cached("forex_rates")
    if cached:
        return cached

    rates = {"EURCZK": 25.10, "USDCZK": 23.20, "GBPCZK": 29.85}
    pairs = {"EURCZK": "EURCZK=X", "USDCZK": "CZK=X", "GBPCZK": "GBPCZK=X"}
    for key, symbol in pairs.items():
        try:
            t = yf.Ticker(symbol)
            hist = t.history(period="2d")
            if not hist.empty:
                rates[key] = round(float(hist["Close"].iloc[-1]), 2)
        except Exception:
            pass

    set_cached("forex_rates", rates)
    return rates


def fetch_ticker_news(ticker: str, limit: int = 6) -> List[Dict[str, str]]:
    ticker = ticker.upper().strip()
    cache_key = f"news_{ticker}"
    cached = get_cached(cache_key)
    if cached is not None:
        return cached

    articles: List[Dict[str, str]] = []
    try:
        t = yf.Ticker(ticker)
        raw_news = t.news or []
        for item in raw_news[:limit]:
            content = item.get("content", item)
            title = content.get("title") or item.get("title") or "Tržní aktualita"
            provider = (
                content.get("provider", {}).get("displayName")
                if isinstance(content.get("provider"), dict)
                else item.get("publisher", "Finanční zpravodajství")
            )
            link = "#"
            if isinstance(content.get("canonicalUrl"), dict):
                link = content["canonicalUrl"].get("url", "#")
            elif item.get("link"):
                link = item["link"]

            pub_date = content.get("pubDate") or ""
            if not pub_date and item.get("providerPublishTime"):
                pub_date = datetime.datetime.fromtimestamp(
                    item["providerPublishTime"]
                ).strftime("%d.%m.%Y %H:%M")

            articles.append({
                "title": title,
                "publisher": provider or "Tržní zdroj",
                "link": link,
                "published": pub_date[:16].replace("T", " ") if pub_date else "Aktuálně"
            })
    except Exception:
        pass

    if not articles:
        articles = [
            {
                "title": f"Analytický přehled a tržní sentiment pro titul {ticker}",
                "publisher": "Tensegrity Intelligence",
                "link": f"https://finance.yahoo.com/quote/{ticker}",
                "published": datetime.datetime.now().strftime("%d.%m.%Y")
            }
        ]

    set_cached(cache_key, articles)
    return articles


def compute_valuation(
    ticker: str,
    leverage_pct: float = 30.0,
    ebitda_growth_pct: float = 5.0,
    exit_multiple: float = 13.5
) -> Dict[str, Any]:
    ticker = ticker.upper().strip()
    cache_key = f"val_{ticker}_{leverage_pct}_{ebitda_growth_pct}_{exit_multiple}"
    cached = get_cached(cache_key)
    if cached:
        return cached

    try:
        stock = yf.Ticker(ticker)
        info = stock.info or {}
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Nepodařilo se načíst data pro {ticker}: {e}")

    company_name = info.get("longName") or info.get("shortName") or ticker
    currency = info.get("currency") or "USD"
    sector = info.get("sector") or "Technologie / Průmysl"
    industry = info.get("industry") or "Obecný sektor"

    current_price = safe_float(
        info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"),
        100.0
    )
    shares = safe_float(info.get("sharesOutstanding"), 1_000_000_000.0)
    market_cap = safe_float(info.get("marketCap"), current_price * shares)
    total_debt = safe_float(info.get("totalDebt"), market_cap * 0.15)
    total_cash = safe_float(info.get("totalCash"), market_cap * 0.10)
    net_debt = total_debt - total_cash
    enterprise_value = market_cap + net_debt

    revenue = safe_float(info.get("totalRevenue"), market_cap * 0.35)
    ebitda = safe_float(info.get("ebitda"), revenue * 0.28)
    if ebitda <= 0:
        ebitda = max(revenue * 0.20, market_cap * 0.06)

    fcf = safe_float(info.get("freeCashflow"), ebitda * 0.65)
    if fcf <= 0:
        fcf = ebitda * 0.55

    beta = safe_float(info.get("beta"), 1.15)
    risk_free = 0.042
    erp = 0.052
    cost_of_equity = risk_free + beta * erp
    cost_of_debt = 0.055 * (1 - 0.21)

    debt_weight = max(0.0, min(0.90, leverage_pct / 100.0))
    equity_weight = 1.0 - debt_weight
    wacc = max(0.065, min(0.18, equity_weight * cost_of_equity + debt_weight * cost_of_debt))

    g = ebitda_growth_pct / 100.0
    terminal_g = 0.025

    # 5letá DCF projekce
    projected_fcf = []
    pv_fcfs = 0.0
    curr_fcf = fcf
    for yr in range(1, 6):
        curr_fcf = curr_fcf * (1.0 + g)
        disc_fcf = curr_fcf / ((1.0 + wacc) ** yr)
        pv_fcfs += disc_fcf
        projected_fcf.append(round(curr_fcf, 2))

    terminal_ebitda = ebitda * ((1.0 + g) ** 5)
    terminal_value_mult = terminal_ebitda * exit_multiple
    terminal_value_gordon = (projected_fcf[-1] * (1.0 + terminal_g)) / max(0.02, (wacc - terminal_g))
    blended_tv = 0.5 * terminal_value_mult + 0.5 * terminal_value_gordon
    pv_terminal = blended_tv / ((1.0 + wacc) ** 5)

    implied_ev = pv_fcfs + pv_terminal
    implied_equity_val = max(0.0, implied_ev - net_debt)
    dcf_intrinsic_price = implied_equity_val / shares if shares > 0 else current_price

    # Valuační cenová pásma (+/- 15 % kolem střední vnitřní hodnoty)
    lower_band = round(dcf_intrinsic_price * 0.85, 2)
    upper_band = round(dcf_intrinsic_price * 1.15, 2)

    if current_price < lower_band:
        signal_color = "green"
        signal_title = "Model indikuje podhodnocení"
        band_label = "ZELENÉ PÁSMO (PODHODNOCENO)"
    elif current_price <= upper_band:
        signal_color = "amber"
        signal_title = "Model indikuje férové ocenění"
        band_label = "ŽLUTOORANŽOVÉ PÁSMO (FÉROVÁ HODNOTA)"
    else:
        signal_color = "red"
        signal_title = "Model indikuje nadhodnocení"
        band_label = "ČERVENÉ PÁSMO (NADHODNOCENO)"

    upside_pct = ((dcf_intrinsic_price - current_price) / current_price * 100.0) if current_price > 0 else 0.0

    # LBO model (5 let)
    entry_ev = enterprise_value
    lbo_debt_start = entry_ev * debt_weight
    lbo_equity_start = max(entry_ev * 0.1, entry_ev - lbo_debt_start)
    interest_rate = 0.068
    amortization_schedule = []
    rem_debt = lbo_debt_start
    curr_ebitda = ebitda
    for yr in range(1, 6):
        curr_ebitda *= (1.0 + g)
        interest = rem_debt * interest_rate
        cash_avail = max(0.0, curr_ebitda * 0.62 - interest)
        repayment = min(rem_debt, cash_avail * 0.75)
        rem_debt = max(0.0, rem_debt - repayment)
        amortization_schedule.append({
            "year": f"Rok {yr}",
            "ebitda": round(curr_ebitda / 1e6, 1),
            "interest": round(interest / 1e6, 1),
            "repayment": round(repayment / 1e6, 1),
            "ending_debt": round(rem_debt / 1e6, 1)
        })

    exit_ev_lbo = curr_ebitda * exit_multiple
    exit_equity_lbo = max(0.0, exit_ev_lbo - rem_debt)
    moic = exit_equity_lbo / lbo_equity_start if lbo_equity_start > 0 else 1.0
    irr = ((moic ** (1.0 / 5.0)) - 1.0) * 100.0 if moic > 0 else 0.0

    # SOTP (Sum-of-the-Parts) rozklad
    core_ev = implied_ev * 0.72
    growth_ev = implied_ev * 0.28
    sotp_price = max(0.0, (core_ev + growth_ev - net_debt) / shares) if shares > 0 else dcf_intrinsic_price

    # Citlivostní matice (WACC vs. Exit násobek)
    wacc_steps = [round(wacc - 0.015, 4), round(wacc - 0.0075, 4), round(wacc, 4), round(wacc + 0.0075, 4), round(wacc + 0.015, 4)]
    mult_steps = [round(exit_multiple - 3.0, 1), round(exit_multiple - 1.5, 1), round(exit_multiple, 1), round(exit_multiple + 1.5, 1), round(exit_multiple + 3.0, 1)]
    sensitivity_matrix = []
    for w_val in wacc_steps:
        row_prices = []
        pv_f = sum(projected_fcf[i] / ((1.0 + w_val) ** (i + 1)) for i in range(5))
        for m_val in mult_steps:
            tv_m = terminal_ebitda * m_val
            tv_g = (projected_fcf[-1] * (1.0 + terminal_g)) / max(0.02, (w_val - terminal_g))
            pv_t = (0.5 * tv_m + 0.5 * tv_g) / ((1.0 + w_val) ** 5)
            eq_v = max(0.0, pv_f + pv_t - net_debt)
            row_prices.append(round(eq_v / shares, 2) if shares > 0 else 0.0)
        sensitivity_matrix.append({
            "wacc_pct": round(w_val * 100.0, 2),
            "prices": row_prices
        })

    # Srovnávací násobky (Trading Comps)
    ev_ebitda_curr = round(enterprise_value / ebitda, 1) if ebitda > 0 else 0.0
    pe_ratio = round(safe_float(info.get("trailingPE"), current_price / max(0.1, current_price * 0.05)), 1)
    ev_rev = round(enterprise_value / revenue, 2) if revenue > 0 else 0.0

    # Historický cenový graf (6 měsíců)
    history_labels = []
    history_prices = []
    try:
        hist = stock.history(period="6mo", interval="1wk")
        for idx, row in hist.iterrows():
            history_labels.append(idx.strftime("%d.%m."))
            history_prices.append(round(float(row["Close"]), 2))
    except Exception:
        history_labels = ["M-5", "M-4", "M-3", "M-2", "M-1", "Dnes"]
        history_prices = [round(current_price * x, 2) for x in [0.88, 0.91, 0.95, 0.93, 0.98, 1.0]]

    result = {
        "ticker": ticker,
        "company_name": company_name,
        "currency": currency,
        "sector": sector,
        "industry": industry,
        "current_price": round(current_price, 2),
        "dcf_price": round(dcf_intrinsic_price, 2),
        "sotp_price": round(sotp_price, 2),
        "lower_band": lower_band,
        "upper_band": upper_band,
        "signal_color": signal_color,
        "signal_title": signal_title,
        "band_label": band_label,
        "upside_pct": round(upside_pct, 1),
        "wacc_pct": round(wacc * 100.0, 2),
        "exit_multiple": exit_multiple,
        "market_cap_m": round(market_cap / 1e6, 1),
        "enterprise_value_m": round(enterprise_value / 1e6, 1),
        "net_debt_m": round(net_debt / 1e6, 1),
        "revenue_m": round(revenue / 1e6, 1),
        "ebitda_m": round(ebitda / 1e6, 1),
        "fcf_m": round(fcf / 1e6, 1),
        "lbo": {
            "entry_ev_m": round(entry_ev / 1e6, 1),
            "initial_debt_m": round(lbo_debt_start / 1e6, 1),
            "initial_equity_m": round(lbo_equity_start / 1e6, 1),
            "exit_equity_m": round(exit_equity_lbo / 1e6, 1),
            "moic": round(moic, 2),
            "irr_pct": round(irr, 1),
            "schedule": amortization_schedule
        },
        "comps": {
            "ev_ebitda": ev_ebitda_curr,
            "sector_ev_ebitda": round(max(10.0, exit_multiple), 1),
            "pe_ratio": pe_ratio,
            "ev_revenue": ev_rev
        },
        "sensitivity": {
            "multiples": mult_steps,
            "rows": sensitivity_matrix
        },
        "history": {
            "labels": history_labels,
            "prices": history_prices
        }
    }
    set_cached(cache_key, result)
    return result


# ==============================================================================
# ODESÍLÁNÍ E-MAILOVÉHO NEWSLETTERU (JEDNORÁZOVĚ I PLÁNOVANĚ)
# ==============================================================================
def build_newsletter_html(val_data: Dict[str, Any], news_list: List[Dict[str, str]], mode_label: str) -> str:
    ticker = val_data["ticker"]
    curr = val_data["currency"]
    news_items_html = "".join(
        f"<li style='margin-bottom:10px;'>"
        f"<a href='{n['link']}' style='color:#0ea5e9;text-decoration:none;font-weight:600;'>{n['title']}</a>"
        f"<div style='font-size:12px;color:#94a3b8;'>{n['publisher']} • {n['published']}</div>"
        f"</li>"
        for n in news_list[:5]
    )

    return f"""
    <html>
    <body style="font-family:Inter,Arial,sans-serif;background-color:#0b1120;color:#e2e8f0;padding:24px;">
      <div style="max-width:640px;margin:0 auto;background:#111827;border:1px solid #1e293b;border-radius:12px;padding:24px;">
        <div style="font-size:12px;letter-spacing:2px;color:#38bdf8;font-weight:700;">TENSEGRITY TERMINAL • {mode_label.upper()}</div>
        <h2 style="margin:8px 0 4px 0;color:#ffffff;">Investiční bulletin: {val_data['company_name']} ({ticker})</h2>
        <p style="color:#94a3b8;font-size:13px;margin-top:0;">Vygenerováno dne {datetime.datetime.now().strftime('%d.%m.%Y v %H:%M')}</p>

        <div style="background:#0f172a;border:1px solid #1e293b;border-radius:8px;padding:16px;margin:18px 0;">
          <div style="font-size:13px;color:#94a3b8;">Aktuální tržní cena: <strong style="color:#ffffff;">{val_data['current_price']} {curr}</strong></div>
          <div style="font-size:13px;color:#94a3b8;margin-top:6px;">Střední vnitřní hodnota modelu: <strong style="color:#38bdf8;">{val_data['dcf_price']} {curr}</strong> ({val_data['upside_pct']:+.1f} %)</div>
          <div style="font-size:13px;color:#94a3b8;margin-top:6px;">Valuační pásmo: <strong style="color:#f8fafc;">{val_data['band_label']}</strong></div>
          <div style="font-size:12px;color:#64748b;margin-top:8px;">
            Zelené pásmo: 0,00 – {val_data['lower_band']} {curr} | Férové pásmo: {val_data['lower_band']} – {val_data['upper_band']} {curr} | Červené pásmo: nad {val_data['upper_band']} {curr}
          </div>
        </div>

        <h3 style="color:#f8fafc;font-size:15px;border-bottom:1px solid #1e293b;padding-bottom:6px;">Aktuální tržní zprávy ({ticker})</h3>
        <ul style="padding-left:18px;margin-top:12px;">
          {news_items_html}
        </ul>
      </div>
    </body>
    </html>
    """


def dispatch_email_report(
    recipient_email: str,
    ticker: str,
    leverage: float = 30.0,
    growth: float = 5.0,
    exit_mult: float = 13.5,
    mode_label: str = "Jednorázový přehled"
) -> Dict[str, Any]:
    val_data = compute_valuation(ticker, leverage, growth, exit_mult)
    news_list = fetch_ticker_news(ticker, limit=5)
    html_body = build_newsletter_html(val_data, news_list, mode_label)

    if not SENDER_EMAIL or not SENDER_PASSWORD:
        return {
            "sent_via_smtp": False,
            "message": (
                f"Přehled pro {ticker} byl připraven pro {recipient_email}, ale na serveru zatím nejsou "
                f"nastaveny proměnné prostředí SENDER_EMAIL a SENDER_PASSWORD pro přímé odeslání pošty."
            )
        }

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Tensegrity Terminal | {ticker}: {val_data['signal_title']} ({val_data['current_price']} {val_data['currency']})"
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, [recipient_email], msg.as_string())
        return {
            "sent_via_smtp": True,
            "message": f"Investiční bulletin pro {ticker} byl úspěšně odeslán na adresu {recipient_email}."
        }
    except Exception as e:
        return {
            "sent_via_smtp": False,
            "message": f"Chyba při odesílání e-mailu přes poštovní server: {e}"
        }


def scheduler_worker() -> None:
    """Vlákno na pozadí, které každou minutu kontroluje plánované odesílání pravidelných odběrů."""
    day_map = {
        0: "MONDAY",
        1: "TUESDAY",
        2: "WEDNESDAY",
        3: "THURSDAY",
        4: "FRIDAY",
        5: "SATURDAY",
        6: "SUNDAY"
    }
    while True:
        try:
            now = datetime.datetime.now()
            current_hhmm = now.strftime("%H:%M")
            today_str = now.strftime("%Y-%m-%d")
            weekday_code = day_map[now.weekday()]
            is_workday = now.weekday() < 5

            subscribers = load_subscribers()
            changed = False

            for sub in subscribers:
                if sub.get("last_sent_date") == today_str:
                    continue
                if sub.get("schedule_time") != current_hhmm:
                    continue

                sched_day = sub.get("schedule_day", "MONDAY")
                should_send = (
                    sched_day == "DAILY"
                    or (sched_day == "WORKDAYS" and is_workday)
                    or (sched_day == weekday_code)
                )

                if should_send:
                    dispatch_email_report(
                        recipient_email=sub["email"],
                        ticker=sub.get("ticker", "NVDA"),
                        leverage=safe_float(sub.get("leverage"), 30.0),
                        growth=safe_float(sub.get("growth"), 5.0),
                        exit_mult=safe_float(sub.get("exit_mult"), 13.5),
                        mode_label="Pravidelný investiční odběr"
                    )
                    sub["last_sent_date"] = today_str
                    changed = True

            if changed:
                save_subscribers(subscribers)
        except Exception:
            pass

        time.sleep(30)


@app.on_event("startup")
def start_background_scheduler() -> None:
    t = threading.Thread(target=scheduler_worker, daemon=True)
    t.start()


# ==============================================================================
# API ENDPOINTY
# ==============================================================================
@app.get("/favicon.ico")
def favicon() -> Response:
    return Response(status_code=204)


@app.get("/api/forex-rates")
def api_forex_rates() -> Dict[str, float]:
    return get_live_forex()


@app.get("/api/valuation/{ticker}")
def api_valuation(
    ticker: str,
    leverage: float = Query(30.0),
    growth: float = Query(5.0),
    exit_mult: float = Query(13.5)
) -> Dict[str, Any]:
    return compute_valuation(ticker, leverage, growth, exit_mult)


@app.get("/api/news/{ticker}")
def api_news(ticker: str) -> Dict[str, Any]:
    return {"ticker": ticker.upper(), "articles": fetch_ticker_news(ticker, limit=6)}


@app.get("/api/advisor/{ticker}")
def api_advisor(
    ticker: str,
    leverage: float = Query(30.0),
    growth: float = Query(5.0),
    exit_mult: float = Query(13.5)
) -> Dict[str, Any]:
    val = compute_valuation(ticker, leverage, growth, exit_mult)
    diff = val["upside_pct"]
    if diff > 15:
        summary = (
            f"Titul {val['ticker']} se obchoduje s výrazným diskontem ({diff:+.1f} %) vůči vnitřní hodnotě "
            f"{val['dcf_price']} {val['currency']}. Bezpečnostní polštář je dostatečný a model podporuje akumulaci pozice."
        )
    elif diff >= -15:
        summary = (
            f"Aktuální tržní cena {val['current_price']} {val['currency']} se pohybuje uvnitř neutrálního valuačního pásma "
            f"({val['lower_band']} – {val['upper_band']} {val['currency']}). Ocenění odpovídá fundamentálnímu výhledu."
        )
    else:
        summary = (
            f"Tržní cena {val['current_price']} {val['currency']} překračuje odhadovanou vnitřní hodnotu "
            f"o {abs(diff):.1f} %. Trh započítává nadstandardní budoucí růst, což zvyšuje valuační riziko při poklesu násobků."
        )
    return {"ticker": val["ticker"], "recommendation": val["band_label"], "analysis": summary}


@app.post("/api/newsletter/send-now")
def api_send_newsletter_now(req: OneOffNewsletterRequest) -> Dict[str, Any]:
    email = req.email.strip()
    if "@" not in email or "." not in email:
        raise HTTPException(status_code=400, detail="Zadejte prosím platnou e-mailovou adresu.")
    result = dispatch_email_report(
        recipient_email=email,
        ticker=req.ticker,
        leverage=req.leverage,
        growth=req.growth,
        exit_mult=req.exit_mult,
        mode_label="Jednorázový investiční přehled"
    )
    return {"status": "ok", **result}


@app.get("/api/subscribers")
def api_list_subscribers() -> Dict[str, Any]:
    return {"subscribers": load_subscribers()}


@app.post("/subscribe")
def api_subscribe_regular(req: RegularSubscriptionRequest) -> Dict[str, Any]:
    email = req.email.strip()
    if "@" not in email or "." not in email:
        raise HTTPException(status_code=400, detail="Zadejte prosím platnou e-mailovou adresu.")

    subscribers = load_subscribers()
    record = {
        "email": email,
        "ticker": req.ticker.upper().strip(),
        "schedule_day": req.schedule_day,
        "schedule_time": req.schedule_time,
        "leverage": req.leverage,
        "growth": req.growth,
        "exit_mult": req.exit_mult,
        "last_sent_date": ""
    }

    # Aktualizace existujícího záznamu pro stejný e-mail a titul, nebo přidání nového
    updated = False
    for idx, existing in enumerate(subscribers):
        if existing.get("email", "").lower() == email.lower() and existing.get("ticker", "").upper() == record["ticker"]:
            subscribers[idx] = record
            updated = True
            break

    if not updated:
        subscribers.append(record)

    save_subscribers(subscribers)
    return {
        "status": "ok",
        "message": f"E-mail {email} byl zaevidován pro pravidelné zasílání přehledu ({record['ticker']}).",
        "subscribers": subscribers
    }


@app.delete("/subscribe")
def api_unsubscribe(email: str = Query(...), ticker: str = Query("")) -> Dict[str, Any]:
    subscribers = load_subscribers()
    filtered = [
        s for s in subscribers
        if not (s.get("email", "").lower() == email.lower() and (not ticker or s.get("ticker", "").upper() == ticker.upper()))
    ]
    save_subscribers(filtered)
    return {"status": "ok", "subscribers": filtered}


@app.get("/api/export/excel/{ticker}")
def api_export_excel(
    ticker: str,
    leverage: float = Query(30.0),
    growth: float = Query(5.0),
    exit_mult: float = Query(13.5)
):
    val = compute_valuation(ticker, leverage, growth, exit_mult)
    if not HAS_OPENPYXL:
        raise HTTPException(status_code=500, detail="Knihovna openpyxl není nainstalována.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Valuace {val['ticker']}"

    ws.append(["TENSEGRITY TERMINAL - VALUAČNÍ MODEL", val["company_name"], val["ticker"]])
    ws.append([])
    ws.append(["Ukazatel", "Hodnota", "Jednotka"])
    ws.append(["Aktuální tržní cena", val["current_price"], val["currency"]])
    ws.append(["Střední vnitřní hodnota", val["dcf_price"], val["currency"]])
    ws.append(["Dolní hranice férového pásma", val["lower_band"], val["currency"]])
    ws.append(["Horní hranice férového pásma", val["upper_band"], val["currency"]])
    ws.append(["Valuační pásmo", val["band_label"], ""])
    ws.append(["Diskontní sazba", val["wacc_pct"], "%"])
    ws.append(["Výstupní násobek", val["exit_multiple"], "x"])
    ws.append(["Vnitřní míra výnosnosti odkupu", val["lbo"]["irr_pct"], "%"])
    ws.append(["Násobek investovaného kapitálu", val["lbo"]["moic"], "x"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Tensegrity_{val['ticker']}.xlsx"}
    )


@app.get("/api/export/pdf/{ticker}")
def api_export_pdf(
    ticker: str,
    leverage: float = Query(30.0),
    growth: float = Query(5.0),
    exit_mult: float = Query(13.5)
):
    val = compute_valuation(ticker, leverage, growth, exit_mult)
    if not HAS_FPDF:
        raise HTTPException(status_code=500, detail="Knihovna fpdf2 není nainstalována.")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, f"TENSEGRITY TERMINAL - {val['ticker']}", ln=True)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"Spolecnost: {val['company_name'].encode('ascii', 'ignore').decode()}", ln=True)
    pdf.cell(0, 8, f"Aktualni cena: {val['current_price']} {val['currency']}", ln=True)
    pdf.cell(0, 8, f"Vnitrni hodnota: {val['dcf_price']} {val['currency']}", ln=True)
    pdf.cell(0, 8, f"Ferove pasmo: {val['lower_band']} - {val['upper_band']} {val['currency']}", ln=True)
    pdf.cell(0, 8, f"Potencial: {val['upside_pct']:+.1f} %", ln=True)

    raw_bytes = bytes(pdf.output())
    return Response(
        content=raw_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename=Tensegrity_{val['ticker']}.pdf"}
    )


# ==============================================================================
# KOMPLETNÍ UŽIVATELSKÉ ROZHRANÍ (HTML / TAILWIND / JAVASCRIPT)
# ==============================================================================
HTML_PAGE = """<!DOCTYPE html>
<html lang="cs" class="dark">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Tensegrity Terminal Intelligence</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600;700&display=swap" rel="stylesheet">
  <style>
    body {
      font-family: 'Inter', sans-serif;
      background-color: #070b14;
      color: #e2e8f0;
    }
    .font-mono-num {
      font-family: 'JetBrains Mono', monospace;
    }
    .glass-card {
      background: linear-gradient(180deg, #0e1525 0%, #0b101d 100%);
      border: 1px solid #1b2538;
      border-radius: 14px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.35);
    }
    .band-active-green {
      border-color: #10b981 !important;
      background: linear-gradient(180deg, rgba(16, 185, 129, 0.14) 0%, rgba(11, 16, 29, 0.95) 100%) !important;
      box-shadow: 0 0 25px rgba(16, 185, 129, 0.18);
    }
    .band-active-amber {
      border-color: #f59e0b !important;
      background: linear-gradient(180deg, rgba(245, 158, 11, 0.14) 0%, rgba(11, 16, 29, 0.95) 100%) !important;
      box-shadow: 0 0 25px rgba(245, 158, 11, 0.18);
    }
    .band-active-red {
      border-color: #ef4444 !important;
      background: linear-gradient(180deg, rgba(239, 68, 68, 0.16) 0%, rgba(11, 16, 29, 0.95) 100%) !important;
      box-shadow: 0 0 25px rgba(239, 68, 68, 0.22);
    }
  </style>
</head>
<body class="min-h-screen p-3 sm:p-6">
  <div class="max-w-7xl mx-auto space-y-5">

    <!-- HORNÍ LIŠTA -->
    <header class="glass-card p-4 flex flex-wrap items-center justify-between gap-4">
      <div>
        <div class="text-sm sm:text-base font-bold tracking-[0.2em] text-white">TENSEGRITY TERMINAL</div>
        <div class="text-[10px] tracking-[0.35em] text-slate-400 uppercase mt-0.5">INTELLIGENCE</div>
      </div>

      <div class="flex flex-wrap items-center gap-3 flex-1 justify-end">
        <input
          id="tickerInput"
          type="text"
          value="NVDA"
          placeholder="Zadejte symbol (např. NVDA, AAPL, CEZ.PR)..."
          class="bg-[#080d18] border border-slate-700/80 focus:border-sky-500 rounded-lg px-3.5 py-2 text-sm font-mono-num font-semibold text-white uppercase outline-none w-36 sm:w-44"
        />

        <!-- KURZY MĚN -->
        <div class="hidden xl:flex items-center gap-4 bg-[#080d18] border border-slate-800 rounded-lg px-4 py-2 text-xs font-mono-num">
          <span class="text-slate-400">EUR/CZK <strong id="fxEur" class="text-white ml-1">--</strong></span>
          <span class="text-slate-400">USD/CZK <strong id="fxUsd" class="text-white ml-1">--</strong></span>
          <span class="text-slate-400">GBP/CZK <strong id="fxGbp" class="text-white ml-1">--</strong></span>
        </div>

        <button onclick="runAnalysis()" class="bg-sky-500 hover:bg-sky-400 text-slate-950 font-semibold text-xs sm:text-sm px-4 py-2 rounded-lg transition cursor-pointer">
          Analyzovat
        </button>
        <button onclick="downloadExport('excel')" class="bg-emerald-950/70 hover:bg-emerald-900/80 border border-emerald-700/60 text-emerald-300 font-medium text-xs sm:text-sm px-3.5 py-2 rounded-lg transition cursor-pointer">
          Excel Model
        </button>
        <button onclick="downloadExport('pdf')" class="bg-rose-950/70 hover:bg-rose-900/80 border border-rose-700/60 text-rose-300 font-medium text-xs sm:text-sm px-3.5 py-2 rounded-lg transition cursor-pointer">
          Přehled ke stažení
        </button>
        <button onclick="openNewsletterModal()" class="bg-slate-800 hover:bg-slate-700 border border-slate-600/70 text-slate-100 font-medium text-xs sm:text-sm px-4 py-2 rounded-lg transition cursor-pointer">
          Newsletter &amp; Tržní zprávy
        </button>
      </div>
    </header>

    <!-- PARAMETRY SCÉNÁŘE -->
    <section class="glass-card p-4 flex flex-wrap items-center gap-6">
      <div class="text-xs font-bold tracking-wider text-sky-400 uppercase">PARAMETRY SCÉNÁŘE</div>

      <div class="flex items-center gap-2 text-xs sm:text-sm text-slate-300">
        <span>Cizí kapitál (% Dluhu)</span>
        <input id="leverageInput" type="number" value="30" min="0" max="90" step="5"
          class="bg-[#080d18] border border-slate-700 rounded-md px-2.5 py-1 w-20 text-right font-mono-num text-white font-semibold" />
        <span class="text-slate-400">%</span>
      </div>

      <div class="flex items-center gap-2 text-xs sm:text-sm text-slate-300">
        <span>Růst provozního zisku</span>
        <input id="growthInput" type="number" value="5" min="-20" max="50" step="1"
          class="bg-[#080d18] border border-slate-700 rounded-md px-2.5 py-1 w-20 text-right font-mono-num text-white font-semibold" />
        <span class="text-slate-400">%</span>
      </div>

      <div class="flex items-center gap-2 text-xs sm:text-sm text-slate-300">
        <span>Výstupní násobek hodnoty</span>
        <input id="exitMultInput" type="number" value="13.5" min="3" max="40" step="0.5"
          class="bg-[#080d18] border border-slate-700 rounded-md px-2.5 py-1 w-20 text-right font-mono-num text-white font-semibold" />
        <span class="text-slate-400">x</span>
      </div>
    </section>

    <!-- INVESTIČNÍ SEMAFOR & TEZE -->
    <section class="glass-card p-5">
      <div class="flex items-center justify-between mb-4">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase">INVESTIČNÍ SEMAFOR &amp; TEZE</div>
        <span id="badgeTicker1" class="text-xs font-mono-num font-semibold text-sky-400 bg-sky-950/70 border border-sky-800/70 px-2.5 py-0.5 rounded">NVDA</span>
      </div>

      <div class="flex flex-col sm:flex-row sm:items-center gap-5">
        <div class="flex items-center gap-2.5 bg-[#080d18] border border-slate-800 rounded-full px-4 py-2.5 w-fit">
          <span id="lightGreen" class="w-4 h-4 rounded-full bg-emerald-950 transition-all"></span>
          <span id="lightAmber" class="w-4 h-4 rounded-full bg-amber-950 transition-all"></span>
          <span id="lightRed" class="w-4 h-4 rounded-full bg-rose-950 transition-all"></span>
        </div>

        <div>
          <h2 id="signalTitle" class="text-base sm:text-lg font-bold text-white">Načítání analýzy...</h2>
          <p id="signalSubtitle" class="text-xs sm:text-sm text-slate-400 mt-1">Výpočet vnitřní hodnoty a diskontní sazby...</p>
        </div>
      </div>
    </section>

    <!-- VALUAČNÍ CENOVÁ PÁSMA AKCIE -->
    <section class="glass-card p-5 space-y-4">
      <div class="flex items-center justify-between">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase">VALUAČNÍ CENOVÁ PÁSMA AKCIE (AKTUÁLNÍ CENA VS. INTERVALY)</div>
        <span id="badgeTicker2" class="text-xs font-mono-num font-semibold text-sky-400 bg-sky-950/70 border border-sky-800/70 px-2.5 py-0.5 rounded">NVDA</span>
      </div>

      <div class="bg-[#080d18] border border-slate-800/90 rounded-xl p-4 grid grid-cols-1 md:grid-cols-3 gap-4 items-center">
        <div>
          <div class="text-[11px] font-semibold tracking-wider text-slate-400 uppercase">AKTUÁLNÍ TRŽNÍ CENA AKCIE</div>
          <div id="bandCurrentPrice" class="text-xl sm:text-2xl font-mono-num font-bold text-white mt-1">--</div>
        </div>

        <div>
          <div class="text-[11px] font-semibold tracking-wider text-slate-400 uppercase">STŘEDNÍ VNITŘNÍ HODNOTA</div>
          <div id="bandIntrinsicPrice" class="text-xl sm:text-2xl font-mono-num font-bold text-sky-400 mt-1">--</div>
        </div>

        <div class="md:text-right">
          <div class="text-[11px] font-semibold tracking-wider text-slate-400 uppercase">AKTUÁLNÍ PÁSMO</div>
          <span id="bandCurrentBadge" class="inline-block mt-1.5 text-xs font-bold px-3 py-1 rounded border border-slate-700 bg-slate-900 text-slate-200">--</span>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
        <!-- ZELENÉ PÁSMO -->
        <div id="cardGreenBand" class="bg-[#080d18] border border-emerald-900/60 rounded-xl p-4 transition-all">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold tracking-wider text-emerald-400 uppercase">ZELENÉ PÁSMO (PODHODNOCENO)</span>
            <span id="tagGreenActive" class="hidden text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 rounded">AKTUÁLNÍ CENA</span>
          </div>
          <div id="rangeGreen" class="text-lg font-mono-num font-bold text-white mt-2">--</div>
          <p class="text-xs text-slate-400 mt-2">Bezpečnostní marže ≥ 15 % pod vnitřní hodnotou modelu (atraktivní valuace).</p>
        </div>

        <!-- ŽLUTOORANŽOVÉ PÁSMO -->
        <div id="cardAmberBand" class="bg-[#080d18] border border-amber-900/60 rounded-xl p-4 transition-all">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold tracking-wider text-amber-400 uppercase">ŽLUTOORANŽOVÉ PÁSMO (FÉROVÁ HODNOTA)</span>
            <span id="tagAmberActive" class="hidden text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 rounded">AKTUÁLNÍ CENA</span>
          </div>
          <div id="rangeAmber" class="text-lg font-mono-num font-bold text-white mt-2">--</div>
          <p class="text-xs text-slate-400 mt-2">Neutrální koridor ± 15 % kolem odhadované reálné hodnoty akcie.</p>
        </div>

        <!-- ČERVENÉ PÁSMO -->
        <div id="cardRedBand" class="bg-[#080d18] border border-rose-900/60 rounded-xl p-4 transition-all">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold tracking-wider text-rose-400 uppercase">ČERVENÉ PÁSMO (NADHODNOCENO)</span>
            <span id="tagRedActive" class="hidden text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 rounded">AKTUÁLNÍ CENA</span>
          </div>
          <div id="rangeRed" class="text-lg font-mono-num font-bold text-white mt-2">--</div>
          <p class="text-xs text-slate-400 mt-2">Tržní cena přesahuje vnitřní hodnotu o více než 15 % (zvýšené valuační riziko).</p>
        </div>
      </div>
    </section>

    <!-- METRIKY A OCENĚNÍ -->
    <section class="grid grid-cols-1 md:grid-cols-3 gap-5">
      <div class="glass-card p-5 space-y-3">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase">TRŽNÍ PROFIL &amp; KAPITALIZACE</div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Tržní kapitalizace</span><strong id="mCap" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Hodnota podniku</span><strong id="mEv" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Čistý dluh</span><strong id="mDebt" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Roční tržby</span><strong id="mRev" class="font-mono-num text-white">--</strong></div>
      </div>

      <div class="glass-card p-5 space-y-3">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase">OCENĚNÍ METODOU SOUČTU ČÁSTÍ</div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Vnitřní hodnota rozkladu</span><strong id="sotpPrice" class="font-mono-num text-sky-400">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Násobek ceny k zisku</span><strong id="compPe" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Násobek hodnoty k zisku</span><strong id="compEvEbitda" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Násobek hodnoty k tržbám</span><strong id="compEvRev" class="font-mono-num text-white">--</strong></div>
      </div>

      <div class="glass-card p-5 space-y-3">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase">PÁKOVÝ ODKUP &amp; NÁVRATNOST</div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Vnitřní míra výnosnosti</span><strong id="lboIrr" class="font-mono-num text-emerald-400">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Násobek kapitálu</span><strong id="lboMoic" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Počáteční dluh transakce</span><strong id="lboDebt" class="font-mono-num text-white">--</strong></div>
        <div class="flex justify-between text-sm"><span class="text-slate-400">Výstupní hodnota podílu</span><strong id="lboExitEq" class="font-mono-num text-white">--</strong></div>
      </div>
    </section>

    <!-- GRAF A AMORTIZAČNÍ KALENDÁŘ -->
    <section class="grid grid-cols-1 lg:grid-cols-2 gap-5">
      <div class="glass-card p-5">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase mb-3">CENOVÝ VÝVOJ (6 MĚSÍCŮ)</div>
        <div class="h-64">
          <canvas id="priceChart"></canvas>
        </div>
      </div>

      <div class="glass-card p-5 overflow-x-auto">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase mb-3">5LETÝ SPLÁTKOVÝ KALENDÁŘ DLUHU (V MIL.)</div>
        <table class="w-full text-left text-xs font-mono-num">
          <thead>
            <tr class="border-b border-slate-800 text-slate-400">
              <th class="py-2">Období</th>
              <th class="py-2">Provozní zisk</th>
              <th class="py-2">Úroky</th>
              <th class="py-2">Splátka jistiny</th>
              <th class="py-2">Zůstatek dluhu</th>
            </tr>
          </thead>
          <tbody id="amortTableBody" class="divide-y divide-slate-800/60"></tbody>
        </table>
      </div>
    </section>

    <!-- CITLIVOSTNÍ MATICE A ANALYTICKÝ KOMENTÁŘ -->
    <section class="grid grid-cols-1 lg:grid-cols-2 gap-5">
      <div class="glass-card p-5 overflow-x-auto">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase mb-3">CITLIVOSTNÍ MATICE VNITŘNÍ HODNOTY (DISKONTNÍ SAZBA VS. VÝSTUPNÍ NÁSOBEK)</div>
        <table class="w-full text-center text-xs font-mono-num" id="sensitivityTable"></table>
      </div>

      <div class="glass-card p-5 flex flex-col justify-between">
        <div>
          <div class="text-xs font-bold tracking-wider text-sky-400 uppercase mb-2">INVESTIČNÍ SHRNUTÍ MODELU</div>
          <p id="advisorText" class="text-sm text-slate-300 leading-relaxed">Načítání analytického komentáře...</p>
        </div>
        <div class="mt-4 pt-3 border-t border-slate-800 text-xs text-slate-500">
          Všechny výpočty vycházejí z živých tržních dat a uživatelských parametrů scénáře.
        </div>
      </div>
    </section>

  </div>

  <!-- MODÁLNÍ OKNO: NEWSLETTER & TRŽNÍ ZPRÁVY -->
  <div id="newsletterModal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/75 backdrop-blur-sm p-4">
    <div class="bg-[#0e1525] border border-slate-700/80 rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-5 max-h-[90vh] overflow-y-auto">
      <div class="flex items-center justify-between border-b border-slate-800 pb-3">
        <h3 class="text-sm font-bold tracking-wider text-white uppercase">
          NEWSLETTER &amp; TRŽNÍ ZPRÁVY (<span id="modalTickerLabel" class="text-sky-400">NVDA</span>)
        </h3>
        <button onclick="closeNewsletterModal()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg px-2.5 py-1 text-sm cursor-pointer">
          ✕
        </button>
      </div>

      <!-- ZADÁNÍ E-MAILU A VOLBA REŽIMU ODESLÁNÍ -->
      <div class="space-y-4">
        <div>
          <label class="block text-xs font-semibold text-slate-400 mb-1.5">E-mailová adresa příjemce</label>
          <input
            id="subscribeEmailInput"
            type="email"
            placeholder="Zadejte e-mail pro zaslání investičního bulletinu..."
            class="w-full bg-[#080d18] border border-slate-700 focus:border-sky-500 rounded-xl px-4 py-2.5 text-sm text-white outline-none"
          />
        </div>

        <!-- NASTAVENÍ PRAVIDELNÉHO HARMONOGRAMU -->
        <div class="bg-[#080d18] border border-slate-800 rounded-xl p-3.5 grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block text-[11px] font-semibold text-slate-400 uppercase mb-1">Dny pravidelného odesílání</label>
            <select id="scheduleDaySelect" class="w-full bg-[#0e1525] border border-slate-700 rounded-lg px-3 py-2 text-xs sm:text-sm text-white outline-none">
              <option value="WORKDAYS">Každý pracovní den (Po–Pá)</option>
              <option value="DAILY">Denně (Po–Ne)</option>
              <option value="MONDAY" selected>Každé pondělí</option>
              <option value="WEDNESDAY">Každou středu</option>
              <option value="FRIDAY">Každý pátek</option>
            </select>
          </div>

          <div>
            <label class="block text-[11px] font-semibold text-slate-400 uppercase mb-1">Čas odeslání</label>
            <input
              id="scheduleTimeInput"
              type="time"
              value="08:00"
              class="w-full bg-[#0e1525] border border-slate-700 rounded-lg px-3 py-2 text-xs sm:text-sm font-mono-num text-white outline-none"
            />
          </div>
        </div>

        <!-- DVĚ SAMOSTATNÁ TLAČÍTKA PRO JEDNORÁZOVÉ ODESLÁNÍ A PRAVIDELNOU EVIDENCI -->
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <button
            id="btnOneOffSend"
            onclick="sendOneOffNewsletter()"
            class="w-full bg-sky-500 hover:bg-sky-400 text-slate-950 font-semibold text-xs sm:text-sm py-2.5 px-4 rounded-xl transition cursor-pointer"
          >
            Odeslat jednorázově nyní
          </button>

          <button
            id="btnRegularSubscribe"
            onclick="subscribeRegularNewsletter()"
            class="w-full bg-emerald-500/20 hover:bg-emerald-500/30 border border-emerald-500/60 text-emerald-300 font-semibold text-xs sm:text-sm py-2.5 px-4 rounded-xl transition cursor-pointer"
          >
            Evidovat pro pravidelný odběr
          </button>
        </div>

        <!-- STAVOVÁ ZPRÁVA -->
        <div id="newsletterStatusBox" class="hidden text-xs rounded-xl p-3 border"></div>
      </div>

      <!-- SEZNAM EVIDOVANÝCH PRAVIDELNÝCH ODBĚRATELŮ -->
      <div class="border-t border-slate-800 pt-4">
        <div class="flex items-center justify-between mb-2">
          <span class="text-xs font-bold tracking-wider text-slate-400 uppercase">EVIDOVANÉ E-MAILY PRO PRAVIDELNÉ ZASÍLÁNÍ</span>
          <span id="subscribersCountBadge" class="text-[11px] font-mono-num text-sky-400">0 odběratelů</span>
        </div>
        <div id="subscribersListContainer" class="space-y-2 max-h-36 overflow-y-auto text-xs">
          <div class="text-slate-500 italic">Žádné e-maily zatím nejsou evidovány.</div>
        </div>
      </div>

      <!-- AKTUÁLNÍ TRŽNÍ ZPRÁVY PRO ZVOLENÝ TITUL -->
      <div class="border-t border-slate-800 pt-4">
        <div class="text-xs font-bold tracking-wider text-slate-400 uppercase mb-2.5">AKTUÁLNÍ TRŽNÍ ZPRÁVY</div>
        <div id="newsFeedContainer" class="space-y-2.5">
          <div class="text-xs text-slate-400">Načítání tržních zpráv...</div>
        </div>
      </div>
    </div>
  </div>

  <script>
    let priceChartInstance = null;

    const DAY_LABELS = {
      "WORKDAYS": "Pracovní dny (Po–Pá)",
      "DAILY": "Denně (Po–Ne)",
      "MONDAY": "Každé pondělí",
      "TUESDAY": "Každé úterý",
      "WEDNESDAY": "Každou středu",
      "THURSDAY": "Každý čtvrtek",
      "FRIDAY": "Každý pátek"
    };

    function fmtNum(num, decimals = 2) {
      return Number(num).toLocaleString('cs-CZ', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
      });
    }

    async function loadForexRates() {
      try {
        const res = await fetch('/api/forex-rates');
        const data = await res.json();
        document.getElementById('fxEur').textContent = fmtNum(data.EURCZK, 2);
        document.getElementById('fxUsd').textContent = fmtNum(data.USDCZK, 2);
        document.getElementById('fxGbp').textContent = fmtNum(data.GBPCZK, 2);
      } catch (e) {
        console.warn('Chyba načítání kurzů:', e);
      }
    }

    async function runAnalysis() {
      const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
      const leverage = document.getElementById('leverageInput').value || 30;
      const growth = document.getElementById('growthInput').value || 5;
      const exitMult = document.getElementById('exitMultInput').value || 13.5;

      document.getElementById('badgeTicker1').textContent = ticker;
      document.getElementById('badgeTicker2').textContent = ticker;
      document.getElementById('modalTickerLabel').textContent = ticker;

      try {
        const query = `?leverage=${leverage}&growth=${growth}&exit_mult=${exitMult}`;
        const [valRes, advRes] = await Promise.all([
          fetch(`/api/valuation/${encodeURIComponent(ticker)}${query}`),
          fetch(`/api/advisor/${encodeURIComponent(ticker)}${query}`)
        ]);

        if (!valRes.ok) throw new Error('Nepodařilo se načíst valuační data.');
        const val = await valRes.json();
        const adv = advRes.ok ? await advRes.json() : { analysis: '' };

        updateDashboardUI(val, adv.analysis);
      } catch (err) {
        document.getElementById('signalTitle').textContent = 'Chyba při načítání dat pro ' + ticker;
        document.getElementById('signalSubtitle').textContent = err.message;
      }
    }

    function updateDashboardUI(val, advisorSummary) {
      const curr = val.currency;

      // Semafor
      const lg = document.getElementById('lightGreen');
      const la = document.getElementById('lightAmber');
      const lr = document.getElementById('lightRed');
      lg.className = 'w-4 h-4 rounded-full bg-emerald-950 transition-all';
      la.className = 'w-4 h-4 rounded-full bg-amber-950 transition-all';
      lr.className = 'w-4 h-4 rounded-full bg-rose-950 transition-all';

      if (val.signal_color === 'green') {
        lg.className = 'w-4 h-4 rounded-full bg-emerald-400 shadow-[0_0_14px_#34d399]';
      } else if (val.signal_color === 'amber') {
        la.className = 'w-4 h-4 rounded-full bg-amber-400 shadow-[0_0_14px_#fbbf24]';
      } else {
        lr.className = 'w-4 h-4 rounded-full bg-rose-500 shadow-[0_0_14px_#f43f5e]';
      }

      document.getElementById('signalTitle').textContent = val.signal_title;
      document.getElementById('signalSubtitle').textContent =
        `Vnitřní hodnota modelu ${fmtNum(val.dcf_price)} ${curr} proti aktuální tržní ceně ${fmtNum(val.current_price)} ${curr}. ` +
        `Výsledek závisí na diskontní sazbě ${fmtNum(val.wacc_pct)} % a výstupním násobku ${val.exit_multiple}x.`;

      // Cenová pásma
      document.getElementById('bandCurrentPrice').textContent = `${fmtNum(val.current_price)} ${curr}`;
      document.getElementById('bandIntrinsicPrice').textContent = `${fmtNum(val.dcf_price)} ${curr}`;
      document.getElementById('rangeGreen').textContent = `0,00 – ${fmtNum(val.lower_band)} ${curr}`;
      document.getElementById('rangeAmber').textContent = `${fmtNum(val.lower_band)} – ${fmtNum(val.upper_band)} ${curr}`;
      document.getElementById('rangeRed').textContent = `nad ${fmtNum(val.upper_band)} ${curr}`;

      const badge = document.getElementById('bandCurrentBadge');
      badge.textContent = val.band_label;

      const cGreen = document.getElementById('cardGreenBand');
      const cAmber = document.getElementById('cardAmberBand');
      const cRed = document.getElementById('cardRedBand');
      cGreen.classList.remove('band-active-green');
      cAmber.classList.remove('band-active-amber');
      cRed.classList.remove('band-active-red');
      document.getElementById('tagGreenActive').classList.add('hidden');
      document.getElementById('tagAmberActive').classList.add('hidden');
      document.getElementById('tagRedActive').classList.add('hidden');

      if (val.signal_color === 'green') {
        cGreen.classList.add('band-active-green');
        document.getElementById('tagGreenActive').classList.remove('hidden');
        badge.className = 'inline-block mt-1.5 text-xs font-bold px-3 py-1 rounded border border-emerald-700 bg-emerald-950/70 text-emerald-300';
      } else if (val.signal_color === 'amber') {
        cAmber.classList.add('band-active-amber');
        document.getElementById('tagAmberActive').classList.remove('hidden');
        badge.className = 'inline-block mt-1.5 text-xs font-bold px-3 py-1 rounded border border-amber-700 bg-amber-950/70 text-amber-300';
      } else {
        cRed.classList.add('band-active-red');
        document.getElementById('tagRedActive').classList.remove('hidden');
        badge.className = 'inline-block mt-1.5 text-xs font-bold px-3 py-1 rounded border border-rose-700 bg-rose-950/70 text-rose-300';
      }

      // Metriky
      document.getElementById('mCap').textContent = `${fmtNum(val.market_cap_m, 0)} mil. ${curr}`;
      document.getElementById('mEv').textContent = `${fmtNum(val.enterprise_value_m, 0)} mil. ${curr}`;
      document.getElementById('mDebt').textContent = `${fmtNum(val.net_debt_m, 0)} mil. ${curr}`;
      document.getElementById('mRev').textContent = `${fmtNum(val.revenue_m, 0)} mil. ${curr}`;

      document.getElementById('sotpPrice').textContent = `${fmtNum(val.sotp_price)} ${curr}`;
      document.getElementById('compPe').textContent = `${fmtNum(val.comps.pe_ratio, 1)}x`;
      document.getElementById('compEvEbitda').textContent = `${fmtNum(val.comps.ev_ebitda, 1)}x`;
      document.getElementById('compEvRev').textContent = `${fmtNum(val.comps.ev_revenue, 2)}x`;

      document.getElementById('lboIrr').textContent = `${fmtNum(val.lbo.irr_pct, 1)} %`;
      document.getElementById('lboMoic').textContent = `${fmtNum(val.lbo.moic, 2)}x`;
      document.getElementById('lboDebt').textContent = `${fmtNum(val.lbo.initial_debt_m, 0)} mil. ${curr}`;
      document.getElementById('lboExitEq').textContent = `${fmtNum(val.lbo.exit_equity_m, 0)} mil. ${curr}`;

      // Splátkový kalendář
      const tbody = document.getElementById('amortTableBody');
      tbody.innerHTML = val.lbo.schedule.map(r => `
        <tr>
          <td class="py-2 text-slate-300">${r.year}</td>
          <td class="py-2 text-white">${fmtNum(r.ebitda, 1)}</td>
          <td class="py-2 text-rose-400">${fmtNum(r.interest, 1)}</td>
          <td class="py-2 text-emerald-400">${fmtNum(r.repayment, 1)}</td>
          <td class="py-2 text-sky-400">${fmtNum(r.ending_debt, 1)}</td>
        </tr>
      `).join('');

      // Citlivostní matice
      const sTable = document.getElementById('sensitivityTable');
      let headerHtml = `<thead><tr class="border-b border-slate-800 text-slate-400"><th class="py-2">Sazba \\ Násobek</th>`;
      val.sensitivity.multiples.forEach(m => {
        headerHtml += `<th class="py-2">${fmtNum(m, 1)}x</th>`;
      });
      headerHtml += `</tr></thead><tbody class="divide-y divide-slate-800/60">`;

      val.sensitivity.rows.forEach(row => {
        headerHtml += `<tr><td class="py-2 font-bold text-sky-400">${fmtNum(row.wacc_pct, 2)} %</td>`;
        row.prices.forEach(p => {
          const isHigher = p >= val.current_price;
          headerHtml += `<td class="py-2 ${isHigher ? 'text-emerald-400' : 'text-slate-300'}">${fmtNum(p, 1)}</td>`;
        });
        headerHtml += `</tr>`;
      });
      headerHtml += `</tbody>`;
      sTable.innerHTML = headerHtml;

      document.getElementById('advisorText').textContent = advisorSummary;

      // Graf
      renderPriceChart(val.history.labels, val.history.prices, val.ticker, curr);
    }

    function renderPriceChart(labels, prices, ticker, curr) {
      const ctx = document.getElementById('priceChart').getContext('2d');
      if (priceChartInstance) {
        priceChartInstance.destroy();
      }
      priceChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
          labels: labels,
          datasets: [{
            label: `${ticker} (${curr})`,
            data: prices,
            borderColor: '#38bdf8',
            backgroundColor: 'rgba(56, 189, 248, 0.12)',
            borderWidth: 2,
            fill: true,
            tension: 0.3,
            pointRadius: 2
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8', font: { size: 11 } } },
            y: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8', font: { size: 11 } } }
          }
        }
      });
    }

    function downloadExport(type) {
      const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
      const leverage = document.getElementById('leverageInput').value || 30;
      const growth = document.getElementById('growthInput').value || 5;
      const exitMult = document.getElementById('exitMultInput').value || 13.5;
      window.open(`/api/export/${type}/${encodeURIComponent(ticker)}?leverage=${leverage}&growth=${growth}&exit_mult=${exitMult}`, '_blank');
    }

    // ==========================================================================
    // NEWSLETTER A PRAVIDELNÝ ODBĚR
    // ==========================================================================
    function showNewsletterStatus(message, isSuccess = true) {
      const box = document.getElementById('newsletterStatusBox');
      box.textContent = message;
      box.classList.remove('hidden', 'border-emerald-700', 'bg-emerald-950/60', 'text-emerald-300', 'border-rose-700', 'bg-rose-950/60', 'text-rose-300');
      if (isSuccess) {
        box.classList.add('border-emerald-700', 'bg-emerald-950/60', 'text-emerald-300');
      } else {
        box.classList.add('border-rose-700', 'bg-rose-950/60', 'text-rose-300');
      }
    }

    async function openNewsletterModal() {
      const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
      document.getElementById('modalTickerLabel').textContent = ticker;
      const modal = document.getElementById('newsletterModal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
      document.getElementById('newsletterStatusBox').classList.add('hidden');

      await Promise.all([loadSubscribersList(), loadTickerNews(ticker)]);
    }

    function closeNewsletterModal() {
      const modal = document.getElementById('newsletterModal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
    }

    async function loadTickerNews(ticker) {
      const container = document.getElementById('newsFeedContainer');
      container.innerHTML = '<div class="text-xs text-slate-400">Načítání aktuálních zpráv...</div>';
      try {
        const res = await fetch(`/api/news/${encodeURIComponent(ticker)}`);
        const data = await res.json();
        const articles = data.articles || [];
        if (!articles.length) {
          container.innerHTML = '<div class="text-xs text-slate-500">Žádné aktuální zprávy nebyly nalezeny.</div>';
          return;
        }
        container.innerHTML = articles.map(a => `
          <a href="${a.link}" target="_blank" rel="noopener noreferrer"
             class="block bg-[#080d18] hover:bg-slate-900/90 border border-slate-800 rounded-xl p-3 transition">
            <div class="text-xs font-semibold text-slate-100">${a.title}</div>
            <div class="text-[11px] text-slate-400 mt-1">${a.publisher} • ${a.published}</div>
          </a>
        `).join('');
      } catch (e) {
        container.innerHTML = '<div class="text-xs text-rose-400">Nepodařilo se načíst tržní zprávy.</div>';
      }
    }

    async function loadSubscribersList() {
      const container = document.getElementById('subscribersListContainer');
      const badge = document.getElementById('subscribersCountBadge');
      try {
        const res = await fetch('/api/subscribers');
        const data = await res.json();
        const list = data.subscribers || [];
        badge.textContent = `${list.length} odběratelů`;
        if (!list.length) {
          container.innerHTML = '<div class="text-slate-500 italic">Žádné e-maily zatím nejsou evidovány.</div>';
          return;
        }
        container.innerHTML = list.map(item => {
          const dayText = DAY_LABELS[item.schedule_day] || item.schedule_day;
          return `
            <div class="flex items-center justify-between bg-[#080d18] border border-slate-800 rounded-lg px-3 py-2">
              <div>
                <div class="font-semibold text-slate-200">${item.email} <span class="text-sky-400 font-mono-num ml-1">[${item.ticker}]</span></div>
                <div class="text-[11px] text-slate-400">${dayText} v ${item.schedule_time}</div>
              </div>
              <button onclick="removeSubscriber('${item.email}', '${item.ticker}')"
                class="text-rose-400 hover:text-rose-300 text-xs font-medium px-2 py-1 rounded bg-rose-950/40 border border-rose-900/60 cursor-pointer">
                Odebrat
              </button>
            </div>
          `;
        }).join('');
      } catch (e) {
        container.innerHTML = '<div class="text-rose-400">Chyba načtení seznamu odběratelů.</div>';
      }
    }

    async function sendOneOffNewsletter() {
      const email = (document.getElementById('subscribeEmailInput').value || '').trim();
      if (!email || !email.includes('@')) {
        showNewsletterStatus('Zadejte prosím platnou e-mailovou adresu.', false);
        return;
      }

      const btn = document.getElementById('btnOneOffSend');
      const origText = btn.textContent;
      btn.disabled = true;
      btn.textContent = 'Odesílám...';

      try {
        const res = await fetch('/api/newsletter/send-now', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: email,
            ticker: (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase(),
            leverage: Number(document.getElementById('leverageInput').value || 30),
            growth: Number(document.getElementById('growthInput').value || 5),
            exit_mult: Number(document.getElementById('exitMultInput').value || 13.5)
          })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Odeslání selhalo.');
        showNewsletterStatus(data.message, data.sent_via_smtp !== false);
      } catch (err) {
        showNewsletterStatus(err.message, false);
      } finally {
        btn.disabled = false;
        btn.textContent = origText;
      }
    }

    async function subscribeRegularNewsletter() {
      const email = (document.getElementById('subscribeEmailInput').value || '').trim();
      if (!email || !email.includes('@')) {
        showNewsletterStatus('Zadejte prosím platnou e-mailovou adresu.', false);
        return;
      }

      const btn = document.getElementById('btnRegularSubscribe');
      const origText = btn.textContent;
      btn.disabled = true;
      btn.textContent = 'Ukládám...';

      try {
        const res = await fetch('/subscribe', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: email,
            ticker: (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase(),
            schedule_day: document.getElementById('scheduleDaySelect').value,
            schedule_time: document.getElementById('scheduleTimeInput').value || '08:00',
            leverage: Number(document.getElementById('leverageInput').value || 30),
            growth: Number(document.getElementById('growthInput').value || 5),
            exit_mult: Number(document.getElementById('exitMultInput').value || 13.5)
          })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Registrace odběru selhala.');
        showNewsletterStatus(data.message, true);
        await loadSubscribersList();
      } catch (err) {
        showNewsletterStatus(err.message, false);
      } finally {
        btn.disabled = false;
        btn.textContent = origText;
      }
    }

    async function removeSubscriber(email, ticker) {
      try {
        await fetch(`/subscribe?email=${encodeURIComponent(email)}&ticker=${encodeURIComponent(ticker)}`, {
          method: 'DELETE'
        });
        await loadSubscribersList();
      } catch (e) {
        console.warn('Chyba při odebírání odběratele:', e);
      }
    }

    document.getElementById('tickerInput').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') runAnalysis();
    });

    window.addEventListener('DOMContentLoaded', () => {
      loadForexRates();
      runAnalysis();
    });
  </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def serve_terminal() -> HTMLResponse:
    return HTMLResponse(content=HTML_PAGE)