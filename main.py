# -*- coding: utf-8 -*-
from contextlib import asynccontextmanager
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr
from html import escape
import asyncio
import io
import json
import logging
import os
import re
import smtplib
import ssl
from typing import Optional
from zoneinfo import ZoneInfo

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
import yfinance as yf

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

PRAGUE_TZ = ZoneInfo("Europe/Prague")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
SUBSCRIBERS_FILE = os.path.join(BASE_DIR, "subscribers.json")

SENDER_EMAIL = os.getenv("SENDER_EMAIL", os.getenv("SMTP_USER", "kubenkamatyas@gmail.com"))
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", os.getenv("SMTP_PASSWORD", ""))

FOREX_PAIRS = {
    "EUR/CZK": "EURCZK=X",
    "USD/CZK": "USDCZK=X",
    "GBP/CZK": "GBPCZK=X",
}
FOREX_FALLBACK_RATES = {"EUR/CZK": 25.0, "USD/CZK": 23.0, "GBP/CZK": 29.0}

# Pevný rozvrh pravidelného newsletteru: Středa 7:00, Pátek 7:00, Neděle 19:00 (Europe/Prague)
FIXED_SCHEDULE_SLOTS = [
    (2, 7, 0),   # Středa 07:00
    (4, 7, 0),   # Pátek 07:00
    (6, 19, 0),  # Neděle 19:00
]
_last_dispatched_slot: Optional[str] = None


def get_smtp_config():
    username = (os.getenv("SENDER_EMAIL") or os.getenv("SMTP_USER") or SENDER_EMAIL).strip()
    password = (os.getenv("SENDER_PASSWORD") or os.getenv("SMTP_PASSWORD") or SENDER_PASSWORD).replace(" ", "")
    if not username or not password:
        return None
    return {"host": "smtp.gmail.com", "port": 465, "username": username, "password": password}


def is_valid_email_address(email: str) -> bool:
    if not isinstance(email, str) or len(email) > 254:
        return False
    parsed = parseaddr(email)
    return parsed[1] == email and bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email))


def get_subscribers() -> list[dict]:
    if not os.path.exists(SUBSCRIBERS_FILE):
        return []
    try:
        with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, list):
            return []
        normalized = []
        for item in data:
            if isinstance(item, str) and item.strip():
                normalized.append({
                    "email": item.strip(),
                    "ticker": "NVDA",
                    "created_at": datetime.now(PRAGUE_TZ).strftime("%d.%m.%Y %H:%M"),
                })
            elif isinstance(item, dict) and item.get("email"):
                normalized.append({
                    "email": str(item["email"]).strip(),
                    "ticker": str(item.get("ticker") or "NVDA").strip().upper(),
                    "created_at": str(item.get("created_at") or datetime.now(PRAGUE_TZ).strftime("%d.%m.%Y %H:%M")),
                })
        return normalized
    except (OSError, json.JSONDecodeError):
        return []


def save_subscribers_list(subscribers: list[dict]) -> None:
    temporary_path = SUBSCRIBERS_FILE + ".tmp"
    with open(temporary_path, "w", encoding="utf-8") as file:
        json.dump(subscribers, file, ensure_ascii=False, indent=2)
        file.flush()
        os.fsync(file.fileno())
    os.replace(temporary_path, SUBSCRIBERS_FILE)


def upsert_subscriber(email: str, ticker: str = "NVDA") -> bool:
    normalized_email = email.strip()
    normalized_ticker = (ticker or "NVDA").strip().upper()
    subscribers = get_subscribers()
    for entry in subscribers:
        if entry["email"].casefold() == normalized_email.casefold():
            entry["ticker"] = normalized_ticker
            save_subscribers_list(subscribers)
            return False
    subscribers.append({
        "email": normalized_email,
        "ticker": normalized_ticker,
        "created_at": datetime.now(PRAGUE_TZ).strftime("%d.%m.%Y %H:%M"),
    })
    save_subscribers_list(subscribers)
    return True


def remove_subscriber(email: str) -> bool:
    normalized_email = email.strip().casefold()
    subscribers = get_subscribers()
    filtered = [s for s in subscribers if s["email"].casefold() != normalized_email]
    if len(filtered) == len(subscribers):
        return False
    save_subscribers_list(filtered)
    return True


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
            logger.warning("Nelze načíst živý kurz %s: %s", pair, error)
        rates[pair] = {"rate": rate, "live": live}
    return {
        "rates": rates,
        "updated_at": datetime.now(PRAGUE_TZ).isoformat(timespec="minutes"),
    }


def fetch_ticker_news(ticker: str, limit: int = 6) -> list[dict]:
    try:
        raw_news = yf.Ticker(ticker).news or []
        articles = []
        for item in raw_news[:limit]:
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
        return articles
    except Exception as error:
        logger.warning("Chyba při načítání zpráv pro %s: %s", ticker, error)
        return []


def run_valuation_logic(
    ticker: str,
    leverage: Optional[float] = None,
    growth: Optional[float] = None,
    exit_mult: Optional[float] = None,
):
    ticker_clean = ticker.strip().upper()
    try:
        stock = yf.Ticker(ticker_clean)
        info = stock.info or {}
        financials = stock.financials
        cashflow_df = stock.cashflow
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Nepodařilo se načíst tržní data pro {ticker_clean}: {exc}")

    current_price = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
    shares_outstanding = info.get("sharesOutstanding")
    celkovy_dluh = float(info.get("totalDebt") or 0.0)
    hotovost = float(info.get("totalCash") or 0.0)
    ebitda = float(info.get("ebitda") or 0.0)
    trzby = float(info.get("totalRevenue") or 0.0)
    cisty_zisk = float(info.get("netIncomeToCommon") or info.get("netIncome") or 0.0)
    market_cap = float(info.get("marketCap") or 0.0)
    currency = str(info.get("currency") or "USD")
    company_name = str(info.get("longName") or info.get("shortName") or ticker_clean)

    if not ebitda and financials is not None and not financials.empty:
        try:
            ebitda = float(financials.loc["EBITDA"].iloc[0]) if "EBITDA" in financials.index else trzby * 0.25
        except Exception:
            ebitda = trzby * 0.25

    if not market_cap and current_price and shares_outstanding:
        market_cap = float(current_price) * float(shares_outstanding)

    trzni_kapitalizace = market_cap if market_cap > 0 else 100_000_000_000.0
    celkova_hodnota_podniku = max(1.0, trzni_kapitalizace + celkovy_dluh - hotovost)

    # Skutečný volný peněžní tok (Free Cash Flow) z výkazů společnosti
    base_fcf = float(info.get("freeCashflow") or 0.0)
    if base_fcf <= 0 and cashflow_df is not None and not cashflow_df.empty:
        try:
            op_cf = float(cashflow_df.loc["Operating Cash Flow"].iloc[0]) if "Operating Cash Flow" in cashflow_df.index else 0.0
            capex = abs(float(cashflow_df.loc["Capital Expenditure"].iloc[0])) if "Capital Expenditure" in cashflow_df.index else 0.0
            if op_cf > 0:
                base_fcf = max(op_cf - capex, op_cf * 0.65)
        except Exception:
            pass
    if base_fcf <= 0:
        base_fcf = max(ebitda * 0.68, cisty_zisk * 0.85, trzni_kapitalizace * 0.035)

    # Dynamická kalibrace očekávaného růstu podle fundamentů společnosti
    raw_earnings_growth = info.get("earningsGrowth")
    raw_revenue_growth = info.get("revenueGrowth")
    detected_growth = 0.08
    growth_candidates = [g for g in (raw_earnings_growth, raw_revenue_growth) if isinstance(g, (int, float)) and g > -0.5]
    if growth_candidates:
        avg_g = sum(growth_candidates) / len(growth_candidates)
        # Konzervativní 5letý průměrný roční růst odvozený z aktuálního tempa růstu firmy (6 % až 28 %)
        detected_growth = max(0.06, min(0.28, avg_g * 0.65))

    applied_growth = float(growth) if growth is not None else round(detected_growth, 4)

    # Dynamická kalibrace výstupního násobku podle sektoru a tržního ocenění podniku
    market_ev_ebitda = info.get("enterpriseToEbitda")
    if isinstance(market_ev_ebitda, (int, float)) and 4.0 <= market_ev_ebitda <= 65.0:
        # Výstupní násobek v roce 5 konverguje směrem k udržitelnému sektorovému násobku
        detected_exit_mult = round(max(10.0, min(32.0, float(market_ev_ebitda) * 0.72)), 1)
    else:
        detected_exit_mult = 15.0

    applied_exit_mult = float(exit_mult) if exit_mult is not None else detected_exit_mult

    # Váha cizího kapitálu pro model odkupu za cizí kapitál (výchozí 30 %)
    applied_leverage = float(leverage) if leverage is not None else 0.30

    efektivni_danova_sazba = 0.21
    implikovana_sazba_dluhu = 0.048

    # Vážené průměrné náklady kapitálu (WACC) z tržních vah vlastního a cizího kapitálu
    try:
        beta = float(info.get("beta") or 1.05)
    except (TypeError, ValueError):
        beta = 1.05
    beta = max(0.65, min(1.65, beta))
    bezrizikova_sazba = 0.042
    rizikova_premie_trhu = 0.046
    naklady_vlastniho_kapitalu = bezrizikova_sazba + beta * rizikova_premie_trhu

    trzni_kapital_celkem = trzni_kapitalizace + celkovy_dluh
    vaha_dluhu_trzni = celkovy_dluh / trzni_kapital_celkem if trzni_kapital_celkem > 0 else 0.15
    vaha_vlastniho_trzni = 1.0 - vaha_dluhu_trzni

    wacc = (
        vaha_dluhu_trzni * implikovana_sazba_dluhu * (1.0 - efektivni_danova_sazba)
        + vaha_vlastniho_trzni * naklady_vlastniho_kapitalu
    )
    wacc = max(0.065, min(0.135, wacc))

    # 5letá projekce volných peněžních toků s postupným zpomalováním růstu (fade-down) k dlouhodobé rovnováze
    terminalni_rust = 0.032
    dcf_cash_flows = []
    projected_fcf = base_fcf
    for rok in range(1, 6):
        # Postupný přechod z počátečního tempa růstu směrem k terminálnímu růstu
        fade_weight = (rok - 1) / 5.0
        rocni_rust = applied_growth * (1.0 - 0.35 * fade_weight) + terminalni_rust * (0.35 * fade_weight)
        projected_fcf *= (1.0 + rocni_rust)
        diskontni_faktor = 1.0 / ((1.0 + wacc) ** rok)
        dcf_cash_flows.append({
            "rok": rok,
            "volny_penezni_tok": round(projected_fcf, 2),
            "diskontni_faktor": round(diskontni_faktor, 6),
            "soucasna_hodnota_fcf": round(projected_fcf * diskontni_faktor, 2),
        })

    # Terminální hodnota (kombinace Gordonova růstového modelu a výstupního násobku volného peněžního toku)
    gordon_tv = (projected_fcf * (1.0 + terminalni_rust)) / max(0.025, (wacc - terminalni_rust))
    multiple_tv = projected_fcf * max(12.0, applied_exit_mult * 1.15)
    terminal_value = 0.55 * gordon_tv + 0.45 * multiple_tv
    present_value_terminal_value = terminal_value * dcf_cash_flows[-1]["diskontni_faktor"]

    dcf_enterprise_value = sum(item["soucasna_hodnota_fcf"] for item in dcf_cash_flows) + present_value_terminal_value
    dcf_equity_value = max(0.0, dcf_enterprise_value - celkovy_dluh + hotovost)

    shares = float(shares_outstanding) if shares_outstanding else (trzni_kapitalizace / float(current_price) if current_price else 1_000_000_000.0)
    dcf_share_price = dcf_equity_value / shares if shares > 0 else 0.0

    # Srovnání konkurence (Trading Comps)
    sector = str(info.get("sector") or "Technology")
    forward_pe = info.get("forwardPE")
    trailing_pe = info.get("trailingPE")
    pe_multiple = float(forward_pe) if isinstance(forward_pe, (int, float)) and 5.0 <= forward_pe <= 80.0 else (
        float(trailing_pe) if isinstance(trailing_pe, (int, float)) and 5.0 <= trailing_pe <= 80.0 else 24.0
    )
    ev_ebitda_comp = float(market_ev_ebitda) if isinstance(market_ev_ebitda, (int, float)) and 5.0 <= market_ev_ebitda <= 60.0 else applied_exit_mult
    ev_rev_raw = info.get("enterpriseToRevenue")
    ev_rev_comp = float(ev_rev_raw) if isinstance(ev_rev_raw, (int, float)) and 0.5 <= ev_rev_raw <= 35.0 else 4.5

    forward_ebitda_y1 = (ebitda if ebitda > 0 else base_fcf * 1.35) * (1.0 + applied_growth)
    forward_net_income_y1 = (cisty_zisk if cisty_zisk > 0 else base_fcf * 0.95) * (1.0 + applied_growth)
    forward_rev_y1 = (trzby if trzby > 0 else forward_ebitda_y1 * 3.2) * (1.0 + applied_growth)

    comps_ev_ebitda_eq = max(0.0, forward_ebitda_y1 * ev_ebitda_comp - celkovy_dluh + hotovost)
    comps_pe_eq = max(0.0, forward_net_income_y1 * pe_multiple)
    comps_rev_eq = max(0.0, forward_rev_y1 * ev_rev_comp - celkovy_dluh + hotovost)

    peer_comps_rows = [
        {
            "method": "Poměr hodnoty podniku k provoznímu zisku",
            "multiple": round(ev_ebitda_comp, 1),
            "forward_metric": round(forward_ebitda_y1, 2),
            "implied_share_price": round(comps_ev_ebitda_eq / shares, 2) if shares > 0 else None,
        },
        {
            "method": "Poměr ceny akcie k čistému zisku",
            "multiple": round(pe_multiple, 1),
            "forward_metric": round(forward_net_income_y1, 2),
            "implied_share_price": round(comps_pe_eq / shares, 2) if shares > 0 else None,
        },
        {
            "method": "Poměr hodnoty podniku k tržbám",
            "multiple": round(ev_rev_comp, 1),
            "forward_metric": round(forward_rev_y1, 2),
            "implied_share_price": round(comps_rev_eq / shares, 2) if shares > 0 else None,
        },
    ]

    valid_comp_prices = [r["implied_share_price"] for r in peer_comps_rows if r["implied_share_price"] and r["implied_share_price"] > 0]
    comps_avg_price = (sum(valid_comp_prices) / len(valid_comp_prices)) if valid_comp_prices else dcf_share_price

    # Institucionální střední vnitřní hodnota (70 % diskontované peněžní toky + 30 % srovnání násobků)
    intrinsic_value_per_share = round(0.70 * dcf_share_price + 0.30 * comps_avg_price, 2)

    # Metoda součtu částí (SOTP) - rozdělení hlavního segmentu a růstových divizí
    sotp_ev = (forward_ebitda_y1 * 0.75 * applied_exit_mult) + (forward_ebitda_y1 * 0.25 * (applied_exit_mult * 1.25))
    sotp_equity = max(0.0, sotp_ev - celkovy_dluh + hotovost)

    # Model odkupu za cizí kapitál (LBO Base Case)
    vstupni_cizi_kapital = celkova_hodnota_podniku * applied_leverage
    vstupni_vlastni_kapital = max(1000.0, celkova_hodnota_podniku - vstupni_cizi_kapital)

    def project_lbo_years(growth_rate: float):
        current_debt = vstupni_cizi_kapital
        proj_ebitda = ebitda if ebitda > 0 else base_fcf * 1.35
        proj_fcf = base_fcf
        projected_years = []
        for year in range(1, 6):
            proj_ebitda *= (1.0 + growth_rate)
            proj_fcf *= (1.0 + growth_rate)
            urokove_naklady = current_debt * implikovana_sazba_dluhu
            volny_tok_pro_splaceni = max(0.0, proj_fcf - urokove_naklady * (1.0 - efektivni_danova_sazba))
            splatka_jistiny = min(current_debt, volny_tok_pro_splaceni * 0.85)
            current_debt -= splatka_jistiny
            projected_years.append({
                "rok": year,
                "projektovany_provozni_zisk_pred_odpisy": round(proj_ebitda, 2),
                "volny_tok_pro_splaceni": round(volny_tok_pro_splaceni, 2),
                "zbyvajici_cizi_kapital": round(current_debt, 2),
                "mor_jistiny": round(splatka_jistiny, 2),
            })
        return projected_years

    projektovane_roky = project_lbo_years(applied_growth)
    konecna_ebitda = projektovane_roky[-1]["projektovany_provozni_zisk_pred_odpisy"]
    konecny_dluh = projektovane_roky[-1]["zbyvajici_cizi_kapital"]
    vystupni_hodnota_podniku = konecna_ebitda * applied_exit_mult
    vystupni_hodnota_vlastniho_kapitalu = max(0.0, vystupni_hodnota_podniku - konecny_dluh)

    moic = (vystupni_hodnota_vlastniho_kapitalu / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0.0
    irr = ((moic ** (1.0 / 5.0)) - 1.0) * 100.0 if moic > 0 else -100.0

    # Citlivostní matice vnitřní míry výnosnosti
    citlivostni_matice_irr = []
    testovane_rustove_sazby = [max(-0.50, applied_growth + offset) for offset in [-0.04, -0.02, 0.0, 0.02, 0.04]]
    testovane_nasobky = [max(2.0, round(applied_exit_mult + offset, 1)) for offset in [-3.0, -1.5, 0.0, 1.5, 3.0]]

    for g_rate in testovane_rustove_sazby:
        radek_matice = {"tempo_rustu_ebitda_procent": round(g_rate * 100.0, 1), "vysledky_irr_dle_nasobku": []}
        sim_years = project_lbo_years(g_rate)
        ebitda_t5 = sim_years[-1]["projektovany_provozni_zisk_pred_odpisy"]
        debt_t5 = sim_years[-1]["zbyvajici_cizi_kapital"]
        for mult in testovane_nasobky:
            eq_t5 = max(0.0, ebitda_t5 * mult - debt_t5)
            sim_moic = (eq_t5 / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0.0
            sim_irr = ((sim_moic ** (1.0 / 5.0)) - 1.0) * 100.0 if sim_moic > 0 else -100.0
            radek_matice["vysledky_irr_dle_nasobku"].append({
                "vystupni_nasobek": mult,
                "irr_procent": round(sim_irr, 2),
            })
        citlivostni_matice_irr.append(radek_matice)

    # Výpočet cenových pásem kolem střední vnitřní hodnoty (-15 % / +15 %)
    zelena_horni = round(intrinsic_value_per_share * 0.85, 2)
    cervena_dolni = round(intrinsic_value_per_share * 1.15, 2)

    current_price_float = float(current_price) if current_price else None
    upside_downside_percent = (
        round(((intrinsic_value_per_share / current_price_float) - 1.0) * 100.0, 2)
        if current_price_float and current_price_float > 0
        else None
    )

    if current_price_float is None:
        aktualni_pasmo = "NEZNÁMÉ"
    elif current_price_float <= zelena_horni:
        aktualni_pasmo = "ZELENÉ"
    elif current_price_float <= cervena_dolni:
        aktualni_pasmo = "ŽLUTOORANŽOVÉ"
    else:
        aktualni_pasmo = "ČERVENÉ"

    return {
        "spolecnost": ticker_clean,
        "nazev_spolecnosti": company_name,
        "mena": currency,
        "trzni_kapitalizace": round(trzni_kapitalizace, 2),
        "celkovy_dluh": round(celkovy_dluh, 2),
        "hotovost": round(hotovost, 2),
        "celkova_hodnota_podniku": round(celkova_hodnota_podniku, 2),
        "metoda_souctu_casti": {
            "vypoctena_hodnota_podniku": round(sotp_ev, 2),
            "vypoctena_hodnota_vlastniho_kapitalu": round(sotp_equity, 2),
        },
        "model_odkupu_za_cizi_kapital_base_case": {
            "vstupni_vlastni_kapital": round(vstupni_vlastni_kapital, 2),
            "vstupni_cizi_kapital": round(vstupni_cizi_kapital, 2),
            "vystupni_hodnota_vlastniho_kapitalu": round(vystupni_hodnota_vlastniho_kapitalu, 2),
            "moic": round(moic, 2),
            "irr_procent": round(irr, 2),
        },
        "dcf_valuation": {
            "wacc": round(wacc, 6),
            "terminal_value": round(terminal_value, 2),
            "present_value_terminal_value": round(present_value_terminal_value, 2),
            "enterprise_value": round(dcf_enterprise_value, 2),
            "equity_value": round(dcf_equity_value, 2),
            "intrinsic_value_per_share": intrinsic_value_per_share,
            "dcf_only_share_price": round(dcf_share_price, 2),
            "comps_only_share_price": round(comps_avg_price, 2),
            "current_share_price": round(current_price_float, 2) if current_price_float else None,
            "upside_downside_percent": upside_downside_percent,
            "cash_flows": dcf_cash_flows,
        },
        "valuacni_pasma": {
            "aktualni_cena": round(current_price_float, 2) if current_price_float else None,
            "stredni_vnitrni_hodnota": intrinsic_value_per_share,
            "zelene_pasmo_od": 0.0,
            "zelene_pasmo_do": zelena_horni,
            "zlutooranzove_pasmo_od": zelena_horni,
            "zlutooranzove_pasmo_do": cervena_dolni,
            "cervene_pasmo_od": cervena_dolni,
            "aktualni_pasmo": aktualni_pasmo,
        },
        "peer_comps": {
            "sector": sector,
            "multiples": peer_comps_rows,
        },
        "parametry_scenare": {
            "leverage": round(applied_leverage, 4),
            "growth": round(applied_growth, 4),
            "exit_mult": round(applied_exit_mult, 2),
            "testovane_nasobky": testovane_nasobky,
        },
        "projektovane_roky": projektovane_roky,
        "citlivostni_matice_lbo_irr": citlivostni_matice_irr,
    }


def build_newsletter_email_html(val_data: dict, news_articles: list[dict], is_scheduled: bool = False) -> str:
    ticker = escape(val_data["spolecnost"])
    nazev = escape(val_data["nazev_spolecnosti"])
    mena = escape(val_data["mena"])
    pasma = val_data["valuacni_pasma"]
    dcf = val_data["dcf_valuation"]

    aktualni = f"{pasma['aktualni_cena']:,.2f} {mena}" if pasma["aktualni_cena"] is not None else "Nedostupná"
    vnitrni = f"{pasma['stredni_vnitrni_hodnota']:,.2f} {mena}"
    zelene = f"0,00 – {pasma['zelene_pasmo_do']:,.2f} {mena}"
    zlute = f"{pasma['zlutooranzove_pasmo_od']:,.2f} – {pasma['zlutooranzove_pasmo_do']:,.2f} {mena}"
    cervene = f"nad {pasma['cervene_pasmo_od']:,.2f} {mena}"
    pasmo_nazev = escape(pasma["aktualni_pasmo"])
    typ_zpravy = "Pravidelný investiční přehled (Středa 7:00 • Pátek 7:00 • Neděle 19:00)" if is_scheduled else "Jednorázový investiční přehled"

    news_html = ""
    for art in news_articles[:5]:
        title = escape(art.get("title", ""))
        source = escape(art.get("source", ""))
        url = escape(art.get("url", "#"))
        news_html += (
            f"<li style='margin-bottom:10px;'>"
            f"<a href='{url}' style='color:#0284c7;text-decoration:none;font-weight:600;'>{title}</a>"
            f"<br><span style='color:#64748b;font-size:12px;'>{source}</span></li>"
        )

    if not news_html:
        news_html = "<li>Žádné aktuální zprávy pro tento titul nebyly nalezeny.</li>"

    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; background-color: #f1f5f9; color: #0f172a; padding: 24px;">
        <div style="max-width: 640px; margin: 0 auto; background: #ffffff; border-radius: 12px; padding: 28px; border: 1px solid #cbd5e1;">
            <p style="font-size: 11px; letter-spacing: 2px; color: #0284c7; font-weight: bold; text-transform: uppercase; margin: 0 0 8px 0;">
                TENSEGRITY TERMINAL • {escape(typ_zpravy)}
            </p>
            <h2 style="margin: 0 0 16px 0; color: #0f172a;">Valuační přehled: {nazev} ({ticker})</h2>
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px;">
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">Aktuální tržní cena akcie:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold; text-align: right;">{aktualni}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">Střední vnitřní hodnota:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #0284c7; text-align: right;">{vnitrni}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">Aktuální cenové pásmo:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold; text-align: right;">{pasmo_nazev}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #059669;">Zelené pásmo (Podhodnoceno):</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; text-align: right;">{zelene}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #d97706;">Žlutooranžové pásmo (Férová hodnota):</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; text-align: right;">{zlute}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; color: #dc2626;">Červené pásmo (Nadhodnoceno):</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; text-align: right;">{cervene}</td></tr>
                <tr><td style="padding: 8px;">Vážené průměrné náklady kapitálu:</td><td style="padding: 8px; text-align: right;">{dcf['wacc'] * 100:.2f} %</td></tr>
            </table>
            <h3 style="margin: 20px 0 10px 0; font-size: 16px;">Aktuální tržní zprávy ({ticker})</h3>
            <ul style="padding-left: 18px; margin: 0 0 20px 0; line-height: 1.5;">
                {news_html}
            </ul>
            <p style="font-size: 11px; color: #64748b; margin: 0;">
                Vygenerováno platformou Tensegrity Terminal dne {datetime.now(PRAGUE_TZ).strftime('%d.%m.%Y v %H:%M')}.
            </p>
        </div>
    </body>
    </html>
    """


def send_email_sync(recipient: str, subject: str, html_body: str) -> tuple[bool, str]:
    smtp_cfg = get_smtp_config()
    if not smtp_cfg:
        return False, "E-mailový server není nakonfigurován (chybí proměnné SENDER_EMAIL a SENDER_PASSWORD)."
    try:
        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        message["From"] = f"Tensegrity Terminal <{smtp_cfg['username']}>"
        message["To"] = recipient
        message.attach(MIMEText(html_body, "html", "utf-8"))

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], context=context, timeout=20) as server:
            server.login(smtp_cfg["username"], smtp_cfg["password"])
            server.sendmail(smtp_cfg["username"], recipient, message.as_string())
        return True, f"Investiční přehled byl úspěšně odeslán na adresu {recipient}."
    except Exception as exc:
        logger.error("Chyba při odesílání e-mailu pro %s: %s", recipient, exc)
        return False, f"Nepodařilo se odeslat e-mail: {exc}"


async def scheduled_newsletter_worker():
    global _last_dispatched_slot
    while True:
        try:
            now_prg = datetime.now(PRAGUE_TZ)
            current_tuple = (now_prg.weekday(), now_prg.hour, now_prg.minute)
            slot_key = now_prg.strftime("%Y-%m-%d-%H-%M")

            if current_tuple in FIXED_SCHEDULE_SLOTS and _last_dispatched_slot != slot_key:
                _last_dispatched_slot = slot_key
                subscribers = get_subscribers()
                if subscribers:
                    logger.info("Spouštím pravidelné rozesílání newsletteru pro %d odběratelů.", len(subscribers))
                    for sub in subscribers:
                        email = sub["email"]
                        ticker = sub.get("ticker") or "NVDA"
                        try:
                            val_data = await asyncio.to_thread(run_valuation_logic, ticker)
                            news_articles = await asyncio.to_thread(fetch_ticker_news, ticker, 5)
                            html_body = build_newsletter_email_html(val_data, news_articles, is_scheduled=True)
                            subject = f"Tensegrity Terminal | Pravidelný přehled: {val_data['spolecnost']}"
                            await asyncio.to_thread(send_email_sync, email, subject, html_body)
                        except Exception as item_err:
                            logger.error("Chyba při plánovaném odeslání pro %s (%s): %s", email, ticker, item_err)
        except Exception as loop_err:
            logger.error("Chyba v plánovači newsletteru: %s", loop_err)

        await asyncio.sleep(25)


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(scheduled_newsletter_worker())
    yield
    task.cancel()


app = FastAPI(title="Tensegrity Terminal API", version="13.0", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    logo_file = os.path.join(STATIC_DIR, "logo.png")
    if os.path.exists(logo_file):
        return FileResponse(logo_file)
    raise HTTPException(status_code=404, detail="Ikona nenalezena.")


@app.get("/api/forex-rates")
async def get_forex_rates():
    return await asyncio.to_thread(get_live_forex_rates)


@app.get("/api/valuation/{ticker}")
async def get_institutional_terminal_valuation(
    ticker: str,
    leverage: Optional[float] = Query(None, ge=0.0, le=1.0),
    growth: Optional[float] = Query(None, gt=-0.9, le=5.0),
    exit_mult: Optional[float] = Query(None, gt=0.5, le=100.0),
):
    return await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)


@app.get("/api/news/{ticker}")
async def get_ticker_news_endpoint(ticker: str):
    articles = await asyncio.to_thread(fetch_ticker_news, ticker, 8)
    return {"ticker": ticker.upper(), "articles": articles}


@app.get("/api/subscribers")
async def list_subscribers_endpoint():
    return {"subscribers": get_subscribers()}


@app.delete("/api/subscribers")
async def delete_subscriber_endpoint(email: str = Query(...)):
    removed = remove_subscriber(email)
    if not removed:
        raise HTTPException(status_code=404, detail="E-mailová adresa nebyla v evidenci nalezena.")
    return {"status": "ok", "message": f"Adresa {email} byla z pravidelného odběru odhlášena."}


@app.post("/subscribe")
async def handle_subscribe(request: Request, email: str = Query("")):
    candidate = email.strip()
    ticker = "NVDA"
    if request.method == "POST":
        try:
            payload = await request.json()
            if not candidate:
                candidate = str(payload.get("email", "")).strip()
            ticker = str(payload.get("ticker", "NVDA")).strip().upper() or "NVDA"
        except Exception:
            pass
    if not is_valid_email_address(candidate):
        raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")
    created = upsert_subscriber(candidate, ticker)
    return {
        "status": "ok",
        "message": (
            f"Adresa {candidate} byla zaevidována. Přehled pro {ticker} bude pravidelně odesílán každou středu v 7:00, pátek v 7:00 a neděli v 19:00."
            if created
            else f"Adresa {candidate} je již v evidenci (sledovaný titul aktualizován na {ticker})."
        ),
    }


@app.post("/api/newsletter/send-now")
async def send_newsletter_now_endpoint(request: Request):
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Neplatný formát požadavku.")

    recipient = str(payload.get("email", "")).strip()
    ticker = str(payload.get("ticker", "NVDA")).strip().upper() or "NVDA"
    leverage = payload.get("leverage")
    growth = payload.get("growth")
    exit_mult = payload.get("exit_mult")

    if not is_valid_email_address(recipient):
        raise HTTPException(status_code=422, detail="Zadejte platnou e-mailovou adresu.")

    val_data = await asyncio.to_thread(
        run_valuation_logic,
        ticker,
        float(leverage) if leverage is not None else None,
        float(growth) if growth is not None else None,
        float(exit_mult) if exit_mult is not None else None,
    )
    news_articles = await asyncio.to_thread(fetch_ticker_news, ticker, 5)
    html_body = build_newsletter_email_html(val_data, news_articles, is_scheduled=False)
    subject = f"Tensegrity Terminal | Okamžitý valuační přehled: {val_data['spolecnost']}"

    sent, msg = await asyncio.to_thread(send_email_sync, recipient, subject, html_body)
    if not sent:
        raise HTTPException(status_code=500, detail=msg)
    return {"status": "ok", "message": msg}


@app.post("/api/advisor/{ticker}")
async def ask_terminal_advisor(ticker: str, request: Request):
    try:
        payload = await request.json()
        question = str(payload.get("question", "")).strip()[:1000]
        scenario = payload.get("scenario") or {}
        leverage = scenario.get("leverage")
        growth = scenario.get("growth")
        exit_mult = scenario.get("exit_mult")

        if not question:
            raise HTTPException(status_code=422, detail="Zadejte prosím svůj dotaz.")

        data = await asyncio.to_thread(
            run_valuation_logic,
            ticker,
            float(leverage) if leverage is not None else None,
            float(growth) if growth is not None else None,
            float(exit_mult) if exit_mult is not None else None,
        )
        pasma = data["valuacni_pasma"]
        dcf = data["dcf_valuation"]
        mena = data["mena"]

        answer = (
            f"Analytický přehled pro {data['nazev_spolecnosti']} ({data['spolecnost']}): "
            f"Aktuální tržní cena činí {pasma['aktualni_cena']} {mena}, zatímco střední vnitřní hodnota modelu "
            f"vychází na {pasma['stredni_vnitrni_hodnota']} {mena} (při vážených průměrných nákladech kapitálu "
            f"{dcf['wacc'] * 100:.2f} % a očekávaném ročním růstu {data['parametry_scenare']['growth'] * 100:.1f} %). "
            f"Akcie se nyní nachází v pásmu: {pasma['aktualni_pasmo']}. "
            f"Zelené pásmo podhodnocení je do {pasma['zelene_pasmo_do']} {mena}, férové pásmo leží mezi "
            f"{pasma['zlutooranzove_pasmo_od']} a {pasma['zlutooranzove_pasmo_do']} {mena}, a nad {pasma['cervene_pasmo_od']} {mena} "
            f"začíná červené pásmo nadhodnocení. K vašemu dotazu („{question}“): klíčovou citlivostí ocenění je "
            f"udržení tempa růstu volných peněžních toků a výstupního násobku v 5. roce."
        )
        return {"answer": answer}
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=400, detail=str(error))


@app.get("/api/export/excel/{ticker}")
async def export_valuation_excel(
    ticker: str,
    leverage: Optional[float] = Query(None, ge=0.0, le=1.0),
    growth: Optional[float] = Query(None, gt=-0.9, le=5.0),
    exit_mult: Optional[float] = Query(None, gt=0.5, le=100.0),
):
    data = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Valuace - {data['spolecnost']}"
    pasma = data["valuacni_pasma"]
    mena = data["mena"]

    ws.append(["TENSEGRITY TERMINAL — INSTITUCIONÁLNÍ VALUAČNÍ ZPRÁVA"])
    ws.append([f"Společnost: {data['nazev_spolecnosti']} ({data['spolecnost']}) | Datum: {datetime.now(PRAGUE_TZ).strftime('%d.%m.%Y')}"])
    ws.append([])
    ws.append(["Ukazatel", f"Hodnota ({mena})"])
    ws.append(["Aktuální tržní cena akcie", pasma["aktualni_cena"]])
    ws.append(["Střední vnitřní hodnota na akcii", pasma["stredni_vnitrni_hodnota"]])
    ws.append(["Zelené pásmo (Podhodnoceno) — horní hranice", pasma["zelene_pasmo_do"]])
    ws.append(["Žlutooranžové pásmo (Férová hodnota) — horní hranice", pasma["zlutooranzove_pasmo_do"]])
    ws.append(["Červené pásmo (Nadhodnoceno) — od", pasma["cervene_pasmo_od"]])
    ws.append(["Tržní kapitalizace", data["trzni_kapitalizace"]])
    ws.append(["Celkový dluh", data["celkovy_dluh"]])
    ws.append(["Celková hodnota podniku", data["celkova_hodnota_podniku"]])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Tensegrity_Valuace_{data['spolecnost']}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.get("/api/export/pdf/{ticker}")
async def export_valuation_pdf(
    ticker: str,
    leverage: Optional[float] = Query(None, ge=0.0, le=1.0),
    growth: Optional[float] = Query(None, gt=-0.9, le=5.0),
    exit_mult: Optional[float] = Query(None, gt=0.5, le=100.0),
):
    data = await asyncio.to_thread(run_valuation_logic, ticker, leverage, growth, exit_mult)
    pasma = data["valuacni_pasma"]
    mena = data["mena"]

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=45)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph(f"Tensegrity Terminal — Valuace: {data['spolecnost']}", styles["Heading1"]),
        Spacer(1, 10),
        Paragraph(f"Aktualni trzni cena akcie: {pasma['aktualni_cena']} {mena}", styles["Normal"]),
        Paragraph(f"Stredni vnitrni hodnota: {pasma['stredni_vnitrni_hodnota']} {mena}", styles["Normal"]),
        Paragraph(f"Zelene pasmo (podhodnoceno): 0.00 - {pasma['zelene_pasmo_do']} {mena}", styles["Normal"]),
        Paragraph(f"Zlutooranzove pasmo (ferova hodnota): {pasma['zlutooranzove_pasmo_od']} - {pasma['zlutooranzove_pasmo_do']} {mena}", styles["Normal"]),
        Paragraph(f"Cervene pasmo (nadhodnoceno): nad {pasma['cervene_pasmo_od']} {mena}", styles["Normal"]),
        Spacer(1, 10),
        Paragraph(f"Celkova hodnota podniku: {data['celkova_hodnota_podniku']:,.0f} {mena}", styles["Normal"]),
        Paragraph(f"Trzni kapitalizace: {data['trzni_kapitalizace']:,.0f} {mena}", styles["Normal"]),
    ]
    doc.build(elements)
    buffer.seek(0)
    filename = f"Tensegrity_Valuace_{data['spolecnost']}.pdf"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


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
            --bg-card: rgba(15, 22, 37, 0.88);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --accent-blue: #38bdf8;
            --accent-blue-glow: rgba(56, 189, 248, 0.22);
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
        .brand-title { font-size: 16px; font-weight: 800; letter-spacing: 4px; color: #fff; text-transform: uppercase; }
        .brand-subtitle { font-size: 10px; letter-spacing: 6px; color: var(--text-secondary); text-transform: uppercase; margin-top: 2px; }
        .controls { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
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
            font-size: 14px; font-weight: 600; padding: 10px 16px; border-radius: 10px;
            outline: none; width: 140px; text-transform: uppercase; letter-spacing: 1px;
        }
        .search-wrapper input:focus { border-color: var(--accent-blue); box-shadow: 0 0 12px var(--accent-blue-glow); }
        .btn {
            font-family: 'Plus Jakarta Sans', sans-serif; font-size: 13px; font-weight: 600;
            padding: 10px 18px; border-radius: 10px; border: 1px solid transparent; cursor: pointer;
            display: inline-flex; align-items: center; justify-content: center; gap: 8px; transition: all 0.2s ease;
        }
        .btn:disabled { opacity: 0.6; cursor: not-allowed; }
        .btn-primary { background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%); color: #040914; font-weight: 700; }
        .btn-primary:hover:not(:disabled) { opacity: 0.92; transform: translateY(-1px); }
        .btn-excel { background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.35); color: #34d399; }
        .btn-pdf { background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.35); color: #f87171; }
        .btn-news { background: rgba(148, 163, 184, 0.14); border: 1px solid var(--border-subtle); color: var(--text-primary); }
        .btn-news:hover { border-color: var(--accent-blue); color: var(--accent-blue); }
        .btn-subtle { background: rgba(56, 189, 248, 0.1); border: 1px solid rgba(56, 189, 248, 0.25); color: var(--accent-blue); font-size: 12px; padding: 6px 12px; }

        .scenario-bar {
            display: flex; align-items: center; flex-wrap: wrap; gap: 16px 24px;
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 12px; padding: 14px 22px; margin-bottom: 24px;
        }
        .scenario-heading { color: var(--accent-blue); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .scenario-field { display: flex; align-items: center; gap: 10px; color: var(--text-secondary); font-size: 12px; }
        .scenario-input { display: inline-flex; align-items: center; gap: 6px; }
        .scenario-input input {
            width: 78px; padding: 7px 8px; border: 1px solid var(--border-subtle); border-radius: 6px;
            background: rgba(6, 9, 17, 0.75); color: var(--text-primary); font: 600 13px 'JetBrains Mono', monospace; text-align: right;
        }

        .grid-3 { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 24px; margin-bottom: 24px; }
        .card {
            background: var(--bg-card); border: 1px solid var(--border-subtle);
            border-radius: 16px; padding: 24px 28px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        }
        .matrix-card { margin-bottom: 24px; }
        .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid var(--border-subtle); gap: 12px; }
        .card-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: var(--text-secondary); }
        .badge { font-size: 11px; font-family: 'JetBrains Mono', monospace; padding: 3px 9px; border-radius: 6px; background: rgba(56, 189, 248, 0.1); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.25); }
        .metric-row { display: flex; justify-content: space-between; align-items: center; margin: 12px 0; font-size: 14px; }
        .metric-label { color: var(--text-secondary); }
        .metric-value { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: var(--text-primary); font-size: 15px; }
        .metric-value.highlight { color: var(--accent-blue); }

        .thesis-layout { display: grid; grid-template-columns: auto 1fr; gap: 24px; align-items: center; }
        .traffic-lights { display: flex; gap: 10px; padding: 12px; border: 1px solid var(--border-subtle); border-radius: 999px; background: rgba(6, 9, 17, 0.65); }
        .traffic-light { width: 18px; height: 18px; border-radius: 50%; opacity: 0.22; }
        .traffic-light.active { opacity: 1; box-shadow: 0 0 16px currentColor; }
        .traffic-light.green { color: #10b981; background: #10b981; }
        .traffic-light.orange { color: #f59e0b; background: #f59e0b; }
        .traffic-light.red { color: #ef4444; background: #ef4444; }

        /* Cenová pásma akcie */
        .bands-summary {
            display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 16px;
            padding: 16px 20px; border-radius: 12px; background: rgba(6, 9, 17, 0.65);
            border: 1px solid var(--border-subtle); margin-bottom: 18px;
        }
        .bands-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
        .band-box {
            padding: 18px 20px; border-radius: 12px; border: 1px solid var(--border-subtle);
            background: rgba(6, 9, 17, 0.55); transition: all 0.2s ease;
        }
        .band-box.green { border-color: rgba(16, 185, 129, 0.35); }
        .band-box.orange { border-color: rgba(245, 158, 11, 0.35); }
        .band-box.red { border-color: rgba(239, 68, 68, 0.35); }
        .band-box.active-green { background: rgba(16, 185, 129, 0.14); border-color: #10b981; box-shadow: 0 0 20px rgba(16, 185, 129, 0.15); }
        .band-box.active-orange { background: rgba(245, 158, 11, 0.14); border-color: #f59e0b; box-shadow: 0 0 20px rgba(245, 158, 11, 0.15); }
        .band-box.active-red { background: rgba(239, 68, 68, 0.14); border-color: #ef4444; box-shadow: 0 0 20px rgba(239, 68, 68, 0.15); }
        .band-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; }
        .band-range { font-family: 'JetBrains Mono', monospace; font-size: 17px; font-weight: 700; color: var(--text-primary); margin-bottom: 6px; }
        .band-desc { font-size: 12px; color: var(--text-secondary); line-height: 1.45; }
        .active-pill { font-size: 10px; padding: 2px 7px; border-radius: 4px; background: rgba(255,255,255,0.12); color: #fff; font-family: 'JetBrains Mono', monospace; }

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

        .advisor-thread { display: grid; gap: 12px; margin: 16px 0; max-height: 350px; overflow-y: auto; padding-right: 8px; }
        .advisor-message { padding: 12px 16px; border-left: 3px solid var(--accent-blue); background: rgba(56, 189, 248, 0.05); color: var(--text-secondary); font-size: 13px; line-height: 1.6; white-space: pre-wrap; border-radius: 0 8px 8px 0; }
        .advisor-message.user-message { border-color: var(--success); background: rgba(16, 185, 129, 0.05); color: var(--text-primary); }
        .advisor-form { display: flex; gap: 12px; margin-top: 12px; flex-wrap: wrap; }
        .advisor-form input { flex: 1; min-width: 220px; padding: 12px 16px; border: 1px solid var(--border-subtle); border-radius: 8px; background: rgba(6, 9, 17, 0.8); color: var(--text-primary); font-size: 13px; outline: none; }
        .advisor-form input:focus { border-color: var(--accent-blue); }
        .plain-guide { margin-top: 12px; font-size: 12px; color: var(--text-secondary); line-height: 1.5; }

        /* Modální okno Newsletteru */
        .modal-backdrop { display: none; position: fixed; inset: 0; background: rgba(4, 7, 15, 0.84); backdrop-filter: blur(6px); z-index: 999; align-items: center; justify-content: center; padding: 20px; }
        .modal-backdrop.open { display: flex; }
        .modal-box { background: #0f1625; border: 1px solid var(--border-subtle); border-radius: 16px; max-width: 740px; width: 100%; padding: 28px; box-shadow: 0 20px 50px rgba(0,0,0,0.75); max-height: 88vh; overflow-y: auto; }
        .newsletter-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 12px; }
        @media (max-width: 600px) { .newsletter-actions { grid-template-columns: 1fr; } }
        .schedule-note { margin-top: 12px; padding: 10px 14px; border-radius: 8px; background: rgba(56, 189, 248, 0.08); border: 1px solid rgba(56, 189, 248, 0.22); font-size: 12px; color: var(--text-secondary); }
        .subscribers-box { margin-top: 18px; padding-top: 16px; border-top: 1px solid var(--border-subtle); }
        .subscriber-row { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; background: rgba(6, 9, 17, 0.6); border: 1px solid var(--border-subtle); border-radius: 8px; margin-top: 8px; font-size: 12px; }
        .news-item { padding: 12px 0; border-bottom: 1px solid var(--border-subtle); }
        .news-item a { color: var(--accent-blue); text-decoration: none; font-weight: 600; font-size: 14px; }
        .news-item a:hover { text-decoration: underline; }
        .news-meta { font-size: 11px; color: var(--text-muted); margin-top: 4px; font-family: 'JetBrains Mono', monospace; }

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
                    <input type="text" id="tickerInput" value="NVDA" placeholder="NVDA" aria-label="Symbol společnosti" onkeydown="if(event.key==='Enter') loadValuation(true)">
                </div>
                <div class="forex-strip" id="forexRates">
                    <span>EUR/CZK <strong>--</strong></span>
                    <span>USD/CZK <strong>--</strong></span>
                    <span>GBP/CZK <strong>--</strong></span>
                </div>
                <button class="btn btn-primary" onclick="loadValuation(true)">Analyzovat</button>
                <button class="btn btn-excel" onclick="downloadExcel()">Excel Model</button>
                <button class="btn btn-pdf" onclick="downloadPDF()">Přehled ke stažení</button>
                <button class="btn btn-news" onclick="openNewsletterModal()">Newsletter &amp; Tržní zprávy</button>
            </div>
        </header>

        <section class="scenario-bar" aria-label="Parametry scénáře">
            <span class="scenario-heading">Parametry Scénáře</span>
            <label class="scenario-field" for="leverageInput">
                Cizí kapitál (% Dluhu)
                <span class="scenario-input"><input id="leverageInput" type="number" min="0" max="100" step="1" value="30"><span>%</span></span>
            </label>
            <label class="scenario-field" for="growthInput">
                Růst provozního zisku
                <span class="scenario-input"><input id="growthInput" type="number" min="-50" max="150" step="0.5" value="15"><span>%</span></span>
            </label>
            <label class="scenario-field" for="exitMultipleInput">
                Výstupní násobek hodnoty
                <span class="scenario-input"><input id="exitMultipleInput" type="number" min="1" max="100" step="0.5" value="22.0"><span>x</span></span>
            </label>
            <button class="btn btn-subtle" onclick="loadValuation(false)">Recalculate s vlastními parametry</button>
        </section>

        <section class="card matrix-card" id="investmentThesis">
            <div class="card-header">
                <span class="card-title">Investiční Semafor &amp; Teze</span>
                <span class="badge" id="thesisTicker">NVDA</span>
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

        <section class="card matrix-card" id="priceBandsCard">
            <div class="card-header">
                <span class="card-title">Valuační cenová pásma akcie (Aktuální cena vs. intervaly)</span>
                <span class="badge" id="bandsTicker">NVDA</span>
            </div>
            <div class="bands-summary">
                <div>
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 1px;">Aktuální tržní cena akcie</div>
                    <div class="metric-value" style="font-size: 22px; margin-top: 4px;" id="bandCurrentPrice">-</div>
                </div>
                <div>
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 1px;">Střední vnitřní hodnota</div>
                    <div class="metric-value highlight" style="font-size: 22px; margin-top: 4px;" id="bandIntrinsicPrice">-</div>
                </div>
                <div>
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 1px;">Aktuální pásmo</div>
                    <div style="margin-top: 6px;"><span class="valuation-status neutral" id="bandCurrentStatus">-</span></div>
                </div>
            </div>
            <div class="bands-grid">
                <div class="band-box green" id="bandBoxGreen">
                    <div class="band-title" style="color: #34d399;">
                        <span>Zelené pásmo (Podhodnoceno)</span>
                        <span class="active-pill" id="pillGreen" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="band-range" id="rangeGreen">-</div>
                    <div class="band-desc">Bezpečnostní marže alespoň 15 % pod vnitřní hodnotou modelu (atraktivní valuace).</div>
                </div>
                <div class="band-box orange" id="bandBoxOrange">
                    <div class="band-title" style="color: #fbbf24;">
                        <span>Žlutooranžové pásmo (Férová hodnota)</span>
                        <span class="active-pill" id="pillOrange" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="band-range" id="rangeOrange">-</div>
                    <div class="band-desc">Neutrální koridor plus minus 15 % kolem odhadované reálné hodnoty akcie.</div>
                </div>
                <div class="band-box red" id="bandBoxRed">
                    <div class="band-title" style="color: #f87171;">
                        <span>Červené pásmo (Nadhodnoceno)</span>
                        <span class="active-pill" id="pillRed" style="display:none;">AKTUÁLNÍ CENA</span>
                    </div>
                    <div class="band-range" id="rangeRed">-</div>
                    <div class="band-desc">Tržní cena přesahuje vnitřní hodnotu o více než 15 % (zvýšené valuační riziko).</div>
                </div>
            </div>
        </section>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Tržní Data</span>
                    <span class="badge" id="valTicker">NVDA</span>
                </div>
                <div class="metric-row"><span class="metric-label">Tržní kapitalizace:</span><span class="metric-value" id="valCap">-</span></div>
                <div class="metric-row"><span class="metric-label">Celkový dluh:</span><span class="metric-value" id="valDebt">-</span></div>
                <div class="metric-row"><span class="metric-label">Celková hodnota podniku:</span><span class="metric-value highlight" id="valEV">-</span></div>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Metoda součtu částí</span>
                    <span class="badge">Rozklad divizí</span>
                </div>
                <div class="metric-row"><span class="metric-label">Hodnota podniku:</span><span class="metric-value" id="sotpEV">-</span></div>
                <div class="metric-row"><span class="metric-label">Vlastní kapitál:</span><span class="metric-value highlight" id="sotpEquity">-</span></div>
                <p class="plain-guide">Ocenění podniku metodou součtu jednotlivých segmentů po odečtení čistého dluhu.</p>
            </div>

            <div class="card">
                <div class="card-header">
                    <span class="card-title">Odkup za cizí kapitál</span>
                    <span class="badge">Základní scénář</span>
                </div>
                <div class="metric-row"><span class="metric-label">Vstupní vlastní kapitál:</span><span class="metric-value" id="lboInEq">-</span></div>
                <div class="metric-row"><span class="metric-label">Výstupní vlastní kapitál:</span><span class="metric-value" id="lboOutEq">-</span></div>
                <div class="metric-row"><span class="metric-label">Násobek kapitálu:</span><span class="metric-value highlight" id="lboMoic">-</span></div>
                <div class="metric-row"><span class="metric-label">Vnitřní míra výnosnosti:</span><span class="metric-value highlight" id="lboIrr">-</span></div>
            </div>
        </div>

        <div class="grid-3">
            <div class="card">
                <div class="card-header">
                    <span class="card-title">Diskontované peněžní toky</span>
                    <span class="badge">Vnitřní hodnota</span>
                </div>
                <div class="metric-row"><span class="metric-label">Střední vnitřní hodnota / akcii:</span><span class="metric-value highlight" id="dcfIntrinsicValue">-</span></div>
                <div class="metric-row"><span class="metric-label">Vážené náklady kapitálu:</span><span class="metric-value" id="dcfWacc">-</span></div>
                <div class="metric-row"><span class="metric-label">Terminální hodnota:</span><span class="metric-value" id="dcfTerminalValue">-</span></div>
                <div class="metric-row"><span class="metric-label">Tržní postavení:</span><span class="valuation-status" id="dcfValuationStatus">-</span></div>
            </div>

            <div class="card" style="grid-column: span 2;">
                <div class="card-header">
                    <span class="card-title">Srovnání konkurence dle násobků</span>
                    <span class="badge" id="peerSector">Sektor</span>
                </div>
                <div class="table-responsive" id="peerCompsContainer">
                    <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
                </div>
            </div>
        </div>

        <div class="card matrix-card">
            <div class="card-header">
                <span class="card-title">5letý amortizační kalendář &amp; peněžní toky</span>
                <span class="badge">Splácení dluhu</span>
            </div>
            <div class="table-responsive" id="lboScheduleContainer">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
        </div>

        <div class="card matrix-card">
            <div class="card-header">
                <span class="card-title">Citlivostní matice vnitřní míry výnosnosti (%)</span>
                <span class="badge">Růst provozního zisku vs. výstupní násobek</span>
            </div>
            <div class="table-responsive" id="sensitivityMatrix">
                <div style="text-align: center; padding: 20px;"><div class="loader"></div></div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <span class="card-title">Investiční poradce Tensegrity</span>
                <span class="badge">Analytický komentář</span>
            </div>
            <div class="advisor-thread" id="advisorThread">
                <div class="advisor-message">Dobrý den! Zeptejte se mě na cokoli ohledně ocenění společnosti, valuačních pásem, diskontní sazby nebo citlivosti modelu.</div>
            </div>
            <form class="advisor-form" onsubmit="askAdvisor(event)">
                <input id="advisorQuestion" type="text" placeholder="Zadejte dotaz k ocenění akcie..." required>
                <button class="btn btn-primary" type="submit">Odeslat dotaz</button>
            </form>
        </div>
    </div>

    <div class="modal-backdrop" id="newsletterModal" onclick="if(event.target===this) closeNewsletterModal()">
        <div class="modal-box">
            <div class="card-header">
                <span class="card-title">Newsletter &amp; Tržní zprávy (<span id="newsModalTicker">NVDA</span>)</span>
                <button class="btn btn-news" style="padding: 4px 10px;" onclick="closeNewsletterModal()">✕</button>
            </div>

            <div class="advisor-form" style="margin-top: 0;">
                <input id="newsletterEmailInput" type="email" placeholder="Zadejte e-mailovou adresu..." required>
            </div>

            <div class="newsletter-actions">
                <button class="btn btn-primary" id="btnSendOneOff" type="button" onclick="sendNewsletterOnce()">
                    Odeslat jednorázově nyní
                </button>
                <button class="btn btn-excel" id="btnSubscribeRegular" type="button" onclick="registerRegularNewsletter()">
                    Zaevidovat pravidelný odběr
                </button>
            </div>

            <div class="schedule-note">
                <strong>Pevný rozvrh pravidelného odesílání:</strong> Každou středu v <strong>7:00</strong>, každý pátek v <strong>7:00</strong> a každou neděli v <strong>19:00</strong> (seč).
            </div>

            <div id="newsletterStatusMsg" style="font-size: 13px; margin-top: 12px; font-weight: 500;"></div>

            <div class="subscribers-box">
                <div style="font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--text-secondary);">
                    Evidované e-maily pro pravidelný odběr
                </div>
                <div id="subscribersListContainer"></div>
            </div>

            <div style="margin-top: 22px; padding-top: 16px; border-top: 1px solid var(--border-subtle);">
                <div style="font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--text-secondary); margin-bottom: 8px;">
                    Aktuální tržní zprávy
                </div>
                <div id="newsFeedList">Načítám zprávy...</div>
            </div>
        </div>
    </div>

    <script>
        let currentCurrency = 'USD';

        function formatCurrency(val, curr = currentCurrency) {
            if (val === null || val === undefined || isNaN(val)) return 'Nedostupné';
            return Number(val).toLocaleString('cs-CZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ' + curr;
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
                console.error('Chyba kurzů měn:', err);
            }
        }

        async function loadValuation(autoCalibrate = true) {
            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            const queryParams = new URLSearchParams();

            if (!autoCalibrate) {
                const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
                const growth = document.getElementById('growthInput').valueAsNumber / 100;
                const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;
                if (!isNaN(leverage)) queryParams.set('leverage', leverage);
                if (!isNaN(growth)) queryParams.set('growth', growth);
                if (!isNaN(exitMult)) queryParams.set('exit_mult', exitMult);
            }

            const url = '/api/valuation/' + encodeURIComponent(ticker) + (queryParams.toString() ? '?' + queryParams.toString() : '');

            try {
                const res = await fetch(url);
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Valuaci se nepodařilo načíst.');

                currentCurrency = data.mena || 'USD';
                document.getElementById('tickerInput').value = data.spolecnost;
                document.getElementById('thesisTicker').textContent = data.spolecnost;
                document.getElementById('bandsTicker').textContent = data.spolecnost;
                document.getElementById('valTicker').textContent = data.spolecnost;

                if (autoCalibrate && data.parametry_scenare) {
                    document.getElementById('leverageInput').value = Math.round(data.parametry_scenare.leverage * 100);
                    document.getElementById('growthInput').value = (data.parametry_scenare.growth * 100).toFixed(1);
                    document.getElementById('exitMultipleInput').value = data.parametry_scenare.exit_mult.toFixed(1);
                }

                // Cenová pásma akcie
                const pasma = data.valuacni_pasma;
                document.getElementById('bandCurrentPrice').textContent = formatCurrency(pasma.aktualni_cena);
                document.getElementById('bandIntrinsicPrice').textContent = formatCurrency(pasma.stredni_vnitrni_hodnota);
                document.getElementById('rangeGreen').textContent = '0,00 – ' + formatCurrency(pasma.zelene_pasmo_do);
                document.getElementById('rangeOrange').textContent = formatCurrency(pasma.zlutooranzove_pasmo_od) + ' – ' + formatCurrency(pasma.zlutooranzove_pasmo_do);
                document.getElementById('rangeRed').textContent = 'nad ' + formatCurrency(pasma.cervene_pasmo_od);

                ['bandBoxGreen', 'bandBoxOrange', 'bandBoxRed'].forEach(id => {
                    document.getElementById(id).classList.remove('active-green', 'active-orange', 'active-red');
                });
                ['pillGreen', 'pillOrange', 'pillRed'].forEach(id => {
                    document.getElementById(id).style.display = 'none';
                });

                const bandStatusEl = document.getElementById('bandCurrentStatus');
                if (pasma.aktualni_pasmo === 'ZELENÉ') {
                    document.getElementById('bandBoxGreen').classList.add('active-green');
                    document.getElementById('pillGreen').style.display = 'inline-block';
                    bandStatusEl.className = 'valuation-status undervalued';
                    bandStatusEl.textContent = 'Zelené pásmo (Podhodnoceno)';
                } else if (pasma.aktualni_pasmo === 'ŽLUTOORANŽOVÉ') {
                    document.getElementById('bandBoxOrange').classList.add('active-orange');
                    document.getElementById('pillOrange').style.display = 'inline-block';
                    bandStatusEl.className = 'valuation-status neutral';
                    bandStatusEl.textContent = 'Žlutooranžové pásmo (Férová hodnota)';
                } else if (pasma.aktualni_pasmo === 'ČERVENÉ') {
                    document.getElementById('bandBoxRed').classList.add('active-red');
                    document.getElementById('pillRed').style.display = 'inline-block';
                    bandStatusEl.className = 'valuation-status overvalued';
                    bandStatusEl.textContent = 'Červené pásmo (Nadhodnoceno)';
                } else {
                    bandStatusEl.className = 'valuation-status neutral';
                    bandStatusEl.textContent = 'Tržní cena nedostupná';
                }

                // Tržní data, SOTP, LBO
                document.getElementById('valCap').textContent = formatCurrency(data.trzni_kapitalizace);
                document.getElementById('valDebt').textContent = formatCurrency(data.celkovy_dluh);
                document.getElementById('valEV').textContent = formatCurrency(data.celkova_hodnota_podniku);

                document.getElementById('sotpEV').textContent = formatCurrency(data.metoda_souctu_casti.vypoctena_hodnota_podniku);
                document.getElementById('sotpEquity').textContent = formatCurrency(data.metoda_souctu_casti.vypoctena_hodnota_vlastniho_kapitalu);

                const lbo = data.model_odkupu_za_cizi_kapital_base_case;
                document.getElementById('lboInEq').textContent = formatCurrency(lbo.vstupni_vlastni_kapital);
                document.getElementById('lboOutEq').textContent = formatCurrency(lbo.vystupni_hodnota_vlastniho_kapitalu);
                document.getElementById('lboMoic').textContent = lbo.moic.toFixed(2) + 'x';
                document.getElementById('lboIrr').textContent = lbo.irr_procent.toFixed(2) + ' %';

                // Diskontované peněžní toky a Semafor
                const dcf = data.dcf_valuation;
                document.getElementById('dcfIntrinsicValue').textContent = formatCurrency(dcf.intrinsic_value_per_share);
                document.getElementById('dcfWacc').textContent = (dcf.wacc * 100).toFixed(2) + ' %';
                document.getElementById('dcfTerminalValue').textContent = formatCurrency(dcf.terminal_value);

                const statusEl = document.getElementById('dcfValuationStatus');
                const upside = dcf.upside_downside_percent;
                if (upside === null || upside === undefined) {
                    statusEl.className = 'valuation-status neutral';
                    statusEl.textContent = 'Cena neuvedena';
                } else {
                    statusEl.className = 'valuation-status ' + (upside >= 0 ? 'undervalued' : 'overvalued');
                    statusEl.textContent = (upside >= 0 ? 'Podhodnoceno o ' : 'Nadhodnoceno o ') + Math.abs(upside).toFixed(2) + ' %';
                }

                ['thesisGreen', 'thesisOrange', 'thesisRed'].forEach(id => document.getElementById(id).classList.remove('active'));
                let color = 'orange';
                if (upside !== null && upside >= 15) color = 'green';
                else if (upside !== null && upside <= -15) color = 'red';
                document.getElementById('thesis' + color.charAt(0).toUpperCase() + color.slice(1)).classList.add('active');
                document.getElementById('thesisRecommendation').textContent =
                    upside >= 15 ? 'Model indikuje podhodnocení akcie' :
                    upside <= -15 ? 'Model indikuje nadhodnocení akcie' :
                    'Model se nachází v pásmu férové hodnoty';
                document.getElementById('thesisSummary').textContent =
                    'Střední vnitřní hodnota ' + formatCurrency(dcf.intrinsic_value_per_share) +
                    ' proti aktuální tržní ceně ' + formatCurrency(dcf.current_share_price) +
                    '. Výsledek vychází z vážených nákladů kapitálu ' + (dcf.wacc * 100).toFixed(2) +
                    ' % a výstupního násobku ' + data.parametry_scenare.exit_mult + 'x.';

                // Srovnání konkurence
                document.getElementById('peerSector').textContent = data.peer_comps.sector || 'Sektor';
                let peerHtml = '<table><thead><tr><th>Metoda ocenění</th><th>Násobek</th><th>Odhadovaný ukazatel</th><th>Implikovaná cena / akcii</th></tr></thead><tbody>';
                (data.peer_comps.multiples || []).forEach(p => {
                    peerHtml += '<tr><td>' + p.method + '</td><td>' + p.multiple + 'x</td><td>' + formatCurrency(p.forward_metric) + '</td><td>' + formatCurrency(p.implied_share_price) + '</td></tr>';
                });
                peerHtml += '</tbody></table>';
                document.getElementById('peerCompsContainer').innerHTML = peerHtml;

                // Amortizační kalendář
                let schedHtml = '<table><thead><tr><th>Období</th><th>Provozní zisk před odpisy</th><th>Volný tok pro splácení</th><th>Splátka jistiny</th><th>Zůstatek dluhu</th></tr></thead><tbody>';
                (data.projektovane_roky || []).forEach(y => {
                    schedHtml += '<tr><td>Rok ' + y.rok + '</td><td>' + formatCurrency(y.projektovany_provozni_zisk_pred_odpisy) + '</td><td>' + formatCurrency(y.volny_tok_pro_splaceni) + '</td><td>' + formatCurrency(y.mor_jistiny) + '</td><td>' + formatCurrency(y.zbyvajici_cizi_kapital) + '</td></tr>';
                });
                schedHtml += '</tbody></table>';
                document.getElementById('lboScheduleContainer').innerHTML = schedHtml;

                // Citlivostní matice
                const matrix = data.citlivostni_matice_lbo_irr || [];
                const multiHeader = data.parametry_scenare.testovane_nasobky || [];
                let matHtml = '<table><thead><tr><th>Růst provozního zisku / Násobek</th>' + multiHeader.map(m => '<th>' + Number(m).toFixed(1) + 'x</th>').join('') + '</tr></thead><tbody>';
                matrix.forEach(row => {
                    matHtml += '<tr><td><strong>' + row.tempo_rustu_ebitda_procent + ' %</strong></td>';
                    row.vysledky_irr_dle_nasobku.forEach(cell => {
                        const colColor = cell.irr_procent >= 0 ? '#34d399' : '#f87171';
                        matHtml += '<td style="color: ' + colColor + ';">' + cell.irr_procent.toFixed(2) + ' %</td>';
                    });
                    matHtml += '</tr>';
                });
                matHtml += '</tbody></table>';
                document.getElementById('sensitivityMatrix').innerHTML = matHtml;
            } catch (err) {
                console.error(err);
                document.getElementById('thesisRecommendation').textContent = 'Chyba při načítání dat: ' + err.message;
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
                aiMsg.textContent = 'Chyba při komunikaci s poradcem.';
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

        async function openNewsletterModal() {
            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            document.getElementById('newsModalTicker').textContent = ticker;
            document.getElementById('newsletterStatusMsg').textContent = '';
            document.getElementById('newsletterModal').classList.add('open');
            await Promise.all([loadSubscribersList(), loadTickerNews(ticker)]);
        }

        function closeNewsletterModal() {
            document.getElementById('newsletterModal').classList.remove('open');
        }

        async function loadSubscribersList() {
            const container = document.getElementById('subscribersListContainer');
            try {
                const res = await fetch('/api/subscribers');
                const data = await res.json();
                const subs = data.subscribers || [];
                if (subs.length === 0) {
                    container.innerHTML = '<p class="plain-guide">Zatím nejsou evidovány žádné e-mailové adresy.</p>';
                    return;
                }
                container.innerHTML = subs.map(s =>
                    `<div class="subscriber-row">
                        <div><strong>${s.email}</strong> • sledovaný titul: <span style="color:#38bdf8;">${s.ticker}</span></div>
                        <button type="button" class="btn btn-pdf" style="padding: 4px 10px; font-size: 11px;" onclick="deleteSubscriber('${s.email}')">Odhlásit</button>
                    </div>`
                ).join('');
            } catch (err) {
                container.innerHTML = '<p class="plain-guide">Nepodařilo se načíst seznam odběratelů.</p>';
            }
        }

        async function deleteSubscriber(email) {
            const statusEl = document.getElementById('newsletterStatusMsg');
            try {
                const res = await fetch('/api/subscribers?email=' + encodeURIComponent(email), { method: 'DELETE' });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Chyba při odhlášení.');
                statusEl.style.color = '#34d399';
                statusEl.textContent = data.message;
                await loadSubscribersList();
            } catch (err) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = err.message;
            }
        }

        async function sendNewsletterOnce() {
            const emailInput = document.getElementById('newsletterEmailInput');
            const statusEl = document.getElementById('newsletterStatusMsg');
            const btn = document.getElementById('btnSendOneOff');
            const email = emailInput.value.trim();

            if (!email) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = 'Zadejte nejprve svou e-mailovou adresu.';
                return;
            }

            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            const leverage = document.getElementById('leverageInput').valueAsNumber / 100;
            const growth = document.getElementById('growthInput').valueAsNumber / 100;
            const exitMult = document.getElementById('exitMultipleInput').valueAsNumber;

            btn.disabled = true;
            statusEl.style.color = '#38bdf8';
            statusEl.textContent = 'Odesílám jednorázový valuační přehled...';

            try {
                const res = await fetch('/api/newsletter/send-now', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email, ticker, leverage, growth, exit_mult: exitMult })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Odeslání se nezdařilo.');
                statusEl.style.color = '#34d399';
                statusEl.textContent = data.message;
            } catch (err) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = err.message;
            } finally {
                btn.disabled = false;
            }
        }

        async function registerRegularNewsletter() {
            const emailInput = document.getElementById('newsletterEmailInput');
            const statusEl = document.getElementById('newsletterStatusMsg');
            const btn = document.getElementById('btnSubscribeRegular');
            const email = emailInput.value.trim();

            if (!email) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = 'Zadejte nejprve e-mailovou adresu pro zaevidování.';
                return;
            }

            const ticker = (document.getElementById('tickerInput').value || 'NVDA').trim().toUpperCase();
            btn.disabled = true;

            try {
                const res = await fetch('/subscribe', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email, ticker })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Chyba při evidenci odběru.');
                statusEl.style.color = '#34d399';
                statusEl.textContent = data.message;
                emailInput.value = '';
                await loadSubscribersList();
            } catch (err) {
                statusEl.style.color = '#f87171';
                statusEl.textContent = err.message;
            } finally {
                btn.disabled = false;
            }
        }

        async function loadTickerNews(ticker) {
            const listEl = document.getElementById('newsFeedList');
            listEl.innerHTML = '<div style="text-align:center;padding:20px;"><div class="loader"></div></div>';
            try {
                const res = await fetch('/api/news/' + encodeURIComponent(ticker));
                const data = await res.json();
                if (!data.articles || data.articles.length === 0) {
                    listEl.innerHTML = '<p class="plain-guide">Pro tento symbol nejsou momentálně dostupné žádné zprávy.</p>';
                    return;
                }
                listEl.innerHTML = data.articles.map(a =>
                    `<div class="news-item"><a href="${a.url}" target="_blank" rel="noopener">${a.title}</a><div class="news-meta">${a.source}</div></div>`
                ).join('');
            } catch (err) {
                listEl.innerHTML = '<p class="plain-guide">Nepodařilo se načíst tržní zprávy.</p>';
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            loadForexRates();
            loadValuation(true);
        });
    </script>
</body>
</html>
"""