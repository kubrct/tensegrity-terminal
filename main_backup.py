from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
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

app = FastAPI(title="Tensegrity Terminal API", version="12.1")

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
                </div>
            </header>

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
                </div>
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
        </div>

        <script>
            async function loadValuation() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                const container = document.getElementById('matrixContainer');
                container.innerHTML = '<div style="text-align: center; padding: 30px; color: var(--text-muted);"><div class="loader"></div><div style="margin-top: 10px;">Načítám data pro ' + ticker + '...</div></div>';

                try {
                    const response = await fetch(`/api/valuation/${ticker}`);
                    const data = await response.json();

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
                    container.innerHTML = '<div style="text-align: center; padding: 20px; color: #f87171;">Chyba při načítání dat: ' + error + '</div>';
                }
            }

            function downloadExcel() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                window.location.href = `/api/export/excel/${ticker}`;
            }

            function downloadPDF() {
                const ticker = document.getElementById('tickerInput').value.trim().toUpperCase();
                window.location.href = `/api/export/pdf/${ticker}`;
            }

            window.onload = loadValuation;
        </script>
    </body>
    </html>
    """

def run_valuation_logic(ticker: str):
    stock = yf.Ticker(ticker)
    info = stock.info
    financials = stock.financials

    current_price = info.get("currentPrice")
    shares_outstanding = info.get("sharesOutstanding")
    celkovy_dluh = info.get("totalDebt", 0)
    hotovost = info.get("totalCash", 0)
    
    trzni_kapitalizace = (current_price * shares_outstanding) if (current_price and shares_outstanding) else info.get("marketCap", 0)
    celkova_hodnota_podniku = trzni_kapitalizace + celkovy_dluh - hotovost

    ebitda = info.get("ebitda", 0)
    trzby = info.get("totalRevenue", 0)
    cisty_zisk = info.get("netIncomeToCommon", 0)

    try:
        pretax_income = financials.loc["Pretax Income"].iloc[0] if "Pretax Income" in financials.index else cisty_zisk * 1.25
        tax_provision = financials.loc["Tax Provision"].iloc[0] if "Tax Provision" in financials.index else cisty_zisk * 0.25
        efektivni_danova_sazba = max(0.0, min(0.40, tax_provision / pretax_income)) if pretax_income > 0 else 0.21
        
        raw_interest_rate = abs(financials.loc["Interest Expense"].iloc[0]) / celkovy_dluh if "Interest Expense" in financials.index and celkovy_dluh > 0 else 0.05
        implikovana_sazba_dluhu = max(0.03, min(0.07, raw_interest_rate))
    except Exception:
        efektivni_danova_sazba = 0.21
        implikovana_sazba_dluhu = 0.045

    hodnota_prvniho = (ebitda * 0.80) * 14.0
    hodnota_druheho = (ebitda * 0.20) * 9.0
    celkova_hodnota_podniku_sotp = hodnota_prvniho + hodnota_druheho
    hodnota_vlastniho_kapitalu_sotp = celkova_hodnota_podniku_sotp - celkovy_dluh + hotovost

    transakcni_dluh_procento = 0.30
    vstupni_cizi_kapital = celkova_hodnota_podniku * transakcni_dluh_procento
    vstupni_vlastni_kapital = celkova_hodnota_podniku - vstupni_cizi_kapital

    roční_rust_ebitda = 0.05
    aktualni_dluh = vstupni_cizi_kapital
    projektovane_roky = []
    bezkusova_ebitda = ebitda

    for rok in range(1, 6):
        bezkusova_ebitda *= (1 + roční_rust_ebitda)
        odpisy = trzby * 0.025
        provozni_zisk_ebit = bezkusova_ebitda - odpisy
        
        urokove_naklady = aktualni_dluh * implikovana_sazba_dluhu
        zisk_pred_danemi = max(0.0, provozni_zisk_ebit - urokove_naklady)
        dan_z_prijmu = zisk_pred_danemi * efektivni_danova_sazba
        
        investicni_vydaje = trzby * 0.035
        zmena_pracovniho_kapitalu = trzby * 0.01

        cisty_penezni_tok = provozni_zisk_ebit - dan_z_prijmu + odpisy - investicni_vydaje - zmena_pracovniho_kapitalu
        volny_tok_pro_splaceni = max(0.0, cisty_penezni_tok - urokove_naklady)

        splatka_jistiny = min(aktualni_dluh, volny_tok_pro_splaceni * 0.80)
        aktualni_dluh -= splatka_jistiny

        projektovane_roky.append({
            "rok": rok,
            "projektovany_provozni_zisk_pred_odpisy": round(bezkusova_ebitda, 2),
            "zbyvajici_cizi_kapital": round(aktualni_dluh, 2),
            "mor_jistiny": round(splatka_jistiny, 2)
        })

    vystupni_nasobek = 13.5
    vystupni_hodnota_podniku = bezkusova_ebitda * vystupni_nasobek
    vystupni_hodnota_vlastniho_kapitalu = vystupni_hodnota_podniku - aktualni_dluh

    moic = (vystupni_hodnota_vlastniho_kapitalu / vstupni_vlastni_kapital) if vstupni_vlastni_kapital > 0 else 0
    irr = ((moic ** (1 / 5)) - 1) * 100 if moic > 0 else 0

    citlivostni_matice_irr = []
    testovane_rustove_sazby = [0.03, 0.04, 0.05, 0.06, 0.07]
    testovane_nasobky = [11.0, 12.0, 13.5, 15.0, 16.5]

    for g in testovane_rustove_sazby:
        radek_matice = {"tempo_rustu_ebitda_procent": round(g * 100, 1), "vysledky_irr_dle_nasobku": []}
        ebitda_t5_g = ebitda * ((1 + g) ** 5)
        for mult in testovane_nasobky:
            ev_t5 = ebitda_t5_g * mult
            final_debt_sim = max(0.0, vstupni_cizi_kapital - (ebitda_t5_g * 0.15 * 5))
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
        "citlivostni_matice_lbo_irr": citlivostni_matice_irr
    }

@app.get("/api/valuation/{ticker}")
def get_institutional_terminal_valuation(ticker: str):
    try:
        return run_valuation_logic(ticker)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/export/excel/{ticker}")
def export_valuation_excel(ticker: str):
    try:
        data = run_valuation_logic(ticker)
        
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
def export_valuation_pdf(ticker: str):
    try:
        data = run_valuation_logic(ticker)
        
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

        # 5. CITLIVOSTNÍ MATICE IRR
        elements.append(Paragraph("3. Citlivostní Matice Vnitřní Míry Výnosnosti (IRR %)", section_style))
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