import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
from scipy.stats import linregress, norm

# Configuration de la page
st.set_page_config(
    page_title="Financial Dashboard Pro",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# PALETTE DE COULEURS ET CSS SUR MESURE
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    /* Arrière-plan global */
    .stApp {
        background-color: #F4F8F8;
        color: #3F5056;
        font-family: 'Inter', 'Segoe UI', sans-serif;
    }

    /* Sidebar sombre moderne */
    section[data-testid="stSidebar"] {
        background-color: #E8EDEF !important;
    }
    /* Titres et labels de la sidebar en blanc/clair */
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label {
        color: #147394 !important;
        font-weight: 600;
    }

    /* 🟢 CORRECTION : Texte à l'intérieur des boîtes de saisie et listes déroulantes */
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] div[role="combobox"],
    section[data-testid="stSidebar"] div[data-baseweb="select"] * {
        color: #3F5056 !important;
        background-color: #ffffff !important;
    }

    /* Cartes de synthèse style "UI Cards" */
    .metric-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
        border: 1px solid #E2ECEB;
        margin-bottom: 15px;
    }

    .card-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #147394;
        margin-bottom: 15px;
        padding-bottom: 8px;
        border-bottom: 2px solid #8ACABB;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .metric-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 8px 0;
        border-bottom: 1px dashed #EEF4F4;
        font-size: 0.90rem;
    }
    .metric-row:last-child {
        border-bottom: none;
    }

    .metric-label {
        color: #3F5056;
        font-weight: 500;
    }

    .metric-value {
        font-weight: 700;
        color: #147394;
        font-family: 'Consolas', monospace;
    }

    /* Badges */
    .badge {
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
        display: inline-block;
    }
    .badge-green { background-color: #E6F7F2; color: #00915A; }
    .badge-orange { background-color: #FFF4E6; color: #D97706; }
    .badge-red { background-color: #FDE8E8; color: #E11D48; }K

    /* En-tête du Dashboard */
    .dashboard-header {
        background: linear-gradient(135deg, #147394 0%, #009FB1 100%);
        padding: 20px 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(20, 115, 148, 0.15);
    }
    .dashboard-header h1 {
        color: white !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
    }
    .dashboard-header p {
        color: #ADD9D3;
        margin: 5px 0 0 0;
        font-size: 0.95rem;
    }
    /* En-tête du Dashboard avec cours à droite */
    .dashboard-header {
        background: linear-gradient(135deg, #147394 0%, #009FB1 100%);
        padding: 20px 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(20, 115, 148, 0.15);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .dashboard-header-left h1 {
        color: white !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 500;
    }
    .dashboard-header-left p {
        color: #ADD9D3;
        margin: 5px 0 0 0;
        font-size: 0.95rem;
    }
    .dashboard-header-right {
        text-align: right;
        background: rgba(255, 255, 255, 0.15);
        padding: 10px 18px;
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    .dashboard-header-right .price-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        color: #ffffff;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .dashboard-header-right .price-value {
        font-size: 1,7rem;
        font-weight: 800;
        color: #ffffff;
        font-family: 'Consolas', monospace;
    }
    .dashboard-header-right .price-change {
        font-size: 1rem;
        font-weight: 700;
        margin-left: 8px;
    }
        .dashboard-header-right .target_str {
        font-size: 0.75rem;
        text-transform: uppercase;
        color: #ffffff;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1. BARRE LATÉRALE - SÉLECTION DU TICKER (3 MODES)
# -----------------------------------------------------------------------------
st.sidebar.markdown("### SELECTION TICKER")

# 1. Liste prédéfinie : Portefeuille
tickers_portfolio = {
    "Air Liquide": "AI.PA",
    "Crédit Agricole": "ACA.PA",
    "Eli Lilly": "LLY",
    "Investor AB": "INVE-B.ST",
    "Mauna Kea Techologies": "ALMKT.PA",
    "Novo Nordisk": "NOV.DE",
    "Microsoft":"MSFT",
    "TotalEnergies": "TTE.PA",
    "Visa": "V",
    "Amundi MSCI World": "DCAM.PA",
    "Amundi Nasdaq 100": "UST.PA",
    "BNP Stox 600": "ETZ.PA",
    "Ishares Semiconductor": "SEME.PA",
    "Bitcoin": "BTC-EUR",
    "Solana": "SOL-EUR",
}


# 2. Chargement de la liste globale (ticker.xlsx)
@st.cache_data
def charger_liste_tickers(fichier_excel="ticker.xlsx"):
    try:
        df_tickers = pd.read_excel(fichier_excel)
        df_tickers = df_tickers.dropna(subset=["Name", "full ticker"])
        return dict(
            zip(df_tickers["Name"].astype(str), df_tickers["full ticker"].astype(str))
        )
    except Exception:
        return {}

# Purge uniquement le cache de la fonction qui lit l'Excel
charger_liste_tickers.clear()

# Puis chargement normal
tickers_excel = charger_liste_tickers("ticker.xlsx")

# --- COMPOSANTS DE SÉLECTION ---

# 1. Liste déroulante : Portefeuille
choix_portefeuille = st.sidebar.selectbox(
    "Tickers en portefeuille :",
    options=["-- Aucun --"] + list(tickers_portfolio.keys()),
    index=0,
)

# 2. Liste déroulante : Dictionnaire global (ticker.xlsx)
choix_excel = st.sidebar.selectbox(
    "Actions à selectionner :",
    options=["-- Aucun --"] + sorted(list(tickers_excel.keys())),
    index=0,
)

# Determination de la valeur par défaut du champ manuel selon la sélection
ticker_suggere = ""
if choix_portefeuille != "-- Aucun --":
    ticker_suggere = tickers_portfolio[choix_portefeuille]
elif choix_excel != "-- Aucun --":
    ticker_suggere = tickers_excel[choix_excel]
else:
    ticker_suggere = "AI.PA"  # Valeur par défaut si aucun filtre n'est sélectionné

# 3. Zone de recherche / saisie manuelle
ticker_input = st.sidebar.text_input(
    "Rechercher Action / ETF :", value=ticker_suggere
)

# -----------------------------------------------------------------------------
# 2. CHARGEMENT ET TRAITEMENT DES DONNÉES
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)
def charger_donnees(symbol):
    ticker = yf.Ticker(symbol)
    df = ticker.history(period="10y")
    if df.empty:
        return None, None
    try:
        info = ticker.info
    except Exception:
        info = {}
    return df, info

df, info = charger_donnees(ticker_input)

if df is None or df.empty:
    st.error(f"Impossible de charger les données pour le symbole '{ticker_input}'.")
    st.stop()


# Récupérer l'ISIN
ticker_isin = yf.Ticker(ticker_input)
isin_code = ticker_isin.isin

from urllib.parse import urlparse
import streamlit as st

# 1. Détails de la compagnie via Yahoo Finance
stock = yf.Ticker(ticker_input)
company_name = stock.info.get("shortName", "N/A")
sector = stock.info.get("sector", "N/A")
website = stock.info.get("website", "")
currency = stock.info.get("currency", "N/A")
exchange = stock.info.get("exchange", "N/A")
api_token = "pk_cgCfOY41ROGNIKMC3E0Gmg"


# 2. Fonction pour nettoyer l'URL et la réduire au domaine principal
def extract_domain(url: str) -> str:
  if not url:
    return ""
  if not url.startswith(("http://", "https://")):
    url = "https://" + url
  domain = urlparse(url).netloc
  return domain.replace("www.", "")


# 3. Fonction pour générer l'URL du logo
def get_logo_url(domain: str, token: str) -> str:
  if not domain:
    return ""
  return f"https://img.logo.dev/{domain}?token={token}"


# 4. Traitement
domain_name = extract_domain(website)
logo_url = get_logo_url(domain_name, api_token)


# -----------------------------------------------------------------------------
# 3. CALCULS MATHÉMATIQUES ET FINANCIERS
# -----------------------------------------------------------------------------
# Mappage de la devise vers son symbole
SYMBOLES_DEVISES = {
    "EUR": "€",
    "USD": "$",
    "GBP": "£",
    "GBp": "p",  # Pence britannique
    "CHF": "CHF",
    "CAD": "C$",
    "JPY": "¥",
}

# Extraction de la devise via yfinance (avec "EUR" par défaut)
code_devise = info.get("currency", "EUR")
symbole_devise = SYMBOLES_DEVISES.get(code_devise, code_devise)

# Fonction de formatage des nombres avec suffixe
def fmt(value, decimals, suffix=""):
    if value is None:
        return "N/A"
    return f"{value:,.{decimals}f}{suffix}".replace(",", " ").replace(".", ",")

df = df[['Close']].dropna()
df['Time_Idx'] = np.arange(len(df))

log_prices = np.log(df['Close'].values)
slope, intercept, r_value, p_value, std_err = linregress(df['Time_Idx'].values, log_prices)

df['Log_Pred'] = intercept + slope * df['Time_Idx']
df['Theoretical'] = np.exp(df['Log_Pred'])
residuals = log_prices - df['Log_Pred']
std_dev = np.std(residuals)

df['+1std'] = np.exp(df['Log_Pred'] + std_dev)
df['-1std'] = np.exp(df['Log_Pred'] - std_dev)
df['+2std'] = np.exp(df['Log_Pred'] + 2 * std_dev)
df['-2std'] = np.exp(df['Log_Pred'] - 2 * std_dev)

cours_actuel = df['Close'].iloc[-1]
valeur_theorique = df['Theoretical'].iloc[-1]
ecart_regression = ((valeur_theorique - cours_actuel) / valeur_theorique) * 100
badge_class = "badge-red" if ecart_regression < 0 else "badge-green"
position_sigma = (np.log(cours_actuel) - df['Log_Pred'].iloc[-1]) / std_dev

nb_jours_projection = 1260
derniere_date = df.index[-1]
futures_dates = [derniere_date + timedelta(days=i*1.44) for i in range(1, nb_jours_projection + 1)]

index_futurs = np.arange(len(df), len(df) + nb_jours_projection)
log_pred_futur = intercept + slope * index_futurs
valeur_prevue_5a = np.exp(log_pred_futur[-1])
date_cible_5a = futures_dates[-1]

cagr = (np.exp(slope * 252) - 1) * 100
perf_12m = ((cours_actuel / df['Close'].iloc[-252]) - 1) * 100 if len(df) >= 252 else 0.0
perf_5a = ((cours_actuel / df['Close'].iloc[-1260]) - 1) * 100 if len(df) >= 1260 else 0.0

dividende_annuel = info.get('dividendRate', 0.0) or 0.0
rendement_div = (dividende_annuel / cours_actuel * 100) if (dividende_annuel > 0 and cours_actuel > 0) else 0.0

dernier_div = info.get('lastDividendValue', 0.0) or 0.0
date_dernier_div = info.get('exDividendDate', None)
date_dernier_div = datetime.fromtimestamp(date_dernier_div).strftime('%d/%m/%Y') if date_dernier_div else "N/A"

# Payout Ratio & Badge
payout_ratio = info.get('payoutRatio', None)
if payout_ratio is not None and payout_ratio > 0:
    payout_pct = payout_ratio * 100
    if payout_pct < 50:
        payout_badge = "<span class='badge badge-green'>★★★★★</span>"
    elif 50 <= payout_pct <= 75:
        payout_badge = "<span class='badge badge-green'>★★★★</span>"
    elif 75 < payout_pct <= 100:
        payout_badge = "<span class='badge badge-orange'>★★★</span>"
    else:
        payout_badge = "<span class='badge badge-red'>★★</span>"
    payout_ratio_str = f"{payout_pct:.1f}% {payout_badge}"
elif payout_ratio == 0:
    payout_ratio_str = "0.0% <span class='badge badge-red'>★</span>" # Cas où payoutRatio est 0
else:
    payout_ratio_str = "N/A"

# Risque 6M & Badge
log_returns = np.log(df['Close'] / df['Close'].shift(1)).dropna()
volatilite_annualisee = log_returns.std() * np.sqrt(252)
risque_6m = volatilite_annualisee * np.sqrt(0.5) * 100

if risque_6m < 15:
    risque_badge = "<span class='badge badge-green'>★★★★</span>"
elif 15 <= risque_6m < 30:
    risque_badge = "<span class='badge badge-green'>★★★</span>"
elif 30 <= risque_6m < 45:
    risque_badge = "<span class='badge badge-orange'>★★</span>"
else:
    risque_badge = "<span class='badge badge-red'>★</span>"
risque_6m_str = f"{risque_6m:.1f}% {risque_badge}"

# --- VOLATILITÉ ET MODÈLE STOCHASTIQUE BLACK-SCHOLES-MERTON (BSM) ---
log_returns = np.log(df['Close'] / df['Close'].shift(1)).dropna()
volatilite_annualisee = log_returns.std() * np.sqrt(252)
sigma_bsm = volatilite_annualisee

# Risque classique 6M
risque_6m = volatilite_annualisee * np.sqrt(0.5) * 100
if risque_6m < 15: risque_badge = "<span class='badge badge-green'>★★★★</span>"
elif 15 <= risque_6m < 30: risque_badge = "<span class='badge badge-green'>★★★</span>"
elif 30 <= risque_6m < 45: risque_badge = "<span class='badge badge-orange'>★★</span>"
else: risque_badge = "<span class='badge badge-red'>★</span>"
risque_6m_str = f"{risque_6m:.1f}% {risque_badge}"

# Calcul du Risque BSM (Probabilité de chute de 20% à 6M)
T_6m = 0.5  # Horizon 6M
r_f = 0.03  # Taux sans risque (3%)
seuil_baisse = 0.20
S_cible = cours_actuel * (1 - seuil_baisse)

if sigma_bsm > 0 and cours_actuel > 0:
    d2 = (np.log(cours_actuel / S_cible) + (r_f - 0.5 * sigma_bsm**2) * T_6m) / (sigma_bsm * np.sqrt(T_6m))
    prob_baisse_bsm = norm.cdf(-d2) * 100
else:
    prob_baisse_bsm = 0.0

if prob_baisse_bsm < 10: bsm_badge = "<span class='badge badge-green'>★★★★</span>"
elif 10 <= prob_baisse_bsm < 20: bsm_badge = "<span class='badge badge-green'>★★★</span>"
elif 20 <= prob_baisse_bsm < 35: bsm_badge = "<span class='badge badge-orange'>★★</span>"
else: bsm_badge = "<span class='badge badge-red'>★</span>"

bsm_risk_str = f"{prob_baisse_bsm:.1f}% {bsm_badge}"

# Forward P/E & Badge
forward_pe = info.get('forwardPE', None)
if forward_pe is not None and forward_pe > 0:
    if forward_pe < 10:
        pe_badge = "<span class='badge badge-green'>★★★★</span>"
    elif 10 <= forward_pe <= 17:
        pe_badge = "<span class='badge badge-green'>★★★</span>"
    elif 17 < forward_pe <= 25:
        pe_badge = "<span class='badge badge-orange'>★★</span>"
    else:
        pe_badge = "<span class='badge badge-red'>★</span>"
    forward_pe_str = f"x {forward_pe:.1f} {pe_badge}"
else:
    forward_pe_str = "N/A"

    S_cible = cours_actuel * (1 - seuil_baisse)

if sigma_bsm > 0:
    d2 = (
        np.log(cours_actuel / S_cible) + (r_f - 0.5 * sigma_bsm**2) * T_6m
    ) / (sigma_bsm * np.sqrt(T_6m))
    prob_baisse_bsm = norm.cdf(-d2) * 100
else:
    prob_baisse_bsm = 0.0

# Système de notation par étoiles (★)
if prob_baisse_bsm < 10:
    bsm_badge = "<span class='badge badge-green'>★★★★</span>"
elif 10 <= prob_baisse_bsm < 20:
    bsm_badge = "<span class='badge badge-green'>★★★</span>"
elif 20 <= prob_baisse_bsm < 35:
    bsm_badge = "<span class='badge badge-orange'>★★</span>"
else:
    bsm_badge = "<span class='badge badge-red'>★</span>"

bsm_risk_str = f"{prob_baisse_bsm:.1f}% {bsm_badge}"

# -----------------------------------------------------------------------------
# CALCULS : SANTÉ FINANCIÈRE, RENTABILITÉ & SCORE GLOBAL
# -----------------------------------------------------------------------------

# 1. Extraction des métriques fondamentales depuis yfinance
ebitda = info.get('ebitda', None)
total_cash = info.get('totalCash', 0.0) or 0.0
total_debt = info.get('totalDebt', 0.0) or 0.0

# Net Debt / EBITDA
if ebitda and ebitda > 0:
    net_debt = total_debt - total_cash
    net_debt_ebitda = net_debt / ebitda
    net_debt_ebitda_str = f"{net_debt_ebitda:.2f}x"
else:
    net_debt_ebitda = None
    net_debt_ebitda_str = "N/A"

# Current Ratio
current_ratio = info.get('currentRatio', None)
current_ratio_str = f"{current_ratio:.2f}" if current_ratio else "N/A"

# ROE (Return on Equity)
roe = info.get('returnOnEquity', None)
roe_pct = roe * 100 if roe else None
roe_str = f"{roe_pct:.1f}%" if roe_pct else "N/A"

# ROIC (Return on Invested Capital)
# yfinance ne fournissant pas directement le ROIC dans info, nous l'estimons via ROA ou EBIT/(Capital Investi)
roa = info.get('returnOnAssets', None)
roic_pct = (roa * 100 * 1.3) if roa else (roe_pct * 0.75 if roe_pct else None) # Approximation robuste
roic_str = f"{roic_pct:.1f}%" if roic_pct else "N/A"

# PEG Ratio
peg_ratio = info.get('pegRatio', None)
peg_str = f"{peg_ratio:.2f}" if peg_ratio else "N/A"

# 2. Algorithme de calcul du Score de Santé Financière (sur 100 pts)
score_pts = 0
max_pts = 0

# Test Endettement (Net Debt / EBITDA) - Poids : 25 pts
if net_debt_ebitda is not None:
    max_pts += 25
    if net_debt_ebitda <= 1.5: score_pts += 25
    elif net_debt_ebitda <= 3.0: score_pts += 15
    elif net_debt_ebitda <= 4.5: score_pts += 5

# Test Liquidité (Current Ratio) - Poids : 20 pts
if current_ratio is not None:
    max_pts += 20
    if current_ratio >= 1.5: score_pts += 20
    elif current_ratio >= 1.0: score_pts += 12
    else: score_pts += 4

# Test Rentabilité (ROE) - Poids : 20 pts
if roe_pct is not None:
    max_pts += 20
    if roe_pct >= 15: score_pts += 20
    elif roe_pct >= 8: score_pts += 12
    elif roe_pct > 0: score_pts += 5

# Test Rentabilité des Investissements (ROIC) - Poids : 20 pts
if roic_pct is not None:
    max_pts += 20
    if roic_pct >= 12: score_pts += 20
    elif roic_pct >= 6: score_pts += 12
    elif roic_pct > 0: score_pts += 5

# Test Valorisation / Croissance (PEG) - Poids : 15 pts
if peg_ratio is not None and peg_ratio > 0:
    max_pts += 15
    if peg_ratio <= 1.0: score_pts += 15
    elif peg_ratio <= 2.0: score_pts += 8
    else: score_pts += 3

# Normalisation du score final sur 100
score_sante = int((score_pts / max_pts) * 100) if max_pts > 0 else 50

# Fonction de calcul du RSI
def calculer_rsi(series, fenetre=14):
    """Calcule l'indicateur RSI (Relative Strength Index) sur 14 périodes."""
    delta = series.diff()
    gain = delta.where(delta > 0, 0.0)
    perte = -delta.where(delta < 0, 0.0)

    # Lissage exponentiel (méthode de Wilder)
    avg_gain = gain.ewm(alpha=1 / fenetre, adjust=False).mean()
    avg_perte = perte.ewm(alpha=1 / fenetre, adjust=False).mean()

    rs = avg_gain / avg_perte
    rsi = 100.0 - (100.0 / (1.0 + rs))
    return rsi

# Calcul de la colonne RSI (sur les 14 derniers jours)
df["RSI"] = calculer_rsi(df["Close"], fenetre=14)
rsi_actuel = df["RSI"].iloc[-1]

    

# -----------------------------------------------------------------------------
# 4. EN-TÊTE DU DASHBOARD & CARTES UI DE SYNTHÈSE
# -----------------------------------------------------------------------------

# --- A. DÉFINITION DE LA FONCTION DU SCORE & SVG (DOIT ÊTRE PLACÉE EN PREMIER) ---
def generer_svg_potentiel(
    sigma, perf_12m, cagr, pe_fwd, cours_actuel, target_price
):
    # --- 1. CALCULS DES SCORES INDIVIDUELS ---

    # A. Sigma (Prix vs Valorisation Théorique)
    score_sigma = (
        np.clip((1 - (sigma / 3.0)) * 50, 0, 100) if sigma is not None else 50
    )

    # B. P/E Forward (12M)
    pe_val = float(pe_fwd) if pe_fwd and pe_fwd != "N/A" else 20
    score_pe = np.clip((1 - (pe_val - 5) / 35) * 100, 0, 100)

    # C. CAGR
    cagr_val = cagr if cagr is not None else 0
    score_cagr = np.clip(((cagr_val + 0.10) / 0.40) * 100, 0, 100)

    # D. Perf 12M
    perf_val = perf_12m if perf_12m is not None else 0
    score_perf = np.clip(((perf_val + 0.30) / 0.80) * 100, 0, 100)

    # E. Delta Val (Écart Cours Actuel vs Target Analystes)
    if target_price and target_price > 0 and cours_actuel and cours_actuel > 0:
        potentiel_pct = (target_price - cours_actuel) / cours_actuel
        score_delta = np.clip(50 + (potentiel_pct * 100), 0, 120)
    else:
        delta_val = 50
        score_delta = delta_val

    # --- 2. CALCUL DU SCORE GLOBAL PONDÉRÉ (TOTAL = 1.00) ---
    score_brut = (
        0.35 * score_sigma
        + 0.15 * score_pe
        + 0.25 * score_cagr
        + 0.15 * score_perf
        + 0.10 * score_delta
    )

    # Normalisation finale entre 0 et 100 sans décimale
    score = int(round(np.clip(score_brut, 0, 100), 0))

    # 2. Couleur du badge selon l'intervalle
    if score >= 80:
        bg_badge = "#3F5056"
    elif score >= 60:
        bg_badge = "#147394"
    elif score >= 40:
        bg_badge = "#009FB1"
    elif score >= 20:
        bg_badge = "#51BBB4"
    else:
        bg_badge = "#ADD9D3"

    pos_x = score
  
    svg_code = f'''<div style="display:flex;justify-content:flex-end;width:100%;">
<svg viewBox="0 0 320 85" width="280" height="75" style="overflow:visible;">
<defs><clipPath id="bar-clip"><rect x="0" y="45" width="320" height="18" rx="9" ry="9"/></clipPath></defs>
<g clip-path="url(#bar-clip)">
<rect x="0" y="45" width="64" height="18" fill="#ADD9D3"/>
<rect x="64" y="45" width="64" height="18" fill="#51BBB4"/>
<rect x="128" y="45" width="64" height="18" fill="#009FB1"/>
<rect x="192" y="45" width="64" height="18" fill="#147394"/>
<rect x="256" y="45" width="64" height="18" fill="#3F5056"/>
<line x1="64" y1="45" x2="64" y2="63" stroke="#fff" stroke-width="2"/>
<line x1="128" y1="45" x2="128" y2="63" stroke="#fff" stroke-width="2"/>
<line x1="192" y1="45" x2="192" y2="63" stroke="#fff" stroke-width="2"/>
<line x1="256" y1="45" x2="256" y2="63" stroke="#fff" stroke-width="2"/>
</g>
<text x="32" y="75" font-size="10" font-weight="bold" fill="#3F5056" text-anchor="middle">VENDRE</text>
<text x="96" y="75" font-size="10" font-weight="bold" fill="#3F5056" text-anchor="middle">ALLEGER</text>
<text x="160" y="75" font-size="10" font-weight="bold" fill="#3F5056" text-anchor="middle">CONSERVER</text>
<text x="224" y="75" font-size="10" font-weight="bold" fill="#3F5056" text-anchor="middle">RENFORCER</text>
<text x="288" y="75" font-size="10" font-weight="bold" fill="#3F5056" text-anchor="middle">ACHETER</text>
<g transform="translate({(pos_x / 100) * 320}, 0)">
<rect x="-24" y="0" width="48" height="28" rx="6" ry="6" fill="{bg_badge}" stroke="#ffffff" stroke-width="1.5"/>
<text x="0" y="19" font-size="14" font-weight="bold" fill="#ffffff" text-anchor="middle">{score}</text>
<polygon points="0,35 -5,28 5,28" fill="{bg_badge}"/>
<circle cx="0" cy="54" r="3.5" fill="#ffffff" stroke="{bg_badge}" stroke-width="2"/>
</g>
</svg>
</div>'''
    return svg_code


# --- B. CALCULS PRÉALABLES POUR L'EN-TÊTE ---
if len(df) >= 2:
    cours_precedents = df['Close'].iloc[-2]
    variation_jour_pct = ((cours_actuel - cours_precedents) / cours_precedents) * 100
else:
    variation_jour_pct = 0.0

couleur_var_header = "#8ACABB" if variation_jour_pct >= 0 else "#F59D7D"
signe_var = "+" if variation_jour_pct >= 0 else ""

target_price = info.get('targetMeanPrice', None)

if target_price and cours_actuel > 0:
    potentiel_target = ((target_price - cours_actuel) / cours_actuel) * 100
    signe_target = "+" if potentiel_target >= 0 else ""
    target_str = f"Objectif Analystes (1 an) : {fmt(target_price, 2, symbole_devise)} ({signe_target}{fmt(potentiel_target, 1, '%')})"
else:
    target_str = "Objectif Analystes (1 an) : Non disponible"

try:
    sigma_cours = (cours_actuel - valeur_theorique) / df['Close'].std()
except Exception:
    sigma_cours = 0.0

pe_fwd_val = info.get('forwardPE', None)


# --- C. GÉNÉRATION DE L'INDICATEUR & AFFICHAGE DANS STREAMLIT ---
# Appel mis à jour avec les 6 arguments requis
svg_indicator_html = generer_svg_potentiel(
    sigma=sigma_cours,
    perf_12m=perf_12m if "perf_12m" in locals() else 0,
    cagr=cagr if "cagr" in locals() else 0,
    pe_fwd=pe_fwd_val,
    cours_actuel=cours_actuel,
    target_price=target_price,
)
st.markdown(
    f"""<div style="display: flex; justify-content: flex-end; align-items: center; background-color: rgba(255, 255, 255, 0.05); padding: 8px 20px; border-radius: 10px; margin-bottom: 10px;">
    {svg_indicator_html}
</div>""",
    unsafe_allow_html=True,
)
st.write("")

# Génération du logo avec lien cliquable vers le site web (domain_name)
if logo_url and domain_name:
  website_url = (
      domain_name
      if domain_name.startswith("http")
      else f"https://{domain_name}"
  )
  logo_html = f'<a href="{website_url}" target="_blank"><img src="{logo_url}" style="height: 80px; width: 80px; object-fit: cover; border-radius: 50%; cursor: pointer;" alt="Logo"></a>'
elif logo_url:
  logo_html = f'<img src="{logo_url}" style="height: 80px; width: 80px; object-fit: cover; border-radius: 50%;" alt="Logo">'
else:
  initiales = ticker_input[:2].upper()
  logo_html = f'<div style="height: 80px; width: 80px; border-radius: 50%; background-color: rgba(255, 255, 255, 0.2); display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 1.5rem; color: white;">{initiales}</div>'

# Affichage de l'en-tête principal
st.markdown(
    f"""<div class="dashboard-header" style="display: flex; justify-content: space-between; align-items: center; padding: 15px 20px;">
    <div class="dashboard-header-left" style="display: flex; align-items: center; gap: 15px;">
        {logo_html}
        <div>
            <h2 style="margin: 0 0 -15px 0; line-height: 1.0;">{info.get('shortName', ticker_input)} <span style="font-size: 1.2rem; opacity: 0.9;"></span></h2>
            <div style="display: flex; gap: 8px; margin-top: 6px;">
                <span style="background-color: rgba(255, 255, 255, 0.15); padding: 4px 8px; border-radius: 6px; font-size: 0.85rem;">
                    <strong>ISIN :</strong> {isin_code}
                </span>
                <span style="background-color: rgba(255, 255, 255, 0.15); padding: 4px 8px; border-radius: 6px; font-size: 0.85rem;">
                    <strong>Secteur :</strong> {sector}
                </span>
                <span style="background-color: rgba(255, 255, 255, 0.15); padding: 4px 8px; border-radius: 6px; font-size: 0.85rem;">
                    <strong>Bourse :</strong> {exchange} ({currency})
                </span>
            </div>
        </div>
    </div>
    <div class="dashboard-header-right" style="text-align: right;">
        <div class="price-label">Cours Actuel</div>
        <div>
            <span class="price-value" style="font-size: 2.2rem; font-weight: bold;">{fmt(cours_actuel, 2, symbole_devise)}</span>
            <br>
            <span class="price-change" style="color: {couleur_var_header}; font-weight: bold; font-size: 1.1rem;">({signe_var}{fmt(variation_jour_pct, 2, '%')})</span>
        </div>
        <div class="target_str">{target_str}</div>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 0.9rem;"></p>
    </div>
</div>""",
    unsafe_allow_html=True,
)

# --- D. CARTES D'ANALYSE ---
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="card-title">📈 Données Actuelles</div>
        <div class="metric-row"><span class="metric-label">Valeur théorique</span><span class="metric-value">{fmt(valeur_theorique, 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Écart vs régression</span><span class="metric-value"><span class="badge {badge_class}">{fmt(ecart_regression, 2, '%')}</span></span></div>
        <div class="metric-row"><span class="metric-label">Rendement div.</span><span class="metric-value">{fmt(rendement_div, 2, '%')}</span></div>
        <div class="metric-row"><span class="metric-label">Payout Ratio</span><span>{payout_ratio_str}</span></div>
        <div class="metric-row"><span class="metric-label">Div. annuel</span><span class="metric-value">{fmt(dividende_annuel, 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Dernier div.</span><span class="metric-value">{fmt(dernier_div, 2, symbole_devise)}</span></div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="card-title">📊 Analyse Technique</div>
        <div class="metric-row"><span class="metric-label">CAGR (10 ans)</span><span class="metric-value">{fmt(cagr, 2, '%')}</span></div>
        <div class="metric-row"><span class="metric-label">Borne +2σ</span><span class="metric-value">{fmt(df['+2std'].iloc[-1], 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Borne +1σ</span><span class="metric-value">{fmt(df['+1std'].iloc[-1], 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Borne -1σ</span><span class="metric-value">{fmt(df['-1std'].iloc[-1], 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Borne -2σ</span><span class="metric-value">{fmt(df['-2std'].iloc[-1], 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">Risque 6M</span><span>{risque_6m_str}</span></div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="card-title">📅 Prévisions & Perf.</div>
        <div class="metric-row"><span class="metric-label">Perf. 12M</span><span class="metric-value">{fmt(perf_12m, 2, '%')}</span></div>
        <div class="metric-row"><span class="metric-label">Perf. 5A</span><span class="metric-value">{fmt(perf_5a, 2, '%')}</span></div>
        <div class="metric-row"><span class="metric-label">Prév. (5A)</span><span class="metric-value">{fmt(valeur_prevue_5a, 2, symbole_devise)}</span></div>
        <div class="metric-row"><span class="metric-label">P/E (12M fwd)</span><span>{forward_pe_str}</span></div>
        <div class="metric-row"><span class="metric-label">Volatilité BSM</span><span class="metric-value">{fmt(sigma_bsm * 100, 1, '%')}</span></div>
        <div class="metric-row"><span class="metric-label">Prob. Baisse 6M</span><span>{bsm_risk_str}</span></div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    couleur_score = "#00915A" if score_sante >= 70 else ("#009FB1" if score_sante >= 45 else "#E11D48")

    st.markdown(f"""
    <div class="metric-card">
        <div class="card-title">🏥 Santé & Rentabilité</div>
        <div class="metric-row"><span class="metric-label">Net Debt / EBITDA</span><span class="metric-value">{net_debt_ebitda_str}</span></div>
        <div class="metric-row"><span class="metric-label">Current Ratio</span><span class="metric-value">{current_ratio_str}</span></div>
        <div class="metric-row"><span class="metric-label">ROE</span><span class="metric-value">{roe_str}</span></div>
        <div class="metric-row"><span class="metric-label">ROIC (est.)</span><span class="metric-value">{roic_str}</span></div>
        <div class="metric-row"><span class="metric-label">PEG Ratio</span><span class="metric-value">{peg_str}</span></div>
        <div class="metric-row" style="border-bottom: none; flex-direction: column; align-items: stretch; gap: 3px; padding-top: 4px; padding-bottom: 0px;">
            <div style="display: flex; justify-content: space-between; font-size: 0.90rem;">
                <span class="metric-label">Score global</span>
                <span class="metric-value" style="color: {couleur_score};">{score_sante} %</span>
            </div>
            <div style="width: 100%; background-color: #EEF4F4; border-radius: 6px; height: 10px; overflow: hidden; display: flex; align-items: center;">
                <div style="width: {score_sante}%; background-color: {couleur_score}; height: 100%; border-radius: 6px; transition: width 0.5s;"></div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# -----------------------------------------------------------------------------
# 5. JAUGES PERSONNALISÉES AVEC LA NOUVELLE PALETTE
# -----------------------------------------------------------------------------
gauge_cols = st.columns(5)

def creer_jauge_custom(titre, valeur, min_val, max_val, suffixe=""):
    FONT_FAMILY = "Source Sans Pro, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif"

    # 🟢 Sécurité : Si la valeur est None ou NaN, on force 0 pour éviter de casser Plotly
    if valeur is None or (isinstance(valeur, float) and np.isnan(valeur)):
        valeur_safe = 0.0
    else:
        valeur_safe = float(valeur)

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=valeur_safe,
        number={
            'suffix': suffixe,
            'valueformat': '.2f' if 'σ' in suffixe else '.1f',
            'font': {'color': '#147394', 'size': 20, 'family': FONT_FAMILY}
        },
        title={
            'text': titre,
            'font': {'size': 13, 'color': '#3F5056', 'family': FONT_FAMILY}
        },
        gauge={
            'axis': {
                'range': [min_val, max_val],
                'tickcolor': "#3F5056",
                'tickfont': {'family': FONT_FAMILY, 'size': 10, 'color': '#3F5056'}
            },
            'bar': {'color': "#147394", 'thickness': 0.3},
            'bgcolor': "#ffffff",
            'borderwidth': 0,
            'steps': [
                {'range': [min_val, min_val + (max_val-min_val)*0.5], 'color': "#ADD9D3"},
                {'range': [min_val + (max_val-min_val)*0.5, max_val], 'color': "#51BBB4"}
            ],
        }
    ))

    fig.update_layout(
        height=210,
        margin=dict(l=20, r=20, t=50, b=10),
        separators=', ',  # Virgule pour les décimales
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family=FONT_FAMILY)
    )
    return fig

# 🟢 Affichage des 5 jauges en dehors de la fonction
with gauge_cols[0]:
    st.plotly_chart(creer_jauge_custom("CAGR 10A(%)", cagr, -10, 30, "%"), use_container_width=True)

with gauge_cols[1]:
    st.plotly_chart(creer_jauge_custom("R² (Qualité)", r_value**2, 0, 1), use_container_width=True)

with gauge_cols[2]:
    st.plotly_chart(creer_jauge_custom("Perf. 12M (%)", perf_12m, -50, 50, "%"), use_container_width=True)

with gauge_cols[3]:
    st.plotly_chart(creer_jauge_custom("Perf. 5A (%)", perf_5a, -50, 150, "%"), use_container_width=True)

with gauge_cols[4]:
    st.plotly_chart(creer_jauge_custom("Position (σ)", position_sigma, -3, 3, "σ"), use_container_width=True)

st.write("")

# -----------------------------------------------------------------------------
# 6. GRAPHIQUE LOGARITHMIQUE SUR 15 ANS
# -----------------------------------------------------------------------------
st.markdown("<div class='card-title' style='background:white; padding:15px; border-radius:12px 12px 0 0; margin-bottom:0;'>📉 Graphique de Régression Log-Normale & Projection 5 ans</div>", unsafe_allow_html=True)

fig_chart = go.Figure()

# Historique
fig_chart.add_trace(go.Scatter(
    x=df.index, y=df['Close'],
    mode='lines', name='Historique',
    line=dict(color='#147394', width=2)
))

# Régression
fig_chart.add_trace(go.Scatter(
    x=df.index, y=df['Theoretical'],
    mode='lines', name='Régression',
    line=dict(color='#009FB1', width=2)
))

# Prévision
fig_chart.add_trace(go.Scatter(
    x=futures_dates, y=np.exp(log_pred_futur),
    mode='lines', name='Prévision 5A',
    line=dict(color='#00915A', dash='dash', width=2)
))

# Bandes d'écart-type
df_futur_dates = list(df.index) + futures_dates
all_log_pred = list(df['Log_Pred'].values) + list(log_pred_futur)

fig_chart.add_trace(go.Scatter(x=df_futur_dates, y=np.exp(np.array(all_log_pred) + 2*std_dev), mode='lines', name='+2σ', line=dict(color='#61C3D7', dash='dot', width=1.2)))
fig_chart.add_trace(go.Scatter(x=df_futur_dates, y=np.exp(np.array(all_log_pred) + std_dev), mode='lines', name='+1σ', line=dict(color='#8ACABB', dash='dashdot', width=1.2)))
fig_chart.add_trace(go.Scatter(x=df_futur_dates, y=np.exp(np.array(all_log_pred) - std_dev), mode='lines', name='-1σ', line=dict(color='#8ACABB', dash='dashdot', width=1.2)))
fig_chart.add_trace(go.Scatter(x=df_futur_dates, y=np.exp(np.array(all_log_pred) - 2*std_dev), mode='lines', name='-2σ', line=dict(color='#61C3D7', dash='dot', width=1.2)))

# Ligne horizontale pour le prix cible des analystes
#if target_price is not None and target_price > 0:
#    fig_chart.add_hline(y=target_price, line_dash="dash", line_color="#00915A", line_width=1.5, name='Prix Cible Analystes', annotation_text="Prix Cible Analystes", annotation_position="bottom right", annotation_font=dict(color="#00915A", size=10))

if target_price is not None and target_price > 0:
    fig_chart.add_hline(
        y=target_price,
        line_dash="dash",
        line_color="#3F5056",
        line_width=1.5,
        )
else:
    pass

fig_chart.update_layout(
    yaxis_type="log",
    yaxis=dict(title=f"Prix ({symbole_devise})", side="left", gridcolor="#eef4f4"),
    xaxis=dict(title="Année", gridcolor="#eef4f4"),
    hovermode="x unified",
    height=500,
    paper_bgcolor='#ffffff',
    plot_bgcolor='#ffffff',
    margin=dict(l=20, r=20, t=20, b=20),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)

st.plotly_chart(fig_chart, use_container_width=True)

# -----------------------------------------------------------------------------
# 6b. SECTION RSI (LIMITÉE À 1 AN / 365 JOURS)
# -----------------------------------------------------------------------------
st.write("")

# 1. Calcul de la valeur actuelle du RSI
rsi_actuel = df['RSI'].iloc[-1]

# 2. Détermination de la zone RSI et création du badge (Résout le NameError)
if rsi_actuel >= 70:
    rsi_status = "<span class='badge badge-red'>🚨 Surachat (>70)</span>"
elif rsi_actuel <= 30:
    rsi_status = "<span class='badge badge-green'>🟢 Survente (<30)</span>"
else:
    rsi_status = "<span class='badge badge-green'>🟡 Neutre</span>"

# 3. Filtrage du DataFrame uniquement sur la dernière année (365 jours)
date_limite_1an = df.index[-1] - pd.Timedelta(days=365)
df_rsi_1an = df[df.index >= date_limite_1an]

# 4. En-tête avec Score RSI à droite (grande police) + Badge
st.markdown(f"""
<div class='card-title' style='background:white; padding:15px 20px; border-radius:12px 12px 0 0; margin-bottom:0; display:flex; justify-content:space-between; align-items:center;'>
    <div>📉 Indicateur RSI (14 jours - Glissant 1 an)</div>
    <div style='display:flex; align-items:center; gap:12px;'>
        <span style='font-size:1.4rem; font-weight:800; color:#147394; font-family:"Consolas", monospace;'>
            {fmt(rsi_actuel, 1)}
        </span>
        {rsi_status}
    </div>
</div>
""", unsafe_allow_html=True)

# 5. Création et affichage du graphique Plotly
fig_rsi = go.Figure()

fig_rsi.add_trace(go.Scatter(
    x=df_rsi_1an.index,
    y=df_rsi_1an['RSI'],
    mode='lines',
    name='RSI (14)',
    line=dict(color='#147394', width=1.8)
))

# Zone neutre (30-70)
fig_rsi.add_hrect(
    y0=30, y1=70,
    fillcolor="#EEF4F4", opacity=0.5,
    layer="below", line_width=0
)

# Lignes de surachat (70) et survente (30)
fig_rsi.add_hline(y=70, line_dash="dash", line_color="#E11D48", line_width=1.2, annotation_text="Surachat (70)", annotation_position="top left", annotation_font=dict(color="#E11D48", size=10))
fig_rsi.add_hline(y=30, line_dash="dash", line_color="#00915A", line_width=1.2, annotation_text="Survente (30)", annotation_position="bottom left", annotation_font=dict(color="#00915A", size=10))

fig_rsi.update_layout(
    yaxis=dict(range=[0, 100], gridcolor="#EEF4F4", tickvals=[0, 30, 50, 70, 100]),
    xaxis=dict(gridcolor="#EEF4F4"),
    height=220,
    paper_bgcolor='#ffffff',
    plot_bgcolor='#ffffff',
    margin=dict(l=20, r=20, t=10, b=20),
    showlegend=False
)

st.plotly_chart(fig_rsi, use_container_width=True, config={'displayModeBar': False})

# -----------------------------------------------------------------------------
# 7. HEATMAPS CÔTE À CÔTE : PERFORMANCE (%) & VOLATILITÉ (%)
# -----------------------------------------------------------------------------
st.write("")
st.markdown("<div class='card-title' style='background:white; padding:15px; border-radius:12px 12px 0 0; margin-bottom:0;'>🔥 Analyses Mensuelles Historiques (10 ans)</div>", unsafe_allow_html=True)

col_hm1, col_hm2 = st.columns(2)

# =============================================================================
# 8. HEATMAP PERFORMANCE MENSUELLE (%) - Style Zonebourse
# =============================================================================
df_perf_m = df[['Close']].resample('ME').last()
df_perf_m['Perf_%'] = df_perf_m['Close'].pct_change() * 100
df_perf_m['Year'] = df_perf_m.index.year
df_perf_m['Month'] = df_perf_m.index.month

perf_matrix = df_perf_m.pivot(index='Year', columns='Month', values='Perf_%')
mois_nom = ['Janv', 'Fév', 'Mars', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sept', 'Oct', 'Nov', 'Déc']
perf_matrix.columns = [mois_nom[m-1] for m in perf_matrix.columns if m <= 12]

# Palette Performance : Rouge (Baisse) -> Blanc (Neutre) -> Vert (#00915A)
colorscale_perf = [

    [0.0, "#E11D48"],   # Rouge
    [0.5, "#E8EDEF"],   # Blanc neutre
    [0.8, "#147394"],
    [1.0, "#3F5056"]
]

fig_perf = go.Figure(data=go.Heatmap(
    z=perf_matrix.values,
    x=perf_matrix.columns,
    y=perf_matrix.index,
    colorscale=colorscale_perf,
    zmid=0,
    colorbar=dict(title="%"),
    text=np.round(perf_matrix.values, 1),
    texttemplate="%{text}%",
    hoverongaps=False
))

fig_perf.update_layout(
    title=dict(text="📊 Performance Mensuelle (%)", font=dict(size=14, color="#147394")),
    xaxis_title="Mois",
    yaxis_title="Année",
    height=420,
    yaxis=dict(autorange="reversed"),
    paper_bgcolor='#ffffff',
    plot_bgcolor='#ffffff',
    margin=dict(l=10, r=10, t=40, b=20)
)

with col_hm1:
    st.plotly_chart(fig_perf, use_container_width=True)

# =============================================================================
# 9. HEATMAP VOLATILITÉ MENSUELLE (%) - Annualisée
# =============================================================================
df_vol = df[['Close']].copy()
df_vol['Returns'] = np.log(df_vol['Close'] / df_vol['Close'].shift(1))
df_vol['Year'] = df_vol.index.year
df_vol['Month'] = df_vol.index.month

vol_matrix = df_vol.groupby(['Year', 'Month'])['Returns'].std() * np.sqrt(252) * 100
vol_matrix = vol_matrix.unstack(level='Month')
vol_matrix.columns = [mois_nom[m-1] for m in vol_matrix.columns if m <= 12]

# Palette Volatilité (de votre palette) : #ADD9D3 -> #61C3D7 -> #147394 -> #3F5056
colorscale_vol = [
    [0.0, "#F4F8F8"],
    [0.2, "#ADD9D3"],
    [0.5, "#61C3D7"],
    [0.8, "#147394"],
    [1.0, "#3F5056"]
]

fig_vol = go.Figure(data=go.Heatmap(
    z=vol_matrix.values,
    x=vol_matrix.columns,
    y=vol_matrix.index,
    colorscale=colorscale_vol,
    colorbar=dict(title="%"),
    text=np.round(vol_matrix.values, 1),
    texttemplate="%{text}%",
    hoverongaps=False
))

fig_vol.update_layout(
    title=dict(text="⚡ Volatilité Mensuelle Annualisée (%)", font=dict(size=14, color="#147394")),
    xaxis_title="Mois",
    yaxis_title="Année",
    height=420,
    yaxis=dict(autorange="reversed"),
    paper_bgcolor='#ffffff',
    plot_bgcolor='#ffffff',
    margin=dict(l=10, r=10, t=40, b=20)
)

with col_hm2:
    st.plotly_chart(fig_vol, use_container_width=True)


# -----------------------------------------------------------------------------
# 7. HISTORIQUE COURS : DRAWDOWN (10A) & COURS (10A)
# -----------------------------------------------------------------------------
st.write("")
st.markdown("<div class='card-title' style='background:white; padding:15px; border-radius:12px 12px 0 0; margin-bottom:0;'>📅 Historique de cours (10 ans)</div>", unsafe_allow_html=True)

col_h1, col_h2 = st.columns(2)

# =============================================================================
# 10. DRAWDOWN 10 ANS
# =============================================================================
# 1. Calcul du drawdown et de sa médiane (arrondie à l'entier)
df['Peak'] = df['Close'].cummax()
df['Drawdown'] = ((df['Close'] - df['Peak']) / df['Peak']) * 100
current_drawdown = df['Drawdown'].iloc[-1]

# 2. Création du graphique Plotly
fig = go.Figure()

# Ajout de la zone remplie
fig.add_trace(go.Scatter(
    x=df.index,
    y=df['Drawdown'],
    mode='lines',
    name='Drawdown',
    fill='tozeroy',
    line=dict(color='rgba(240, 89, 89, 1)', width=2),
    fillcolor='rgba(240, 89, 89, 0.1)',
    hoverinfo='skip'
))

# 3. Configuration de l'affichage

# Style de l'axe Y (étiquettes à droite, format pourcentage)
fig.update_yaxes(
    tickformat=".0f",
    ticksuffix="%",
    ticks="outside",
    showline=True,
    linecolor='lightgray',
    side="right",
    range=[df['Drawdown'].min() * 1.1, 0]
)

# Style de l'axe X (graduations annuelles)
fig.update_xaxes(
    showline=True,
    linecolor='lightgray',
    ticks="inside",
    dtick= "M12",
    ticklabelmode="period", 
    tickformat="%Y",
    tickangle=0
)

# Ajout de l'annotation pour le drawdown actuel
annotation_text = f"{current_drawdown:.1f}%"
fig.add_annotation(
    x=df.index[-1], # Position x = dernière date
    y=current_drawdown, # Position y = valeur actuelle
    text=annotation_text,
    showarrow=False,
    xref="x",
    yref="y",
    font=dict(color="white", size=12),
    bgcolor="rgba(240, 89, 89, 0.95)", # Fond rouge
    bordercolor="white",
    borderwidth=1,
    borderpad=3,
    ax=30, # Décalage x de l'étiquette
    ay=0 # Décalage y de l'étiquette
)
# 2. Ajout de la ligne horizontale rouge du Max Drawdown
fig.add_trace(go.Scatter(
    x=[df.index[0], df.index[-1]],
    y=[df['Drawdown'].min(), df['Drawdown'].min()],
    mode='lines',
    line=dict(color='rgba(240, 89, 89, 0.9)', width=1, dash='solid'),
    showlegend=False,
    hoverinfo='skip'
    ))

# Ajout du bandeau d'information 
fig.add_annotation(
    x=df.index[0],
    y=df['Drawdown'].min() * 0.95, # Position basse
    text=f"<b>{ticker_input}  {df['Drawdown'].min():.1f}%</b>",
    showarrow=False,
    xref="x",
    yref="y",
    xanchor="left",
    font=dict(color="black", size=14)
)

# Mise en page générale
fig.update_layout(
    title=f"<b>Drawdown {info.get('shortName', ticker_input)}</b>",
    xaxis_title="", # Pas de titre sur l'axe X (les années suffisent)
    yaxis_title="", # pas de titre sur l'axe Y
    template="plotly_white", # Fond clair
    margin=dict(l=10, r=80, t=40, b=50), # Marges ajustées, surtout à droite
    height=400,
    hovermode="x unified", # Ligne verticale au survol montrant la valeur à cette date
    showlegend=False # Pas de légende
)

# --- 4. Affichage dans Streamlit ---
with col_h1:
    st.plotly_chart(fig, use_container_width=True)

# =============================================================================
# 11. COURS 10 ANS
# =============================================================================
# 1. Récupération des données de prix et calcul des repères
current_price = df['Close'].iloc[-1]
min_price = df['Close'].min()
max_price = df['Close'].max()

# 2. Création du graphique Plotly
fig = go.Figure()

# Ajout de la zone remplie pour le cours de l'action
fig.add_trace(go.Scatter(
    x=df.index,
    y=df['Close'],
    mode='lines',
    name='Cours',
    fill='tozeroy',
    line=dict(color='#009FB1', width=2),  # Bleu professionnel
    fillcolor='rgba(41, 98, 255, 0.1)',              # Remplissage bleu très transparent
    hoverinfo='skip'
))

# 3. Configuration de l'affichage

# Style de l'axe Y (étiquettes à droite, format monétaire/nombre)
fig.update_yaxes(
    tickformat=",.2f",         # Format décimal pour un prix (ex: 150.50)
    ticks="outside",
    showline=True,
    linecolor='lightgray',
    side="right",              # Étiquettes à droite
    range=[min_price * 0.95, max_price * 1.05]  # Plage ajustée avec un peu de marge haut/bas
)

# Style de l'axe X (graduations annuelles)
fig.update_xaxes(
    showline=True,
    linecolor='lightgray',
    ticks="inside",
    dtick="M12",
    ticklabelmode="period", 
    tickformat="%Y",
    tickangle=0
)

# Ajout de l'annotation pour le prix actuel (bulle en haut/milieu à droite)
annotation_text = f"{current_price:.2f}"
fig.add_annotation(
    x=df.index[-1],          # Position x = dernière date
    y=current_price,         # Position y = prix actuel
    text=annotation_text,
    showarrow=False,
    xref="x",
    yref="y",
    font=dict(color="white", size=12),
    bgcolor="rgba(41, 98, 255, 0.95)", # Fond bleu assorti
    bordercolor="white",
    borderwidth=1,
    borderpad=3,
    ax=30, 
    ay=0 
)

# Mise en page générale
fig.update_layout(
    title=f"<b>Cours {info.get('shortName', ticker_input)}</b>",
    xaxis_title="", 
    yaxis_title="", 
    template="plotly_white", 
    margin=dict(l=10, r=80, t=40, b=50), 
    height=400,
    hovermode="x unified", 
    showlegend=False 
)

# --- 4. Affichage dans Streamlit ---
with col_h2:
    st.plotly_chart(fig, use_container_width=True)

