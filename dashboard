from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import streamlit as st
import yfinance as yf

# Configuration de la page Streamlit (mode sombre automatique ou forcé)
st.set_page_config(
    page_title="Mon Portefeuille Boursier", page_icon="📈", layout="centered"
)

# Style CSS personnalisé pour imiter l'affichage sombre de l'image
st.markdown(
    """
    <style>
    .stApp {
        background-color: #1c1c1e;
        color: #ffffff;
    }
    .stock-card {
        background-color: #2c2c2e;
        padding: 15px;
        border-radius: 12px;
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Mon Portefeuille")
st.markdown(
    "<p style='color: #8e8e93;'>Suivi des cours sur les 6 derniers mois</p>",
    unsafe_allow_html=True,
)

# Votre liste de valeurs
portfolio = [
    ("ALV.DE", "Allianz"),
    ("AMAT", "Applied Materials"),
    ("AI.PA", "Air Liquide"),
    ("ACA.PA", "Crédit Agricole"),
    ("LLY", "Eli Lilly"),
    ("INVE-B.ST", "Investor AB"),
    ("GTT.PA", "Gaztransport & Technogaz"),
    ("HAUTO.OL", "Hoegh Autoliners"),
    ("ALMKT.PA", "Mauna Kea Technologies"),
    ("NOV.DE", "Novo"),
    ("SU.PA", "Schneider Electric"),
    ("SNDK", "Sandisk"),
    ("TTE.PA", "TotalEnergies"),
    ("V", "Visa"),
    ("DCAM.PA", "Msci World"),
    ("UST.PA", "Nasdaq 100"),
    ("GBSE.MI", "Gold"),
    ("BTC-EUR", "Bitcoin"),
    ("SOL-EUR", "Solana"),
]


@st.cache_data(ttl=600)  # Met en cache les données pendant 10 minutes
def load_data(ticker_symbol):
  end_date = datetime.now()
  start_date = end_date - timedelta(days=180)
  try:
    stock = yf.Ticker(ticker_symbol)
    df = stock.history(start=start_date, end=end_date)
    if df.empty:
      return None, 0, 0, []
    current_price = df["Close"].iloc[-1]
    prev_close = stock.info.get(
        "previousClose", df["Close"].iloc[-2] if len(df) > 1 else current_price
    )
    change = current_price - prev_close
    return df, current_price, change, df["Close"].tolist()
  except Exception:
    return None, 0, 0, []


# Affichage de chaque valeur sous forme de "carte"
for symbol, name in portfolio:
  df, price, change, prices = load_data(symbol)

  if df is None or len(prices) == 0:
    continue

  is_positive = change >= 0
  color = "#30d158" if is_positive else "#ff453a"
  arrow = "▲" if is_positive else "▼"
  change_str = f"+{change:.2f}" if is_positive else f"{change:.2f}"

  # Structure en colonnes (Symbole/Nom | Mini-graphe | Prix/Variation)
  with st.container():
    st.markdown('<div class="stock-card">', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([2, 2, 1.5])

    with col1:
      st.markdown(
          f"<h3 style='margin:0; color:{color}; font-size:18px;'>{arrow}"
          f" {symbol}</h3>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<p style='margin:0; color:#8e8e93; font-size:13px;'>{name}</p>",
          unsafe_allow_html=True,
      )

    with col2:
      # Génération du mini-graphique (sparkline)
      fig, ax = plt.subplots(figsize=(2, 0.6), dpi=100)
      fig.patch.set_facecolor("#2c2c2e")
      ax.set_facecolor("#2c2c2e")
      ax.plot(prices, color=color, linewidth=1.5)
      ax.axis("off")
      fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
      st.pyplot(fig, use_container_width=False)
      plt.close(fig)

    with col3:
      st.markdown(
          f"<p style='margin:0; text-align:right; font-weight:bold;"
          f" font-size:16px;'>{price:.2f}</p>",
          unsafe_allow_html=True,
      )
      st.markdown(
          f"<p style='margin:0; text-align:right; color:{color};"
          f" font-size:13px; font-weight:bold;'>{change_str}</p>",
          unsafe_allow_html=True,
      )

    st.markdown("</div>", unsafe_allow_html=True)
