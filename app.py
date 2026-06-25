"""
app.py — Dashboard Analítica de CX · Xtay
Streamlit app: lê dados do Google Sheets e renderiza o dashboard.
Deploy: Streamlit Community Cloud
"""
import streamlit as st
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
import json, os

from processar_dados import processar

st.set_page_config(
    page_title="Dashboard CX · Xtay",
    page_icon="📨",
    layout="wide",
)

# Remove padding padrão do Streamlit para o dashboard ocupar a tela toda
st.markdown("""
<style>
    .block-container { padding: 0 !important; }
    header { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
</style>
""", unsafe_allow_html=True)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

@st.cache_data(ttl=300, show_spinner="Carregando dados do Google Sheets…+")
def load_data():
    creds_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    gc = gspread.authorize(creds)

    sheet_id = st.secrets["google_sheets"]["sheet_id"]
    sh = gc.open_by_key(sheet_id)

    # Aba de atendimentos — aceita múltiplos nomes (ordem de prioridade)
    ws_droz = None
    for name in ["Atendimentos", "Input Droz (Abr-Mai)", "Resultado da consulta"]:
        try:
            ws_droz = sh.worksheet(name)
            break
        except Exception:
            pass
    if ws_droz is None:
        ws_droz = sh.get_worksheet(0)

    rows_droz = ws_droz.get_all_values()[1:]  # pula header

    # Aba de OCC (opcional)
    rows_occ = None
    try:
        ws_occ = sh.worksheet("Reservas e OCC")
        rows_occ = ws_occ.get_all_values()[1:]
    except Exception:
        pass

    return rows_droz, rows_occ

# Carrega template HTML
_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(_dir, "template.html"), "r", encoding="utf-8") as f:
    TEMPLATE = f.read()

try:
    rows_droz, rows_occ = load_data()
    payload = processar(rows_droz, rows_occ)
    html = TEMPLATE.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    components.html(html, height=3600, scrolling=True)
except Exception as e:
    st.error(f"Erro ao carregar dados: {e}")
    st.info("Verifique se os secrets estão configurados corretamente no Streamlit Cloud.")
