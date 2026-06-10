import streamlit as st

from lt_code import render_app


st.set_page_config(
    page_title="WTA Loco Tequila",
    layout="wide",
)

try:
    render_app()
except Exception as exc:
    st.warning(f"No se pudo generar el dashboard: {exc}")
