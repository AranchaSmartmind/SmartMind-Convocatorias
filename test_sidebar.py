import streamlit as st

st.set_page_config(
    page_title="Test Sidebar",
    layout="wide",
    initial_sidebar_state="expanded"
)

with st.sidebar:
    st.write("¿Me ves?")
    st.button("Botón de prueba")

st.write("Contenido principal")