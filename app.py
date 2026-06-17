import streamlit as st

st.write("Verificando configuração...")

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    if api_key:
        st.success("✅ Chave GEMINI_API_KEY configurada com sucesso!")
        st.write(f"Primeiros caracteres: {api_key[:10]}...")
    else:
        st.error("❌ Chave está vazia")
except KeyError:
    st.error("❌ Chave GEMINI_API_KEY não encontrada nos Secrets")
