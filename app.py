import streamlit as st
import google.generativeai as genai

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    
    model = genai.GenerativeModel("gemini-3.5-flash")
    response = model.generate_content("Olá! Você está funcionando?")
    
    st.success("✅ API do Google Generative AI está funcionando!")
    st.write(response.text)
except Exception as e:
    st.error(f"❌ Erro: {str(e)}")
