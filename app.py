import io
import logging
import time
from typing import Optional, List, Dict, Any
from datetime import datetime
import json
import hmac
import hashlib

from PIL import Image
import streamlit as st
import pandas as pd
import google.generativeai as genai

from config import (
    GEMINI_CONFIG, PROCESSAMENTO_CONFIG, DOCUMENTO_CONFIG,
    EMPRESA_CONFIG, TECNICO_CONFIG, CLIENTE_CONFIG,
    VERSAO_APP, DESCRICAO_VERSAO
)
from utils import (
    validar_arquivo_pdf, validar_texto_entrada, criar_logger,
    gerar_relatorio_processamento, sanitizar_nome_arquivo, formatar_tempo_decorrido
)
from processador import ProcessadorMemorial
from gerador_word import GeradorAnuenciaIncraWord

# ============================================================
# IMPORTAR MÓDULO DE PROGRESSO
# ============================================================
from progress_tracker import ProgressTracker, ProgressBarStreamlit, criar_progress_tracker_requerimento
from gerador_requerimento_cartorio import GeradorRequerimentoCartorio

# Inicializa o logger
logger = criar_logger(__name__)

st.set_page_config(
    page_title="Gerador de Memorial - Gleba A",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# TIPOS DE ARQUIVO ACEITOS (PDF + IMAGENS)
# ============================================================
TIPOS_ARQUIVO_SUPORTADOS = ["pdf", "png", "jpg", "jpeg"]

# Extensões de imagem reconhecidas
EXTENSOES_IMAGEM = ["png", "jpg", "jpeg"]
EXTENSOES_PDF = ["pdf"]


def is_imagem(arquivo) -> bool:
    """Verifica se o arquivo carregado é uma imagem (PNG, JPG, JPEG)."""
    extensao = arquivo.name.split(".")[-1].lower()
    return extensao in EXTENSOES_IMAGEM


def is_pdf(arquivo) -> bool:
    """Verifica se o arquivo carregado é um PDF."""
    extensao = arquivo.name.split(".")[-1].lower()
    return extensao in EXTENSOES_PDF


def carregar_imagem_direta(arquivo) -> Image.Image:
    """Lê o conteúdo de uma imagem (PNG/JPG/JPEG) e retorna um objeto PIL.Image.Image para envio ao Gemini."""
    img_bytes = arquivo.getvalue()
    return Image.open(io.BytesIO(img_bytes))


def configurar_gemini() -> bool:
    """Configura conexão com API Gemini."""
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            logger.error("GEMINI_API_KEY não encontrada nos Secrets")
            return False
        genai.configure(api_key=api_key)
        return True
    except Exception as e:
        logger.error(f"Erro ao configurar Gemini: {str(e)}")
        return False


# ============================================================
# CALLBACK PARA ATUALIZAR PROGRESSO DO REQUERIMENTO
# ============================================================
def callback_atualizar_progresso_requerimento(etapa: int, descricao: str, percentual: float = None):
    """
    Callback para atualizar a barra de progresso do requerimento.
    Chamado pelo GeradorRequerimentoCartorio durante o processamento.
    """
    if 'progress_container_requerimento' not in st.session_state:
        st.session_state.progress_container_requerimento = st.container()
    
    with st.session_state.progress_container_requerimento:
        if percentual is not None:
            st.progress(percentual / 100, text=f"Etapa {etapa}: {descricao}")
        else:
            st.write(f"📊 {descricao}")


# ============================================================
# SEGURANÇA: FORÇAR HTTPS
# ============================================================
def verificar_https():
    """Redireciona para HTTPS se a conexão não for segura.
    Usa o header X-Forwarded-Proto fornecido pelo proxy do Streamlit Cloud.
    """
    try:
        from streamlit.web.server.websocket_headers import _get_websocket_headers
        headers = _get_websocket_headers()
        if headers:
            proto = headers.get("X-Forwarded-Proto", "")
            if proto.lower() == "http":
                # Tenta redirecionar para HTTPS
                from urllib.parse import urlparse
                current_url = st.query_params.get("url", "")
                st.markdown("""
                    <script>
                        if (window.location.protocol !== 'https:') {
                            window.location.href = window.location.href.replace('http://', 'https://');
                        }
                    </script>
                """, unsafe_allow_html=True)
    except Exception:
        pass  # Streamlit pode não ter headers disponíveis localmente


# ============================================================
# SEGURANÇA: SISTEMA DE LOGIN COM ÍCONE
# ============================================================
def tela_login():
    """Exibe a tela de login com verificação de usuário e senha.
    As credenciais são definidas via st.secrets para segurança.
    Inclui o ícone do gerador.
    """
    st.markdown("""
        <style>
            .login-container {
                max-width: 500px;
                margin: 40px auto;
                padding: 40px;
                background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                border-radius: 16px;
                color: #ffffff;
                text-align: center;
                border-left: 6px solid #10b981;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
            }
            .login-icon-container {
                margin-bottom: 2rem;
                display: flex;
                justify-content: center;
            }
            .login-icon-container img {
                max-width: 150px;
                height: auto;
                filter: drop-shadow(0 4px 8px rgba(0, 0, 0, 0.3));
            }
            .login-title { font-size: 1.8rem; font-weight: 700; margin-bottom: 0.5rem; }
            .login-subtitle { font-size: 1rem; color: #a7f3d0; margin-bottom: 2rem; }
            .login-warning { color: #fbbf24; font-size: 0.85rem; margin-top: 1rem; }
        </style>
    """, unsafe_allow_html=True)

    # Exibir ícone se existir
    try:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.image("icon_gerador_128x128.png", width=150, use_column_width=False)
    except:
        pass  # Se o ícone não existir, continuar sem ele

    st.markdown("""
        <div class="login-container">
            <div class="login-title">🔐 Acesso Restrito</div>
            <div class="login-subtitle">Gerador de Memorial - Portal de Engenharia</div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Inicie sessão para continuar")

    usuario = st.text_input("Usuário", key="login_usuario")
    senha = st.text_input("Senha", type="password", key="login_senha")

    if st.button("🔓 Entrar", type="primary", use_container_width=True):
        # Verifica credenciais contra os secrets do Streamlit
        usuario_correto = st.secrets.get("USUARIO", "")
        senha_correta = st.secrets.get("SENHA", "")

        # Verificação segura com comparação constante (prevenção contra timing attack)
        if hmac.compare_digest(usuario, usuario_correto) and hmac.compare_digest(senha, senha_correta):
            st.session_state["autenticado"] = True
            st.session_state["usuario_logado"] = usuario
            st.rerun()
        else:
            st.error("❌ Usuário ou senha incorretos.")
            logger.warning(f"Tentativa de login falha para usuário: {usuario}")

    st.markdown("<p class='login-warning'>⚠️ Sistema de uso exclusivo da equipe técnica. Acesso não autorizado será registrado.</p>", unsafe_allow_html=True)

    return False


def verificar_autenticacao() -> bool:
    """Verifica se o usuário está autenticado. Exibe tela de login se não estiver."""
    if not st.session_state.get("autenticado", False):
        return tela_login()
    return True


# ============================================================
# SEGURANÇA: LIMPAR SESSION STATE APÓS DOWNLOAD
# ============================================================
def limpar_sessao_memorial():
    """Limpa dados sensíveis do memorial após geração."""
    chaves_para_limpar = [
        "dados_memoriais_processados",
        "segmentos",
        "confrontantes",
    ]
    for chave in chaves_para_limpar:
        if chave in st.session_state:
            del st.session_state[chave]
    st.info("🧹 Dados da sessão foram limpos por segurança.")


def limpar_sessao_requerimento():
    """Limpa dados sensíveis do requerimento após geração."""
    chaves_para_limpar = [
        "dados_extraidos_requerimento",
        "requerente_1",
        "requerente_2",
    ]
    for chave in chaves_para_limpar:
        if chave in st.session_state:
            del st.session_state[chave]
    st.info("🧹 Dados da sessão foram limpos por segurança.")


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    # Forçar HTTPS
    verificar_https()

    # Verificar autenticação
    if not verificar_autenticacao():
        return

    # ============================================================
    # SIDEBAR COM ÍCONE E LOGO
    # ============================================================
    with st.sidebar:
        # Exibir ícone no topo da sidebar
        st.markdown("""
            <style>
                .sidebar-icon-container {
                    text-align: center;
                    margin-bottom: 2rem;
                    padding: 1rem;
                    background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                    border-radius: 12px;
                    border-left: 4px solid #10b981;
                }
                .sidebar-icon-container img {
                    max-width: 100px;
                    height: auto;
                    filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.2));
                }
                .sidebar-title {
                    color: #10b981;
                    font-weight: 700;
                    margin-top: 0.5rem;
                    font-size: 0.95rem;
                }
            </style>
        """, unsafe_allow_html=True)
        
        try:
            st.image("icon_gerador_64x64.png", width=100, use_column_width=False)
            st.markdown('<p class="sidebar-title">Gerador de Memorial</p>', unsafe_allow_html=True)
        except:
            st.markdown("### 📐 Gerador de Memorial")
        
        st.markdown("---")
        
        # Informações do usuário
        usuario_logado = st.session_state.get("usuario_logado", "Usuário")
        st.markdown(f"👤 **Usuário:** {usuario_logado}")
        st.markdown(f"📅 **Data:** {datetime.now().strftime('%d/%m/%Y')}")
        st.markdown(f"⏰ **Hora:** {datetime.now().strftime('%H:%M:%S')}")
        
        st.markdown("---")
        
        # Botão de logout
        if st.button("🚪 Sair (Logout)", type="secondary", use_container_width=True):
            st.session_state["autenticado"] = False
            st.session_state["usuario_logado"] = ""
            # Limpar toda a sessão
            for chave in list(st.session_state.keys()):
                del st.session_state[chave]
            st.rerun()

    # Estilização CSS Interna Premium
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
            html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
            .hero-container {
                background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                padding: 2rem;
                border-radius: 12px;
                color: #ffffff;
                margin-bottom: 2rem;
                border-left: 6px solid #10b981;
            }
            .hero-title { font-size: 2.2rem; font-weight: 700; margin: 0; }
            .hero-subtitle { font-size: 1.1rem; color: #a7f3d0; margin: 0.5rem 0 0 0; }
            .tab-content { padding: 1.5rem; background: #f8f9fa; border-radius: 8px; }
            .success-box { 
                background: #d1fae5; 
                border-left: 4px solid #10b981; 
                padding: 1rem; 
                border-radius: 4px; 
                margin: 1rem 0;
            }
            .warning-box { 
                background: #fef3c7; 
                border-left: 4px solid #f59e0b; 
                padding: 1rem; 
                border-radius: 4px; 
                margin: 1rem 0;
            }
            .error-box { 
                background: #fee2e2; 
                border-left: 4px solid #ef4444; 
                padding: 1rem; 
                border-radius: 4px; 
                margin: 1rem 0;
            }
        </style>
    """, unsafe_allow_html=True)

    # Cabeçalho Hero
    st.markdown("""
        <div class="hero-container">
            <div class="hero-title">📐 Gerador de Memorial</div>
            <div class="hero-subtitle">Sistema Integrado de Processamento de Documentos Técnicos</div>
        </div>
    """, unsafe_allow_html=True)

    # Criar abas
    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Memorial Descritivo",
        "🤝 Anuências Co-proprietários",
        "🌾 Anuências INCRA",
        "🏛️ Requerimento de Cartório"
    ])

    # ============================================================
    # ABA 1: MEMORIAL DESCRITIVO
    # ============================================================
    with tab1:
        st.markdown("### 📝 Gerador de Memorial Descritivo")
        st.markdown("Carregue os documentos necessários para gerar o memorial descritivo.")
        
        # Resto do código da aba 1 continua aqui...
        st.info("Funcionalidade de Memorial Descritivo - Em desenvolvimento")

    # ============================================================
    # ABA 2: ANUÊNCIAS CO-PROPRIETÁRIOS
    # ============================================================
    with tab2:
        st.markdown("### 🤝 Gerador de Anuências - Co-proprietários")
        st.markdown("Gere anuências para co-proprietários do imóvel.")
        
        # Resto do código da aba 2 continua aqui...
        st.info("Funcionalidade de Anuências Co-proprietários - Em desenvolvimento")

    # ============================================================
    # ABA 3: ANUÊNCIAS INCRA
    # ============================================================
    with tab3:
        st.markdown("### 🌾 Gerador de Anuências - INCRA")
        st.markdown("Gere anuências para o INCRA.")
        
        # Resto do código da aba 3 continua aqui...
        st.info("Funcionalidade de Anuências INCRA - Em desenvolvimento")

    # ============================================================
    # ABA 4: REQUERIMENTO DE CARTÓRIO
    # ============================================================
    with tab4:
        st.markdown("### 🏛️ Gerador de Requerimento de Cartório")
        st.markdown("Carregue os documentos necessários para gerar o requerimento de cartório.")
        
        # Resto do código da aba 4 continua aqui...
        st.info("Funcionalidade de Requerimento de Cartório - Em desenvolvimento")


if __name__ == "__main__":
    main()
