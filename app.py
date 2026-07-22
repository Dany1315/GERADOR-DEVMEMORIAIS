# GERADOR DE MEMORIAL DESCRITIVO - Versão Final com INCRA e Cartório
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
from gerador_memorial_word import GeradorMemorialWord
from gerador_anuencias import GeradorAnuenciaWord
from gerador_anuencia_incra import GeradorAnuenciaIncraWord as GeradorIncraCompleto
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
# SEGURANÇA: FORÇAR HTTPS
# ============================================================
def verificar_https():
    """Redireciona para HTTPS se a conexão não for segura."""
    try:
        from streamlit.web.server.websocket_headers import _get_websocket_headers
        headers = _get_websocket_headers()
        if headers:
            proto = headers.get("X-Forwarded-Proto", "")
            if proto.lower() == "http":
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
        pass


# ============================================================
# SEGURANÇA: SISTEMA DE LOGIN
# ============================================================
def tela_login():
    """Exibe a tela de login com verificação de usuário e senha."""
    st.markdown("""
        <style>
            .login-container {
                max-width: 400px;
                margin: 80px auto;
                padding: 40px;
                background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                border-radius: 16px;
                color: #ffffff;
                text-align: center;
                border-left: 6px solid #10b981;
            }
            .login-title { font-size: 1.8rem; font-weight: 700; margin-bottom: 0.5rem; }
            .login-subtitle { font-size: 1rem; color: #a7f3d0; margin-bottom: 2rem; }
            .login-warning { color: #fbbf24; font-size: 0.85rem; margin-top: 1rem; }
        </style>
    """, unsafe_allow_html=True)

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
        usuario_correto = st.secrets.get("USUARIO", "")
        senha_correta = st.secrets.get("SENHA", "")

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
    """Verifica se o usuário está autenticado."""
    if not st.session_state.get("autenticado", False):
        return tela_login()
    return True


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    verificar_https()

    if not verificar_autenticacao():
        return

    with st.sidebar:
        if st.button("🚪 Sair (Logout)", type="secondary", use_container_width=True):
            st.session_state["autenticado"] = False
            st.session_state["usuario_logado"] = ""
            for chave in list(st.session_state.keys()):
                del st.session_state[chave]
            st.rerun()

    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;600;700&display=swap');
            html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
            .hero-container {
                background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                padding: 2.5rem; border-radius: 16px; color: #ffffff;
                margin-bottom: 2rem; border-left: 6px solid #10b981;
            }
            .hero-title { font-size: 2.2rem; font-weight: 700; margin-bottom: 0.5rem; }
            .hero-subtitle { font-size: 1.1rem; color: #a7f3d0; font-weight: 300; }
            div.stButton > button:first-child {
                background: linear-gradient(90deg, #065f46 0%, #022c22 100%);
                color: white; border: none; padding: 0.75rem 1.5rem;
                font-weight: 600; border-radius: 8px; transition: all 0.3s ease;
            }
            div.stButton > button:first-child:hover {
                transform: translateY(-2px);
                background: linear-gradient(90deg, #047857 0%, #064e3b 100%);
            }
            section[data-testid="stSidebar"] {
                background: linear-gradient(180deg, #022c22 0%, #064e3b 100%) !important;
                border-right: 1px solid #047857;
            }
            section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] p, 
            section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] span {
                color: #f8fafc !important;
            }
            section[data-testid="stSidebar"] input {
                background-color: #042f2e !important; color: #ffffff !important;
                border: 1px solid #115e59 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div class="hero-container">
            <div class="hero-title">📐 Portal de Engenharia & Topografia</div>
            <div class="hero-subtitle">Gerador Inteligente de Memoriais Descritivos & Gestão de Anuências</div>
        </div>
    """, unsafe_allow_html=True)

    usuario_logado = st.session_state.get("usuario_logado", "Usuário")
    st.caption(f"🔒 Sessão ativa: **{usuario_logado}** | Conexão segura")

    # ============================================================
    # PAINEL DE CONTROLE (SIDEBAR)
    # ============================================================
    with st.sidebar:
        st.markdown("<h2 style='font-size: 1.5rem;'>⚙️ Painel de Controle</h2>", unsafe_allow_html=True)
        tab_side_empresa, tab_side_cliente = st.tabs(["🏢 Empresa / Técnico", "👥 Cliente & Imóvel"])

        # Inicializar dados do painel lateral em session_state
        if "painel_empresa_nome" not in st.session_state:
            st.session_state["painel_empresa_nome"] = EMPRESA_CONFIG.NOME
        if "painel_empresa_endereco" not in st.session_state:
            st.session_state["painel_empresa_endereco"] = EMPRESA_CONFIG.ENDERECO
        if "painel_empresa_telefone" not in st.session_state:
            st.session_state["painel_empresa_telefone"] = EMPRESA_CONFIG.TELEFONE
        if "painel_empresa_email" not in st.session_state:
            st.session_state["painel_empresa_email"] = EMPRESA_CONFIG.EMAIL
        if "painel_tecnico_nome" not in st.session_state:
            st.session_state["painel_tecnico_nome"] = TECNICO_CONFIG.NOME
        if "painel_tecnico_cfta" not in st.session_state:
            st.session_state["painel_tecnico_cfta"] = TECNICO_CONFIG.CFTA
        if "painel_cpf_tecnico" not in st.session_state:
            st.session_state["painel_cpf_tecnico"] = "111.985.197-11"
        if "painel_cliente_imovel" not in st.session_state:
            st.session_state["painel_cliente_imovel"] = CLIENTE_CONFIG.IMOVEL
        if "painel_cliente_proprietario" not in st.session_state:
            st.session_state["painel_cliente_proprietario"] = CLIENTE_CONFIG.PROPRIETARIO
        if "painel_cliente_local" not in st.session_state:
            st.session_state["painel_cliente_local"] = CLIENTE_CONFIG.LOCAL
        if "painel_cliente_area" not in st.session_state:
            st.session_state["painel_cliente_area"] = CLIENTE_CONFIG.AREA
        if "painel_cliente_perimetro" not in st.session_state:
            st.session_state["painel_cliente_perimetro"] = CLIENTE_CONFIG.PERIMETRO
        if "painel_cliente_comarca" not in st.session_state:
            st.session_state["painel_cliente_comarca"] = "N/A"
        if "painel_cliente_matricula" not in st.session_state:
            st.session_state["painel_cliente_matricula"] = "N/A"
        if "painel_tecnico_trt" not in st.session_state:
            st.session_state["painel_tecnico_trt"] = "N/A"

        with tab_side_empresa:
            st.markdown("### Configurações Institucionais")
            st.session_state["painel_empresa_nome"] = st.text_input("Nome da Empresa", value=st.session_state["painel_empresa_nome"], key="input_empresa_nome")
            st.session_state["painel_empresa_endereco"] = st.text_input("Endereço", value=st.session_state["painel_empresa_endereco"], key="input_empresa_endereco")
            st.session_state["painel_empresa_telefone"] = st.text_input("Telefone", value=st.session_state["painel_empresa_telefone"], key="input_empresa_telefone")
            st.session_state["painel_empresa_email"] = st.text_input("Email", value=st.session_state["painel_empresa_email"], key="input_empresa_email")

            st.markdown("---")
            st.session_state["painel_tecnico_nome"] = st.text_input("Nome do Técnico", value=st.session_state["painel_tecnico_nome"], key="input_tecnico_nome")
            st.session_state["painel_tecnico_cfta"] = st.text_input("CFTA", value=st.session_state["painel_tecnico_cfta"], key="input_tecnico_cfta")
            st.session_state["painel_cpf_tecnico"] = st.text_input("CPF do Responsável Técnico", value=st.session_state["painel_cpf_tecnico"], key="input_cpf_tecnico")
            st.session_state["painel_tecnico_trt"] = st.text_input("TRT", value=st.session_state["painel_tecnico_trt"], key="input_tecnico_trt")

        with tab_side_cliente:
            st.markdown("### Especificações do Projeto")
            st.session_state["painel_cliente_imovel"] = st.text_input("📍 Identificação do Imóvel", value=st.session_state["painel_cliente_imovel"], key="input_cliente_imovel")
            st.session_state["painel_cliente_proprietario"] = st.text_input("👤 Proprietário", value=st.session_state["painel_cliente_proprietario"], key="input_cliente_proprietario")
            st.session_state["painel_cliente_local"] = st.text_input("🗺️ Município / Localidade", value=st.session_state["painel_cliente_local"], key="input_cliente_local")
            
            col_area, col_per = st.columns(2)
            with col_area:
                st.session_state["painel_cliente_area"] = st.text_input("📐 Área (ha)", value=st.session_state["painel_cliente_area"], key="input_cliente_area")
            with col_per:
                st.session_state["painel_cliente_perimetro"] = st.text_input("🏃 Perímetro (m)", value=st.session_state["painel_cliente_perimetro"], key="input_cliente_perimetro")
            
            st.session_state["painel_cliente_comarca"] = st.text_input("⚖️ Comarca", value=st.session_state["painel_cliente_comarca"], key="input_cliente_comarca")
            st.session_state["painel_cliente_matricula"] = st.text_input("📋 Matrícula", value=st.session_state["painel_cliente_matricula"], key="input_cliente_matricula")

    # Abas Principais
    tab_memorial, tab_anuencias, tab_anuencias_incra, tab_requerimento = st.tabs([
        "📝 Memorial Descritivo", 
        "🤝 Anuências Co-proprietários", 
        "🌾 Anuências INCRA",
        "🏛️ Requerimento de Cartório"
    ])

    with tab_memorial:
        st.markdown("### 📥 Entrada de Dados do Memorial")
        tab_pdf, tab_manual = st.tabs([
            "📁 Processamento de Arquivos (PDF ou Imagem)",
            "📝 Colagem de Texto Manual"
        ])

        pdf_planta, pdf_roteiro = None, None
        texto_planta_manual, texto_roteiro_manual = "", ""

        with tab_pdf:
            st.markdown("**Envie os arquivos da planta e do roteiro perimetral:**")
            col1, col2 = st.columns(2)
            with col1:
                pdf_planta = st.file_uploader("📄 Planta (PDF/Imagem)", type=TIPOS_ARQUIVO_SUPORTADOS, key="upload_planta")
            with col2:
                pdf_roteiro = st.file_uploader("📄 Roteiro Perimetral (PDF/Imagem)", type=TIPOS_ARQUIVO_SUPORTADOS, key="upload_roteiro")

        with tab_manual:
            st.markdown("**Ou cole o texto manualmente:**")
            col1, col2 = st.columns(2)
            with col1:
                texto_planta_manual = st.text_area("Texto da Planta", height=150, key="texto_planta")
            with col2:
                texto_roteiro_manual = st.text_area("Texto do Roteiro", height=150, key="texto_roteiro")

        if st.button("🔍 Analisar Documentos", type="primary", use_container_width=True):
            try:
                if not configurar_gemini():
                    st.error("❌ Falha ao configurar Gemini API")
                    return

                processador = ProcessadorMemorial()

                # Processar planta
                if pdf_planta:
                    conteudo_planta = pdf_planta.getvalue()
                    st.session_state["arquivo_memorial_bytes"] = conteudo_planta
                    st.session_state["nome_arquivo_memorial"] = pdf_planta.name
                else:
                    conteudo_planta = texto_planta_manual.encode()

                # Processar roteiro
                if pdf_roteiro:
                    conteudo_roteiro = pdf_roteiro.getvalue()
                else:
                    conteudo_roteiro = texto_roteiro_manual.encode()

                # Processar com IA
                with st.spinner("⏳ Processando com Gemini..."):
                    segmentos = processador.processar(conteudo_planta, conteudo_roteiro)

                st.session_state["segmentos"] = segmentos
                st.session_state["dados_memoriais_processados"] = True

                # Preparar dados_finais
                dados_finais = {
                    "cliente": {
                        "proprietario": st.session_state["painel_cliente_proprietario"],
                        "local": st.session_state["painel_cliente_local"],
                        "imovel": st.session_state["painel_cliente_imovel"],
                        "area": st.session_state["painel_cliente_area"],
                        "perimetro": st.session_state["painel_cliente_perimetro"],
                        "comarca": st.session_state["painel_cliente_comarca"],
                        "matricula": st.session_state["painel_cliente_matricula"],
                    },
                    "empresa": {
                        "nome": st.session_state["painel_empresa_nome"],
                        "endereco": st.session_state["painel_empresa_endereco"],
                        "telefone": st.session_state["painel_empresa_telefone"],
                        "email": st.session_state["painel_empresa_email"],
                    },
                    "tecnico": {
                        "nome": st.session_state["painel_tecnico_nome"],
                        "cfta": st.session_state["painel_tecnico_cfta"],
                        "cpf": st.session_state["painel_cpf_tecnico"],
                        "trt": st.session_state["painel_tecnico_trt"],
                    },
                    "segmentos": segmentos,
                }

                st.session_state["dados_finais"] = dados_finais

                st.success("✅ Processamento concluído!")
                st.info(f"📊 {len(segmentos)} segmentos extraídos com sucesso")

            except Exception as e:
                st.error(f"❌ Erro ao processar: {str(e)}")
                logger.error(f"Erro no processamento: {str(e)}")

        # Gerar Memorial
        if st.session_state.get("dados_memoriais_processados"):
            st.markdown("---")
            st.markdown("### 📄 Gerar Memorial Descritivo")

            if st.button("📝 Gerar Memorial", type="primary", use_container_width=True):
                try:
                    gerador = GeradorMemorialWord(st.session_state.get("dados_finais", {}))
                    arquivo_docx = gerador.gerar_documento()

                    st.download_button(
                        label="📥 BAIXAR MEMORIAL DESCRITIVO (.DOCX)",
                        data=arquivo_docx.getvalue(),
                        file_name=f"MEMORIAL_{sanitizar_nome_arquivo(st.session_state['painel_cliente_proprietario'].upper())}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_download_memorial"
                    )
                    st.success("✅ Memorial gerado com sucesso!")

                except Exception as e:
                    st.error(f"❌ Erro ao gerar memorial: {str(e)}")
                    logger.error(f"Erro ao gerar memorial: {str(e)}")

    # ============================================================
    # ABA 2: ANUÊNCIAS CO-PROPRIETÁRIOS
    # ============================================================
    with tab_anuencias:
        st.markdown("### 🤝 Gerador de Anuências Co-proprietários")

        if st.session_state.get("dados_memoriais_processados"):
            segmentos = st.session_state.get("segmentos", [])

            if isinstance(segmentos, list) and len(segmentos) > 0:
                confrontantes_unicos = list(set([s.get("confrontante", "N/A") for s in segmentos if s.get("confrontante")]))

                if confrontantes_unicos:
                    st.markdown(f"**Confrontantes encontrados:** {len(confrontantes_unicos)}")

                    for idx, confrontante in enumerate(confrontantes_unicos, 1):
                        with st.expander(f"👤 {idx}. {confrontante}"):
                            col1, col2 = st.columns(2)
                            with col1:
                                nome_anuencia = st.text_input(f"Nome completo ({confrontante}):", value=confrontante, key=f"nome_anu_{idx}")
                            with col2:
                                cpf_anuencia = st.text_input(f"CPF ({confrontante}):", key=f"cpf_anu_{idx}")

                            endereco_anuencia = st.text_area(f"Endereço ({confrontante}):", key=f"end_anu_{idx}")

                            trt_anuencia = st.text_input(f"TRT ({confrontante}):", key=f"trt_anu_{idx}")

                            if st.button(f"📄 Gerar Anuência para {confrontante}", key=f"btn_anu_{idx}"):
                                try:
                                    segmentos_confrontante = [s for s in st.session_state.get("segmentos", []) if s.get("confrontante", "").upper() == confrontante.upper()]

                                    dados_anuencia = {
                                        "confrontante": nome_anuencia,
                                        "proprietario": st.session_state["painel_cliente_proprietario"],
                                        "local": st.session_state["painel_cliente_local"],
                                        "segmentos": segmentos_confrontante
                                    }

                                    gerador_anu = GeradorAnuenciaWord(st.session_state.get("dados_finais", {}).get("empresa", {}), st.session_state.get("dados_finais", {}).get("tecnico", {}))
                                    arquivo_anuencia = gerador_anu.gerar_documento(dados_anuencia)

                                    st.download_button(
                                        label=f"📥 BAIXAR ANUÊNCIA - {nome_anuencia}",
                                        data=arquivo_anuencia.getvalue(),
                                        file_name=f"ANUENCIA_{sanitizar_nome_arquivo(nome_anuencia.upper())}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"download_anu_{idx}"
                                    )
                                    st.success(f"✅ Anuência gerada para {nome_anuencia}")
                                except Exception as e:
                                    st.error(f"❌ Erro ao gerar anuência: {str(e)}")
                                    logger.error(f"Erro ao gerar anuência para {nome_anuencia}: {str(e)}")
                else:
                    st.info("ℹ️ Nenhum confrontante foi encontrado.")
            else:
                st.warning("⚠️ Nenhum segmento foi processado. Volte à aba 1 e processe um arquivo.")
        else:
            st.info("💡 Processe um memorial na aba 1 para gerar anuências.")

    # ============================================================
    # ABA 3: ANUÊNCIAS INCRA
    # ============================================================
    with tab_anuencias_incra:
        st.markdown("### 🌾 Gerador de Anuências INCRA")

        if st.session_state.get("dados_memoriais_processados"):
            st.info("📄 Processando memorial para gerar anuências INCRA...")

            try:
                dados_projeto = {
                    "proprietario": st.session_state.get("painel_cliente_proprietario", "N/A"),
                    "cpf_proprietario": st.session_state.get("painel_cpf_tecnico", "N/A"),
                    "imovel": st.session_state.get("painel_cliente_imovel", "N/A"),
                    "localidade": st.session_state.get("painel_cliente_local", "N/A")
                }

                gerador_incra = GeradorIncraCompleto(
                    st.session_state.get("dados_finais", {}).get("empresa", {}),
                    st.session_state.get("dados_finais", {}).get("tecnico", {})
                )

                if st.session_state.get("arquivo_memorial_bytes"):
                    documentos = gerador_incra.gerar_documentos_pelo_memorial(
                        st.session_state["arquivo_memorial_bytes"],
                        st.session_state.get("nome_arquivo_memorial", "memorial.pdf"),
                        dados_projeto
                    )

                    if documentos:
                        st.success(f"✅ {len(documentos)} anuência(s) INCRA gerada(s)!")

                        col1, col2 = st.columns(2)

                        with col1:
                            st.markdown("**Download Individual:**")
                            for nome_confrontante, buffer in documentos:
                                st.download_button(
                                    label=f"📥 {nome_confrontante}",
                                    data=buffer.getvalue(),
                                    file_name=f"ANUENCIA_INCRA_{sanitizar_nome_arquivo(nome_confrontante.upper())}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"incra_{nome_confrontante}"
                                )

                        with col2:
                            st.markdown("**Download em Lote (ZIP):**")
                            zip_buffer = GeradorIncraCompleto.gerar_zip_anuencias(documentos)
                            st.download_button(
                                label="📦 Baixar Todas (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name="ANUENCIAS_INCRA.zip",
                                mime="application/zip",
                                key="incra_zip"
                            )
                    else:
                        st.warning("⚠️ Nenhuma anuência INCRA foi gerada.")
                else:
                    st.warning("⚠️ Nenhum memorial foi processado. Volte à aba 1 e processe um arquivo.")
            except Exception as e:
                st.error(f"❌ Erro ao gerar anuências INCRA: {str(e)}")
                logger.error(f"Erro INCRA: {str(e)}")
        else:
            st.info("💡 Processe um memorial na aba 1 para gerar anuências INCRA.")

    # ============================================================
    # ABA 4: REQUERIMENTO DE CARTÓRIO
    # ============================================================
    with tab_requerimento:
        st.markdown("### 🏛️ Gerador de Requerimento de Cartório")

        st.markdown("**Preencha os dados dos requerentes:**")

        col1, col2 = st.columns(2)
        with col1:
            requerente_1_nome = st.text_input("Nome Completo (Requerente 1):", key="req_nome_1")
            requerente_1_cpf = st.text_input("CPF (Requerente 1):", key="req_cpf_1")
            requerente_1_endereco = st.text_area("Endereço (Requerente 1):", key="req_end_1")

        with col2:
            requerente_2_nome = st.text_input("Nome Completo (Requerente 2):", key="req_nome_2")
            requerente_2_cpf = st.text_input("CPF (Requerente 2):", key="req_cpf_2")
            requerente_2_endereco = st.text_area("Endereço (Requerente 2):", key="req_end_2")

        st.markdown("---")
        st.markdown("**Dados do Imóvel:**")

        imovel_descricao = st.text_area("Descrição do Imóvel:", key="req_imovel")
        municipio = st.text_input("Município:", value=st.session_state.get("painel_cliente_local", ""), key="req_municipio")
        cartorio_nome = st.text_input("Nome do Cartório:", key="req_cartorio")

        if st.button("📝 Gerar Requerimento", type="primary", use_container_width=True):
            try:
                dados_requerimento = {
                    "requerente_1_nome": requerente_1_nome,
                    "requerente_1_cpf": requerente_1_cpf,
                    "requerente_1_endereco": requerente_1_endereco,
                    "requerente_2_nome": requerente_2_nome,
                    "requerente_2_cpf": requerente_2_cpf,
                    "requerente_2_endereco": requerente_2_endereco,
                    "imovel_descricao": imovel_descricao,
                    "municipio": municipio,
                    "cartorio_nome": cartorio_nome,
                    "tecnico_nome": st.session_state.get("painel_tecnico_nome", ""),
                    "tecnico_cfta": st.session_state.get("painel_tecnico_cfta", ""),
                }

                gerador_cartorio = GeradorRequerimentoCartorio()
                arquivo_requerimento = gerador_cartorio.gerar_documento(dados_requerimento)

                st.download_button(
                    label="📥 BAIXAR REQUERIMENTO (.DOCX)",
                    data=arquivo_requerimento.getvalue(),
                    file_name=f"REQUERIMENTO_CARTORIO_{sanitizar_nome_arquivo(municipio.upper())}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="btn_download_requerimento"
                )
                st.success("✅ Requerimento gerado com sucesso!")

            except Exception as e:
                st.error(f"❌ Erro ao gerar requerimento: {str(e)}")
                logger.error(f"Erro ao gerar requerimento: {str(e)}")


if __name__ == "__main__":
    main()
