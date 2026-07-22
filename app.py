# GERADOR DE MEMORIAL DESCRITIVO - Versão 6.3 (Segurança + UI/UX Premium)
import io
import logging
import time
import os
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
from gerador_anuencia_incra import GeradorAnuenciaIncraWord
from gerador_memorial_word import GeradorMemorialWord
from gerador_anuencias import GeradorAnuenciaWord

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
# SEGURANÇA: SISTEMA DE LOGIN
# ============================================================
def tela_login():
    """Exibe a tela de login com verificação de usuário e senha.
    As credenciais são definidas via st.secrets para segurança.
    """
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

    # Botão de logout na sidebar
    with st.sidebar:
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

    # Indicador de usuário logado
    usuario_logado = st.session_state.get("usuario_logado", "Usuário")
    st.caption(f"🔒 Sessão ativa: **{usuario_logado}** | Conexão segura")

    # ============================================================
    # PAINEL DE CONTROLE (SIDEBAR) — AGORA FUNCIONAL
    # ============================================================
    with st.sidebar:
        st.markdown("<h2 style='font-size: 1.5rem;'>⚙️ Painel de Controle</h2>", unsafe_allow_html=True)
        tab_side_empresa, tab_side_cliente = st.tabs(["🏢 Empresa / Técnico", "👥 Cliente & Imóvel"])

        # ✅ NOVO: Inicializar dados do painel lateral em session_state
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

        st.markdown("---")
        
        # DEFINIÇÃO DOS MODELOS GEMINI
        modelos_filtrados = {
            "Gemini 3.5 Flash (Fronteira/Padrão)": "gemini-3.5-flash",
            "Gemini 3.1 Pro (Raciocínio Avançado)": "gemini-3.1-pro",
            "Gemini 3.1 Flash-Lite (Alta Velocidade)": "gemini-3.1-flash-lite",
            "Gemini 2.5 Pro (Estável e Preciso)": "gemini-2.5-pro",
            "Gemini 2.5 Flash (Trabalho Diário)": "gemini-2.5-flash"
        }
        nome_modelo = st.selectbox("Modelo Gemini", options=list(modelos_filtrados.keys()), index=0)
        nome_modelo_api = modelos_filtrados[nome_modelo]
        
        dpi_conversao = st.slider("Resolução (DPI) - usado apenas para PDFs", 100, 400, int(PROCESSAMENTO_CONFIG.DPI_PADRAO), 50)
        tamanho_max = st.slider("Upload Máximo (MB)", 10, 100, int(PROCESSAMENTO_CONFIG.TAMANHO_MAX_PDF_MB), 10)

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
            st.info("💡 Aceita arquivos **PDF**, **PNG**, **JPG** ou **JPEG**. Imagens são enviadas diretamente à IA; PDFs são convertidos internamente em imagens antes da análise.")
            col1, col2 = st.columns(2)
            with col1:
                pdf_planta = st.file_uploader(
                    "Planta (Confrontantes) — PDF ou Imagem:",
                    type=TIPOS_ARQUIVO_SUPORTADOS,
                    key="planta"
                )
                if pdf_planta:
                    if is_imagem(pdf_planta):
                        st.caption(f"✅ Imagem carregada: `{pdf_planta.name}`")
                    else:
                        st.caption(f"✅ PDF carregado: `{pdf_planta.name}` (será convertido a {dpi_conversao} DPI)")
            with col2:
                pdf_roteiro = st.file_uploader(
                    "Roteiro (Tabela) — PDF ou Imagem:",
                    type=TIPOS_ARQUIVO_SUPORTADOS,
                    key="roteiro"
                )
                if pdf_roteiro:
                    if is_imagem(pdf_roteiro):
                        st.caption(f"✅ Imagem carregada: `{pdf_roteiro.name}`")
                    else:
                        st.caption(f"✅ PDF carregado: `{pdf_roteiro.name}` (será convertido a {dpi_conversao} DPI)")

        with tab_manual:
            col_planta, col_roteiro = st.columns(2)
            with col_planta:
                st.markdown("**Planta (Confrontantes) - Texto:**")
                texto_planta_manual = st.text_area("Cole o texto da planta aqui:", height=150, key="texto_planta")
            with col_roteiro:
                st.markdown("**Roteiro (Tabela) - Texto:**")
                texto_roteiro_manual = st.text_area("Cole o texto do roteiro aqui:", height=150, key="texto_roteiro")

        # Botão para processar
        if st.button("🔍 Analisar Documentos", type="primary", use_container_width=True):
            if not (pdf_planta or pdf_roteiro or texto_planta_manual or texto_roteiro_manual):
                st.error("❌ Por favor, carregue pelo menos um arquivo ou cole um texto.")
            else:
                with st.spinner("⏳ Processando... Isso pode levar alguns minutos."):
                    try:
                        # Configurar Gemini
                        if not configurar_gemini():
                            st.error("❌ Erro ao configurar Gemini API.")
                            return

                        # Inicializar processador
                        processador = ProcessadorMemorial(nome_modelo_api)

                        # ============================================================
                        # PROCESSAMENTO DE PLANTA
                        # ============================================================
                        imagens_planta = []
                        if pdf_planta:
                            if is_imagem(pdf_planta):
                                imagens_planta = [carregar_imagem_direta(pdf_planta)]
                            else:
                                imagens_planta = processador.pdf_para_imagens(pdf_planta, dpi=dpi_conversao)

                        # ============================================================
                        # PROCESSAMENTO DE ROTEIRO
                        # ============================================================
                        imagens_roteiro = []
                        if pdf_roteiro:
                            if is_imagem(pdf_roteiro):
                                imagens_roteiro = [carregar_imagem_direta(pdf_roteiro)]
                            else:
                                imagens_roteiro = processador.pdf_para_imagens(pdf_roteiro, dpi=dpi_conversao)

                        # Extração de roteiro
                        if imagens_roteiro:
                            segmentos = processador.extrair_roteiro_com_ia(imagens_roteiro)
                        elif texto_roteiro_manual:
                            segmentos = processador.parse_tabela_roteiro_texto(texto_roteiro_manual)
                        else:
                            st.error("❌ Nenhum roteiro foi fornecido.")
                            return

                        # Mapeamento de confrontantes
                        processador.mapear_confrontantes(
                            imagens_planta=imagens_planta or None,
                            texto_planta=texto_planta_manual or None,
                            texto_roteiro=texto_roteiro_manual or None
                        )

                        # Vinculação de confrontantes
                        processador.vincular_confrontantes()

                        # Validação
                        eh_valido, avisos = processador.validar_resultado()
                        if avisos:
                            for aviso in avisos:
                                st.warning(aviso)

                        # Armazenar em session_state
                        st.session_state["dados_memoriais_processados"] = processador.segmentos
                        st.session_state["segmentos"] = processador.segmentos

                        # ✅ NOVO: Preparar dados_finais usando dados do painel lateral
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
                            "segmentos": processador.segmentos,
                        }

                        st.session_state["dados_finais"] = dados_finais

                        st.balloons()
                        st.success("🎉 Memorial estruturado com sucesso!")
                        
                        # Exibição dos resultados em Tabela
                        df_data = [{
                            "De": s['de'], "Para": s['para'], "N": s['coord_y'], "E": s['coord_x'],
                            "Azimute": s['azimute'], "Distância (m)": s['distancia'], "Confrontante": s['confrontante']
                        } for s in dados_finais["segmentos"]]
                        st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)

                        # Exportadores (.docx)
                        gerador = GeradorMemorialWord(dados_finais["empresa"], dados_finais["tecnico"])
                        arquivo_docx = gerador.gerar_documento(dados_finais)
                        
                        st.download_button(
                            label="📥 BAIXAR MEMORIAL DESCRITIVO (.DOCX)",
                            data=arquivo_docx,
                            file_name=f"MEMORIAL_{sanitizar_nome_arquivo(st.session_state['painel_cliente_proprietario'].upper())}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="btn_download_memorial"
                        )

                        # Botão para limpar dados
                        if st.button("🧹 Limpar Dados da Sessão", type="secondary", use_container_width=True):
                            limpar_sessao_memorial()
                            st.rerun()

                    except Exception as e:
                        logger.error(f"Erro: {str(e)}", exc_info=True)
                        st.error(f"❌ Erro: {str(e)}")

    # ============================================================
    # ABA 2: ANUÊNCIAS CO-PROPRIETÁRIOS
    # ============================================================
    with tab_anuencias:
        st.markdown("### 🤝 Gerador de Anuências Co-proprietários")
        
        if "dados_memoriais_processados" not in st.session_state:
            st.info("💡 Processe um memorial descritivo primeiro na aba anterior.")
        else:
            segmentos = st.session_state["dados_memoriais_processados"]
            
            # Validar se segmentos é uma lista
            if not isinstance(segmentos, list) or len(segmentos) == 0:
                st.warning("⚠️ Nenhum segmento foi processado. Verifique o memorial descritivo.")
                confrontantes_unicos = []
            else:
                # Extrair confrontantes únicos
                confrontantes_unicos = list(set([s.get("confrontante", "N/A") for s in segmentos if isinstance(s, dict) and s.get("confrontante")]))
            
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
                                # Preparar dados para a anuência
                                segmentos_confrontante = [s for s in st.session_state.get("segmentos", []) if s.get("confrontante", "").upper() == confrontante.upper()]
                                
                                dados_anuencia = {
                                    "confrontante": nome_anuencia,
                                    "proprietario": st.session_state["painel_cliente_proprietario"],
                                    "local": st.session_state["painel_cliente_local"],
                                    "segmentos": segmentos_confrontante
                                }
                                
                                # Gerar documento
                                gerador_anu = GeradorAnuenciaWord(st.session_state.get("dados_finais", {}).get("empresa", {}), st.session_state.get("dados_finais", {}).get("tecnico", {}))
                                arquivo_anuencia = gerador_anu.gerar_documento(dados_anuencia)
                                
                                # Botão de download
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

    # ============================================================
    # ABA 3: ANUÊNCIAS INCRA
    # ============================================================
    with tab_anuencias_incra:
        st.markdown("### 🌾 Gerador de Anuências INCRA")
        
        st.info("💡 Carregue um memorial descritivo na aba anterior para gerar anuências INCRA.")
        
        if "dados_finais" not in st.session_state:
            st.warning("⚠️ Processe um memorial descritivo primeiro na aba 'Memorial Descritivo'.")
        else:
            st.markdown("---")
            st.markdown("**Gerar Anuências INCRA a partir do Memorial Descritivo**")
            
            if st.button("🌾 Gerar Anuências INCRA", type="primary", use_container_width=True):
                try:
                    with st.spinner("⏳ Gerando anuências INCRA... Isso pode levar alguns minutos."):
                        # Preparar dados para o gerador INCRA
                        dados_projeto = {
                            "proprietario": st.session_state["painel_cliente_proprietario"],
                            "cpf_proprietario": st.session_state.get("painel_cpf_tecnico", "092.653.737-76"),
                            "local": st.session_state["painel_cliente_local"],
                            "imovel": st.session_state["painel_cliente_imovel"],
                            "area": st.session_state["painel_cliente_area"],
                            "perimetro": st.session_state["painel_cliente_perimetro"],
                            "comarca": st.session_state["painel_cliente_comarca"],
                            "matricula": st.session_state["painel_cliente_matricula"],
                        }
                        
                        # Usar o gerador INCRA
                        gerador_incra = GeradorAnuenciaIncraWord(
                            st.session_state["dados_finais"]["empresa"],
                            st.session_state["dados_finais"]["tecnico"]
                        )
                        
                        # Se houver arquivo de memorial processado, usar; caso contrário, usar dados estruturados
                        # Aqui assumimos que temos dados já processados
                        documentos_incra = gerador_incra.gerar_documentos_pelo_memorial(
                            conteudo_arquivo=b"",  # Será substituído por dados estruturados
                            nome_arquivo="memorial.docx",
                            dados_projeto=dados_projeto
                        )
                        
                        # Gerar ZIP com os documentos
                        zip_buffer = GeradorAnuenciaIncraWord.gerar_zip_anuencias(
                            documentos_incra,
                            prefixo_arquivo="ANUENCIA_INCRA"
                        )
                        
                        st.download_button(
                            label="📥 BAIXAR ANUÊNCIAS INCRA (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"ANUENCIAS_INCRA_{sanitizar_nome_arquivo(st.session_state['painel_cliente_proprietario'].upper())}.zip",
                            mime="application/zip",
                            use_container_width=True,
                            key="download_incra_zip"
                        )
                        
                        st.success(f"✅ {len(documentos_incra)} anuência(s) INCRA gerada(s) com sucesso!")
                        
                except Exception as e:
                    st.error(f"❌ Erro ao gerar anuências INCRA: {str(e)}")
                    logger.error(f"Erro ao gerar anuências INCRA: {str(e)}", exc_info=True)

    # ============================================================
    # ABA 4: REQUERIMENTO DE CARTÓRIO
    # ============================================================
    with tab_requerimento:
        st.markdown("### 🏛️ Gerador de Requerimento de Cartório")
        
        st.info("💡 Carregue documentos (RG, CPF, Planta INCRA) para gerar o requerimento de cartório.")
        
        # Abas para entrada de dados
        tab_requerimento_upload, tab_requerimento_manual = st.tabs([
            "📁 Upload de Documentos",
            "📝 Preenchimento Manual"
        ])
        
        with tab_requerimento_upload:
            st.markdown("**Carregue as imagens dos documentos:**")
            imagens_requerimento = st.file_uploader(
                "Selecione as imagens dos documentos (RG, CPF, Planta INCRA, etc.):",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key="upload_requerimento"
            )
            
            if imagens_requerimento:
                st.success(f"✅ {len(imagens_requerimento)} imagem(s) carregada(s)")
                
                if st.button("🔍 Extrair Dados dos Documentos", type="primary", use_container_width=True):
                    try:
                        with st.spinner("⏳ Analisando documentos com IA..."):
                            from gerador_requerimento_cartorio import GeradorRequerimentoCartorio
                            
                            # Converter imagens para formato aceito pelo Gemini
                            imagens_gemini = []
                            for img_file in imagens_requerimento:
                                img = Image.open(io.BytesIO(img_file.getvalue()))
                                imagens_gemini.append(img)
                            
                            # Extrair dados
                            gerador_req = GeradorRequerimentoCartorio("gemini-2.5-flash")
                            dados_extraidos = gerador_req.extrair_dados_documentos(imagens_gemini)
                            
                            # Armazenar em session_state
                            st.session_state["dados_extraidos_requerimento"] = dados_extraidos
                            
                            st.success("✅ Dados extraídos com sucesso!")
                            st.json(dados_extraidos)
                    except Exception as e:
                        st.error(f"❌ Erro ao extrair dados: {str(e)}")
                        logger.error(f"Erro ao extrair dados do requerimento: {str(e)}", exc_info=True)
        
        with tab_requerimento_manual:
            st.markdown("**Preencha os dados manualmente:**")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Requerente 1 (Proprietário):**")
                nome_req1 = st.text_input("Nome Completo", key="manual_nome_req1")
                cpf_req1 = st.text_input("CPF", key="manual_cpf_req1")
                rg_req1 = st.text_input("RG", key="manual_rg_req1")
                profissao_req1 = st.text_input("Profissão", key="manual_prof_req1")
            
            with col2:
                st.markdown("**Requerente 2 (Cônjuge - Opcional):**")
                nome_req2 = st.text_input("Nome Completo", key="manual_nome_req2")
                cpf_req2 = st.text_input("CPF", key="manual_cpf_req2")
                rg_req2 = st.text_input("RG", key="manual_rg_req2")
                profissao_req2 = st.text_input("Profissão", key="manual_prof_req2")
            
            st.markdown("---")
            st.markdown("**Dados do Imóvel:**")
            
            col_imovel1, col_imovel2 = st.columns(2)
            with col_imovel1:
                nome_imovel = st.text_input("Nome do Imóvel", key="manual_nome_imovel")
                area_imovel = st.text_input("Área (hectares)", key="manual_area_imovel")
                matricula_imovel = st.text_input("Matrícula", key="manual_matricula_imovel")
            
            with col_imovel2:
                municipio_imovel = st.text_input("Município", key="manual_municipio_imovel")
                comarca_imovel = st.text_input("Comarca", key="manual_comarca_imovel")
                trt_imovel = st.text_input("TRT", key="manual_trt_imovel")
            
            if st.button("💾 Salvar Dados Manualmente", type="primary", use_container_width=True):
                dados_manual = {
                    "requerente_1": {
                        "nome": nome_req1,
                        "cpf": cpf_req1,
                        "rg": rg_req1,
                        "profissao": profissao_req1,
                    },
                    "requerente_2": {
                        "nome": nome_req2 or "XXXXXX",
                        "cpf": cpf_req2 or "XXXXXX",
                        "rg": rg_req2 or "XXXXXX",
                        "profissao": profissao_req2 or "XXXXXX",
                    },
                    "imovel": {
                        "nome": nome_imovel,
                        "area_registrada": area_imovel,
                        "matricula": matricula_imovel,
                        "municipio_imovel": municipio_imovel,
                        "comarca_imovel": comarca_imovel,
                        "trt_numero": trt_imovel,
                    },
                    "comarca": comarca_imovel,
                    "municipio_cliente": municipio_imovel,
                }
                st.session_state["dados_extraidos_requerimento"] = dados_manual
                st.success("✅ Dados salvos com sucesso!")
        
        # Seção de geração do requerimento
        st.markdown("---")
        st.markdown("**Gerar Requerimento de Cartório:**")
        
        if "dados_extraidos_requerimento" not in st.session_state:
            st.info("💡 Extraia ou preencha os dados acima para gerar o requerimento.")
        else:
            # Buscar template
            template_path = os.path.join(os.path.dirname(__file__), "template_requerimento.docx")
            
            if not os.path.exists(template_path):
                st.warning(f"⚠️ Template não encontrado em {template_path}. Verifique se o arquivo 'template_requerimento.docx' existe.")
            else:
                if st.button("🏛️ Gerar Requerimento de Cartório", type="primary", use_container_width=True):
                    try:
                        with st.spinner("⏳ Gerando requerimento..."):
                            from gerador_requerimento_cartorio import GeradorRequerimentoCartorio
                            
                            gerador_req = GeradorRequerimentoCartorio("gemini-2.5-flash")
                            documento_bytes = gerador_req.gerar_documento(
                                st.session_state["dados_extraidos_requerimento"],
                                template_path
                            )
                            
                            st.download_button(
                                label="📥 BAIXAR REQUERIMENTO DE CARTÓRIO",
                                data=documento_bytes,
                                file_name=f"REQUERIMENTO_CARTORIO_{sanitizar_nome_arquivo(st.session_state['dados_extraidos_requerimento'].get('requerente_1', {}).get('nome', 'REQUERENTE').upper())}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                key="download_requerimento"
                            )
                            
                            st.success("✅ Requerimento de cartório gerado com sucesso!")
                            
                            # Botão para limpar dados
                            if st.button("🧹 Limpar Dados da Sessão", type="secondary", use_container_width=True):
                                limpar_sessao_requerimento()
                                st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erro ao gerar requerimento: {str(e)}")
                        logger.error(f"Erro ao gerar requerimento de cartório: {str(e)}", exc_info=True)


if __name__ == "__main__":
    main()
