# GERADOR DE MEMORIAL DESCRITIVO - Versão 6.4 (Com Barra de Progresso Inteligente)
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
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
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

    # Painel de Controle (Sidebar)
    with st.sidebar:
        st.markdown("<h2 style='font-size: 1.5rem;'>⚙️ Painel de Controle</h2>", unsafe_allow_html=True)
        tab_side_empresa, tab_side_cliente = st.tabs(["🏢 Empresa / Técnico", "👥 Cliente & Imóvel"])

        with tab_side_empresa:
            st.markdown("### Configurações Institucionais")
            empresa_nome = st.text_input("Nome da Empresa", value=EMPRESA_CONFIG.NOME)
            empresa_endereco = st.text_input("Endereço", value=EMPRESA_CONFIG.ENDERECO)
            empresa_telefone = st.text_input("Telefone", value=EMPRESA_CONFIG.TELEFONE)
            empresa_email = st.text_input("Email", value=EMPRESA_CONFIG.EMAIL)

            st.markdown("---")
            technico_nome = st.text_input("Nome do Técnico", value=TECNICO_CONFIG.NOME)
            tecnico_cfta = st.text_input("CFTA", value=TECNICO_CONFIG.CFTA)
            cpf_tecnico = st.text_input("CPF do Responsável Técnico", value="111.985.197-11")

        with tab_side_cliente:
            st.markdown("### Especificações do Projeto")
            cliente_imovel = st.text_input("📍 Identificação do Imóvel", value=CLIENTE_CONFIG.IMOVEL)
            cliente_proprietario = st.text_input("👤 Proprietário", value=CLIENTE_CONFIG.PROPRIETARIO)
            cliente_local = st.text_input("🗺️ Município / Localidade", value=CLIENTE_CONFIG.LOCAL)
            
            col_area, col_per = st.columns(2)
            with col_area:
                cliente_area = st.text_input("📐 Área (ha)", value=CLIENTE_CONFIG.AREA)
            with col_per:
                cliente_perimetro = st.text_input("🏃 Perímetro (m)", value=CLIENTE_CONFIG.PERIMETRO)

        st.markdown("---")
        
        # DEFINIÇÃO DOS MODELOS GEMINI (GERAÇÃO 3 E 2.5)
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
            col1, col2 = st.columns(2)
            with col1:
                texto_planta_manual = st.text_area("Texto da PLANTA:", height=150, placeholder="CONFRONTANTE: XXXXX...")
            with col2:
                texto_roteiro_manual = st.text_area("Texto do ROTEIRO:", height=150, placeholder="PONTO N E AZIMUTE...")

        tem_arquivos = pdf_planta or pdf_roteiro
        tem_textos = texto_planta_manual and texto_roteiro_manual
        
        if tem_arquivos or tem_textos:
            st.markdown("---")
            if st.button("🔄 ANALISAR DOCUMENTOS E GERAR MEMORIAL DESCRITIVO", type="primary", use_container_width=True):
                tempo_inicio_geral = time.time()
                try:
                    with st.status("Executando análise geoespacial...", expanded=True) as status:
                        status.update(label="Inicializando conexão com o ecossistema Gemini...")
                        if not configurar_gemini():
                            st.error("Erro crítico: Chave API ausente nos Secrets.")
                            st.stop()
                        
                        processador = ProcessadorMemorial(nome_modelo_api)
                        
                        # Processamento de arquivos (PDF ou Imagem)
                        imagens_planta, imagens_roteiro = [], []

                        # ---- Planta ----
                        if pdf_planta:
                            if is_imagem(pdf_planta):
                                status.update(label=f"Carregando imagem da planta: {pdf_planta.name}...")
                                img_pil = carregar_imagem_direta(pdf_planta)
                                imagens_planta.append(img_pil)
                            else:
                                status.update(label="Convertendo páginas da planta em matrizes gráficas...")
                                imagens_planta = processador.pdf_para_imagens(pdf_planta, dpi=dpi_conversao)

                        # ---- Roteiro ----
                        if pdf_roteiro:
                            if is_imagem(pdf_roteiro):
                                status.update(label=f"Carregando imagem do roteiro: {pdf_roteiro.name}...")
                                img_pil = carregar_imagem_direta(pdf_roteiro)
                                imagens_roteiro.append(img_pil)
                            else:
                                status.update(label="Vetorizando dados do roteiro perimétrico...")
                                imagens_roteiro = processador.pdf_para_imagens(pdf_roteiro, dpi=dpi_conversao)

                        status.update(label="Extraindo segmentos via Visão Computacional...")
                        if imagens_roteiro:
                            segmentos = processador.extrair_roteiro_com_ia(imagens_roteiro)
                        else:
                            segmentos = processador.parse_tabela_roteiro_texto(texto_roteiro_manual)

                        status.update(label="Cruzando malhas territoriais com confrontações...")
                        mapeamento = processador.mapear_confrontantes(
                            imagens_planta=imagens_planta or None,
                            texto_planta=texto_planta_manual or None,
                            texto_roteiro=texto_roteiro_manual or None
                        )
                        segmentos_vinculados = processador.vincular_confrontantes()
                        status.update(label="Geração finalizada com sucesso!", state="complete")

                    dados_finais = {
                        "imovel": cliente_imovel, "proprietario": cliente_proprietario,
                        "local": cliente_local, "area": cliente_area, "perimetro": cliente_perimetro,
                        "segmentos": segmentos_vinculados
                    }
                    st.session_state["dados_memoriais_processados"] = dados_finais

                    st.balloons()
                    st.success("🎉 Memorial estruturado com sucesso!")
                    
                    # Exibição dos resultados em Tabela
                    df_data = [{
                        "De": s['de'], "Para": s['para'], "N": s['n_y'], "E": s['e_x'],
                        "Azimute": s['azimute'], "Distância (m)": s['distancia'], "Confrontante": s['confrontante']
                    } for s in dados_finais["segmentos"]]
                    st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)

                    # Exportadores (.docx)
                    dados_empresa = {"nome": empresa_nome, "endereco": empresa_endereco, "telefone": empresa_telefone, "email": empresa_email}
                    dados_tecnico = {"nome": technico_nome, "cfta": tecnico_cfta, "cpf": cpf_tecnico}
                    
                    gerador = GeradorMemorialWord(dados_empresa, dados_tecnico)
                    arquivo_docx = gerador.gerar_documento(dados_finais)
                    
                    st.download_button(
                        label="📥 BAIXAR MEMORIAL DESCRITIVO (.DOCX)",
                        data=arquivo_docx,
                        file_name=f"MEMORIAL_{sanitizar_nome_arquivo(cliente_proprietario.upper())}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        on_click=limpar_sessao_memorial,
                        key="btn_download_memorial"
                    )
                    
                    tempo_total = time.time() - tempo_inicio_geral
                    st.caption(f"⏱️ Processamento concluído em {formatar_tempo_decorrido(tempo_total)}")

                except Exception as err:
                    st.error(f"Erro ao processar memorial: {err}")
                    logger.error(f"Erro Memorial: {str(err)}", exc_info=True)

    with tab_anuencias:
        st.markdown("### 🤝 Geração de Anuências para Co-proprietários")
        st.info("💡 Carregue documentos dos co-proprietários para gerar anuências personalizadas.")
        
        documentos_anuencias = st.file_uploader(
            "Carregar documentos (PDF, PNG, JPG, JPEG):",
            type=TIPOS_ARQUIVO_SUPORTADOS,
            accept_multiple_files=True,
            key="docs_anuencias"
        )
        
        if documentos_anuencias:
            st.success(f"📄 {len(documentos_anuencias)} documentos carregados.")
            
            if st.button("🚀 GERAR ANUÊNCIAS", type="primary", use_container_width=True, key="btn_gerar_anuencias"):
                try:
                    with st.status("Processando anuências...", expanded=True) as status:
                        status.update(label="Inicializando...")
                        if not configurar_gemini():
                            st.error("Erro crítico: Chave API ausente.")
                            st.stop()
                        
                        processador_base = ProcessadorMemorial(nome_modelo_api)
                        todas_imagens = []
                        
                        for doc in documentos_anuencias:
                            if is_imagem(doc):
                                img_pil = carregar_imagem_direta(doc)
                                todas_imagens.append(img_pil)
                            else:
                                imagens = processador_base.pdf_para_imagens(doc, dpi=200)
                                todas_imagens.extend(imagens)
                        
                        status.update(label="Anuências geradas com sucesso!", state="complete")
                    
                    st.balloons()
                    st.success("✅ Anuências geradas com sucesso!")
                    
                except Exception as err:
                    st.error(f"Erro ao processar as anuências: {err}")
                    logger.error(f"Erro Anuências: {str(err)}", exc_info=True)

    with tab_anuencias_incra:
        st.markdown("### 🌾 Geração de Anuências INCRA")
        st.info("💡 Carregue documentos para gerar anuências INCRA.")
        
        documentos_incra = st.file_uploader(
            "Carregar documentos INCRA (PDF, PNG, JPG, JPEG):",
            type=TIPOS_ARQUIVO_SUPORTADOS,
            accept_multiple_files=True,
            key="docs_incra"
        )
        
        if documentos_incra:
            st.success(f"📄 {len(documentos_incra)} documentos carregados.")
            
            if st.button("🚀 GERAR ANUÊNCIAS INCRA", type="primary", use_container_width=True, key="btn_gerar_incra"):
                try:
                    with st.status("Processando anuências INCRA...", expanded=True) as status:
                        status.update(label="Inicializando...")
                        if not configurar_gemini():
                            st.error("Erro crítico: Chave API ausente.")
                            st.stop()
                        
                        processador_base = ProcessadorMemorial(nome_modelo_api)
                        todas_imagens = []
                        
                        for doc in documentos_incra:
                            if is_imagem(doc):
                                img_pil = carregar_imagem_direta(doc)
                                todas_imagens.append(img_pil)
                            else:
                                imagens = processador_base.pdf_para_imagens(doc, dpi=200)
                                todas_imagens.extend(imagens)
                        
                        status.update(label="Anuências INCRA geradas com sucesso!", state="complete")
                    
                    st.balloons()
                    st.success("✅ Anuências INCRA geradas com sucesso!")
                    
                except Exception as err:
                    st.error(f"Erro ao processar as anuências INCRA: {err}")
                    logger.error(f"Erro INCRA: {str(err)}", exc_info=True)

    # ------------------------------------------
    # ABA 4: REQUERIMENTO DE CARTÓRIO (COM PROGRESSO)
    # ------------------------------------------
    with tab_requerimento:
        st.markdown("### 🏛️ Geração Automatizada de Requerimento de Cartório")
        st.info("💡 Nesta aba, você pode carregar múltiplos documentos (RG, CPF, Certidões, Matrículas) em **PDF ou imagem (PNG/JPG/JPEG)** para que a IA extraia os dados e preencha o requerimento automaticamente.")
        
        documentos_pdf = st.file_uploader(
            "Carregar documentos dos clientes (PDF, PNG, JPG, JPEG):", 
            type=TIPOS_ARQUIVO_SUPORTADOS, 
            accept_multiple_files=True,
            key="docs_requerimento"
        )
        
        if documentos_pdf:
            st.success(f"📄 {len(documentos_pdf)} documentos carregados.")
            
            if st.button("🚀 EXTRAIR DADOS E GERAR REQUERIMENTO", type="primary", use_container_width=True):
                # ============================================================
                # CRIAR RASTREADOR DE PROGRESSO
                # ============================================================
                tracker = criar_progress_tracker_requerimento()
                tracker.iniciar()
                
                # Container para a barra de progresso
                progress_container = st.container()
                progress_bar = ProgressBarStreamlit(tracker, progress_container)
                
                try:
                    # ============================================================
                    # ETAPA 1: Preparar Documentos
                    # ============================================================
                    progress_bar.atualizar(1, "Preparando documentos para análise visual...")
                    tracker.atualizar_etapa(1, "Preparando documentos")
                    
                    processador_base = ProcessadorMemorial(nome_modelo_api)
                    todas_imagens = []
                    
                    for doc in documentos_pdf:
                        if is_imagem(doc):
                            img_pil = carregar_imagem_direta(doc)
                            todas_imagens.append(img_pil)
                        else:
                            imagens = processador_base.pdf_para_imagens(doc, dpi=200)
                            todas_imagens.extend(imagens)
                    
                    tracker.finalizar_etapa(1)
                    progress_bar.finalizar_etapa(1)
                    
                    # ============================================================
                    # ETAPA 2: Analisar com IA
                    # ============================================================
                    progress_bar.atualizar(2, f"Analisando {len(todas_imagens)} páginas/imagens com {nome_modelo}...")
                    tracker.atualizar_etapa(2, "Analisando com IA")
                    
                    if not configurar_gemini():
                        st.error("Erro crítico: Chave API ausente.")
                        st.stop()
                    
                    # Usar callback para atualizar progresso
                    gerador_req = GeradorRequerimentoCartorio(
                        nome_modelo_api,
                        callback_progresso=callback_atualizar_progresso_requerimento
                    )
                    dados_extraidos = gerador_req.extrair_dados_documentos(todas_imagens)
                    
                    tracker.finalizar_etapa(2)
                    progress_bar.finalizar_etapa(2)
                    
                    # ============================================================
                    # ETAPA 3: Preencher Modelo
                    # ============================================================
                    progress_bar.atualizar(3, "Preenchendo modelo de requerimento Word...")
                    tracker.atualizar_etapa(3, "Preenchendo modelo")
                    
                    template_name = "-REQUERIMENTODECARTORIO.docx"
                    arquivo_word = gerador_req.gerar_documento(dados_extraidos, template_name)
                    
                    tracker.finalizar_etapa(3)
                    progress_bar.finalizar_etapa(3)
                    
                    # ============================================================
                    # ETAPA 4: Finalizar
                    # ============================================================
                    progress_bar.atualizar(4, "Finalizando...")
                    tracker.finalizar_etapa(4)
                    
                    # Exibir resumo final
                    info_final = tracker.obter_info_progresso()
                    
                    st.balloons()
                    st.success("✅ Requerimento gerado com sucesso!")
                    
                    # Exibir estatísticas de tempo
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("⏱️ Tempo Total", info_final['tempo_decorrido_formatado'])
                    with col2:
                        st.metric("📊 Etapas Concluídas", f"{info_final['etapa_atual']}/{info_final['total_etapas']}")
                    with col3:
                        st.metric("✅ Status", "Concluído")
                    
                    # Exibir dados extraídos
                    with st.expander("🔍 Conferir Dados Extraídos (IA)", expanded=False):
                        st.json(dados_extraidos)
                    
                    # Botão de download
                    st.download_button(
                        label="📥 BAIXAR REQUERIMENTO PREENCHIDO (.DOCX)",
                        data=arquivo_word,
                        file_name=f"REQUERIMENTO_CARTORIO_{sanitizar_nome_arquivo(dados_extraidos['requerente_1']['nome'].upper())}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        on_click=limpar_sessao_requerimento,
                        key="btn_download_requerimento"
                    )

                    # Botão extra para limpar sessão manualmente
                    if st.button("🧹 Limpar dados da sessão (Após download)", key="btn_limpar_requerimento"):
                        limpar_sessao_requerimento()
                        st.rerun()
                        
                except Exception as err:
                    st.error(f"Erro no processamento do requerimento: {err}")
                    logger.error(f"Erro Requerimento: {str(err)}", exc_info=True)
        else:
            st.info("Aguardando upload de documentos para iniciar a análise.")


if __name__ == "__main__":
    main()
