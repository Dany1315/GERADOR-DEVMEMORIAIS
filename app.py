#"""
#GERADOR DE MEMORIAL DESCRITIVO - Versão 6.2 (UI/UX Premium All-Green Edition com Módulo de Anuências)
#Refatoração visual focada em experiência do usuário e design corporativo com paleta verde escuro integral.
#"""

import io
import logging
import time
from typing import Optional, List, Dict, Any
from datetime import datetime
import json  # Adicionado para evitar erro caso ocorra o except json.JSONDecodeError

import streamlit as st
import pandas as pd
import google.generativeai as genai

from config import (
    GEMINI_CONFIG,
    PROCESSAMENTO_CONFIG,
    DOCUMENTO_CONFIG,
    EMPRESA_CONFIG,
    TECNICO_CONFIG,
    CLIENTE_CONFIG,
    VERSAO_APP,
    DESCRICAO_VERSAO
)
from utils import (
    validar_arquivo_pdf,
    validar_texto_entrada,
    criar_logger,
    gerar_relatorio_processamento,
    sanitizar_nome_arquivo,
    formatar_tempo_decorrido
)
from processador import ProcessadorMemorial
from gerador_word import GeradorMemorialWord

# ==========================================
# CONFIGURAÇÃO
# ==========================================

logger = criar_logger(__name__)

st.set_page_config(
    page_title="Gerador de Memorial - Gleba A",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================

def configurar_gemini() -> bool:
    """Configura conexão com API Gemini."""
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
        
        if not api_key:
            logger.error("GEMINI_API_KEY não encontrada nos Secrets")
            return False
        
        genai.configure(api_key=api_key)
        logger.info("✅ Gemini configurado com sucesso")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao configurar Gemini: {str(e)}")
        return False


# ==========================================
# INTERFACE STREAMLIT (UI/UX PREMIUM)
# ==========================================

def main():
    """Função principal da aplicação com interface aprimorada."""
    
    # 🎨 ESTILIZAÇÃO CUSTOMIZADA (CSS)
    st.markdown("""
        <style>
            /* Importar fonte moderna */
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;600;700&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Inter', sans-serif;
            }
            
            /* Banner de Cabeçalho Corporativo - Verde Escuro Elegante */
            .hero-container {
                background: linear-gradient(135deg, #064e3b 0%, #022c22 100%);
                padding: 2.5rem;
                border-radius: 16px;
                color: #ffffff;
                margin-bottom: 2rem;
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
                border-left: 6px solid #10b981;
            }
            .hero-title {
                font-size: 2.2rem;
                font-weight: 700;
                margin-bottom: 0.5rem;
                letter-spacing: -0.025em;
            }
            .hero-subtitle {
                font-size: 1.1rem;
                color: #a7f3d0;
                font-weight: 300;
            }
            
            /* Ajustes finos nos cards e botões - Verde Escuro Premium */
            div.stButton > button:first-child {
                background: linear-gradient(90deg, #065f46 0%, #022c22 100%);
                color: white;
                border: none;
                padding: 0.75rem 1.5rem;
                font-weight: 600;
                font-size: 1.1rem;
                border-radius: 8px;
                transition: all 0.3s ease;
                box-shadow: 0 4px 6px -1px rgba(6, 95, 70, 0.2);
            }
            div.stButton > button:first-child:hover {
                transform: translateY(-2px);
                box-shadow: 0 10px 15px -3px rgba(6, 95, 70, 0.3);
                background: linear-gradient(90deg, #047857 0%, #064e3b 100%);
            }
            
            /* Customização de alertas e informativos */
            .stAlert {
                border-radius: 8px !important;
                border: none !important;
            }
            
            /* 🟢 SIDEBAR PREMIUM VERDE ESCURO 🟢 */
            section[data-testid="stSidebar"] {
                background: linear-gradient(180deg, #022c22 0%, #064e3b 100%) !important;
                border-right: 1px solid #047857;
            }
            
            /* Forçar textos, labels e títulos da sidebar a ficarem claros */
            section[data-testid="stSidebar"] .stMarkdown, 
            section[data-testid="stSidebar"] p, 
            section[data-testid="stSidebar"] label, 
            section[data-testid="stSidebar"] h1, 
            section[data-testid="stSidebar"] h2, 
            section[data-testid="stSidebar"] h3, 
            section[data-testid="stSidebar"] h4, 
            section[data-testid="stSidebar"] h5, 
            section[data-testid="stSidebar"] h6,
            section[data-testid="stSidebar"] span {
                color: #f8fafc !important;
            }

            /* Pequeno ajuste nos subtítulos explicativos da sidebar */
            section[data-testid="stSidebar"] div[style*="color:#64748b"] {
                color: #a7f3d0 !important;
            }

            /* Estilo dos inputs dentro da sidebar para melhor contraste */
            section[data-testid="stSidebar"] input {
                background-color: #042f2e !important;
                color: #ffffff !important;
                border: 1px solid #115e59 !important;
            }

            /* Customização das Abas (Tabs) internas da sidebar */
            section[data-testid="stSidebar"] button[data-baseweb="tab"] {
                color: #a7f3d0 !important;
            }
            section[data-testid="stSidebar"] button[aria-selected="true"] {
                color: #ffffff !important;
                border-bottom-color: #10b981 !important;
            }
            
            /* Estilização para métricas de validação */
            [data-testid="stMetricValue"] {
                font-size: 1.8rem;
                font-weight: 700;
                color: #0f172a;
            }
            
            /* Estilização premium para as abas principais do sistema */
            .stTabs [data-baseweb="tab-list"] {
                gap: 10px;
            }
            .stTabs [data-baseweb="tab"] {
                background-color: #f1f5f9;
                border-radius: 8px 8px 0px 0px;
                padding: 10px 20px;
                font-weight: 600;
            }
            .stTabs [aria-selected="true"] {
                background-color: #064e3b !important;
                color: white !important;
            }

            /* ⚙️ SISTEMA DE ENGRENAGENS EM CSS ⚙️ */
            @keyframes spin-clockwise {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            @keyframes spin-counter {
                0% { transform: rotate(360deg); }
                100% { transform: rotate(0deg); }
            }
            
            .construction-container {
                text-align: center;
                padding: 3.5rem;
                border-radius: 12px;
                background-color: #f8fafc;
                border: 1px dashed #cbd5e1;
                margin-top: 2rem;
            }

            .gears-wrapper {
                position: relative;
                width: 260px;
                height: 180px;
                margin: 0 auto 1.5rem auto;
            }

            .gear {
                position: absolute;
                display: block;
            }

            /* Engrenagem Maior (Esquerda - Verde Escura - Horário) */
            .gear-large {
                width: 110px;
                height: 110px;
                top: 30px;
                left: 20px;
                color: #022c22;
                animation: spin-clockwise 10s linear infinite;
            }

            /* Engrenagem Média (Direita Superior - Verde Médio - Anti-horário) */
            .gear-medium {
                width: 80px;
                height: 80px;
                top: 15px;
                left: 118px;
                color: #047857;
                animation: spin-counter 7.27s linear infinite; /* Velocidade proporcional ao tamanho */
            }

            /* Engrenagem Pequena (Direita Inferior - Verde Claro - Horário) */
            .gear-small {
                width: 55px;
                height: 55px;
                top: 85px;
                left: 175px;
                color: #10b981;
                animation: spin-clockwise 5s linear infinite; /* Mais rápida */
            }
        </style>
    """, unsafe_allow_html=True)

    # 🏛️ HERO HEADER (CABEÇALHO CORPORATIVO)
    st.markdown("""
        <div class="hero-container">
            <div class="hero-title">📐 Portal de Engenharia & Topografia</div>
            <div class="hero-subtitle">Gerador Inteligente de Memoriais Descritivos — Gleba A</div>
        </div>
    """, unsafe_allow_html=True)

    # ==========================================
    # CRIAÇÃO DAS ABAS PRINCIPAIS DO SISTEMA (BEM NO TOPO)
    # ==========================================
    tab_memorial, tab_anuencias = st.tabs(["📝 Memorial Descritivo", "🤝 Anuências"])

    # ==========================================
    # SIDEBAR COM CONFIGURAÇÕES
    # ==========================================
    
    with st.sidebar:
        st.markdown("<h2 style='font-size: 1.5rem; margin-bottom: 1.5rem;'>⚙️ Painel de Controle</h2>", unsafe_allow_html=True)

        # Usando abas na sidebar para reduzir a rolagem vertical infinita
        tab_side_empresa, tab_side_cliente = st.tabs(["🏢 Empresa / Técnico", "👥 Cliente & Imóvel"])

        with tab_side_empresa:
            st.markdown("### Configurações Institucionais")
            with st.expander("📋 Dados da Empresa", expanded=True):
                empresa_nome = st.text_input(
                    "Nome da Empresa",
                    value=EMPRESA_CONFIG.NOME,
                    help="Nome completo da empresa"
                )
                empresa_endereco = st.text_input(
                    "Endereço",
                    value=EMPRESA_CONFIG.ENDERECO,
                    help="Endereço completo com CEP"
                )
                empresa_telefone = st.text_input(
                    "Telefone",
                    value=EMPRESA_CONFIG.TELEFONE,
                    help="Telefone de contato"
                )
                empresa_email = st.text_input(
                    "Email",
                    value=EMPRESA_CONFIG.EMAIL,
                    help="Email para contato"
                )

            with st.expander("👤 Responsável Técnico", expanded=True):
                tecnico_nome = st.text_input(
                    "Nome do Técnico",
                    value=TECNICO_CONFIG.NOME,
                    help="Nome completo do técnico"
                )
                tecnico_cfta = st.text_input(
                    "CFTA",
                    value=TECNICO_CONFIG.CFTA,
                    help="Número do CFTA"
                )

        with tab_side_cliente:
            st.markdown("### Especificações do Projeto")
            
            cliente_imovel = st.text_input(
                "📍 Identificação do Imóvel",
                value=CLIENTE_CONFIG.IMOVEL,
                help="Tipo/identificação do imóvel"
            )
            cliente_proprietario = st.text_input(
                "👤 Proprietário",
                value=CLIENTE_CONFIG.PROPRIETARIO,
                help="Nome completo do proprietário"
            )
            cliente_local = st.text_input(
                "🗺️ Município / Localidade",
                value=CLIENTE_CONFIG.LOCAL,
                help="Localização/município"
            )
            
            col_area, col_per = st.columns(2)
            with col_area:
                cliente_area = st.text_input(
                    "📐 Área (ha)",
                    value=CLIENTE_CONFIG.AREA,
                    help="Área total em hectares"
                )
            with col_per:
                cliente_perimetro = st.text_input(
                    "🏃 Perímetro (m)",
                    value=CLIENTE_CONFIG.PERIMETRO,
                    help="Perímetro total em metros"
                )

        st.markdown("---")
        st.markdown("### 🤖 Ajustes do Motor de IA")
        
        nome_modelo = st.selectbox(
            "Modelo Gemini",
            options=list(GEMINI_CONFIG.MODELOS_DISPONIVEIS.keys()),
            index=0,
            help="Modelos 'Flash': rápidos e econômicos. Modelos 'Pro': maior capacidade analítica."
        )

        with st.expander("🖼️ Parâmetros Técnicos de Imagem", expanded=False):
            dpi_conversao = st.slider(
                "Resolução de Leitura (DPI)",
                min_value=PROCESSAMENTO_CONFIG.DPI_MINIMO,
                max_value=PROCESSAMENTO_CONFIG.DPI_MAXIMO,
                value=PROCESSAMENTO_CONFIG.DPI_PADRAO,
                step=50,
                help="DPI maior aumenta a precisão da leitura, mas torna o processamento mais lento. Padrão: 250 DPI."
            )
            tamanho_max = st.slider(
                "Upload Máximo (MB)",
                min_value=10,
                max_value=100,
                value=PROCESSAMENTO_CONFIG.TAMANHO_MAX_PDF_MB,
                step=10,
                help="Limite de segurança para o tamanho dos arquivos enviados."
            )

        st.markdown(
            "<div style='font-size:0.8rem; color:#a7f3d0; margin-top:1.5rem;'>💡 As definições acima serão replicadas em todos os documentos processados no decorrer desta sessão ativa.</div>",
            unsafe_allow_html=True
        )

    # ------------------------------------------
    # ABA 1: MEMORIAL DESCRITIVO
    # ------------------------------------------
    with tab_memorial:
        # Info sobre versão recolhida em aba ou expander limpo
        with st.expander("ℹ️ Detalhes da Plataforma & Recursos Ativos", expanded=False):
            st.markdown(f"""
            **Gleba A Processor** — `Versão {VERSAO_APP}` — *{DESCRICAO_VERSAO}*
            * **Visão Computacional:** Processamento multimodal via IA (sem dependências locais pesadas).
            * **Infraestrutura:** Execução direta em nuvem com tolerância a falhas (*exponential backoff*).
            * **Saída:** Geração automatizada de relatórios técnicos de alta precisão em formato Microsoft Word (`.docx`).
            """)

        st.markdown("### 📥 Entrada de Dados do Memorial")
        
        # Separando o upload de PDF e a colagem manual por abas para limpar a tela
        tab_pdf, tab_manual = st.tabs(["📁 Processamento de Arquivos PDF", "📝 Colagem de Texto Manual"])

        with tab_pdf:
            st.write("Insira os documentos técnicos vetorizados da Gleba A para processamento inteligente.")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("<div style='font-weight: 600; margin-bottom: 0.5rem;'>📋 Planta e Confrontantes</div>", unsafe_allow_html=True)
                pdf_planta = st.file_uploader(
                    "Arraste ou selecione o PDF da Planta:",
                    type=["pdf"],
                    key="planta",
                    help="PDF contendo a relação de confrontantes por intervalos de pontos"
                )
                if pdf_planta:
                    valido, msg = validar_arquivo_pdf(pdf_planta, tamanho_max)
                    if valido:
                        st.success(f"{msg}")
                    else:
                        st.error(f"{msg}")

            with col2:
                st.markdown("<div style='font-weight: 600; margin-bottom: 0.5rem;'>📊 Tabela de Roteiro Perimétrico</div>", unsafe_allow_html=True)
                pdf_roteiro = st.file_uploader(
                    "Arraste ou selecione o PDF do Roteiro:",
                    type=["pdf"],
                    key="roteiro",
                    help="PDF contendo a tabela com coordenadas, azimutes e distâncias"
                )
                if pdf_roteiro:
                    valido, msg = validar_arquivo_pdf(pdf_roteiro, tamanho_max)
                    if valido:
                        st.success(f"{msg}")
                    else:
                        st.error(f"{msg}")

        with tab_manual:
            st.write("Cole os dados brutos de texto caso não possua os arquivos PDF em mãos.")
            col1, col2 = st.columns(2)
            with col1:
                texto_planta_manual = st.text_area(
                    "Texto copiado da PLANTA (Confrontantes):",
                    height=150,
                    key="texto_planta",
                    placeholder="Exemplo: De 1 para 2 confronta com João da Silva..."
                )
            with col2:
                texto_roteiro_manual = st.text_area(
                    "Texto copiado do ROTEIRO (Tabela técnica):",
                    height=150,
                    key="texto_roteiro",
                    placeholder="Exemplo: PONTO N E AZIMUTE DISTANCIA..."
                )

        st.markdown("<br>", unsafe_allow_html=True)

        # DETECÇÃO DE ENTRADA & PROCESSAMENTO
        tem_pdfs = pdf_planta and pdf_roteiro
        tem_textos = texto_planta_manual and texto_roteiro_manual
        
        if tem_pdfs or tem_textos:
            st.markdown("---")
            st.markdown("<h3 style='text-align: center;'>⚡ Pronto para Processar!</h3>", unsafe_allow_html=True)
            
            if st.button("🔄 ANALISAR DOCUMENTOS E GERAR MEMORIAL DESCRITIVO", type="primary", use_container_width=True):
                tempo_inicio_geral = time.time()
                try:
                    # Etapa 1: Configurar Gemini
                    st.info("🔑 **Etapa 1:** Inicializando conexão com os servidores do Google Gemini...")
                    if not configurar_gemini():
                        st.error("❌ Erro crítico: A variável de ambiente GEMINI_API_KEY não foi encontrada nas configurações do Streamlit.")
                        st.stop()
                    st.success("✅ Conexão com o motor de IA estabelecida!")

                    # Mapear nome amigável para API
                    nome_modelo_api = GEMINI_CONFIG.MODELOS_DISPONIVEIS.get(nome_modelo, "gemini-3.5-flash")

                    # Etapa 2: Converter PDFs em imagens
                    st.info("🖼️ **Etapa 2:** Executando renderização de alta definição das páginas...")
                    processador = ProcessadorMemorial(nome_modelo_api)
                    
                    imagens_planta = []
                    imagens_roteiro = []
                    
                    if pdf_planta:
                        progress_planta = st.progress(0)
                        imagens_planta = processador.pdf_para_imagens(
                            pdf_planta,
                            dpi=dpi_conversao,
                            progress_callback=lambda p, msg: progress_planta.progress(p, msg)
                        )
                    
                    if pdf_roteiro:
                        progress_roteiro = st.progress(0)
                        imagens_roteiro = processador.pdf_para_imagens(
                            pdf_roteiro,
                            dpi=dpi_conversao,
                            progress_callback=lambda p, msg: progress_roteiro.progress(p, msg)
                        )
                    st.success("✅ Renderização concluída!")

                    # Etapa 3: Extrair tabela de roteiro
                    st.info("📊 **Etapa 3:** Analisando a planilha de roteiro por vetorização...")
                    if imagens_roteiro:
                        segmentos = processador.extrair_roteiro_com_ia(imagens_roteiro)
                    else:
                        segmentos = processador.parse_tabela_roteiro_texto(texto_roteiro_manual)

                    if not segmentos:
                        st.warning("⚠️ Atenção: Não conseguimos extrair segmentos legíveis do roteiro perimétrico. Revise a qualidade de conversão ou aumente a taxa de DPI.")
                        st.stop()
                    st.success(f"✅ Sucesso! {len(segmentos)} segmentos georreferenciados identificados.")

                    # Etapa 4: Mapear confrontantes
                    st.info("🤖 **Etapa 4:** Mapeando relações espaciais e limites territoriais...")
                    mapeamento = processador.mapear_confrontantes(
                        imagens_planta=imagens_planta if imagens_planta else None,
                        texto_planta=texto_planta_manual if texto_planta_manual else None,
                        texto_roteiro=texto_roteiro_manual if texto_roteiro_manual else None,
                    )
                    st.success(f"✅ {len(mapeamento.regras)} polígonos de confrontação mapeados com sucesso!")

                    # Etapa 5: Vincular confrontantes
                    st.info("🔗 **Etapa 5:** Consolidando dados topográficos e confrontações...")
                    segmentos_vinculados = processador.vincular_confrontantes()
                    st.success("✅ Consolidação de vértices realizada!")

                    # Validar resultado
                    valido, avisos = processador.validar_resultado()
                    if avisos:
                        for aviso in avisos:
                            st.warning(aviso)

                    # Preparar dados finais
                    dados_finais = {
                        "imovel": cliente_imovel,
                        "proprietario": cliente_proprietario,
                        "local": cliente_local,
                        "area": cliente_area,
                        "perimetro": cliente_perimetro,
                        "segmentos": segmentos_vinculados
                    }

                    # Salvar dados na sessão para que fiquem disponíveis para o módulo de Anuências
                    st.session_state["dados_memoriais_processados"] = dados_finais

                    # Resumo de validação
                    st.balloons()
                    st.success("🎉 Memorial estruturado com sucesso!")
                    st.markdown("### 🔍 Validação e Auditoria dos Dados")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        prop_truncado = (dados_finais['proprietario'][:25] + "...") if len(dados_finais['proprietario']) > 25 else dados_finais['proprietario']
                        st.metric("👤 Proprietário", prop_truncado)
                    with col2:
                        st.metric("📐 Área Total Declarada", f"{dados_finais['area']} ha")
                    with col3:
                        st.metric("🏃 Perímetro Estimado", f"{dados_finais['perimetro']} m")

                    st.markdown("<br>", unsafe_allow_html=True)
                    with st.expander("📋 Tabela Gerada: Malha de Confrontação e Poligonais", expanded=True):
                        df_data = []
                        for seg in dados_finais["segmentos"]:
                            df_data.append({
                                "De": seg['de'],
                                "Para": seg['para'],
                                "N": seg['n_y'],
                                "E": seg['e_x'],
                                "Azimute": seg['azimute'],
                                "Distância (m)": seg['distancia'],
                                "Confrontante": seg['confrontante']
                            })
                        df = pd.DataFrame(df_data)
                        st.dataframe(df, use_container_width=True, hide_index=True)
                        st.caption("⚠️ Nota de Responsabilidade: Os dados acima foram estruturados por algoritmos de visão de IA. Sempre revise os resultados antes de protocolar a peça técnica.")

                    # Gerar documento Word
                    st.info("📝 Redigindo arquivo final no padrão Word (.docx)...")
                    dados_empresa = {"nome": empresa_nome, "endereco": empresa_endereco, "telefone": empresa_telefone, "email": empresa_email}
                    dados_tecnico = {"nome": tecnico_nome, "cfta": tecnico_cfta}
                    
                    gerador = GeradorMemorialWord(dados_empresa, dados_tecnico)
                    arquivo_docx = gerador.gerar_documento(dados_finais)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_down1, col_down2 = st.columns(2)
                    
                    with col_down1:
                        nome_arquivo = sanitizar_nome_arquivo(cliente_proprietario.upper())
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        st.download_button(
                            label="📥 BAIXAR MEMORIAL DESCRITIVO (.DOCX)",
                            data=arquivo_docx,
                            file_name=f"MEMORIAL_DESCRITIVO_{nome_arquivo}_{timestamp}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    
                    with col_down2:
                        tempo_fim = time.time()
                        relatorio = gerar_relatorio_processamento(dados_finais, tempo_inicio_geral, tempo_fim, processador.tempo_gemini)
                        st.download_button(
                            label="📥 BAIXAR RELATÓRIO DE EXECUÇÃO (.TXT)",
                            data=relatorio,
                            file_name=f"relatorio_{timestamp}.txt",
                            use_container_width=True
                        )

                    with st.expander("📊 Relatório Detalhado das Operações", expanded=False):
                        st.text(relatorio)

                except ValueError as e:
                    st.error(f"❌ Erro de Validação: {str(e)}")
                    logger.error(f"Erro de validação: {str(e)}")
                except json.JSONDecodeError as e:
                    st.error(f"❌ Erro ao processar resposta da IA: {str(e)}")
                    logger.error(f"Erro JSON: {str(e)}")
                except Exception as e:
                    st.error(f"❌ Erro inesperado: {str(e)}")
                    logger.error(f"Erro geral: {str(e)}", exc_info=True)
                    with st.expander("🔧 Detalhes Técnicos (Debug/Diagnóstico)"):
                        import traceback
                        st.code(traceback.format_exc())
        else:
            st.markdown("""
                <div style="background-color: #f1f5f9; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #064e3b; margin-top: 2rem;">
                    <h4 style="margin-top:0; color: #0f172a;">🏁 Primeiros passos para iniciar</h4>
                    <ol style="margin-bottom:0; color: #334155;">
                        <li>Carregue ambos os arquivos PDFs correspondentes à <b>Planta</b> e ao <b>Roteiro</b> (ou insira-os manualmente na aba ao lado).</li>
                        <li>Certifique-se de preencher as informações cadastrais do Proprietário e do Técnico na barra lateral esquerda.</li>
                        <li>Clique no botão de análise que aparecerá na tela para processar a malha geográfica.</li>
                    </ol>
                </div>
            """, unsafe_allow_html=True)

   # =========================================================================
    # ABA 2: ANUÊNCIAS (INTEGRAÇÃO DIRETA E AUTOMÁTICA)
    # =========================================================================
    with tab_anuencias:
        st.markdown("### 🤝 Geração Automatizada de Declarações de Anuência")
        st.write("Gere as declarações individuais de reconhecimento de limites baseadas nos confrontantes do Memorial processado.")

        # 1. Verifica se existem dados processados vindos do Processador da Aba 1
        # O processador armazena o resultado no st.session_state (ex: st.session_state.dados_finais ou similar)
        # Ajuste a chave abaixo para bater exatamente com a variável onde você guarda os dados finais do memorial
        dados_memoriais = st.session_state.get("dados_finais") or st.session_state.get("dados_processados")

        if not dados_memoriais or "segmentos" not in dados_memoriais:
            st.warning("⚠️ Nenhum dado de memorial foi localizado. Por favor, carregue e processe os PDFs na aba 'Memorial Descritivo' primeiro.")
        else:
            segmentos = dados_memoriais.get("segmentos", [])
            proprietario_principal = dados_memoriais.get("proprietario", CLIENTE_CONFIG.PROPRIETARIO)
            local_principal = dados_memoriais.get("local", CLIENTE_CONFIG.LOCAL)

            # 2. Filtragem inteligente dos confrontantes legítimos (ignora ruas e avenidas)
            termos_ignorados = ["AV.", "RUA", "AVENIDA", "ESTRADA", "PROJEÇÃO", "VALA", "CORREGO"]
            confrontantes_validos = []
            
            for seg in segmentos:
                conf_nome = str(seg.get("confrontante", "")).strip().upper()
                if conf_nome and not any(termo in conf_nome for termo in termos_ignorados):
                    if conf_nome not in confrontantes_validos and "ERRO" not in conf_nome and "NÃO ENCONTRADA" not in conf_nome:
                        confrontantes_validos.append(conf_nome)
            
            confrontantes_validos = sorted(confrontantes_validos)

            if not confrontantes_validos:
                st.info("ℹ️ Nenhum proprietário confrontante nominal elegível foi localizado nos segmentos deste memorial.")
            else:
                st.success(f"🔍 Identificado(s) **{len(confrontantes_validos)}** confrontante(s) apto(s) para assinatura de anuência!")
                
                # Inputs complementares para o termo técnico exigido no modelo físico
                st.markdown("#### 📄 Informações do Termo de Responsabilidade")
                col_trt1, col_trt2 = st.columns(2)
                with col_trt1:
                    trt_numero = st.text_input("Número da TRT / ART correspondente:", value=TECNICO_CONFIG.TRT, key="trt_anuencia_input")
                with col_trt2:
                    cpf_tecnico = st.text_input("CPF do Responsável Técnico:", value="111.985.197-11", key="cpf_tecnico_anuencia")

                st.markdown("---")
                st.write("### 🗂️ Documentos Prontos para Emissão:")

                # Importa o módulo isolado que gerencia a inteligência artificial do Gemini e o Word
                from gerador_anuencias import GeradorAnuenciaWord

                # Instancia o gerador passando os dicionários de configuração padrão do seu site
                dados_empresa_dict = {
                    "nome": EMPRESA_CONFIG.NOME,
                    "endereco": EMPRESA_CONFIG.ENDERECO,
                    "telefone": EMPRESA_CONFIG.TELEFONE,
                    "email": EMPRESA_CONFIG.EMAIL
                }
                dados_tecnico_dict = {
                    "nome": TECNICO_CONFIG.NOME,
                    "cfta": TECNICO_CONFIG.CFTA,
                    "trt": trt_numero,
                    "cpf": cpf_tecnico
                }
                
                gerador_anuencia_modulo = GeradorAnuenciaWord(dados_empresa_dict, dados_tecnico_dict)

                # Loop dinâmico criando um painel de controle individual para cada confrontante
                for idx, conf in enumerate(confrontantes_validos):
                    # Filtra apenas a malha vetorial que pertence àquele confrontante específico
                    segmentos_do_confrontante = [s for s in segmentos if str(s.get("confrontante", "")).strip().upper() == conf]
                    
                    with st.expander(f"👤 Declaração de Limites: {conf}", expanded=True):
                        col_dados, col_acao = st.columns([3, 1])
                        
                        with col_dados:
                            st.markdown(f"**Vértices abrangidos:** Do `{segmentos_do_confrontante[0]['de']}` ao `{segmentos_do_confrontante[-1]['para']}`")
                            
                            # Mostra uma mini-tabela técnica do trecho no próprio site
                            df_trecho = pd.DataFrame([{
                                "De": s.get('de'), 
                                "Para": s.get('para'), 
                                "Azimute": s.get('azimute'), 
                                "Distância (m)": s.get('distancia')
                            } for s in segmentos_do_confrontante])
                            st.dataframe(df_trecho, use_container_width=True, hide_index=True)
                        
                        with col_acao:
                            st.write("") # Ajuste de alinhamento vertical
                            
                            # Executa a chamada do Gemini e gera o .docx binário ao clicar
                            if st.button(f"⚡ Estruturar via Gemini", key=f"btn_gemini_an_{idx}"):
                                with st.spinner("🤖 O Gemini está analisando o perímetro e redigindo o trecho..."):
                                    dados_anuencia_payload = {
                                        "proprietario": proprietario_principal,
                                        "confrontante": conf,
                                        "local": local_principal,
                                        "segmentos": segmentos_do_confrontante
                                    }
                                    
                                    # Invoca a lógica do gerador_anuencias.py
                                    buffer_docx = gerador_anuencia_modulo.gerar_documento(dados_anuencia_payload)
                                    st.session_state[f"buffer_anuencia_{idx}"] = buffer_docx.getvalue()
                                    st.success("Redação técnica concluída!")

                            # Se o documento já foi gerado na memória do site, habilita o botão de Download
                            if f"buffer_anuencia_{idx}" in st.session_state:
                                nome_arquivo_limpo = sanitizar_nome_arquivo(conf).upper()
                                st.download_button(
                                    label="📥 Baixar Word (.docx)",
                                    data=st.session_state[f"buffer_anuencia_{idx}"],
                                    file_name=f"ANUENCIA_{nome_arquivo_limpo}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"download_an_btn_{idx}",
                                    use_container_width=True
                                )
