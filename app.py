"""
""" GERADOR DE MEMORIAL DESCRITIVO - Versão 6.0 Refatoração completa com arquitetura melhorada,
melhor performance e UX  Principais melhorias: ✅ Separação de responsabilidades
(config, utils, processador, gerador_word) ✅
Retry automático com exponential backoff para API Gemini ✅
Validação robusta de arquivos e entrada ✅
Cache de sessão para evitar reprocessamento ✅
Progress bars para operações longas ✅
Relatórios de processamento detalhados ✅
Tratamento aprimorado de erros ✅ 
Type hints completos ✅
Logging estruturado ✅
Documentação inline completa✅
GERADOR DE MEMORIAL DESCRITIVO - Versão 6.0 (UI/UX Premium Edition)
Refatoração visual completa focada em experiência do usuário e design corporativo.
"""

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
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Inter', sans-serif;
            }
            
            /* Banner de Cabeçalho Corporativo */
            .hero-container {
                background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
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
                color: #94a3b8;
                font-weight: 300;
            }
            
            /* Ajustes finos nos cards e botões */
            div.stButton > button:first-child {
                background: linear-gradient(90deg, #10b981 0%, #059669 100%);
                color: white;
                border: none;
                padding: 0.75rem 1.5rem;
                font-weight: 600;
                font-size: 1.1rem;
                border-radius: 8px;
                transition: all 0.3s ease;
                box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.2);
            }
            div.stButton > button:first-child:hover {
                transform: translateY(-2px);
                box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.3);
                background: linear-gradient(90deg, #059669 0%, #047857 100%);
            }
            
            /* Customização de alertas e informativos */
            .stAlert {
                border-radius: 8px !important;
                border: none !important;
            }
            
            /* Sidebar Elegante */
            section[data-testid="stSidebar"] {
                background-color: #f8fafc;
                border-right: 1px solid #e2e8f0;
            }
            
            /* Estilização para métricas de validação */
            [data-testid="stMetricValue"] {
                font-size: 1.8rem;
                font-weight: 700;
                color: #0f172a;
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

    # Info sobre versão recolhida em aba ou expander limpo
    with st.expander("ℹ️ Detalhes da Plataforma & Recursos Ativos", expanded=False):
        st.markdown(f"""
        **Gleba A Processor** — `Versão {VERSAO_APP}` — *{DESCRICAO_VERSAO}*
        * **Visão Computacional:** Processamento multimodal via IA (sem dependências locais pesadas).
        * **Infraestrutura:** Execução direta em nuvem com tolerância a falhas (*exponential backoff*).
        * **Saída:** Geração automatizada de relatórios técnicos de alta precisão em formato Microsoft Word (`.docx`).
        """)

    # ==========================================
    # SIDEBAR COM CONFIGURAÇÕES (REORGANIZADO)
    # ==========================================
    
    with st.sidebar:
        st.markdown("<h2 style='font-size: 1.5rem; color: #0f172a; margin-bottom: 1.5rem;'>⚙️ Painel de Controle</h2>", unsafe_allow_html=True)

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
            "<div style='font-size:0.8rem; color:#64748b; margin-top:1.5rem;'>💡 As definições acima serão replicadas em todos os documentos processados no decorrer desta sessão ativa.</div>",
            unsafe_allow_html=True
        )

    # ==========================================
    # CORPO PRINCIPAL (TABS PARA FLUXO LIMPO)
    # ==========================================

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

    # ==========================================
    # DETECÇÃO DE ENTRADA & PROCESSAMENTO
    # ==========================================

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
                    st.error(
                        "❌ Erro crítico: A variável de ambiente GEMINI_API_KEY não foi encontrada nas configurações do Streamlit."
                    )
                    st.stop()
                st.success("✅ Conexão com o motor de IA estabelecida!")

                # Mapear nome amigável para API
                nome_modelo_api = GEMINI_CONFIG.MODELOS_DISPONIVEIS.get(
                    nome_modelo,
                    "gemini-3.5-flash"
                )

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
                    st.warning(
                        "⚠️ Atenção: Não conseguimos extrair segmentos legíveis do roteiro perimétrico. "
                        "Revise a qualidade de conversão ou aumente a taxa de DPI."
                    )
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

                # Resumo de validação
                st.balloons()
                st.success("🎉 Memorial estruturado com sucesso!")
                
                st.markdown("### 🔍 Validação e Auditoria dos Dados")
                
                # Cards de métricas mais limpos e polidos
                col1, col2, col3 = st.columns(3)
                with col1:
                    prop_truncado = (dados_finais['proprietario'][:25] + "...") if len(dados_finais['proprietario']) > 25 else dados_finais['proprietario']
                    st.metric("👤 Proprietário", prop_truncado)
                with col2:
                    st.metric("📐 Área Total Declarada", f"{dados_finais['area']} ha")
                with col3:
                    st.metric("🏃 Perímetro Estimado", f"{dados_finais['perimetro']} m")

                # Tabela detalhada
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
                    st.caption(
                        "⚠️ Nota de Responsabilidade: Os dados acima foram estruturados por algoritmos de visão de IA. "
                        "Sempre revise os resultados antes de protocolar a peça técnica."
                    )

                # Gerar documento Word
                st.info("📝 Redigindo arquivo final no padrão Word (.docx)...")
                
                dados_empresa = {
                    "nome": empresa_nome,
                    "endereco": empresa_endereco,
                    "telefone": empresa_telefone,
                    "email": empresa_email
                }
                
                dados_tecnico = {
                    "nome": tecnico_nome,
                    "cfta": tecnico_cfta
                }
                
                gerador = GeradorMemorialWord(dados_empresa, dados_tecnico)
                arquivo_docx = gerador.gerar_documento(dados_finais)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Seção de exportação com destaque visual
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
                    # Relatório de processamento
                    tempo_fim = time.time()
                    relatorio = gerar_relatorio_processamento(
                        dados_finais,
                        tempo_inicio_geral,
                        tempo_fim,
                        processador.tempo_gemini
                    )
                    
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
        # Mensagem inicial de tutorial de fluxo (substituindo a info box antiga por algo mais amigável)
        st.markdown("""
            <div style="background-color: #f1f5f9; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #64748b; margin-top: 2rem;">
                <h4 style="margin-top:0; color: #334155;">🏁 Primeiros passos para iniciar</h4>
                <ol style="margin-bottom:0; color: #475569;">
                    <li>Carregue ambos os arquivos PDFs correspondentes à <b>Planta</b> e ao <b>Roteiro</b> (ou insira-os manualmente na aba ao lado).</li>
                    <li>Certifique-se de preencher as informações cadastrais do Proprietário e do Técnico na barra lateral esquerda.</li>
                    <li>Clique no botão de análise que aparecerá na tela para processar a malha geográfica.</li>
                </ol>
            </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
