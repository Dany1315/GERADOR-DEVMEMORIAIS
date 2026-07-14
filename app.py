"""
GERADOR DE MEMORIAL DESCRITIVO - Versão 6.0
Refatoração completa com arquitetura melhorada, melhor performance e UX

Principais melhorias:
✅ Separação de responsabilidades (config, utils, processador, gerador_word)
✅ Retry automático com exponential backoff para API Gemini
✅ Validação robusta de arquivos e entrada
✅ Cache de sessão para evitar reprocessamento
✅ Progress bars para operações longas
✅ Relatórios de processamento detalhados
✅ Tratamento aprimorado de erros
✅ Type hints completos
✅ Logging estruturado
✅ Documentação inline completa
"""

import io
import logging
import time
from typing import Optional, List, Dict, Any
from datetime import datetime

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
    page_title="Gerador de Memorial Descritivo - Gleba A",
    page_icon="📄",
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
# INTERFACE STREAMLIT
# ==========================================

def main():
    """Função principal da aplicação."""
    
    st.title("📄 Processador de Memoriais Descritivos - Gleba A")
    st.write(
        "Insira os dois arquivos da Gleba A para estruturar automaticamente o Memorial "
        "Descritivo com precisão e conformidade técnica."
    )

    # Info sobre versão
    st.info(f"""
    ✅ **Versão {VERSAO_APP}** - {DESCRICAO_VERSAO}
    
    **Funcionalidades:**
    - ☁️ Funciona 100% na nuvem, sem Poppler/Tesseract
    - 📖 Lê PDFs de CAD (VectorDraw, etc.) convertendo em imagem
    - 🤖 Extrai tabelas com visão multimodal do Gemini
    - 🔄 Retry automático com exponential backoff para rate limiting
    - ⚡ Cache de sessão para operações rápidas
    - 📊 Relatórios detalhados de processamento
    - ✔️ Validação robusta de entrada e saída
    """)

    # ==========================================
    # SIDEBAR COM CONFIGURAÇÕES
    # ==========================================
    
    with st.sidebar:
        st.header("⚙️ Configurações")

        with st.expander("📋 Dados da Empresa", expanded=False):
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

        with st.expander("👤 Dados do Técnico Responsável", expanded=False):
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

        with st.expander("🤖 Modelo de IA", expanded=True):
            nome_modelo = st.selectbox(
                "Modelo Gemini",
                options=list(GEMINI_CONFIG.MODELOS_DISPONIVEIS.keys()),
                index=0,
                help="Modelos 'Flash': rápidos e econômicos. "
                     "Modelos 'Pro': maior capacidade para documentos complexos"
            )

        with st.expander("👥 Dados do Cliente", expanded=True):
            cliente_imovel = st.text_input(
                "Imóvel",
                value=CLIENTE_CONFIG.IMOVEL,
                help="Tipo/identificação do imóvel"
            )
            cliente_proprietario = st.text_input(
                "Nome do Proprietário",
                value=CLIENTE_CONFIG.PROPRIETARIO,
                help="Nome completo do proprietário"
            )
            cliente_local = st.text_input(
                "Local",
                value=CLIENTE_CONFIG.LOCAL,
                help="Localização/município"
            )
            cliente_area = st.text_input(
                "Área (ha)",
                value=CLIENTE_CONFIG.AREA,
                help="Área total em hectares"
            )
            cliente_perimetro = st.text_input(
                "Perímetro (m)",
                value=CLIENTE_CONFIG.PERIMETRO,
                help="Perímetro total em metros"
            )

        with st.expander("🖼️ Processamento", expanded=False):
            dpi_conversao = st.slider(
                "Qualidade da imagem (DPI)",
                min_value=PROCESSAMENTO_CONFIG.DPI_MINIMO,
                max_value=PROCESSAMENTO_CONFIG.DPI_MAXIMO,
                value=PROCESSAMENTO_CONFIG.DPI_PADRAO,
                step=50,
                help="DPI maior = leitura mais precisa, porém mais lenta. "
                     "Recomendado: 250 DPI"
            )
            tamanho_max = st.slider(
                "Tamanho máximo de PDF (MB)",
                min_value=10,
                max_value=100,
                value=PROCESSAMENTO_CONFIG.TAMANHO_MAX_PDF_MB,
                step=10,
                help="Limite de tamanho para upload"
            )

        st.info(
            "💡 **Dica:** As configurações acima são usadas em todos os documentos "
            "gerados nesta sessão."
        )

    # ==========================================
    # UPLOAD DE ARQUIVOS
    # ==========================================

    st.subheader("📁 Carregue os Arquivos")

    col1, col2 = st.columns(2)
    
    with col1:
        pdf_planta = st.file_uploader(
            "Carregue o PDF com os DADOS DA PLANTA:",
            type=["pdf"],
            key="planta",
            help="PDF contendo a relação de confrontantes por intervalos de pontos"
        )
        if pdf_planta:
            valido, msg = validar_arquivo_pdf(pdf_planta, tamanho_max)
            if valido:
                st.success(f"✅ {msg}")
            else:
                st.error(f"❌ {msg}")

    with col2:
        pdf_roteiro = st.file_uploader(
            "Carregue o PDF da TABELA DE ROTEIRO PERIMÉTRICO:",
            type=["pdf"],
            key="roteiro",
            help="PDF contendo a tabela com coordenadas, azimutes e distâncias"
        )
        if pdf_roteiro:
            valido, msg = validar_arquivo_pdf(pdf_roteiro, tamanho_max)
            if valido:
                st.success(f"✅ {msg}")
            else:
                st.error(f"❌ {msg}")

    # Alternativa: Colagem de texto
    with st.expander("📝 Alternativa: colar o texto manualmente (opcional)"):
        st.write(
            "Use apenas se preferir não enviar os PDFs, ou como reforço de contexto "
            "além dos PDFs enviados acima."
        )
        col1, col2 = st.columns(2)
        with col1:
            texto_planta_manual = st.text_area(
                "Cole o texto da PLANTA aqui (opcional):",
                height=100,
                key="texto_planta"
            )
        with col2:
            texto_roteiro_manual = st.text_area(
                "Cole o texto do ROTEIRO aqui (opcional):",
                height=100,
                key="texto_roteiro"
            )

    # ==========================================
    # PROCESSAMENTO
    # ==========================================

    tem_pdfs = pdf_planta and pdf_roteiro
    tem_textos = texto_planta_manual and texto_roteiro_manual
    
    if tem_pdfs or tem_textos:
        if st.button("🔄 Analisar Documentos e Gerar Memorial", type="primary", use_container_width=True):
            
            tempo_inicio_geral = time.time()
            
            try:
                # Etapa 1: Configurar Gemini
                st.info("🔑 Etapa 1: Configurando API Gemini...")
                if not configurar_gemini():
                    st.error(
                        "❌ Erro: Chave GEMINI_API_KEY não configurada nos Streamlit Secrets. "
                        "Configure a chave e tente novamente."
                    )
                    st.stop()
                st.success("✅ Gemini configurado")

                # Mapear nome amigável para API
                nome_modelo_api = GEMINI_CONFIG.MODELOS_DISPONIVEIS.get(
                    nome_modelo,
                    "gemini-3.5-flash"
                )

                # Etapa 2: Converter PDFs em imagens
                st.info("🖼️ Etapa 2: Convertendo PDFs em imagens...")
                
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
                
                st.success("✅ PDFs convertidos em imagem")

                # Etapa 3: Extrair tabela de roteiro
                st.info("📊 Etapa 3: Lendo a tabela de roteiro perimétrico...")
                
                if imagens_roteiro:
                    segmentos = processador.extrair_roteiro_com_ia(imagens_roteiro)
                else:
                    segmentos = processador.parse_tabela_roteiro_texto(texto_roteiro_manual)

                if not segmentos:
                    st.warning(
                        "⚠️ Nenhum segmento foi extraído da tabela de roteiro. "
                        "Confira se o PDF/texto enviado está correto. "
                        "Tente aumentar o DPI para melhorar a leitura."
                    )
                    st.stop()

                st.success(f"✅ {len(segmentos)} segmentos extraídos")

                # Etapa 4: Mapear confrontantes
                st.info("🤖 Etapa 4: Mapeando confrontantes com IA...")
                
                mapeamento = processador.mapear_confrontantes(
                    imagens_planta=imagens_planta if imagens_planta else None,
                    texto_planta=texto_planta_manual if texto_planta_manual else None,
                    texto_roteiro=texto_roteiro_manual if texto_roteiro_manual else None,
                )
                
                st.success(f"✅ {len(mapeamento.regras)} regras de confrontantes extraídas")

                # Etapa 5: Vincular confrontantes
                st.info("🔗 Etapa 5: Vinculando confrontantes aos segmentos...")
                
                segmentos_vinculados = processador.vincular_confrontantes()
                
                st.success("✅ Confrontantes vinculados")

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
                st.success("🎉 Processamento concluído com sucesso!")
                
                st.subheader("🔍 Resumo de Validação")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(
                        "Proprietário",
                        (dados_finais['proprietario'][:25] + "...")
                        if len(dados_finais['proprietario']) > 25
                        else dados_finais['proprietario']
                    )
                with col2:
                    st.metric("Área Total", dados_finais['area'])
                with col3:
                    st.metric("Perímetro", dados_finais['perimetro'])

                # Tabela detalhada
                with st.expander("📋 Clique para conferir a malha de confrontações vinculadas", expanded=True):
                    df_data = []
                    for seg in dados_finais["segmentos"]:
                        df_data.append({
                            "De": seg['de'],
                            "Para": seg['para'],
                            "N": seg['n_y'],
                            "E": seg['e_x'],
                            "Azimute": seg['azimute'],
                            "Distância": seg['distancia'],
                            "Confrontante": seg['confrontante']
                        })

                    df = pd.DataFrame(df_data)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    st.caption(
                        "⚠️ Confira os valores acima antes de usar o memorial oficialmente. "
                        "A leitura é feita por IA e pode conter erros, especialmente em desenhos "
                        "com baixa qualidade."
                    )

                # Gerar documento Word
                st.info("📝 Gerando documento Word...")
                
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
                
                st.success("✅ Documento gerado com sucesso!")

                # Download
                nome_arquivo = sanitizar_nome_arquivo(cliente_proprietario.upper())
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                st.download_button(
                    label="📥 Baixar Memorial Descritivo (.docx)",
                    data=arquivo_docx,
                    file_name=f"MEMORIAL_DESCRITIVO_{nome_arquivo}_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                # Relatório de processamento
                tempo_fim = time.time()
                relatorio = gerar_relatorio_processamento(
                    dados_finais,
                    tempo_inicio_geral,
                    tempo_fim,
                    processador.tempo_gemini
                )
                
                with st.expander("📊 Relatório de Processamento", expanded=False):
                    st.text(relatorio)
                    
                    st.download_button(
                        label="📥 Baixar Relatório (.txt)",
                        data=relatorio,
                        file_name=f"relatorio_{timestamp}.txt",
                        use_container_width=True
                    )

            except ValueError as e:
                st.error(f"❌ Erro de Validação: {str(e)}")
                logger.error(f"Erro de validação: {str(e)}")

            except json.JSONDecodeError as e:
                st.error(f"❌ Erro ao processar resposta da IA: {str(e)}")
                logger.error(f"Erro JSON: {str(e)}")

            except Exception as e:
                st.error(f"❌ Erro inesperado: {str(e)}")
                logger.error(f"Erro geral: {str(e)}", exc_info=True)

                with st.expander("🔧 Detalhes Técnicos (Debug)"):
                    import traceback
                    st.code(traceback.format_exc())

    else:
        st.info(
            "👆 **Próximos passos:**\n"
            "1. Carregue ambos os PDFs (planta + roteiro), OU\n"
            "2. Preencha o texto manual de ambos na seção expandida abaixo\n"
            "3. Clique em 'Analisar Documentos' para começar o processamento"
        )


if __name__ == "__main__":
    main()
