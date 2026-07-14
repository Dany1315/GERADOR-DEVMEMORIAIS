"""
GERADOR DE MEMORIAL DESCRITIVO - Versão 6.0
Refatoração completa com arquitetura melhorada, melhor performance e UX
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
from gerador_word import GeradorMemorialWord, GeradorAnuenciaWord

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
    
    st.title("📄 Processador Topográfico - Gleba A")
    
    # Criar abas
    tab_memorial, tab_anuencia = st.tabs(["📄 Gerador de Memorial", "🤝 Gerador de Anuência"])
    
    # ==========================================
    # SIDEBAR COM CONFIGURAÇÕES
    # ==========================================
    
    with st.sidebar:
        st.header("⚙️ Configurações")

        with st.expander("📋 Dados da Empresa", expanded=False):
            empresa_nome = st.text_input("Nome da Empresa", value=EMPRESA_CONFIG.NOME)
            empresa_endereco = st.text_input("Endereço", value=EMPRESA_CONFIG.ENDERECO)
            empresa_telefone = st.text_input("Telefone", value=EMPRESA_CONFIG.TELEFONE)
            empresa_email = st.text_input("Email", value=EMPRESA_CONFIG.EMAIL)

        with st.expander("👤 Dados do Técnico Responsável", expanded=False):
            tecnico_nome = st.text_input("Nome do Técnico", value=TECNICO_CONFIG.NOME)
            tecnico_cfta = st.text_input("CFTA", value=TECNICO_CONFIG.CFTA)

        with st.expander("🤖 Modelo de IA", expanded=True):
            nome_modelo = st.selectbox(
                "Modelo Gemini",
                options=list(GEMINI_CONFIG.MODELOS_DISPONIVEIS.keys()),
                index=0
            )

        with st.expander("👥 Dados do Cliente", expanded=True):
            cliente_imovel = st.text_input("Imóvel", value=CLIENTE_CONFIG.IMOVEL)
            cliente_proprietario = st.text_input("Nome do Proprietário", value=CLIENTE_CONFIG.PROPRIETARIO)
            cliente_local = st.text_input("Local", value=CLIENTE_CONFIG.LOCAL)
            cliente_area = st.text_input("Área (ha)", value=CLIENTE_CONFIG.AREA)
            cliente_perimetro = st.text_input("Perímetro (m)", value=CLIENTE_CONFIG.PERIMETRO)

        with st.expander("🖼️ Processamento", expanded=False):
            dpi_conversao = st.slider("Qualidade (DPI)", 150, 400, 250, 50)
            tamanho_max = st.slider("Tamanho PDF (MB)", 10, 100, 50, 10)

        st.info("💡 **Dica:** As configurações acima são usadas em todos os documentos.")

    with tab_memorial:
        st.subheader("📁 Carregue os Arquivos")
        col1, col2 = st.columns(2)
        with col1:
            pdf_planta = st.file_uploader("PDF PLANTA:", type=["pdf"], key="planta")
        with col2:
            pdf_roteiro = st.file_uploader("PDF ROTEIRO:", type=["pdf"], key="roteiro")

        with st.expander("📝 Alternativa: colar o texto manualmente"):
            col1, col2 = st.columns(2)
            with col1:
                texto_planta_manual = st.text_area("Texto PLANTA:", height=100, key="texto_planta")
            with col2:
                texto_roteiro_manual = st.text_area("Texto ROTEIRO:", height=100, key="texto_roteiro")

        tem_pdfs = pdf_planta and pdf_roteiro
        tem_textos = texto_planta_manual and texto_roteiro_manual
        
        if tem_pdfs or tem_textos:
            if st.button("🔄 Gerar Memorial", type="primary", use_container_width=True):
                tempo_inicio_geral = time.time()
                try:
                    if not configurar_gemini():
                        st.error("❌ Erro na API Gemini")
                        st.stop()
                    
                    nome_modelo_api = GEMINI_CONFIG.MODELOS_DISPONIVEIS.get(nome_modelo, "gemini-3.5-flash")
                    processador = ProcessadorMemorial(nome_modelo_api)
                    
                    # Processamento
                    st.info("⏳ Processando...")
                    imagens_planta = processador.pdf_para_imagens(pdf_planta, dpi=dpi_conversao) if pdf_planta else []
                    imagens_roteiro = processador.pdf_para_imagens(pdf_roteiro, dpi=dpi_conversao) if pdf_roteiro else []
                    
                    segmentos = processador.extrair_roteiro_com_ia(imagens_roteiro) if imagens_roteiro else processador.parse_tabela_roteiro_texto(texto_roteiro_manual)
                    mapeamento = processador.mapear_confrontantes(imagens_planta, texto_planta_manual, texto_roteiro_manual)
                    segmentos_vinculados = processador.vincular_confrontantes()
                    
                    dados_finais = {
                        "imovel": cliente_imovel,
                        "proprietario": cliente_proprietario,
                        "local": cliente_local,
                        "area": cliente_area,
                        "perimetro": cliente_perimetro,
                        "segmentos": segmentos_vinculados
                    }
                    
                    st.session_state['dados_finais'] = dados_finais
                    st.session_state['processamento_concluido'] = True
                    
                    st.success("🎉 Memorial Gerado!")
                    
                    # Word
                    gerador = GeradorMemorialWord({"nome": empresa_nome, "endereco": empresa_endereco, "telefone": empresa_telefone, "email": empresa_email}, {"nome": tecnico_nome, "cfta": tecnico_cfta})
                    arquivo_docx = gerador.gerar_documento(dados_finais)
                    
                    st.download_button("📥 Baixar Memorial (.docx)", data=arquivo_docx, file_name=f"MEMORIAL_{sanitizar_nome_arquivo(cliente_proprietario)}.docx", use_container_width=True)
                    
                    # Tabela
                    st.dataframe(pd.DataFrame(dados_finais["segmentos"]), use_container_width=True)

                except Exception as e:
                    st.error(f"❌ Erro: {str(e)}")
        else:
            st.info("👆 Carregue os arquivos para começar.")

    with tab_anuencia:
        st.header("🤝 Geração de Termos de Anuência")
        if not st.session_state.get('processamento_concluido'):
            st.warning("⚠️ Primeiro, processe o memorial na aba ao lado.")
        else:
            dados_finais = st.session_state['dados_finais']
            segmentos = dados_finais['segmentos']
            confrontantes_unicos = sorted(list(set([s['confrontante'] for s in segmentos if s['confrontante'] not in ["", "CONFRONTAÇÃO NÃO ENCONTRADA"]])))
            
            if not confrontantes_unicos:
                st.error("❌ Nenhum confrontante identificado.")
            else:
                confrontante_selecionado = st.selectbox("Selecione o Confrontante:", options=confrontantes_unicos)
                segmentos_confrontante = [s for s in segmentos if s['confrontante'] == confrontante_selecionado]
                st.write(f"📊 **{len(segmentos_confrontante)}** segmentos encontrados.")
                
                if st.button("📝 Gerar Termo de Anuência", type="primary"):
                    gerador_an = GeradorAnuenciaWord({"nome": empresa_nome, "endereco": empresa_endereco, "telefone": empresa_telefone, "email": empresa_email}, {"nome": tecnico_nome, "cfta": tecnico_cfta})
                    doc_anuencia = gerador_an.gerar_documento({"proprietario": cliente_proprietario, "confrontante": confrontante_selecionado, "local": cliente_local, "segmentos": segmentos_confrontante})
                    st.success("✅ Anuência Gerada!")
                    st.download_button(f"📥 Baixar Anuência - {confrontante_selecionado}", data=doc_anuencia, file_name=f"ANUENCIA_{sanitizar_nome_arquivo(confrontante_selecionado)}.docx", use_container_width=True)

if __name__ == "__main__":
    main()
