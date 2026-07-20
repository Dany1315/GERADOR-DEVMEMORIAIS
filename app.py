# GERADOR DE MEMORIAL DESCRITIVO - Versão 6.2 (UI/UX Premium All-Green Edition)
import io
import logging
import time
from typing import Optional, List, Dict, Any
from datetime import datetime
import json

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


def carregar_imagem_direta(arquivo) -> bytes:
    """Lê o conteúdo bruto de uma imagem (PNG/JPG/JPEG) já como bytes para envio ao Gemini."""
    return arquivo.getvalue()


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

def main():
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
                texto_planta_manual = st.text_area("Texto da PLANTA:", height=150, placeholder="De 1 para 2 confronta com...")
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
                                img_bytes = carregar_imagem_direta(pdf_planta)
                                imagens_planta.append(img_bytes)
                            else:
                                status.update(label="Convertendo páginas da planta em matrizes gráficas...")
                                imagens_planta = processador.pdf_para_imagens(pdf_planta, dpi=dpi_conversao)

                        # ---- Roteiro ----
                        if pdf_roteiro:
                            if is_imagem(pdf_roteiro):
                                status.update(label=f"Carregando imagem do roteiro: {pdf_roteiro.name}...")
                                img_bytes = carregar_imagem_direta(pdf_roteiro)
                                imagens_roteiro.append(img_bytes)
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
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erro no processamento: {str(e)}")
                    logger.error(f"Erro: {str(e)}", exc_info=True)
        else:
            st.info("💡 Carregue os documentos técnicos obrigatórios para liberar a esteira de processamento por inteligência artificial.")

    # ------------------------------------------
    # ABA 2: ANUÊNCIAS
    # ------------------------------------------
    with tab_anuencias:
        st.markdown("### 🤝 Geração Automatizada de Declarações de Anuência")
        dados_memoriais = st.session_state.get("dados_memoriais_processados")

        if not dados_memoriais:
            st.warning("⚠️ Nenhum memorial descritivo processado nesta sessão. Processe os dados na aba anterior.")
        else:
            segmentos = dados_memoriais.get("segmentos", [])
            termos_indigo = ["AV.", "RUA", "AVENIDA", "ESTRADA", "PROJEÇÃO", "VALA", "CORREGO", "VALAO"]
            
            confrontantes_validos = sorted(list(set(
                [str(s.get("confrontante", "")).strip().upper() for s in segmentos
                 if s.get("confrontante") and not any(t in str(s.get("confrontante", "")).strip().upper() for t in termos_indigo)]
            )))

            if not confrontantes_validos:
                st.info("ℹ️ Nenhum proprietário confrontante individual elegível foi mapeado na poligonal deste imóvel.")
            else:
                st.success(f"🔍 Encontrados **{len(confrontantes_validos)}** confrontantes elegíveis para Termos de Anuência.")
                trt_numero = st.text_input("Número da TRT / ART vinculada:", value="", key="trt_anuencias")

                # Importação no escopo correto para evitar redundâncias
                from gerador_anuencias import GeradorAnuenciaWord
                
                dados_empresa_dict = {"nome": empresa_nome, "endereco": empresa_endereco, "telefone": empresa_telefone, "email": empresa_email}
                dados_tecnico_dict = {"nome": technico_nome, "cfta": tecnico_cfta, "trt": trt_numero, "cpf": cpf_tecnico}
                
                gerador_anuencia_modulo = GeradorAnuenciaWord(dados_empresa_dict, dados_tecnico_dict)

                for idx, conf in enumerate(confrontantes_validos):
                    with st.expander(f"👤 Termo Unitário: {conf.title()}", expanded=True):
                        seg_filtrados = [s for s in segmentos if str(s.get("confrontante", "")).strip().upper() == conf]
                        intervalos = ", ".join([f"{s['de']} ao {s['para']}" for s in seg_filtrados])
                        
                        st.write(f"**Vértices de Intersecção:** {intervalos}")
                        
                        dados_anuencia = {
                            "proprietario": dados_memoriais.get("proprietario"),
                            "local": cliente_local,
                            "imovel": dados_memoriais.get("imovel"),
                            "confrontante": conf,
                            "intervalos": intervalos,
                            "segmentos": seg_filtrados
                        }
                        
                        if st.button(f"⚙️ Construir Documento Word — {conf.title()}", key=f"build_{idx}"):
                            with st.spinner("Redigindo descrição técnica jurídica..."):
                                try:
                                    arq_anuencia = gerador_anuencia_modulo.gerar_documento(dados_anuencia)
                                    st.download_button(
                                        label="📥 Fazer Download do Termo (.DOCX)",
                                        data=arq_anuencia,
                                        file_name=f"ANUENCIA_{sanitizar_nome_arquivo(conf)}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"down_{idx}"
                                    )
                                except Exception as err:
                                    st.error(f"Falha na compilação: {err}")

    # ------------------------------------------
    # ABA 3: ANUÊNCIAS INCRA (OTIMIZADA PARA TODOS OS VIZINHOS)
    # ------------------------------------------
    with tab_anuencias_incra:
        st.markdown("### 🌾 Geração Automatizada de Declarações de Anuência INCRA")
        
        # Campo para carregar o arquivo "memorial" necessário para o INCRA
        memorial_incra_file = st.file_uploader(
            "Carregar arquivo Memorial (TXT, PDF ou DOCX correspondente):", 
            type=["txt", "pdf", "docx"], 
            key="memorial_incra"
        )
        
        if not memorial_incra_file:
            st.info("💡 Por favor, carregue o arquivo de Memorial correspondente acima para iniciar a geração das anuências do INCRA.")
        else:
            st.success("📄 Arquivo de Memorial carregado com sucesso!")
            trt_incra_numero = st.text_input("Número da TRT / ART vinculada (INCRA):", value="", key="trt_incra")
            
            # Botão de processamento da Anuência INCRA
            if st.button("⚙️ PROCESSAR E GERAR ANUÊNCIAS DE TODOS OS VIZINHOS", type="primary", use_container_width=True):
                with st.spinner("Analisando o memorial e gerando as anuências de todos os confrontantes..."):
                    try:
                        # Importação dinâmica do novo módulo correspondente
                        from gerador_anuencia_incra import GeradorAnuenciaIncraWord
                        
                        dados_empresa_dict = {
                            "nome": empresa_nome, 
                            "endereco": empresa_endereco, 
                            "telefone": empresa_telefone, 
                            "email": empresa_email
                        }
                        dados_tecnico_dict = {
                            "nome": technico_nome, 
                            "cfta": tecnico_cfta, 
                            "trt": trt_incra_numero, 
                            "cpf": cpf_tecnico
                        }
                        
                        # Instanciação do gerador
                        gerador_incra = GeradorAnuenciaIncraWord(dados_empresa_dict, dados_tecnico_dict)
                        
                        # Processamento do arquivo enviado pelo usuário
                        arquivo_conteudo = memorial_incra_file.read()
                        
                        # GERAÇÃO DA LISTA DE DOCUMENTOS (Um para cada confrontante identificado)
                        documentos_gerados = gerador_incra.gerar_documentos_pelo_memorial(
                            arquivo_conteudo, 
                            memorial_incra_file.name,
                            {
                                "proprietario": cliente_proprietario,
                                "imovel": cliente_imovel,
                                "local": cliente_local,
                                "area": cliente_area,
                                "perimetro": cliente_perimetro
                            }
                        )
                        
                        st.balloons()
                        st.success(f"🎉 Processamento concluído! Foram geradas **{len(documentos_gerados)}** anuências individuais.")
                        
                        # -------------------------------------------------------------
                        # Opção 1: Botão para Baixar Todas as Anuências Juntas em ZIP
                        # -------------------------------------------------------------
                        zip_buffer = gerador_incra.gerar_zip_anuencias(documentos_gerados)
                        st.download_button(
                            label="📥 BAIXAR TODAS AS ANUÊNCIAS EM UM ÚNICO ARQUIVO (.ZIP)",
                            data=zip_buffer,
                            file_name=f"ANUENCIAS_INCRA_LOTE_{sanitizar_nome_arquivo(cliente_proprietario.upper())}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                        
                        st.markdown("---")
                        st.markdown("#### 👤 Downloads Individuais por Vizinho:")
                        
                        # -------------------------------------------------------------
                        # Opção 2: Lista com Botões de Download Individuais
                        # -------------------------------------------------------------
                        for nome_confrontante, doc_buffer in documentos_gerados:
                            col_nome, col_btn = st.columns([3, 1])
                            with col_nome:
                                st.markdown(f"**Confrontante:** {nome_confrontante.upper()}")
                            with col_btn:
                                st.download_button(
                                    label="Baixar Word (.docx)",
                                    data=doc_buffer,
                                    file_name=f"ANUENCIA_INCRA_{sanitizar_nome_arquivo(nome_confrontante.upper())}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"down_incra_{sanitizar_nome_arquivo(nome_confrontante.upper())}",
                                    use_container_width=True
                                )
                                
                    except Exception as err:
                        st.error(f"Erro ao processar as anuências INCRA: {err}")
                        logger.error(f"Erro INCRA: {str(err)}", exc_info=True)

    # ------------------------------------------
    # ABA 4: REQUERIMENTO DE CARTÓRIO
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
                with st.status("Processando documentos via Gemini...", expanded=True) as status:
                    try:
                        from gerador_requerimento_cartorio import GeradorRequerimentoCartorio
                        from processador import ProcessadorMemorial
                        
                        if not configurar_gemini():
                            st.error("Erro crítico: Chave API ausente.")
                            st.stop()
                        
                        status.update(label="Preparando documentos para análise visual...")
                        processador_base = ProcessadorMemorial(nome_modelo_api)
                        todas_imagens = []
                        for doc in documentos_pdf:
                            if is_imagem(doc):
                                # Imagem já está pronta para envio à IA
                                img_bytes = carregar_imagem_direta(doc)
                                todas_imagens.append(img_bytes)
                            else:
                                # PDF: converter em imagens
                                imagens = processador_base.pdf_para_imagens(doc, dpi=200)
                                todas_imagens.extend(imagens)
                        
                        status.update(label=f"Analisando {len(todas_imagens)} páginas/imagens com {nome_modelo}...")
                        gerador_req = GeradorRequerimentoCartorio(nome_modelo_api)
                        dados_extraidos = gerador_req.extrair_dados_documentos(todas_imagens)
                        
                        status.update(label="Preenchendo modelo de requerimento Word...")
                        template_name = "-REQUERIMENTODECARTORIO.docx"
                        arquivo_word = gerador_req.gerar_documento(dados_extraidos, template_name)
                        
                        status.update(label="Requerimento gerado com sucesso!", state="complete")
                        
                        st.balloons()
                        st.success("✅ Dados extraídos e requerimento preenchido!")
                        
                        # Exibe os dados extraídos para conferência
                        with st.expander("🔍 Conferir Dados Extraídos (IA)", expanded=False):
                            st.json(dados_extraidos)
                        
                        st.download_button(
                            label="📥 BAIXAR REQUERIMENTO PREENCHIDO (.DOCX)",
                            data=arquivo_word,
                            file_name=f"REQUERIMENTO_CARTORIO_{sanitizar_nome_arquivo(dados_extraidos['requerente_1']['nome'].upper())}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                    except Exception as err:
                        st.error(f"Erro no processamento do requerimento: {err}")
                        logger.error(f"Erro Requerimento: {str(err)}", exc_info=True)
        else:
            st.info("Aguardando upload de documentos para iniciar a análise.")

if __name__ == "__main__":
    main()
