#"""
#MÓDULO INDEPENDENTE: GERADOR E ANALISADOR DE ANUÊNCIAS VIA GEMINI API
#Este arquivo roda de forma isolada e não afeta o código do Gerador de Memoriais.
#"""

import io
import re
import logging
from datetime import datetime
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai

# Configuração de Logging para segurança
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuração da página isolada
st.set_page_config(
    page_title="Painel de Anuências Independente",
    page_icon="🤝",
    layout="wide"
)

# Inicializa API do Gemini a partir dos Secrets do Streamlit de forma segura
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("🔑 Chave 'GEMINI_API_KEY' não encontrada nos Secrets do Streamlit.")

def chamar_gemini_para_texto(confrontante: str, proprietario: str, trecho_dados: str) -> str:
    """Comunica com a API do Gemini para redigir o parágrafo técnico da anuência"""
    try:
        prompt = f"""
        Você é um engenheiro agrimensor especialista em retificação de registro imobiliário.
        Redija um parágrafo técnico formal e fluido (em português) para uma DECLARAÇÃO DE RECONHECIMENTO DE LIMITES.
        O imóvel é de propriedade de {proprietario} e o trecho analisado confronta com {confrontante}.
        Dados georreferenciados do trecho: {trecho_dados}.
        Retorne APENAS o parágrafo corrido, sem saudações, marcações em negrito ou introduções.
        """
        model = genai.GenerativeModel('gemini-1.5-flash')
        resposta = model.generate_content(prompt)
        return resposta.text.strip()
    except Exception as e:
        logger.error(f"Erro na API Gemini: {str(e)}")
        return f"O perímetro delimitado confronta com {confrontante} acompanhando as amarrações técnicas e coordenadas descritas."

def gerar_docx_anuencia(confrontante: str, proprietario: str, local: str, dados_tecnico: dict, segmentos: list, texto_ia: str) -> bytes:
    """Gera o arquivo Word (.docx) baseado estritamente no modelo físico enviado"""
    doc = Document()
    
    # Configuração de margens (2,5 cm padrão)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # Título
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.add_run("DECLARAÇÃO DE RECONHECIMENTO DE LIMITES").bold = True
    p_tit.runs[0].font.size = Pt(14)
    doc.add_paragraph("")

    # Abertura
    texto_abertura = (
        f"Eu, {confrontante.upper()}, proprietário do imóvel confrontante, e eu, "
        f"{proprietario.upper()}, proprietário do imóvel urbano, declaramos não "
        f"existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis."
    )
    doc.add_paragraph(texto_abertura).paragraph_format.line_spacing = 1.15
    
    # Texto descritivo da IA
    p_desc_tit = doc.add_paragraph()
    p_desc_tit.add_run("Descrição do trecho de confrontação:").bold = True
    doc.add_paragraph(texto_ia).paragraph_format.line_spacing = 1.15
    doc.add_paragraph("")

    # Tabela Técnica
    tabela = doc.add_table(rows=1, cols=7)
    tabela.style = 'Light Shading Accent 1'
    hdr_cells = tabela.rows[0].cells
    headers = ["De", "Para", "Azimute", "Distância (m)", "E(X)", "N(Y)", "Altitude"]
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True

    total_dist = 0.0
    for s in segmentos:
        row = tabela.add_row().cells
        row[0].text = str(s['de'])
        row[1].text = str(s['para'])
        row[2].text = str(s['azimute'])
        
        dist_val = float(str(s['distancia']).replace(',', '.'))
        total_dist += dist_val
        row[3].text = f"{dist_val:.2f}".replace('.', ',')
        row[4].text = str(s.get('e_x', '0,00'))
        row[5].text = str(s.get('n_y', '0,00'))
        row[6].text = "0,00"

    # Linha Totalizer
    row_tot = tabela.add_row().cells
    row_tot[0].text = "Total"
    row_tot[0].paragraphs[0].runs[0].font.bold = True
    row_tot[3].text = f"{total_dist:.2f}".replace('.', ',')
    row_tot[3].paragraphs[0].runs[0].font.bold = True
    doc.add_paragraph("")

    # Termo Técnico
    texto_tec = (
        f"Declaramos ainda que o profissional {dados_tecnico['nome']} (CFTA {dados_tecnico['cfta']}), "
        f"Resp. Técnico, credenciado pelo INCRA sob o cod. G1D, com a emissão da TRT nº {dados_tecnico['trt']}, "
        f"nos indicou as demarcações do limite entre as nossas propriedades, tanto no campo como nas suas apresentações gráficas. "
        f"Concordamos com essa demarcação e reconhecemos esta descrição como o limite legal entre nossas propriedades."
    )
    doc.add_paragraph(texto_tec).paragraph_format.line_spacing = 1.15
    doc.add_paragraph("")

    # Data e Local
    data_atual = datetime.now().strftime('%d de %m de %Y')
    doc.add_paragraph(f"{local} – ES, {data_atual}.")
    doc.add_paragraph("\n")

    # Assinaturas
    tab_ass = doc.add_table(rows=2, cols=2)
    l1 = tab_ass.rows[0].cells
    l1[0].text = "__________________________________________________"
    l1[1].text = "__________________________________________________"
    l2 = tab_ass.rows[1].cells
    l2[0].text = f"{confrontante.title()}\nProprietário do Imóvel Confrontante"
    l2[1].text = f"{proprietario.title()}\nProprietário do Imóvel"
    doc.add_paragraph("\n")

    p_tec_ass = doc.add_paragraph()
    p_tec_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tec_ass.add_run(f"______________________________\n{dados_tecnico['nome']}\nResp. Técnico\nCFTA: {dados_tecnico['cfta']}")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# --- INTERFACE STREAMLIT ISOLADA ---
st.title("🤝 Analisador e Emissor de Anuências (Módulo Independente)")
st.write("Insira as informações do memorial abaixo para segmentar e gerar os termos de anuência via IA.")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📋 Dados Cadastrais")
    proprietario = st.text_input("Nome do Proprietário do Imóvel", "SEBASTIAO IZOTON")
    local = st.text_input("Cidade / Localidade", "Vila Valério")
    
    st.markdown("**Dados do Responsável Técnico**")
    tecnico_nome = st.text_input("Nome do Técnico", "Régis Campo da Silva")
    tecnico_cfta = st.text_input("CFTA", "11198519711")
    tecnico_trt = st.text_input("Número da TRT", "BR20260312846")

with col2:
    st.subheader("📐 Inserção da Malha de Vértices")
    st.info("Insira abaixo as linhas do memorial ou roteiro para que o sistema separe por confrontante automaticamente.")
    
    # Área para simulação/input dos segmentos coletados do memorial
    exemplo_dados = [
        {"de": "1", "para": "2", "azimute": "109°26'28\"", "distancia": "16.86", "e_x": "356.53", "n_y": "7.884.465,45", "confrontante": "ELISEU BISONE"},
        {"de": "4", "para": "1", "azimute": "28°10'34\"", "distancia": "15.68", "e_x": "356.51", "n_y": "7.884.471,06", "confrontante": "EZEQUIEL PEREIRA DOS SANTOS SILVA"}
    ]
    
    if "lista_segmentos" not in st.session_state:
        st.session_state["lista_segmentos"] = exemplo_dados

    df_editavel = st.data_editor(st.session_state["lista_segmentos"], num_rows="dynamic", use_container_width=True)

if st.button("⚡ Analisar Malha de Confrontantes com o Gemini", type="primary"):
    confrontantes = set([str(row['confrontante']).strip().upper() for row in df_editavel if row.get('confrontante')])
    
    for conf in sorted(confrontantes):
        if not conf or conf in ["RUA", "AVENIDA", "AV."]:
            continue
            
        seg_filtrados = [r for r in df_editavel if str(r['confrontante']).strip().upper() == conf]
        
        with st.container(border=True):
            st.write(f"### 👤 Confrontante: {conf}")
            
            # Converte dados do trecho em texto puro para enviar para a inteligência artificial
            resumo_trecho = "; ".join([f"De {s['de']} para {s['para']} com az {s['azimute']} e dist {s['distancia']}m" for s in seg_filtrados])
            
            with st.spinner(f"O Gemini está analisando e descrevendo o trecho de {conf}..."):
                texto_ia = chamar_gemini_para_texto(conf, proprietario, resumo_trecho)
            
            st.write(f"**Texto sugerido pela IA:** *{texto_ia}*")
            
            dados_tecnico = {"nome": tecnico_nome, "cfta": tecnico_cfta, "trt": tecnico_trt}
            arquivo_word = gerar_docx_anuencia(conf, proprietario, local, dados_tecnico, seg_filtrados, texto_ia)
            
            st.download_button(
                label=f"📥 Baixar ANUENCIA_{conf}.docx",
                data=arquivo_word,
                file_name=f"ANUENCIA_{conf.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
