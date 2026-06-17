import io
import re
from datetime import datetime
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from pypdf import PdfReader
import streamlit as st
from pydantic import BaseModel
import os
import json
from google import genai
from google.genai import types

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gerador de Memorial Descritivo - Gleba A", page_icon="📄", layout="wide"
)

client = genai.Client()

# ==========================================
# 1. EXTRATOR E PARSER DO ROTEIRO PERIMÉTRICO (VIA PYTHON - SEM ERROS)
# ==========================================
def extrair_texto_pdf(arquivo_pdf):
    leitor = PdfReader(arquivo_pdf)
    texto_completo = ""
    for pagina in leitor.pages:
        texto_completo += pagina.extract_text() + "\n"
    return texto_completo

def parse_tabela_roteiro(texto_roteiro):
    """
    Extrai via Regex os dados exatos da tabela do PDF, garantindo que o Python
    leia os números reais sem que a IA invente dados.
    """
    # Regex para capturar as linhas da tabela: De, Para, Coord N, Coord E, Azimute, Distancia
    pattern = r'"(\d+)","(\d+)","([\d\.,]+)","([\d\.,]+)","([^"]+)","([\d\.,]+\s*m)"'
    matches = re.findall(pattern, texto_roteiro)
    
    segmentos = []
    for m in matches:
        # Limpa elementos de LaTeX que possam vir no Azimute
        az = m[4].replace('$', '').replace('\\circ', '°').replace('\\prime\\prime', '"').replace('\\prime', "'").strip()
        az = az.replace('\\:', '')
        
        segmentos.append({
            "de": m[0],
            "para": m[1],
            "n_y": m[2] + " m",
            "e_x": m[3] + " m",
            "azimute": az,
            "distancia": m[5].strip(),
            "confrontante": "" # Será preenchido pela IA
        })
    return segmentos

# ==========================================
# 2. MODELO PYDANTIC REDUZIDO (EVITA ALUCINAÇÃO)
# ==========================================
class RegraConfrontante(BaseModel):
    ponto_inicio: int
    ponto_fim: int
    nome_confrontante: str

class MapeamentoConfrontantes(BaseModel):
    proprietario: str
    municipio: str
    comarca: str
    area: str
    perimetro: str
    regras: list[RegraConfrontante]

# ==========================================
# 3. INTEGRAÇÃO INTELIGENTE COM GEMINI
# ==========================================
def mapear_confrontantes_gemini(texto_planta, texto_roteiro):
    """
    O Gemini aqui serve apenas como o cérebro que lê a lógica de quem confronta com quem.
    """
    prompt = f"""
    Você é um engenheiro agrimensor especialista. Analise os documentos abaixo para mapear os confrontantes da Gleba A.

    DOCUMENTO 1 (DADOS DA PLANTA):
    {texto_planta}

    DOCUMENTO 2 (RESUMO DO ROTEIRO):
    {texto_roteiro[:1000]} (Foco no final para área e perímetro)

    Sua tarefa é extrair os dados cadastrais e criar as regras de transição de confrontantes.
    Exemplo: Se diz 'Do ponto 7 - 21: ES 230', crie uma regra com inicio: 7, fim: 21, nome: 'ES 230'.
    Se houver um ponto único como 'Do ponto 1-2: DEVACIR BOONI', inicio: 1, fim: 2.

    Retorne no formato JSON estruturado:
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=MapeamentoConfrontantes,
            temperature=0.1,
        ),
    )
    return json.loads(response.text)

# ==========================================
# 4. GERADOR DO DOCUMENTO DOCX (IGUAL AO MODELO FORMATADO)
# ==========================================
def gerar_documento_word(dados_finais):
    doc = docx.Document()

    # Margens de 2.5 cm
    for section in doc.sections:
        section.top_margin = docx.shared.Cm(2.5)
        section.bottom_margin = docx.shared.Cm(2.5)
        section.left_margin = docx.shared.Cm(2.5)
        section.right_margin = docx.shared.Cm(2.5)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(11)

    # Cabeçalho da Empresa
    p_empresa = doc.add_paragraph()
    p_empresa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_empresa.paragraph_format.space_after = Pt(18)
    run_emp = p_empresa.add_run(
        "TopoGeo Topografia e Consultoria LTDA\nRua Natalino Cossi, No 114, sala 2 - Vila Valério, CEP 29785-000\nFone 27 99837-1164 - topogeo2014@gmail.com"
    )
    run_emp.font.size = Pt(9)
    run_emp.italic = True

    p_linha = doc.add_paragraph()
    p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_linha.add_run("________________________________________________________________________________")

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(12)
    p_titulo.paragraph_format.space_after = Pt(18)
    run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
    run_tit.bold = True
    run_tit.font.size = Pt(12)

    # Cabeçalho de Dados
    p_dados = doc.add_paragraph()
    p_dados.paragraph_format.line_spacing = 1.15
    p_dados.paragraph_format.space_after = Pt(18)

    p_dados.add_run("Imóvel: ").bold = True
    p_dados.add_run("GLEBA A\n")
    p_dados.add_run("Proprietário: ").bold = True
    p_dados.add_run(f"{dados_finais['proprietario'].upper()}\n")
    p_dados.add_run("Município: ").bold = True
    p_dados.add_run(f"{dados_finais['municipio'].upper()}\n")
    p_dados.add_run("Comarca: ").bold = True
    p_dados.add_run(f"{dados_finais['comarca'].upper()}\n")
    p_dados.add_run("Área: ").bold = True
    p_dados.add_run(f"{dados_finais['area']}\n")
    p_dados.add_run("Perímetro: ").bold = True
    p_dados.add_run(f"{dados_finais['perimetro']}")

    p_desc_tit = doc.add_paragraph()
    p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_desc_tit.add_run("DESCRIÇÃO").bold = True

    p_texto = doc.add_paragraph()
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.paragraph_format.line_spacing = 1.25

    segmentos = dados_finais["segmentos"]
    if segmentos:
        primeiro = segmentos[0]
        p_texto.add_run(
            f"Inicia-se a descrição deste perímetro no vértice {primeiro['de']}, de coordenadas N {primeiro['n_y']} e E {primeiro['e_x']}; "
        )

        for s in segmentos:
            p_texto.add_run(
                f"deste, segue confrontando com {s['confrontante']}, com os seguintes azimutes e distâncias: {s['azimute']} e {s['distancia']} até o vértice {s['para']}, de coordenadas N {s['n_y']} e E {s['e_x']}; "
            )

    p_texto.add_run(
        "ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas no Sistema UTM, referenciadas ao Meridiano Central nº 39° WGr, tendo como datum o SIRGAS2000. Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
    )

    # Assinatura técnica
    p_data = doc.add_paragraph()
    p_data.paragraph_format.space_before = Pt(24)
    p_data.add_run(f"Vila Valério, {datetime.now().day} de junho de {datetime.now().year}")

    p_assinatura = doc.add_paragraph()
    p_assinatura.paragraph_format.space_before = Pt(36)
    p_assinatura.add_run(
        "__________________________________________________\nRégis Campo da Silva\nTÉCNICO EM AGROPECUÁRIA\nCFTA: 11198519711\nTRT: BR20260210971"
    )

    conteudo_arquivo = io.BytesIO()
    doc.save(conteudo_arquivo)
    conteudo_arquivo.seek(0)
    return conteudo_arquivo

# ==========================================
# 5. EXECUÇÃO DO FLUXO (STREAMLIT UI)
# ==========================================
st.title("📄 Processador de Memoriais - Gleba A")

col1, col2 = st.columns(2)
with col1:
    pdf_planta = st.file_uploader("Carregue o PDF com os DADOS DA PLANTA:", type=["pdf"], key="planta")
with col2:
    pdf_roteiro = st.file_uploader("Carregue o PDF da TABELA DE ROTEIRO PERIMETRICO:", type=["pdf"], key="roteiro")

if pdf_planta and pdf_roteiro:
    if st.button("Analisar Documentos e Gerar Memorial", type="primary"):
        with st.spinner("⏳ Processando dados com precisão cirúrgica..."):
            try:
                texto_planta = extrair_texto_pdf(pdf_planta)
                texto_roteiro = extrair_texto_pdf(pdf_roteiro)

                # 1. Extrai a tabela real via Python (Erro zero de matemática)
                segmentos_reais = parse_tabela_roteiro(texto_roteiro)

                # 2. IA extrai apenas o mapeamento lógico dos confrontantes e metadados
                mapeamento = mapear_confrontantes_gemini(texto_planta, texto_roteiro)

                # 3. Vincula os confrontantes encontrados aos segmentos reais por faixa de ID
                for seg in segmentos_reais:
                    v_de = int(seg["de"])
                    v_para = int(seg["para"])
                    
                    confrontante_encontrado = "CONFRONTACAO NAO ENCONTRADA"
                    for regra in mapeamento["regras"]:
                        # Lógica de intervalo (Ex: se está entre o ponto 7 e o 21)
                        if regra.ponto_inicio <= v_de < regra.ponto_fim:
                            confrontante_encontrado = regra.nome_confrontante.upper()
                            break
                        # Caso específico do último nó fechando o ciclo
                        elif v_de >= regra.ponto_inicio and v_para == 1 and regra.ponto_fim == 1:
                            confrontante_encontrado = regra.nome_confrontante.upper()
                            break
                    
                    seg["confrontante"] = confrontante_encontrado

                # Monta a estrutura final para o gerador de Word
                dados_finais = {
                    "proprietario": mapeamento["proprietario"],
                    "municipio": mapeamento["municipio"],
                    "comarca": mapeamento["comarca"],
                    "area": "139.954,68 m² (14,00 ha)" if "139.954,68" in texto_roteiro else mapeamento["area"],
                    "perimetro": "1.655,00 m" if "1.655,00" in texto_roteiro else mapeamento["perimetro"],
                    "segmentos": segmentos_reais
                }

                st.write("### 🔍 Visualização Prévio do Cruzamento:")
                st.success(f"**Proprietário:** {dados_finais['proprietario']} | **Área:** {dados_finais['area']}")
                
                # Gera o arquivo final usando os dados reais e estruturados
                arquivo_docx = gerar_documento_word(dados_finais)
                
                st.download_button(
                    label="📥 Baixar Arquivo Word (.docx) Corrigido",
                    data=arquivo_docx,
                    file_name="MEMORIAL_DESCRITIVO_GLEBA_A_SIDNEU_CALLEGARI_CORRIGIDO.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            except Exception as e:
                st.error(f"❌ Ocorreu um erro no processamento: {str(e)}")
