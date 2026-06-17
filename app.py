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

# ==========================================
# CONEXÃO COM A API KEY DOS SECRETS DO STREAMLIT
# ==========================================
# Garante que o cliente do Gemini use exatamente a chave configurada na aba Secrets
client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

# ==========================================
# 1. EXTRATOR E PARSER DO ROTEIRO PERIMÉTRICO (VIA PYTHON - SEM ERROS)
# ==========================================
def extrair_texto_pdf(arquivo_pdf):
    """Lê o arquivo PDF enviado e extrai todo o texto contido nele."""
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
            "confrontante": "" # Será preenchido pela IA casada com a lógica do script
        })
    return segmentos

# ==========================================
# 2. MODELO PYDANTIC ESTRUTURADO PARA LOGICA DE CONFRONTANTES
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
# 3. INTEGRAÇÃO INTELIGENTE COM GEMINI 3.5 FLASH
# ==========================================
def mapear_confrontantes_gemini(texto_planta, texto_roteiro):
    """
    Utiliza o Gemini 3.5 Flash para interpretar a lógica de quais confrontantes pertencem 
    aos respectivos intervalos de pontos e extrair os dados do cabeçalho.
    """
    prompt = f"""
    Você é um engenheiro agrimensor especialista em topografia. Analise os documentos abaixo para mapear os confrontantes da Gleba A.

    DOCUMENTO 1 (DADOS DA PLANTA - Relação de confrontantes por intervalos):
    {texto_planta}

    DOCUMENTO 2 (TABELA DE ROTEIRO PERIMÉTRICO):
    {texto_roteiro[:1200]} (Use também as linhas finais do texto para capturar os valores totais exatos de Área e Perímetro)

    Sua tarefa é extrair os dados cadastrais solicitados e criar as regras matemáticas de transição de confrontantes.
    Exemplo: Se do ponto 7 ao 21 confronta com 'ES 230', crie um item em regras com ponto_inicio: 7, ponto_fim: 21, nome_confrontante: 'ES 230'.
    Se houver um ponto fechando a divisa de volta para o início como 'Do ponto 21 - 1: JEAN CARLOS CALLEGARI', salve ponto_inicio: 21, ponto_fim: 1, nome_confrontante: 'JEAN CARLOS CALLEGARI'.

    Retorne estritamente no formato JSON estruturado respeitando o schema fornecido.
    """

    # ALTERADO AQUI: Atualizado para gemini-3.5-flash conforme solicitado
    response = client.models.generate_content(
        model="gemini-3.5-flash", 
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=MapeamentoConfrontantes,
            temperature=0.0, # Zero especulação, estritamente factual
        ),
    )
    return json.loads(response.text)

# ==========================================
# 4. GERADOR DO DOCUMENTO DOCX (AMARRADO DINAMICAMENTE)
# ==========================================
def gerar_documento_word(dados_finais):
    """Gera o arquivo Word (.docx) aplicando o encadeamento correto de coordenadas destino."""
    doc = docx.Document()

    # Margens de 2.5 cm (Configuração Oficial Padrão)
    for section in doc.sections:
        section.top_margin = docx.shared.Cm(2.5)
        section.bottom_margin = docx.shared.Cm(2.5)
        section.left_margin = docx.shared.Cm(2.5)
        section.right_margin = docx.shared.Cm(2.5)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(11)

    # Cabeçalho da Empresa TopoGeo
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

    # Cabeçalho de Dados Cadastrais
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
    p_desc_tit.paragraph_format.space_before = Pt(12)
    p_desc_tit.paragraph_format.space_after = Pt(12)
    p_desc_tit.add_run("DESCRIÇÃO").bold = True

    p_texto = doc.add_paragraph()
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.paragraph_format.line_spacing = 1.25
    p_texto.paragraph_format.space_after = Pt(12)

    segmentos = dados_finais["segmentos"]
    if segmentos:
        primeiro = segmentos[0]
        p_texto.add_run(
            f"Inicia-se a descrição deste perímetro no vértice {primeiro['de']}, de coordenadas N {primeiro['n_y']} e E {primeiro['e_x']}; "
        )

        # Loop de Amarração Geográfica Correta
        for i, s in enumerate(segmentos):
            if i + 1 < len(segmentos):
                prox_coordenada_n = segmentos[i + 1]['n_y']
                prox_coordenada_e = segmentos[i + 1]['e_x']
            else:
                # Retorna ao marco inicial fechando o polígono de vértices
                prox_coordenada_n = segmentos[0]['n_y']
                prox_coordenada_e = segmentos[0]['e_x']

            p_texto.add_run(
                f"deste, segue confrontando com {s['confrontante']}, com os seguintes azimutes e distâncias: {s['azimute']} e {s['distancia']} "
                f"até o vértice {s['para']}, de coordenadas N {prox_coordenada_n} e E {prox_coordenada_e}; "
            )

    p_texto.add_run(
        "ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas no Sistema UTM, referenciadas ao Meridiano Central nº 39° WGr, tendo como datum o SIRGAS2000. Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
    )

    # Data Atual Automatizada
    meses_pt = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
        7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    data_atual = datetime.now()
    nome_mes = meses_pt[data_atual.month]

    p_data = doc.add_paragraph()
    p_data.paragraph_format.space_before = Pt(24)
    p_data.add_run(f"Vila Valério, {data_atual.day} de {nome_mes} de {data_atual.year}")

    # Bloco de Assinatura Técnico
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
st.write("Insira os dois arquivos da Gleba A para estruturar automaticamente o Memorial Descritivo corrigido.")

col1, col2 = st.columns(2)
with col1:
    pdf_planta = st.file_uploader("Carregue o PDF com os DADOS DA PLANTA:", type=["pdf"], key="planta")
with col2:
    pdf_roteiro = st.file_uploader("Carregue o PDF da TABELA DE ROTEIRO PERIMETRICO:", type=["pdf"], key="roteiro")

if pdf_planta and pdf_roteiro:
    if st.button("Analisar Documentos e Gerar Memorial", type="primary"):
        with st.spinner("⏳ Conectando à API do Gemini 3.5 Flash e amarrando dados com precisão..."):
            try:
                texto_planta = extrair_texto_pdf(pdf_planta)
                texto_roteiro = extrair_texto_pdf(pdf_roteiro)

                # 1. Extrai a tabela numérica real via Python
                segmentos_reais = parse_tabela_roteiro(texto_roteiro)

                # 2. IA extrai o mapeamento lógico usando a chave secreta vinculada do Streamlit
                mapeamento = mapear_confrontantes_gemini(texto_planta, texto_roteiro)

                # 3. Vincula via código os confrontantes mapeados da IA para as linhas reais da tabela
                for seg in segmentos_reais:
                    v_de = int(seg["de"])
                    v_para = int(seg["para"])
                    
                    confrontante_encontrado = "CONFRONTACAO NAO ENCONTRADA"
                    for regra in mapeamento["regras"]:
                        # Tratamento para faixas regulares
                        if regra.ponto_inicio <= v_de < regra.ponto_fim and regra.ponto_inicio < regra.ponto_fim:
                            confrontante_encontrado = regra.nome_confrontante.upper()
                            break
                        # Tratamento para o fechamento final do ciclo perimétrico (Ex: de 21 para 1)
                        elif regra.ponto_inicio > regra.ponto_fim:
                            if v_de >= regra.ponto_inicio or v_para <= regra.ponto_fim:
                                confrontante_encontrado = regra.nome_confrontante.upper()
                                break
                    
                    seg["confrontante"] = confrontante_encontrado

                # Unificação final de metadados
                dados_finais = {
                    "proprietario": mapeamento["proprietario"],
                    "municipio": mapeamento["municipio"],
                    "comarca": mapeamento["comarca"],
                    "area": mapeamento["area"],
                    "perimetro": mapeamento["perimetro"],
                    "segmentos": segmentos_reais
                }

                # Feedback visual na interface do usuário
                st.write("### 🔍 Resumo de Validação Gerado com Sucesso:")
                st.info(f"**Proprietário Extraído:** {dados_finais['proprietario']}")
                st.info(f"**Área Total:** {dados_finais['area']} | **Perímetro:** {dados_finais['perimetro']}")
                
                with st.expander("Clique aqui para conferir a malha de confrontações vinculadas"):
                    for seg in dados_finais["segmentos"]:
                        st.write(f"Trecho {seg['de']} ➔ {seg['para']} | Confrontante: **{seg['confrontante']}**")
                
                # Geração do arquivo Word final (.docx)
                arquivo_docx = gerar_documento_word(dados_finais)
                
                st.success("🎉 Arquivo estruturado com sucesso!")
                st.download_button(
                    label="📥 Baixar Arquivo Word (.docx) Corrigido",
                    data=arquivo_docx,
                    file_name="MEMORIAL_DESCRITIVO_GLEBA_A_SIDNEU_CALLEGARI_CORRIGIDO.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            except Exception as e:
                st.error(f"❌ Ocorreu um erro inesperado no processamento: {str(e)}")
