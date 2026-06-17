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
from google import genai
from google.genai import types

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gerador de Memorial Descritivo", page_icon="📄", layout="wide"
)

# ==========================================
# 1. CONFIGURAÇÃO DA API DO GEMINI (NOVO SDK)
# ==========================================
import os
from google import genai
from google.genai import types

# Segurança: É melhor NÃO deixar a chave exposta direto no código.
# O os.environ abaixo serve para testar localmente, mas o ideal é preencher nos "Secrets" do GitHub/Streamlit.
if "GEMINI_API_KEY" not in os.environ:
    os.environ["GEMINI_API_KEY"] = "AQ.Ab8RN6IvGAjuzov7gtTG8VkVwtijXx_AofJJKgkJNnZVFCJtIQ"

# Inicializa o cliente padrão para desenvolvedores (sem Vertex AI)
client = genai.Client()
# ==========================================
# 2. EXTRATOR DE TEXTO DE PDF
# ==========================================
def extrair_texto_pdf(arquivo_pdf):
    """Lê o arquivo PDF enviado e extrai todo o texto contido nele."""
    leitor = PdfReader(arquivo_pdf)
    texto_completo = ""
    for pagina in leitor.pages:
        texto_completo += pagina.extract_text() + "\n"
    return texto_completo


# ==========================================
# 3. MODELOS DE ESTRUTURAÇÃO (PYDANTIC)
# ==========================================
class SegmentoPerimetro(BaseModel):
    de: str
    para: str
    n_y: str
    e_x: str
    azimute: str
    distancia: str
    confrontante: str


class DadosMemorial(BaseModel):
    proprietario: str
    municipio: str
    comarca: str
    area: str
    perimetro: str
    segmentos: list[SegmentoPerimetro]


# ==========================================
# 4. INTEGRAÇÃO COM GEMINI (ANÁLISE AVANÇADA)
# ==========================================
def analisar_dados_com_gemini(texto_planta, texto_roteiro):
    """Utiliza a API do Gemini via Vertex AI para correlacionar os dados cadastrais da planta
    com os vértices técnicos e confrontantes da tabela de roteiro perimétrico.
    """
    prompt = f"""
    Você é um assistente especialista em topografia, cartografia e engenharia agrimensura.
    Seu objetivo é cruzar as informações de dois documentos para estruturar um Memorial Descritivo perfeito do imóvel (Gleba A).

    DOCUMENTO 1: DADOS DA PLANTA (Contém a relação de quais confrontantes pertencem a quais intervalos de pontos)
    {texto_planta}

    DOCUMENTO 2: TABELA DE ROTEIRO PERIMÉTRICO (Contém as colunas De, Para, Coord. N, Coord. E, Azimute, Distância, além de Área e Perímetro no final)
    {texto_roteiro}

    REGRAS IMPORTANTES DE TRATAMENTO DE TEXTO:
    - Limpe os azimutes! Remova símbolos de LaTeX como '$', '\\circ', '\\prime', '\\prime\\prime'. Formate exatamente assim como o exemplo: 133°19'54".
    - Associe os confrontantes por trecho. Se o Documento 1 diz que do 'ponto 1-2' é 'DEVACIR BOONI', o segmento DE 1 PARA 2 terá o confrontante 'DEVACIR BOONI'. 
    - Se o Documento 1 diz que do 'ponto 7-21' é 'ES 230', significa que TODOS os segmentos individuais sequenciais entre o 7 e o 21 (7-8, 8-9, 9-10 ... até 20-21) terão como confrontante 'ES 230'.
    - Pegue o valor exato da Área (ex: 139.954,68 m² (14,00 ha)) e do Perímetro (ex: 1.655,00 m) localizados no final do Documento 2.

    Retorne a resposta estritamente no formato JSON estruturado respeitando as chaves abaixo:
    1. "proprietario": Nome do proprietário (Procure no texto da planta ou use "SIDNEU CALLEGARI" se não indicado).
    2. "municipio": Nome do município e estado (ex: VILA VALÉRIO - ES).
    3. "comarca": Nome da comarca (ex: SÃO GABRIEL DA PALHA).
    4. "area": Área total formatada (ex: "139.954,68 m² (14,00 ha)").
    5. "perimetro": Perímetro total formatado (ex: "1.655,00 m").
    6. "segmentos": Lista contendo cada linha da tabela de roteiro:
       - "de": Número do vértice inicial (ex: "1")
       - "para": Número do vértice final (ex: "2")
       - "n_y": Coordenada Norte N(Y) do vértice INICIAL formatada com "m" (ex: "7.902.352,947 m")
       - "e_x": Coordenada Este E(X) do vértice INICIAL formatada com "m" (ex: "351.478,017 m")
       - "azimute": O azimute limpo e formatado (ex: "133°19'54\"")
       - "distancia": A distância formatada com "m" (ex: "256,30 m")
       - "confrontante": O nome do confrontante em letras maiúsculas associado àquele trecho específico.
    """

    # Chamada adaptada para o ambiente Vertex AI com suporte à chave de autorização
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DadosMemorial,
            temperature=0.1,
        ),
    )

    import json
    return json.loads(response.text)

# ==========================================
# 5. GERADOR DO DOCUMENTO DOCX (IGUAL MODELO)
# ==========================================
def gerar_documento_word(dados):
    """Gera o arquivo Word (.docx) idêntico ao modelo corrigido."""
    doc = docx.Document()

    # Margens Padrão (2.5 cm)
    for section in doc.sections:
        section.top_margin = docx.shared.Cm(2.5)
        section.bottom_margin = docx.shared.Cm(2.5)
        section.left_margin = docx.shared.Cm(2.5)
        section.right_margin = docx.shared.Cm(2.5)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(11)

    # Cabeçalho TopoGeo
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
    p_linha.add_run(
        "________________________________________________________________________________"
    )

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(12)
    p_titulo.paragraph_format.space_after = Pt(18)
    run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
    run_tit.bold = True
    run_tit.font.size = Pt(12)

    # Bloco de Dados Cadastrais
    p_dados = doc.add_paragraph()
    p_dados.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_dados.paragraph_format.line_spacing = 1.15
    p_dados.paragraph_format.space_after = Pt(18)

    p_dados.add_run("Imóvel: ").bold = True
    p_dados.add_run("GLEBA A\n")
    p_dados.add_run("Proprietário: ").bold = True
    p_dados.add_run(f"{dados['proprietario'].upper()}\n")
    p_dados.add_run("Município: ").bold = True
    p_dados.add_run(f"{dados['municipio'].upper()}\n")
    if dados.get("comarca"):
        p_dados.add_run("Comarca: ").bold = True
        p_dados.add_run(f"{dados['comarca'].upper()}\n")
    p_dados.add_run("Área: ").bold = True
    p_dados.add_run(f"{dados['area']}\n")
    p_dados.add_run("Perímetro: ").bold = True
    p_dados.add_run(f"{dados['perimetro']}")

    p_desc_tit = doc.add_paragraph()
    p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_desc_tit.paragraph_format.space_after = Pt(12)
    p_desc_tit.add_run("DESCRIÇÃO").bold = True

    p_texto = doc.add_paragraph()
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.paragraph_format.line_spacing = 1.25
    p_texto.paragraph_format.space_after = Pt(12)

    segmentos = dados["segmentos"]
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

    # Data Atual Automatizada
    meses_pt = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
        7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    data_atual = datetime.now()
    nome_mes = meses_pt[data_atual.month]

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_data.paragraph_format.space_before = Pt(24)
    p_data.add_run(
        f"Vila Valério, {data_atual.day} de {nome_mes} de {data_atual.year}"
    )

    p_assinatura = doc.add_paragraph()
    p_assinatura.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_assinatura.paragraph_format.space_before = Pt(36)
    p_assinatura.add_run(
        "__________________________________________________\nRégis Campo da Silva\nTÉCNICO EM AGROPECUÁRIA\nCFTA: 11198519711\nTRT: BR20260210971"
    )

    conteudo_arquivo = io.BytesIO()
    doc.save(conteudo_arquivo)
    conteudo_arquivo.seek(0)
    return conteudo_arquivo


# ==========================================
# 6. INTERFACE WEB (STREAMLIT)
# ==========================================
st.title("📄 Processador de Memoriais por Planta e Roteiro")
st.write(
    "Insira os dois arquivos PDF gerados pelo software de topografia para realizar o cruzamento inteligente de dados."
)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Dados da Planta")
    pdf_planta = st.file_uploader(
        "Carregue o PDF com os DADOS DA PLANTA:", type=["pdf"], key="planta"
    )

with col2:
    st.subheader("2. Roteiro Perimétrico")
    pdf_roteiro = st.file_uploader(
        "Carregue o PDF da TABELA DE ROTEIRO PERIMETRICO:",
        type=["pdf"],
        key="roteiro",
    )

if pdf_planta and pdf_roteiro:
    if st.button("Analisar Documentos e Gerar Memorial", type="primary"):
        with st.spinner("⏳ Extraindo, limpando e cruzando dados dos dois PDFs..."):
            try:
                texto_planta = extrair_texto_pdf(pdf_planta)
                texto_roteiro = extrair_texto_pdf(pdf_roteiro)

                # Processamento inteligente via SDK clássico do Gemini
                dados_estruturados = analisar_dados_com_gemini(
                    texto_planta, texto_roteiro
                )

                st.write("### 🔍 Informações Unificadas com Sucesso:")
                st.info(f"**Proprietário:** {dados_estruturados['proprietario']}")
                st.info(f"**Município/Estado:** {dados_estruturados['municipio']}")
                st.info(
                    f"**Área Extraída:** {dados_estruturados['area']} | **Perímetro:** {dados_estruturados['perimetro']}"
                )

                with st.expander("Verificar amarração lógica de confrontantes"):
                    for seg in dados_estruturados["segmentos"]:
                        st.write(
                            f"Ponto {seg['de']} ➔ {seg['para']} | Vizinho: **{seg['confrontante']}** | Az: {seg['azimute']} | Dist: {seg['distancia']}"
                        )

                # Gera o arquivo final com o layout TopoGeo corrigido
                arquivo_docx = gerar_documento_word(dados_estruturados)

                nome_final_arquivo = "MEMORIAL_DESCRITIVO_GLEBA_B_SIDNEU_CALLEGARI_CORRIGIDO.docx"

                st.success("🎉 Documento estruturado com perfeição!")
                st.download_button(
                    label="📥 Baixar Arquivo Word (.docx) Corrigido",
                    data=arquivo_docx,
                    file_name=nome_final_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            except Exception as e:
                st.error(
                    f"❌ Ocorreu um erro no cruzamento dos dados: {str(e)}"
                )
else:
    st.info(
        "💡 Aguardando o upload de ambos os arquivos (DADOS DA PLANTA + TABELA DE ROTEIRO PERIMETRICO) para liberar o botão de processamento."
    )
