import io
import re
from datetime import datetime
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from pypdf import PdfReader
import streamlit as st
from google import genai
from google.genai import types

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gerador de Memorial Descritivo", page_icon="📄", layout="wide"
)

# ==========================================
# 1. CONFIGURAÇÃO DA API DO GEMINI
# ==========================================
# Chave fornecida inserida diretamente no cliente oficial da Google
GEMINI_API_KEY = "AQ.Ab8RN6LK4VOZSijNDEUarjSOaYyyY4STJ0UVeaSNL-ysxvrPvg"
client = genai.Client(api_key=GEMINI_API_KEY)


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
# 3. INTEGRAÇÃO COM GEMINI (ANÁLISE INTELIGENTE)
# ==========================================
def analisar_dados_com_gemini(texto_documento):
    """Utiliza a API do Gemini para extrair e estruturar os dados técnicos

    e confrontantes do texto, independentemente do formato original.
    """
    prompt = f"""
    Você é um assistente especialista em topografia e georreferenciamento.
    Analise o texto abaixo, que foi extraído de um documento topográfico, e extraia de forma extremamente precisa os dados necessários para gerar um Memorial Descritivo.

    Texto extraído:
    {texto_documento}

    Você deve retornar a resposta estritamente em um formato estruturado (JSON), contendo:
    1. "proprietario": Nome completo do proprietário.
    2. "municipio": Nome do município e estado (ex: VILA VALÉRIO - ES).
    3. "comarca": Nome da comarca (se houver).
    4. "area": Área total com a unidade (ex: 6.002,32 m² (0,60 ha)).
    5. "perimetro": Perímetro total com a unidade (ex: 491,43 m).
    6. "segmentos": Uma lista de objetos para cada trecho do perímetro, contendo:
       - "de": Número do vértice inicial (ex: 1)
       - "para": Número do vértice final (ex: 2)
       - "n_y": Coordenada Norte N(Y) do vértice INICIAL com "m" (ex: "7.901.880,451 m")
       - "e_x": Coordenada Este E(X) do vértice INICIAL com "m" (ex: "351.143,587 m")
       - "azimute": O azimute formatado (ex: 145°20'58")
       - "distancia": A distância formatada com "m" (ex: "2,70 m")
       - "confrontante": O nome do confrontante em letras maiúsculas (ex: "ES 230" ou "ADILSON BRAUN (MAT. 226)")
    """

    # Forçando o Gemini a responder estritamente em JSON usando Pydantic/Structured Outputs
    class SegmentoPerimetro(types.BaseModel):
        de: str
        para: str
        n_y: str
        e_x: str
        azimute: str
        distancia: str
        confrontante: str

    class DadosMemorial(types.BaseModel):
        proprietario: str
        municipio: str
        comarca: str
        area: str
        perimetro: str
        segmentos: list[SegmentoPerimetro]

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DadosMemorial,
            temperature=0.1,  # Baixa temperatura para maior precisão factual
        ),
    )

    # Retorna o objeto JSON estruturado nativamente pelo SDK
    import json

    return json.loads(response.text)


# ==========================================
# 4. GERADOR DO DOCUMENTO DOCX (PADRÃO EXIGIDO)
# ==========================================
def gerar_documento_word(dados):
    """Gera o arquivo Word (.docx) idêntico ao modelo corrigido enviado."""
    doc = docx.Document()

    # Configuração de Margens Padrão (2.5 cm)
    for section in doc.sections:
        section.top_margin = docx.shared.Cm(2.5)
        section.bottom_margin = docx.shared.Cm(2.5)
        section.left_margin = docx.shared.Cm(2.5)
        section.right_margin = docx.shared.Cm(2.5)

    # Configura o padrão de fonte (Arial 11)
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(11)

    # Cabeçalho da Empresa Técnica (TopoGeo)
    p_empresa = doc.add_paragraph()
    p_empresa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_empresa.paragraph_format.space_after = Pt(18)
    run_emp = p_empresa.add_run(
        "TopoGeo Topografia e Consultoria LTDA\nRua Natalino Cossi, No 114, sala 2 - Vila Valério, CEP 29785-000\nFone 27 99837-1164 - topogeo2014@gmail.com"
    )
    run_emp.font.size = Pt(9)
    run_emp.italic = True

    # Linha Divisória
    p_linha = doc.add_paragraph()
    p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_linha.add_run(
        "________________________________________________________________________________"
    )

    # Título Principal
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(12)
    p_titulo.paragraph_format.space_after = Pt(18)
    run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
    run_tit.bold = True
    run_tit.font.size = Pt(12)

    # Bloco de Dados do Imóvel
    p_dados = doc.add_paragraph()
    p_dados.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_dados.paragraph_format.line_spacing = 1.15
    p_dados.paragraph_format.space_after = Pt(18)

    p_dados.add_run("Imóvel: ").bold = True
    p_dados.add_run("GLEBA B\n")
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

    # Subtítulo DESCRIÇÃO
    p_desc_tit = doc.add_paragraph()
    p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_desc_tit.paragraph_format.space_after = Pt(12)
    p_desc_tit.add_run("DESCRIÇÃO").bold = True

    # Texto Técnico Dinâmico Reconstruído baseando-se exatamente no modelo
    p_texto = doc.add_paragraph()
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.paragraph_format.line_spacing = 1.25
    p_texto.paragraph_format.space_after = Pt(12)

    segmentos = dados["segmentos"]
    if segmentos:
        primeiro = segmentos[0]
        # Início no Vértice 1
        p_texto.add_run(
            f"Inicia-se a descrição deste perímetro no vértice {primeiro['de']}, de coordenadas N {primeiro['n_y']} e E {primeiro['e_x']}; "
        )

        # Laço para todos os vértices seguintes
        for s in segmentos:
            p_texto.add_run(
                f"deste, segue confrontando com {s['confrontante']}, com os seguintes azimutes e distâncias: {s['azimute']} e {s['distancia']} até o vértice {s['para']}, de coordenadas N {s['n_y']} e E {s['e_x']}; "
            )

    # Cláusula de encerramento do perímetro padrão do arquivo corrigido
    p_texto.add_run(
        "ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas no Sistema UTM, referenciadas ao Meridiano Central nº 39° WGr, tendo como datum o SIRGAS2000. Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
    )

    # Data (Baseada no dia da execução)
    meses_pt = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro",
    }
    data_atual = datetime.now()
    nome_mes = meses_pt[data_atual.month]

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_data.paragraph_format.space_before = Pt(24)
    p_data.add_run(
        f"Vila Valério, {data_atual.day} de {nome_mes} de {data_atual.year}"
    )

    # Bloco de Assinatura do Técnico Responsável
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
# 5. INTERFACE WEB (STREAMLIT)
# ==========================================
st.title("📄 Processador Inteligente de Memorial Descritivo")
st.write(
    "Faça o upload do seu arquivo PDF ou digite/cole os dados brutos. O Gemini extrairá as coordenadas e confrontantes automaticamente para gerar o documento Word formatado."
)

# Duas opções de entrada de dados
aba_pdf, aba_texto = st.tabs(
    ["📥 Enviar Arquivo PDF", "📝 Colar Texto / Tabela Manual"]
)
texto_para_processar = ""

with aba_pdf:
    arquivo_upload = st.file_uploader(
        "Selecione o PDF topográfico / tabela de vértices:", type=["pdf"]
    )
    if arquivo_upload:
        with st.spinner("Lendo conteúdo do arquivo PDF..."):
            texto_para_processar = extrair_texto_pdf(arquivo_upload)
            st.success("PDF lido com sucesso!")
            with st.expander("Visualizar texto extraído do PDF"):
                st.text(texto_para_processar)

with aba_texto:
    caixa_texto = st.text_area(
        "Cole aqui os dados da tabela ou texto corrido:", height=250
    )
    if caixa_texto:
        texto_para_processar = caixa_texto.strip()

# Botão de processamento unificado
if st.button("Analisar e Gerar Memorial Corrigido", type="primary"):
    if not texto_para_processar:
        st.error(
            "❌ Erro: Por favor, faça o upload de um PDF ou cole o texto antes de prosseguir."
        )
    else:
        with st.spinner(
            "⏳ O Gemini está analisando e estruturando os dados técnicos (Vértices, Coordenadas e Confrontantes)..."
        ):
            try:
                # Envia os dados para a inteligência artificial mapear as variáveis
                dados_estruturados = analisar_dados_com_gemini(
                    texto_para_processar
                )

                # Exibe um resumo dos dados capturados na tela para validação do usuário
                st.write("### 🔍 Dados identificados pelo modelo:")
                st.info(f"**Proprietário:** {dados_estruturados['proprietario']}")
                st.info(f"**Município/Comarca:** {dados_estruturados['municipio']}")
                st.info(
                    f"**Área:** {dados_estruturados['area']} | **Perímetro:** {dados_estruturados['perimetro']}"
                )

                with st.expander("Ver confrontantes detalhados por vértice"):
                    for seg in dados_estruturados["segmentos"]:
                        st.write(
                            f"Vértice {seg['de']} ➔ {seg['para']} | Confrontante: **{seg['confrontante']}** | Az: {seg['azimute']} | Dist: {seg['distancia']}"
                        )

                # Cria o documento baseado nos dados estruturados do Gemini
                arquivo_docx = gerar_documento_word(dados_estruturados)

                st.success("🎉 Memorial gerado perfeitamente!")

                # Define o nome exato solicitado por você
                nome_final_arquivo = "MEMORIAL_DESCRITIVO_GLEBA_B_SIDNEU_CALLEGARI_CORRIGIDO.docx"

                st.download_button(
                    label="📥 Baixar Arquivo Word (.docx) Corrigido",
                    data=arquivo_docx,
                    file_name=nome_final_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            except Exception as e:
                st.error(f"❌ Ocorreu um erro ao processar os dados: {str(e)}")
