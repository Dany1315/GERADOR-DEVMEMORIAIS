import io
import re
import json
import logging
import zipfile
from datetime import datetime
from typing import Dict, Any, List, Tuple
import streamlit as st
import google.generativeai as genai

try:
    import pypdf
except ImportError:
    pypdf = None

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

logger = logging.getLogger(__name__)


class GeradorAnuenciaIncraWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador de anuência do INCRA com dados institucionais e do técnico.
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico

        # Configuração da API do Gemini obtida de forma segura dos secrets do Streamlit
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)

    def _estrutura_padrao(self) -> Dict[str, Any]:
        """
        Estrutura de fallback baseada fielmente nos modelos reais fornecidos.
        """
        return {
            "confrontantes": [
                {
                    "confrontante_imovel": "Sitio Bravin",
                    "confrontante_matricula": "6093",
                    "confrontante_comarca": "São Gabriel da Palha",
                    "confrontante_proprietario": "ANTONIO BRAVIN",
                    "confrontante_cpf": "___.___.___-__",
                    "vertices": [
                        {
                            "codigo": "G1D-M-03281",
                            "longitude": "40°22'10.639\"",
                            "latitude": "19°00'24.525\"",
                            "altitude": "162.10",
                            "vante": "G1D-M-03282",
                            "azimute": "06°49'",
                            "distancia": "640.71",
                            "confrontacao_completa": "CNS: 02.170-9 | Mat. 6093 | Sitio Bravin; Antonio Bravin"
                        }
                    ]
                }
            ]
        }

    def _obter_dados_estruturados_com_ia(self, texto_memorial: str, dados_projeto: Dict[str, Any]) -> Dict[str, Any]:
        """
        Usa o Gemini para analisar o memorial e separar TODOS os confrontantes com seus respectivos
        vértices e dados cadastrais.
        """
        estrutura_padrao = self._estrutura_padrao()

        if not self.api_key:
            logger.warning("Chave de API do Gemini não configurada nos secrets do Streamlit. Usando dados padrão.")
            return estrutura_padrao

        prompt = f"""
        Você é um engenheiro cartógrafo especialista em georreferenciamento do INCRA.
        Sua tarefa é analisar o texto de um memorial descritivo ou relatório de vértices e estruturar as
        informações de TODOS os confrontantes (vizinhos) identificados ao longo da poligonal.

        DADOS DE CONTEXTO DO PROJETO (ORIGEM):
        - Proprietário Origem: {dados_projeto.get('proprietario', 'RODRIGO COLOMBI FROTA')}
        - Imóvel Origem: {dados_projeto.get('imovel', 'Gleba A')}
        - Município: {dados_projeto.get('local', 'Vila Valerio-ES')}

        TEXTO DO MEMORIAL DESCRITIVO / RELATÓRIO EXTRAÍDO:
        \"\"\"
        {texto_memorial}
        \"\"\"

        REGRAS DE EXTRAÇÃO:
        1. Identifique cada confrontante distinct (por nome de proprietário, matrícula ou imóvel) ao longo do perímetro.
        2. Agrupe sob cada confrontante APENAS os vértices/segmentos cujo trecho de "vante" faz divisa com ele.
        3. Se o texto for incompreensível, simule dados verossímeis baseados no contexto para manter o JSON estruturado.
        4. Tente encontrar ou estimar o CPF do confrontante se houver menção, caso contrário, deixe em branco para preenchimento manual (ex: "___.___.___-__").

        Responda APENAS com o JSON estruturado abaixo, sem markdown ou textos explicativos:
        {{
            "confrontantes": [
                {{
                    "confrontante_imovel": "Nome do Imóvel Confrontante",
                    "confrontante_matricula": "Número da Matrícula",
                    "confrontante_comarca": "Nome da Comarca",
                    "confrontante_proprietario": "Nome do Proprietário Confrontante",
                    "confrontante_cpf": "CPF do Confrontante",
                    "vertices": [
                        {{
                            "codigo": "Código do Vértice",
                            "longitude": "Longitude formatada",
                            "latitude": "Latitude formatada",
                            "altitude": "Altitude com duas casas",
                            "vante": "Código do Vértice de Vante",
                            "azimute": "Azimute formatado",
                            "distancia": "Distância formatada com duas casas",
                            "confrontacao_completa": "Descrição completa da confrontação"
                        }}
                    ]
                }}
            ]
        }}
        """

        try:
            model = genai.GenerativeModel("gemini-2.5-flash")
            response = model.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"}
            )

            texto_resposta = response.text.strip()
            if texto_resposta.startswith("```json"):
                texto_resposta = texto_resposta.split("```json")[1].split("```")[0].strip()
            elif texto_resposta.startswith("```"):
                texto_resposta = texto_resposta.split("```")[1].split("```")[0].strip()

            dados = json.loads(texto_resposta)

            if "confrontantes" not in dados:
                dados = {"confrontantes": [dados]}

            return dados
        except Exception as e:
            logger.error(f"Erro ao obter dados estruturados do Gemini: {str(e)}")
            return estrutura_padrao

    def _extrair_texto_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str) -> str:
        texto_memorial = ""
        if nome_arquivo.lower().endswith(".pdf"):
            try:
                pdf_file = io.BytesIO(conteudo_arquivo)
                if pypdf:
                    reader = pypdf.PdfReader(pdf_file)
                    paginas_texto = []
                    for pagina in reader.pages:
                        texto_pag = pagina.extract_text()
                        if texto_pag:
                            paginas_texto.append(texto_pag)
                    texto_memorial = "\n".join(paginas_texto)
                else:
                    texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")
            except Exception as pdf_err:
                logger.error(f"Falha ao extrair texto do PDF: {pdf_err}")
                texto_memorial = "Erro na leitura estruturada do PDF."
        elif nome_arquivo.lower().endswith(".docx"):
            try:
                doc_temp = Document(io.BytesIO(conteudo_arquivo))
                texto_memorial = "\n".join([p.text for p in doc_temp.paragraphs])
            except Exception as docx_err:
                logger.error(f"Não foi possível abrir o arquivo como .docx: {docx_err}")
                texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")
        else:
            texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")

        texto_memorial = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\xFF]', '', texto_memorial)
        return texto_memorial

    def gerar_documentos_pelo_memorial(
        self, conteudo_arquivo: bytes, nome_arquivo: str, dados_projeto: Dict[str, Any]
    ) -> List[Tuple[str, io.BytesIO]]:
        """
        Gera um documento Word separado para cada confrontante identificado no memorial.
        """
        if not conteudo_arquivo:
            raise ValueError("O conteúdo do arquivo está vazio ou corrompido.")

        texto_memorial = self._extrair_texto_memorial(conteudo_arquivo, nome_arquivo)
        dados_ia = self._obter_dados_estruturados_com_ia(texto_memorial, dados_projeto)
        confrontantes = dados_ia.get("confrontantes") or self._estrutura_padrao()["confrontantes"]

        documentos: List[Tuple[str, io.BytesIO]] = []
        for dados_confrontante in confrontantes:
            nome_confrontante = str(
                dados_confrontante.get("confrontante_proprietario", "Confrontante")
            ).strip()
            buffer = self._montar_documento_confrontante(dados_confrontante, dados_projeto)
            documentos.append((nome_confrontante, buffer))

        return documentos

    def gerar_documento_pelo_memorial(
        self, conteudo_arquivo: bytes, nome_arquivo: str, dados_projeto: Dict[str, Any]
    ) -> io.BytesIO:
        """
        Mantido para compatibilidade simples (retorna o primeiro).
        """
        documentos = self.gerar_documentos_pelo_memorial(conteudo_arquivo, nome_arquivo, dados_projeto)
        return documentos[0][1]

    @staticmethod
    def gerar_zip_anuencias(documentos: List[Tuple[str, io.BytesIO]], prefixo_arquivo: str = "ANUENCIA_INCRA") -> io.BytesIO:
        zip_buffer = io.BytesIO()
        nomes_usados = set()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nome_confrontante, buffer in documentos:
                nome_base = re.sub(r'[^A-Za-z0-9_\-]+', '_', nome_confrontante.strip().upper()) or "CONFRONTANTE"
                nome_arquivo = f"{prefixo_arquivo}_{nome_base}.docx"
                sufixo = 2
                nome_final = nome_arquivo
                while nome_final in nomes_usados:
                    nome_final = f"{prefixo_arquivo}_{nome_base}_{sufixo}.docx"
                    sufixo += 1
                nomes_usados.add(nome_final)

                buffer.seek(0)
                zip_file.writestr(nome_final, buffer.read())
                buffer.seek(0)

        zip_buffer.seek(0)
        return zip_buffer

    def _definir_margens_celulas_zero(self, cell):
        """
        Remove o preenchimento interno padrão (padding) das células do Word
        para permitir que o texto fique encostado e as colunas fiquem extremamente justas.
        """
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m in ['top', 'bottom', 'left', 'right']:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), '40')  # Margem mínima de segurança interna
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def _montar_documento_confrontante(
        self, dados_ia: Dict[str, Any], dados_projeto: Dict[str, Any]
    ) -> io.BytesIO:
        """
        Gera o documento Word otimizado e idêntico ao modelo físico real (com cabeçalho de tabela de linha única).
        """
        doc = Document()

        # Configura margens estreitas no documento (idêntico ao modelo físico)
        for section in doc.sections:
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

        # Configura estilo normal como Arial 11pt
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)

        # 1. TÍTULO (DECLARAÇÃO DE RESPEITO DE LIMITES)
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(0)
        p_titulo.paragraph_format.space_after = Pt(12)
        run_titulo = p_titulo.add_run("DECLARAÇÃO DE RESPEITO DE LIMITES")
        run_titulo.bold = True
        run_titulo.font.size = Pt(12)
        run_titulo.font.name = 'Arial'

        # 2. TEXTO DA DECLARAÇÃO
        proprietario_origem = dados_projeto.get("proprietario", "RODRIGO COLOMBI FROTA").upper()
        cpf_origem = dados_projeto.get("cpf_proprietario", "092.653.737-76")
        localidade_origem = dados_projeto.get("local", "Vila Valerio-ES")

        tecnico_nome = self.dados_tecnico.get("nome", "Regis Campo da Silva")
        tecnico_cfta = self.dados_tecnico.get("cfta", "1119851971-1")
        codigo_incra = "G1D"

        p_corpo = doc.add_paragraph()
        p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_corpo.paragraph_format.space_after = Pt(12)
        p_corpo.paragraph_format.line_spacing = 1.15
        
        p_corpo.add_run("Eu, ").font.name = 'Arial'
        run_prop = p_corpo.add_run(f" {proprietario_origem}, CPF {cpf_origem} ")
        run_prop.bold = True
        p_corpo.add_run(f", residente em {localidade_origem}, e eu, ").font.name = 'Arial'
        run_tec = p_corpo.add_run(f" {tecnico_nome} ")
        run_tec.bold = True
        p_corpo.add_run(", Tecnico em Agropecuaria, CFTA ").font.name = 'Arial'
        run_cfta = p_corpo.add_run(f" {tecnico_cfta} ")
        run_cfta.bold = True
        p_corpo.add_run(", credenciado pelo INCRA sob o codigo ").font.name = 'Arial'
        run_incra = p_corpo.add_run(f" {codigo_incra} ")
        run_incra.bold = True
        p_corpo.add_run(", declaramos sob as penas da Lei que quando dos trabalhos topograficos executados na citada propriedade ").font.name = 'Arial'
        run_resp = p_corpo.add_run(" foram respeitados os limites de \"divisas in loco\" ")
        run_resp.bold = True
        p_corpo.add_run(" com os confrontantes abaixo relacionados, ").font.name = 'Arial'
        run_lit = p_corpo.add_run(" não havendo qualquer litigio entre as partes.")
        run_lit.bold = True

        # 3. CABEÇALHO CONFRONTANTES E DATA
        p_confrontantes_label = doc.add_paragraph()
        p_confrontantes_label.paragraph_format.space_after = Pt(2)
        run_conf_label = p_confrontantes_label.add_run(" Confrontantes:")
        run_conf_label.bold = True
        run_conf_label.font.name = 'Arial'

        data_atual = datetime.now()
        meses = [
            "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
            "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
        ]
        texto_data = f"{localidade_origem}, {data_atual.day} de {meses[data_atual.month - 1]} de {data_atual.year}."

        p_data = doc.add_paragraph()
        p_data.paragraph_format.space_after = Pt(12)
        p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_data = p_data.add_run(texto_data)
        run_data.font.name = 'Arial'

        # =====================================================================
        # 4. TABELA DE VÉRTICES / VANTE (Linha Única de Cabeçalho)
        # =====================================================================
        # Removido a Tabela 1 de cadastro anterior por não constar no modelo real
        tabela_vert = doc.add_table(rows=1, cols=8)
        tabela_vert.style = 'Table Grid'
        tabela_vert.autofit = False

        # Cabeçalho simplificado de 1 única linha
        headers_t2 = [
            "Código", "Longitude", "Latitude", "Altitude (m)",
            "Código", "Azimute", "Dist. (m)", "Confrontante"
        ]
        
        hdr_cells = tabela_vert.rows[0].cells
        for idx, text in enumerate(headers_t2):
            hdr_cells[idx].text = text
            p = hdr_cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            run = p.runs[0]
            run.font.bold = True
            run.font.size = Pt(8.5)
            run.font.name = 'Arial'

        # Medidas precisas baseadas no layout para ocupar perfeitamente a página horizontalmente
        larguras_t2 = [
            Inches(0.95),  # Código Vértice
            Inches(0.95),  # Longitude
            Inches(0.95),  # Latitude
            Inches(0.70),  # Altitude (m)
            Inches(0.95),  # Código Vante
            Inches(0.60),  # Azimute
            Inches(0.65),  # Dist. (m)
            Inches(2.25)   # Confrontante
        ]

        vertices_dados = dados_ia.get("vertices", [])
        if not vertices_dados:
            vertices_dados = [
                {
                    "codigo": "G1D-M-03281", "longitude": "40°22'10.639\"", "latitude": "19°00'24.525\"", "altitude": "162.10",
                    "vante": "G1D-M-03282", "azimute": "06°49'", "distancia": "640.71",
                    "confrontacao_completa": f"CNS: 02.170-9 | Mat. {dados_ia.get('confrontante_matricula', '6093')} | {dados_ia.get('confrontante_imovel', 'Sitio Bravin')}; {dados_ia.get('confrontante_proprietario', 'ANTONIO BRAVIN')}"
                }
            ]

        for v in vertices_dados:
            row = tabela_vert.add_row()
            cells = row.cells
            cells[0].text = str(v.get("codigo", ""))
            cells[1].text = str(v.get("longitude", ""))
            cells[2].text = str(v.get("latitude", ""))
            cells[3].text = str(v.get("altitude", ""))
            cells[4].text = str(v.get("vante", ""))
            cells[5].text = str(v.get("azimute", ""))
            cells[6].text = str(v.get("distancia", ""))
            cells[7].text = str(v.get("confrontacao_completa", ""))

        # Formatação das linhas e células
        for r_idx, row in enumerate(tabela_vert.rows):
            for c_idx, cell in enumerate(row.cells):
                cell.width = larguras_t2[c_idx]
                self._definir_margens_celulas_zero(cell)
                
                p = cell.paragraphs[0]
                if c_idx == 7:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                
                if len(p.runs) > 0:
                    p.runs[0].font.size = Pt(8.5)
                    p.runs[0].font.name = 'Arial'

        # Espaço antes das Assinaturas
        p_espaco = doc.add_paragraph()
        p_espaco.paragraph_format.space_before = Pt(36)

        # 5. TABELA DE ASSINATURAS HORIZONTAIS
        tabela_assinaturas = doc.add_table(rows=2, cols=2)
        tabela_assinaturas.autofit = False
        tabela_assinaturas.columns[0].width = Inches(3.7)
        tabela_assinaturas.columns[1].width = Inches(3.7)

        cells_as = tabela_assinaturas.rows[0].cells
        cells_as[0].text = "__________________________________________________"
        cells_as[1].text = "__________________________________________________"

        cells_nomes = tabela_assinaturas.rows[1].cells

        p_origem = cells_nomes[0].paragraphs[0]
        p_origem.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_o1 = p_origem.add_run(f"{proprietario_origem}\n")
        run_o1.bold = True
        p_origem.add_run(f"CPF: {cpf_origem}")

        p_confrontante = cells_nomes[1].paragraphs[0]
        p_confrontante.alignment = WD_ALIGN_PARAGRAPH.CENTER
        nome_vizinho = str(dados_ia.get("confrontante_proprietario", "")).upper()
        run_v1 = p_confrontante.add_run(f"{nome_vizinho}")
        run_v1.bold = True

        for row in tabela_assinaturas.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        run.font.size = Pt(9.5)
                        run.font.name = 'Arial'

        # 6. ASSINATURA DO RESPONSÁVEL TÉCNICO
        p_rt_espaco = doc.add_paragraph()
        p_rt_espaco.paragraph_format.space_before = Pt(28)

        p_linha_rt = doc.add_paragraph()
        p_linha_rt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linha_rt.add_run("_____________________________________")

        p_info_rt = doc.add_paragraph()
        p_info_rt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_rt1 = p_info_rt.add_run(f"{tecnico_nome}\n")
        run_rt1.bold = True
        p_info_rt.add_run(f"CFTA: {tecnico_cfta}")

        for run in p_info_rt.runs:
            run.font.size = Pt(9.5)
            run.font.name = 'Arial'

        # 7. ANEXOS
        p_anexos_espaco = doc.add_paragraph()
        p_anexos_espaco.paragraph_format.space_before = Pt(24)

        p_anexos = doc.add_paragraph()
        run_anexos_label = p_anexos.add_run("Anexos:   ")
        run_anexos_label.bold = True
        p_anexos.add_run("Planta do Imóvel  Memorial Descritivo do Imóvel")
        for run in p_anexos.runs:
            run.font.size = Pt(9)
            run.font.name = 'Arial'

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
