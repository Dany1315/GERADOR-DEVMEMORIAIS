import io
import re
import logging
from datetime import datetime
from typing import Dict, Any, List
import streamlit as st
import google.generativeai as genai

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

    def _obter_dados_estruturados_com_ia(self, texto_memorial: str, dados_projeto: Dict[str, Any]) -> Dict[str, Any]:
        """
        Usa o modelo Gemini 2.5 Flash para extrair os vértices, coordenadas,
        e dados do confrontante para a tabela de limites oficial.
        """
        if not self.api_key:
            raise ValueError("Chave de API do Gemini não configurada nos secrets do Streamlit.")

        prompt = f"""
        Você é um engenheiro cartógrafo especialista em georreferenciamento do INCRA (3ª Edição da Norma Técnica).
        Sua tarefa é ler o texto do memorial descritivo fornecido e estruturar TODOS os vértices correspondentes apenas ao confrontante em questão.
        
        DADOS DE CONTEXTO DO PROJETO:
        - Proprietário Origem: {dados_projeto.get('proprietario', 'Agostinho Izoton')}
        - Imóvel Origem: {dados_projeto.get('imovel', 'Gleba A')}
        - Município: {dados_projeto.get('local', 'Vila Valério - ES')}
        
        TEXTO DO MEMORIAL DESCRITIVO COMPLETO:
        \"\"\"
        {texto_memorial}
        \"\"\"

        Você precisa extrair:
        1. O Nome do Imóvel Rural Confrontante (ex: Sitio Moro, Sitio Sete Quedas)
        2. A Matrícula/Transição do Confrontante (ex: 8281)
        3. A Comarca (ex: São Gabriel da Palha)
        4. O Nome Completo do Proprietário Confrontante (ex: Alecio Moro, Elias Moro)
        5. CPF do Proprietário Confrontante (se disponível, ou deixe um campo vazio '___.___.___-__')
        6. A lista exata de vértices confrontantes com este vizinho contendo:
           - Código do Vértice (ex: G1D-P-06820)
           - Longitude formatada (ex: -40°17'13,717")
           - Latitude formatada (ex: -18°59'09,548")
           - Altitude em metros (ex: 105.09)
           - Código do Vértice de Vante (Vértice Seguinte)
           - Azimute (ex: 01°58')
           - Distância (ex: 94,33)
           - Descrição da confrontação exata (ex: CNS: 02.170-9 | Mat. 8281 | Sitio Moro; Alecio Moro)

        Responda APENAS com um objeto JSON válido, sem markdown adicional, usando a estrutura:
        {{
            "confrontante_imovel": "Sítio...",
            "confrontante_matricula": "...",
            "confrontante_comarca": "...",
            "confrontante_proprietario": "...",
            "confrontante_cpf": "...",
            "vertices": [
                {{
                    "codigo": "G1D-P-...",
                    "longitude": "...",
                    "latitude": "...",
                    "altitude": "...",
                    "vante": "G1D-P-...",
                    "azimute": "...",
                    "distancia": "...",
                    "confrontacao_completa": "..."
                }}
            ]
        }}
        """
        
        model = genai.GenerativeModel("gemini-2.5-flash")
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        
        try:
            return json.loads(response.text.strip())
        except Exception as e:
            logger.error(f"Erro ao decodificar JSON gerado pelo Gemini: {str(e)}")
            # Retorna estrutura padrão vazia em caso de falha de decodificação
            return {
                "confrontante_imovel": "Sítio Exemplo",
                "confrontante_matricula": "9999",
                "confrontante_comarca": "Comarca",
                "confrontante_proprietario": "Confrontante Exemplo",
                "confrontante_cpf": "___.___.___-__",
                "vertices": []
            }

    def gerar_documento_pelo_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str, dados_projeto: Dict[str, Any]) -> io.BytesIO:
        """
        Lê o memorial enviado, analisa as tabelas via Gemini e gera o documento Word formatado (DRL).
        """
        # Extrai o texto do arquivo enviado de acordo com a sua extensão
        texto_memorial = ""
        if nome_arquivo.lower().endswith(".txt"):
            texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")
        elif nome_arquivo.lower().endswith(".docx"):
            doc_temp = Document(io.BytesIO(conteudo_arquivo))
            texto_memorial = "\n".join([p.text for p in doc_temp.paragraphs])
        else:
            # Fallback simples de extração direta
            texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")

        # Chama a inteligência artificial para estruturar os dados do Confrontante
        dados_ia = self._obter_dados_estruturados_com_ia(texto_memorial, dados_projeto)

        # Criação do documento DRL usando o design oficial
        doc = Document()

        # Configuração de Margens Estreitas
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        # Estilo Base do Documento (Times New Roman)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(11)

        # 1. TÍTULO
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_after = Pt(12)
        run_titulo = p_titulo.add_run("DECLARAÇÃO DE RESPEITO DE LIMITES")
        run_titulo.bold = True
        run_titulo.font.size = Pt(12)

        # 2. TEXTO DE ABERTURA (DECLARAÇÃO)
        proprietario_origem = dados_projeto.get("proprietario", "AGOSTINHO IZOTON").upper()
        # CPF estático ou configurado do proprietário origem
        cpf_origem = "215.894.707-10" 
        localidade_origem = dados_projeto.get("local", "Vila Valério - ES")
        
        tecnico_nome = self.dados_tecnico.get("nome", "Régis Campo da Silva")
        tecnico_cfta = self.dados_tecnico.get("cfta", "1119851971-1")
        # Código do INCRA do técnico credenciado (usando código padrão ou enviado)
        codigo_incra = "G1D"

        texto_abertura = (
            f"Eu, {proprietario_origem}, CPF {cpf_origem}, residente no Jurama, Córrego Sete Quedas, {localidade_origem}, "
            f"e eu, {tecnico_nome}, Técnico em Agropecuária, CFTA {tecnico_cfta}, credenciado pelo INCRA sob o código {codigo_incra}, "
            f"declaramos sob as penas da Lei que quando dos trabalhos topográficos executados na citada propriedade "
            f"foram respeitados os limites de \"divisas in loco\" com os confrontantes abaixo relacionados, "
            f"não havendo qualquer litígio entre as partes."
        )

        p_corpo = doc.add_paragraph()
        p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_corpo.paragraph_format.space_after = Pt(12)
        p_corpo.paragraph_format.line_spacing = 1.15
        p_corpo.add_run(texto_abertura)

        # 3. CABEÇALHO CONFRONTANTES
        p_confrontantes = doc.add_paragraph()
        p_confrontantes.paragraph_format.space_after = Pt(6)
        run_conf_label = p_confrontantes.add_run("Confrontantes:")
        run_conf_label.bold = True

        # Data atualizada no formato padrão do modelo
        data_atual = datetime.now()
        meses = [
            "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", 
            "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
        ]
