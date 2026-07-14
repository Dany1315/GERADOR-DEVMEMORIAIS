#"""
#MÓDULO INDEPENDENTE: GERADOR E ANALISADOR DE ANUÊNCIAS VIA GEMINI API
#Este arquivo roda de forma isolada e não afeta o código do Gerador de Memoriais.
#"""

import io
import re
import logging
from datetime import datetimeimport io
import logging
from datetime import datetime
from typing import Dict, List, Any
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)

class GeradorAnuenciaWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador com os dados padrão da empresa e do responsável técnico.
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
        
        # Configuração da API do Gemini obtida de forma segura a partir dos Secrets do Streamlit
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)
        else:
            logger.warning("Chave 'GEMINI_API_KEY' não encontrada nos st.secrets.")

    def consultar_gemini_para_trecho(self, confrontante: str, proprietario: str, segmentos: List[Dict[str, Any]]) -> str:
        """
        Envia os dados georreferenciados do trecho para a API do Gemini 
        para redigir a descrição técnica textual fluida.
        """
        if not self.api_key:
            return "De comum acordo, as partes reconhecem o limite estabelecido pelos vértices informados."

        try:
            # Transforma os segmentos em linhas legíveis para a IA
            linhas_texto = []
            for s in segmentos:
                linhas_texto.append(f"De {s['de']} para {s['para']} com azimute {s['azimute']} e distância {s['distancia']}m")
            roteiro_trecho = "; ".join(linhas_texto)

            prompt = f"""
            Você é um engenheiro agrimensor especialista em retificação de registro imobiliário e topografia jurídica.
            Redija um parágrafo técnico formal, fluido e descritivo (em português) para ser inserido em uma DECLARAÇÃO DE RECONHECIMENTO DE LIMITES.
            O imóvel principal pertence a {proprietario} e o trecho analisado confronta com {confrontante}.
            Dados das linhas do trecho: {roteiro_trecho}.
            
            Retorne APENAS o parágrafo corrido, sem saudações, sem marcações em negrito e sem introduções textuais.
            """
            
            model = genai.GenerativeModel('gemini-1.5-flash')
            resposta = model.generate_content(prompt)
            return resposta.text.strip()
            
        except Exception as e:
            logger.error(f"Erro ao chamar a API do Gemini para Anuência: {str(e)}")
            return f"O limite perimétrico com {confrontante} acompanha as amarrações técnicas e coordenadas descritas na tabela."

    def gerar_documento(self, dados_anuencia: Dict[str, Any]) -> io.BytesIO:
        """
        Gera o arquivo Word (.docx) baseado estritamente no modelo físico de anuência enviado.
        """
        confrontante = dados_anuencia["confrontante"]
        proprietario = dados_anuencia["proprietario"]
        local = dados_anuencia.get("local", "Vila Valério")
        segmentos = dados_anuencia["segmentos"]

        doc = Document()

        # Configuração de Margens Padrão (2.54 cm)
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # Configuração de Estilo e Fonte (Calibri 11)
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        # 1. TÍTULO
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_titulo = p_titulo.add_run("DECLARAÇÃO DE RECONHECIMENTO DE LIMITES")
        run_titulo.bold = True
        run_titulo.font.size = Pt(14)
        doc.add_paragraph("") 

        # 2. TEXTO DE ABERTURA
        texto_abertura = (
            f"Eu, {confrontante.title()}, proprietário do imóvel confrontante, e eu, "
            f"{proprietario.title()}, proprietário do imóvel urbano, declaramos não "
            f"existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis."
        )
        p_abertura = doc.add_paragraph(texto_abertura)
        p_abertura.paragraph_format.line_spacing = 1.15
        p_abertura.paragraph_format.space_after = Pt(12)

        # 3. DESCRIÇÃO DO TRECHO (Via Inteligência Artificial)
        p_desc_tit = doc.add_paragraph()
        p_desc_tit.add_run("Descrição do trecho de confrontação:").bold = True
        
        texto_ia = self.consultar_gemini_para_trecho(confrontante, proprietario, segmentos)
        p_desc_corpo = doc.add_paragraph(texto_ia)
        p_desc_corpo.paragraph_format.line_spacing = 1.15
        p_desc_corpo.paragraph_format.space_after = Pt(12)

        # 4. TABELA TÉCNICA DO TRECHO (Padrão das anuências enviadas)
        tabela = doc.add_table(rows=1, cols=7)
        tabela.style = 'Table Grid'
        
        hdr_cells = tabela.rows[0].cells
        headers = ["De", "Para", "Azimute", "Distância (m)", "E(X)", "N(Y)", "Altitude"]
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True

        total_distancia = 0.0
        for s in segmentos:
            row_cells = tabela.add_row().cells
            row_cells[0].text = str(s['de'])
            row_cells[1].text = str(s['para'])
            row_cells[2].text = str(s['azimute'])
            
            # Tratamento numérico da distância para soma
            dist_val = float(str(s['distancia']).replace(',', '.'))
            total_distancia += dist_val
            row_cells[3].text = f"{dist_val:.2f}".replace('.', ',')
            
            row_cells[4].text = str(s.get('e_x', '0,00'))
            row_cells[5].text = str(s.get('n_y', '0,00'))
            row_cells[6].text = "0,00"

        # Linha Totalizadora
        row_totais = tabela.add_row().cells
        row_totais[0].text = "Total"
        row_totais[0].paragraphs[0].runs[0].font.bold = True
        row_totais[3].text = f"{total_distancia:.2f}".replace('.', ',')
        row_totais[3].paragraphs[0].runs[0].font.bold = True
        
        doc.add_paragraph("") # Espaço pós tabela

        # 5. PARÁGRAFO DE RESPONSABILIDADE TÉCNICA
        texto_tecnico = (
            f"Declaramos ainda que o profissional {self.dados_tecnico.get('nome')} "
            f"(CPF nº {self.dados_tecnico.get('cpf', '111.985.197-11')}), Resp. Técnico "
            f"(CFTA {self.dados_tecnico.get('cfta')}), credenciado pelo INCRA sob o cod. G1D, com a emissão da TRT nº "
            f"{self.dados_tecnico.get('trt')}, nos indicou as demarcações do limite entre as nossas propriedades, tanto no campo como "
            f"nas suas apresentações gráficas. Concordamos com essa demarcação, expressa na planta e no memorial descritivo, "
            f"ambos em anexo, e reconhecemos esta descrição como o limite legal entre nossas propriedades."
        )
        p_tecnico = doc.add_paragraph(texto_tecnico)
        p_tecnico.paragraph_format.line_spacing = 1.15
        p_tecnico.paragraph_format.space_after = Pt(24)

        # 6. ENCERRAMENTO COM DATA
        data_atual = datetime.now().strftime('%d de %m de %Y')
        p_data = doc.add_paragraph(f"{local}, {data_atual}.")
        p_data.paragraph_format.space_after = Pt(36)

        # 7. CAMPOS DE ASSINATURA DOS PROPRIETÁRIOS
        tab_assinatura = doc.add_table(rows=2, cols=2)
        celulas_l1 = tab_assinatura.rows[0].cells
        celulas_l1[0].text = "__________________________________________________"
        celulas_l1[1].text = "__________________________________________________"
        
        celulas_l2 = tab_assinatura.rows[1].cells
        celulas_l2[0].text = f"{confrontante.title()}\nProprietário do Imóvel Confrontante"
        celulas_l2[1].text = f"{proprietario.title()}\nProprietário do Imóvel"
        
        doc.add_paragraph("\n\n")

        # 8. ASSINATURA DO RESPONSÁVEL TÉCNICO
        p_ass_tec = doc.add_paragraph()
        p_ass_tec.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ass_tec.add_run("______________________________\n").bold = True
        p_ass_tec.add_run(f"{self.dados_tecnico.get('nome')}\nResp. Técnico\nCFTA: {self.dados_tecnico.get('cfta')}")

        # Retorna o arquivo formatado em bytes pronto para download
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
