#"""
#MÓDULO INDEPENDENTE: GERADOR E ANALISADOR DE ANUÊNCIAS VIA GEMINI API
#Este arquivo roda de forma isolada e não afeta o código do Gerador de Memoriais.
#"""

import io
import re
import logging
from datetime import datetime
from typing import Dict, List, Any
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

logger = logging.getLogger(__name__)

class GeradorAnuenciaWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador com os dados padrão da empresa e do responsável técnico.
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
        
        # Configuração da API do Gemini (Mantido por compatibilidade de estrutura)
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)

    def _limpar_e_converter_numerico(self, valor: Any) -> float:
        """
        Limpa caracteres indesejados mantendo apenas números, pontos e vírgulas,
        garantindo a conversão para float sem erros de compilação.
        """
        try:
            string_limpa = str(valor).strip()
            string_limpa = re.sub(r"[^\d,.]", "", string_limpa)
            if not string_limpa:
                return 0.0
            
            # Se possui os dois separadores (ex: 7,884,471.06), remove as vírgulas de milhar
            if "," in string_limpa and "." in string_limpa:
                string_limpa = string_limpa.replace(",", "")
            # Caso tenha vindo apenas com vírgula como decimal (ex: 15,68)
            elif "," in string_limpa and "." not in string_limpa:
                string_limpa = string_limpa.replace(",", ".")
                
            return float(string_limpa)
        except Exception:
            return 0.0

    def _formatar_para_padrao_modelo(self, valor_numerico: float) -> str:
        """
        Formata o número usando ponto para decimais e vírgula para milhar (Ex: 7,884,471.06).
        """
        # Formata com vírgulas como separador de milhar americano (1,234,567.89)
        return f"{valor_numerico:,.2f}"

    def gerar_documento(self, dados_anuencia: Dict[str, Any]) -> io.BytesIO:
        """
        Gera o arquivo Word (.docx) baseado estritamente no modelo físico de anuência enviado.
        """
        confrontante = dados_anuencia["confrontante"]
        proprietario = dados_anuencia["proprietario"]
        local = dados_anuencia.get("local", "Vila Valério")
        segmentos = dados_anuencia["segmentos"]

        doc = Document()

        # Configuração de Margens Padrão (2.54 cm / 1 polegada)
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # Configuração de Estilo e Fonte (Calibri 11)
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        # Configuração das Cores do Cabeçalho
        cor_verde_claro = RGBColor(34, 139, 34) # Verde Floresta Claro/Médio bem profissional
        cor_preta = RGBColor(0, 0, 0) # Preto puro para as informações de contato

        # =========================================================================
        # CABEÇALHO CORPORATIVO (Topo da Folha, Centralizado)
        # =========================================================================
        p_head1 = doc.add_paragraph()
        p_head1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head1.paragraph_format.space_before = Pt(0)
        p_head1.paragraph_format.space_after = Pt(1)
        run_h1 = p_head1.add_run("TopoGeo")
        run_h1.bold = True
        run_h1.font.size = Pt(12)
        run_h1.font.color.rgb = cor_verde_claro
        
        p_head2 = doc.add_paragraph()
        p_head2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head2.paragraph_format.space_before = Pt(0)
        p_head2.paragraph_format.space_after = Pt(1)
        run_h2 = p_head2.add_run("Topografia   Consultoria LTDA")
        run_h2.bold = True
        run_h2.font.size = Pt(10)
        run_h2.font.color.rgb = cor_verde_claro
        
        p_head3 = doc.add_paragraph()
        p_head3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head3.paragraph_format.space_before = Pt(0)
        p_head3.paragraph_format.space_after = Pt(1)
        run_h3 = p_head3.add_run("Rua Natalino Cossi, No 114, sala 2 - Vila Valério, CEP 29785-000 Fone 27 99837-1164")
        run_h3.font.size = Pt(9)
        run_h3.font.color.rgb = cor_preta
        
        p_head4 = doc.add_paragraph()
        p_head4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head4.paragraph_format.space_before = Pt(0)
        p_head4.paragraph_format.space_after = Pt(24)
        run_h4 = p_head4.add_run("topogeo2014@gmail.com")
        run_h4.font.size = Pt(9)
        run_h4.font.color.rgb = cor_preta

        # 1. TÍTULO: Negrito, Centralizado e em Letras Maiúsculas
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(0)
        p_titulo.paragraph_format.space_after = Pt(18)
        run_titulo = p_titulo.add_run("DECLARAÇÃO DE RECONHECIMENTO DE LIMITES")
        run_titulo.bold = True
        run_titulo.font.size = Pt(11)

        # 2. TEXTO DE ABERTURA: Alinhamento Justificado
        texto_abertura = (
            f"Eu, {confrontante.title()}, proprietário do imóvel confrontante, e eu, "
            f"{proprietario.title()}, proprietário do imóvel urbano, declaramos não "
            f"existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis."
        )
        p_abertura = doc.add_paragraph(texto_abertura)
        p_abertura.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_abertura.paragraph_format.line_spacing = 1.15
        p_abertura.paragraph_format.space_after = Pt(12)

        # 3. DESCRIÇÃO DO TRECHO (Sem parágrafo extra)
        p_desc_tit = doc.add_paragraph()
        p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_desc_tit.paragraph_format.space_after = Pt(6)
        p_desc_tit.add_run("Descrição do trecho de confrontação:").bold = True

        # 4. TABELA TÉCNICA
        tabela = doc.add_table(rows=1, cols=7)
        tabela.style = 'Table Grid'
        tabela.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        hdr_cells = tabela.rows[0].cells
        headers = ["De", "Para", "Azimute", "Distância (m)", "E(X)", "N(Y)", "Altitude"]
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.bold = True
            run.font.size = Pt(10)

        total_distancia = 0.0
        for s in segmentos:
            row_cells = tabela.add_row().cells
            row_cells[0].text = str(s['de'])
            row_cells[1].text = str(s['para'])
            row_cells[2].text = str(s['azimute'])
            
            dist_val = self._limpar_e_converter_numerico(s['distancia'])
            total_distancia += dist_val
            row_cells[3].text = self._formatar_para_padrao_modelo(dist_val)
            
            val_ex = self._limpar_e_converter_numerico(s.get('e_x', '0.00'))
            val_ny = self._limpar_e_converter_numerico(s.get('n_y', '0.00'))
            
            row_cells[4].text = self._formatar_para_padrao_modelo(val_ex)
            row_cells[5].text = self._formatar_para_padrao_modelo(val_ny)
            row_cells[6].text = "0,00"
            
            for cell in row_cells:
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.size = Pt(9.5)

        # Linha de Totais da Tabela
        row_totais = tabela.add_row().cells
        row_totais[0].text = "Total"
        row_totais[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_totais[0].paragraphs[0].runs[0].font.bold = True
        row_totais[0].paragraphs[0].runs[0].font.size = Pt(9.5)

        row_totais[3].text = self._formatar_para_padrao_modelo(total_distancia)
        row_totais[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_totais[3].paragraphs[0].runs[0].font.bold = True
        row_totais[3].paragraphs[0].runs[0].font.size = Pt(9.5)
        
        for idx in [1, 2, 4, 5, 6]:
            row_totais[idx].text = ""
            row_totais[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        p_espaco = doc.add_paragraph("")
        p_espaco.paragraph_format.space_before = Pt(6)

        # 5. PARÁGRAFO DE RESPONSABILIDADE TÉCNICA
        rg_rt = self.dados_tecnico.get('rg', '1.936.653')
        codigo_incra = self.dados_tecnico.get('codigo_incra', 'G1D')
        
        texto_tecnico = (
            f"Declaramos ainda que o profissional {self.dados_tecnico.get('nome')} "
            f"(RG nº {rg_rt} e CPF nº {self.dados_tecnico.get('cpf', '111.985.197-11')}), Resp. "
            f"Técnico (CFTA {self.dados_tecnico.get('cfta')}), credenciado pelo INCRA sob o cod. {codigo_incra}, com a emissão da
