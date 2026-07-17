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
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

class GeradorAnuenciaIncraWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)

    def _estrutura_padrao(self) -> Dict[str, Any]:
        return {
            "proprietario_origem": "RODRIGO COLOMBI FROTA",
            "cpf_origem": "092.653.737-76",
            "confrontantes": []
        }

    def _formatar_coordenada(self, coord_str: str) -> str:
        if not coord_str: return ""
        limpo = coord_str.replace("-", "").replace('""', '"').strip().replace(",", ".")
        match = re.search(r"(\d+)[°ºd\s]+(\d+)['\'\s]+([\d\.]+)", limpo)
        if match:
            return f"{match.group(1)}°{match.group(2)}'{float(match.group(3)):.3f}\""
        return limpo

    def _formatar_azimute(self, az_str: str) -> str:
        if not az_str: return ""
        limpo = az_str.replace("-", "").strip().replace(",", ".")
        match = re.search(r"(\d+)[°ºd\s]+(\d+)", limpo)
        if match:
            return f"{int(match.group(1)):02d}°{int(match.group(2)):02d}'"
        return limpo

    def _formatar_numero(self, num_str: Any, casas: int = 2) -> str:
        if num_str is None: return ""
        try:
            val = str(num_str).replace(",", ".").strip()
            match = re.search(r"[\d\.]+", val)
            if match:
                return f"{float(match.group(0)):.{casas}f}"
            return str(num_str)
        except Exception: return str(num_str)

    def _obter_dados_estruturados_com_ia(self, texto_memorial: str, dados_projeto: Dict[str, Any]) -> Dict[str, Any]:
        estrutura_padrao = self._estrutura_padrao()
        if not self.api_key: return estrutura_padrao

        prompt = f"""
        Você é engenheiro cartógrafo. Analise o memorial e extraia confrontantes e proprietário.
        Texto: {texto_memorial}
        Responda APENAS com JSON estruturado com: proprietario_origem, cpf_origem e confrontantes (lista com vertices).
        """
        try:
            model = genai.GenerativeModel("gemini-2.5-flash")
            response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            dados = json.loads(response.text.strip())
            return dados if "confrontantes" in dados else {"confrontantes": [dados]}
        except Exception:
            return estrutura_padrao

    def _extrair_texto_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str) -> str:
        texto_memorial = ""
        if nome_arquivo.lower().endswith(".pdf") and pypdf:
            reader = pypdf.PdfReader(io.BytesIO(conteudo_arquivo))
            texto_memorial = "\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        elif nome_arquivo.lower().endswith(".docx"):
            doc_temp = Document(io.BytesIO(conteudo_arquivo))
            texto_memorial = "\n".join([p.text for p in doc_temp.paragraphs])
        else:
            texto_memorial = conteudo_arquivo.decode("utf-8", errors="ignore")
        return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\xFF]', '', texto_memorial)

    def gerar_documentos_pelo_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str, dados_projeto: Dict[str, Any]) -> List[Tuple[str, io.BytesIO]]:
        texto_memorial = self._extrair_texto_memorial(conteudo_arquivo, nome_arquivo)
        dados_ia = self._obter_dados_estruturados_com_ia(texto_memorial, dados_projeto)
        
        dados_projeto_atualizados = dados_projeto.copy()
        dados_projeto_atualizados["proprietario"] = dados_ia.get("proprietario_origem", dados_projeto.get("proprietario", "")).upper()
        dados_projeto_atualizados["cpf_proprietario"] = dados_ia.get("cpf_origem", dados_projeto.get("cpf_proprietario", ""))

        documentos = []
        for dados_confrontante in dados_ia.get("confrontantes", []):
            nome = str(dados_confrontante.get("confrontante_proprietario", "Confrontante")).strip()
            documentos.append((nome, self._montar_documento_confrontante(dados_confrontante, dados_projeto_atualizados)))
        return documentos

    def gerar_zip_anuencias(self, lista_documentos: List[Tuple[str, io.BytesIO]]) -> io.BytesIO:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for nome, doc_buffer in lista_documentos:
                nome_seguro = re.sub(r'[\\/*?:"<>|]', "", nome)
                zip_file.writestr(f"{nome_seguro}.docx", doc_buffer.getvalue())
        zip_buffer.seek(0)
        return zip_buffer

    def _definir_margens_celulas_zero(self, cell):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m in ['top', 'bottom', 'left', 'right']:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), '20')
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def _montar_documento_confrontante(self, dados_ia: Dict[str, Any], dados_projeto: Dict[str, Any]) -> io.BytesIO:
        doc = Document()
        for section in doc.sections:
            section.orientation = WD_ORIENTATION.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)

        tabela_vert = doc.add_table(rows=1, cols=8)
        tabela_vert.style = 'Table Grid'
        tabela_vert.autofit = False

        headers = ["Código", "Longitude", "Latitude", "Altitude (m)", "Código", "Azimute", "Dist. (m)", "Confrontante"]
        hdr_cells = tabela_vert.rows[0].cells
        for i, h in enumerate(headers): hdr_cells[i].text = h

        larguras_t2 = [Cm(1.65), Cm(2.16), Cm(1.98), Cm(1.75), Cm(1.75), Cm(1.5), Cm(1.5), Cm(13.11)]

        for v in dados_ia.get("vertices", []):
            row = tabela_vert.add_row()
            cells = row.cells
            cells[0].text = str(v.get("codigo", ""))
            cells[1].text = self._formatar_coordenada(str(v.get("longitude", "")))
            cells[2].text = self._formatar_coordenada(str(v.get("latitude", "")))
            cells[3].text = self._formatar_numero(v.get("altitude", ""))
            cells[4].text = str(v.get("vante", ""))
            cells[5].text = self._formatar_azimute(str(v.get("azimute", "")))
            cells[6].text = self._formatar_numero(v.get("distancia", ""))
            cells[7].text = str(v.get("confrontacao_completa", ""))

        for row in tabela_vert.rows:
            for c_idx, cell in enumerate(row.cells):
                cell.width = larguras_t2[c_idx]
                self._definir_margens_celulas_zero(cell)
                for p in cell.paragraphs:
                    for run in p.runs: run.font.size = Pt(7)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
