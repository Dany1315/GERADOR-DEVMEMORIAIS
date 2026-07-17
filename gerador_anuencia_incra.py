import io
import re
import json
import logging
import zipfile
from datetime import datetime
from typing import Dict, Any, List, Tuple
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    import pypdf
except ImportError:
    pypdf = None

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
        limpo = str(coord_str).replace("-", "").replace('""', '"').strip().replace(",", ".")
        match = re.search(r"(\d+)[°ºd\s]+(\d+)['\'\s]+([\d\.]+)", limpo)
        if match:
            return f"{match.group(1)}°{match.group(2)}'{float(match.group(3)):.3f}\""
        return limpo

    def _formatar_azimute(self, az_str: str) -> str:
        if not az_str: return ""
        limpo = str(az_str).replace("-", "").strip().replace(",", ".")
        match = re.search(r"(\d+)[°ºd\s]+(\d+)", limpo)
        if match:
            return f"{int(match.group(1)):02d}°{int(match.group(2)):02d}'"
        return limpo

    def _formatar_numero(self, num_str: Any, casas: int = 2) -> str:
        if num_str is None: return ""
        try:
            val = str(num_str).replace(",", ".").strip()
            match = re.search(r"[\d\.]+", val)
            return f"{float(match.group(0)):.{casas}f}" if match else str(num_str)
        except Exception: return str(num_str)

    def _obter_dados_estruturados_com_ia(self, texto_memorial: str, dados_projeto: Dict[str, Any]) -> Dict[str, Any]:
        if not self.api_key: return self._estrutura_padrao()
        prompt = f"Analise o memorial e extraia: proprietario_origem, cpf_origem e lista de confrontantes com vertices.\nTexto: {texto_memorial}"
        try:
            model = genai.GenerativeModel("gemini-2.0-flash") # Atualizado para versão estável
            response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            dados = json.loads(response.text.strip())
            return dados if "confrontantes" in dados else {"confrontantes": [dados]}
        except Exception: return self._estrutura_padrao()

    def _extrair_texto_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str) -> str:
        texto = ""
        if nome_arquivo.lower().endswith(".pdf") and pypdf:
            reader = pypdf.PdfReader(io.BytesIO(conteudo_arquivo))
            texto = "\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        elif nome_arquivo.lower().endswith(".docx"):
            texto = "\n".join([p.text for p in Document(io.BytesIO(conteudo_arquivo)).paragraphs])
        else:
            texto = conteudo_arquivo.decode("utf-8", errors="ignore")
        return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\xFF]', '', texto)

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

        # Tabela
        tabela = doc.add_table(rows=1, cols=8)
        tabela.style = 'Table Grid'
        tabela.autofit = False
        headers = ["Código", "Longitude", "Latitude", "Altitude (m)", "Código", "Azimute", "Dist. (m)", "Confrontante"]
        for i, h in enumerate(headers):
            tabela.rows[0].cells[i].text = h

        larguras = [Cm(1.65), Cm(2.16), Cm(1.98), Cm(1.75), Cm(1.75), Cm(1.5), Cm(1.5), Cm(13.11)]

        for v in dados_ia.get("vertices", []):
            row = tabela.add_row()
            vals = [v.get("codigo"), self._formatar_coordenada(v.get("longitude")), self._formatar_coordenada(v.get("latitude")), 
                    self._formatar_numero(v.get("altitude")), v.get("vante"), self._formatar_azimute(v.get("azimute")), 
                    self._formatar_numero(v.get("distancia")), v.get("confrontacao_completa")]
            for i in range(8):
                row.cells[i].text = str(vals[i])
                row.cells[i].width = larguras[i]
                self._definir_margens_celulas_zero(row.cells[i])
                for p in row.cells[i].paragraphs:
                    for run in p.runs: run.font.size = Pt(7)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def gerar_documentos_pelo_memorial(self, conteudo_arquivo: bytes, nome_arquivo: str, dados_projeto: Dict[str, Any]) -> List[Tuple[str, io.BytesIO]]:
        texto = self._extrair_texto_memorial(conteudo_arquivo, nome_arquivo)
        dados_ia = self._obter_dados_estruturados_com_ia(texto, dados_projeto)
        
        dados_upd = dados_projeto.copy()
        dados_upd["proprietario"] = dados_ia.get("proprietario_origem", dados_projeto.get("proprietario", "")).upper()
        
        docs = []
        for conf in dados_ia.get("confrontantes", []):
            nome = str(conf.get("confrontante_proprietario", "Confrontante")).strip()
            docs.append((nome, self._montar_documento_confrontante(conf, dados_upd)))
        return docs

    def gerar_zip_anuencias(self, lista_documentos: List[Tuple[str, io.BytesIO]]) -> io.BytesIO:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nome, doc_buffer in lista_documentos:
                nome_seguro = re.sub(r'[\\/*?:"<>|]', "", nome)
                zip_file.writestr(f"ANUENCIA_{nome_seguro}.docx", doc_buffer.getvalue())
        zip_buffer.seek(0)
        return zip_buffer
