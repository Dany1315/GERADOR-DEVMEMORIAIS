# ==========================================
# ARQUIVO: gerador_memorial_word.py
# ==========================================
import io
import logging
from typing import Dict, Any, List
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)


class GeradorMemorialWord:
    """Gerador de Memorial Descritivo em formato Word (.docx)"""
    
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador com dados da empresa e técnico.
        
        Args:
            dados_empresa: Dicionário com 'nome', 'endereco', 'telefone', 'email'
            dados_tecnico: Dicionário com 'nome', 'cfta', 'cpf'
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
    
    def gerar_documento(self, dados_finais: Dict[str, Any]) -> io.BytesIO:
        """
        Gera o memorial descritivo em formato Word.
        
        Args:
            dados_finais: Dicionário com 'imovel', 'proprietario', 'local', 'area', 'perimetro', 'segmentos'
        
        Returns:
            io.BytesIO: Buffer contendo o documento Word
        """
        doc = Document()
        
        # Configurar estilo padrão
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        
        # Extrair dados
        imovel = dados_finais.get("imovel", "IMÓVEL")
        proprietario = dados_finais.get("proprietario", "PROPRIETÁRIO")
        local = dados_finais.get("local", "LOCALIDADE")
        area = dados_finais.get("area", "0,00")
        perimetro = dados_finais.get("perimetro", "0,00")
        segmentos = dados_finais.get("segmentos", [])
        
        # ============================================================
        # CABEÇALHO COM DADOS DA EMPRESA
        # ============================================================
        p_empresa = doc.add_paragraph()
        p_empresa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_emp = p_empresa.add_run(self.dados_empresa.get("nome", "EMPRESA"))
        run_emp.bold = True
        run_emp.font.size = Pt(12)
        
        p_endereco = doc.add_paragraph()
        p_endereco.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_endereco.add_run(self.dados_empresa.get("endereco", ""))
        
        p_contato = doc.add_paragraph()
        p_contato.alignment = WD_ALIGN_PARAGRAPH.CENTER
        telefone = self.dados_empresa.get("telefone", "")
        email = self.dados_empresa.get("email", "")
        p_contato.add_run(f"Tel: {telefone} | Email: {email}")
        
        doc.add_paragraph()  # Espaço
        
        # ============================================================
        # TÍTULO DO MEMORIAL
        # ============================================================
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_titulo = p_titulo.add_run("MEMORIAL DESCRITIVO")
        run_titulo.bold = True
        run_titulo.font.size = Pt(14)
        
        p_subtitulo = doc.add_paragraph()
        p_subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_subtitulo.add_run(f"Imóvel: {imovel}")
        
        doc.add_paragraph()  # Espaço
        
        # ============================================================
        # DADOS DO IMÓVEL
        # ============================================================
        p_dados = doc.add_paragraph()
        p_dados.add_run("DADOS DO IMÓVEL:").bold = True
        
        p_prop = doc.add_paragraph()
        p_prop.add_run(f"Proprietário: ").bold = True
        p_prop.add_run(proprietario)
        
        p_local = doc.add_paragraph()
        p_local.add_run(f"Localidade: ").bold = True
        p_local.add_run(local)
        
        p_area = doc.add_paragraph()
        p_area.add_run(f"Área: ").bold = True
        p_area.add_run(f"{area} hectares")
        
        p_perim = doc.add_paragraph()
        p_perim.add_run(f"Perímetro: ").bold = True
        p_perim.add_run(f"{perimetro} metros")
        
        doc.add_paragraph()  # Espaço
        
        # ============================================================
        # DADOS DO TÉCNICO
        # ============================================================
        p_tecnico_titulo = doc.add_paragraph()
        p_tecnico_titulo.add_run("RESPONSÁVEL TÉCNICO:").bold = True
        
        p_tecnico_nome = doc.add_paragraph()
        p_tecnico_nome.add_run(f"Nome: ").bold = True
        p_tecnico_nome.add_run(self.dados_tecnico.get("nome", ""))
        
        p_tecnico_cfta = doc.add_paragraph()
        p_tecnico_cfta.add_run(f"CFTA: ").bold = True
        p_tecnico_cfta.add_run(self.dados_tecnico.get("cfta", ""))
        
        p_tecnico_cpf = doc.add_paragraph()
        p_tecnico_cpf.add_run(f"CPF: ").bold = True
        p_tecnico_cpf.add_run(self.dados_tecnico.get("cpf", ""))
        
        doc.add_paragraph()  # Espaço
        
        # ============================================================
        # TABELA DE SEGMENTOS/CONFRONTAÇÕES
        # ============================================================
        if segmentos:
            p_tabela_titulo = doc.add_paragraph()
            p_tabela_titulo.add_run("ROTEIRO PERIMÉTRICO:").bold = True
            
            # Criar tabela
            tabela = doc.add_table(rows=1, cols=7)
            tabela.style = 'Table Grid'
            
            # Cabeçalho
            headers = ["De", "Para", "N (Y)", "E (X)", "Azimute", "Distância (m)", "Confrontante"]
            row_headers = tabela.rows[0]
            for idx, header in enumerate(headers):
                cell = row_headers.cells[idx]
                cell.text = header
                # Formatar cabeçalho
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(9)
            
            # Adicionar dados dos segmentos
            for seg in segmentos:
                row = tabela.add_row()
                row.cells[0].text = str(seg.get("de", ""))
                row.cells[1].text = str(seg.get("para", ""))
                row.cells[2].text = str(seg.get("n_y", ""))
                row.cells[3].text = str(seg.get("e_x", ""))
                row.cells[4].text = str(seg.get("azimute", ""))
                row.cells[5].text = str(seg.get("distancia", ""))
                row.cells[6].text = str(seg.get("confrontante", ""))
        
        doc.add_paragraph()  # Espaço
        
        # ============================================================
        # RODAPÉ COM DATA E ASSINATURA
        # ============================================================
        p_data = doc.add_paragraph()
        data_atual = datetime.now()
        meses = [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ]
        texto_data = f"{local}, {data_atual.day} de {meses[data_atual.month - 1]} de {data_atual.year}."
        p_data.add_run(texto_data)
        
        doc.add_paragraph()  # Espaço
        doc.add_paragraph()  # Espaço
        
        # Assinatura do técnico
        p_assinatura = doc.add_paragraph()
        p_assinatura.add_run("_" * 50)
        
        p_nome_assinatura = doc.add_paragraph()
        p_nome_assinatura.add_run(self.dados_tecnico.get("nome", "TÉCNICO"))
        p_nome_assinatura.add_run(f"\nCFTA: {self.dados_tecnico.get('cfta', '')}")
        
        # Salvar em buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        logger.info("Memorial descritivo gerado com sucesso")
        return buffer
