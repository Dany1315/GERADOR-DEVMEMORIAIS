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
    
    def _formatar_data(self, municipio: str) -> str:
        """Formata a data em português"""
        data_atual = datetime.now()
        meses = [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ]
        return f"{municipio}, {data_atual.day} de {meses[data_atual.month - 1]} de {data_atual.year}."
    
    def gerar_documento(self, dados_finais: Dict[str, Any]) -> bytes:
        """
        Gera o memorial descritivo em formato Word com estrutura descritiva.
        
        Args:
            dados_finais: Dicionário com dados do memorial
        
        Returns:
            bytes: Conteúdo do documento em bytes
        """
        doc = Document()
        
        # Configurar estilo padrão
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        
        # Extrair dados
        proprietario = dados_finais.get("proprietario", "N/A")
        municipio = dados_finais.get("municipio", "N/A")
        comarca = dados_finais.get("comarca", "N/A")
        trt = dados_finais.get("trt", "N/A")
        perimetro = dados_finais.get("perimetro", "0,00")
        area = dados_finais.get("area", "0,00")
        matricula = dados_finais.get("matricula", "N/A")
        segmentos = dados_finais.get("segmentos", [])
        
        # ============================================================
        # TÍTULO DO MEMORIAL
        # ============================================================
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_titulo = p_titulo.add_run("MEMORIAL DESCRITIVO")
        run_titulo.bold = True
        run_titulo.font.size = Pt(14)
        
        # ============================================================
        # DADOS DO IMÓVEL
        # ============================================================
        p_prop = doc.add_paragraph()
        p_prop.add_run("Proprietário: ").bold = False
        p_prop.add_run(proprietario)
        
        p_mun = doc.add_paragraph()
        p_mun.add_run("Município: ").bold = False
        p_mun.add_run(municipio)
        
        p_com = doc.add_paragraph()
        p_com.add_run("Comarca: ").bold = False
        p_com.add_run(comarca)
        
        p_trt = doc.add_paragraph()
        p_trt.add_run("TRT: ").bold = False
        p_trt.add_run(trt)
        
        p_perim = doc.add_paragraph()
        p_perim.add_run("Perímetro: ").bold = False
        p_perim.add_run(f"{perimetro} m")
        
        p_area = doc.add_paragraph()
        p_area.add_run("Área: ").bold = False
        p_area.add_run(f"{area} m²")
        
        if matricula and matricula != "N/A":
            p_mat = doc.add_paragraph()
            p_mat.add_run("MAT. ").bold = False
            p_mat.add_run(matricula)
        
        # ============================================================
        # SEÇÃO DE DESCRIÇÃO
        # ============================================================
        p_desc_titulo = doc.add_paragraph()
        run_desc = p_desc_titulo.add_run("DESCRIÇÃO")
        run_desc.bold = True
        
        # Gerar parágrafos descritivos para cada segmento
        if segmentos:
            # Primeiro segmento
            primeiro = segmentos[0]
            p_inicio = doc.add_paragraph()
            p_inicio.add_run(
                f"Inicia-se a descrição deste perímetro no vértice {primeiro.get('de', '1')}, "
                f"de coordenadas N(Y) {primeiro.get('n_y', 'N/A')} e E(X) {primeiro.get('e_x', 'N/A')}, "
                f"situado na divisa com {primeiro.get('confrontante', 'N/A')}"
            )
            
            # Segmentos intermediários
            for seg in segmentos[1:]:
                p_seg = doc.add_paragraph()
                p_seg.add_run(
                    f"Segue-se pela divisa com {seg.get('confrontante', 'N/A')} "
                    f"até o vértice {seg.get('para', 'N/A')}, "
                    f"de coordenadas N(Y) {seg.get('n_y', 'N/A')} e E(X) {seg.get('e_x', 'N/A')}, "
                    f"com azimute {seg.get('azimute', 'N/A')} e distância de {seg.get('distancia', 'N/A')} m"
                )
            
            # Parágrafo de fechamento
            p_fechamento = doc.add_paragraph()
            p_fechamento.add_run(
                f"Fechando o perímetro no vértice 1 de origem, totalizando uma área de {area} m²."
            )
        
        # ============================================================
        # DATA E LOCAL
        # ============================================================
        p_data = doc.add_paragraph()
        p_data.add_run(self._formatar_data(municipio))
        
        # ============================================================
        # ASSINATURA DO TÉCNICO
        # ============================================================
        doc.add_paragraph()  # Espaço
        
        p_assinatura = doc.add_paragraph()
        p_assinatura.add_run("_" * 50)
        
        p_nome = doc.add_paragraph()
        p_nome.add_run(self.dados_tecnico.get("nome", "TÉCNICO"))
        
        p_profissao = doc.add_paragraph()
        p_profissao.add_run("Técnico em Agropecuária")
        
        p_dados_tecnico = doc.add_paragraph()
        p_dados_tecnico.add_run(
            f"CFTA: {self.dados_tecnico.get('cfta', '')} | "
            f"TRT: {trt}"
        )
        
        # Salvar em buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        logger.info("Memorial descritivo gerado com sucesso")
        return buffer.getvalue()
