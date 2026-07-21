# ARQUIVO: gerador_memorial_word.py
# ==========================================
import io
import logging
from typing import Dict, Any, List
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches, RGBColor
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
    
    def _formatar_data(self, municipio: str) -> str:
        """Formata a data em português"""
        data_atual = datetime.now()
        meses = [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ]
        return f"{municipio}, {data_atual.day} de {meses[data_atual.month - 1]} de {data_atual.year}."
    
    def _adicionar_cabecalho(self, doc: Document):
        """
        Adiciona o cabeçalho da empresa no header da folha.
        
        ✅ NOVO: Coloca dados da empresa no campo de cabeçalho do Word
        """
        section = doc.sections[0]
        header = section.header
        
        # Limpar header anterior
        for paragraph in header.paragraphs:
            p = paragraph._element
            p.getparent().remove(p)
        
        # Adicionar nome da empresa (em verde)
        p_nome = header.add_paragraph()
        p_nome.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_nome = p_nome.add_run(self.dados_empresa.get("nome", ""))
        run_nome.font.size = Pt(12)
        run_nome.font.bold = True
        run_nome.font.color.rgb = RGBColor(6, 78, 59)  # Verde escuro
        
        # Adicionar tipo de empresa
        p_tipo = header.add_paragraph()
        p_tipo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_tipo = p_tipo.add_run("Topografia e Consultoria LTDA")
        run_tipo.font.size = Pt(11)
        
        # Adicionar endereço + telefone em uma linha
        p_endereco = header.add_paragraph()
        p_endereco.alignment = WD_ALIGN_PARAGRAPH.CENTER
        endereco = self.dados_empresa.get("endereco", "")
        telefone = self.dados_empresa.get("telefone", "")
        run_endereco = p_endereco.add_run(f"{endereco} Fone {telefone}")
        run_endereco.font.size = Pt(10)
        
        # Adicionar email
        p_email = header.add_paragraph()
        p_email.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_email = p_email.add_run(self.dados_empresa.get("email", ""))
        run_email.font.size = Pt(10)
        
        # Adicionar linha separadora
        p_linha = header.add_paragraph()
        p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_linha = p_linha.add_run("_" * 80)
        run_linha.font.size = Pt(10)
        
        # Espaçamento
        header.add_paragraph()
    
    def gerar_documento(self, dados_finais: Dict[str, Any]) -> bytes:
        """
        Gera o memorial descritivo em formato Word com estrutura descritiva.
        
        Args:
            dados_finais: Dicionário com dados do memorial
        
        Returns:
            bytes: Conteúdo do documento em bytes
        """
        doc = Document()
        
        # ✅ NOVO: Adicionar cabeçalho no header da folha
        self._adicionar_cabecalho(doc)
        
        # Configurar estilo padrão
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        
        # Extrair dados
        proprietario = dados_finais.get("cliente", {}).get("proprietario", "N/A")
        municipio = dados_finais.get("cliente", {}).get("local", "N/A")
        comarca = dados_finais.get("comarca", "N/A")
        trt = dados_finais.get("trt", "N/A")
        perimetro = dados_finais.get("cliente", {}).get("perimetro", "0,00")
        area = dados_finais.get("cliente", {}).get("area", "0,00")
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
        p_proprietario = doc.add_paragraph(f"Proprietário: {proprietario}")
        p_municipio = doc.add_paragraph(f"Município: {municipio}")
        p_comarca = doc.add_paragraph(f"Comarca: {comarca}")
        p_trt = doc.add_paragraph(f"TRT: {trt}")
        p_perimetro = doc.add_paragraph(f"Perímetro: {perimetro} m")
        p_area = doc.add_paragraph(f"Area: {area} m²")
        p_matricula = doc.add_paragraph(f"MAT. {matricula}")
        
        # ============================================================
        # DESCRIÇÃO (UM PARÁGRAFO CONTÍNUO)
        # ============================================================
        p_descricao_titulo = doc.add_paragraph()
        p_descricao_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_desc_titulo = p_descricao_titulo.add_run("DESCRIÇÃO")
        run_desc_titulo.bold = True
        
        # Construir descrição contínua
        descricao_texto = self._construir_descricao_continua(segmentos)
        p_descricao = doc.add_paragraph(descricao_texto)
        p_descricao.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # ============================================================
        # DATA E ASSINATURA
        # ============================================================
        doc.add_paragraph()  # Espaço
        
        p_data = doc.add_paragraph(self._formatar_data(municipio))
        p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()  # Espaço
        
        # Linha de assinatura
        p_assinatura_linha = doc.add_paragraph("_" * 40)
        p_assinatura_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Nome do técnico
        p_nome_tecnico = doc.add_paragraph(self.dados_tecnico.get("nome", ""))
        p_nome_tecnico.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Profissão
        p_profissao = doc.add_paragraph("Técnico em Agropecuária")
        p_profissao.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # CFTA, CPF, TRT
        cfta = self.dados_tecnico.get("cfta", "N/A")
        cpf = self.dados_tecnico.get("cpf", "N/A")
        p_dados_tecnico = doc.add_paragraph(f"CFTA: {cfta} | CPF: {cpf} | TRT: {trt}")
        p_dados_tecnico.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Converter para bytes
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output.getvalue()
    
    def _construir_descricao_continua(self, segmentos: List[Dict[str, Any]]) -> str:
        """
        Constrói a descrição como um parágrafo contínuo com todos os segmentos.
        
        Args:
            segmentos: Lista de segmentos com dados de coordenadas e confrontantes
        
        Returns:
            str: Texto da descrição contínua
        """
        if not segmentos:
            return "Sem dados de segmentos disponíveis."
        
        descricao = []
        
        # Primeiro segmento
        primeiro = segmentos[0]
        coord_y_1 = primeiro.get("coord_y", "N/A")
        coord_x_1 = primeiro.get("coord_x", "N/A")
        confrontante_1 = primeiro.get("confrontante", "N/A")
        matricula_1 = primeiro.get("matricula", "N/A")
        
        # Iniciar descrição
        descricao.append(
            f"Inicia-se a descrição deste perímetro no vértice 1, de coordenadas N(Y) {coord_y_1} m e "
            f"E(X) {coord_x_1} m, situado na divisa com {confrontante_1} (MAT. {matricula_1}); "
            f"deste, segue confrontando com {confrontante_1} (MAT. {matricula_1}), "
            f"com o(s) seguinte(s) azimute(s) e distância(s): "
        )
        
        # Processar cada segmento
        for idx, seg in enumerate(segmentos):
            azimute = seg.get("azimute", "N/A")
            distancia = seg.get("distancia", "N/A")
            
            # Próximo vértice
            prox_idx = idx + 1
            if prox_idx < len(segmentos):
                proximo = segmentos[prox_idx]
            else:
                proximo = segmentos[0]
            
            coord_y_prox = proximo.get("coord_y", "N/A")
            coord_x_prox = proximo.get("coord_x", "N/A")
            vértice_prox = prox_idx + 1 if prox_idx < len(segmentos) else 1
            
            # Adicionar segmento
            descricao.append(
                f"{azimute} e {distancia} m até o vértice {vértice_prox}, de coordenadas N(Y) {coord_y_prox} m "
                f"e E(X) {coord_x_prox} m; "
            )
            
            # Adicionar confrontante se mudar
            if prox_idx < len(segmentos):
                proximo_confrontante = proximo.get("confrontante", "N/A")
                proximo_matricula = proximo.get("matricula", "N/A")
                
                if proximo_confrontante != confrontante_1:
                    descricao.append(
                        f"deste, segue confrontando com {proximo_confrontante} (MAT. {proximo_matricula}), "
                        f"com o(s) seguinte(s) azimute(s) e distância(s): "
                    )
                    confrontante_1 = proximo_confrontante
        
        # Fechar descrição
        descricao.append(f"Fechando o perímetro no vértice 1 de origem.")
        
        return "".join(descricao)
