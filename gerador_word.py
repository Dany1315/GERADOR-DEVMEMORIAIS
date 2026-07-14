"""
Gerador de documentos Word com formatação profissional.
"""

import io
import logging
from typing import Dict, Any
from datetime import datetime

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor

from config import (
    DOCUMENTO_CONFIG,
    EMPRESA_CONFIG,
    TECNICO_CONFIG,
    MESES_PT
)
from utils import criar_logger

logger = criar_logger(__name__)


class GeradorMemorialWord:
    """
    Gera documentos Word formatados com memoriais descritivos.
    
    Responsabilidades:
    - Criação de documento Word estruturado
    - Aplicação de formatação profissional
    - Inclusão de dados cadastrais e técnicos
    - Descrição textual do perímetro
    """

    def __init__(self, dados_empresa: Dict, dados_tecnico: Dict):
        """
        Inicializa gerador com dados padrão.
        
        Args:
            dados_empresa: Dict com informações da empresa
            dados_tecnico: Dict com informações do técnico
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
        logger.info("GeradorMemorialWord inicializado")

    def gerar_documento(self, dados_finais: Dict[str, Any]) -> io.BytesIO:
        """
        Gera documento Word completo com memorial descritivo.
        
        Args:
            dados_finais: Dict com todos os dados a incluir no documento
        
        Returns:
            BytesIO com conteúdo do arquivo Word
            
        Raises:
            Exception: Se erro ao gerar documento
        """
        try:
            logger.info("Iniciando geração do documento Word...")
            doc = docx.Document()

            # Configurar margens
            self._configurar_margens(doc)
            
            # Aplicar estilos padrão
            self._aplicar_estilos_padrao(doc)

            # Cabeçalho
            self._criar_cabecalho(doc)

            # Título
            self._criar_titulo(doc)

            # Dados cadastrais
            self._criar_secao_dados_cadastrais(doc, dados_finais)

            # Descrição do perímetro
            self._criar_secao_perimetro(doc, dados_finais)

            # Data e local
            self._criar_secao_data(doc, dados_finais)

            # Assinatura
            self._criar_secao_assinatura(doc)

            # Salvar em BytesIO
            conteudo_arquivo = io.BytesIO()
            doc.save(conteudo_arquivo)
            conteudo_arquivo.seek(0)

            logger.info("✅ Documento Word gerado com sucesso")
            return conteudo_arquivo

        except Exception as e:
            logger.error(f"Erro ao gerar documento Word: {str(e)}")
            raise

    def _configurar_margens(self, doc):
        """Configura margens do documento."""
        margem = Cm(DOCUMENTO_CONFIG.MARGENS_CM)
        for section in doc.sections:
            section.top_margin = margem
            section.bottom_margin = margem
            section.left_margin = margem
            section.right_margin = margem
        logger.debug("Margens configuradas")

    def _aplicar_estilos_padrao(self, doc):
        """Aplica estilos padrão ao documento."""
        style = doc.styles['Normal']
        style.font.name = DOCUMENTO_CONFIG.FONTE_PADRAO
        style.font.size = Pt(DOCUMENTO_CONFIG.TAMANHO_FONTE_PADRAO)
        logger.debug("Estilos padrão aplicados")

    def _criar_cabecalho(self, doc):
        """Cria cabeçalho com informações da empresa."""
        section = doc.sections[0]
        header = section.header
        header.is_linked_to_previous = False

        # Logo/Nome da empresa
        p_logo = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(0)

        run_logo = p_logo.add_run("TopoGeo")
        run_logo.font.size = Pt(14)
        run_logo.bold = True
        run_logo.font.color.rgb = RGBColor(0, 128, 0)  # Verde

        # Informações da empresa
        p_empresa = header.add_paragraph()
        p_empresa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_empresa.paragraph_format.space_after = Pt(6)

        empresa_info = self.dados_empresa
        p_empresa.add_run("Topografia e Consultoria LTDA\n")
        p_empresa.add_run(f"{empresa_info.get('endereco', '')}\n")
        p_empresa.add_run(f"Fone {empresa_info.get('telefone', '')} - {empresa_info.get('email', '')}")

        # Linha separadora
        p_linha = header.add_paragraph()
        p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linha.add_run("_" * 80)
        p_linha.paragraph_format.space_after = Pt(12)

        logger.debug("Cabeçalho criado")

    def _criar_titulo(self, doc):
        """Cria título do documento."""
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(12)
        p_titulo.paragraph_format.space_after = Pt(18)

        run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
        run_tit.bold = True
        run_tit.font.size = Pt(DOCUMENTO_CONFIG.TAMANHO_TITULO)

        logger.debug("Título criado")

    def _criar_secao_dados_cadastrais(self, doc, dados_finais: Dict):
        """Cria seção com dados cadastrais."""
        p_dados = doc.add_paragraph()
        p_dados.paragraph_format.line_spacing = 1.15
        p_dados.paragraph_format.space_after = Pt(18)

        campos = [
            ("Imóvel: ", dados_finais.get('imovel', 'NÃO INFORMADO')),
            ("Proprietário: ", dados_finais.get('proprietario', 'NÃO INFORMADO')),
            ("Local: ", dados_finais.get('local', 'NÃO INFORMADO')),
            ("Área (ha): ", dados_finais.get('area', 'NÃO INFORMADO')),
            ("Perímetro (m): ", dados_finais.get('perimetro', 'NÃO INFORMADO')),
        ]

        for i, (label, valor) in enumerate(campos):
            run_label = p_dados.add_run(label)
            run_label.bold = True
            p_dados.add_run(f"{valor}")
            if i < len(campos) - 1:
                p_dados.add_run("\n")

        logger.debug("Seção de dados cadastrais criada")

    def _criar_secao_perimetro(self, doc, dados_finais: Dict):
        """Cria seção de descrição do perímetro."""
        # Título da seção
        p_desc_tit = doc.add_paragraph()
        p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_desc_tit.paragraph_format.space_before = Pt(12)
        p_desc_tit.paragraph_format.space_after = Pt(12)

        run_desc = p_desc_tit.add_run("DESCRIÇÃO DO PERÍMETRO")
        run_desc.bold = True

        # Texto descritivo
        p_texto = doc.add_paragraph()
        p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_texto.paragraph_format.line_spacing = 1.25
        p_texto.paragraph_format.space_after = Pt(12)

        segmentos = dados_finais.get("segmentos", [])

        if segmentos:
            primeiro = segmentos[0]
            p_texto.add_run(
                f"Inicia-se a descrição deste perímetro no vértice {primeiro['de']}, "
                f"de coordenadas N {primeiro['n_y']} e E {primeiro['e_x']}; "
            )

            for i, s in enumerate(segmentos):
                if i + 1 < len(segmentos):
                    prox_n = segmentos[i + 1]['n_y']
                    prox_e = segmentos[i + 1]['e_x']
                else:
                    prox_n = segmentos[0]['n_y']
                    prox_e = segmentos[0]['e_x']

                confrontante = s.get('confrontante', 'NÃO INFORMADO')
                azimute = s.get('azimute', 'NÃO INFORMADO')
                distancia = s.get('distancia', 'NÃO INFORMADO')

                if i == len(segmentos) - 1:
                    # Último segmento
                    p_texto.add_run(
                        f" {azimute} e {distancia} até o vértice {s['para']}, "
                    )
                else:
                    p_texto.add_run(
                        f"Divisa do imóvel; deste, segue confrontando com {confrontante}, "
                        f"com os seguintes azimutes e distâncias: {azimute} e {distancia} "
                        f"até o vértice {s['para']}, de coordenadas N {prox_n} e E {prox_e}; "
                    )
        else:
            logger.warning("Nenhum segmento disponível para descrição")
            p_texto.add_run("Nenhum segmento foi processado.")

        # Texto final padrão
        p_texto.add_run(
            "ponto inicial da descrição deste perímetro. As coordenadas da base foram "
            "processadas pelo método de Posicionamento por Ponto Preciso (PPP). Todas as "
            "coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro "
            "e encontram-se representadas no Sistema U T M, referenciadas ao Meridiano Central "
            "nº 39°00', fuso -24, tendo como datum o SIRGAS2000. Todos os azimutes e distâncias, "
            "área e perímetro foram calculados no plano de projeção U T M."
        )

        logger.debug("Seção de perímetro criada")

    def _criar_secao_data(self, doc, dados_finais: Dict):
        """Cria seção com data e local."""
        p_data = doc.add_paragraph()
        p_data.paragraph_format.space_before = Pt(24)

        local = dados_finais.get('local', 'Vila Valério')
        data_atual = datetime.now()
        data_formatada = data_atual.strftime('%d/%m/%Y')

        p_data.add_run(f"{local} – ES, {data_formatada}.")

        logger.debug("Seção de data criada")

    def _criar_secao_assinatura(self, doc):
        """Cria seção de assinatura."""
        p_assinatura = doc.add_paragraph()
        p_assinatura.paragraph_format.space_before = Pt(36)

        tecnico_info = self.dados_tecnico
        p_assinatura.add_run(
            f"__________________________________\n"
            f"                {tecnico_info.get('nome', 'NÃO INFORMADO')}\n"
            f"             Resp. Técnico\n"
            f"               CFTA: {tecnico_info.get('cfta', 'NÃO INFORMADO')}\n"
            f"Credenciamento INCRA: G1D"
        )

        logger.debug("Seção de assinatura criada")


class GeradorAnuenciaWord:
    """
    Gera documentos Word formatados com termos de anuência.
    """

    def __init__(self, dados_empresa: Dict, dados_tecnico: Dict):
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico

    def gerar_documento(self, dados_anuencia: Dict[str, Any]) -> io.BytesIO:
        """
        Gera termo de anuência.
        """
        try:
            doc = docx.Document()
            
            # Margens e estilos
            margem = Cm(DOCUMENTO_CONFIG.MARGENS_CM)
            for section in doc.sections:
                section.top_margin = margem
                section.bottom_margin = margem
                section.left_margin = margem
                section.right_margin = margem
            
            style = doc.styles['Normal']
            style.font.name = DOCUMENTO_CONFIG.FONTE_PADRAO
            style.font.size = Pt(DOCUMENTO_CONFIG.TAMANHO_FONTE_PADRAO)

            # Título
            p_titulo = doc.add_paragraph()
            p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_tit = p_titulo.add_run("DECLARAÇÃO DE RECONHECIMENTO DE LIMITES")
            run_tit.bold = True
            run_tit.font.size = Pt(12)

            # Texto inicial
            p_texto = doc.add_paragraph()
            p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_texto.add_run(f"Eu, ").add_run(f"{dados_anuencia['confrontante']}").bold = True
            p_texto.add_run(", proprietário do imóvel confrontante, e eu, ")
            p_texto.add_run(f"{dados_anuencia['proprietario']}").bold = True
            p_texto.add_run(", proprietário do imóvel rural, declaramos não existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis.")

            p_trecho = doc.add_paragraph()
            p_trecho.add_run("Descrição do trecho de confrontação:").bold = True

            # Tabela
            table = doc.add_table(rows=1, cols=7)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'De'
            hdr_cells[1].text = 'Para'
            hdr_cells[2].text = 'Azimute'
            hdr_cells[3].text = 'Distância (m)'
            hdr_cells[4].text = 'E(X)'
            hdr_cells[5].text = 'N(Y)'
            hdr_cells[6].text = 'Altitude'

            distancia_total = 0.0
            for seg in dados_anuencia['segmentos']:
                row_cells = table.add_row().cells
                row_cells[0].text = str(seg['de'])
                row_cells[1].text = str(seg['para'])
                row_cells[2].text = str(seg['azimute'])
                row_cells[3].text = str(seg['distancia']).replace(' m', '').replace(',', '.')
                row_cells[4].text = str(seg['e_x']).replace(' m', '')
                row_cells[5].text = str(seg['n_y']).replace(' m', '')
                row_cells[6].text = '-'
                
                try:
                    d_val = float(row_cells[3].text)
                    distancia_total += d_val
                except:
                    pass

            # Linha de total
            row_total = table.add_row().cells
            row_total[0].text = 'TOTAL'
            row_total[3].text = f"{distancia_total:.2f}".replace('.', ',')

            # Texto final
            p_final = doc.add_paragraph()
            p_final.paragraph_format.space_before = Pt(12)
            p_final.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_final.add_run(f"Declaramos ainda que o profissional {self.dados_tecnico['nome']} (RG n° 1.936.653 e CPF n° 111.985.197-11), Resp. Técnico (CFTA {self.dados_tecnico['cfta']}), credenciado pelo INCRA sob o cod. G1D, nos indicou as demarcações do limite entre as nossas propriedades, tanto no campo como nas suas apresentações gráficas.\n")
            p_final.add_run("Concordamos com essa demarcação, expressa na planta e no memorial descritivo, ambos em anexo, e reconhecemos esta descrição como o limite legal entre nossas propriedades.")

            # Data
            p_data = doc.add_paragraph()
            p_data.paragraph_format.space_before = Pt(12)
            data_atual = datetime.now()
            nome_mes = MESES_PT[data_atual.month]
            p_data.add_run(f"{dados_anuencia['local']}, {data_atual.day} de {nome_mes} de {data_atual.year}")

            # Assinaturas
            p_ass = doc.add_paragraph()
            p_ass.paragraph_format.space_before = Pt(24)
            p_ass.add_run("______________________________\n")
            p_ass.add_run("______________________________\n")
            p_ass.add_run(f"{dados_anuencia['confrontante']}\n").bold = True
            p_ass.add_run(f"{dados_anuencia['proprietario']}\n").bold = True
            p_ass.add_run("Proprietário do Imóvel Confrontante\n")
            p_ass.add_run("Proprietário do Imóvel")

            p_ass_tec = doc.add_paragraph()
            p_ass_tec.paragraph_format.space_before = Pt(12)
            p_ass_tec.add_run("___________________________________\n")
            p_ass_tec.add_run(f"{self.dados_tecnico['nome']}\n")
            p_ass_tec.add_run("Resp. Técnico\n")
            p_ass_tec.add_run(f"CFTA: {self.dados_tecnico['cfta']} | TRT: BR20260600550")

            conteudo_arquivo = io.BytesIO()
            doc.save(conteudo_arquivo)
            conteudo_arquivo.seek(0)
            return conteudo_arquivo

        except Exception as e:
            logger.error(f"Erro ao gerar anuência: {str(e)}")
            raise
