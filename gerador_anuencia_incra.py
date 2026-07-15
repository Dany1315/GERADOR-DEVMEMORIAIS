# =========================================================================
# MODULO: GERADOR DE ANUÊNCIAS INCRA VIA GEMINI API & python-docx
# =========================================================================

import io
import re
import logging
from datetime import datetime
from typing import Dict, Any, List
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

logger = logging.getLogger(__name__)

class GeradorAnuenciaIncraWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador com os dados da empresa e do técnico.
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico
        
        # Configuração da API do Gemini usando o SDK padrão do app.py
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)

    def _extrair_texto_docx(self, bytes_conteudo: bytes) -> str:
        """Extrai texto de arquivos .docx"""
        try:
            doc = Document(io.BytesIO(bytes_conteudo))
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            logger.error(f"Erro ao ler DOCX: {e}")
            return ""

    def _chamar_gemini_para_analise(self, texto_memorial: str, metadados_usuario: Dict[str, Any]) -> str:
        """
        Envia o memorial cru para o Gemini ler, estruturar e redigir o texto 
        da anuência técnica nos padrões rígidos do INCRA.
        """
        if not self.api_key:
            raise ValueError("Chave API do Gemini (GEMINI_API_KEY) não está configurada.")

        # Instruções estritas para garantir que o Gemini retorne o corpo do texto perfeito, sem inventar limites.
        system_instruction = (
            "Você é um engenheiro agrimensor especialista em georreferenciamento de imóveis rurais certificado pelo INCRA.\n"
            "Sua tarefa é analisar o memorial descritivo fornecido e gerar exclusivamente o texto descritivo oficial de confrontação "
            "para o Termo de Anuência do INCRA, integrando as coordenadas dos limites comuns que constam no texto."
        )

        prompt = f"""
        Com base no memorial descritivo enviado e nas informações do projeto, identifique as linhas de divisa, 
        seus respectivos vértices, azimutes e distâncias, e formule o parágrafo descritivo de confrontação oficial para o INCRA.

        INFORMAÇÕES ADICIONAIS DO PROJETO:
        - Proprietário Requerente: {metadados_usuario.get('proprietario', 'Não Informado')}
        - Imóvel Requerente: {metadados_usuario.get('imovel', 'Não Informado')}
        - Local: {metadados_usuario.get('local', 'Não Informado')}
        - Área Declarada: {metadados_usuario.get('area', 'Não Informada')} ha
        - Perímetro Declarado: {metadados_usuario.get('perimetro', 'Não Informado')} m

        TEXTO DO MEMORIAL DESCRITIVO ENVIADO:
        \"\"\"
        {texto_memorial}
        \"\"\"

        REQUISITO:
        Gere uma redação técnica descritiva bem detalhada do trecho, indicando de forma explícita que as partes concordam 
        com a descrição técnica dos limites comuns georreferenciados apresentados. Escreva um texto formal e fluido.
        """

        try:
            # Mantendo consistência com o Gemini 2.5 Flash de texto
            model = genai.GenerativeModel(
                model_name='gemini-2.5-flash',
                generation_config={"temperature": 0.2}
            )
            response = model.generate_content([system_instruction, prompt])
            return response.text
        except Exception as e:
            logger.error(f"Erro na API do Gemini: {e}")
            raise Exception(f"Erro na comunicação com o cérebro da Inteligência Artificial: {str(e)}")

    def gerar_documento_pelo_memorial(self, arquivo_conteudo: bytes, nome_arquivo: str, metadados_usuario: Dict[str, Any]) -> io.BytesIO:
        """
        Processa o memorial técnico enviado (TXT ou DOCX), invoca a IA para análise do INCRA e constrói o arquivo Word final.
        """
        # 1. Extração do Texto Técnico
        texto_extraido = ""
        extensao = nome_arquivo.split(".")[-1].lower()

        if extensao == "docx":
            texto_extraido = self._extrair_texto_docx(arquivo_conteudo)
        elif extensao in ["txt", "csv"]:
            texto_extraido = arquivo_conteudo.decode("utf-8", errors="ignore")
        elif extensao == "pdf":
            # Caso seja PDF, como o app já tem PyMuPDF ou similar, tratamos como texto de emergência ou instruímos leitura de strings
            try:
                import fitz  # PyMuPDF
                doc_pdf = fitz.open(stream=arquivo_conteudo, filetype="pdf")
                texto_extraido = "\n".join([page.get_text() for page in doc_pdf])
            except ImportError:
                texto_extraido = "Erro: Biblioteca PyMuPDF (fitz) não instalada para ler o PDF das anuências do INCRA."
        else:
            raise ValueError(f"Formato de arquivo .{extensao} não suportado para extração automática.")

        if not texto_extraido.strip():
            raise ValueError("O documento enviado está vazio ou não pôde ser lido corretamente.")

        # 2. IA gera a Redação Técnica nos padrões do INCRA
        redacao_tecnica_ia = self._chamar_gemini_para_analise(texto_extraido, metadados_usuario)

        # 3. Montagem do Word (.docx) usando o layout Premium
        doc = Document()

        # Configuração de Margens Padrão (2.54 cm / 1 polegada)
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # Estilo Base
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        cor_verde_claro = RGBColor(34, 139, 34)
        cor_preta = RGBColor(0, 0, 0)

        # =========================================================================
        # CABEÇALHO NATIVO DO WORD
        # =========================================================================
        secao = doc.sections[0]
        cabecalho_nativo = secao.header
        
        p_head1 = cabecalho_nativo.paragraphs[0]
        p_head1.text = ""
        p_head1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h1 = p_head1.add_run(self.dados_empresa.get("nome", "TopoGeo"))
        run_h1.bold = True
        run_h1.font.size = Pt(12)
        run_h1.font.color.rgb = cor_verde_claro
        
        p_head2 = cabecalho_nativo.add_paragraph()
        p_head2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h2 = p_head2.add_run("Serviços Especializados de Topografia & Engenharia")
        run_h2.bold = True
        run_h2.font.size = Pt(10)
        run_h2.font.color.rgb = cor_verde_claro
        
        p_head3 = cabecalho_nativo.add_paragraph()
        p_head3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h3 = p_head3.add_run(f"{self.dados_empresa.get('endereco', '')} - Fone {self.dados_empresa.get('telefone', '')}")
        run_h3.font.size = Pt(9)
        run_h3.font.color.rgb = cor_preta

        # =========================================================================
        # CORPO DO DOCUMENTO
        # =========================================================================
        
        # Título
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(18)
        p_titulo.paragraph_format.space_after = Pt(24)
        run_titulo = p_titulo.add_run("CARTA DE ANUÊNCIA DOS LIMITES CONFRONTANTES - MODELO INCRA")
        run_titulo.bold = True
        run_titulo.font.size = Pt(12)

        # Abertura Formal
        texto_abertura = (
            f"Nos termos do Art. 9º, § 6º do Decreto nº 4.449/2002, "
            f"declaramos para fins de georreferenciamento e certificação do imóvel rural junto ao INCRA, "
            f"conforme especificações técnicas de limites estabelecidas na Lei nº 10.267/2001:"
        )
        p_aberto = doc.add_paragraph(texto_abertura)
        p_aberto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_aberto.paragraph_format.line_spacing = 1.15
        p_aberto.paragraph_format.space_after = Pt(12)

        # Dados do Imóvel Requerente
        p_dados_tit = doc.add_paragraph()
        p_dados_tit.add_run("I - DO IMÓVEL DO REQUERENTE:").bold = True
        p_dados_tit.paragraph_format.space_after = Pt(4)

        p_dados_txt = doc.add_paragraph(
            f"Imóvel: {metadados_usuario.get('imovel')}\n"
            f"Proprietário: {metadados_usuario.get('proprietario')}\n"
            f"Localização/Município: {metadados_usuario.get('local')}\n"
            f"Área Informada: {metadados_usuario.get('area')} ha    |    Perímetro: {metadados_usuario.get('perimetro')} m"
        )
        p_dados_txt.paragraph_format.left_indent = Inches(0.5)
        p_dados_txt.paragraph_format.space_after = Pt(12)

        # Seção de Redação Técnica Analisada pela IA
        p_analise_tit = doc.add_paragraph()
        p_analise_tit.add_run("II - DESCRIÇÃO DO LIMITE COMUM ACORDADO:").bold = True
        p_analise_tit.paragraph_format.space_after = Pt(6)

        # Inserção da descrição contextualizada do Gemini
        p_ia_corpo = doc.add_paragraph(redacao_tecnica_ia)
        p_ia_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ia_corpo.paragraph_format.line_spacing = 1.15
        p_ia_corpo.paragraph_format.space_after = Pt(18)

        # Declaração técnica e Responsabilidade do Profissional
        rg_rt = self.dados_tecnico.get('rg', '1.936.653')
        codigo_incra = self.dados_tecnico.get('codigo_incra', 'G1D')
        
        texto_tecnico = (
            f"Declaramos e atestamos a conformidade técnica destas divisas sob a responsabilidade do profissional "
            f"{self.dados_tecnico.get('nome')} (CPF: {self.dados_tecnico.get('cpf')}, Registro Profissional CFTA: {self.dados_tecnico.get('cfta')}), "
            f"credenciado ao INCRA sob o código de identificação '{codigo_incra}', o qual emitiu a correspondente guia de "
            f"Termo de Responsabilidade Técnica (TRT) sob o nº {self.dados_tecnico.get('trt')}."
        )
        p_tecnico = doc.add_paragraph(texto_tecnico)
        p_tecnico.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_tecnico.paragraph_format.line_spacing = 1.15
        p_tecnico.paragraph_format.space_after = Pt(18)

        # Data e Local
        data_atual = datetime.now().strftime('%d de %m de %Y')
        p_data = doc.add_paragraph(f"{metadados_usuario.get('local', 'Vila Valerio')}, {data_atual}")
        p_data.paragraph_format.space_after = Pt(28)

        # Assinaturas Lado a Lado
        tab_assinatura = doc.add_table(rows=2, cols=2)
        tab_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        tblPr = tab_assinatura._tbl.tblPr
        limite_xml = f'<w:tblBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>'
        tblBorders = parse_xml(limite_xml)
        tblPr.append(tblBorders)

        celulas_l1 = tab_assinatura.rows[0].cells
        celulas_l1[0].text = "______________________________________"
        celulas_l1[1].text = "______________________________________"
        celulas_l1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        celulas_l1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        celulas_l2 = tab_assinatura.rows[1].cells
        
        # Coluna 1: Proprietário Requerente
        c1 = celulas_l2[0]
        c1.text = ""
        p1_c1 = c1.paragraphs[0]
        p1_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1_c1.paragraph_format.space_before = Pt(4)
        run_name_c1 = p1_c1.add_run(metadados_usuario.get('proprietario', '').title())
        run_name_c1.font.size = Pt(10)
        p2_c1 = c1.add_paragraph()
        p2_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_role_c1 = p2_c1.add_run("Proprietário Requerente (INCRA)")
        run_role_c1.font.size = Pt(9)

        # Coluna 2: Confrontante Acordante
        c2 = celulas_l2[1]
        c2.text = ""
        p1_c2 = c2.paragraphs[0]
        p1_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1_c2.paragraph_format.space_before = Pt(4)
        run_name_c2 = p1_c2.add_run("____________________________________")
        run_name_c2.font.size = Pt(10)
        p2_c2 = c2.add_paragraph()
        p2_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_role_c2 = p2_c2.add_run("Proprietário do Imóvel Confrontante Acordante")
        run_role_c2.font.size = Pt(9)

        # Assinatura do Responsável Técnico Centralizado no Fim
        p_rt_l1 = doc.add_paragraph()
        p_rt_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_rt_l1.paragraph_format.space_before = Pt(24)
        p_rt_l1.add_run("________________________")
        
        p_rt_l2 = doc.add_paragraph()
        p_rt_l2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_nome_tec = p_rt_l2.add_run(f"{self.dados_tecnico.get('nome')}")
        run_nome_tec.bold = True
        
        p_rt_l3 = doc.add_paragraph()
        p_rt_l3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_cargo = p_rt_l3.add_run(f"Responsável Técnico Credenciado INCRA (CFTA: {self.dados_tecnico.get('cfta')})")
        run_cargo.font.size = Pt(9)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
