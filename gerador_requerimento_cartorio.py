import io
import json
import logging
import os
from datetime import datetime
from typing import List, Dict, Any
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai

logger = logging.getLogger(__name__)

class GeradorRequerimentoCartorio:
    def __init__(self, model_name: str):
        self.model_name = model_name
        self.model = genai.GenerativeModel(model_name)

    def extrair_dados_documentos(self, imagens: List[Any]) -> Dict[str, Any]:
        """
        Envia as imagens dos documentos para o Gemini e extrai os dados estruturados.
        """
        prompt = """
        Analise as imagens dos documentos fornecidos e extraia as informações para um requerimento de cartório.
        REGRAS IMPORTANTES:
        1. Formate RGs com pontos: 706786 -> 706.786 ou 1706786 -> 1.706.786.
        2. Identifique o sexo do requerente para ajustar 'lavrador/lavradora' ou 'agricultor/agricultora'.
        3. Se não encontrar o cônjuge, preencha com 'XXXXXX'.
        4. Identifique a TRT (começa com BR e tem 11 números).
        5. Identifique a Comarca, Município e Matrícula.
        6. Extraia a área total retificada (encontrada na planta INCRA).
        
        Retorne estritamente em JSON:
        {
            "comarca": "NOME DA COMARCA EM MAIUSCULO",
            "municipio_cliente": "Cidade do Cliente",
            "requerente_1": {
                "nome": "Nome Completo",
                "profissao": "Lavrador ou Lavradora",
                "rg": "0.000.000",
                "orgao": "SSP/ES",
                "cpf": "000.000.000-00",
                "endereco_corrego": "Nome do Córrego (Apenas o nome)"
            },
            "requerente_2": {
                "nome": "Nome Completo da Esposa ou XXXXXX",
                "profissao": "Lavradora ou XXXXXX",
                "rg": "0.000.000 ou XXXXXX",
                "orgao": "SSP/ES ou XXXXXX",
                "cpf": "000.000.000-00 ou XXXXXX",
                "regime_bens": "Comunhão Parcial de Bens ou XXXXXX"
            },
            "imovel": {
                "nome": "Nome do Sítio (Apenas o nome)",
                "area_registrada": "0,0000",
                "municipio_imovel": "Vila Valério",
                "comarca_imovel": "SÃO GABRIEL DA PALHA",
                "matricula": "0.000",
                "area_encontrada": "0,0000",
                "codigo_incra": "000.000.000.000-0",
                "trt_numero": "BR00000000000",
                "area_total_retificada": "0,0000"
            }
        }
        """
        
        try:
            conteudo = [prompt] + imagens
            response = self.model.generate_content(conteudo)
            
            text_response = response.text
            if "```json" in text_response:
                text_response = text_response.split("```json")[1].split("```")[0]
            elif "```" in text_response:
                text_response = text_response.split("```")[1].split("```")[0]
            
            dados = json.loads(text_response.strip())
            return dados
        except Exception as e:
            logger.error(f"Erro na extração via Gemini: {e}")
            raise Exception(f"Falha ao extrair dados: {str(e)}")

    def _ajustar_fonte_arial(self, doc):
        """Ajusta a fonte de todo o documento para Arial."""
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(11)

    def gerar_documento(self, dados: Dict[str, Any], template_name: str) -> io.BytesIO:
        try:
            # Busca do template
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, template_name)
            if not os.path.exists(template_path):
                template_path = template_name
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Modelo {template_name} não encontrado.")

            doc = Document(template_path)
            
            # Data de hoje formatada
            hoje = datetime.now()
            meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            data_formatada = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

            # Mapeamento de substituições (Limpando nomes repetidos)
            substituicoes = {
                "COMARCA DE XXXXXXX – ES": f"COMARCA DE {dados.get('comarca', 'XXXXXX').upper()} – ES",
                "XXXXXX, proprietário": f"{dados['requerente_1'].get('nome', 'XXXXXX')}, proprietário",
                "XXXXX, lavrador": f"{dados['requerente_1'].get('profissao', 'lavrador')}",
                "C.I. n°. XXXX – SSP/ES": f"C.I. n°. {dados['requerente_1'].get('rg', 'XXXX')} – {dados['requerente_1'].get('orgao', 'SSP/ES')}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {dados['requerente_1'].get('cpf', 'XXXXXXX')}",
                "esposa XXXXXX": f"esposa {dados['requerente_2'].get('nome', 'XXXXXX')}",
                "XXXXX – SSP/ES": f"{dados['requerente_2'].get('rg', 'XXXXX')} – {dados['requerente_2'].get('orgao', 'SSP/ES')}",
                "CPF/MF n° XXXXXX": f"CPF/MF n° {dados['requerente_2'].get('cpf', 'XXXXXX')}",
                "comunhão XXXXXX de bens": f"comunhão {dados['requerente_2'].get('regime_bens', 'XXXXXX')}",
                "Córrego XXXXX": f"Córrego {dados['requerente_1'].get('endereco_corrego', 'XXXXX')}",
                "Zona Rural, XXXXXX-ES": f"Zona Rural, {dados.get('municipio_cliente', 'XXXXXX')}-ES",
                "Sitio XXXXX": f"Sítio {dados['imovel'].get('nome', 'XXXXX')}",
                "registrada de XXXXXX ha": f"registrada de {dados['imovel'].get('area_registrada', 'XXXXXX')} ha",
                "município de XXXX - ES": f"município de {dados['imovel'].get('municipio_imovel', 'XXXX')} - ES",
                "comarca de XXXXXXX - ES": f"comarca de {dados['imovel'].get('comarca_imovel', 'XXXXXXX')} - ES",
                "matrícula n°. XXXXXX": f"matrícula n°. {dados['imovel'].get('matricula', 'XXXXXX')}",
                "área de XXXXX ha": f"área de {dados['imovel'].get('area_encontrada', 'XXXXX')} ha",
                "n°. XXX.XXX.XXX.XXX-X": f"n°. {dados['imovel'].get('codigo_incra', 'XXX.XXX.XXX.XXX-X')}",
                "TRT BRXXXXXXX": f"TRT {dados['imovel'].get('trt_numero', 'BRXXXXXXX')}",
                "encontrada de XXXXXXX ha": f"encontrada de {dados['imovel'].get('area_total_retificada', 'XXXXXXX')} ha",
                "XX de XXX de XXXX": data_formatada,
                "XXXXXXXXXXXXXXXX": dados['requerente_1'].get('nome', 'XXXXXXXXXXXXXXXX'),
                "XXXXXXXXXXXXXXXXXX": dados['requerente_2'].get('nome', 'XXXXXXXXXXXXXXXXXX'),
                "CPF: XXX.XXX.XXX-XX": f"CPF: {dados['requerente_1'].get('cpf', 'XXX.XXX.XXX-XX')}"
            }

            # Aplicando substituições
            for p in doc.paragraphs:
                for key, val in substituicoes.items():
                    if key in p.text:
                        p.text = p.text.replace(key, str(val))
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, val in substituicoes.items():
                                if key in p.text:
                                    p.text = p.text.replace(key, str(val))
            
            # Ajuste final de fonte
            self._ajustar_fonte_arial(doc)

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer
        except Exception as e:
            logger.error(f"Erro ao gerar documento: {e}")
            raise e
