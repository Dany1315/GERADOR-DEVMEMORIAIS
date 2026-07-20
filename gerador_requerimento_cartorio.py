#ATENÇÃO NA GERAÇÃO DE REQUERIMENTOS E DESORDEM DE DOCUMENTOS
import io
import json
import logging
from typing import List, Dict, Any
from docx import Document
import google.generativeai as genai
from PIL import Image

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
        Analise as imagens dos documentos fornecidos (RG, CPF, Certidões, Matrículas, etc.) e extraia as seguintes informações para preencher um requerimento de cartório. 
        Retorne os dados estritamente em formato JSON, conforme o esquema abaixo:

        {
            "comarca": "Nome da Comarca",
            "requerente_1": {
                "nome": "Nome Completo",
                "estado_civil": "Estado Civil",
                "profissao": "Profissão",
                "rg": "Número do RG",
                "orgao_emissor": "Órgão Emissor/UF",
                "cpf": "000.000.000-00",
                "endereco": "Endereço Completo (Córrego, Zona, Município-UF)"
            },
            "requerente_2": {
                "nome": "Nome Completo da Esposa/Cônjuge (se houver)",
                "profissao": "Profissão",
                "rg": "Número do RG",
                "orgao_emissor": "Órgão Emissor/UF",
                "cpf": "000.000.000-00",
                "regime_bens": "Regime de Comunhão (ex: Parcial)"
            },
            "imovel": {
                "nome": "Nome do Sítio/Fazenda",
                "area_registrada": "0,0000",
                "municipio": "Nome do Município",
                "comarca_imovel": "Nome da Comarca do Imóvel",
                "matricula": "00.000",
                "valor_fiscal": "000.000,00",
                "valor_extenso": "valor por extenso",
                "area_encontrada": "0,0000",
                "codigo_incra": "000.000.000.000-0",
                "trt_numero": "BR0000000",
                "area_gleba_1": "0,0000",
                "area_gleba_2": "0,0000",
                "area_estrada": "0.000,00",
                "area_total_retificada": "0,0000"
            },
            "data": {
                "dia": "DD",
                "mes": "Nome do Mês",
                "ano": "AAAA"
            }
        }

        Se alguma informação não for encontrada, deixe o valor como "XXXXXX".
        Certifique-se de que o JSON seja válido.
        """
        
        try:
            # Prepara o conteúdo para o Gemini (Imagens + Prompt)
            conteudo = [prompt] + imagens
            response = self.model.generate_content(conteudo)
            
            # Limpa a resposta para extrair apenas o JSON
            text_response = response.text
            if "```json" in text_response:
                text_response = text_response.split("```json")[1].split("```")[0]
            elif "```" in text_response:
                text_response = text_response.split("```")[1].split("```")[0]
            
            dados = json.loads(text_response.strip())
            return dados
        except Exception as e:
            logger.error(f"Erro na extração via Gemini: {e}")
            raise Exception(f"Falha ao extrair dados dos documentos: {str(e)}")

    def gerar_documento(self, dados: Dict[str, Any], template_path: str) -> io.BytesIO:
        """
        Preenche o template Word com os dados extraídos.
        """
        try:
            doc = Document(template_path)
            
            # Mapeamento de substituições
            # Nota: Esta é uma implementação simplificada. Para produção, 
            # o ideal é iterar sobre parágrafos e tabelas.
            
            substituicoes = {
                "XXXXXXX – ES": f"{dados['comarca']} – ES",
                "XXXXXX, proprietário": f"{dados['requerente_1']['nome']}, proprietário",
                "XXXXX, lavrador": f"{dados['requerente_1']['profissao']}, lavrador",
                "C.I. n°. XXXX – SSP/ES": f"C.I. n°. {dados['requerente_1']['rg']} – {dados['requerente_1']['orgao_emissor']}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {dados['requerente_1']['cpf']}",
                "esposa XXXXXX": f"esposa {dados['requerente_2']['nome']}",
                "XXXXX – SSP/ES": f"{dados['requerente_2']['rg']} – {dados['requerente_2']['orgao_emissor']}",
                "CPF/MF n° XXXXXX": f"CPF/MF n° {dados['requerente_2']['cpf']}",
                "comunhão XXXXXX de bens": f"comunhão {dados['requerente_2']['regime_bens']} de bens",
                "Córrego XXXXX": f"Córrego {dados['requerente_1']['endereco'].split(',')[0]}",
                "XXXXXX-ES": f"{dados['requerente_1']['endereco'].split('-')[-1]}",
                "Sitio XXXXX": f"Sitio {dados['imovel']['nome']}",
                "registrada de XXXXXX ha": f"registrada de {dados['imovel']['area_registrada']} ha",
                "município de XXXX": f"município de {dados['imovel']['municipio']}",
                "comarca de XXXXXXX - ES": f"comarca de {dados['imovel']['comarca_imovel']} - ES",
                "matrícula n°. XXXXXX": f"matrícula n°. {dados['imovel']['matricula']}",
                "R$ XX0.000,00": f"R$ {dados['imovel']['valor_fiscal']}",
                "(XXXXXX mil reais)": f"({dados['imovel']['valor_extenso']})",
                "área de XXXXX ha": f"área de {dados['imovel']['area_encontrada']} ha",
                "n°. XXX.XXX.XXX.XXX-X": f"n°. {dados['imovel']['codigo_incra']}",
                "TRT BRXXXXXXX": f"TRT {dados['imovel']['trt_numero']}",
                "Gleba 1” com área de XXXXXX ha": f"Gleba 1” com área de {dados['imovel']['area_gleba_1']} ha",
                "Gleba 2” com área de XXXXX ha": f"Gleba 2” com área de {dados['imovel']['area_gleba_2']} ha",
                "Municipal de X.XXX,XX m²": f"Municipal de {dados['imovel']['area_estrada']} m²",
                "encontrada de XXXXXXX ha": f"encontrada de {dados['imovel']['area_total_retificada']} ha",
                "XX de XXX de XXXX": f"{dados['data']['dia']} de {dados['data']['mes']} de {dados['data']['ano']}",
                "XXXXXXXXXXXXXXXX": dados['requerente_1']['nome'],
                "XXXXXXXXXXXXXXXXXX": dados['requerente_2']['nome'],
                "CPF: XXX.XXX.XXX-XX": f"CPF: {dados['requerente_1']['cpf']}"
            }

            for p in doc.paragraphs:
                for key, value in substituicoes.items():
                    if key in p.text:
                        p.text = p.text.replace(key, str(value))
            
            # Trata tabelas se houver
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, value in substituicoes.items():
                                if key in p.text:
                                    p.text = p.text.replace(key, str(value))

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer
        except Exception as e:
            logger.error(f"Erro ao gerar Word do requerimento: {e}")
            raise e
