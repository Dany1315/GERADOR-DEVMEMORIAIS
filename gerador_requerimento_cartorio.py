import io
import json
import logging
import os
from typing import List, Dict, Any
from docx import Document
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
            raise Exception(f"Falha ao extrair dados dos documentos: {str(e)}")

    def _substituir_texto_no_paragrafo(self, p, substituicoes):
        """
        Substitui texto em um parágrafo mantendo a formatação o melhor possível.
        """
        for key, value in substituicoes.items():
            if key in p.text:
                # Substituição simples no texto completo do parágrafo
                # Nota: Isso pode perder formatação específica de palavras, 
                # mas é mais confiável para capturar chaves que o Word divide em múltiplos 'runs'
                inline = p.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        text = inline[i].text.replace(key, str(value))
                        inline[i].text = text
                
                # Se a chave ainda estiver lá (dividida entre runs), fazemos a troca bruta
                if key in p.text:
                    p.text = p.text.replace(key, str(value))

    def gerar_documento(self, dados: Dict[str, Any], template_name: str) -> io.BytesIO:
        """
        Preenche o template Word com os dados extraídos.
        """
        try:
            # Tenta encontrar o arquivo no diretório atual ou no diretório do script
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, template_name)
            
            if not os.path.exists(template_path):
                # Tenta diretório atual de trabalho se o acima falhar
                template_path = template_name
            
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Modelo não encontrado: {template_name}. Certifique-se de que ele está na mesma pasta do projeto.")

            doc = Document(template_path)
            
            substituicoes = {
                "XXXXXXX – ES": f"{dados.get('comarca', 'XXXXXX')} – ES",
                "XXXXXX, proprietário": f"{dados['requerente_1'].get('nome', 'XXXXXX')}, proprietário",
                "XXXXX, lavrador": f"{dados['requerente_1'].get('profissao', 'XXXXXX')}, lavrador",
                "C.I. n°. XXXX – SSP/ES": f"C.I. n°. {dados['requerente_1'].get('rg', 'XXXX')} – {dados['requerente_1'].get('orgao_emissor', 'SSP/ES')}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {dados['requerente_1'].get('cpf', 'XXXXXXX')}",
                "esposa XXXXXX": f"esposa {dados['requerente_2'].get('nome', 'XXXXXX')}",
                "XXXXX – SSP/ES": f"{dados['requerente_2'].get('rg', 'XXXXX')} – {dados['requerente_2'].get('orgao_emissor', 'SSP/ES')}",
                "CPF/MF n° XXXXXX": f"CPF/MF n° {dados['requerente_2'].get('cpf', 'XXXXXX')}",
                "comunhão XXXXXX de bens": f"comunhão {dados['requerente_2'].get('regime_bens', 'XXXXXX')} de bens",
                "Córrego XXXXX": f"Córrego {dados['requerente_1'].get('endereco', 'XXXXX').split(',')[0]}",
                "Sitio XXXXX": f"Sitio {dados['imovel'].get('nome', 'XXXXX')}",
                "registrada de XXXXXX ha": f"registrada de {dados['imovel'].get('area_registrada', 'XXXXXX')} ha",
                "município de XXXX": f"município de {dados['imovel'].get('municipio', 'XXXX')}",
                "comarca de XXXXXXX - ES": f"comarca de {dados['imovel'].get('comarca_imovel', 'XXXXXXX')} - ES",
                "matrícula n°. XXXXXX": f"matrícula n°. {dados['imovel'].get('matricula', 'XXXXXX')}",
                "R$ XX0.000,00": f"R$ {dados['imovel'].get('valor_fiscal', 'XX0.000,00')}",
                "(XXXXXX mil reais)": f"({dados['imovel'].get('valor_extenso', 'XXXXXX mil reais')})",
                "área de XXXXX ha": f"área de {dados['imovel'].get('area_encontrada', 'XXXXX')} ha",
                "n°. XXX.XXX.XXX.XXX-X": f"n°. {dados['imovel'].get('codigo_incra', 'XXX.XXX.XXX.XXX-X')}",
                "TRT BRXXXXXXX": f"TRT {dados['imovel'].get('trt_numero', 'BRXXXXXXX')}",
                "Gleba 1” com área de XXXXXX ha": f"Gleba 1” com área de {dados['imovel'].get('area_gleba_1', 'XXXXXX')} ha",
                "Gleba 2” com área de XXXXX ha": f"Gleba 2” com área de {dados['imovel'].get('area_gleba_2', 'XXXXX')} ha",
                "Municipal de X.XXX,XX m²": f"Municipal de {dados['imovel'].get('area_estrada', 'X.XXX,XX')} m²",
                "encontrada de XXXXXXX ha": f"encontrada de {dados['imovel'].get('area_total_retificada', 'XXXXXXX')} ha",
                "XX de XXX de XXXX": f"{dados['data'].get('dia', 'XX')} de {dados['data'].get('mes', 'XXX')} de {dados['data'].get('ano', 'XXXX')}",
                "XXXXXXXXXXXXXXXX": dados['requerente_1'].get('nome', 'XXXXXXXXXXXXXXXX'),
                "XXXXXXXXXXXXXXXXXX": dados['requerente_2'].get('nome', 'XXXXXXXXXXXXXXXXXX'),
                "CPF: XXX.XXX.XXX-XX": f"CPF: {dados['requerente_1'].get('cpf', 'XXX.XXX.XXX-XX')}"
            }

            for p in doc.paragraphs:
                self._substituir_texto_no_paragrafo(p, substituicoes)
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            self._substituir_texto_no_paragrafo(p, substituicoes)

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer
        except Exception as e:
            logger.error(f"Erro ao gerar Word do requerimento: {e}")
            raise e
