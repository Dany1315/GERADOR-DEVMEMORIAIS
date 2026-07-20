import io
import json
import logging
import os
import re
from datetime import datetime
from typing import List, Dict, Any
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai

logger = logging.getLogger(__name__)


# ============================================================
# NOMES FEMININOS COMUNS (para detecção de gênero)
# ============================================================
NOMES_FEMININOS = {
    "maria", "ana", "carla", "fernanda", "juliana", "camila",
    "patricia", "andrea", "adriana", "claudia", "denise",
    "elisabete", "eliane", "fabiana", "gabriela", "helena",
    "inez", "jaqueline", "katia", "laura", "lucia", "margarete",
    "marcia", "marta", "monica", "nair", "olivia", "paula",
    "quiteria", "rosa", "silvana", "solange", "teresa",
    "vania", "walquiria", "yasmin", "zuleica", "beatriz",
    "carolina", "daniela", "edite", "fatima", "gisele", "heloisa",
    "isabela", "janice", "lilian", "margaret", "neusa", "odete",
    "priscila", "raquel", "sandra", "tatiane", "ursula", "vitoria",
    "wilma", "yolanda", "zilda", "aline", "bianca", "cristina",
    "diana", "evelin", "flavia", "graziela", "hellen", "ingrid",
    "jessica", "karen", "leticia", "mirela", "natasha", "pamela",
    "renata", "sabrina", "vivian", "angela", "barbara", "cecilia",
    "dolores", "estela", "francisca", "gloria", "hilda", "irma",
    "josefina", "kelly", "lorena", "marina", "paulina", "rita",
    "sueli", "tania", "vera", "wanda", "xenia", "zoraide",
    "amanda", "brenda", "clarice", "daisy", "elenice", "fani",
    "gina", "hebe", "iracema", "joyce", "kika", "melissa",
    "natalia", "rafaela", "silvia", "taina", "vanessa",
    "agostinha", "aparecida", "bete", "creuza", "dalva", "edna",
    "filomena", "gracinda", "ilza", "jandira", "kely", "luana",
    "miriam", "nara", "roseli", "simone", "vanda", "waleska",
    "yanira", "zelia", "ana clara", "vera lucia"
}


def detectar_genero_por_nome(nome: str) -> str:
    """Detecta o gênero de uma pessoa com base no primeiro nome."""
    if not nome or nome.strip().lower() in ("", "xxxxx", "x", "n/a", "-"):
        return "neutro"
    primeiro_nome = nome.strip().lower().split()[0]
    if primeiro_nome in NOMES_FEMININOS:
        return "feminino"
    return "masculino"


def ajustar_profissao_por_genero(profissao: str, genero: str) -> str:
    """Ajusta a profissão conforme o gênero da pessoa."""
    if not profissao:
        return profissao
    if genero == "feminino":
        conjugacoes = {
            "lavrador": "lavrador",  # neutra no Brasil
            "agricultor": "agricultora",
            "pecuarista": "pecuarista",
            "engenheiro": "engenheira",
            "tecnico": "técnica",
            "técnico": "técnica",
            "professor": "professora",
            "enfermeiro": "enfermeira",
            "empresario": "empresária",
            "aposentado": "aposentada",
            "desempregado": "desempregada",
            "trabalhador": "trabalhadora",
            "proprietario": "proprietária",
            "proprietário": "proprietária",
            "doutor": "doutora",
            "advogado": "advogada",
            "contador": "contadora",
            "autonomo": "autônoma",
            "autônomo": "autônoma",
        }
        for masc, fem in conjugacoes.items():
            if masc in profissao.lower():
                profissao = profissao.replace(masc.title(), fem.title())
                profissao = profissao.replace(masc.capitalize(), fem.title())
                profissao = profissao.replace(masc, fem)
                return profissao
    return profissao


def remover_duplicacoes(texto: str) -> str:
    """Remove palavras duplicadas consecutivas no texto.
    Ex: 'Comunhão Comunhão de Bens' -> 'Comunhão de Bens'
    """
    if not texto:
        return texto
    palavras = texto.split()
    resultado = [palavras[0]]
    for i in range(1, len(palavras)):
        if palavras[i].lower() != palavras[i - 1].lower():
            resultado.append(palavras[i])
    return " ".join(resultado)


def limpar_x_do_texto(texto: str) -> str:
    """Remove os 'X' usados como marcadores de gênero ou placeholder.
    Ex: 'Xlavradora' -> 'Lavrador'
    """
    if not texto:
        return texto
    # Remove 'X' no início de palavras que parecem ser marcadores de gênero
    texto = re.sub(r'^X([a-zA-Z])', r'\1', texto)
    texto = re.sub(r'\sX([A-Z][a-z]+)', r' \1', texto)
    return texto


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
           Se for feminino, use 'lavradora' e 'agricultora'. Se masculino, use 'lavrador' e 'agricultor'.
        3. Se não encontrar o cônjuge (esposa), preencha TODOS os campos do requerente_2 com 'XXXXXX'.
        4. Identifique a TRT (começa com BR e tem 11 números).
        5. Identifique a Comarca, Município e Matrícula.
        6. Extraia a área total retificada (encontrada na planta INCRA).
        7. IMPORTANTE: Se encontrar palavras duplicadas como 'comunhão comunhão', corrija para apenas 'comunhão'.
        8. IMPORTANTE: Remova qualquer 'X' usado como marcador de gênero (ex: 'Xlavradora' deve ser 'lavrador' ou 'lavradora' dependendo do gênero).
        9. Se encontrar 2 pessoas (casal), classifique como requerente_1 (proprietário) e requerente_2 (cônjuge/esposa).
        10. Se encontrar apenas 1 pessoa, classifique como requerente_1 e preencha requerente_2 com XXXXXX.
        11. Para o regime de bens, remova duplicações: 'comunhão comunhão de bens' deve virar 'comunhão de bens'.
        
        Retorne estritamente em JSON:
        {
            "comarca": "NOME DA COMARCA EM MAIUSCULO",
            "municipio_cliente": "Cidade do Cliente",
            "requerente_1": {
                "nome": "Nome Completo",
                "profissao": "Lavrador ou Lavradora (ajustado ao gênero)",
                "rg": "0.000.000",
                "orgao": "SSP/ES",
                "cpf": "000.000.000-00",
                "endereco_corrego": "Nome do Córrego (Apenas o nome)"
            },
            "requerente_2": {
                "nome": "Nome Completo da Esposa ou XXXXXX",
                "profissao": "Lavradora ou XXXXXX (ajustado ao gênero)",
                "rg": "0.000.000 ou XXXXXX",
                "orgao": "SSP/ES ou XXXXXX",
                "cpf": "000.000.000-00 ou XXXXXX",
                "regime_bens": "Comunhão Parcial de Bens ou XXXXXX (sem duplicações)"
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
                "area_total_retificada": "0,0000",
                "valor_fiscal": "R$ XXXXXX,00 (XXXXXX mil reais)",
                "area_estrada": "X.XXX,XX",
                "cfta_tecnico": "1119851971-1 ou encontrado",
                "codigo_credenciamento": "G1D ou encontrado"
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
            
            # Pós-processamento: ajustar gênero e profissão da requerente 2
            if dados.get("requerente_2"):
                req2 = dados["requerente_2"]
                nome_req2 = req2.get("nome", "")
                if nome_req2.lower() not in ("", "xxxxx", "x", "n/a", "-"):
                    genero = detectar_genero_por_nome(nome_req2)
                    req2["profissao"] = ajustar_profissao_por_genero(req2.get("profissao", ""), genero)
            
            # Ajustar gênero e profissão do requerente 1
            if dados.get("requerente_1"):
                req1 = dados["requerente_1"]
                genero1 = detectar_genero_por_nome(req1.get("nome", ""))
                req1["profissao"] = ajustar_profissao_por_genero(req1.get("profissao", ""), genero1)
            
            # Remover duplicações do regime de bens
            if dados.get("requerente_2", {}).get("regime_bens"):
                dados["requerente_2"]["regime_bens"] = remover_duplicacoes(dados["requerente_2"]["regime_bens"])
            
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
            meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            data_formatada = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

            # Extrair dados com valores padrão seguros
            req1 = dados.get("requerente_1", {})
            req2 = dados.get("requerente_2", {})
            imovel = dados.get("imovel", {})

            # Detecção de gênero da requerente 2 para ajustar pronome
            genero_req2 = detectar_genero_por_nome(req2.get("nome", ""))
            if genero_req2 == "feminino":
                pronome_conjuge = "esposa"
            elif genero_req2 == "masculino":
                pronome_conjuge = "esposo"
            else:
                pronome_conjuge = "esposa"  # padrão

            # Mapeamento de substituições
            substituicoes = {
                # Cabeçalho / Destinatário
                "COMARCA DE XXXXXXX – ES": f"COMARCA DE {dados.get('comarca', 'XXXXXX').upper()} – ES",

                # Requerente 1
                "XXXXXX, proprietário": f"{req1.get('nome', 'XXXXXX')}, proprietário",
                "XXXXX, lavrador": f"{req1.get('profissao', 'lavrador')}",
                "C.I. n°. XXXX – SSP/ES": f"C.I. n°. {req1.get('rg', 'XXXX')} – {req1.get('orgao', 'SSP/ES')}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {req1.get('cpf', 'XXXXXXX')}",

                # Requerente 2 (Esposa)
                "esposa XXXXXX": f"{pronome_conjuge} {req2.get('nome', 'XXXXXX')}",
                "XXXXX – SSP/ES": f"{req2.get('rg', 'XXXXX')} – {req2.get('orgao', 'SSP/ES')}",
                "CPF/MF n° XXXXXX": f"CPF/MF n° {req2.get('cpf', 'XXXXXX')}",

                # Regime de bens (com remoção de duplicações)
                "comunhão XXXXXX de bens": f"comunhão {req2.get('regime_bens', 'XXXXXX').lower()}",

                # Endereço
                "Córrego XXXXX": f"Córrego {req1.get('endereco_corrego', 'XXXXX')}",
                "Zona Rural, XXXXXX-ES": f"Zona Rural, {dados.get('municipio_cliente', 'XXXXXX')}-ES",

                # Imóvel
                "Sitio XXXXX": f"Sítio {imovel.get('nome', 'XXXXX')}",
                "registrada de XXXXXX ha": f"registrada de {imovel.get('area_registrada', 'XXXXXX')} ha",
                "município de XXXX - ES": f"município de {imovel.get('municipio_imovel', 'XXXX')} - ES",
                "comarca de XXXXXXX - ES": f"comarca de {imovel.get('comarca_imovel', 'XXXXXXX')} - ES",
                "matrícula n°. XXXXXX": f"matrícula n°. {imovel.get('matricula', 'XXXXXX')}",

                # Área encontrada / levantada
                "área de XXXXX ha": f"área de {imovel.get('area_encontrada', 'XXXXX')} ha",

                # INCRA
                "n°. XXX.XXX.XXX.XXX-X": f"n°. {imovel.get('codigo_incra', 'XXX.XXX.XXX.XXX-X')}",

                # Técnico / TRT / CFTA
                "TRT BRXXXXXXX": f"TRT {imovel.get('trt_numero', 'BRXXXXXXX')}",
                "CFTA n°. XXXXXXXXX-X": f"CFTA n°. {imovel.get('cfta_tecnico', 'XXXXXXX')}",
                "código XXX": f"código {imovel.get('codigo_credenciamento', 'XXX')}",

                # Área total retificada (item 10)
                "encontrada de XXXXXXX ha": f"encontrada de {imovel.get('area_total_retificada', 'XXXXXXX')} ha",

                # Valor fiscal (item 2) - mantido como placeholder se não extraído
                "R$ XXXXXX,00 (XXXXXX mil reais)": f"R$ {imovel.get('valor_fiscal', 'XXXXXX')}",

                # Área de estrada (item 7) - mantida como placeholder
                "X.XXX,XX m²": f"{imovel.get('area_estrada', 'X.XXX,XX')} m²",

                # Data
                "XX de XXX de XXXX": data_formatada,

                # Assinaturas
                "XXXXXXXXXXXXXXXX": req1.get('nome', 'XXXXXXXXXXXXXXXX'),
                "XXXXXXXXXXXXXXXXXX": req2.get('nome', 'XXXXXXXXXXXXXXXXXX'),
                "CPF: XXX.XXX.XXX-XX": f"CPF: {req1.get('cpf', 'XXX.XXX.XXX-XX')}",

                # CPF da esposa na assinatura
                "CPF: XXXXXX": f"CPF: {req2.get('cpf', 'XXXXXX')}",
            }

            # Aplicando substituições nos parágrafos
            for p in doc.paragraphs:
                for key, val in substituicoes.items():
                    if key in p.text:
                        p.text = p.text.replace(key, str(val))
            
            # Aplicando substituições nas tabelas
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
