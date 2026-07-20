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
            "lavrador": "lavradora",
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
    """Remove palavras duplicadas consecutivas no texto."""
    if not texto:
        return texto
    palavras = texto.split()
    resultado = [palavras[0]]
    for i in range(1, len(palavras)):
        if palavras[i] != palavras[i-1]:
            resultado.append(palavras[i])
    return " ".join(resultado)


class GeradorRequerimentoCartorio:
    def __init__(self, model_name: str, callback_progresso=None):
        self.model_name = model_name
        self.callback_progresso = callback_progresso
        self._atualizar_progresso(0, "Inicializando gerador...", 0)

    def _atualizar_progresso(self, etapa: int, descricao: str, percentual: float = None):
        """Callback para atualizar progresso durante o processamento."""
        if self.callback_progresso:
            self.callback_progresso(etapa, descricao, percentual)

    def extrair_dados_documentos(self, imagens: List) -> Dict[str, Any]:
        """Extrai dados dos documentos usando Gemini."""
        try:
            self._atualizar_progresso(1, "Analisando documentos com IA Gemini...", 25)
            
            model = genai.GenerativeModel(self.model_name)
            
            prompt = """
            Analise os documentos fornecidos (RG, CPF, Certidões, Matrículas, etc.) e extraia os seguintes dados:
            
            REQUERENTE 1 (Proprietário):
            - nome: Nome completo
            - profissao: Profissão (ex: agricultor, lavrador)
            - rg: Número do RG
            - cpf: CPF (apenas números)
            - endereco_corrego: Nome do córrego/localidade
            
            REQUERENTE 2 (Esposa/Esposo):
            - nome: Nome completo
            - profissao: Profissão
            - rg: Número do RG
            - cpf: CPF (apenas números)
            - regime_bens: Regime de bens (ex: comunhão parcial de bens)
            
            IMÓVEL:
            - nome: Nome do sítio/propriedade
            - area_registrada: Área registrada em hectares (ex: 25,51 ha)
            - area_encontrada: Área encontrada em hectares
            - area_total_retificada: Área total retificada em hectares
            - municipio_imovel: Município do imóvel
            - comarca_imovel: Comarca do imóvel
            - matricula: Número da matrícula
            - codigo_incra: Código INCRA (formato: BR XXXXXXX)
            - trt_numero: Número do TRT
            - valor_fiscal: Valor fiscal (ex: 150000)
            
            OUTROS:
            - comarca: Comarca (ex: SÃO GABRIEL DA PALHA)
            - municipio_cliente: Município do cliente
            - data: Data (formato: DD/MM/YYYY)
            
            Retorne APENAS um JSON válido com esses campos.
            """
            
            response = model.generate_content([prompt] + imagens)
            
            self._atualizar_progresso(2, "Processando resposta da IA...", 50)
            
            # Extrair JSON da resposta
            texto_resposta = response.text
            json_match = re.search(r'\{.*\}', texto_resposta, re.DOTALL)
            
            if json_match:
                dados = json.loads(json_match.group())
            else:
                dados = {}
            
            self._atualizar_progresso(3, "Validando dados extraídos...", 75)
            
            # Estruturar dados com valores padrão
            dados_estruturados = {
                "requerente_1": {
                    "nome": dados.get("nome", "XXXXXX"),
                    "profissao": dados.get("profissao", "XXXXX"),
                    "rg": dados.get("rg", "XXXX"),
                    "cpf": dados.get("cpf", "XXXXXXX"),
                    "endereco_corrego": dados.get("endereco_corrego", "XXXXX"),
                },
                "requerente_2": {
                    "nome": dados.get("nome_esposa", "XXXXXX"),
                    "profissao": dados.get("profissao_esposa", "XXXXX"),
                    "rg": dados.get("rg_esposa", "XXXXX"),
                    "cpf": dados.get("cpf_esposa", "XXXXXX"),
                    "regime_bens": dados.get("regime_bens", "XXXXXX"),
                },
                "imovel": {
                    "nome": dados.get("nome_imovel", "XXXXX"),
                    "area_registrada": dados.get("area_registrada", "XXXXXX"),
                    "area_encontrada": dados.get("area_encontrada", "XXXXX"),
                    "area_total_retificada": dados.get("area_total_retificada", "XXXXXXX"),
                    "municipio_imovel": dados.get("municipio_imovel", "XXXX"),
                    "comarca_imovel": dados.get("comarca_imovel", "XXXXXXX"),
                    "matricula": dados.get("matricula", "XXXXXX"),
                    "codigo_incra": dados.get("codigo_incra", "XXX.XXX.XXX.XXX-X"),
                    "trt_numero": dados.get("trt_numero", "BRXXXXXXX"),
                    "valor_fiscal": dados.get("valor_fiscal", "XX0.000,00"),
                },
                "comarca": dados.get("comarca", "XXXXXXX"),
                "municipio_cliente": dados.get("municipio_cliente", "XXXXXX"),
                "data": dados.get("data", "XX de XXX de XXXX"),
            }
            
            self._atualizar_progresso(4, "Extração concluída!", 100)
            return dados_estruturados
            
        except Exception as e:
            logger.error(f"Erro na extração: {e}")
            raise Exception(f"Falha ao extrair dados: {str(e)}")

    def _substituir_paragrafo_exato(self, doc, indice_paragrafo: int, novo_texto: str, alinhamento=None):
        """
        Substitui o conteúdo EXATO de um parágrafo específico.
        Preserva a formatação e o alinhamento.
        """
        if indice_paragrafo < len(doc.paragraphs):
            para = doc.paragraphs[indice_paragrafo]
            
            # Limpar todos os runs existentes
            for run in para.runs:
                run.text = ""
            
            # Adicionar novo texto
            if para.runs:
                para.runs[0].text = novo_texto
            else:
                para.add_run(novo_texto)
            
            # Aplicar alinhamento se especificado
            if alinhamento:
                para.alignment = alinhamento

    def gerar_documento(self, dados: Dict[str, Any], template_name: str) -> io.BytesIO:
        """Gera o documento Word preenchido com os dados extraídos."""
        try:
            # Buscar template
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, template_name)
            if not os.path.exists(template_path):
                template_path = template_name
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Modelo {template_name} não encontrado.")

            doc = Document(template_path)
            
            # Data formatada
            hoje = datetime.now()
            meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            data_formatada = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

            # Extrair dados
            req1 = dados.get("requerente_1", {})
            req2 = dados.get("requerente_2", {})
            imovel = dados.get("imovel", {})

            # ============================================================
            # USAR PLACEHOLDERS ÚNICOS - CADA CAMPO TEM UM PLACEHOLDER DIFERENTE
            # ============================================================
            substituicoes = {
                # Cabeçalho
                "COMARCA DE (XXXXXXX)": f"COMARCA DE {dados.get('comarca', 'XXXXXXX')}",

                # Requerente 1 - PLACEHOLDERS ÚNICOS
                "XXXXXX, proprietário": f"{req1.get('nome', 'XXXXXX')}, proprietário",
                "XXXXX lavrador": f"{req1.get('profissao', 'XXXXX')} lavrador" if 'lavrador' in req1.get('profissao', '').lower() else f"{req1.get('profissao', 'XXXXX')}",
                "C.I. n°. (NUMERO DA IDENTIDADE)": f"C.I. n°. {req1.get('rg', 'NUMERO DA IDENTIDADE')}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {req1.get('cpf', 'XXXXXXX')}",

                # Requerente 2 - PLACEHOLDERS ÚNICOS
                "e sua esposa (XXXXXX)": f"e sua esposa {req2.get('nome', 'XXXXXX')}",
                "C.I. n° (XXXXX)": f"C.I. n° {req2.get('rg', 'XXXXX')}",
                "CPF/MF n°. (XXXXXX)": f"CPF/MF n°. {req2.get('cpf', 'XXXXXX')}",
                "regime de (XXXXXX)": f"regime de {req2.get('regime_bens', 'XXXXXX')}",

                # Endereço
                "córrego (XXXXX)": f"córrego {req1.get('endereco_corrego', 'XXXXX')}",

                # Imóvel - PLACEHOLDERS ÚNICOS
                "Sitio (XXXXX)": f"Sítio {imovel.get('nome', 'XXXXX')}",
                "área de (XXXXXX há)": f"área de {imovel.get('area_registrada', 'XXXXXX')} ha",
                "(XXXXXX-ES)": f"{dados.get('municipio_cliente', 'XXXXXX')}-ES",
                "(XXXXXXX – ES)": f"{imovel.get('comarca_imovel', 'XXXXXXX')} – ES",
                "matrícula (XXXXXX)": f"matrícula {imovel.get('matricula', 'XXXXXX')}",
                "área de (XXXXX há)": f"área de {imovel.get('area_encontrada', 'XXXXX')} ha",
                "Código INCRA (XXX.XXX.XXX.XXX-X)": f"Código INCRA {imovel.get('codigo_incra', 'XXX.XXX.XXX.XXX-X')}",
                "TRT (BRXXXXXXX)": f"TRT {imovel.get('trt_numero', 'BRXXXXXXX')}",
                "área de Estrada Municipal de (XXXXXX) m²": f"área de Estrada Municipal de {imovel.get('area_estrada', '0,00')} m²",
                "R$ XX0.000,00": f"R$ {imovel.get('valor_fiscal', '0,00')}",

                # Data
                "(XX de XXX de XXXX)": data_formatada,

                # Assinantes - PLACEHOLDERS ÚNICOS
                "(XXXXXX)": req1.get('nome', 'XXXXXX'),
                "(XXXXXXXX)": req2.get('nome', 'XXXXXXXX'),
                "CPF: (XXXXXX)": f"CPF: {req1.get('cpf', 'XXXXXX')}",
                "CPF: (XXXXXXXX)": f"CPF: {req2.get('cpf', 'XXXXXXXX')}",
            }

            # Aplicar substituições
            for p in doc.paragraphs:
                for key, val in substituicoes.items():
                    if key in p.text:
                        p.text = p.text.replace(key, str(val))
            
            # Aplicar em tabelas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, val in substituicoes.items():
                                if key in p.text:
                                    p.text = p.text.replace(key, str(val))

            # Salvar em BytesIO
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            return output.getvalue()

        except Exception as e:
            logger.error(f"Erro ao gerar documento: {e}")
            raise Exception(f"Falha ao gerar documento: {str(e)}")
