import io
import json
import logging
import os
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Callable
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
import time

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
            "lavrador": "lavrador",
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
        if palavras[i].lower() != palavras[i - 1].lower():
            resultado.append(palavras[i])
    return " ".join(resultado)


class GeradorRequerimentoCartorio:
    """
    Gerador de requerimentos de cartório com suporte a barra de progresso.
    
    VERSÃO COM PROGRESSO (v4):
    - Usa placeholders contextuais e únicos
    - Evita conflitos de substituição
    - Implementa barra de progresso com tempo estimado
    - Fornece logging detalhado
    - Compatível com a estrutura do projeto GitHub
    """
    
    def __init__(self, model_name: str, callback_progresso: Optional[Callable] = None):
        """
        Inicializa o gerador.
        
        Args:
            model_name: Nome do modelo Gemini
            callback_progresso: Função callback para atualizar progresso (opcional)
        """
        self.model_name = model_name
        self.model = genai.GenerativeModel(model_name)
        self.callback_progresso = callback_progresso
        self.tempo_inicio = None
        
    def _atualizar_progresso(self, etapa: int, descricao: str, percentual: float = None):
        """Atualiza o progresso via callback."""
        if self.callback_progresso:
            try:
                self.callback_progresso(etapa, descricao, percentual)
            except Exception as e:
                logger.warning(f"Erro ao atualizar progresso: {e}")

    def extrair_dados_documentos(self, imagens: List[Any]) -> Dict[str, Any]:
        """
        Envia as imagens dos documentos para o Gemini e extrai os dados estruturados.
        """
        self._atualizar_progresso(1, "Analisando documentos com IA Gemini...", 25)
        
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
        8. Se encontrar 2 pessoas (casal), classifique como requerente_1 (proprietário) e requerente_2 (cônjuge/esposa).
        9. Se encontrar apenas 1 pessoa, classifique como requerente_1 e preencha requerente_2 com XXXXXX.
        10. Para o regime de bens, remova duplicações: 'comunhão comunhão de bens' deve virar 'comunhão de bens'.
        
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
                "cfta_tecnico": "1119851971-1 ou encontrado",
                "codigo_credenciamento": "G1D ou encontrado",
                "valor_fiscal": "0,00",
                "area_estrada": "0,00"
            }
        }
        """
        
        try:
            conteudo = [prompt] + imagens
            response = self.model.generate_content(conteudo)
            
            self._atualizar_progresso(1, "Processando resposta da IA...", 50)
            
            text_response = response.text
            if "```json" in text_response:
                text_response = text_response.split("```json")[1].split("```")[0]
            elif "```" in text_response:
                text_response = text_response.split("```")[1].split("```")[0]
            
            dados = json.loads(text_response.strip())
            
            # Pós-processamento: ajustar gênero e profissão
            if dados.get("requerente_2"):
                req2 = dados["requerente_2"]
                nome_req2 = req2.get("nome", "")
                if nome_req2.lower() not in ("", "xxxxx", "x", "n/a", "-"):
                    genero = detectar_genero_por_nome(nome_req2)
                    req2["profissao"] = ajustar_profissao_por_genero(req2.get("profissao", ""), genero)
            
            if dados.get("requerente_1"):
                req1 = dados["requerente_1"]
                genero1 = detectar_genero_por_nome(req1.get("nome", ""))
                req1["profissao"] = ajustar_profissao_por_genero(req1.get("profissao", ""), genero1)
            
            # Remover duplicações do regime de bens
            if dados.get("requerente_2", {}).get("regime_bens"):
                dados["requerente_2"]["regime_bens"] = remover_duplicacoes(dados["requerente_2"]["regime_bens"])
            
            self._atualizar_progresso(1, "Dados extraídos com sucesso!", 75)
            logger.info("✅ Dados extraídos com sucesso da IA")
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

    def _substituir_em_run(self, text_element, substituicoes: Dict[str, str]):
        """
        Substitui placeholders em um elemento de texto (parágrafo ou cell).
        Trabalha no nível dos RUNS para preservar formatação.
        Realiza substituições com ordem específica para evitar conflitos.
        """
        # Junta todo o texto do elemento
        texto_completo = ""
        for run in text_element.runs:
            texto_completo += run.text
        
        # Verifica se há algum placeholder para substituir
        tem_placeholder = False
        for key in substituicoes:
            if key in texto_completo:
                tem_placeholder = True
                break
        
        if not tem_placeholder:
            return
        
        # Aplica substituições em ORDEM ESPECÍFICA (do mais específico para o mais genérico)
        ordem_substituicao = sorted(substituicoes.keys(), key=len, reverse=True)
        
        for key in ordem_substituicao:
            val = substituicoes[key]
            if key in texto_completo:
                texto_completo = texto_completo.replace(key, str(val), 1)
        
        # Reescreve o texto: primeiro run recebe tudo, demais ficam vazios
        if text_element.runs:
            text_element.runs[0].text = texto_completo
            for run in text_element.runs[1:]:
                run.text = ""

    def gerar_documento(self, dados: Dict[str, Any], template_name: str) -> io.BytesIO:
        """
        Gera o documento Word preenchendo os placeholders com os dados extraídos.
        
        VERSÃO COM PROGRESSO: Exibe barra de progresso durante a geração.
        """
        try:
            self._atualizar_progresso(2, "Carregando template de requerimento...", 10)
            
            # Busca do template
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, template_name)
            if not os.path.exists(template_path):
                template_path = template_name
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Modelo {template_name} não encontrado.")

            doc = Document(template_path)
            
            self._atualizar_progresso(2, "Preparando dados para preenchimento...", 30)
            
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
                pronome_conjuge = "esposa"

            # ============================================================
            # MAPEAMENTO DE SUBSTITUIÇÕES COM CONTEXTO ÚNICO
            # ============================================================
            substituicoes = {
                "COMARCA DE XXXXXXX – ES": f"COMARCA DE {dados.get('comarca', 'XXXXXX').upper()} – ES",
                "XXXXXX, proprietário": f"{req1.get('nome', 'XXXXXX')}, proprietário",
                "XXXXX, lavrador": f"{req1.get('profissao', 'lavrador')}",
                "C.I. n°. XXXX – SSP/ES": f"C.I. n°. {req1.get('rg', 'XXXX')} – {req1.get('orgao', 'SSP/ES')}",
                "CPF/MF n°. XXXXXXX": f"CPF/MF n°. {req1.get('cpf', 'XXXXXXX')}",
                "esposa XXXXXX": f"{pronome_conjuge} {req2.get('nome', 'XXXXXX')}",
                "XXXXX – SSP/ES": f"{req2.get('rg', 'XXXXX')} – {req2.get('orgao', 'SSP/ES')}",
                "CPF/MF n° XXXXXX": f"CPF/MF n° {req2.get('cpf', 'XXXXXX')}",
                "comunhão XXXXXX de bens": f"comunhão {req2.get('regime_bens', 'XXXXXX').lower()} de bens",
                "Córrego XXXXX": f"Córrego {req1.get('endereco_corrego', 'XXXXX')}",
                "Zona Rural, XXXXXX-ES": f"Zona Rural, {dados.get('municipio_cliente', 'XXXXXX')}-ES",
                "Sitio XXXXX": f"Sítio {imovel.get('nome', 'XXXXX')}",
                "registrada de XXXXXX ha": f"registrada de {imovel.get('area_registrada', 'XXXXXX')} ha",
                "município de XXXX - ES": f"município de {imovel.get('municipio_imovel', 'XXXX')} - ES",
                "comarca de XXXXXXX - ES": f"comarca de {imovel.get('comarca_imovel', 'XXXXXXX')} - ES",
                "matrícula n°. XXXXXX": f"matrícula n°. {imovel.get('matricula', 'XXXXXX')}",
                "área de XXXXX ha": f"área de {imovel.get('area_encontrada', 'XXXXX')} ha",
                "n°. XXX.XXX.XXX.XXX-X": f"n°. {imovel.get('codigo_incra', 'XXX.XXX.XXX.XXX-X')}",
                "TRT BRXXXXXXX": f"TRT {imovel.get('trt_numero', 'BRXXXXXXX')}",
                "CFTA n°. XXXXXXXXX-X": f"CFTA n°. {imovel.get('cfta_tecnico', 'XXXXXXX')}",
                "código XXX": f"código {imovel.get('codigo_credenciamento', 'XXX')}",
                "encontrada de XXXXXXX ha": f"encontrada de {imovel.get('area_total_retificada', 'XXXXXXX')} ha",
                "R$ XXXXXX,00 (XXXXXX mil reais)": f"R$ {imovel.get('valor_fiscal', 'XXXXXX')},00",
                "X.XXX,XX m²": f"{imovel.get('area_estrada', 'X.XXX,XX')} m²",
                "XX de XXX de XXXX": data_formatada,
                "XXXXXXXXXXXXXXXX": req1.get('nome', 'XXXXXXXXXXXXXXXX'),
                "XXXXXXXXXXXXXXXXXX": req2.get('nome', 'XXXXXXXXXXXXXXXXXX'),
                "CPF: XXX.XXX.XXX-XX": f"CPF: {req1.get('cpf', 'XXX.XXX.XXX-XX')}",
                "CPF: XXXXXX": f"CPF: {req2.get('cpf', 'XXXXXX')}",
            }

            self._atualizar_progresso(2, "Preenchendo parágrafos...", 50)
            
            # Aplicando substituições nos parágrafos
            for i, p in enumerate(doc.paragraphs):
                self._substituir_em_run(p, substituicoes)
                # Atualizar progresso a cada 10 parágrafos
                if i % 10 == 0:
                    self._atualizar_progresso(2, f"Preenchendo parágrafos ({i}/{len(doc.paragraphs)})...", 50)
            
            self._atualizar_progresso(2, "Preenchendo tabelas...", 70)
            
            # Aplicando substituições nas tabelas
            for table_idx, table in enumerate(doc.tables):
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            self._substituir_em_run(p, substituicoes)
                # Atualizar progresso
                if table_idx % 5 == 0:
                    self._atualizar_progresso(2, f"Preenchendo tabelas ({table_idx}/{len(doc.tables)})...", 70)
            
            self._atualizar_progresso(2, "Ajustando formatação...", 85)
            
            # Ajuste final de fonte
            self._ajustar_fonte_arial(doc)

            self._atualizar_progresso(2, "Salvando documento...", 95)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            self._atualizar_progresso(2, "Documento gerado com sucesso!", 100)
            logger.info("✅ Documento gerado com sucesso")
            return buffer
            
        except Exception as e:
            logger.error(f"Erro ao gerar documento: {e}")
            raise e
