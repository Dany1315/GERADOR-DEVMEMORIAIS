import io
import json
import logging
import os
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Callable
from docx import Document
from docx.shared import Pt, Inches
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
    "yanira", "zelia", "ana clara", "vera lucia", "maria josé"
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


class GeradorRequerimentoCartorio:
    """
    Gerador de requerimentos de cartório seguindo o modelo CORRETO.
    
    VERSÃO FINAL CORRIGIDA (v8):
    - Segue estrutura exata do modelo correto
    - Dados separados em parágrafos distintos
    - Nomes corretos nas assinantes
    - Todos os placeholders substituídos
    - Formatação correta preservada
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
        3. Se não encontrar o cônjuge (esposa), preencha TODOS os campos do requerente_2 com 'XXXXXX'.
        4. Identifique a TRT (começa com BR e tem 11 números).
        5. Identifique a Comarca, Município e Matrícula.
        6. Extraia a área total retificada (encontrada na planta INCRA).
        7. IMPORTANTE: Remova duplicações de palavras (ex: 'comunhão comunhão' -> 'comunhão').
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
                "estado_civil": "casado ou solteiro",
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
            
            self._atualizar_progresso(1, "Dados extraídos com sucesso!", 75)
            logger.info("✅ Dados extraídos com sucesso da IA")
            return dados
        except Exception as e:
            logger.error(f"Erro na extração via Gemini: {e}")
            raise Exception(f"Falha ao extrair dados: {str(e)}")

    def gerar_documento(self, dados: Dict[str, Any], template_name: str) -> io.BytesIO:
        """
        Gera o documento Word seguindo EXATAMENTE o modelo correto.
        
        VERSÃO CORRIGIDA: Estrutura idêntica ao modelo enviado.
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
            data_formatada = f"{hoje.day:02d} de {meses[hoje.month-1]} de {hoje.year}"

            # Extrair dados com valores padrão seguros
            req1 = dados.get("requerente_1", {})
            req2 = dados.get("requerente_2", {})
            imovel = dados.get("imovel", {})

            self._atualizar_progresso(2, "Preenchendo documento...", 50)
            
            # ============================================================
            # SUBSTITUIÇÃO SEGUINDO O MODELO CORRETO
            # ============================================================
            
            # Parágrafo 0: Cabeçalho (já está no template)
            # Parágrafo 3: Dados do requerente (PRINCIPAL)
            self._substituir_paragrafo_exato(doc, 3, 
                f"{req1.get('nome', 'XXXXXX')}, proprietário, brasileiro, {req1.get('estado_civil', 'casado')}, "
                f"{req1.get('profissao', 'lavrador')}, C.I. n° {req1.get('rg', 'XXXXXX')} {req1.get('orgao', 'SSP/ES')}, "
                f"CPF/MF n°. {req1.get('cpf', 'XXXXXX')}, e sua esposa "
                f"{req2.get('nome', 'XXXXXX')}, {req2.get('profissao', 'XXXXXX')}, "
                f"C.I. n° {req2.get('rg', 'XXXXXX')} {req2.get('orgao', 'XXXXXX')}, "
                f"CPF/MF n°. {req2.get('cpf', 'XXXXXX')}, brasileiros, casados sob o regime de "
                f"{req2.get('regime_bens', 'XXXXXX')} de bens, residentes e domiciliados no Córrego "
                f"{req1.get('endereco_corrego', 'XXXXXX')}, Zona Rural, "
                f"{imovel.get('municipio_imovel', 'XXXXXX')}-ES; E o responsável técnico pela medição "
                f"Régis Campo da Silva, brasileiro, casado, técnico em agropecuária, C.I. n°. 1.936.653 – SPTC/ES, "
                f"CPF/MF n°. 111.985.197-11, residente e domiciliado no Córrego Groner, Zona Rural, Vila Valério-ES, "
                f"vem expor e requerer o que segue:"
            )
            
            # Parágrafo 5: Descrição do imóvel
            self._substituir_paragrafo_exato(doc, 5,
                f"Que são senhores e legítimos proprietários de uma área de terras denominada "
                f"\"{imovel.get('nome', 'XXXXXX')}\", com área registrada de {imovel.get('area_registrada', 'XXXXXX')} ha "
                f"situada no município de {imovel.get('municipio_imovel', 'XXXXXX')}-ES e registrada na comarca de "
                f"{imovel.get('comarca_imovel', 'XXXXXX')} – ES, a qual se acha devidamente registrada, descrita e "
                f"caracterizada na matrícula n°. {imovel.get('matricula', 'XXXXXX')}, dessa circunscrição imobiliária."
            )
            
            # Parágrafo 7: Valor fiscal
            self._substituir_paragrafo_exato(doc, 7,
                f"Que o imóvel acima mencionado está avaliado pelos proprietários para fins fiscais no valor de "
                f"R$ {imovel.get('valor_fiscal', 'XXXXXX')}, conforme item 8 das Notas, da Tabela 11 de Emolumentos "
                f"editada pela CGJ/ES, bem como o artigo 98, do Código de Normas da Corregedoria Geral da Justiça deste "
                f"Estado do ES;"
            )
            
            # Parágrafo 9: Levantamento perimetral
            self._substituir_paragrafo_exato(doc, 9,
                f"Que foi procedido o levantamento perimetral do imóvel, sendo encontrado a área de "
                f"{imovel.get('area_encontrada', 'XXXXXX')} ha;"
            )
            
            # Parágrafo 11: Certificação INCRA
            self._substituir_paragrafo_exato(doc, 11,
                f"Que referido levantamento foi certificado pelo Instituto Nacional de Colonização e Reforma Agrária – INCRA, "
                f"sob o n°. {imovel.get('codigo_incra', 'XXXXXX')}."
            )
            
            # Parágrafo 14: Técnico
            self._substituir_paragrafo_exato(doc, 14,
                f"Que os trabalhos topográficos foram elaborados pelo técnico em agropecuária, Regis Campo da Silva, "
                f"CFTA n°. {imovel.get('cfta_tecnico', 'XXXXXX')}, credenciamento no INCRA sob o código "
                f"{imovel.get('codigo_credenciamento', 'XXXXXX')} da TRT {imovel.get('trt_numero', 'XXXXXX')}."
            )
            
            # Parágrafo 16: Glebas
            self._substituir_paragrafo_exato(doc, 16,
                f"Que se trata de um imóvel dividido por uma estrada municipal formando duas glebas distintas, autônomas e "
                f"independentes, sendo elas: Uma gleba denominada \"Gleba 1\" com área de {imovel.get('area_registrada', 'XXXXXX')} ha; "
                f"uma gleba denominada \"Gleba 2\" com área de {imovel.get('area_encontrada', 'XXXXXX')} ha."
            )
            
            # Parágrafo 18: Estrada
            self._substituir_paragrafo_exato(doc, 18,
                f"Que foi encontrada uma área de Estrada Municipal de {imovel.get('area_estrada', 'XXXXXX')} m². "
                f"Sendo assim requerida a averbação de afetação por finalidade pública."
            )
            
            # Parágrafo 32: Data
            self._substituir_paragrafo_exato(doc, 32,
                f"Vila Valério – ES, {data_formatada}."
            )
            
            # Parágrafo 38: Nomes dos assinantes (CORRETO!)
            self._substituir_paragrafo_exato(doc, 38,
                f"                       {req1.get('nome', 'XXXXXX')}                    "
                f"{req2.get('nome', 'XXXXXX')}"
            )
            
            # Parágrafo 39: CPFs dos assinantes (CORRETO!)
            self._substituir_paragrafo_exato(doc, 39,
                f"                   CPF:{req1.get('cpf', 'XXXXXX')}                    "
                f"CPF: {req2.get('cpf', 'XXXXXX')}"
            )

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

    def _substituir_paragrafo_exato(self, doc, indice_paragrafo: int, novo_texto: str):
        """Substitui o conteúdo exato de um parágrafo."""
        if indice_paragrafo < len(doc.paragraphs):
            paragrafo = doc.paragraphs[indice_paragrafo]
            # Limpar o parágrafo
            for run in paragrafo.runs:
                run.text = ""
            # Adicionar novo texto
            paragrafo.text = novo_texto

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
