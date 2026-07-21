"""
Classe principal de processamento de memoriais descritivos.
Encapsula toda a lógica de negócio, separada da interface Streamlit.
VERSÃO CORRIGIDA: Campos mapeados corretamente, coordenadas do próximo vértice, matrícula separada
"""

import io
import re
import json
import logging
import time
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime

import fitz
from PIL import Image
from pypdf import PdfReader
from pydantic import BaseModel, ValidationError
import google.generativeai as genai
from google.generativeai import types

from config import GEMINI_CONFIG, PROCESSAMENTO_CONFIG
from utils import retry_com_backoff, criar_logger

logger = criar_logger(__name__)

# ==========================================
# MODELOS PYDANTIC
# ==========================================

class RegraConfrontante(BaseModel):
    ponto_inicio: int
    ponto_fim: int
    nome_confrontante: str

    class Config:
        str_strip_whitespace = True


class MapeamentoConfrontantes(BaseModel):
    regras: List[RegraConfrontante]

    class Config:
        str_strip_whitespace = True


class SegmentoRoteiro(BaseModel):
    de: str
    para: str
    n_y: str
    e_x: str
    azimute: str
    distancia: str

    class Config:
        str_strip_whitespace = True


class ExtracaoRoteiro(BaseModel):
    segmentos: List[SegmentoRoteiro]

    class Config:
        str_strip_whitespace = True


# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================

def extrair_matricula_e_nome(texto_confrontante: str) -> Tuple[str, str]:
    """
    Extrai matrícula e nome do confrontante.
    
    Entrada: "MATRÍCULA 9448 JOSÉ ANTONIO RIBEIRO"
    Saída: ("9448", "JOSÉ ANTONIO RIBEIRO")
    
    Entrada: "ESTRADA MUNICIPAL"
    Saída: ("N/A", "ESTRADA MUNICIPAL")
    
    Args:
        texto_confrontante: Texto com matrícula e nome
        
    Returns:
        Tuple[str, str]: (matricula, nome)
    """
    if not texto_confrontante:
        return "N/A", "N/A"
    
    texto = texto_confrontante.strip()
    
    # Se começa com "MATRÍCULA"
    if texto.startswith("MATRÍCULA"):
        partes = texto.split()
        if len(partes) >= 2:
            matricula = partes[1]
            nome = ' '.join(partes[2:]) if len(partes) > 2 else "N/A"
            return matricula, nome
    
    # Se é apenas "ESTRADA MUNICIPAL" ou similar (sem matrícula)
    return "N/A", texto


def limpar_unidade(valor: str) -> str:
    """
    Remove unidades (m, m², etc) do valor.
    
    Entrada: "207,12 m"
    Saída: "207,12"
    
    Args:
        valor: Valor com possível unidade
        
    Returns:
        str: Valor sem unidade
    """
    if not valor:
        return ""
    
    return valor.strip().replace(" m", "").replace(" m²", "").strip()


# ==========================================
# CLASSE PROCESSADOR
# ==========================================

class ProcessadorMemorial:
    """
    Orquestra todo o processamento de memoriais descritivos.
    
    Responsabilidades:
    - Extração de roteiros de PDFs
    - Mapeamento de confrontantes com IA
    - Vinculação de dados
    - Validação e tratamento de erros
    """

    def __init__(self, nome_modelo: str):
        """
        Inicializa o processador.
        
        Args:
            nome_modelo: Nome do modelo Gemini a usar
        """
        self.nome_modelo = nome_modelo
        self.segmentos: List[Dict] = []
        self.mapeamento: Optional[MapeamentoConfrontantes] = None
        self.tempo_inicio = time.time()
        self.tempo_gemini = 0.0
        logger.info(f"ProcessadorMemorial inicializado com modelo: {nome_modelo}")

    # ==========================================
    # EXTRAÇÃO DE PDFs
    # ==========================================

    def pdf_para_imagens(
        self,
        arquivo_pdf,
        dpi: int = 250,
        progress_callback=None
    ) -> List[Image.Image]:
        """
        Converte páginas de PDF em imagens PIL usando PyMuPDF.
        
        Não depende de Poppler/Tesseract (funciona 100% em nuvem).
        Suporta PDFs de CAD sem texto extraível (ex: VectorDraw).
        
        Args:
            arquivo_pdf: Arquivo do Streamlit
            dpi: Qualidade da imagem (150-400)
            progress_callback: Função para atualizar barra de progresso
        
        Returns:
            Lista de imagens PIL
            
        Raises:
            ValueError: Se PDF estiver corrompido ou vazio
        """
        try:
            arquivo_pdf.seek(0)
            dados = arquivo_pdf.read()
            
            try:
                documento = fitz.open(stream=dados, filetype="pdf")
            except Exception as e:
                logger.error(f"PDF parece corrompido: {e}")
                raise ValueError(
                    "PDF está corrompido ou em formato não suportado. "
                    "Tente salvar novamente em um leitor de PDF."
                )
            
            if len(documento) == 0:
                raise ValueError("PDF está vazio (sem páginas)")
            
            zoom = dpi / 72
            matriz = fitz.Matrix(zoom, zoom)
            
            imagens = []
            total_paginas = len(documento)
            
            logger.info(f"Convertendo PDF em {total_paginas} página(s) a {dpi} DPI")
            
            for idx, pagina in enumerate(documento):
                pix = pagina.get_pixmap(matrix=matriz)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                imagens.append(img)
                
                if progress_callback:
                    progress = (idx + 1) / total_paginas
                    progress_callback(progress, f"Página {idx + 1}/{total_paginas}")
            
            documento.close()
            logger.info(f"✅ {len(imagens)} página(s) convertida(s) com sucesso")
            return imagens
            
        except ValueError:
            raise
        except Exception as e:
            logger.error(f"Erro ao rasterizar PDF: {str(e)}")
            raise ValueError(f"Erro ao processar PDF: {str(e)}")

    def extrair_texto_pdf(self, arquivo_pdf) -> str:
        """Extrai texto de PDF com camada de texto (fallback)."""
        try:
            arquivo_pdf.seek(0)
            leitor = PdfReader(arquivo_pdf)
            texto_completo = ""
            
            logger.info(f"Extraindo texto de PDF com {len(leitor.pages)} páginas")
            
            for idx, pagina in enumerate(leitor.pages):
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
                else:
                    logger.warning(f"Página {idx + 1} sem texto extraível")
            
            if len(texto_completo.strip()) < 100:
                logger.warning("PDF não contém texto suficiente")
            
            return texto_completo
            
        except Exception as e:
            logger.error(f"Erro ao extrair texto: {str(e)}")
            raise

    # ==========================================
    # EXTRAÇÃO COM IA
    # ==========================================

    @retry_com_backoff(
        max_retries=GEMINI_CONFIG.MAX_RETRIES,
        delay_inicial=GEMINI_CONFIG.DELAY_RETRY_INICIAL
    )
    def extrair_roteiro_com_ia(
        self,
        imagens_roteiro: List[Image.Image]
    ) -> List[Dict[str, str]]:
        """
        Extrai tabela de roteiro usando visão do Gemini.
        
        Com retry automático para rate limiting.
        
        Args:
            imagens_roteiro: Lista de imagens da tabela de roteiro
        
        Returns:
            Lista de segmentos extraídos
            
        Raises:
            ValueError: Se resposta da IA for inválida
        """
        try:
            prompt = """
            Você é um especialista em leitura de plantas topográficas. A(s) imagem(ns)
            contém uma TABELA DE ROTEIRO PERIMÉTRICO com vértices, coordenadas, azimutes
            e distâncias.

            Leia a tabela LINHA POR LINHA, na ordem exata, extraindo:
            - de: número do vértice de origem
            - para: número do vértice de destino
            - n_y: coordenada N (incluindo unidade "m")
            - e_x: coordenada E (incluindo unidade "m")
            - azimute: em formato graus/minutos/segundos (ex: 45°12'33")
            - distancia: incluindo unidade "m"

            Transcreva EXATAMENTE os valores. Não invente dados.
            """

            tempo_inicio_gemini = time.time()
            logger.info("Chamando Gemini para extrair tabela de roteiro...")

            model = genai.GenerativeModel(
                model_name=self.nome_modelo,
                generation_config=types.GenerationConfig(
                    response_mime_type="application/json",
                    response_schema=ExtracaoRoteiro,
                    temperature=GEMINI_CONFIG.TEMPERATURE,
                ),
            )

            conteudo = [prompt] + list(imagens_roteiro)
            response = model.generate_content(conteudo)
            
            tempo_gemini = time.time() - tempo_inicio_gemini
            self.tempo_gemini += tempo_gemini
            logger.info(f"Gemini respondeu em {tempo_gemini:.1f}s")

            dados = json.loads(response.text)
            extracao = ExtracaoRoteiro(**dados)

            # ============================================================
            # CORREÇÃO 1: Mapear campos corretamente e adicionar coordenadas do próximo vértice
            # ============================================================
            segmentos = []
            for idx, seg in enumerate(extracao.segmentos):
                # Extrair dados do segmento
                coord_y = limpar_unidade(seg.n_y)
                coord_x = limpar_unidade(seg.e_x)
                distancia = limpar_unidade(seg.distancia)
                
                # Determinar coordenadas do próximo vértice
                if idx + 1 < len(extracao.segmentos):
                    proximo = extracao.segmentos[idx + 1]
                    coord_y_para = limpar_unidade(proximo.n_y)
                    coord_x_para = limpar_unidade(proximo.e_x)
                else:
                    # Último segmento volta para o primeiro
                    primeiro = extracao.segmentos[0]
                    coord_y_para = limpar_unidade(primeiro.n_y)
                    coord_x_para = limpar_unidade(primeiro.e_x)
                
                # Criar segmento com campos corretos
                item = {
                    'de': seg.de,
                    'para': seg.para,
                    'coord_y': coord_y,              # ✅ CORRIGIDO: era 'n_y'
                    'coord_x': coord_x,              # ✅ CORRIGIDO: era 'e_x'
                    'coord_y_para': coord_y_para,    # ✅ NOVO: coordenada Y do próximo vértice
                    'coord_x_para': coord_x_para,    # ✅ NOVO: coordenada X do próximo vértice
                    'azimute': seg.azimute,
                    'distancia': distancia,          # ✅ CORRIGIDO: sem " m"
                    'confrontante': '',              # Será preenchido depois
                    'matricula': '',                 # ✅ NOVO: será preenchido depois
                }
                segmentos.append(item)

            logger.info(f"✅ {len(segmentos)} segmentos extraídos via IA (com coordenadas do próximo vértice)")
            self.segmentos = segmentos
            return segmentos

        except ValidationError as e:
            logger.error(f"Resposta da IA inválida: {str(e)}")
            raise ValueError(f"Resposta inválida da IA: {str(e)}")
        except json.JSONDecodeError as e:
            logger.error(f"JSON inválido da IA: {str(e)}")
            raise ValueError(f"Resposta não é JSON válido: {str(e)}")
        except Exception as e:
            logger.error(f"Erro ao extrair roteiro: {str(e)}")
            raise

    def parse_tabela_roteiro_texto(self, texto_roteiro: str) -> List[Dict[str, str]]:
        """Parse manual de tabela (fallback quando não há PDF)."""
        try:
            # Padrão 1: Com aspas
            pattern1 = r'"(\d+)","(\d+)","([\d\.,]+)","([\d\.,]+)","([^"]+)","([\d\.,]+\s*m)"'
            matches = re.findall(pattern1, texto_roteiro)
            
            if not matches:
                logger.warning("Padrão 1 não encontrado, tentando padrão 2...")
                # Padrão 2: Sem aspas
                pattern2 = r'(\d+)\s+(\d+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([°\d\'\"\s\.]+)\s+([\d\.,]+\s*m)'
                matches = re.findall(pattern2, texto_roteiro)
            
            if not matches:
                logger.warning("Nenhum padrão encontrou correspondências")
                return []
            
            # ============================================================
            # CORREÇÃO 2: Mapear campos corretamente no parse manual
            # ============================================================
            segmentos = []
            for idx, m in enumerate(matches):
                try:
                    az = m[4].replace('$', '').replace('\\circ', '°').strip()
                    
                    # Coordenadas do próximo vértice
                    if idx + 1 < len(matches):
                        proximo = matches[idx + 1]
                        coord_y_para = proximo[2]
                        coord_x_para = proximo[3]
                    else:
                        primeiro = matches[0]
                        coord_y_para = primeiro[2]
                        coord_x_para = primeiro[3]
                    
                    segmento = {
                        "de": m[0],
                        "para": m[1],
                        "coord_y": m[2],                    # ✅ CORRIGIDO
                        "coord_x": m[3],                    # ✅ CORRIGIDO
                        "coord_y_para": coord_y_para,       # ✅ NOVO
                        "coord_x_para": coord_x_para,       # ✅ NOVO
                        "azimute": az,
                        "distancia": limpar_unidade(m[5]), # ✅ CORRIGIDO: sem " m"
                        "confrontante": "",
                        "matricula": "",                    # ✅ NOVO
                    }
                    segmentos.append(segmento)
                except Exception as e:
                    logger.warning(f"Erro ao processar segmento: {str(e)}")
                    continue
            
            logger.info(f"✅ {len(segmentos)} segmentos extraídos do texto")
            self.segmentos = segmentos
            return segmentos
            
        except Exception as e:
            logger.error(f"Erro ao fazer parse do texto: {str(e)}")
            raise

    @retry_com_backoff(
        max_retries=GEMINI_CONFIG.MAX_RETRIES,
        delay_inicial=GEMINI_CONFIG.DELAY_RETRY_INICIAL
    )
    def mapear_confrontantes(
        self,
        imagens_planta: Optional[List[Image.Image]] = None,
        texto_planta: Optional[str] = None,
        texto_roteiro: Optional[str] = None,
    ) -> MapeamentoConfrontantes:
        """
        Mapeia confrontantes usando Gemini.
        
        Com retry automático e tratamento de rate limiting.
        
        Args:
            imagens_planta: Imagens da planta (opcional)
            texto_planta: Texto da planta (opcional)
            texto_roteiro: Texto do roteiro como contexto (opcional)
        
        Returns:
            MapeamentoConfrontantes com regras extraídas
            
        Raises:
            ValueError: Se mapeamento for inválido
        """
        try:
            prompt = """
            Você é um engenheiro agrimensor. Analise o(s) documento(s) para mapear
            confrontantes de forma sistemática.

            Para cada confrontante:
            1. Determine intervalo de pontos (ponto_inicio e ponto_fim)
            2. Exemplo: Pontos 7-21 confrontam com 'Matrícula 1234 JOAO'
            3. Para ciclos: ponto_inicio > ponto_fim (ex: 21→1)
            4. Inclua matrícula e nomes (sem CPF)
            5. Não invente dados

            Retorne ESTRITAMENTE em JSON estruturado.
            """

            if texto_planta:
                prompt += f"\n\nPLANTA (texto):\n{texto_planta}\n"
            if texto_roteiro:
                prompt += f"\n\nROTEIRO (contexto):\n{texto_roteiro}\n"

            tempo_inicio_gemini = time.time()
            logger.info("Chamando Gemini para mapeamento de confrontantes...")

            model = genai.GenerativeModel(
                model_name=self.nome_modelo,
                generation_config=types.GenerationConfig(
                    response_mime_type="application/json",
                    response_schema=MapeamentoConfrontantes,
                    temperature=GEMINI_CONFIG.TEMPERATURE,
                ),
            )

            conteudo = [prompt] + list(imagens_planta or [])
            response = model.generate_content(conteudo)
            
            tempo_gemini = time.time() - tempo_inicio_gemini
            self.tempo_gemini += tempo_gemini
            logger.info(f"Gemini respondeu em {tempo_gemini:.1f}s")

            response_data = json.loads(response.text)
            mapeamento = MapeamentoConfrontantes(**response_data)
            
            self.mapeamento = mapeamento
            logger.info(f"✅ {len(mapeamento.regras)} regras de confrontantes extraídas")
            return mapeamento

        except ValidationError as e:
            logger.error(f"Resposta da IA inválida: {str(e)}")
            raise ValueError(f"Resposta inválida: {str(e)}")
        except json.JSONDecodeError as e:
            logger.error(f"JSON inválido: {str(e)}")
            raise ValueError(f"Resposta não é JSON válido: {str(e)}")
        except Exception as e:
            logger.error(f"Erro ao mapear confrontantes: {str(e)}")
            raise

    # ==========================================
    # VINCULAÇÃO E VALIDAÇÃO
    # ==========================================

    def vincular_confrontantes(self) -> List[Dict]:
        """
        Vincula confrontantes aos segmentos.
        
        ✅ CORRIGIDO: Agora extrai matrícula e nome separadamente
        
        Returns:
            Lista de segmentos com confrontantes preenchidos
        """
        if not self.mapeamento:
            logger.warning("Sem mapeamento de confrontantes")
            return self.segmentos
        
        logger.info("Vinculando confrontantes aos segmentos...")
        
        for seg in self.segmentos:
            try:
                v_de = int(seg["de"])
                
                confrontante_encontrado = None
                
                for regra in self.mapeamento.regras:
                    if regra.ponto_inicio < regra.ponto_fim:
                        if regra.ponto_inicio <= v_de < regra.ponto_fim:
                            confrontante_encontrado = regra.nome_confrontante.upper()
                            break
                    else:  # Ciclo fechado
                        if v_de >= regra.ponto_inicio or v_de <= regra.ponto_fim:
                            confrontante_encontrado = regra.nome_confrontante.upper()
                            break
                
                # ============================================================
                # CORREÇÃO 3: Extrair matrícula e nome separadamente
                # ============================================================
                if confrontante_encontrado:
                    matricula, nome = extrair_matricula_e_nome(confrontante_encontrado)
                    seg["matricula"] = matricula
                    seg["confrontante"] = nome
                else:
                    seg["matricula"] = "N/A"
                    seg["confrontante"] = "CONFRONTAÇÃO NÃO ENCONTRADA"
                
            except (ValueError, IndexError) as e:
                logger.warning(f"Erro ao vincular segmento {seg}: {str(e)}")
                seg["confrontante"] = "ERRO NA VINCULAÇÃO"
                seg["matricula"] = "N/A"
        
        logger.info("✅ Vinculação concluída (com matrícula e nome separados)")
        return self.segmentos

    def validar_resultado(self) -> Tuple[bool, List[str]]:
        """
        Valida resultado do processamento.
        
        Returns:
            Tuple[bool, List[str]]: (é_válido, lista_de_avisos)
        """
        avisos = []
        
        if not self.segmentos:
            avisos.append("⚠️ Nenhum segmento foi extraído")
        
        confrontacoes_invalidas = [
            s for s in self.segmentos
            if "ERRO" in s.get("confrontante", "") or "NÃO ENCONTRADA" in s.get("confrontante", "")
        ]
        if confrontacoes_invalidas:
            avisos.append(f"⚠️ {len(confrontacoes_invalidas)} segmento(s) com confrontação inválida")
        
        # ✅ NOVO: Validar coordenadas do próximo vértice
        coordenadas_vazias = [
            s for s in self.segmentos
            if not s.get("coord_y") or not s.get("coord_x") or not s.get("coord_y_para") or not s.get("coord_x_para")
        ]
        if coordenadas_vazias:
            avisos.append(f"⚠️ {len(coordenadas_vazias)} segmento(s) com coordenadas faltando")
        
        # ✅ NOVO: Validar matrícula
        matriculas_vazias = [
            s for s in self.segmentos
            if not s.get("matricula")
        ]
        if matriculas_vazias:
            avisos.append(f"⚠️ {len(matriculas_vazias)} segmento(s) com matrícula faltando")
        
        return len(avisos) == 0, avisos

    def obter_tempo_decorrido(self) -> float:
        """Retorna tempo decorrido desde inicialização."""
        return time.time() - self.tempo_inicio
