"""
Lógica pura e testável separada da interface Streamlit.
Este módulo contém funções que podem ser testadas unitariamente sem dependências de UI.
"""

import io
import re
import logging
import json
import time
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple

import fitz  # PyMuPDF
from PIL import Image
from pypdf import PdfReader
from pydantic import BaseModel, ValidationError
import google.generativeai as genai
from google.generativeai import types
from google.api_core import exceptions

logger = logging.getLogger(__name__)

# ==========================================
# CONSTANTES E PROMPTS
# ==========================================
PROMPT_ROTEIRO = """
Você é um especialista em leitura de plantas e tabelas topográficas. A(s) imagem(ns)
anexada(s) contém uma TABELA DE ROTEIRO PERIMÉTRICO de um levantamento topográfico
(linhas de vértices/pontos com coordenadas, azimute e distância entre cada ponto e o
próximo).

Leia a tabela LINHA POR LINHA, na ordem em que aparecem, sem pular nenhuma, e para
cada linha/segmento extraia os seguintes campos, **todos são obrigatórios e devem estar presentes no JSON, mesmo que vazios se não puderem ser lidos**:
- de: número do vértice de origem (string)
- para: número do vértice de destino (próximo vértice da linha) (string)
- n_y: coordenada N (Norte/Y) do vértice de origem, incluindo a unidade "m" se houver (string)
- e_x: coordenada E (Este/X) do vértice de origem, incluindo a unidade "m" se houver (string)
- azimute: azimute do segmento no formato graus/minutos/segundos (ex: 45°12'33") (string)
- distancia: distância do segmento, incluindo a unidade "m" (string)

Transcreva EXATAMENTE os valores que aparecem na imagem. Não invente, arredonde ou
preencha valores que não conseguir ler com certeza — nesse caso deixe o campo como
string vazia "".
"""

PROMPT_CONFRONTANTES = """
Você é um engenheiro agrimensor especialista em topografia. Analise o(s) documento(s)
abaixo (imagem e/ou texto) para mapear os confrontantes da Gleba A.

Sua tarefa é extrair os dados cadastrais solicitados e criar as regras matemáticas de transição de confrontantes.

INSTRUÇÕES CRÍTICAS:
1. Para cada confrontante, determine o intervalo de pontos (ponto_inicio e ponto_fim)
2. Exemplo: Se do ponto 7 ao 21 confronta com 'Matrícula nº 1234 propriedade de JOAO', crie: ponto_inicio: 7, ponto_fim: 21, nome_confrontante: 'Matrícula nº 1234 propriedade de JOAO'
3. Se houver fechamento do ciclo (ex: de ponto 21 para 1), use: ponto_inicio: 21, ponto_fim: 1
4. Tente incluir a matrícula e os nomes dos vizinhos, não é necessário o CPF.
5. Retorne ESTRITAMENTE no formato JSON estruturado fornecido
6. NÃO invente dados. Se não conseguir extrair um campo, deixe como string vazia ""
"""

PADROES_TABELA = [
    r'"(\d+)","(\d+)","([\d\.,]+)","([\d\.,]+)","([^"]+)","([\d\.,]+\s*m)"',
    r'(\d+)\s+(\d+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([°\d\'\"\s\.]+)\s+([\d\.,]+\s*m)'
]

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
# VALIDAÇÃO
# ==========================================
def validar_dados_cliente(
    imovel: str,
    proprietario: str,
    local: str,
    area: str,
    perimetro: str
) -> List[str]:
    """Valida os dados do cliente e retorna lista de erros (vazia se válido)."""
    erros = []

    if not imovel or not imovel.strip():
        erros.append("Imóvel não pode estar vazio")
    
    if not proprietario or not proprietario.strip():
        erros.append("Proprietário não pode estar vazio")
    
    if not local or not local.strip():
        erros.append("Local não pode estar vazio")

    try:
        area_num = float(area.replace(",", ".").replace(" ha", "").strip())
        if area_num <= 0:
            erros.append("Área deve ser maior que zero")
    except ValueError:
        erros.append("Área deve ser um número válido (ex: 0,16 ou 0.16)")

    try:
        perim_num = float(perimetro.replace(",", ".").replace(" m", "").strip())
        if perim_num <= 0:
            erros.append("Perímetro deve ser maior que zero")
    except ValueError:
        erros.append("Perímetro deve ser um número válido (ex: 206,42 ou 206.42)")

    return erros


# ==========================================
# PDF PROCESSING
# ==========================================
def verificar_pdf_tipo(arquivo_pdf) -> Tuple[str, bool]:
    """Verifica se o PDF contém texto ou é apenas imagens."""
    try:
        arquivo_pdf.seek(0)
        leitor = PdfReader(arquivo_pdf)
        
        texto_total = ""
        for pagina in leitor.pages:
            texto = pagina.extract_text()
            if texto:
                texto_total += texto
        
        tem_texto = len(texto_total.strip()) > 100
        tipo = "Texto" if tem_texto else "Imagem"
        
        logger.info(f"PDF detectado como: {tipo}")
        return tipo, tem_texto
        
    except Exception as e:
        logger.error(f"Erro ao verificar tipo de PDF: {str(e)}")
        return "Desconhecido", False


def pdf_paginas_para_imagens(arquivo_pdf, dpi: int = 200) -> List[Image.Image]:
    """Converte cada página de um PDF em uma imagem PIL usando PyMuPDF.
    
    Usa context manager para garantir liberação de memória.
    """
    documento = None
    try:
        arquivo_pdf.seek(0)
        dados = arquivo_pdf.read()
        documento = fitz.open(stream=dados, filetype="pdf")

        zoom = dpi / 72
        matriz = fitz.Matrix(zoom, zoom)

        imagens = []
        for pagina in documento:
            pix = pagina.get_pixmap(matrix=matriz)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            imagens.append(img)

        logger.info(f"✅ PDF convertido em {len(imagens)} página(s) a {dpi} DPI")
        return imagens

    except Exception as e:
        logger.error(f"Erro ao rasterizar PDF: {str(e)}")
        raise ValueError(f"Não foi possível abrir o PDF para leitura visual: {str(e)}")
    finally:
        if documento:
            documento.close()


# ==========================================
# PARSING E EXTRAÇÃO
# ==========================================
def _processar_matches_tabela(matches: List[tuple]) -> List[Dict[str, str]]:
    """Processa matches de regex da tabela de roteiro."""
    segmentos = []
    for m in matches:
        try:
            az = m[4].replace('$', '').replace('\\circ', '°')\
                     .replace('\\prime\\prime', '"').replace('\\prime', "'")\
                     .replace('\\:', '').strip()
            
            segmento = {
                "de": m[0],
                "para": m[1],
                "n_y": m[2] + " m",
                "e_x": m[3] + " m",
                "azimute": az,
                "distancia": m[5].strip(),
                "confrontante": ""
            }
            segmentos.append(segmento)
            logger.debug(f"Segmento extraído: {m[0]} → {m[1]}")
            
        except Exception as e:
            logger.warning(f"Erro ao processar segmento {m}: {str(e)}")
            continue
    
    return segmentos


def parse_tabela_roteiro(texto_roteiro: str) -> List[Dict[str, str]]:
    """Extrai dados da tabela a partir de TEXTO já extraído."""
    try:
        for pattern in PADROES_TABELA:
            matches = re.findall(pattern, texto_roteiro)
            if matches:
                segmentos = _processar_matches_tabela(matches)
                if segmentos:
                    logger.info(f"✅ Total de segmentos extraídos: {len(segmentos)}")
                    return segmentos
        
        logger.warning("Nenhum padrão encontrou correspondências")
        return []
        
    except Exception as e:
        logger.error(f"Erro ao fazer parse da tabela de roteiro: {str(e)}")
        raise


# ==========================================
# CHAMADAS À API GEMINI COM RETRY
# ==========================================
def _chamar_gemini_com_retry(
    model,
    conteudo: List,
    max_tentativas: int = 3
):
    """Wrapper para chamadas ao Gemini com retry e backoff exponencial."""
    for tentativa in range(max_tentativas):
        try:
            response = model.generate_content(conteudo)
            return response
        except exceptions.ResourceExhausted:
            if tentativa < max_tentativas - 1:
                espera = 2 ** tentativa
                logger.warning(f"⏳ Rate limit atingido. Aguardando {espera}s (tentativa {tentativa + 1}/{max_tentativas})")
                time.sleep(espera)
            else:
                raise ValueError("Rate limit da API excedido após várias tentativas")
        except exceptions.ApiError as e:
            logger.error(f"Erro na API Gemini (tentativa {tentativa + 1}): {str(e)}")
            if tentativa < max_tentativas - 1:
                time.sleep(2 ** tentativa)
            else:
                raise


def extrair_roteiro_com_ia(
    imagens_roteiro: List[Image.Image],
    nome_modelo: str
) -> List[Dict[str, str]]:
    """Lê a(s) imagem(ns) da TABELA DE ROTEIRO PERIMÉTRICO usando a visão do Gemini."""
    try:
        if not imagens_roteiro:
            raise ValueError("Nenhuma imagem de roteiro fornecida")

        logger.info(f"Chamando API Gemini (visão) para extrair {len(imagens_roteiro)} página(s)...")

        model = genai.GenerativeModel(
            model_name=nome_modelo,
            generation_config=types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=ExtracaoRoteiro,
                temperature=0.0,
            ),
        )

        conteudo = list(imagens_roteiro) + [PROMPT_ROTEIRO]
        response = _chamar_gemini_com_retry(model, conteudo)

        if not response.text or not response.text.strip():
            raise ValueError("Resposta vazia da API Gemini")

        dados = json.loads(response.text)
        extracao = ExtracaoRoteiro(**dados)

        segmentos = []
        for seg in extracao.segmentos:
            item = seg.model_dump()
            item["confrontante"] = ""
            segmentos.append(item)

        logger.info(f"✅ Total de segmentos extraídos via IA: {len(segmentos)}")
        return segmentos

    except ValidationError as e:
        logger.error(f"Erro de validação ao processar tabela de roteiro: {str(e)}")
        raise ValueError(f"Resposta da IA inválida: {str(e)}")
    except json.JSONDecodeError as e:
        logger.error(f"Erro ao fazer parse JSON: {str(e)}")
        raise ValueError(f"Resposta da IA não é JSON válido: {str(e)}")
    except Exception as e:
        logger.error(f"Erro ao extrair tabela de roteiro com IA: {str(e)}")
        raise


def mapear_confrontantes_gemini(
    nome_modelo: str,
    imagens_planta: Optional[List[Image.Image]] = None,
    texto_planta: Optional[str] = None,
    texto_roteiro: Optional[str] = None,
) -> MapeamentoConfrontantes:
    """Mapeia confrontantes usando Gemini."""
    try:
        prompt = PROMPT_CONFRONTANTES

        if texto_planta and texto_planta.strip():
            prompt += f"\n\nDOCUMENTO (DADOS DA PLANTA em texto):\n{texto_planta}\n"
        if texto_roteiro and texto_roteiro.strip():
            prompt += f"\n\nDOCUMENTO (TABELA DE ROTEIRO PERIMÉTRICO em texto):\n{texto_roteiro}\n"

        if not imagens_planta and not texto_planta:
            logger.warning("Nenhuma imagem ou texto da planta fornecido")
            return MapeamentoConfrontantes(regras=[])

        logger.info("Chamando API Gemini para mapeamento de confrontantes...")

        model = genai.GenerativeModel(
            model_name=nome_modelo,
            generation_config=types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=MapeamentoConfrontantes,
                temperature=0.0,
            ),
        )

        conteudo = list(imagens_planta or []) + [prompt]
        response = _chamar_gemini_com_retry(model, conteudo)
        
        if not response.text or not response.text.strip():
            raise ValueError("Resposta vazia da API Gemini")

        response_data = json.loads(response.text)
        mapeamento = MapeamentoConfrontantes(**response_data)
        logger.info(f"✅ Mapeamento extraído: {len(mapeamento.regras)} regra(s)")
        return mapeamento
        
    except ValidationError as e:
        logger.error(f"Erro de validação ao processar resposta do Gemini: {str(e)}")
        raise ValueError(f"Resposta da IA inválida: {str(e)}")
    except json.JSONDecodeError as e:
        logger.error(f"Erro ao fazer parse JSON: {str(e)}")
        raise ValueError(f"Resposta da IA não é JSON válido: {str(e)}")
    except Exception as e:
        logger.error(f"Erro ao mapear confrontantes: {str(e)}")
        raise


# ==========================================
# VINCULAÇÃO
# ==========================================
def vincular_confrontantes(
    segmentos: List[Dict],
    mapeamento: MapeamentoConfrontantes
) -> List[Dict]:
    """Vincula confrontantes aos segmentos."""
    logger.info("Iniciando vinculação de confrontantes aos segmentos...")
    
    for seg in segmentos:
        try:
            v_de = int(seg["de"])
            v_para = int(seg["para"])
            
            confrontante_encontrado = None
            
            for regra in mapeamento.regras:
                if regra.ponto_inicio < regra.ponto_fim:
                    # Faixa regular: inicio < fim
                    if regra.ponto_inicio <= v_de < regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Faixa regular encontrada")
                        break
                
                elif regra.ponto_inicio > regra.ponto_fim:
                    # Ciclo fechado: inicio > fim
                    if v_de >= regra.ponto_inicio or v_de <= regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Ciclo fechado encontrado")
                        break
            
            if not confrontante_encontrado:
                confrontante_encontrado = "CONFRONTAÇÃO NÃO ENCONTRADA"
                logger.warning(f"Segmento {v_de}→{v_para}: Nenhuma regra correspondente")
            
            seg["confrontante"] = confrontante_encontrado
            
        except ValueError as e:
            logger.error(f"Erro ao converter vértices: {str(e)}")
            seg["confrontante"] = "ERRO NA CONVERSÃO"
        except Exception as e:
            logger.error(f"Erro ao vincular confrontante: {str(e)}")
            seg["confrontante"] = "ERRO NO PROCESSAMENTO"
    
    logger.info("✅ Vinculação de confrontantes concluída")
    return segmentos
