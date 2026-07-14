"""
Funções utilitárias para validação, tratamento de erros e otimizações.
"""

import io
import time
import logging
from typing import Tuple, Optional, Callable, TypeVar, Any
import streamlit as st

logger = logging.getLogger(__name__)

# ==========================================
# VALIDAÇÃO DE ARQUIVOS
# ==========================================

def validar_arquivo_pdf(arquivo, tamanho_max_mb: int = 50, tamanho_min_mb: float = 0.1) -> Tuple[bool, str]:
    """
    Valida arquivo PDF antes do processamento.
    
    Args:
        arquivo: Arquivo do Streamlit file_uploader
        tamanho_max_mb: Tamanho máximo em MB
        tamanho_min_mb: Tamanho mínimo em MB
    
    Returns:
        Tuple[bool, str]: (é_válido, mensagem)
    """
    try:
        if arquivo is None:
            return False, "Arquivo não foi selecionado"
        
        arquivo.seek(0, 2)  # Vai para o fim
        tamanho_bytes = arquivo.tell()
        tamanho_mb = tamanho_bytes / (1024 * 1024)
        
        if tamanho_mb > tamanho_max_mb:
            return False, f"PDF muito grande ({tamanho_mb:.1f}MB). Máximo: {tamanho_max_mb}MB"
        
        if tamanho_mb < tamanho_min_mb:
            return False, "PDF muito pequeno (parece vazio)"
        
        arquivo.seek(0)
        return True, "OK"
        
    except Exception as e:
        logger.error(f"Erro ao validar arquivo PDF: {str(e)}")
        return False, f"Erro ao validar arquivo: {str(e)}"


def validar_texto_entrada(texto: str, min_length: int = 10) -> Tuple[bool, str]:
    """
    Valida texto de entrada manual.
    
    Args:
        texto: Texto a validar
        min_length: Comprimento mínimo
    
    Returns:
        Tuple[bool, str]: (é_válido, mensagem)
    """
    if not texto or len(texto.strip()) < min_length:
        return False, f"Texto muito curto (mínimo {min_length} caracteres)"
    return True, "OK"


# ==========================================
# RETRY COM EXPONENTIAL BACKOFF
# ==========================================

def retry_com_backoff(
    max_retries: int = 3,
    delay_inicial: float = 2.0,
    backoff_multiplier: float = 2.0
) -> Callable:
    """
    Decorator para retry com exponential backoff.
    
    Útil para lidar com rate limiting da API Gemini e timeouts temporários.
    
    Args:
        max_retries: Número máximo de tentativas
        delay_inicial: Delay inicial em segundos
        backoff_multiplier: Multiplicador de delay a cada retry
    
    Returns:
        Função decorada com retry
    """
    def decorator(func: Callable) -> Callable:
        def wrapper(*args, **kwargs) -> Any:
            delay = delay_inicial
            ultima_excecao = None
            
            for tentativa in range(max_retries):
                try:
                    return func(*args, **kwargs)
                    
                except Exception as e:
                    ultima_excecao = e
                    eh_rate_limit = "429" in str(e) or "quota" in str(e).lower()
                    eh_timeout = "timeout" in str(e).lower() or "deadline" in str(e).lower()
                    
                    if (eh_rate_limit or eh_timeout) and tentativa < max_retries - 1:
                        logger.warning(
                            f"Tentativa {tentativa + 1}/{max_retries} falhou. "
                            f"Aguardando {delay}s antes de retry... (Erro: {str(e)[:50]})"
                        )
                        time.sleep(delay)
                        delay *= backoff_multiplier
                    else:
                        logger.error(f"Falha permanente ou última tentativa: {str(e)}")
                        raise
            
            # Se chegou aqui, todas as tentativas falharam
            if ultima_excecao:
                raise ultima_excecao
                
        return wrapper
    return decorator


# ==========================================
# CACHE COM STREAMLIT
# ==========================================

def limpar_cache_sessao():
    """Limpa o cache da sessão do Streamlit."""
    if hasattr(st, 'session_state'):
        for key in list(st.session_state.keys()):
            if key.startswith("cache_"):
                del st.session_state[key]
        logger.info("Cache da sessão limpo")


def cache_sessao(chave: str) -> Callable:
    """
    Decorator para cache na sessão do Streamlit.
    
    Evita reprocessar os mesmos PDFs múltiplas vezes na mesma sessão.
    
    Args:
        chave: Identificador único para o cache
    
    Returns:
        Função com cache
    """
    def decorator(func: Callable) -> Callable:
        def wrapper(*args, **kwargs) -> Any:
            chave_cache = f"cache_{chave}_{hash(str(args) + str(kwargs))}"
            
            if chave_cache in st.session_state:
                logger.debug(f"Retornando resultado do cache: {chave_cache}")
                return st.session_state[chave_cache]
            
            resultado = func(*args, **kwargs)
            st.session_state[chave_cache] = resultado
            logger.debug(f"Resultado armazenado em cache: {chave_cache}")
            return resultado
        
        return wrapper
    return decorator


# ==========================================
# FORMATAÇÃO E CONVERSÃO
# ==========================================

def formatar_bytes_para_mb(bytes_size: int) -> str:
    """Converte bytes para MB de forma legível."""
    mb = bytes_size / (1024 * 1024)
    return f"{mb:.2f} MB"


def formatar_tempo_decorrido(segundos: float) -> str:
    """Formata tempo decorrido em formato legível."""
    if segundos < 60:
        return f"{segundos:.1f}s"
    elif segundos < 3600:
        minutos = segundos / 60
        return f"{minutos:.1f}m"
    else:
        horas = segundos / 3600
        return f"{horas:.1f}h"


def sanitizar_nome_arquivo(nome: str) -> str:
    """Remove caracteres inválidos do nome do arquivo."""
    caracteres_invalidos = r'<>:"/\|?*'
    for char in caracteres_invalidos:
        nome = nome.replace(char, '')
    return nome.strip()


# ==========================================
# LOGGING E MONITORAMENTO
# ==========================================

class LoggerEstruturado:
    """Logger com contexto estruturado."""
    
    def __init__(self, nome: str):
        self.logger = logging.getLogger(nome)
    
    def info_com_contexto(self, mensagem: str, **contexto):
        """Log com contexto estruturado."""
        msg = f"{mensagem} | {' | '.join(f'{k}={v}' for k, v in contexto.items())}"
        self.logger.info(msg)
    
    def erro_com_contexto(self, mensagem: str, **contexto):
        """Log de erro com contexto estruturado."""
        msg = f"{mensagem} | {' | '.join(f'{k}={v}' for k, v in contexto.items())}"
        self.logger.error(msg, exc_info=True)


def criar_logger(nome: str) -> logging.Logger:
    """Cria um logger configurado."""
    logger = logging.getLogger(nome)
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    return logger


# ==========================================
# GERAÇÃO DE RELATÓRIOS
# ==========================================

def gerar_relatorio_processamento(
    dados_finais: dict,
    tempo_inicio: float,
    tempo_fim: float,
    tempo_gemini: float = 0
) -> str:
    """
    Gera relatório completo do processamento.
    
    Args:
        dados_finais: Dados processados
        tempo_inicio: Timestamp do início
        tempo_fim: Timestamp do fim
        tempo_gemini: Tempo total gasto em chamadas Gemini
    
    Returns:
        String com o relatório formatado
    """
    from datetime import datetime
    from config import VERSAO_APP
    
    tempo_total = tempo_fim - tempo_inicio
    
    relatorio = f"""
{'='*80}
RELATÓRIO DE PROCESSAMENTO - GERADOR DE MEMORIAL DESCRITIVO
{'='*80}

VERSÃO: {VERSAO_APP}
DATA/HORA: {datetime.fromtimestamp(tempo_fim).strftime('%d/%m/%Y %H:%M:%S')}

INFORMAÇÕES DO PROCESSAMENTO:
{'-'*80}
Proprietário:        {dados_finais.get('proprietario', 'NÃO INFORMADO')}
Imóvel:             {dados_finais.get('imovel', 'NÃO INFORMADO')}
Local:              {dados_finais.get('local', 'NÃO INFORMADO')}
Área:               {dados_finais.get('area', 'NÃO INFORMADO')}
Perímetro:          {dados_finais.get('perimetro', 'NÃO INFORMADO')}

RESULTADOS:
{'-'*80}
Segmentos Processados:  {len(dados_finais.get('segmentos', []))}
Confrontantes Únicas:   {len(set(s.get('confrontante', '') for s in dados_finais.get('segmentos', [])))}

DESEMPENHO:
{'-'*80}
Tempo Total:        {formatar_tempo_decorrido(tempo_total)}
Tempo Gemini API:   {formatar_tempo_decorrido(tempo_gemini)}
Tempo Local:        {formatar_tempo_decorrido(tempo_total - tempo_gemini)}

{'='*80}
"""
    return relatorio
