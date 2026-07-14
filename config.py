# ==========================================
# CONFIGURAÇÕES CENTRALIZADAS
# ==========================================
"""
Arquivo de configurações globais para o Gerador de Memorial Descritivo.
Centraliza todas as constantes e configurações da aplicação.
"""

from dataclasses import dataclass
from typing import Dict


@dataclass
class ConfiguracaoGemini:
    """Configurações da API Gemini"""
    MODELOS_DISPONIVEIS: Dict[str, str] = None
    TIMEOUT_PADRAO: int = 120
    MAX_RETRIES: int = 3
    DELAY_RETRY_INICIAL: int = 2
    TEMPERATURE: float = 0.0  # Para extração de dados, sem criatividade

    def __post_init__(self):
        if self.MODELOS_DISPONIVEIS is None:
            self.MODELOS_DISPONIVEIS = {
                "Gemini 3.5 Flash": "gemini-3.5-flash",
                "Gemini 3.1 Pro": "gemini-3.1-pro",
                "Gemini 3.1 Flash Lite": "gemini-3.1-flash-lite",
                "Gemini 2.5 Pro": "gemini-2.5-pro",
                "Gemini 2.5 Flash": "gemini-2.5-flash",
            }


@dataclass
class ConfiguracaoProcessamento:
    """Configurações de processamento de PDFs"""
    DPI_PADRAO: int = 250
    DPI_MINIMO: int = 150
    DPI_MAXIMO: int = 400
    TAMANHO_MAX_PDF_MB: int = 50
    TAMANHO_MIN_PDF_MB: float = 0.1


@dataclass
class ConfiguracaoDocumento:
    """Configurações do documento Word gerado"""
    MARGENS_CM: float = 2.5
    FONTE_PADRAO: str = "Arial"
    TAMANHO_FONTE_PADRAO: int = 11
    TAMANHO_TITULO: int = 12
    TAMANHO_SUBTITULO: int = 10


@dataclass
class ConfiguracaoEmpresa:
    """Configurações padrão da empresa"""
    NOME: str = "TopoGeo Topografia e Consultoria LTDA"
    ENDERECO: str = "Rua Natalino Cossi, No 114, sala 2 - Vila Valério, CEP 29785-000"
    TELEFONE: str = "27 99837-1164"
    EMAIL: str = "topogeo2014@gmail.com"


@dataclass
class ConfiguracaoTecnico:
    """Configurações padrão do técnico responsável"""
    NOME: str = "Régis Campo da Silva"
    CARGO: str = "TÉCNICO EM AGROPECUÁRIA"
    CFTA: str = "11198519711"
    TRT: str = "BR20260210971"


@dataclass
class ConfiguracaoCliente:
    """Configurações padrão do cliente"""
    IMOVEL: str = "Lote"
    PROPRIETARIO: str = "SEBASTIAO IZOTON"
    LOCAL: str = "Vila Valério"
    AREA: str = "0,16 ha"
    PERIMETRO: str = "206,42 m"


# Instâncias singleton para fácil acesso
GEMINI_CONFIG = ConfiguracaoGemini()
PROCESSAMENTO_CONFIG = ConfiguracaoProcessamento()
DOCUMENTO_CONFIG = ConfiguracaoDocumento()
EMPRESA_CONFIG = ConfiguracaoEmpresa()
TECNICO_CONFIG = ConfiguracaoTecnico()
CLIENTE_CONFIG = ConfiguracaoCliente()

# Meses em português para o documento
MESES_PT: Dict[int, str] = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

# Versão da aplicação
VERSAO_APP = "6.0"
DESCRICAO_VERSAO = "Refatoração completa com melhorias de performance, UX e manutenibilidade"
