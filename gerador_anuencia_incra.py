# ==========================================
# ARQUIVO: gerador_anuencia_incra.py (VERSÃO FINAL PAISAGEM)
# ==========================================
import io
import re
import json
import logging
import zipfile
from datetime import datetime
from typing import Dict, Any, List, Tuple
import streamlit as st
import google.generativeai as genai

try:
    import pypdf
except ImportError:
    pypdf = None

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

logger = logging.getLogger(__name__)


class GeradorAnuenciaIncraWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        """
        Inicializa o gerador de anuência do INCRA com dados institucionais e do técnico.
        """
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico

        # Configuração da API do Gemini
        self.api_key = st.secrets.get("GEMINI_API_KEY")
        if self.api_key:
            genai.configure(api_key=self.api_key)

    def _estrutura_padrao(self) -> Dict[str, Any]:
        """
        Estrutura de fallback baseada nos modelos reais.
        """
        return {
            "imovel_origem": {
                "proprietario": "AGOSTINHO IZOTON",
                "cpf": "215.894.707-10",
                "imovel": "GLEBA A",
                "localidade": "Vila Val rio-ES"
            },
            "confrontantes": [
                {
                    "confrontante_imovel": "Sitio Sete Quedas",
                    "confrontante_matricula": "8280",
                    "confrontante_comarca": "o Gabriel da Palha",
                    "confrontante_proprietario": "Elias Moro, Luiz Valentin Moro",
                    "confrontante_cpf": "780.485.677-68",
                    "vertices": [
                        {
                            "codigo": "G1D-P-06815",
                            "longitude": "40°17'14,014\"",
                            "latitude": "18°59'22,007\"",
                            "altitude": "58.39",
                            "vante": "G1D-P-06816",
                            "azimute": "02°15'",
                            "distancia": "41,66",
                            "confrontacao_completa": "CNS: 02.170-9 | Mat. 8280 | Sitio Sete Quedas; Elias Moro, Luiz Valentin Moro"
                        }
                    ]
                }
            ]
        }

    def _obter_dados_estruturados_com_ia(self, texto_memorial: str) -> Dict[str, Any]:
        """
        Usa o Gemini para analisar o memorial e extrair TANTO os dados do proprietário
        de origem quanto os dados de TODOS os confrontantes com seus vértices.
        """
        estrutura_padrao = self._estrutura_padrao()

        if not self.api_key:
            logger.warning("Chave de API do Gemini não configurada nos secrets. Usando dados padrão.")
            return estrutura_padrao

        prompt = f"""
        Você é um engenheiro cartógrafo especialista em georreferenciamento do INCRA.
        Sua tarefa é analisar o texto de um memorial descritivo ou relatório de cálculo/vértices e estruturar de forma impecável as informações do IMÓVEL DE ORIGEM (objeto do memorial) e de TODOS os confrontantes (vizinhos) identificados.

        TEXTO DO MEMORIAL DESCRITIVO / RELATÓRIO EXTRAÍDO:
        \"\"\"
        {texto_memorial}
        \"\"\"

        REGRAS DE EXTRAÇÃO:
        1. Identifique os dados do proprietário principal/origem do memorial (Nome completo, CPF, Nome do Imóvel/Gleba e Localização).
        2. Identifique cada confrontante distinto ao longo da poligonal periférica.
        3. Agrupe sob cada confrontante APENAS os vértices/segmentos cujo trecho de "vante" faz divisa com ele.
        4. Caso o CPF do proprietário de origem ou dos confrontantes não esteja explícito no texto, tente extrair se houver, caso contrário, retorne no formato "___.___.___-__" para preenchimento posterior.

        Responda APENAS com o JSON estruturado abaixo, sem markdown, sem tags ```json ou textos explicativos:
        {{
            "imovel_origem": {{
                "proprietario": "NOME COMPLETO DO PROPRIETÁRIO PRINCIPAL",
                "cpf": "CPF DO PROPRIETÁRIO PRINCIPAL",
                "imovel": "NOME DO IMÓVEL RURAL DE ORIGEM",
                "localidade": "MUNICÍPIO OU LOCALIDADE DE ORIGEM"
            }},
            "confrontantes": [
                {{
                    "confrontante_imovel": "Nome do Imóvel Confrontante",
                    "confrontante_matricula": "Número da Matrícula/Transcrição",
                    "confrontante_comarca": "Nome da Comarca",
                    "confrontante_proprietario": "Nome do Proprietário Confrontante",
                    "confrontante_cpf": "CPF do Confrontante",
                    "vertices": [
                        {{
                            "codigo": "Código do Vértice",
                            "longitude": "Longitude formatada (Graus Minutos Segundos)",
                            "latitude": "Latitude formatada (Graus Minutos Segundos)",
                            "altitude": "Altitude com duas casas",
                            "vante": "Código do Vértice de Vante",
                            "azimute": "Azimute formatado",
                            "distancia": "Distância formatada com duas casas",
                            "confrontacao_completa": "Descrição completa da confrontação"
                        }}
                    ]
                }}
            ]
        }}
        """

        try:
            model = genai.GenerativeModel("gemini-2.5-flash")
            response = model.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"}
            )

            texto_resposta = response.text.strip()
            # Tratamento de segurança para wraps de markdown
            if texto_resposta.startswith("```json"):
                texto_resposta = texto_resposta.split("```json")[1].split("```")[0].strip()
            elif texto_resposta.startswith("```"):
                texto_resposta = texto_resposta.split("
