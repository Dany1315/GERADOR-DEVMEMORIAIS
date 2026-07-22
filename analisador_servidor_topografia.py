"""
Módulo para análise de pastas do servidor de topografia
Monitora a estrutura de pastas por ano e identifica pendências
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Tuple
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class AnalisadorServidorTopografia:
    """
    Analisa o servidor de topografia e identifica pastas incompletas
    Estrutura: \\Servidor\topografia\ANO\CLIENTE\arquivos
    """
    
    # Arquivos e pastas obrigatórios para uma pasta estar completa
    ARQUIVOS_OBRIGATORIOS = {
        "pastas": [
            "ANUENCIAS",
            "DOCUMENTOS",
            "Requerimento Cartório",
            "Boleto",
            "ODS"
        ],
        "arquivos": [
            "memorial",
            "planta",
            "sigef_planilha"
        ],
        "padroes": {
            "trt": "BR",  # Começa com BR
            "desenho": "DESENHO1"  # Pode ter variações
        }
    }
    
    def __init__(self, caminho_servidor: str):
        """
        Inicializa o analisador
        
        Args:
            caminho_servidor: Caminho para \\REDE\SERVIDOR\TOPOGRAFIA
        """
        self.caminho_servidor = caminho_servidor
        self.anos = list(range(2013, 2027))  # 2013 a 2026
        self.resultados = {}
        
    def analisar_servidor(self) -> Dict:
        """
        Analisa todas as pastas do servidor
        
        Returns:
            Dicionário com análise completa
        """
        if not os.path.exists(self.caminho_servidor):
            logger.error(f"Caminho do servidor não encontrado: {self.caminho_servidor}")
            return {
                "status": "erro",
                "mensagem": f"Caminho não encontrado: {self.caminho_servidor}",
                "pastas_analisadas": 0,
                "pendencias": []
            }
        
        pendencias = []
        pastas_analisadas = 0
        pastas_completas = 0
        
        # Iterar sobre cada ano
        for ano in self.anos:
            caminho_ano = os.path.join(self.caminho_servidor, str(ano))
            
            if not os.path.exists(caminho_ano):
                logger.info(f"Pasta do ano {ano} não existe")
                continue
            
            # Listar todas as pastas de clientes dentro do ano
            try:
                pastas_cliente = [d for d in os.listdir(caminho_ano) 
                                 if os.path.isdir(os.path.join(caminho_ano, d))]
            except PermissionError:
                logger.warning(f"Sem permissão para acessar: {caminho_ano}")
                continue
            
            # Analisar cada pasta de cliente
            for pasta_cliente in pastas_cliente:
                caminho_pasta_cliente = os.path.join(caminho_ano, pasta_cliente)
                pastas_analisadas += 1
                
                # Verificar completude da pasta do cliente
                resultado_analise = self._analisar_pasta(caminho_pasta_cliente, ano, pasta_cliente)
                
                if not resultado_analise["completa"]:
                    pendencias.append(resultado_analise)
                else:
                    pastas_completas += 1
        
        return {
            "status": "sucesso",
            "data_analise": datetime.now().isoformat(),
            "pastas_analisadas": pastas_analisadas,
            "pastas_completas": pastas_completas,
            "pastas_incompletas": len(pendencias),
            "taxa_conclusao": f"{(pastas_completas / pastas_analisadas * 100):.1f}%" if pastas_analisadas > 0 else "0%",
            "pendencias": pendencias
        }
    
    def _analisar_pasta(self, caminho_pasta: str, ano: int, nome_cliente: str) -> Dict:
        """
        Analisa uma pasta de cliente individual
        
        Args:
            caminho_pasta: Caminho completo da pasta do cliente
            ano: Ano da pasta
            nome_cliente: Nome da pasta do cliente (ex: GEO AGOSTINHO IZOTON - DONALDSON)
            
        Returns:
            Dicionário com resultado da análise
        """
        resultado = {
            "ano": ano,
            "cliente": nome_cliente,
            "caminho": caminho_pasta,
            "completa": True,
            "faltam": [],
            "encontrados": [],
            "percentual_conclusao": 100
        }
        
        try:
            conteudo = os.listdir(caminho_pasta)
        except PermissionError:
            resultado["completa"] = False
            resultado["faltam"].append("Sem permissão para acessar")
            return resultado
        
        conteudo_lower = [item.lower() for item in conteudo]
        encontrados = 0
        total_esperado = len(self.ARQUIVOS_OBRIGATORIOS["pastas"]) + \
                        len(self.ARQUIVOS_OBRIGATORIOS["arquivos"]) + 2  # +2 para TRT e DESENHO
        
        # Verificar pastas obrigatórias
        for pasta_obrigatoria in self.ARQUIVOS_OBRIGATORIOS["pastas"]:
            if pasta_obrigatoria.lower() in conteudo_lower:
                resultado["encontrados"].append(f"📁 {pasta_obrigatoria}")
                encontrados += 1
            else:
                resultado["completa"] = False
                resultado["faltam"].append(f"📁 {pasta_obrigatoria}")
        
        # Verificar arquivos obrigatórios
        for arquivo_obrigatorio in self.ARQUIVOS_OBRIGATORIOS["arquivos"]:
            # Procurar por arquivo com extensão
            encontrado = False
            for item in conteudo:
                if arquivo_obrigatorio.lower() in item.lower():
                    resultado["encontrados"].append(f"📄 {item}")
                    encontrados += 1
                    encontrado = True
                    break
            
            if not encontrado:
                resultado["completa"] = False
                resultado["faltam"].append(f"📄 {arquivo_obrigatorio}")
        
        # Verificar TRT (começa com BR)
        trt_encontrado = False
        for item in conteudo:
            if item.startswith("BR"):
                resultado["encontrados"].append(f"🔢 TRT: {item}")
                encontrados += 1
                trt_encontrado = True
                break
        
        if not trt_encontrado:
            resultado["completa"] = False
            resultado["faltam"].append("🔢 TRT (BR*)")
        
        # Verificar DESENHO1 (com possíveis variações)
        desenho_encontrado = False
        for item in conteudo:
            if "desenho" in item.lower():
                resultado["encontrados"].append(f"🎨 {item}")
                encontrados += 1
                desenho_encontrado = True
                break
        
        if not desenho_encontrado:
            resultado["completa"] = False
            resultado["faltam"].append("🎨 DESENHO1 (ou variação)")
        
        # Calcular percentual
        resultado["percentual_conclusao"] = int((encontrados / total_esperado) * 100)
        
        return resultado
    
    def obter_pendencias_formatadas(self) -> str:
        """
        Retorna as pendências em formato legível
        
        Returns:
            String formatada com as pendências
        """
        if not self.resultados:
            return "Nenhuma análise realizada ainda"
        
        pendencias = self.resultados.get("pendencias", [])
        
        if not pendencias:
            return "✅ Todos os clientes estão com documentação completa!"
        
        texto = f"📊 ANÁLISE DO SERVIDOR DE TOPOGRAFIA\n"
        texto += f"Data: {self.resultados['data_analise']}\n"
        texto += f"Clientes analisados: {self.resultados['pastas_analisadas']}\n"
        texto += f"Clientes com documentação completa: {self.resultados['pastas_completas']}\n"
        texto += f"Clientes com pendências: {self.resultados['pastas_incompletas']}\n"
        texto += f"Taxa de conclusão: {self.resultados['taxa_conclusao']}\n\n"
        
        texto += "⚠️ CLIENTES COM PENDÊNCIAS:\n"
        texto += "=" * 80 + "\n\n"
        
        for pendencia in pendencias:
            texto += f"📅 ANO: {pendencia['ano']}\n"
            texto += f"👤 CLIENTE: {pendencia['cliente']}\n"
            texto += f"📍 CAMINHO: {pendencia['caminho']}\n"
            texto += f"✅ CONCLUSÃO: {pendencia['percentual_conclusao']}%\n\n"
            
            if pendencia['encontrados']:
                texto += "✅ Encontrados:\n"
                for item in pendencia['encontrados']:
                    texto += f"   {item}\n"
                texto += "\n"
            
            if pendencia['faltam']:
                texto += "❌ Faltam:\n"
                for item in pendencia['faltam']:
                    texto += f"   {item}\n"
                texto += "\n"
            
            texto += "-" * 80 + "\n\n"
        
        return texto
    
    def obter_pendencias_json(self) -> str:
        """
        Retorna as pendências em formato JSON
        
        Returns:
            String JSON com as pendências
        """
        return json.dumps(self.resultados, indent=2, ensure_ascii=False)
    
    def executar_analise_completa(self) -> Dict:
        """
        Executa análise completa e armazena resultado
        
        Returns:
            Dicionário com resultado completo
        """
        self.resultados = self.analisar_servidor()
        return self.resultados


class AnalisadorComGemini:
    """
    Integra análise do servidor com Gemini para insights inteligentes
    """
    
    def __init__(self, api_key: str):
        """
        Inicializa o analisador com Gemini
        
        Args:
            api_key: Chave da API do Gemini
        """
        self.api_key = api_key
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
        except Exception as e:
            logger.error(f"Erro ao configurar Gemini: {str(e)}")
            self.model = None
    
    def analisar_com_gemini(self, dados_analise: Dict) -> str:
        """
        Usa Gemini para gerar insights sobre as pendências
        
        Args:
            dados_analise: Dicionário com resultado da análise
            
        Returns:
            Análise em texto gerada por Gemini
        """
        if not self.model:
            return "Modelo Gemini não disponível"
        
        # Preparar prompt
        pendencias = dados_analise.get("pendencias", [])
        
        prompt = f"""
Você é um especialista em gerenciamento de projetos de topografia. 
Analise os seguintes dados de pendências de clientes do servidor:

Total de clientes analisados: {dados_analise['pastas_analisadas']}
Clientes com documentação completa: {dados_analise['pastas_completas']}
Clientes com pendências: {dados_analise['pastas_incompletas']}
Taxa de conclusão: {dados_analise['taxa_conclusao']}

Pendências encontradas (primeiros 10 clientes):
{json.dumps(pendencias[:10], indent=2, ensure_ascii=False)}

Por favor, forneça:
1. Um resumo executivo da situação
2. Os 3 principais problemas encontrados
3. Quais clientes têm mais pendências
4. Recomendações de priorização
5. Estimativa de tempo para conclusão

Responda em português brasileiro de forma concisa e profissional.
"""
        
        try:
            response = self.model.generate_content(prompt)
            return response.text
        except Exception as e:
            logger.error(f"Erro ao gerar análise com Gemini: {str(e)}")
            return f"Erro ao gerar análise: {str(e)}"


def testar_conexao_servidor(caminho_servidor: str) -> Tuple[bool, str]:
    """
    Testa se consegue acessar o servidor
    
    Args:
        caminho_servidor: Caminho para testar
        
    Returns:
        Tupla (sucesso, mensagem)
    """
    try:
        if os.path.exists(caminho_servidor):
            conteudo = os.listdir(caminho_servidor)
            anos_encontrados = [d for d in conteudo if os.path.isdir(os.path.join(caminho_servidor, d)) and d.isdigit()]
            return True, f"✅ Servidor acessível. Encontrados {len(anos_encontrados)} anos com dados."
        else:
            return False, f"❌ Caminho não encontrado: {caminho_servidor}"
    except PermissionError:
        return False, f"❌ Sem permissão para acessar: {caminho_servidor}"
    except Exception as e:
        return False, f"❌ Erro ao acessar servidor: {str(e)}"
