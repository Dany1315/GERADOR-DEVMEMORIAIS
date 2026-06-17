#Correções Implementadas:
#- Importações corretas do Google Generative AI
#- Tratamento robusto de exceções
#- Validação completa de dados
#- Lógica de mapeamento de confrontantes corrigida
#- Configuração flexível de dados da empresa e técnico
#- Logging e debugging melhorados
#- Melhor tratamento de erros em tempo de execução
#"""

import io
import re
import logging
import json
from datetime import datetime
from typing import Optional, List, Dict, Any

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from pypdf import PdfReader
import streamlit as st
from pydantic import BaseModel, ValidationError
import google.generativeai as genai
from google.generativeai import types

# ==========================================
# CONFIGURAÇÃO DE LOGGING
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==========================================
# CONFIGURAÇÃO DA PÁGINA DO STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Gerador de Memorial Descritivo - Gleba A",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# CONFIGURAÇÕES PADRÃO (FLEXÍVEIS)
# ==========================================
EMPRESA_INFO = {
    "nome": "TopoGeo Topografia e Consultoria LTDA",
    "endereco": "Rua Natalino Cossi, Nº 114, sala 2 - Vila Valério, CEP 29785-000",
    "telefone": "27 99837-1164",
    "email": "topogeo2014@gmail.com"
}

TECNICO_INFO = {
    "nome": "Régis Campo da Silva",
    "cargo": "TÉCNICO EM AGROPECUÁRIA",
    "cfta": "11198519711",
    "trt": "BR20260210971"
}

MARGENS_CM = 2.5
FONTE_PADRAO = "Arial"
TAMANHO_FONTE_PADRAO = 11

# ==========================================
# MODELOS PYDANTIC ESTRUTURADOS
# ==========================================
class RegraConfrontante(BaseModel):
    """Modelo para regra de confrontante com validação"""
    ponto_inicio: int
    ponto_fim: int
    nome_confrontante: str

    class Config:
        str_strip_whitespace = True


class MapeamentoConfrontantes(BaseModel):
    """Modelo para mapeamento completo de confrontantes"""
    proprietario: str
    municipio: str
    comarca: str
    area: str
    perimetro: str
    regras: List[RegraConfrontante]

    class Config:
        str_strip_whitespace = True


# ==========================================
# FUNÇÕES DE EXTRAÇÃO E PARSING
# ==========================================
def extrair_texto_pdf(arquivo_pdf) -> str:
    """
    Lê o arquivo PDF enviado e extrai todo o texto contido nele.
    
    Args:
        arquivo_pdf: Arquivo PDF carregado via Streamlit
        
    Returns:
        str: Texto completo extraído do PDF
        
    Raises:
        Exception: Se houver erro na leitura do PDF
    """
    try:
        leitor = PdfReader(arquivo_pdf)
        texto_completo = ""
        
        logger.info(f"Extraindo texto de PDF com {len(leitor.pages)} páginas")
        
        for idx, pagina in enumerate(leitor.pages):
            texto_pagina = pagina.extract_text()
            if texto_pagina:
                texto_completo += texto_pagina + "\n"
            else:
                logger.warning(f"Página {idx + 1} não contém texto extraível")
        
        if not texto_completo.strip():
            raise ValueError("Nenhum texto foi extraído do PDF. Verifique se o arquivo é válido.")
        
        logger.info(f"Total de caracteres extraídos: {len(texto_completo)}")
        return texto_completo
        
    except Exception as e:
        logger.error(f"Erro ao extrair texto do PDF: {str(e)}")
        raise


def parse_tabela_roteiro(texto_roteiro: str) -> List[Dict[str, str]]:
    """
    Extrai via Regex os dados exatos da tabela do PDF.
    Garante que o Python leia os números reais sem que a IA invente dados.
    
    Args:
        texto_roteiro: Texto extraído do PDF do roteiro perimétrico
        
    Returns:
        List[Dict]: Lista de segmentos com dados de coordenadas e azimutes
    """
    try:
        # Regex para capturar as linhas da tabela: De, Para, Coord N, Coord E, Azimute, Distancia
        pattern = r'"(\d+)","(\d+)","([\d\.,]+)","([\d\.,]+)","([^"]+)","([\d\.,]+\s*m)"'
        matches = re.findall(pattern, texto_roteiro)
        
        if not matches:
            logger.warning("Nenhuma correspondência encontrada com a regex padrão. Tentando padrão alternativo...")
            # Padrão alternativo mais flexível
            pattern_alt = r'(\d+)\s+(\d+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([°\d\'\"\s\.]+)\s+([\d\.,]+\s*m)'
            matches = re.findall(pattern_alt, texto_roteiro)
        
        segmentos = []
        for m in matches:
            try:
                # Limpa elementos de LaTeX que possam vir no Azimute
                az = m[4].replace('$', '').replace('\\circ', '°').replace('\\prime\\prime', '"').replace('\\prime', "'").strip()
                az = az.replace('\\:', '').strip()
                
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
        
        if not segmentos:
            logger.warning("Nenhum segmento foi extraído da tabela")
        else:
            logger.info(f"Total de segmentos extraídos: {len(segmentos)}")
        
        return segmentos
        
    except Exception as e:
        logger.error(f"Erro ao fazer parse da tabela de roteiro: {str(e)}")
        raise


# ==========================================
# INTEGRAÇÃO COM GOOGLE GENERATIVE AI
# ==========================================
def configurar_gemini() -> bool:
    """
    Configura a conexão com a API do Google Generative AI.
    
    Returns:
        bool: True se configurado com sucesso, False caso contrário
    """
    try:
        # Tenta obter a chave do Streamlit Secrets
        api_key = st.secrets.get("GEMINI_API_KEY")
        
        if not api_key:
            logger.error("Chave GEMINI_API_KEY não encontrada nos Secrets")
            return False
        
        genai.configure(api_key=api_key)
        logger.info("Gemini configurado com sucesso")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao configurar Gemini: {str(e)}")
        return False


def mapear_confrontantes_gemini(texto_planta: str, texto_roteiro: str) -> Optional[MapeamentoConfrontantes]:
    """
    Utiliza o Gemini 3.5 Flash para interpretar a lógica de quais confrontantes
    pertencem aos respectivos intervalos de pontos e extrair os dados do cabeçalho.
    
    Args:
        texto_planta: Texto extraído da planta com dados de confrontantes
        texto_roteiro: Texto extraído do roteiro perimétrico
        
    Returns:
        MapeamentoConfrontantes: Objeto com mapeamento de confrontantes
        
    Raises:
        Exception: Se houver erro na chamada à API
    """
    try:
        prompt = f"""
        Você é um engenheiro agrimensor especialista em topografia. Analise os documentos abaixo para mapear os confrontantes da Gleba A.

        DOCUMENTO 1 (DADOS DA PLANTA - Relação de confrontantes por intervalos):
        {texto_planta}

        DOCUMENTO 2 (TABELA DE ROTEIRO PERIMÉTRICO):
        {texto_roteiro}

        Sua tarefa é extrair os dados cadastrais solicitados e criar as regras matemáticas de transição de confrontantes.
        
        INSTRUÇÕES CRÍTICAS:
        1. Extraia EXATAMENTE como aparecem nos documentos: proprietário, município, comarca, área total e perímetro total
        2. Para cada confrontante, determine o intervalo de pontos (ponto_inicio e ponto_fim)
        3. Exemplo: Se do ponto 7 ao 21 confronta com 'ES 230', crie: ponto_inicio: 7, ponto_fim: 21, nome_confrontante: 'ES 230'
        4. Se houver fechamento do ciclo (ex: de ponto 21 para 1), use: ponto_inicio: 21, ponto_fim: 1
        5. Retorne ESTRITAMENTE no formato JSON estruturado fornecido
        6. NÃO invente dados. Se não conseguir extrair um campo, deixe como string vazia ""
        """

        logger.info("Chamando API Gemini 3.5 Flash para mapeamento de confrontantes...")
        
        model = genai.GenerativeModel(
            model_name="gemini-3.5-flash",
            generation_config=types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=MapeamentoConfrontantes,
                temperature=0.0,  # Zero especulação, estritamente factual
            ),
        )
        
        response = model.generate_content(prompt)
        
        logger.info("Resposta recebida da API Gemini")
        
        # Parse da resposta JSON
        try:
            response_data = json.loads(response.text)
            mapeamento = MapeamentoConfrontantes(**response_data)
            logger.info(f"Mapeamento extraído: {len(mapeamento.regras)} regras de confrontantes")
            return mapeamento
            
        except ValidationError as e:
            logger.error(f"Erro de validação ao processar resposta do Gemini: {str(e)}")
            raise ValueError(f"Resposta da IA inválida: {str(e)}")
        except json.JSONDecodeError as e:
            logger.error(f"Erro ao fazer parse JSON da resposta: {str(e)}")
            raise ValueError(f"Resposta da IA não é JSON válido: {str(e)}")
        
    except Exception as e:
        logger.error(f"Erro ao mapear confrontantes com Gemini: {str(e)}")
        raise


# ==========================================
# LÓGICA DE VINCULAÇÃO DE CONFRONTANTES
# ==========================================
def vincular_confrontantes(segmentos: List[Dict], mapeamento: MapeamentoConfrontantes) -> List[Dict]:
    """
    Vincula os confrontantes mapeados pela IA aos segmentos reais da tabela.
    Implementa lógica corrigida para tratamento de faixas regulares e ciclos fechados.
    
    Args:
        segmentos: Lista de segmentos extraídos da tabela
        mapeamento: Mapeamento de confrontantes da IA
        
    Returns:
        List[Dict]: Segmentos com confrontantes vinculados
    """
    logger.info("Iniciando vinculação de confrontantes aos segmentos...")
    
    for seg in segmentos:
        try:
            v_de = int(seg["de"])
            v_para = int(seg["para"])
            
            confrontante_encontrado = None
            
            # Itera pelas regras de confrontantes
            for regra in mapeamento.regras:
                # Tratamento para faixas regulares (ponto_inicio < ponto_fim)
                if regra.ponto_inicio < regra.ponto_fim:
                    # Verifica se v_de está dentro do intervalo
                    if regra.ponto_inicio <= v_de < regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Faixa regular encontrada: {confrontante_encontrado}")
                        break
                
                # Tratamento para fechamento final do ciclo perimétrico (ponto_inicio > ponto_fim)
                # Ex: de 21 para 1 (ciclo fechado)
                elif regra.ponto_inicio > regra.ponto_fim:
                    if v_de >= regra.ponto_inicio or v_de <= regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Ciclo fechado encontrado: {confrontante_encontrado}")
                        break
            
            # Se não encontrou confrontante, usa padrão
            if not confrontante_encontrado:
                confrontante_encontrado = "CONFRONTAÇÃO NÃO ENCONTRADA"
                logger.warning(f"Segmento {v_de}→{v_para}: Nenhuma regra correspondente encontrada")
            
            seg["confrontante"] = confrontante_encontrado
            
        except ValueError as e:
            logger.error(f"Erro ao converter vértices para inteiro: {str(e)}")
            seg["confrontante"] = "ERRO NA CONVERSÃO"
        except Exception as e:
            logger.error(f"Erro ao vincular confrontante ao segmento {seg}: {str(e)}")
            seg["confrontante"] = "ERRO NO PROCESSAMENTO"
    
    logger.info("Vinculação de confrontantes concluída")
    return segmentos


# ==========================================
# GERADOR DO DOCUMENTO WORD
# ==========================================
def gerar_documento_word(dados_finais: Dict[str, Any]) -> io.BytesIO:
    """
    Gera o arquivo Word (.docx) com o memorial descritivo.
    Aplica o encadeamento correto de coordenadas e formatação profissional.
    
    Args:
        dados_finais: Dicionário com todos os dados finais processados
        
    Returns:
        io.BytesIO: Buffer com o arquivo Word gerado
    """
    try:
        logger.info("Iniciando geração do documento Word...")
        
        doc = docx.Document()

        # ========== MARGENS ==========
        for section in doc.sections:
            section.top_margin = Cm(MARGENS_CM)
            section.bottom_margin = Cm(MARGENS_CM)
            section.left_margin = Cm(MARGENS_CM)
            section.right_margin = Cm(MARGENS_CM)

        # ========== ESTILO PADRÃO ==========
        style = doc.styles["Normal"]
        font = style.font
        font.name = FONTE_PADRAO
        font.size = Pt(TAMANHO_FONTE_PADRAO)

        # ========== CABEÇALHO DA EMPRESA ==========
        p_empresa = doc.add_paragraph()
        p_empresa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_empresa.paragraph_format.space_after = Pt(18)
        run_emp = p_empresa.add_run(
            f"{EMPRESA_INFO['nome']}\n"
            f"{EMPRESA_INFO['endereco']}\n"
            f"Fone {EMPRESA_INFO['telefone']} - {EMPRESA_INFO['email']}"
        )
        run_emp.font.size = Pt(9)
        run_emp.italic = True

        # ========== LINHA SEPARADORA ==========
        p_linha = doc.add_paragraph()
        p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linha.add_run("_" * 80)

        # ========== TÍTULO ==========
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(12)
        p_titulo.paragraph_format.space_after = Pt(18)
        run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
        run_tit.bold = True
        run_tit.font.size = Pt(12)

        # ========== DADOS CADASTRAIS ==========
        p_dados = doc.add_paragraph()
        p_dados.paragraph_format.line_spacing = 1.15
        p_dados.paragraph_format.space_after = Pt(18)

        # Validação e tratamento de dados
        proprietario = dados_finais.get('proprietario', '').upper() or 'NÃO INFORMADO'
        municipio = dados_finais.get('municipio', '').upper() or 'NÃO INFORMADO'
        comarca = dados_finais.get('comarca', '').upper() or 'NÃO INFORMADO'
        area = dados_finais.get('area', 'NÃO INFORMADO')
        perimetro = dados_finais.get('perimetro', 'NÃO INFORMADO')

        p_dados.add_run("Imóvel: ").bold = True
        p_dados.add_run("GLEBA A\n")
        p_dados.add_run("Proprietário: ").bold = True
        p_dados.add_run(f"{proprietario}\n")
        p_dados.add_run("Município: ").bold = True
        p_dados.add_run(f"{municipio}\n")
        p_dados.add_run("Comarca: ").bold = True
        p_dados.add_run(f"{comarca}\n")
        p_dados.add_run("Área: ").bold = True
        p_dados.add_run(f"{area}\n")
        p_dados.add_run("Perímetro: ").bold = True
        p_dados.add_run(f"{perimetro}")

        # ========== TÍTULO DA DESCRIÇÃO ==========
        p_desc_tit = doc.add_paragraph()
        p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_desc_tit.paragraph_format.space_before = Pt(12)
        p_desc_tit.paragraph_format.space_after = Pt(12)
        run_desc = p_desc_tit.add_run("DESCRIÇÃO")
        run_desc.bold = True

        # ========== CORPO DO MEMORIAL ==========
        p_texto = doc.add_paragraph()
        p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_texto.paragraph_format.line_spacing = 1.25
        p_texto.paragraph_format.space_after = Pt(12)

        segmentos = dados_finais.get("segmentos", [])
        
        if segmentos:
            primeiro = segmentos[0]
            p_texto.add_run(
                f"Inicia-se a descrição deste perímetro no vértice {primeiro['de']}, "
                f"de coordenadas N {primeiro['n_y']} e E {primeiro['e_x']}; "
            )

            # Loop de Amarração Geográfica Correta
            for i, s in enumerate(segmentos):
                # Determina as coordenadas do próximo vértice
                if i + 1 < len(segmentos):
                    prox_coordenada_n = segmentos[i + 1]['n_y']
                    prox_coordenada_e = segmentos[i + 1]['e_x']
                else:
                    # Retorna ao marco inicial fechando o polígono
                    prox_coordenada_n = segmentos[0]['n_y']
                    prox_coordenada_e = segmentos[0]['e_x']

                confrontante = s.get('confrontante', 'NÃO INFORMADO')
                azimute = s.get('azimute', 'NÃO INFORMADO')
                distancia = s.get('distancia', 'NÃO INFORMADO')

                p_texto.add_run(
                    f"deste, segue confrontando com {confrontante}, "
                    f"com os seguintes azimutes e distâncias: {azimute} e {distancia} "
                    f"até o vértice {s['para']}, de coordenadas N {prox_coordenada_n} e E {prox_coordenada_e}; "
                )
        else:
            logger.warning("Nenhum segmento disponível para gerar descrição")
            p_texto.add_run("Nenhum segmento foi processado.")

        # ========== FECHAMENTO DO MEMORIAL ==========
        p_texto.add_run(
            "ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas "
            "estão georreferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas "
            "no Sistema UTM, referenciadas ao Meridiano Central nº 39° WGr, tendo como datum o SIRGAS2000. "
            "Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
        )

        # ========== DATA ATUAL ==========
        meses_pt = {
            1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
            7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
        }
        data_atual = datetime.now()
        nome_mes = meses_pt[data_atual.month]

        p_data = doc.add_paragraph()
        p_data.paragraph_format.space_before = Pt(24)
        p_data.add_run(f"Vila Valério, {data_atual.day} de {nome_mes} de {data_atual.year}")

        # ========== BLOCO DE ASSINATURA ==========
        p_assinatura = doc.add_paragraph()
        p_assinatura.paragraph_format.space_before = Pt(36)
        p_assinatura.add_run(
            f"__________________________________________________\n"
            f"{TECNICO_INFO['nome']}\n"
            f"{TECNICO_INFO['cargo']}\n"
            f"CFTA: {TECNICO_INFO['cfta']}\n"
            f"TRT: {TECNICO_INFO['trt']}"
        )

        # ========== SALVAR EM BUFFER ==========
        conteudo_arquivo = io.BytesIO()
        doc.save(conteudo_arquivo)
        conteudo_arquivo.seek(0)
        
        logger.info("Documento Word gerado com sucesso")
        return conteudo_arquivo
        
    except Exception as e:
        logger.error(f"Erro ao gerar documento Word: {str(e)}")
        raise


# ==========================================
# INTERFACE STREAMLIT
# ==========================================
def main():
    """Função principal da aplicação Streamlit"""
    
    st.title("📄 Processador de Memoriais Descritivos - Gleba A")
    st.write(
        "Insira os dois arquivos da Gleba A para estruturar automaticamente o Memorial Descritivo "
        "com precisão e conformidade técnica."
    )

    # ========== SIDEBAR COM CONFIGURAÇÕES ==========
    with st.sidebar:
        st.header("⚙️ Configurações")
        
        st.subheader("Dados da Empresa")
        empresa_nome = st.text_input("Nome da Empresa", value=EMPRESA_INFO["nome"])
        empresa_email = st.text_input("Email", value=EMPRESA_INFO["email"])
        
        st.subheader("Dados do Técnico Responsável")
        tecnico_nome = st.text_input("Nome do Técnico", value=TECNICO_INFO["nome"])
        tecnico_cfta = st.text_input("CFTA", value=TECNICO_INFO["cfta"])
        
        st.info(
            "💡 **Dica:** Modifique os dados acima se necessário. "
            "Eles serão usados em todos os documentos gerados nesta sessão."
        )

    # ========== UPLOAD DE ARQUIVOS ==========
    st.subheader("📁 Carregue os Arquivos")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_planta = st.file_uploader(
            "Carregue o PDF com os DADOS DA PLANTA:",
            type=["pdf"],
            key="planta",
            help="PDF contendo a relação de confrontantes por intervalos de pontos"
        )
    with col2:
        pdf_roteiro = st.file_uploader(
            "Carregue o PDF da TABELA DE ROTEIRO PERIMÉTRICO:",
            type=["pdf"],
            key="roteiro",
            help="PDF contendo a tabela com coordenadas, azimutes e distâncias"
        )

    # ========== PROCESSAMENTO ==========
    if pdf_planta and pdf_roteiro:
        if st.button("🔄 Analisar Documentos e Gerar Memorial", type="primary", use_container_width=True):
            
            # Atualizar configurações
            EMPRESA_INFO["nome"] = empresa_nome
            EMPRESA_INFO["email"] = empresa_email
            TECNICO_INFO["nome"] = tecnico_nome
            TECNICO_INFO["cfta"] = tecnico_cfta
            
            with st.spinner("⏳ Processando documentos..."):
                try:
                    # ========== ETAPA 1: EXTRAÇÃO DE TEXTOS ==========
                    st.info("📖 Etapa 1: Extraindo texto dos PDFs...")
                    texto_planta = extrair_texto_pdf(pdf_planta)
                    texto_roteiro = extrair_texto_pdf(pdf_roteiro)
                    st.success("✅ Textos extraídos com sucesso")

                    # ========== ETAPA 2: PARSING DA TABELA ==========
                    st.info("📊 Etapa 2: Analisando tabela de roteiro...")
                    segmentos_reais = parse_tabela_roteiro(texto_roteiro)
                    
                    if not segmentos_reais:
                        st.warning(
                            "⚠️ Nenhum segmento foi extraído da tabela. "
                            "Verifique o formato do PDF ou a qualidade da imagem."
                        )
                        st.stop()
                    
                    st.success(f"✅ {len(segmentos_reais)} segmentos extraídos")

                    # ========== ETAPA 3: CONFIGURAÇÃO DO GEMINI ==========
                    st.info("🔑 Etapa 3: Configurando API Gemini...")
                    if not configurar_gemini():
                        st.error(
                            "❌ Erro: Chave GEMINI_API_KEY não configurada nos Streamlit Secrets. "
                            "Configure a chave e tente novamente."
                        )
                        st.stop()
                    st.success("✅ Gemini configurado")

                    # ========== ETAPA 4: MAPEAMENTO COM IA ==========
                    st.info("🤖 Etapa 4: Mapeando confrontantes com IA...")
                    mapeamento = mapear_confrontantes_gemini(texto_planta, texto_roteiro)
                    st.success(f"✅ {len(mapeamento.regras)} regras de confrontantes extraídas")

                    # ========== ETAPA 5: VINCULAÇÃO ==========
                    st.info("🔗 Etapa 5: Vinculando confrontantes aos segmentos...")
                    segmentos_reais = vincular_confrontantes(segmentos_reais, mapeamento)
                    st.success("✅ Confrontantes vinculados")

                    # ========== ETAPA 6: UNIFICAÇÃO DE DADOS ==========
                    dados_finais = {
                        "proprietario": mapeamento.proprietario,
                        "municipio": mapeamento.municipio,
                        "comarca": mapeamento.comarca,
                        "area": mapeamento.area,
                        "perimetro": mapeamento.perimetro,
                        "segmentos": segmentos_reais
                    }

                    # ========== RESUMO DE VALIDAÇÃO ==========
                    st.success("🎉 Processamento concluído com sucesso!")
                    
                    st.subheader("🔍 Resumo de Validação")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Proprietário", dados_finais['proprietario'][:30] + "...")
                    with col2:
                        st.metric("Área Total", dados_finais['area'])
                    with col3:
                        st.metric("Perímetro", dados_finais['perimetro'])

                    # ========== TABELA DE CONFRONTAÇÕES ==========
                    with st.expander("📋 Clique para conferir a malha de confrontações vinculadas"):
                        df_data = []
                        for seg in dados_finais["segmentos"]:
                            df_data.append({
                                "De": seg['de'],
                                "Para": seg['para'],
                                "Azimute": seg['azimute'],
                                "Distância": seg['distancia'],
                                "Confrontante": seg['confrontante']
                            })
                        
                        import pandas as pd
                        df = pd.DataFrame(df_data)
                        st.dataframe(df, use_container_width=True, hide_index=True)

                    # ========== GERAÇÃO DO DOCUMENTO ==========
                    st.info("📝 Gerando documento Word...")
                    arquivo_docx = gerar_documento_word(dados_finais)
                    st.success("✅ Documento gerado com sucesso!")

                    # ========== DOWNLOAD ==========
                    st.download_button(
                        label="📥 Baixar Memorial Descritivo (.docx)",
                        data=arquivo_docx,
                        file_name=f"MEMORIAL_DESCRITIVO_GLEBA_A_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

                except ValueError as e:
                    st.error(f"❌ Erro de Validação: {str(e)}")
                    logger.error(f"Erro de validação: {str(e)}")
                    
                except json.JSONDecodeError as e:
                    st.error(f"❌ Erro ao processar resposta da IA: {str(e)}")
                    logger.error(f"Erro JSON: {str(e)}")
                    
                except Exception as e:
                    st.error(f"❌ Erro inesperado: {str(e)}")
                    logger.error(f"Erro geral: {str(e)}", exc_info=True)
                    
                    # Mostrar mais detalhes em modo debug
                    with st.expander("🔧 Detalhes Técnicos (Debug)"):
                        import traceback
                        st.code(traceback.format_exc())

    else:
        st.info("👆 Carregue ambos os arquivos PDF para começar o processamento")


if __name__ == "__main__":
    main()
