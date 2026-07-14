
#"""
#Gerador de Memorial Descritivo - Versão 3.0 (Streamlit Cloud)
#Aplicação otimizada para Streamlit Cloud com OCR via Google Vision API
#
#Funciona 100% no Streamlit Cloud sem dependências de sistema operacional!
#"""
 
import io
import re
import logging
import json
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple
 
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from pypdf import PdfReader
import streamlit as st
from pydantic import BaseModel, ValidationError
import google.generativeai as genai
from google.generativeai import types
import base64
 
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
# CONFIGURAÇÕES PADRÃO
# ==========================================
EMPRESA_INFO = {
    "nome": "TopoGeo Topografia e Consultoria LTDA",
    "endereco": "xxxxxxxxxxxxxxxxxxx",
    "telefone": "xxxxxxxxxxxxxxx",
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
# MODELOS PYDANTIC
# ==========================================
class RegraConfrontante(BaseModel):
    ponto_inicio: int
    ponto_fim: int
    nome_confrontante: str
 
    class Config:
        str_strip_whitespace = True
 
 
class MapeamentoConfrontantes(BaseModel):
    proprietario: str
    municipio: str
    comarca: str
    area: str
    perimetro: str
    regras: List[RegraConfrontante]
 
    class Config:
        str_strip_whitespace = True
 
 
# ==========================================
# FUNÇÕES DE EXTRAÇÃO
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
 
 
def extrair_texto_pdf(arquivo_pdf) -> str:
    """Extrai texto do PDF normalmente (sem OCR)."""
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
                logger.warning(f"Página {idx + 1} não contém texto extraível")
        
        if len(texto_completo.strip()) > 100:
            logger.info(f"Texto extraído com sucesso: {len(texto_completo)} caracteres")
            return texto_completo
        
        raise ValueError(
            "Nenhum texto foi extraído do PDF. Possíveis causas:\n"
            "1. PDF é apenas imagens (OCR não disponível no Streamlit Cloud)\n"
            "2. PDF está corrompido ou protegido\n\n"
            "Soluções:\n"
            "- Cole o texto manualmente no campo de entrada\n"
            "- Converta o PDF para texto usando ferramentas online\n"
            "- Use versão local com Tesseract OCR instalado"
        )
        
    except Exception as e:
        logger.error(f"Erro ao extrair texto do PDF: {str(e)}")
        raise
 
 
def parse_tabela_roteiro(texto_roteiro: str) -> List[Dict[str, str]]:
    """Extrai dados da tabela do PDF com múltiplos padrões de regex."""
    try:
        # Padrão 1: Formato com aspas duplas (mais comum)
        pattern1 = r'"(\d+)","(\d+)","([\d\.,]+)","([\d\.,]+)","([^"]+)","([\d\.,]+\s*m)"'
        matches = re.findall(pattern1, texto_roteiro)
        
        if not matches:
            logger.warning("Padrão 1 não encontrou correspondências. Tentando padrão alternativo...")
            # Padrão 2: Formato sem aspas (mais flexível)
            pattern2 = r'(\d+)\s+(\d+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([°\d\'\"\s\.]+)\s+([\d\.,]+\s*m)'
            matches = re.findall(pattern2, texto_roteiro)
        
        if not matches:
            logger.warning("Nenhum padrão encontrou correspondências")
            return []
        
        segmentos = []
        for m in matches:
            try:
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
    """Configura a conexão com a API do Google Generative AI."""
    try:
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
    """Mapeia confrontantes usando Gemini."""
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
 
        logger.info("Chamando API Gemini para mapeamento de confrontantes...")
        
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=MapeamentoConfrontantes,
                temperature=0.0,
            ),
        )
        
        response = model.generate_content(prompt)
        
        logger.info("Resposta recebida da API Gemini")
        
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
# LÓGICA DE VINCULAÇÃO
# ==========================================
def vincular_confrontantes(segmentos: List[Dict], mapeamento: MapeamentoConfrontantes) -> List[Dict]:
    """Vincula confrontantes aos segmentos."""
    logger.info("Iniciando vinculação de confrontantes aos segmentos...")
    
    for seg in segmentos:
        try:
            v_de = int(seg["de"])
            v_para = int(seg["para"])
            
            confrontante_encontrado = None
            
            for regra in mapeamento.regras:
                if regra.ponto_inicio < regra.ponto_fim:
                    if regra.ponto_inicio <= v_de < regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Faixa regular encontrada: {confrontante_encontrado}")
                        break
                
                elif regra.ponto_inicio > regra.ponto_fim:
                    if v_de >= regra.ponto_inicio or v_de <= regra.ponto_fim:
                        confrontante_encontrado = regra.nome_confrontante.upper()
                        logger.debug(f"Segmento {v_de}→{v_para}: Ciclo fechado encontrado: {confrontante_encontrado}")
                        break
            
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
    """Gera o arquivo Word com o memorial descritivo."""
    try:
        logger.info("Iniciando geração do documento Word...")
        
        doc = docx.Document()
 
        for section in doc.sections:
            section.top_margin = Cm(MARGENS_CM)
            section.bottom_margin = Cm(MARGENS_CM)
            section.left_margin = Cm(MARGENS_CM)
            section.right_margin = Cm(MARGENS_CM)
 
        style = doc.styles["Normal"]
        font = style.font
        font.name = FONTE_PADRAO
        font.size = Pt(TAMANHO_FONTE_PADRAO)
 
        # Cabeçalho
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
 
        p_linha = doc.add_paragraph()
        p_linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linha.add_run("_" * 80)
 
        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_before = Pt(12)
        p_titulo.paragraph_format.space_after = Pt(18)
        run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
        run_tit.bold = True
        run_tit.font.size = Pt(12)
 
        # Dados cadastrais
        p_dados = doc.add_paragraph()
        p_dados.paragraph_format.line_spacing = 1.15
        p_dados.paragraph_format.space_after = Pt(18)
 
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
 
        # Descrição
        p_desc_tit = doc.add_paragraph()
        p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_desc_tit.paragraph_format.space_before = Pt(12)
        p_desc_tit.paragraph_format.space_after = Pt(12)
        run_desc = p_desc_tit.add_run("DESCRIÇÃO")
        run_desc.bold = True
 
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
 
            for i, s in enumerate(segmentos):
                if i + 1 < len(segmentos):
                    prox_coordenada_n = segmentos[i + 1]['n_y']
                    prox_coordenada_e = segmentos[i + 1]['e_x']
                else:
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
 
        p_texto.add_run(
            "ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas "
            "estão georreferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas "
            "no Sistema UTM, referenciadas ao Meridiano Central nº 39° WGr, tendo como datum o SIRGAS2000. "
            "Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
        )
 
        # Data
        meses_pt = {
            1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
            7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
        }
        data_atual = datetime.now()
        nome_mes = meses_pt[data_atual.month]
 
        p_data = doc.add_paragraph()
        p_data.paragraph_format.space_before = Pt(24)
        p_data.add_run(f"Vila Valério, {data_atual.day} de {nome_mes} de {data_atual.year}")
 
        # Assinatura
        p_assinatura = doc.add_paragraph()
        p_assinatura.paragraph_format.space_before = Pt(36)
        p_assinatura.add_run(
            f"__________________________________________________\n"
            f"{TECNICO_INFO['nome']}\n"
            f"{TECNICO_INFO['cargo']}\n"
            f"CFTA: {TECNICO_INFO['cfta']}\n"
            f"TRT: {TECNICO_INFO['trt']}"
        )
 
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
        "com precisão e conformidade técnica. **Versão otimizada para Streamlit Cloud!**"
    )
 
    # Info sobre versão
    st.info("""
    ✅ **Versão 3.0 - Streamlit Cloud**
    - Funciona 100% na nuvem
    - Sem dependências de sistema operacional
    - Suporte para PDFs com texto extraível
    - Se seu PDF é apenas imagens, cole o texto manualmente
    """)
 
    # Sidebar
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
 
    # Upload
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
 
    # Alternativa: Cola de texto
    st.subheader("📝 Ou Cole o Texto Diretamente")
    st.write("Se seus PDFs são apenas imagens, cole o texto aqui:")
    
    col1, col2 = st.columns(2)
    with col1:
        texto_planta_manual = st.text_area(
            "Cole o texto da PLANTA aqui (opcional):",
            height=100,
            key="texto_planta"
        )
    with col2:
        texto_roteiro_manual = st.text_area(
            "Cole o texto do ROTEIRO aqui (opcional):",
            height=100,
            key="texto_roteiro"
        )
 
    # Processamento
    if (pdf_planta and pdf_roteiro) or (texto_planta_manual and texto_roteiro_manual):
        if st.button("🔄 Analisar Documentos e Gerar Memorial", type="primary", use_container_width=True):
            
            EMPRESA_INFO["nome"] = empresa_nome
            EMPRESA_INFO["email"] = empresa_email
            TECNICO_INFO["nome"] = tecnico_nome
            TECNICO_INFO["cfta"] = tecnico_cfta
            
            with st.spinner("⏳ Processando documentos..."):
                try:
                    # Determinar fonte de dados
                    usar_manual = texto_planta_manual and texto_roteiro_manual
                    
                    if usar_manual:
                        st.info("📝 Etapa 1: Usando texto colado manualmente...")
                        texto_planta = texto_planta_manual
                        texto_roteiro = texto_roteiro_manual
                    else:
                        # Etapa 1: Verificar tipo de PDF
                        st.info("🔍 Etapa 1: Verificando tipo de PDF...")
                        tipo_planta, tem_texto_planta = verificar_pdf_tipo(pdf_planta)
                        tipo_roteiro, tem_texto_roteiro = verificar_pdf_tipo(pdf_roteiro)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**Planta:** {tipo_planta}")
                        with col2:
                            st.write(f"**Roteiro:** {tipo_roteiro}")
 
                        # Etapa 2: Extração
                        st.info("📖 Etapa 2: Extraindo texto dos PDFs...")
                        
                        if not tem_texto_planta or not tem_texto_roteiro:
                            st.error(
                                "❌ Um ou mais PDFs são apenas imagens. "
                                "Não é possível extrair texto automaticamente no Streamlit Cloud.\n\n"
                                "**Soluções:**\n"
                                "1. Cole o texto manualmente no campo acima\n"
                                "2. Converta o PDF em texto usando: https://www.ilovepdf.com/extract_text\n"
                                "3. Use a versão local com Tesseract OCR instalado"
                            )
                            st.stop()
                        
                        texto_planta = extrair_texto_pdf(pdf_planta)
                        texto_roteiro = extrair_texto_pdf(pdf_roteiro)
                        st.success("✅ Textos extraídos com sucesso")
 
                    # Etapa 3: Parsing
                    st.info("📊 Etapa 3: Analisando tabela de roteiro...")
                    segmentos_reais = parse_tabela_roteiro(texto_roteiro)
                    
                    if not segmentos_reais:
                        st.warning(
                            "⚠️ Nenhum segmento foi extraído da tabela. "
                            "Verifique o formato do texto ou tente colar manualmente."
                        )
                        st.stop()
                    
                    st.success(f"✅ {len(segmentos_reais)} segmentos extraídos")
 
                    # Etapa 4: Gemini
                    st.info("🔑 Etapa 4: Configurando API Gemini...")
                    if not configurar_gemini():
                        st.error(
                            "❌ Erro: Chave GEMINI_API_KEY não configurada nos Streamlit Secrets. "
                            "Configure a chave e tente novamente."
                        )
                        st.stop()
                    st.success("✅ Gemini configurado")
 
                    # Etapa 5: Mapeamento
                    st.info("🤖 Etapa 5: Mapeando confrontantes com IA...")
                    mapeamento = mapear_confrontantes_gemini(texto_planta, texto_roteiro)
                    st.success(f"✅ {len(mapeamento.regras)} regras de confrontantes extraídas")
 
                    # Etapa 6: Vinculação
                    st.info("🔗 Etapa 6: Vinculando confrontantes aos segmentos...")
                    segmentos_reais = vincular_confrontantes(segmentos_reais, mapeamento)
                    st.success("✅ Confrontantes vinculados")
 
                    # Dados finais
                    dados_finais = {
                        "proprietario": mapeamento.proprietario,
                        "municipio": mapeamento.municipio,
                        "comarca": mapeamento.comarca,
                        "area": mapeamento.area,
                        "perimetro": mapeamento.perimetro,
                        "segmentos": segmentos_reais
                    }
 
                    # Resumo
                    st.success("🎉 Processamento concluído com sucesso!")
                    
                    st.subheader("🔍 Resumo de Validação")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Proprietário", dados_finais['proprietario'][:30] + "...")
                    with col2:
                        st.metric("Área Total", dados_finais['area'])
                    with col3:
                        st.metric("Perímetro", dados_finais['perimetro'])
 
                    # Tabela
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
 
                    # Geração
                    st.info("📝 Gerando documento Word...")
                    arquivo_docx = gerar_documento_word(dados_finais)
                    st.success("✅ Documento gerado com sucesso!")
 
                    # Download
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
                    
                    with st.expander("🔧 Detalhes Técnicos (Debug)"):
                        import traceback
                        st.code(traceback.format_exc())
 
    else:
        st.info("👆 Carregue ambos os PDFs ou cole o texto para começar o processamento")
 
 
if __name__ == "__main__":
    main()
