import os
import re
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io
import streamlit as st

# Configuração da página do Streamlit
st.set_page_config(page_title="Gerador de Memorial Descritivo", page_icon="📄", layout="wide")

# ==========================================
# 1. MAPEAMENTO FIXO DOS CONFRONTANTES (PLANTA)
# ==========================================
def obter_confrontantes_planta():
    return [
        {"de": 1, "para": 2, "nome": "WILSON SCHULTAIS"},
        {"de": 2, "para": 3, "nome": "WILSON SCHULTAIS"},
        {"de": 3, "para": 4, "nome": "ROSIMERE CAMPOS DELORTO LAGASSI"},
        {"de": 4, "para": 5, "nome": "ROSIMERE CAMPOS DELORTO LAGASSI"},
        {"de": 5, "para": 6, "nome": "ROSIMERE CAMPOS DELORTO LAGASSI"},
        {"de": 6, "para": 7, "nome": "ESTRADA MUNICIPAL"},
        {"de": 7, "para": 8, "nome": "ESTRADA MUNICIPAL"},
        {"de": 8, "para": 9, "nome": "GILMAR CAMPOS DELORTO"},
        {"de": 9, "para": 10, "nome": "GILMAR CAMPOS DELORTO"},
        {"de": 10, "para": 1, "nome": "ESTRADA MUNICIPAL"}
    ]

# ==========================================
# 2. PROCESSADOR DO TEXTO COLADO
# ==========================================
def processar_texto_copiado(texto_bruto):
    regras = obter_confrontantes_planta()
    segmentos = []
    
    linhas = texto_bruto.split('\n')
    
    for i, linha in enumerate(linhas):
        linha = linha.strip()
        
        # Identifica o padrão de início com dois números (vértices)
        match_vertices = re.match(r'^["\s]*(\d+)\s+["\s]*(\d+)', linha)
        if match_vertices:
            de = int(match_vertices.group(1))
            para = int(match_vertices.group(2))
            
            # Captura números no formato de coordenadas topográficas
            coordenadas = re.findall(r'(\d[\d\.]*,\d{2})', linha)
            
            # Se os dados quebraram de linha na cópia, junta com a linha de baixo
            if len(coordenadas) < 2 and (i + 1) < len(linhas):
                linha_seguinte = linhas[i+1].strip()
                coordenadas += re.findall(r'(\d[\d\.]*,\d{2})', linha_seguinte)
                linha = linha + " " + linha_seguinte
                
            ny, nx = "", ""
            for coord in coordenadas:
                if coord.startswith('7'):
                    ny = coord
                elif coord.startswith('3'):
                    nx = coord
            
            # Limpa e formata o Azimute (remove chaves e traduz LaTeX)
            match_az = re.search(r'\$(.*?)\$', linha)
            azimute = ""
            if match_az:
                azimute = match_az.group(1)
                azimute = azimute.replace(r'\circ', '°').replace(r'\prime\prime', '"').replace(r'\prime', "'")
                azimute = azimute.replace('{', '').replace('}', '')
                azimute = re.sub(r'\s+', '', azimute)
            else:
                match_az_txt = re.search(r'(\d+°\d+[\'\"]\d+[\'\"]|\d+°\d+\'\d+\")', linha)
                if match_az_txt:
                    azimute = match_az_txt.group(1)
            
            # REMOÇÃO EXTRA DE CARACTERES POLUENTES
            if azimute:
                azimute = azimute.replace('^', '').replace('"', '').replace("'", "'")
                azimute = re.sub(r'\s+', '', azimute)
            
            # Captura a Distância
            match_dist = re.search(r'(\d+,\d+|\d+\.\d+)\s*m', linha, re.IGNORECASE)
            distancia = match_dist.group(1) if match_dist else ""
            if distancia:
                distancia = distancia.replace('.', ',')
            
            if ny and azimute and distancia:
                nome_confrontante = "ESTRADA MUNICIPAL"
                for conf in regras:
                    if conf["de"] == de and conf["para"] == para:
                        nome_confrontante = conf["nome"]
                        break
                
                segmentos.append({
                    "de": de, "para": para, "y": ny, "x": nx if nx else "361.359,84",
                    "azimute": azimute, "distancia": distancia,
                    "confrontante": nome_confrontante
                })
                
    segmentos.sort(key=lambda x: x['de'])
    return segmentos

# ==========================================
# 3. GERADOR DO DOCUMENTO EM MEMÓRIA
# ==========================================
def gerar_documento_word(segmentos, nome_proprietario, nome_municipio, area_total, perimetro_total):
    doc = docx.Document()
    
    # Configura o padrão de fonte do documento (Arial 11)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    # Cabeçalho Centralizado
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tit = p_titulo.add_run("MEMORIAL DESCRITIVO")
    run_tit.bold = True
    run_tit.font.size = Pt(12)
    
    # Dados Cadastrais da Propriedade
    p_dados = doc.add_paragraph()
    p_dados.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_dados.paragraph_format.line_spacing = 1.15
    
    p_dados.add_run("Proprietário: ").bold = True
    p_dados.add_run(f"{nome_proprietario}\n")
    p_dados.add_run("Município: ").bold = True
    p_dados.add_run(f"{nome_municipio}\n")
    p_dados.add_run("Perímetro: ").bold = True
    p_dados.add_run(f"{perimetro_total}\n")
    p_dados.add_run("Área: ").bold = True
    p_dados.add_run(f"{area_total}")
    
    # Título da Descrição
    p_desc_tit = doc.add_paragraph()
    p_desc_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_desc_tit.add_run("\nDESCRIÇÃO").bold = True
    
    # Texto Corrido do Perímetro
    p_texto = doc.add_paragraph()
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.paragraph_format.line_spacing = 1.15
    
    trechos_texto = []
    for s in segmentos:
        texto_ponto = (
            f"do ponto {s['de']} (N(Y): {s['y']} m; E(X): {s['x']} m) segue na direção {s['azimute']}, "
            f"percorrendo uma distância de {s['distancia']} m até o ponto {s['para']}, "
            f"confrontando com {s['confrontante']}"
        )
        trechos_texto.append(texto_ponto)
        
    texto_final = "Do " + ", do ".join(trechos_texto) + "."
    p_texto.add_run(texto_final)
    
    # Observações Técnicas Padrão
    p_obs_tit = doc.add_paragraph()
    p_obs_tit.add_run("\nObservações:").bold = True
    
    p_obs_txt = doc.add_paragraph()
    p_obs_txt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_obs_txt.add_run(
        "O presente memorial descritivo foi elaborado conforme as normas técnicas vigentes e "
        "representa fielmente os limites do imóvel conforme levantamento topográfico realizado."
    )
    
    # Data em Português
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    data_atual = datetime.now()
    nome_mes_pt = meses_pt[data_atual.month]
    
    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_data.add_run(f"\nVila Valério, {data_atual.day} de {nome_mes_pt} de {data_atual.year}")
    
    # Bloco de Assinatura
    p_assinatura = doc.add_paragraph()
    p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_assinatura.add_run("\n______________________________\n").bold = True
    p_assinatura.add_run("Régis Campo da Silva\n").bold = True
    p_assinatura.add_run("Resp. Técnico\n")
    p_assinatura.add_run("CFTA: 11198519711")
    
    # Salva o documento em um objeto BytesIO (na memória RAM)
    conteudo_arquivo = io.BytesIO()
    doc.save(conteudo_arquivo)
    conteudo_arquivo.seek(0)
    return conteudo_arquivo

# ==========================================
# 4. INTERFACE WEB (STREAMLIT)
# ==========================================
st.title("📄 Gerador de Memorial Descritivo Estruturado")
st.write("Preencha os dados abaixo e cole a tabela do roteiro perimétrico para gerar o arquivo Word.")

# Organização dos campos em colunas
col1, col2 = st.columns(2)
with col1:
    txt_proprietario = st.text_input('Proprietário:', value='RONIVON CAMPOS DELORTO')
    txt_perimetro = st.text_input('Perímetro:', value='1.199,78 m')
with col2:
    txt_municipio = st.text_input('Município:', value='VILA VALÉRIO')
    txt_area = st.text_input('Área Total:', value='47.863,57 m²')

caixa_texto = st.text_area(
    'Texto Tabela:', 
    placeholder='Cole as linhas de texto da tabela do roteiro perimétrico aqui...',
    height=200
)

# Botão de processamento
if st.button('Processar Dados do Memorial', type='primary'):
    texto_inserido = caixa_texto.strip()
    
    if not texto_inserido:
        st.error("❌ Erro: O campo com o texto da tabela está vazio!")
    else:
        with st.spinner("⏳ Processando dados estruturados..."):
            try:
                segmentos = processar_texto_copiado(texto_inserido)
                if not segmentos:
                    st.warning("⚠️ Erro: Não foi possível identificar os vértices no texto colado.")
                else:
                    # Gera o arquivo na memória
                    arquivo_word = gerar_documento_word(
                        segmentos, 
                        txt_proprietario.strip(), 
                        txt_municipio.strip(), 
                        txt_area.strip(), 
                        txt_perimetro.strip()
                    )
                    
                    st.success("🎉 Concluído com sucesso! O documento foi gerado.")
                    
                    # Nome dinâmico para o download
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nome_slug = re.sub(r'[\\/*?:"<>| ]', '_', txt_proprietario.upper()[:20])
                    nome_arquivo = f"MEMORIAL_{nome_slug}_{timestamp}.docx"
                    
                    # Cria o botão oficial de baixar o arquivo no navegador
                    st.download_button(
                        label="📥 Baixar Arquivo Word (.docx)",
                        data=arquivo_word,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"❌ Erro crítico no processamento: {str(e)}")
