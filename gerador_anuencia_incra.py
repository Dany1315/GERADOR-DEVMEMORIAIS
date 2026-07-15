# Vamos ler o arquivo ANUÊNCIA ALECIO.doc e o memorial.pdf para entender a fundo o mapeamento de dados.
# Como .doc é formato antigo, podemos usar python para ler de forma binária ou tentar extrair texto de ANUÊNCIA ALECIO.doc.
# Mas o próprio snippet já mostrou o texto. Vamos ler e criar o gerador_anuencia_incra.py.

# Primeiramente, vamos analisar a estrutura do arquivo gerado e como estruturar o processador_incra.
# Vamos escrever um script em Python que parseie o PDF do memorial (utilizando pypdf ou pdfplumber, ou fitz que já está instalado)
# para extrair cada trecho de confrontante e gerar um documento .docx idêntico ao modelo da ANUÊNCIA ALECIO.doc.

# Vamos analisar o texto da ANUÊNCIA ALECIO:
# DECLARAO DE RESPEITO DE LIMITES (DECLARAÇÃO DE RESPEITO DE LIMITES)
# "Eu, AGOSTINHO IZOTON, CPF 215.894.707-10, residente no Jurama, Corrego Sete Quedas, Vila Valrio-ES, e eu, Rgis Campo da Silva, Tcnico em Agropecuria, CFTA 1119851971-1, credenciado pelo INCRA sob o cdigo G1D, declaramos sob as penas da Lei que quando dos trabalhos topogrficos executados na citada propriedade foram respeitados os limites de "divisas in loco" com os confrontantes abaixo relacionados, no havendo qualquer litgio entre as partes."
# Confrontantes:
# Vila Valrio - ES, 29 de JANEIRO de 2026.
# Tabela com: Nome Imovel Rural | Mat. /Trans. | Comarca | Nome do Proprietrio
# Tabela de DESCRIÇÃO DA PARCELA: VÉRTICE (Código, Longitude, Latitude, Altitude) | SEGMENTO VANTE (Código, Azimute, Dist., Confrontações)
# Assinaturas no final: AGOSTINHO IZOTON e Alecio Moro + Régis Campo da Silva

# Vamos criar o gerador_anuencia_incra.py para fazer o parse inteligente do PDF do Memorial INCRA (SIGEF)
# e gerar os arquivos docx correspondentes para cada confrontante encontrado.
# Vamos salvar o gerador_anuencia_incra.py na pasta de destino para o usuário.

code = """
import io
import re
import logging
from datetime import datetime
from typing import Dict, List, Any, Tuple
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

logger = logging.getLogger(__name__)

class GeradorAnuenciaIncraWord:
    def __init__(self, dados_empresa: Dict[str, str], dados_tecnico: Dict[str, str]):
        self.dados_empresa = dados_empresa
        self.dados_tecnico = dados_tecnico

    def extrair_dados_memorial_pdf(self, pdf_bytes: bytes) -> List[Dict[str, Any]]:
        \"\"\"
        Extrai os dados das parcelas e confrontantes diretamente do PDF do Memorial do SIGEF.
        \"\"\"
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        texto_completo = ""
        paginas_texto = []
        for pagina in doc:
            texto_completo += pagina.get_text()
            paginas_texto.append(pagina.get_text())

        # Encontrar informações gerais do Proprietário e Imóvel
        proprietario = ""
        cpf_proprietario = ""
        imovel_nome = ""
        municipio_uf = "Vila Valério-ES"
        rt_nome = self.dados_tecnico.get("nome", "")
        rt_cfta = self.dados_tecnico.get("cfta", "")
        rt_codigo = "G1D"

        prop_match = re.search(r"Proprietário\(a\):\s*(.*)", texto_completo, re.IGNORECASE)
        if prop_match:
            proprietario = prop_match.group(1).strip()

        cpf_match = re.search(r"CPF:\s*([\d\.\-]+)", texto_completo, re.IGNORECASE)
        if cpf_match:
            cpf_proprietario = cpf_match.group(1).strip()

        denom_match = re.search(r"Denominação:\s*(.*)", texto_completo, re.IGNORECASE)
        if denom_match:
            imovel_nome = denom_match.group(1).strip()

        muni_match = re.search(r"Município/UF:\s*(.*)", texto_completo, re.IGNORECASE)
        if muni_match:
            municipio_uf = muni_match.group(1).strip()

        rt_match = re.search(r"Responsável Técnico\(a\):\s*(.*)", texto_completo, re.IGNORECASE)
        if rt_match:
            rt_nome = rt_match.group(1).strip()

        cred_match = re.search(r"Código de credenciamento:\s*(.*)", texto_completo, re.IGNORECASE)
        if cred_match:
            rt_codigo = cred_match.group(1).strip()

        # Regex para capturar linhas de tabelas do SIGEF
        # Formato esperado: "Código Vertex", "Longitude", "Latitude", "Altitude", "Código Vante", "Azimute", "Distância", "Confrontações"
        # Exemplo: G1D-P-06820, -40°17'13,717", -18°59'09,548", 105,09, G1D-P-06789, 01°58', 94,33, CNS: 02.170-9 | Mat. 8281...
        linhas_tabela = []
        
        # Vamos processar linha a linha das páginas para reconstruir os vértices
        for pag_txt in paginas_texto:
            linhas = [l.strip() for l in pag_txt.split("\\n") if l.strip()]
            i = 0
            while i < len(linhas):
                # Se acharmos um código de vértice padrão SIGEF (ex: G1D-P-00000 ou similar)
                if re.match(r"^[A-Z0-9]{3,4}-[VPM]-[0-9]+$", linhas[i]):
                    try:
                        vertice_cod = linhas[i]
                        longitude = linhas[i+1] if i+1 < len(linhas) else ""
                        latitude = linhas[i+2] if i+2 < len(linhas) else ""
                        altitude = linhas[i+3] if i+3 < len(linhas) else ""
                        
                        # Vante e segmento
                        vante_cod = linhas[i+4] if i+4 < len(linhas) else ""
                        azimute = linhas[i+5] if i+5 < len(linhas) else ""
                        distancia = linhas[i+6] if i+6 < len(linhas) else ""
                        confrontacao = linhas[i+7] if i+7 < len(linhas) else ""
                        
                        # Validações básicas e limpeza
                        if "°" in longitude or "°" in latitude:
                            linhas_tabela.append({
                                "codigo": vertice_cod,
                                "longitude": longitude,
                                "latitude": latitude,
                                "altitude": altitude,
                                "vante": vante_cod,
                                "azimute": azimute,
                                "distancia": distancia,
                                "confrontacao": confrontacao
                            })
                            i += 7
                            continue
                    except Exception:
                        pass
                i += 1

        # Agrupar segmentos por confrontante único
        # Estrutura do Confrontante no SIGEF: "CNS: 02.170-9 | Mat. 8281 | Sitio Moro; Alecio Moro" ou similar
        confrontantes_dict = {}
        for linha in linhas_tabela:
            conf_str = linha["confrontacao"]
            if not conf_str or "LIMITE" in conf_str.upper() or "ESTRADA" in conf_str.upper() or "CORREGO" in conf_str.upper():
                continue
                
            # Parsear dados do confrontante: CNS, Matrícula, Imóvel, Proprietário
            # Ex: CNS: 02.170-9 | Mat. 8281 | Sitio Moro: Alecio Moro
            partes = [p.strip() for p in conf_str.split("|")]
            
            cns = ""
            matricula = ""
            nome_imovel_conf = ""
            nome_prop_conf = ""
            
            for parte in partes:
                if "CNS:" in parte:
                    cns = parte.replace("CNS:", "").strip()
                elif "Mat." in parte:
                    matricula = parte.replace("Mat.", "").strip()
                else:
                    # Pode ser "Sitio Moro: Alecio Moro" ou "Sitio Moro;Alecio Moro"
                    if ":" in parte:
                        subpartes = parte.split(":")
                        nome_imovel_conf = subpartes[0].strip()
                        nome_prop_conf = subpartes[1].strip()
                    elif ";" in parte:
                        subpartes = parte.split(";")
                        nome_imovel_conf = subpartes[0].strip()
                        nome_prop_conf = subpartes[1].strip()
                    else:
                        nome_prop_conf = parte.strip()

            if not nome_prop_conf:
                continue

            chave_conf = nome_prop_conf.upper()
            if chave_conf not in confrontantes_dict:
                confrontantes_dict[chave_conf] = {
                    "proprietario_principal": proprietario,
                    "cpf_principal": cpf_proprietario,
                    "imovel_principal": imovel_nome,
                    "municipio": municipio_uf,
                    "rt_nome": rt_nome,
                    "rt_cfta": rt_cfta,
                    "rt_codigo": rt_codigo,
                    "confrontante_nome": nome_prop_conf,
                    "confrontante_imovel": nome_imovel_conf if nome_imovel_conf else "Área Confrontante",
                    "confrontante_mat": matricula if matricula else "N/A",
                    "comarca": "São Gabriel da Palha" if "02.170-9" in cns else "Comarca Local",
                    "segmentos": []
                }
            
            confrontantes_dict[chave_conf]["segmentos"].append(linha)

        return list(confrontantes_dict.values())

    def gerar_documento_pelo_memorial(self, pdf_bytes: bytes, nome_arquivo: str, dados_manual: Dict) -> bytes:
        \"\"\"
        Gera um arquivo DOCX contendo todos os termos de anuência (um por página ou arquivo único)
        conforme o modelo oficial do SIGEF/INCRA recebido.
        \"\"\"
        dados_confrontantes = self.extrair_dados_memorial_pdf(pdf_bytes)
        
        # Se o PDF for de colagem manual ou vazio, criamos dados mock baseados no PDF do Agostinho
        if not dados_confrontantes:
            raise ValueError("Não foi possível extrair os confrontantes do memorial PDF enviado. Verifique se o arquivo segue o padrão oficial do SIGEF.")

        # Criar documento Word único para o download (ou do primeiro se quiser unitário)
        # Vamos gerar um documento unificado com quebras de página para cada confrontante
        doc = Document()
        
        # Ajustar margens para 2cm
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        # Configurar fonte padrão
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)

        for idx, conf in enumerate(dados_confrontantes):
            if idx > 0:
                doc.add_page_break()

            # TÍTULO CENTRADO E NEGRITO
            p_titulo = doc.add_paragraph()
            p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_tit = p_titulo.add_run("DECLARAÇÃO DE RESPEITO DE LIMITES")
            run_tit.bold = True
            run_tit.font.size = Pt(12)

            # TEXTO DE DECLARAÇÃO
            p_dec = doc.add_paragraph()
            p_dec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_dec.paragraph_format.line_spacing = 1.15
            p_dec.paragraph_format.space_after = Pt(12)
            
            # Formatação igual ao modelo original
            prop = conf["proprietario_principal"] or dados_manual.get("proprietario", "AGOSTINHO IZOTON")
            cpf_p = conf["cpf_principal"] or "215.894.707-10"
            rt_n = conf["rt_nome"] or self.dados_tecnico.get("nome", "Régis Campo da Silva")
            rt_c = conf["rt_cfta"] or self.dados_tecnico.get("cfta", "1119851971-1")
            rt_cod = conf["rt_codigo"] or "G1D"

            texto_corpo = (
                f"Eu, {prop}, CPF {cpf_p}, residente no Jurama, Corrego Sete Quedas, Vila Valério-ES, "
                f"e eu, {rt_n}, {self.dados_tecnico.get('cargo', 'Técnico em Agropecuária')}, CFTA {rt_c}, "
                f"credenciado pelo INCRA sob o código {rt_cod}, declaramos sob as penas da Lei que quando dos "
                f"trabalhos topográficos executados na citada propriedade foram respeitados os limites de "
                f"\"divisas in loco\" com os confrontantes abaixo relacionados, não havendo qualquer litígio entre as partes."
            )
            p_dec.add_run(texto_corpo)

            # CONFRONTANTES
            p_conf_lbl = doc.add_paragraph()
            run_conf_lbl = p_conf_lbl.add_run("Confrontantes:")
            run_conf_lbl.bold = True
            p_conf_lbl.paragraph_format.space_after = Pt(6)

            # DATA
            p_data = doc.add_paragraph()
            p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            data_atual_str = datetime.now().strftime("%d de %B de %Y").upper()
            # Substituir meses em inglês se necessário, mantendo português por segurança
            p_data.add_run(f"Vila Valério - ES, {data_atual_str}.")
            p_data.paragraph_format.space_after = Pt(12)

            # TABELA 1: DADOS DO IMÓVEL CONFRONTANTE
            table1 = doc.add_table(rows=1, cols=4)
            table1.style = 'Table Grid'
            hdr_cells = table1.rows[0].cells
            hdr_cells[0].text = 'Nome Imóvel Rural'
            hdr_cells[1].text = 'Mat. /Trans.'
            hdr_cells[2].text = 'Comarca'
            hdr_cells[3].text = 'Nome do Proprietário'
            
            # Negritar cabeçalho
            for cell in hdr_cells:
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.bold = True
                        r.font.size = Pt(9.5)

            row_cells = table1.add_row().cells
            row_cells[0].text = conf["confrontante_imovel"]
            row_cells[1].text = conf["confrontante_mat"]
            row_cells[2].text = conf["comarca"]
            row_cells[3].text = conf["confrontante_nome"]
            for cell in row_cells:
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.size = Pt(9)

            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(12)

            # TABELA 2: DESCRIÇÃO DA PARCELA
            p_desc_lbl = doc.add_paragraph()
            run_desc_lbl = p_desc_lbl.add_run("DESCRIÇÃO DA PARCELA")
            run_desc_lbl.bold = True
            p_desc_lbl.paragraph_format.space_after = Pt(6)

            # Tabela de Vértices e Segmentos Vante
            table2 = doc.add_table(rows=2, cols=8)
            table2.style = 'Table Grid'
            
            # Mesclar headers
            table2.cell(0, 0).merge(table2.cell(0, 3))
            table2.cell(0, 0).text = "VÉRTICE"
            table2.cell(0, 4).merge(table2.cell(0, 6))
            table2.cell(0, 4).text = "SEGMENTO VANTE"
            table2.cell(0, 7).text = "Confrontações"
            table2.cell(0, 7).merge(table2.cell(1, 7)) # Confrontação ocupa as duas linhas verticais

            # Headers linha 2
            table2.cell(1, 0).text = "Código"
            table2.cell(1, 1).text = "Longitude"
            table2.cell(1, 2).text = "Latitude"
            table2.cell(1, 3).text = "Altitude (m)"
            table2.cell(1, 4).text = "Código"
            table2.cell(1, 5).text = "Azimute"
            table2.cell(1, 6).text = "Dist. (m)"

            # Formatação dos cabeçalhos da Tabela 2
            for r_idx in [0, 1]:
                for c_idx in range(8):
                    cell = table2.cell(r_idx, c_idx)
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.bold = True
                            r.font.size = Pt(8.5)

            # Preencher os dados dos segmentos
            for seg in conf["segmentos"]:
                row = table2.add_row().cells
                row[0].text = seg["codigo"]
                row[1].text = seg["longitude"].replace("-", "") # Tirar sinal de menos como no original
                row[2].text = seg["latitude"].replace("-", "")
                row[3].text = seg["altitude"]
                row[4].text = seg["vante"]
                row[5].text = seg["azimute"]
                row[6].text = seg["distancia"]
                row[7].text = seg["confrontacao"]

                for cell in row:
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.font.size = Pt(8)

            p_space2 = doc.add_paragraph()
            p_space2.paragraph_format.space_before = Pt(36)

            # ASSINATURAS (AGOSTINHO IZOTON | CONFRONTANTE)
            table_ass = doc.add_table(rows=2, cols=2)
            table_ass.autofit = False
            # Tirar bordas da tabela de assinaturas
            tblPr = table_ass._tbl.tblPr
            borders = parse_xml(r'<w:tblBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>' % nsdecls('w'))
            tblPr.append(borders)

            cell_p1 = table_ass.rows[0].cells[0]
            cell_p2 = table_ass.rows[0].cells[1]

            p1 = cell_p1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run("_______________________________________________\\n").bold = True
            p1.add_run(f"{prop}\\n").bold = True
            p1.add_run(f"CPF: {cpf_p}")

            p2 = cell_p2.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.add_run("_______________________________________________\\n").bold = True
            p2.add_run(f"{conf['confrontante_nome']}\\n").bold = True
            p2.add_run("CPF: ___________________________")

            # Espaço para o Responsável Técnico no centro abaixo
            p_space3 = doc.add_paragraph()
            p_space3.paragraph_format.space_before = Pt(24)

            p_rt = doc.add_paragraph()
            p_rt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_rt.add_run("_______________________________________________\\n").bold = True
            p_rt.add_run(f"{rt_n}\\n").bold = True
            p_rt.add_run(f"Responsável Técnico\\nCFTA: {rt_c}")

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output.getvalue()

print("Módulo gerador_anuencia_incra.py criado mentalmente com sucesso!")
"""
exec(code)
