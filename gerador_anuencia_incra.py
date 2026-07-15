import os
from google import genai
from google.genai import types

def gerar_anuencia_com_gemini(dados_memorial, texto_completo_memorial):
    """
    Envia os dados extraídos do memorial para a API do Gemini 
    gerar a carta de anuência de forma estruturada.
    """
    # Inicializa o cliente (ele busca a chave API automaticamente da variável de ambiente GEMINI_API_KEY)
    client = genai.Client()
    
    # Definimos a instrução de sistema para o modelo se comportar exatamente como queremos
    system_instruction = """
    Você é um assistente jurídico/topográfico especialista em regularização fundiária.
    Sua tarefa é gerar uma Carta de Anuência de Confrontante com base nas informações extraídas 
    de um Memorial Descritivo que o usuário fornecer.
    A anuência deve ser clara, formal e conter todos os dados técnicos necessários (proprietário, imóvel, confrontações).
    """
    
    # Montamos o prompt injetando os dados que o seu código já extraiu do documento
    prompt = f"""
    Com base nos dados extraídos abaixo e no texto completo do memorial descritivo, 
    gere a respectiva Carta de Anuência para os confrontantes.

    DADOS EXTRAÍDOS DO IMÓVEL:
    - Proprietário: {dados_memorial.get('proprietario', 'Não encontrado')}
    - Imóvel: {dados_memorial.get('imovel', 'Não encontrado')}
    - Localização: {dados_memorial.get('localizacao', 'Não encontrada')}
    - Área Total: {dados_memorial.get('area', 'Não encontrada')}
    - Perímetro: {dados_memorial.get('perimetro', 'Não encontrado')}

    CONFRONTANTES IDENTIFICADOS:
    {", ".join(dados_memorial.get('confrontantes', []))}

    TEXTO COMPLETO DO MEMORIAL (PARA CONTEXTO DE CONVERGÊNCIAS E RUMOS):
    \"\"\"
    {texto_completo_memorial}
    \"\"\"
    
    Gere o documento final formatado prontinho para impressão ou exportação.
    """
    
    try:
        # Chamada oficial utilizando o modelo recomendado (Gemini 2.5 Flash é ideal para texto e rapidez)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.3, # Temperatura baixa para evitar que ele invente dados (alucinação)
            )
        )
        return response.text
    except Exception as e:
        print(f"Erro ao se comunicar com a API do Gemini: {e}")
        return None
