# Guia Rápido - GERADOR DE MEMORIAL DESCRITIVO v6.3

## 🎯 Visão Geral

O aplicativo possui 4 abas principais, cada uma com uma função específica:

```
┌─────────────────────────────────────────────────────────────┐
│  📝 Memorial  │  🤝 Anuências  │  🌾 INCRA  │  🏛️ Cartório  │
└─────────────────────────────────────────────────────────────┘
```

---

## 📝 Aba 1: Memorial Descritivo

**Objetivo**: Processar planta e roteiro para gerar o memorial descritivo.

### Fluxo:
1. Acesse a aba **"📝 Memorial Descritivo"**
2. Escolha entre:
   - **📁 Processamento de Arquivos**: Carregue PDF ou imagens da planta e roteiro
   - **📝 Colagem de Texto Manual**: Cole o texto diretamente
3. Clique em **"🔍 Analisar Documentos"**
4. Baixe o memorial em DOCX

### Saída:
- ✅ Memorial Descritivo (.docx)
- ✅ Tabela com segmentos processados

---

## 🤝 Aba 2: Anuências Co-proprietários

**Objetivo**: Gerar anuências convencionais para cada confrontante.

### Fluxo:
1. Processe um memorial na aba anterior
2. Acesse a aba **"🤝 Anuências Co-proprietários"**
3. Preencha os dados de cada confrontante:
   - Nome completo
   - CPF
   - Endereço
   - TRT
4. Clique em **"📄 Gerar Anuência"** para cada confrontante
5. Baixe as anuências

### Saída:
- ✅ Anuências Convencionais (.docx) - uma por confrontante

---

## 🌾 Aba 3: Anuências INCRA

**Objetivo**: Gerar anuências INCRA com layout específico (paisagem).

### Fluxo:
1. Processe um memorial na aba "Memorial Descritivo"
2. Acesse a aba **"🌾 Anuências INCRA"**
3. Clique em **"🌾 Gerar Anuências INCRA"**
4. O sistema irá:
   - Extrair dados do memorial
   - Criar uma anuência por confrontante
   - Incluir tabela de vértices
   - Adicionar assinaturas
5. Baixe o arquivo ZIP com todas as anuências

### Características:
- 📐 Layout em paisagem (A4)
- 📊 Tabela com 8 colunas de vértices
- 🖊️ Assinaturas de proprietário, confrontante e técnico
- 📅 Data automática em português

### Saída:
- ✅ Anuências INCRA (.zip) - múltiplos arquivos

---

## 🏛️ Aba 4: Requerimento de Cartório

**Objetivo**: Gerar requerimento de cartório para registro imobiliário.

### Fluxo A: Upload de Documentos (Automático)
1. Acesse a aba **"🏛️ Requerimento de Cartório"**
2. Clique na aba **"📁 Upload de Documentos"**
3. Carregue imagens dos documentos:
   - RG do requerente
   - CPF do requerente
   - Planta INCRA
   - Outros documentos relevantes
4. Clique em **"🔍 Extrair Dados dos Documentos"**
5. O sistema extrairá os dados automaticamente

### Fluxo B: Preenchimento Manual
1. Acesse a aba **"🏛️ Requerimento de Cartório"**
2. Clique na aba **"📝 Preenchimento Manual"**
3. Preencha os dados:
   - **Requerente 1** (Proprietário)
   - **Requerente 2** (Cônjuge - opcional)
   - **Dados do Imóvel**
4. Clique em **"💾 Salvar Dados Manualmente"**

### Geração do Requerimento
1. Após extrair/preencher dados, clique em **"🏛️ Gerar Requerimento de Cartório"**
2. O sistema irá:
   - Validar os dados
   - Ajustar profissões conforme gênero
   - Remover duplicações
   - Preencher o template
3. Baixe o requerimento em DOCX

### Saída:
- ✅ Requerimento de Cartório (.docx)

---

## ⚙️ Painel de Controle (Sidebar)

O painel lateral permite configurar dados institucionais:

### 🏢 Empresa / Técnico
- Nome da empresa
- Endereço
- Telefone
- Email
- Nome do técnico
- CFTA do técnico
- CPF do técnico
- TRT

### 👥 Cliente & Imóvel
- Identificação do imóvel
- Proprietário
- Município/Localidade
- Área (hectares)
- Perímetro (metros)
- Comarca
- Matrícula

---

## 🔧 Configurações Adicionais

### Modelo Gemini
Escolha entre diferentes modelos de IA:
- Gemini 3.5 Flash (Fronteira/Padrão)
- Gemini 3.1 Pro (Raciocínio Avançado)
- Gemini 3.1 Flash-Lite (Alta Velocidade)
- Gemini 2.5 Pro (Estável e Preciso)
- Gemini 2.5 Flash (Trabalho Diário)

### Resolução (DPI)
- Padrão: 150 DPI
- Intervalo: 100-400 DPI
- Maior DPI = melhor qualidade, mas mais lento

### Upload Máximo
- Padrão: 50 MB
- Intervalo: 10-100 MB

---

## 📋 Checklist de Uso

- [ ] Configurar dados da empresa no painel lateral
- [ ] Configurar dados do técnico no painel lateral
- [ ] Carregar ou colar planta e roteiro
- [ ] Gerar memorial descritivo
- [ ] Revisar tabela de segmentos
- [ ] Gerar anuências co-proprietários
- [ ] Gerar anuências INCRA
- [ ] Preparar documentos para cartório
- [ ] Gerar requerimento de cartório
- [ ] Baixar todos os documentos

---

## ⚠️ Dicas Importantes

1. **Sempre comece pela aba "Memorial Descritivo"** - as outras abas dependem dos dados processados aqui
2. **Configure o painel lateral antes de começar** - garante que todos os documentos tenham os dados corretos
3. **Use o modelo Gemini 2.5 Flash para melhor equilíbrio** entre velocidade e qualidade
4. **Revise os dados extraídos** antes de gerar os documentos finais
5. **Mantenha os arquivos PDF/imagens em boa qualidade** para melhor extração de dados
6. **Limpe a sessão após terminar** para liberar memória

---

## 🆘 Resolução de Problemas

### "Erro ao processar memorial"
- Verifique se o PDF/imagem está legível
- Tente aumentar a resolução (DPI)
- Verifique se o texto está em português

### "Dados não foram extraídos"
- Verifique a conexão com a internet
- Tente novamente com uma imagem mais clara
- Verifique se a chave de API do Gemini está configurada

### "Template não encontrado"
- Verifique se o arquivo `template_requerimento.docx` existe no diretório
- Copie o arquivo para o diretório correto

### "Erro ao gerar anuência INCRA"
- Verifique se processou um memorial primeiro
- Tente novamente com dados mais completos no painel lateral

---

## 📞 Suporte

Para dúvidas ou problemas:
- Verifique o arquivo `CORRECOES_REALIZADAS.md` para histórico de correções
- Consulte o `README.md` para informações técnicas
- Revise os logs do aplicativo para mensagens de erro detalhadas

---

**Versão**: 6.3
**Data**: 22 de Julho de 2026
**Status**: ✅ Totalmente Funcional
