# 📤 Gerador de Coleção Postman - Fornecedores NIBO

Ferramenta para converter planilhas Excel de agendamentos de fornecedores em coleções Postman prontas para envio à API NIBO.

## 🎯 Objetivo

Simplificar o processo de envio de múltiplos agendamentos de fornecedores para a API NIBO, gerando automaticamente uma coleção Postman com todas as requisições HTTP necessárias.

## 📋 Requisitos

- Python 3.8+
- Streamlit
- pandas
- openpyxl

## 🚀 Como Usar

### 1. Instalação

```bash
cd C:\Users\bened\OneDrive\Documentos\Gerador-Fornecedor-NIBO
pip install streamlit pandas openpyxl
```

### 2. Executar

```bash
streamlit run main_fornecedor.py
```

### 3. Preparar Planilha

Sua planilha Excel deve conter as seguintes colunas:

#### Obrigatórias:
- `stakeholderId` - ID do fornecedor no NIBO
- `categoryId` - ID da categoria no NIBO  
- `value` - Valor do agendamento (número)
- `costCenterId` - ID do centro de custo
- `date` - Data do agendamento (formato: YYYY-MM-DD)
- `Vencimento` - Data de vencimento (formato: YYYY-MM-DD)
- `description` - Descrição do agendamento

#### Opcionais:
- `accountId` - ID da conta bancária
- `reference` - Referência/Número do documento
- `Data de competência` - Data de competência (formato: YYYY-MM-DD)

### 4. Exemplo de Linha

```
ID: 570968
stakeholderId: e00a5c53-3f79-4e37-8808-d9c8261daf7f
categoryId: dc99f0b0-3696-489b-bea1-fa72f24dbe28
value: 82.04
costCenterId: bba9250e-09c5-486b-b999-620dc6e79545
date: 2025-10-10
Vencimento: 2025-10-10
Data de competência: 2025-10-10
description: Material BALMAIN SHOP CIDADE JARDIM NF: 3126473
accountId: e876abc3-0bac-4a31-b966-d453d814723d
reference: ITAU SALARIO
```

### 5. Processo no Sistema

1. **Configure API** (menu lateral):
   - URL da API NIBO
   - Token de autenticação
   - Salve o perfil

2. **Upload da planilha**:
   - Arraste ou selecione o arquivo Excel
   - Sistema valida automaticamente

3. **Gere a coleção**:
   - Clique em "Gerar Coleção Postman"
   - Aguarde o processamento

4. **Baixe o JSON**:
   - Download do arquivo de coleção
   - Arquivo pronto para Postman

### 6. Importar no Postman

1. Abra o Postman
2. Clique em "Import"
3. Selecione o arquivo JSON baixado
4. A coleção aparecerá com todas as requisições

### 7. Executar Requisições

**Opção 1 - Individual:**
- Clique em cada requisição
- Clique em "Send"

**Opção 2 - Em Lote:**
- Clique em "..." na coleção
- Selecione "Run Collection"
- Configure delay entre requisições (recomendado: 100-500ms)
- Clique em "Run"

## ✨ Funcionalidades

- ✅ Upload de planilhas Excel
- ✅ Validação automática de colunas
- ✅ Preview dos dados
- ✅ Estatísticas (total de registros, valor total, fornecedores únicos)
- ✅ Geração automática de requisições HTTP
- ✅ Configuração centralizada de API
- ✅ Perfis salvos de configuração
- ✅ Download da coleção Postman (JSON)
- ✅ Download dos dados processados (Excel)
- ✅ Instruções integradas

## 🔧 Configuração da API

As configurações são salvas de forma segura usando o módulo `config_segura.py`.

**Campos necessários:**
- **URL da API**: `https://api.nibo.com.br/empresas/v1/`
- **Token**: Seu token de autenticação NIBO
- **Nome do Perfil**: Para salvar múltiplas configurações

## 📊 Estrutura da Coleção Gerada

```json
{
  "info": {
    "name": "Agendamentos Fornecedores - DD/MM/YYYY",
    "schema": "https://schema.getpostman.com/json/collection/v2.1.0/collection.json"
  },
  "item": [
    {
      "name": "1 - Descrição do Agendamento",
      "request": {
        "method": "POST",
        "header": [...],
        "body": {...},
        "url": "{{base_url}}/financial-schedules"
      }
    }
  ],
  "variable": [
    {"key": "base_url", "value": "..."},
    {"key": "token", "value": "..."}
  ]
}
```

## ⚠️ Avisos Importantes

1. **Validação**: Sempre valide os dados na planilha antes de gerar
2. **API Rate Limit**: Configure delay entre requisições no Postman
3. **Backup**: Mantenha backup da planilha original
4. **Token**: Nunca compartilhe seu token de API
5. **Testes**: Teste com poucos registros primeiro

## 🐛 Solução de Problemas

### Erro "Colunas não encontradas"
- Verifique se os nomes das colunas estão exatamente como especificado
- Colunas são case-sensitive

### Erro ao gerar coleção
- Verifique se URL e Token estão configurados
- Confira se os IDs (stakeholderId, categoryId, etc) são válidos

### Valores não aparecem corretamente
- Verifique formato de datas (YYYY-MM-DD)
- Verifique formato de valores (use ponto como decimal)

## 📞 Suporte

Para dúvidas ou problemas, verifique:
1. Os logs de erro no próprio sistema
2. A validação das colunas
3. Os dados de exemplo fornecidos

## 📝 Licença

Este projeto é de uso interno.

---

**Desenvolvido para facilitar o envio de agendamentos de fornecedores à API NIBO** 🚀
