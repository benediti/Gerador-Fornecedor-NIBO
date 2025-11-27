"""
Gerador de Coleção Postman para Agendamentos de Fornecedores - NIBO
Versão simplificada focada em gerar requisições HTTP para a API NIBO
"""

import streamlit as st
import pandas as pd
import json
import uuid
from datetime import datetime
from io import BytesIO
import sys
import os

# Adicionar diretório raiz ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_segura import ConfigSegura

# ========================================
# CONFIGURAÇÃO DA PÁGINA
# ========================================
st.set_page_config(
    page_title="Gerador Postman - Fornecedores NIBO",
    page_icon="📤",
    layout="wide"
)

st.markdown("""
# 📤 Gerador de Coleção Postman - Fornecedores NIBO
### Converta sua planilha de agendamentos em requisições HTTP prontas para envio
---
""")

# ========================================
# SIDEBAR - CONFIGURAÇÕES
# ========================================
config_segura = ConfigSegura()

with st.sidebar:
    st.header("⚙️ Configuração da API")
    
    perfis_salvos = config_segura.listar_perfis()
    
    perfil_selecionado = ""
    if perfis_salvos:
        perfil_selecionado = st.selectbox("Perfil salvo:", [""] + perfis_salvos)
    
    if perfil_selecionado:
        api_url_salva, api_token_salvo = config_segura.carregar_config(perfil_selecionado)
        api_url = api_url_salva or ""
        api_token = api_token_salvo or ""
        st.success(f"✅ Perfil '{perfil_selecionado}' carregado")
    else:
        api_url, api_token = config_segura.carregar_config("default")
        api_url = api_url or ""
        api_token = api_token or ""
    
    api_url_input = st.text_input(
        "URL da API:", 
        value=api_url, 
        placeholder="https://api.nibo.com.br/empresas/v1/"
    )
    api_token_input = st.text_input("Token:", value=api_token, type="password")
    
    col1, col2 = st.columns(2)
    with col1:
        nome_perfil = st.text_input("Nome perfil:", value="default")
    with col2:
        if st.button("💾 Salvar"):
            if api_url_input and api_token_input and nome_perfil:
                if config_segura.salvar_config(api_url_input, api_token_input, nome_perfil):
                    st.success("✅ Salvo!")
                    st.rerun()

# ========================================
# ÁREA PRINCIPAL
# ========================================

# Upload de arquivo
st.header("📤 Upload da Planilha de Fornecedores")

st.info("""
💡 **Formato esperado da planilha:**
- **Colunas obrigatórias:** ID, stakeholderId, categoryId, value, costCenterId, date, Vencimento, Data de competência, description, accountId, reference
- Os dados já devem estar processados e prontos para envio
""")

uploaded_file = st.file_uploader(
    "Selecione a planilha Excel com os agendamentos",
    type=['xlsx', 'xls'],
    help="Arquivo deve conter todas as colunas necessárias"
)

if uploaded_file is not None:
    try:
        # Ler arquivo
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Registros encontrados: {len(df)}")
        
        # Verificar colunas necessárias
        colunas_necessarias = [
            'stakeholderId', 'categoryId', 'value', 'costCenterId', 
            'date', 'Vencimento', 'description'
        ]
        
        colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]
        
        if colunas_faltando:
            st.error(f"❌ Colunas obrigatórias não encontradas: {colunas_faltando}")
            st.info(f"Colunas disponíveis: {', '.join(df.columns.tolist())}")
        else:
            # Preview dos dados
            with st.expander("👁️ Visualizar dados", expanded=False):
                st.dataframe(df.head(20), use_container_width=True)
                if len(df) > 20:
                    st.info(f"Mostrando as primeiras 20 linhas de {len(df)} registros")
            
            # Estatísticas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Total de Registros", f"{len(df):,}")
            with col2:
                if 'value' in df.columns:
                    st.metric("💰 Valor Total", f"R$ {df['value'].sum():,.2f}")
            with col3:
                if 'stakeholderId' in df.columns:
                    st.metric("👥 Fornecedores Únicos", f"{df['stakeholderId'].nunique():,}")
            
            # Configurações para geração
            st.markdown("---")
            st.header("⚙️ Configurações da Coleção")
            
            nome_colecao = st.text_input(
                "Nome da Coleção:",
                value=f"Agendamentos Fornecedores - {datetime.now().strftime('%d/%m/%Y')}"
            )
            
            st.info("📍 **Endpoint:** `https://api.nibo.com.br/empresas/v1/schedules/debit`")
            
            # Botão para gerar coleção
            if st.button("🚀 Gerar Coleção Postman", type="primary", use_container_width=True):
                if not api_url_input or not api_token_input:
                    st.error("❌ Configure a URL e Token da API antes de gerar a coleção!")
                else:
                    with st.spinner("🔄 Gerando coleção Postman..."):
                        # Criar estrutura da coleção Postman
                        collection = {
                            "info": {
                                "_postman_id": str(uuid.uuid4()),
                                "name": nome_colecao,
                                "schema": "https://schema.getpostman.com/json/collection/v2.1.0/collection.json",
                                "description": f"Coleção gerada automaticamente em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                            },
                            "item": [],
                            "variable": [
                                {
                                    "key": "base_url",
                                    "value": api_url_input.rstrip('/'),
                                    "type": "string"
                                },
                                {
                                    "key": "token",
                                    "value": api_token_input,
                                    "type": "string"
                                }
                            ]
                        }
                        
                        # Gerar requisições
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        for idx, row in df.iterrows():
                            progress = (idx + 1) / len(df)
                            progress_bar.progress(progress)
                            status_text.text(f"Gerando requisição {idx + 1}/{len(df)}...")
                            
                            # Preparar body da requisição no formato NIBO
                            body = {
                                "stakeholderId": str(row['stakeholderId']) if pd.notna(row['stakeholderId']) else "",
                                "description": str(row.get('description', '')) if pd.notna(row.get('description')) else "",
                                "reference": str(row.get('reference', '')) if pd.notna(row.get('reference')) else "",
                                "scheduleDate": str(row['date'])[:10] if pd.notna(row['date']) else "",
                                "dueDate": str(row['Vencimento'])[:10] if pd.notna(row['Vencimento']) else "",
                                "accrualDate": str(row.get('Data de competência', row['date']))[:10] if pd.notna(row.get('Data de competência', row['date'])) else "",
                                "categories": [
                                    {
                                        "categoryId": str(row['categoryId']) if pd.notna(row['categoryId']) else "",
                                        "value": float(row['value']) if pd.notna(row['value']) else 0.0
                                    }
                                ],
                                "costCenterValueType": 0,
                                "costCenters": [
                                    {
                                        "costCenterId": str(row['costCenterId']) if pd.notna(row['costCenterId']) else "",
                                        "value": float(row['value']) if pd.notna(row['value']) else 0.0
                                    }
                                ]
                            }
                            
                            # Criar requisição
                            request_name = f"Agendamento {idx + 1}"
                            if 'description' in row and pd.notna(row['description']):
                                request_name = f"{idx + 1} - {str(row['description'])[:50]}"
                            
                            request_item = {
                                "name": request_name,
                                "request": {
                                    "method": "POST",
                                    "header": [
                                        {
                                            "key": "Content-Type",
                                            "value": "application/json"
                                        },
                                        {
                                            "key": "ApiToken",
                                            "value": "{{token}}"
                                        }
                                    ],
                                    "body": {
                                        "mode": "raw",
                                        "raw": json.dumps(body, indent=2, ensure_ascii=False)
                                    },
                                    "url": {
                                        "raw": "https://api.nibo.com.br/empresas/v1/schedules/debit",
                                        "protocol": "https",
                                        "host": ["api", "nibo", "com", "br"],
                                        "path": ["empresas", "v1", "schedules", "debit"]
                                    }
                                },
                                "response": []
                            }
                            
                            collection["item"].append(request_item)
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        st.success(f"✅ Coleção gerada com {len(df)} requisições!")
                        
                        # Preparar download
                        collection_json = json.dumps(collection, indent=2, ensure_ascii=False)
                        
                        st.markdown("---")
                        st.header("📥 Download da Coleção")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.download_button(
                                label="📦 Download Coleção Postman (JSON)",
                                data=collection_json,
                                file_name=f"colecao_fornecedores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                                mime="application/json",
                                type="primary",
                                use_container_width=True
                            )
                        
                        with col2:
                            # Excel com resumo
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, sheet_name='Dados Processados')
                            
                            st.download_button(
                                label="📊 Download Dados (Excel)",
                                data=output.getvalue(),
                                file_name=f"dados_fornecedores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        # Instruções
                        st.markdown("---")
                        st.markdown("### 🎯 Próximos Passos")
                        st.info("""
                        1. ✅ **Baixe** o arquivo JSON da coleção
                        2. 📂 **Abra** o Postman
                        3. ➕ **Import** → Selecione o arquivo JSON baixado
                        4. ✏️ **Configure** as variáveis de ambiente se necessário
                        5. ▶️ **Execute** as requisições (pode usar "Run Collection" para executar todas de uma vez)
                        6. ✅ **Verifique** os resultados na API NIBO
                        """)
                        
                        st.success("🎉 Coleção pronta para uso no Postman!")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivo: {str(e)}")
        import traceback
        with st.expander("🔍 Detalhes do erro"):
            st.code(traceback.format_exc())

else:
    # Instruções quando não há arquivo
    st.info("📁 Faça upload da planilha para começar")
    
    with st.expander("📋 Instruções de Uso", expanded=True):
        st.markdown("""
        ### 🎯 Objetivo:
        Gerar uma coleção Postman com todas as requisições HTTP necessárias para enviar agendamentos de fornecedores à API NIBO.
        
        ### 📊 Formato da Planilha:
        A planilha Excel deve conter as seguintes colunas:
        
        **Obrigatórias:**
        - `stakeholderId` - ID do fornecedor no NIBO
        - `categoryId` - ID da categoria no NIBO
        - `value` - Valor do agendamento
        - `costCenterId` - ID do centro de custo
        - `date` - Data do agendamento
        - `Vencimento` - Data de vencimento
        - `description` - Descrição do agendamento
        
        **Opcionais:**
        - `accountId` - ID da conta bancária
        - `reference` - Referência/Número do documento
        - `Data de competência` - Data de competência
        
        ### � Passo a Passo:
        1. **Prepare** sua planilha Excel com os dados dos agendamentos
        2. **Configure** a API NIBO (URL e Token) no menu lateral
        3. **Faça upload** da planilha
        4. **Clique** em "Gerar Coleção Postman"
        5. **Baixe** o arquivo JSON gerado
        6. **Importe** no Postman
        7. **Execute** as requisições
        
        ### ✅ Vantagens:
        - ✓ Geração automática de todas as requisições
        - ✓ Configuração centralizada de URL e Token
        - ✓ Fácil importação no Postman
        - ✓ Execução em lote
        - ✓ Rastreabilidade de cada agendamento
        """)

# Footer
st.markdown("---")
st.markdown("*📤 Gerador de Coleção Postman - Fornecedores NIBO v1.0*")
