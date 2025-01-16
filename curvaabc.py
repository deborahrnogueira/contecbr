import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

def carregar_dados(uploaded_file):
    """Carrega e prepara os dados da planilha"""
    try:
        # Verifica a extensão do arquivo
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'csv':
            df = pd.read_csv(uploaded_file)
        elif file_extension in ['xlsx', 'xlsm']:
            # Ler especificamente as colunas B até I, começando da linha 11
            df = pd.read_excel(
                uploaded_file,
                sheet_name='CURVA ABC',
                usecols='B:I',  # Específica as colunas B até I
                skiprows=10,    # Pula as primeiras 10 linhas
            )
        else:
            raise Exception("Formato de arquivo não suportado. Use .xlsx, .xlsm ou .csv")
        
        # Definir os nomes das colunas corretamente
        df.columns = [
            'ITEM', 'SIGLA', 'DESCRIÇÃO', 'UND', 'QTD', 'TOTAL',
            'INCIDÊNCIA DO ITEM (%)', 'INCIDÊNCIA DO ITEM (%) ACUMULADO'
        ]
        
        return df
    except Exception as e:
        raise Exception(f"Erro ao carregar arquivo: {str(e)}")

def limpar_valor_monetario(valor):
    """Limpa e converte valores monetários para float"""
    if pd.isna(valor):
        return 0.0
    
    if isinstance(valor, (int, float)):
        return float(valor)
    
    try:
        # Remove caracteres não numéricos exceto . e ,
        valor_str = str(valor)
        if 'R$' in valor_str:
            valor_str = valor_str.replace('R$', '').strip()
        if '.' in valor_str and ',' in valor_str:
            # Caso brasileiro: 1.234,56
            valor_str = valor_str.replace('.', '').replace(',', '.')
        elif ',' in valor_str:
            # Caso com vírgula decimal: 1234,56
            valor_str = valor_str.replace(',', '.')
        valor_limpo = ''.join(c for c in valor_str if c.isdigit() or c == '.')
        return float(valor_limpo)
    except:
        return 0.0

def formatar_moeda_real(valor):
    """Formata valor para R$ com separadores de milhares"""
    try:
        # Converte para float caso não seja
        valor = float(valor)
        # Formata com R$, separador de milhares e 2 casas decimais
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

@st.cache_data
def processar_dados(df):
    """Processa os dados para análise ABC"""
    try:
        # Criar cópia do DataFrame
        df_processado = df.copy()
        
        # Limpar e converter a coluna TOTAL
        df_processado['TOTAL'] = df_processado['TOTAL'].apply(limpar_valor_monetario)
        
        # Remover linhas com total zero ou nulo
        df_processado = df_processado[df_processado['TOTAL'] > 0].reset_index(drop=True)
        
        # Ordenar por valor total
        df_processado = df_processado.sort_values('TOTAL', ascending=False).reset_index(drop=True)
        
        # Calcular percentuais
        total_geral = df_processado['TOTAL'].sum()
        df_processado['INCIDÊNCIA DO ITEM (%)'] = (df_processado['TOTAL'] / total_geral * 100).round(2)
        df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'] = df_processado['INCIDÊNCIA DO ITEM (%)'].cumsum().round(2)
        
        # Classificar itens
        df_processado['CLASSIFICAÇÃO'] = df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'].apply(
            lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
        
        # Adicionar número do item e formatar total
        df_processado['NÚMERO DO ITEM'] = range(1, len(df_processado) + 1)
        df_processado['TOTAL_FORMATADO'] = df_processado['TOTAL'].apply(formatar_moeda_real)
        
        return df_processado
        
    except Exception as e:
        raise Exception(f"Erro ao processar dados: {str(e)}")

# [Resto das funções permanece igual...]

def main():
    st.title('Análise Curva ABC')
    
    # Sidebar para upload com mensagem sobre formatos aceitos
    st.sidebar.title('Configurações')
    st.sidebar.info('Formatos aceitos: .xlsx, .xlsm e .csv')
    uploaded_file = st.sidebar.file_uploader("Carregue a planilha", type=['xlsx', 'xlsm', 'csv'])
    
    if uploaded_file is not None:
        try:
            # Carregar e processar dados
            df = carregar_dados(uploaded_file)
            df_processado = processar_dados(df)
            
            # Interface com abas
            tab1, tab2, tab3 = st.tabs(['Visão Geral', 'Análise Detalhada', 'Filtros'])
            
            # [... código das abas permanece igual ...]
            
            # Corrigindo o botão de download
            try:
                buffer = io.BytesIO()
                
                try:
                    # Tenta usar xlsxwriter primeiro
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_processado.to_excel(writer, index=False, sheet_name='Dados Processados')
                except ImportError:
                    # Se xlsxwriter não estiver disponível, usa openpyxl
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_processado.to_excel(writer, index=False, sheet_name='Dados Processados')
                
                st.sidebar.download_button(
                    label="📥 Baixar Dados Processados (Excel)",
                    data=buffer.getvalue(),
                    file_name="curva_abc_processada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.sidebar.error(f"Erro ao gerar arquivo para download: {str(e)}")
                # Oferece alternativa em CSV
                csv = df_processado.to_csv(index=False).encode('utf-8')
                st.sidebar.download_button(
                    label="📥 Baixar Dados Processados (CSV)",
                    data=csv,
                    file_name="curva_abc_processada.csv",
                    mime="text/csv"
                )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.info('Verifique se o arquivo está no formato correto e contém todos os dados necessários.')

if __name__ == '__main__':
    main()
