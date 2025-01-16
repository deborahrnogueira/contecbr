import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from datetime import datetime
from fpdf import FPDF

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

# [Mantido o resto das funções auxiliares como está...]

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
            
            # [Mantido o código das abas como está...]
            
            # Corrigindo o botão de download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_processado.to_excel(writer, index=False, sheet_name='Dados Processados')
                writer.close()
            
            st.sidebar.download_button(
                label="📥 Baixar Dados Processados (Excel)",
                data=buffer.getvalue(),
                file_name="curva_abc_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.info('Verifique se o arquivo está no formato correto e contém todos os dados necessários.')

if __name__ == '__main__':
    main()
