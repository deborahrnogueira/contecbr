import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

def criar_curva_abc(df):
    # Usar a coluna com espaços " TOTAL "
    df = df.sort_values(' TOTAL ', ascending=False)
    
    # Converter valor monetário para float
    df[' TOTAL '] = df[' TOTAL '].str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float)
    
    # Calcular percentuais
    total_geral = df[' TOTAL '].sum()
    df['INCIDÊNCIA DO ITEM (%)'] = (df[' TOTAL '] / total_geral) * 100
    df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA DO ITEM (%)'].cumsum()
    
    # Classificar
    def get_classe(acumulado):
        if acumulado <= 80:
            return 'A'
        elif acumulado <= 95:
            return 'B'
        else:
            return 'C'
    
    df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(get_classe)
    return df

def main():
    st.title('Análise Curva ABC')
    
    uploaded_file = st.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Ler arquivo da aba CURVA ABC
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            
            # Verificar colunas disponíveis
            st.write("Colunas disponíveis:", df.columns.tolist())
            
            # Processar dados
            df_classificado = criar_curva_abc(df)
            
            st.subheader('Dados Classificados')
            st.dataframe(df_classificado)
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')

if __name__ == '__main__':
    main()
