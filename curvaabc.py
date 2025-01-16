import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np

# Configuração para melhorar performance
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# Cache da leitura do arquivo
@st.cache_data
def load_data(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name='CURVA ABC')

# Cache do processamento principal
@st.cache_data
def criar_curva_abc(df):
    """Cria a análise da curva ABC a partir do DataFrame"""
    # Selecionar apenas as linhas 12 a 310
    df = df.iloc[11:310].copy()
    
    # Definir os nomes das colunas
    colunas = ['Unnamed: 0', 'Item', 'SIGLA', 'DESCRIÇÃO', 'UND', 'QTD', 'TOTAL',
               'INCIDÊNCIA DO ITEM (%)', 'INCIDÊNCIA DO ITEM (%) ACUMULADO',
               'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12',
               'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15']
    
    # Ajustar o número de colunas se necessário
    while len(df.columns) < len(colunas):
        df[f'Unnamed: {len(df.columns)}'] = None
    
    df.columns = colunas
    
    # Converter coluna TOTAL para float
    df['TOTAL'] = pd.to_numeric(df['TOTAL'].str.replace('R\$ ', '').str.replace('.', '').str.replace(',', '.'), errors='coerce')
    
    # Remover linhas com total zero ou NaN
    df = df[df['TOTAL'] > 0].reset_index(drop=True)
    
    # Ordenar por valor total em ordem decrescente
    df = df.sort_values('TOTAL', ascending=False)
    
    # Calcular percentuais
    total_geral = df['TOTAL'].sum()
    df['INCIDÊNCIA DO ITEM (%)'] = (df['TOTAL'] / total_geral) * 100
    df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA DO ITEM (%)'].cumsum()
    
    # Classificar em A, B ou C
    conditions = [
        df['INCIDÊNCIA ACUMULADA (%)'] <= 80,
        df['INCIDÊNCIA ACUMULADA (%)'] <= 95
    ]
    choices = ['A', 'B']
    df['CLASSIFICAÇÃO'] = np.select(conditions, choices, default='C')
    
    # Adicionar número do item
    df['NÚMERO DO ITEM'] = range(1, len(df) + 1)
    
    return df

# Cache da criação do gráfico
@st.cache_data
def criar_grafico(df):
    fig, ax1 = plt.subplots(figsize=(12, 6))
    
    # Criar barras para cada item
    bars = ax1.bar(df['NÚMERO DO ITEM'], df['INCIDÊNCIA DO ITEM (%)'], 
                  alpha=0.3, color=df['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'}))
    
    # Adicionar linha de incidência acumulada
    ax2 = ax1.twinx()
    line = ax2.plot(df['NÚMERO DO ITEM'], df['INCIDÊNCIA ACUMULADA (%)'], 
                   color='red', linewidth=2, label='Incidência Acumulada')
    
    # Adicionar linhas de referência
    ax2.axhline(y=80, color='red', linestyle='--', alpha=0.5)
    ax2.axhline(y=95, color='orange', linestyle='--', alpha=0.5)
    
    # Configurar eixos
    ax1.set_xlabel('Número de Itens')
    ax1.set_ylabel('Incidência do Item (%)')
    ax2.set_ylabel('Incidência Acumulada (%)')
    
    # Adicionar legendas
    ax1.legend(bars[:3], ['Classe A', 'Classe B', 'Classe C'])
    ax2.legend(line, ['Incidência Acumulada'])
    
    plt.title('Curva ABC com Distribuição de Itens')
    return fig

def main():
    st.title('Análise Curva ABC')
    
    uploaded_file = st.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Carregar dados com cache
            df = load_data(uploaded_file)
            
            # Processar dados com cache
            df_classificado = criar_curva_abc(df)
            
            # Interface dividida em colunas
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader('Gráfico da Curva ABC')
                fig = criar_grafico(df_classificado)
                st.pyplot(fig)
            
            with col2:
                st.subheader('Resumo da Classificação')
                resumo = pd.DataFrame({
                    'Quantidade de Itens': df_classificado['CLASSIFICAÇÃO'].value_counts(),
                    'Total (R$)': df_classificado.groupby('CLASSIFICAÇÃO')['TOTAL'].sum().round(2)
                }).reset_index()
                resumo.columns = ['Classe', 'Quantidade de Itens', 'Total (R$)']
                st.dataframe(resumo)
            
            # Mostra apenas as primeiras 100 linhas por padrão
            st.subheader('Dados Classificados (Primeiras 100 linhas)')
            st.dataframe(df_classificado.head(100))
            
            # Botão para mostrar dados completos
            if st.button('Mostrar Todos os Dados'):
                st.dataframe(df_classificado)
            
            # Download dos resultados
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_classificado.to_excel(writer, sheet_name='CURVA ABC', index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 Baixar Resultados",
                data=output.getvalue(),
                file_name="curva_abc_classificada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')

if __name__ == '__main__':
    main()
