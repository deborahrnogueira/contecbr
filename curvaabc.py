import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np

def validar_dados(df):
    """Valida os dados de entrada"""
    if df.empty:
        raise ValueError("A planilha está vazia")
    
    if df.shape[1] < 7:  # Verifica se tem pelo menos 7 colunas (até a coluna G)
        raise ValueError("A planilha não tem o número mínimo de colunas necessário")
    
    # Verifica se há dados suficientes após o slice
    if len(df.iloc[11:310]) == 0:
        raise ValueError("Não há dados suficientes nas linhas especificadas (12 a 310)")
    
    return True

def limpar_valor_monetario(valor):
    """Limpa e converte valores monetários para float"""
    try:
        if isinstance(valor, (int, float)):
            return float(valor)
        return float(str(valor).replace('R$ ', '').replace('.', '').replace(',', '.'))
    except:
        return 0.0

def criar_curva_abc(df):
    """
    Cria a análise da curva ABC a partir do DataFrame
    """
    try:
        # Selecionar apenas as linhas 12 a 310
        df = df.iloc[11:310].copy()
        
        # Converter coluna G (TOTAL) para float, tratando possíveis erros
        df['TOTAL'] = df.iloc[:, 6].apply(limpar_valor_monetario)
        
        # Remover linhas com total zero
        df = df[df['TOTAL'] > 0]
        
        # Ordenar por valor total em ordem decrescente
        df = df.sort_values('TOTAL', ascending=False)
        
        # Calcular percentuais
        total_geral = df['TOTAL'].sum()
        df['INCIDÊNCIA DO ITEM (%)'] = (df['TOTAL'] / total_geral) * 100
        df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA DO ITEM (%)'].cumsum()
        
        # Classificar em A, B ou C
        def get_classe(acumulado):
            if acumulado <= 80:
                return 'A'
            elif acumulado <= 95:
                return 'B'
            else:
                return 'C'
        
        df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(get_classe)
        
        # Adicionar contagem por classificação
        df['NÚMERO DO ITEM'] = range(1, len(df) + 1)
        
        return df
    except Exception as e:
        raise Exception(f"Erro ao processar a curva ABC: {str(e)}")

def criar_grafico(df):
    """
    Cria o gráfico da curva ABC
    """
    plt.clf()  # Limpa a figura atual
    fig, ax1 = plt.subplots(figsize=(12, 7))
    
    # Plotar a curva ABC
    ax1.plot(df['NÚMERO DO ITEM'], df['INCIDÊNCIA ACUMULADA (%)'], 
            marker='o', color='blue', linewidth=2, markersize=4)
    
    # Adicionar linha de referência em 80%
    ax1.axhline(y=80, color='r', linestyle='--', alpha=0.5)
    ax1.axhline(y=95, color='orange', linestyle='--', alpha=0.5)
    
    # Configurar os eixos
    ax1.set_xlabel('Número de Itens', fontsize=10)
    ax1.set_ylabel('Incidência Acumulada (%)', fontsize=10)
    ax1.grid(True, alpha=0.3)
    
    # Adicionar título e legendas
    plt.title('Curva ABC', fontsize=12, pad=20)
    
    # Adicionar descrição das classes
    plt.text(0, 82, 'Classe A (0-80%)', color='r', fontsize=8)
    plt.text(0, 97, 'Classe B (80-95%)', color='orange', fontsize=8)
    plt.text(0, 99, 'Classe C (95-100%)', color='g', fontsize=8)
    
    return fig

def main():
    st.set_page_config(page_title="Análise Curva ABC", layout="wide")
    
    st.title('Análise Curva ABC')
    st.markdown("""
    Esta aplicação realiza a análise da Curva ABC a partir de uma planilha Excel.
    Por favor, carregue um arquivo Excel contendo os dados na aba 'CURVA ABC'.
    """)
    
    uploaded_file = st.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            
            # Validar os dados
            validar_dados(df)
            
            # Processar os dados e criar a curva ABC
            df_classificado = criar_curva_abc(df)
            
            # Criar duas colunas para organizar o layout
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader('Gráfico da Curva ABC')
                fig = criar_grafico(df_classificado)
                st.pyplot(fig)
            
            with col2:
                st.subheader('Resumo da Classificação')
                resumo = df_classificado['CLASSIFICAÇÃO'].value_counts().reset_index()
                resumo.columns = ['Classe', 'Quantidade de Itens']
                resumo['Percentual de Itens'] = (resumo['Quantidade de Itens'] / len(df_classificado) * 100).round(2)
                st.dataframe(resumo)
            
            st.subheader('Dados Classificados')
            st.dataframe(df_classificado)
            
            # Exportar os resultados para Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_classificado.to_excel(writer, sheet_name='CURVA ABC', index=False)
            
            output.seek(0)  # Reset pointer
            
            st.download_button(
                label="📥 Baixar Resultados em Excel",
                data=output.getvalue(),
                file_name="curva_abc_classificada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.info('Por favor, verifique se o arquivo está no formato correto e contém todos os dados necessários.')

if __name__ == '__main__':
    main()
