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
                sheet_name='ANALISE',
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
            
            # Aba 1: Visão Geral
            with tab1:
                # Criar gráficos
                fig_pareto = go.Figure()
                
                # Barras de incidência
                fig_pareto.add_trace(go.Bar(
                    x=df_processado['NÚMERO DO ITEM'],
                    y=df_processado['INCIDÊNCIA DO ITEM (%)'],
                    name='Incidência',
                    marker_color=df_processado['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'}),
                    text=df_processado['ITEM'],
                    customdata=np.column_stack((df_processado['TOTAL_FORMATADO'], df_processado['CLASSIFICAÇÃO']))
                ))
                
                # Linha de acumulado
                fig_pareto.add_trace(go.Scatter(
                    x=df_processado['NÚMERO DO ITEM'],
                    y=df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'],
                    name='Acumulado',
                    line=dict(color='red', width=2),
                    yaxis='y2'
                ))
                
                # Configuração do layout
                fig_pareto.update_layout(
                    title='Curva ABC (Pareto)',
                    xaxis_title='Número do Item',
                    yaxis_title='Incidência Individual (%)',
                    yaxis2=dict(
                        title='Incidência Acumulada (%)',
                        overlaying='y',
                        side='right'
                    ),
                    showlegend=True,
                    height=600
                )
                
                # Gráfico de Pizza
                df_pizza = df_processado.groupby('CLASSIFICAÇÃO').agg({
                    'TOTAL': 'sum',
                    'ITEM': 'count'
                }).reset_index()
                
                df_pizza['Percentual'] = (df_pizza['TOTAL'] / df_pizza['TOTAL'].sum() * 100).round(2)
                
                fig_pizza = px.pie(
                    df_pizza,
                    values='TOTAL',
                    names='CLASSIFICAÇÃO',
                    title='Distribuição por Classe'
                )
                
                # Exibir gráficos em colunas
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.plotly_chart(fig_pareto, use_container_width=True)
                with col2:
                    st.plotly_chart(fig_pizza, use_container_width=True)
                
                # Resumo
                st.subheader('Resumo por Classe')
                resumo = df_processado.groupby('CLASSIFICAÇÃO').agg({
                    'ITEM': 'count',
                    'TOTAL': ['sum', 'mean']
                }).round(2)
                resumo.columns = ['Quantidade de Itens', 'Valor Total', 'Valor Médio']
                st.dataframe(resumo)
            
            # Aba 2: Análise Detalhada
            with tab2:
                st.dataframe(df_processado)
            
            # Aba 3: Filtros
            with tab3:
                col1, col2 = st.columns(2)
                with col1:
                    classes_selecionadas = st.multiselect(
                        'Filtrar por classe',
                        ['A', 'B', 'C'],
                        default=['A', 'B', 'C']
                    )
                
                with col2:
                    valor_min = float(df_processado['TOTAL'].min())
                    valor_max = float(df_processado['TOTAL'].max())
                    valores = st.slider(
                        'Faixa de valores',
                        min_value=valor_min,
                        max_value=valor_max,
                        value=(valor_min, valor_max)
                    )
                
                busca = st.text_input('Buscar por descrição')
                
                # Aplicar filtros
                mask = (
                    (df_processado['CLASSIFICAÇÃO'].isin(classes_selecionadas)) &
                    (df_processado['TOTAL'] >= valores[0]) &
                    (df_processado['TOTAL'] <= valores[1])
                )
                
                if busca:
                    mask &= df_processado['DESCRIÇÃO'].str.contains(busca, case=False, na=False)
                
                df_filtrado = df_processado[mask]
                st.dataframe(df_filtrado)
            
            # Botão de download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_processado.to_excel(writer, index=False, sheet_name='Dados Processados')
            
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
