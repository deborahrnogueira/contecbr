import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# [Funções anteriores permanecem iguais até o main()]

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
