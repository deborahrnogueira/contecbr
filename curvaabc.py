import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io
from fpdf import FPDF
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# Funções auxiliares
def formatar_moeda(valor):
    """Formata valor para R$ com separadores de milhares"""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def validar_dados(df):
    """Validação robusta dos dados de entrada"""
    erros = []
    
    if df.empty:
        erros.append("A planilha está vazia")
    
    colunas_necessarias = ['Item', 'DESCRIÇÃO', 'UND', 'QTD', 'TOTAL']
    colunas_existentes = df.columns.tolist()
    
    for coluna in colunas_necessarias:
        if coluna not in colunas_existentes:
            erros.append(f"Coluna '{coluna}' não encontrada")
    
    if len(erros) > 0:
        raise ValueError("\n".join(erros))
    
    return True

def limpar_valor_monetario(valor):
    """Limpa e converte valores monetários para float"""
    if pd.isna(valor):
        return 0.0
    
    if isinstance(valor, (int, float)):
        return float(valor)
    
    try:
        # Remove caracteres não numéricos exceto . e ,
        valor_limpo = ''.join(c for c in str(valor) if c.isdigit() or c in '.,')
        valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
        return float(valor_limpo)
    except:
        return 0.0

@st.cache_data
def processar_dados(df):
    """Processa os dados e cria a análise ABC"""
    # Selecionar apenas as linhas relevantes
    df = df.iloc[11:310].copy()
    
    # Converter coluna TOTAL para float
    df['TOTAL'] = df.iloc[:, 6].apply(limpar_valor_monetario)
    
    # Remover linhas com total zero
    df = df[df['TOTAL'] > 0].reset_index(drop=True)
    
    # Ordenar por valor total
    df = df.sort_values('TOTAL', ascending=False).reset_index(drop=True)
    
    # Calcular métricas
    total_geral = df['TOTAL'].sum()
    df['INCIDÊNCIA (%)'] = (df['TOTAL'] / total_geral * 100).round(2)
    df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA (%)'].cumsum().round(2)
    
    # Classificar
    df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(
        lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
    
    df['NÚMERO DO ITEM'] = range(1, len(df) + 1)
    
    return df

def criar_graficos(df):
    """Cria diferentes visualizações dos dados"""
    graficos = {}
    
    # Gráfico de Pareto
    fig_pareto = go.Figure()
    fig_pareto.add_trace(go.Bar(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA (%)'],
        name='Incidência',
        marker_color=df['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'})
    ))
    fig_pareto.add_trace(go.Scatter(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA ACUMULADA (%)'],
        name='Acumulado',
        line=dict(color='red', width=2),
        yaxis='y2'
    ))
    fig_pareto.update_layout(
        title='Curva ABC (Pareto)',
        xaxis_title='Número do Item',
        yaxis_title='Incidência (%)',
        yaxis2=dict(
            title='Acumulado (%)',
            overlaying='y',
            side='right'
        ),
        showlegend=True
    )
    graficos['pareto'] = fig_pareto
    
    # Gráfico de Pizza
    resumo_classes = df.groupby('CLASSIFICAÇÃO').agg({
        'TOTAL': 'sum',
        'NÚMERO DO ITEM': 'count'
    }).reset_index()
    
    fig_pizza = px.pie(
        resumo_classes,
        values='TOTAL',
        names='CLASSIFICAÇÃO',
        title='Distribuição por Classe',
        hover_data=['NÚMERO DO ITEM']
    )
    graficos['pizza'] = fig_pizza
    
    # Heatmap
    matriz_heatmap = df.pivot_table(
        values='TOTAL',
        index='CLASSIFICAÇÃO',
        aggfunc=['count', 'sum', 'mean']
    )
    
    fig_heatmap = px.imshow(
        matriz_heatmap,
        title='Heatmap de Valores por Classe',
        labels=dict(x='Métricas', y='Classe', color='Valor')
    )
    graficos['heatmap'] = fig_heatmap
    
    return graficos

def gerar_pdf(df, graficos):
    """Gera relatório PDF com análises"""
    pdf = FPDF()
    pdf.add_page()
    
    # Cabeçalho
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Relatório de Análise Curva ABC', ln=True, align='C')
    pdf.ln(10)
    
    # Sumário Executivo
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Sumário Executivo', ln=True)
    pdf.set_font('Arial', '', 10)
    
    resumo = df.groupby('CLASSIFICAÇÃO').agg({
        'TOTAL': ['count', 'sum', lambda x: (x.sum() / df['TOTAL'].sum() * 100)],
    }).round(2)
    
    for classe in ['A', 'B', 'C']:
        dados_classe = resumo.loc[classe]
        pdf.multi_cell(0, 5, f"""
        Classe {classe}:
        - Quantidade de itens: {dados_classe[('TOTAL', 'count')]}
        - Valor total: {formatar_moeda(dados_classe[('TOTAL', 'sum')])}
        - Percentual do valor total: {dados_classe[('TOTAL', '<lambda_0>')]}%
        """)
    
    return pdf

def main():
    # Sidebar
    st.sidebar.title('Configurações')
    
    # Área de upload
    formatos_aceitos = ['xlsx', 'xls', 'csv']
    tipo_arquivo = st.sidebar.selectbox('Formato do arquivo', formatos_aceitos)
    uploaded_file = st.file_uploader(f"Carregue o arquivo ({', '.join(formatos_aceitos)})", type=formatos_aceitos)
    
    if uploaded_file is not None:
        try:
            # Carregar dados
            if tipo_arquivo in ['xlsx', 'xls']:
                df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            else:
                df = pd.read_csv(uploaded_file)
            
            # Validar e processar dados
            validar_dados(df)
            df_processado = processar_dados(df)
            
            # Interface principal com abas
            tab1, tab2, tab3, tab4 = st.tabs(['Visão Geral', 'Análise Detalhada', 'Filtros', 'Relatório'])
            
            # Aba 1: Visão Geral
            with tab1:
                graficos = criar_graficos(df_processado)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.plotly_chart(graficos['pareto'], use_container_width=True)
                with col2:
                    st.plotly_chart(graficos['pizza'], use_container_width=True)
                
                st.plotly_chart(graficos['heatmap'], use_container_width=True)
            
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
                    valor_min, valor_max = st.slider(
                        'Faixa de valores',
                        float(df_processado['TOTAL'].min()),
                        float(df_processado['TOTAL'].max()),
                        (float(df_processado['TOTAL'].min()), float(df_processado['TOTAL'].max()))
                    )
                
                busca = st.text_input('Buscar por descrição')
                
                df_filtrado = df_processado[
                    (df_processado['CLASSIFICAÇÃO'].isin(classes_selecionadas)) &
                    (df_processado['TOTAL'] >= valor_min) &
                    (df_processado['TOTAL'] <= valor_max)
                ]
                
                if busca:
                    df_filtrado = df_filtrado[
                        df_filtrado['DESCRIÇÃO'].str.contains(busca, case=False, na=False)
                    ]
                
                st.dataframe(df_filtrado)
            
            # Aba 4: Relatório
            with tab4:
                if st.button('Gerar Relatório PDF'):
                    pdf = gerar_pdf(df_processado, graficos)
                    pdf_output = pdf.output(dest='S').encode('latin1')
                    st.download_button(
                        label="📥 Baixar Relatório PDF",
                        data=pdf_output,
                        file_name=f"relatorio_curva_abc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf"
                    )
            
            # Botões de download
            st.sidebar.download_button(
                label="📥 Baixar Dados Processados (Excel)",
                data=df_processado.to_excel(index=False).getvalue(),
                file_name="curva_abc_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.info('Verifique se o arquivo está no formato correto e contém todos os dados necessários.')

if __name__ == '__main__':
    main()
