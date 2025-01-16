import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
import numpy as np
from fpdf import FPDF
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# Estilo do Streamlit
st.markdown("""
    <style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Função para validar o arquivo de entrada
@st.cache_data
def validate_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        required_columns = ['Item', 'DESCRIÇÃO', 'UND', 'QTD', 'TOTAL']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return False, f"Colunas ausentes: {', '.join(missing_columns)}"
        
        return True, df
    except Exception as e:
        return False, f"Erro ao ler arquivo: {str(e)}"

# Função para processar os dados
@st.cache_data
def criar_curva_abc(df):
    """Cria a análise da curva ABC a partir do DataFrame"""
    # Selecionar apenas as linhas relevantes
    df = df.iloc[11:310].copy()
    
    # Converter coluna TOTAL para numérico, tratando diferentes formatos
    df['TOTAL'] = df['TOTAL'].apply(lambda x: str(x) if pd.notnull(x) else "0")
    df['TOTAL'] = df['TOTAL'].apply(lambda x: 
        float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip())
        if isinstance(x, (str, int, float)) else 0
    )
    
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

# Função para criar gráfico interativo com Plotly
def criar_grafico_interativo(df):
    # Gráfico de barras + linha
    fig = go.Figure()
    
    # Adicionar barras para cada classe
    cores = {'A': 'blue', 'B': 'orange', 'C': 'green'}
    for classe in ['A', 'B', 'C']:
        mask = df['CLASSIFICAÇÃO'] == classe
        fig.add_trace(go.Bar(
            x=df[mask]['NÚMERO DO ITEM'],
            y=df[mask]['INCIDÊNCIA DO ITEM (%)'],
            name=f'Classe {classe}',
            marker_color=cores[classe],
            opacity=0.3,
            hovertemplate="<b>Item %{x}</b><br>" +
                         "Incidência: %{y:.2f}%<br>" +
                         "Classe: " + classe + "<extra></extra>"
        ))
    
    # Adicionar linha de incidência acumulada
    fig.add_trace(go.Scatter(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA ACUMULADA (%)'],
        name='Incidência Acumulada',
        line=dict(color='red', width=2),
        hovertemplate="<b>Item %{x}</b><br>" +
                     "Incidência Acumulada: %{y:.2f}%<extra></extra>"
    ))
    
    # Adicionar linhas de referência
    fig.add_hline(y=80, line_dash="dash", line_color="red", opacity=0.5)
    fig.add_hline(y=95, line_dash="dash", line_color="orange", opacity=0.5)
    
    fig.update_layout(
        title='Curva ABC com Distribuição de Itens',
        xaxis_title='Número de Itens',
        yaxis_title='Incidência (%)',
        hovermode='x unified',
        showlegend=True
    )
    
    return fig

# Função para criar gráfico de pizza
def criar_grafico_pizza(df):
    resumo = df.groupby('CLASSIFICAÇÃO').agg({
        'TOTAL': 'sum',
        'NÚMERO DO ITEM': 'count'
    }).reset_index()
    
    fig = px.pie(
        resumo,
        values='TOTAL',
        names='CLASSIFICAÇÃO',
        title='Distribuição por Classe',
        hover_data=['NÚMERO DO ITEM'],
        labels={'NÚMERO DO ITEM': 'Quantidade de Itens'}
    )
    return fig

# Função para criar heatmap
def criar_heatmap(df):
    pivot = pd.pivot_table(
        df,
        values='TOTAL',
        index='CLASSIFICAÇÃO',
        columns=pd.qcut(df['INCIDÊNCIA DO ITEM (%)'], q=10, labels=['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10']),
        aggfunc='sum'
    )
    
    fig = px.imshow(
        pivot,
        title='Heatmap de Concentração de Valores',
        labels=dict(x='Decil', y='Classificação', color='Valor Total')
    )
    return fig

# Função para gerar PDF
def gerar_pdf(df, figs):
    pdf = FPDF()
    
    # Primeira página - Sumário Executivo
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Análise Curva ABC - Sumário Executivo', ln=True, align='C')
    
    pdf.set_font('Arial', '', 12)
    resumo = df.groupby('CLASSIFICAÇÃO').agg({
        'TOTAL': ['count', 'sum'],
        'INCIDÊNCIA DO ITEM (%)': 'mean'
    }).round(2)
    
    # Adicionar texto do sumário
    pdf.set_font('Arial', '', 10)
    pdf.multi_cell(0, 10, f"""
    Data da análise: {datetime.now().strftime('%d/%m/%Y')}
    
    Principais conclusões:
    - Classe A: {resumo.loc['A', ('TOTAL', 'count')]} itens, representando {resumo.loc['A', ('INCIDÊNCIA DO ITEM (%)', 'mean')]}% do valor total
    - Classe B: {resumo.loc['B', ('TOTAL', 'count')]} itens, representando {resumo.loc['B', ('INCIDÊNCIA DO ITEM (%)', 'mean')]}% do valor total
    - Classe C: {resumo.loc['C', ('TOTAL', 'count')]} itens, representando {resumo.loc['C', ('INCIDÊNCIA DO ITEM (%)', 'mean')]}% do valor total
    """)
    
    # Adicionar gráficos
    for name, fig in figs.items():
        pdf.add_page()
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, name, ln=True)
        
        # Salvar figura temporariamente
        img_path = f"temp_{name}.png"
        fig.write_image(img_path)
        pdf.image(img_path, x=10, y=30, w=190)
    
    return pdf

def main():
    # Sidebar
    with st.sidebar:
        st.title('Configurações')
        st.markdown('### Filtros')
        
        # Placeholder para filtros
        st.markdown('#### Será aplicado aos dados carregados')
    
    # Tabs principais
    tabs = st.tabs(['Upload & Análise', 'Visualizações', 'Relatório', 'Documentação'])
    
    with tabs[0]:
        st.title('Análise Curva ABC')
        
        # Upload de arquivo
        uploaded_file = st.file_uploader(
            "Carregue a planilha (Excel ou CSV)",
            type=['xlsx', 'csv', 'xls']
        )
        
        if uploaded_file is not None:
            # Validar arquivo
            is_valid, result = validate_file(uploaded_file)
            
            if not is_valid:
                st.error(result)
            else:
                df = result
                df_classificado = criar_curva_abc(df)
                
                # Interface principal
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.subheader('Gráfico da Curva ABC')
                    fig_curva = criar_grafico_interativo(df_classificado)
                    st.plotly_chart(fig_curva, use_container_width=True)
                
                with col2:
                    st.subheader('Resumo da Classificação')
                    resumo = pd.DataFrame({
                        'Quantidade de Itens': df_classificado['CLASSIFICAÇÃO'].value_counts(),
                        'Total (R$)': df_classificado.groupby('CLASSIFICAÇÃO')['TOTAL'].sum().round(2)
                    }).reset_index()
                    resumo.columns = ['Classe', 'Quantidade de Itens', 'Total (R$)']
                    resumo['Total (R$)'] = resumo['Total (R$)'].apply(
                        lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    )
                    st.dataframe(resumo)
    
    with tabs[1]:
        if 'df_classificado' in locals():
            st.subheader('Visualizações Adicionais')
            
            viz_type = st.selectbox(
                'Selecione o tipo de visualização',
                ['Gráfico de Pizza', 'Heatmap', 'Todos']
            )
            
            if viz_type in ['Gráfico de Pizza', 'Todos']:
                st.plotly_chart(criar_grafico_pizza(df_classificado))
            
            if viz_type in ['Heatmap', 'Todos']:
                st.plotly_chart(criar_heatmap(df_classificado))
    
    with tabs[2]:
        if 'df_classificado' in locals():
            st.subheader('Relatório')
            
            if st.button('Gerar Relatório PDF'):
                figs = {
                    'Curva ABC': fig_curva,
                    'Distribuição por Classe': criar_grafico_pizza(df_classificado),
                    'Heatmap': criar_heatmap(df_classificado)
                }
                
                pdf = gerar_pdf(df_classificado, figs)
                
                # Salvar PDF em bytes
                pdf_output = pdf.output(dest='S').encode('latin1')
                
                st.download_button(
                    label="📥 Baixar Relatório PDF",
                    data=pdf_output,
                    file_name="relatorio_curva_abc.pdf",
                    mime="application/pdf"
                )
    
    with tabs[3]:
        st.subheader('Documentação')
        st.markdown("""
        ### Guia de Uso
        
        1. **Upload de Dados**
           - Formatos aceitos: Excel (.xlsx, .xls) e CSV
           - A planilha deve conter as colunas: Item, DESCRIÇÃO, UND, QTD, TOTAL
        
        2. **Interpretação dos Resultados**
           - **Classe A**: Itens mais importantes (80% do valor acumulado)
           - **Classe B**: Itens de importância intermediária (15% do valor acumulado)
           - **Classe C**: Itens menos importantes (5% do valor acumulado)
        
        3. **Visualizações**
           - **Curva ABC**: Mostra a distribuição dos itens e sua importância relativa
           - **Gráfico de Pizza**: Apresenta a proporção de cada classe
           - **Heatmap**: Indica a concentração de valores por decil
        
        4. **Relatório**
           - Gera um PDF completo com todas as análises
           - Inclui sumário executivo e recomendações
        """)

if __name__ == '__main__':
    main()
