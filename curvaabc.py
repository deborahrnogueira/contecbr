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
    
    colunas_esperadas = ['ITEM', 'SIGLA', 'DESCRIÇÃO', 'UND', 'QTD', 'TOTAL', 
                        'INCIDÊNCIA DO ITEM (%)', 'INCIDÊNCIA DO ITEM (%) ACUMULADO']
    
    # Verificar se todas as colunas necessárias existem
    colunas_existentes = df.columns.tolist()
    for coluna in colunas_esperadas:
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
    try:
        # Criar cópia do DataFrame
        df_processado = df.copy()
        
        # Garantir que TOTAL seja numérico
        df_processado['TOTAL'] = df_processado['TOTAL'].apply(limpar_valor_monetario)
        
        # Remover linhas com total zero
        df_processado = df_processado[df_processado['TOTAL'] > 0].reset_index(drop=True)
        
        # Ordenar por valor total
        df_processado = df_processado.sort_values('TOTAL', ascending=False).reset_index(drop=True)
        
        # Calcular métricas se não existirem
        total_geral = df_processado['TOTAL'].sum()
        
        if 'INCIDÊNCIA DO ITEM (%)' not in df_processado.columns:
            df_processado['INCIDÊNCIA DO ITEM (%)'] = (df_processado['TOTAL'] / total_geral * 100).round(2)
        
        if 'INCIDÊNCIA DO ITEM (%) ACUMULADO' not in df_processado.columns:
            df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'] = df_processado['INCIDÊNCIA DO ITEM (%)'].cumsum().round(2)
        
        # Classificar itens
        df_processado['CLASSIFICAÇÃO'] = df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'].apply(
            lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
        
        # Adicionar número do item
        df_processado['NÚMERO DO ITEM'] = range(1, len(df_processado) + 1)
        
        # Formatar valores monetários para exibição
        df_processado['TOTAL_FORMATADO'] = df_processado['TOTAL'].apply(formatar_moeda)
        
        return df_processado
        
    except Exception as e:
        raise Exception(f"Erro ao processar dados: {str(e)}")

def criar_graficos(df):
    """Cria diferentes visualizações dos dados"""
    graficos = {}
    
    # Gráfico de Pareto
    fig_pareto = go.Figure()
    
    # Barras de incidência individual
    fig_pareto.add_trace(go.Bar(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA DO ITEM (%)'],
        name='Incidência',
        marker_color=df['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'}),
        hovertemplate="<b>Item %{x}</b><br>" +
                     "Incidência: %{y:.2f}%<br>" +
                     "Valor: %{customdata}<br>" +
                     "Classificação: %{text}<extra></extra>",
        text=df['CLASSIFICAÇÃO'],
        customdata=df['TOTAL_FORMATADO']
    ))
    
    # Linha de acumulado
    fig_pareto.add_trace(go.Scatter(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA DO ITEM (%) ACUMULADO'],
        name='Acumulado',
        line=dict(color='red', width=2),
        yaxis='y2',
        hovertemplate="<b>Item %{x}</b><br>" +
                     "Acumulado: %{y:.2f}%<extra></extra>"
    ))
    
    # Linhas de referência
    fig_pareto.add_hline(y=80, line_dash="dash", line_color="red", opacity=0.5,
                        annotation_text="Limite A (80%)", annotation_position="right")
    fig_pareto.add_hline(y=95, line_dash="dash", line_color="orange", opacity=0.5,
                        annotation_text="Limite B (95%)", annotation_position="right")
    
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
        hovermode='x unified'
    )
    graficos['pareto'] = fig_pareto
    
    # Gráfico de Pizza
    resumo_classes = df.groupby('CLASSIFICAÇÃO').agg({
        'TOTAL': 'sum',
        'ITEM': 'count'
    }).reset_index()
    
    resumo_classes['Percentual'] = (resumo_classes['TOTAL'] / resumo_classes['TOTAL'].sum() * 100).round(2)
    resumo_classes['TOTAL_FORMATADO'] = resumo_classes['TOTAL'].apply(formatar_moeda)
    
    fig_pizza = px.pie(
        resumo_classes,
        values='TOTAL',
        names='CLASSIFICAÇÃO',
        title='Distribuição do Valor Total por Classe',
        hover_data=['ITEM', 'Percentual', 'TOTAL_FORMATADO'],
        custom_data=['TOTAL_FORMATADO', 'ITEM', 'Percentual']
    )
    
    fig_pizza.update_traces(
        hovertemplate="<b>Classe %{label}</b><br>" +
                     "Valor Total: %{customdata[0]}<br>" +
                     "Quantidade de Itens: %{customdata[1]}<br>" +
                     "Percentual: %{customdata[2]:.2f}%<extra></extra>"
    )
    
    graficos['pizza'] = fig_pizza
    
    # Heatmap
    heatmap_data = pd.DataFrame({
        'Classe': ['A', 'B', 'C'],
        'Qtd_Itens': df.groupby('CLASSIFICAÇÃO')['ITEM'].count(),
        'Valor_Total': df.groupby('CLASSIFICAÇÃO')['TOTAL'].sum(),
        'Valor_Medio': df.groupby('CLASSIFICAÇÃO')['TOTAL'].mean()
    }).round(2)
    
    fig_heatmap = go.Figure(data=go.Heatmap(
        z=[heatmap_data['Qtd_Itens'],
           heatmap_data['Valor_Total'],
           heatmap_data['Valor_Medio']],
        x=['A', 'B', 'C'],
        y=['Quantidade de Itens', 'Valor Total', 'Valor Médio'],
        hoverongaps=False,
        colorscale='Blues'
    ))
    
    fig_heatmap.update_layout(
        title='Heatmap de Métricas por Classe',
        xaxis_title='Classe',
        yaxis_title='Métrica'
    )
    
    graficos['heatmap'] = fig_heatmap
    
    return graficos

def gerar_relatorio_pdf(df):
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
    
    # Análise por classe
    for classe in ['A', 'B', 'C']:
        df_classe = df[df['CLASSIFICAÇÃO'] == classe]
        qtd_itens = len(df_classe)
        valor_total = df_classe['TOTAL'].sum()
        percentual = (valor_total / df['TOTAL'].sum() * 100)
        
        pdf.multi_cell(0, 5, f"""
        Classe {classe}:
        - Quantidade de itens: {qtd_itens}
        - Valor total: {formatar_moeda(valor_total)}
        - Percentual do valor total: {percentual:.2f}%
        """)
    
    return pdf

def main():
    # Sidebar
    st.sidebar.title('Configurações')
    
    # Upload do arquivo
    uploaded_file = st.sidebar.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            
            # Validar dados
            validar_dados(df)
            
            # Processar dados
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
                    pdf = gerar_relatorio_pdf(df_processado)
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
