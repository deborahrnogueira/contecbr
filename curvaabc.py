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
        valor_limpo = ''.join(c for c in valor_str if c.isdigit() or c in '.,')
        valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
        return float(valor_limpo)
    except:
        return 0.0

def formatar_moeda(valor):
    """Formata valor para R$ com separadores de milhares"""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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
        df_processado['TOTAL_FORMATADO'] = df_processado['TOTAL'].apply(formatar_moeda)
        
        return df_processado
        
    except Exception as e:
        raise Exception(f"Erro ao processar dados: {str(e)}")

def criar_graficos(df):
    """Cria as visualizações dos dados"""
    graficos = {}
    
    # Gráfico de Pareto
    fig_pareto = go.Figure()
    
    # Barras de incidência
    fig_pareto.add_trace(go.Bar(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA DO ITEM (%)'],
        name='Incidência',
        marker_color=df['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'}),
        hovertemplate="<b>Item: %{text}</b><br>" +
                     "Incidência: %{y:.2f}%<br>" +
                     "Valor: %{customdata[0]}<br>" +
                     "Classificação: %{customdata[1]}<extra></extra>",
        text=df['ITEM'],
        customdata=np.column_stack((df['TOTAL_FORMATADO'], df['CLASSIFICAÇÃO']))
    ))
    
    # Linha de acumulado
    fig_pareto.add_trace(go.Scatter(
        x=df['NÚMERO DO ITEM'],
        y=df['INCIDÊNCIA DO ITEM (%) ACUMULADO'],
        name='Acumulado',
        line=dict(color='red', width=2),
        yaxis='y2',
        hovertemplate="<b>Acumulado: %{y:.2f}%</b><extra></extra>"
    ))
    
    # Linhas de referência
    fig_pareto.add_hline(y=80, line_dash="dash", line_color="red", opacity=0.5,
                        annotation_text="Limite A (80%)")
    fig_pareto.add_hline(y=95, line_dash="dash", line_color="orange", opacity=0.5,
                        annotation_text="Limite B (95%)")
    
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
        title='Distribuição por Classe',
        hover_data=['ITEM', 'Percentual', 'TOTAL_FORMATADO']
    )
    graficos['pizza'] = fig_pizza
    
    return graficos

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
                graficos = criar_graficos(df_processado)
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.plotly_chart(graficos['pareto'], use_container_width=True)
                with col2:
                    st.plotly_chart(graficos['pizza'], use_container_width=True)
                
                # Resumo
                st.subheader('Resumo por Classe')
                resumo = df_processado.groupby('CLASSIFICAÇÃO').agg({
                    'ITEM': 'count',
                    'TOTAL': ['sum', 'mean']
                })
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
