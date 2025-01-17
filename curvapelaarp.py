import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from datetime import datetime
from fpdf import FPDF
import base64
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# Constantes
TOOLTIPS = {
    'INCIDÊNCIA DO ITEM (%)': 'Percentual que o item representa no valor total',
    'INCIDÊNCIA DO ITEM (%) ACUMULADO': 'Soma acumulada dos percentuais',
    'CLASSIFICAÇÃO': 'A: Items críticos (80% do valor)\nB: Items intermediários (15% do valor)\nC: Items menos críticos (5% do valor)'
}

def limpar_valor_monetario(valor):
    """Converte uma string de valor monetário para float"""
    try:
        if pd.isna(valor):
            return 0.0
        if isinstance(valor, str):
            # Remove R$, espaços e converte para float
            valor = valor.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        return float(valor)
    except:
        return 0.0

def formatar_moeda_real(valor):
    """Formata um número para o formato de moeda brasileira"""
    try:
        return f"R$ {valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except:
        return "R$ 0,00"

def carregar_dados(uploaded_file):
    """Carrega e prepara os dados da planilha, removendo linhas específicas"""
    try:
        # Lista de linhas a serem removidas (mantendo a referência original)
        linhas_excluir = [14, 15, 30, 59, 262]
        
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension in ['xlsx', 'xlsm']:
            # Primeiro, vamos ler todas as linhas para identificar corretamente as que precisamos remover
            df_temp = pd.read_excel(uploaded_file, sheet_name='ARP', usecols='A:Q', header=None)
            
            # Encontrar o índice real da linha que contém os cabeçalhos (normalmente linha 11)
            header_index = None
            for idx, row in df_temp.iterrows():
                if 'ITEM' in str(row[0]).upper():
                    header_index = idx
                    break
            
            if header_index is None:
                raise Exception("Não foi possível encontrar o cabeçalho da planilha")
            
            # Agora vamos ler o arquivo novamente, mas pulando até o cabeçalho
            df = pd.read_excel(
                uploaded_file,
                sheet_name='ARP',
                usecols='A:Q',
                header=header_index
            )
            
            # Converter os números das linhas originais para índices do DataFrame
            linhas_para_remover = [x - (header_index + 2) for x in linhas_excluir]
            
        elif file_extension == 'csv':
            df = pd.read_csv(uploaded_file)
            linhas_para_remover = [x - 1 for x in linhas_excluir]
        else:
            raise Exception("Formato de arquivo não suportado. Use .xlsx, .xlsm ou .csv")
        
        # Renomear colunas para garantir consistência
        df.columns = [
            'ITEM', 'SIGLA', 'DESCRIÇÃO', 'UND', 'QTD', 'MO', 'MAT', 'EQUIP', 
            'SUBTOTAL', 'BDI MO', 'BDI MAT', 'BDI EQUIP', 'EQUIP S/ BDI',
            'MO C/ BDI', 'MAT C/ BDI', 'EQUIP C/ BDI2', 'TOTAL'
        ]
        
        # Remover as linhas específicas
        df = df.drop(linhas_para_remover, errors='ignore')
        
        # Limpar e converter a coluna TOTAL
        df['TOTAL'] = df['TOTAL'].apply(limpar_valor_monetario)
        
        # Remover linhas com valores inválidos ou zero
        df = df[df['TOTAL'] > 0]
        
        # Resetar o índice
        df = df.reset_index(drop=True)
        
        # Verificação final
        if df.empty:
            raise Exception("Nenhum dado válido encontrado após o processamento")
            
        return df
        
    except Exception as e:
        raise Exception(f"Erro ao carregar arquivo: {str(e)}")

def processar_dados(df):
    """Processa os dados para análise ABC"""
    try:
        # Verificar se há dados válidos
        if df.empty:
            raise Exception("DataFrame está vazio")
            
        # Criar cópia do DataFrame
        df_processado = df.copy()
        
        # Ordenar por valor total
        df_processado = df_processado.sort_values('TOTAL', ascending=False).reset_index(drop=True)
        
        # Calcular percentuais
        total_geral = df_processado['TOTAL'].sum()
        if total_geral <= 0:
            raise Exception("Soma total dos valores é zero ou negativa")
            
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

def criar_guia_uso():
    """Cria o guia de uso da ferramenta"""
    guia = """
    # Guia de Uso da Ferramenta de Análise Curva ABC
    
    ## 1. Carregamento de Dados
    - Use arquivos .xlsx, .xlsm ou .csv
    - Os dados devem estar na aba 'ARP'
    - O cabeçalho deve estar na linha 12
    
    ## 2. Interpretação dos Resultados
    
    ### Classificação ABC
    - Classe A: Itens que representam 80% do valor total
    - Classe B: Itens que representam 15% do valor total
    - Classe C: Itens que representam 5% do valor total
    
    ### Gráficos
    - Pareto: Mostra a distribuição dos itens e o acumulado
    - Pizza: Apresenta a proporção de cada classe
    
    ## 3. Filtros e Análises
    - Use os filtros para focar em classes específicas
    - Utilize a busca para encontrar itens específicos
    - Exporte os dados processados em Excel
    
    ## 4. Recomendações de Uso
    - Atualize os dados periodicamente
    - Foque nos itens classe A para maior impacto
    - Use os filtros para análises específicas
    """
    return guia

def gerar_pdf(df_processado, fig_pareto, fig_pizza, resumo):
    """Gera relatório PDF com análises"""
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 15)
            self.cell(0, 10, 'Relatório de Análise Curva ABC', 0, 1, 'C')
            self.ln(10)
        
        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    pdf = PDF()
    pdf.add_page()
    
    # Sumário Executivo
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Sumário Executivo', 0, 1)
    pdf.set_font('Arial', '', 10)
    
    total_itens = len(df_processado)
    total_valor = df_processado['TOTAL'].sum()
    
    pdf.multi_cell(0, 10, 
        f'Esta análise contempla {total_itens} itens, totalizando {formatar_moeda_real(total_valor)}. '
        'Os itens foram classificados em três categorias (A, B e C) de acordo com sua representatividade financeira.')
    
    # Análise de Pareto
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Análise de Pareto', 0, 1)
    
    pareto_img = BytesIO()
    fig_pareto.write_image(pareto_img, format='png', width=800, height=400)
    pdf.image(pareto_img, x=10, y=30, w=190)
    
    # Gráfico de Pizza
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Distribuição por Classe', 0, 1)
    
    pizza_img = BytesIO()
    fig_pizza.write_image(pizza_img, format='png', width=800, height=400)
    pdf.image(pizza_img, x=10, y=30, w=190)
    
    # Resumo em tabela
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Resumo por Classe', 0, 1)
    
    # Cabeçalho da tabela
    col_width = 47
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(col_width, 10, 'Classe', 1)
    pdf.cell(col_width, 10, 'Qtd Itens', 1)
    pdf.cell(col_width, 10, 'Valor Total', 1)
    pdf.cell(col_width, 10, 'Valor Médio', 1)
    pdf.ln()
    
    # Dados da tabela
    pdf.set_font('Arial', '', 10)
    for idx, row in resumo.iterrows():
        pdf.cell(col_width, 10, str(idx), 1)
        pdf.cell(col_width, 10, str(row['Quantidade de Itens']), 1)
        pdf.cell(col_width, 10, str(row['Valor Total']), 1)
        pdf.cell(col_width, 10, str(row['Valor Médio']), 1)
        pdf.ln()
    
    # Recomendações
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Recomendações', 0, 1)
    pdf.set_font('Arial', '', 10)
    
    recomendacoes = [
        f"1. Foco em itens classe A: Priorizar a gestão dos {len(df_processado[df_processado['CLASSIFICAÇÃO']=='A'])} itens que representam 80% do valor total.",
        "2. Revisão periódica: Estabelecer ciclos de revisão trimestral para itens classe A.",
        "3. Otimização de estoques: Adequar políticas de estoque de acordo com a classificação."
    ]
    
    for rec in recomendacoes:
        pdf.multi_cell(0, 10, rec)
    
    return BytesIO(pdf.output(dest='S').encode('latin1'))

def main():
    st.title('Análise Curva ABC')
    
    # Sidebar
    st.sidebar.title('Configurações')
    st.sidebar.info('Formatos aceitos: .xlsx, .xlsm e .csv')
    
    # Botão para mostrar guia
    if st.sidebar.button('📚 Mostrar Guia de Uso'):
        st.sidebar.markdown(criar_guia_uso())
    
    uploaded_file = st.sidebar.file_uploader("Carregue a planilha", type=['xlsx', 'xlsm', 'csv'])
    
    if uploaded_file is not None:
        try:
            df = carregar_dados(uploaded_file)
            df_processado = processar_dados(df)
            
            tab1, tab2, tab3, tab4 = st.tabs(['Visão Geral', 'Análise Detalhada', 'Filtros', 'Relatório'])
            
            # Aba 1: Visão Geral
            with tab1:
                # Gráfico de Pareto
                fig_pareto = go.Figure()
                
                fig_pareto.add_trace(go.Bar(
                    x=df_processado['NÚMERO DO ITEM'],
                    y=df_processado['INCIDÊNCIA DO ITEM (%)'],
                    name='Incidência',
                    marker_color=df_processado['CLASSIFICAÇÃO'].map({'A': 'blue', 'B': 'orange', 'C': 'green'}),
                    text=df_processado['ITEM'],
                    customdata=np.column_stack((
                        df_processado['TOTAL_FORMATADO'],
                        df_processado['CLASSIFICAÇÃO'],
                        df_processado['DESCRIÇÃO']
                    )),
                    hovertemplate="""
                    <b>Item:</b> %{text}<br>
                    <b>Descrição:</b> %{customdata[2]}<br>
                    <b>Valor:</b> %{customdata[0]}<br>
                    <b>Classe:</b> %{customdata[1]}<br>
                    <b>Incidência:</b> %{y:.2f}%
                    <extra></extra>
                    """
                ))
                
                fig_pareto.add_trace(go.Scatter(
                    x=df_processado['NÚMERO DO ITEM'],
                    y=df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'],
                    name='Acumulado',
                    line=dict(color='red', width=2),
                    yaxis='y2',
                    hovertemplate="<b>Acumulado:</b> %{y:.2f}%

                fig_pareto.add_trace(go.Scatter(
                    x=df_processado['NÚMERO DO ITEM'],
                    y=df_processado['INCIDÊNCIA DO ITEM (%) ACUMULADO'],
                    name='Acumulado',
                    line=dict(color='red', width=2),
                    yaxis='y2',
                    hovertemplate="<b>Acumulado:</b> %{y:.2f}%<extra></extra>"
                ))
                
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
                    height=600,
                    hovermode='x unified'
                )
                
                df_pizza = df_processado.groupby('CLASSIFICAÇÃO').agg({
                    'TOTAL': 'sum',
                    'ITEM': 'count'
                }).reset_index()
                
                df_pizza['Percentual'] = (df_pizza['TOTAL'] / df_pizza['TOTAL'].sum() * 100).round(2)
                
                fig_pizza = px.pie(
                    df_pizza,
                    values='TOTAL',
                    names='CLASSIFICAÇÃO',
                    title='Distribuição por Classe',
                    hover_data=['Percentual'],
                    custom_data=['ITEM']
                )
                
                fig_pizza.update_traces(
                    hovertemplate="<b>Classe:</b> %{label}<br>" +
                    "<b>Valor Total:</b> R$ %{value:,.2f}<br>" +
                    "<b>Percentual:</b> %{customdata[0]:.1f}%<br>" +
                    "<b>Quantidade de Itens:</b> %{customdata[1]}<extra></extra>"
                )
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.plotly_chart(fig_pareto, use_container_width=True)
                    # Botão para download do gráfico Pareto
                    if st.button('📥 Download Gráfico Pareto'):
                        buffer = BytesIO()
                        fig_pareto.write_image(buffer, format='png')
                        st.download_button(
                            label="💾 Salvar Gráfico Pareto",
                            data=buffer.getvalue(),
                            file_name="pareto.png",
                            mime="image/png"
                        )
                
                with col2:
                    st.plotly_chart(fig_pizza, use_container_width=True)
                    # Botão para download do gráfico Pizza
                    if st.button('📥 Download Gráfico Pizza'):
                        buffer = BytesIO()
                        fig_pizza.write_image(buffer, format='png')
                        st.download_button(
                            label="💾 Salvar Gráfico Pizza",
                            data=buffer.getvalue(),
                            file_name="pizza.png",
                            mime="image/png"
                        )
                
                st.subheader('Resumo por Classe')
                resumo = df_processado.groupby('CLASSIFICAÇÃO').agg({
                    'ITEM': 'count',
                    'TOTAL': ['sum', 'mean']
                }).round(2)
                resumo.columns = ['Quantidade de Itens', 'Valor Total', 'Valor Médio']
                
                # Formatar valores monetários no resumo
                resumo['Valor Total'] = resumo['Valor Total'].apply(formatar_moeda_real)
                resumo['Valor Médio'] = resumo['Valor Médio'].apply(formatar_moeda_real)
                
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
                        default=['A', 'B', 'C'],
                        help=TOOLTIPS['CLASSIFICAÇÃO']
                    )
                
                with col2:
                    valor_min = float(df_processado['TOTAL'].min())
                    valor_max = float(df_processado['TOTAL'].max())
                    valores = st.slider(
                        'Faixa de valores',
                        min_value=valor_min,
                        max_value=valor_max,
                        value=(valor_min, valor_max),
                        format=lambda x: formatar_moeda_real(x)
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

            # Aba 4: Relatório
            with tab4:
                st.subheader('Relatório Completo')
                
                # Sumário Executivo
                st.markdown('### Sumário Executivo')
                total_itens = len(df_processado)
                total_valor = df_processado['TOTAL'].sum()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric('Total de Itens', total_itens)
                with col2:
                    st.metric('Valor Total', formatar_moeda_real(total_valor))
                with col3:
                    st.metric('Média por Item', 
                             formatar_moeda_real(total_valor/total_itens))
                
                # Análise por Classe
                st.markdown('### Análise por Classe')
                for classe in ['A', 'B', 'C']:
                    df_classe = df_processado[df_processado['CLASSIFICAÇÃO'] == classe]
                    st.markdown(f'**Classe {classe}**')
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric(f'Itens Classe {classe}', 
                                len(df_classe))
                    with col2:
                        valor_classe = df_classe['TOTAL'].sum()
                        st.metric(f'Valor Classe {classe}', 
                                formatar_moeda_real(valor_classe))
                    with col3:
                        percentual = (valor_classe/total_valor * 100)
                        st.metric(f'Percentual Classe {classe}', 
                                f'{percentual:.1f}%')
                
                # Recomendações
                st.markdown('### Recomendações')
                st.markdown("""
                1. **Itens Classe A:**
                   - Implementar controle rigoroso
                   - Revisar preços frequentemente
                   - Manter relacionamento próximo com fornecedores
                
                2. **Itens Classe B:**
                   - Estabelecer controles moderados
                   - Revisar preços periodicamente
                   - Manter estoque de segurança adequado
                
                3. **Itens Classe C:**
                   - Simplificar controles
                   - Fazer pedidos em maiores quantidades
                   - Automatizar processos de compra
                """)
                
                # Botão para gerar PDF
                if st.button('📄 Gerar Relatório PDF'):
                    pdf_buffer = gerar_pdf(
                        df_processado, 
                        fig_pareto, 
                        fig_pizza, 
                        resumo
                    )
                    st.download_button(
                        label="💾 Baixar Relatório PDF",
                        data=pdf_buffer.getvalue(),
                        file_name=f"relatorio_curva_abc_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
            
            # Botão de download do Excel com gráficos
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Dados processados
                df_processado.to_excel(writer, sheet_name='Dados Processados', index=False)
                
                # Adicionar gráficos
                workbook = writer.book
                worksheet = writer.sheets['Dados Processados']
                
                # Salvar gráficos como imagens temporárias
                pareto_img = BytesIO()
                fig_pareto.write_image(pareto_img)
                
                pizza_img = BytesIO()
                fig_pizza.write_image(pizza_img)
                
                # Inserir imagens na planilha
                worksheet.insert_image('A1', 'pareto.png', 
                                     {'image_data': pareto_img})
                worksheet.insert_image('J1', 'pizza.png', 
                                     {'image_data': pizza_img})
            
            st.sidebar.download_button(
                label="📥 Baixar Dados e Gráficos (Excel)",
                data=buffer.getvalue(),
                file_name=f"curva_abc_completa_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.info('Verifique se o arquivo está no formato correto e contém todos os dados necessários.')

if __name__ == '__main__':
    main()
