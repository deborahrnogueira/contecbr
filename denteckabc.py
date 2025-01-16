import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

def criar_curva_abc(df):
    # Usar os dados da planilha fornecida
    df = df.copy()
    
    # Calcular percentuais
    total_geral = df[' TOTAL '].str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float).sum()
    df['INCIDÊNCIA DO ITEM (%)'] = df[' TOTAL '].str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float) / total_geral * 100
    
    # Ordenar por incidência em ordem decrescente
    df_ordenado = df.sort_values('INCIDÊNCIA DO ITEM (%)', ascending=False)
    
    # Calcular incidência acumulada
    df_ordenado['INCIDÊNCIA ACUMULADA (%)'] = df_ordenado['INCIDÊNCIA DO ITEM (%)'].cumsum()
    
    # Classificar
    def classificar_abc(valor):
        if valor <= 80:
            return 'A'
        elif valor <= 95:
            return 'B'
        else:
            return 'C'
    
    df_ordenado['CLASSIFICAÇÃO'] = df_ordenado['INCIDÊNCIA ACUMULADA (%)'].apply(classificar_abc)
    
    return df_ordenado

def main():
    st.title('Análise Curva ABC')
    
    uploaded_file = st.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Ler arquivo da aba CURVA ABC
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            
            # Verificar colunas
            if ' TOTAL ' not in df.columns:
                st.error("Colunas disponíveis:")
                st.write(df.columns.tolist())
                return
            
            # Processar dados
            df_classificado = criar_curva_abc(df)
            
            # Mostrar resultados
            st.subheader('Dados Classificados')
            st.dataframe(df_classificado)
            
            # Criar gráfico da curva ABC
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=list(range(len(df_classificado))),
                y=df_classificado['INCIDÊNCIA ACUMULADA (%)'],
                mode='lines+markers',
                name='Curva ABC'
            ))
            
            fig.update_layout(
                title='Curva ABC',
                xaxis_title='Itens',
                yaxis_title='Percentual Acumulado (%)'
            )
            
            st.plotly_chart(fig)
            
            # Download dos resultados
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_classificado.to_excel(writer, sheet_name='CURVA ABC', index=False)
            
            st.download_button(
                label="Baixar Curva ABC",
                data=output.getvalue(),
                file_name="curva_abc_classificada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f'Erro ao processar arquivo: {str(e)}')
            st.write("Detalhes do erro:", e)

if __name__ == '__main__':
    main()
