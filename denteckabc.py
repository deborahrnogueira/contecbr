import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

def criar_curva_abc(df):
    # Selecionar apenas as linhas 12 a 310
    df = df.iloc[11:310].copy()
    
    # Limpar e converter a coluna TOTAL
    df[' TOTAL '] = df[' TOTAL '].str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float)
    
    # Ordenar por valor total em ordem decrescente
    df = df.sort_values(' TOTAL ', ascending=False)
    
    # Calcular percentuais
    total_geral = df[' TOTAL '].sum()
    df['INCIDÊNCIA DO ITEM (%)'] = (df[' TOTAL '] / total_geral) * 100
    df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA DO ITEM (%)'].cumsum()
    
    # Classificar
    def get_classe(acumulado):
        if acumulado <= 80:
            return 'A'
        elif acumulado <= 95:
            return 'B'
        else:
            return 'C'
    
    df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(get_classe)
    return df

def main():
    st.title('Análise Curva ABC')
    
    uploaded_file = st.file_uploader("Carregue a planilha Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Ler arquivo especificando a aba correta
            df = pd.read_excel(uploaded_file, sheet_name='CURVA ABC')
            
            if ' TOTAL ' not in df.columns:
                st.error("Colunas disponíveis na planilha:")
                st.write(df.columns.tolist())
                return
                
            df_classificado = criar_curva_abc(df)
            
            st.subheader('Dados Classificados')
            st.dataframe(df_classificado)
            
            # Criar gráfico
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=list(range(len(df_classificado))),
                y=df_classificado['INCIDÊNCIA ACUMULADA (%)'],
                mode='lines',
                name='Curva ABC'
            ))
            
            fig.update_layout(
                title='Curva ABC',
                xaxis_title='Itens',
                yaxis_title='Percentual Acumulado (%)'
            )
            
            st.plotly_chart(fig)
            
            # Download
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

if __name__ == '__main__':
    main()
