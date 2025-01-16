import streamlit as st
import pandas as pd
import plotly.graph_objects as go

def criar_curva_abc(df):
    # Ordenar por INCIDÊNCIA DO ITEM (%) em ordem decrescente
    df = df.sort_values('INCIDÊNCIA DO ITEM (%)', ascending=False)
    
    # Calcular percentual acumulado
    df['INCIDÊNCIA ACUMULADA (%)'] = df['INCIDÊNCIA DO ITEM (%)'].cumsum()
    
    # Classificar itens
    def classificar_abc(valor):
        if valor <= 80:
            return 'A'
        elif valor <= 95:
            return 'B'
        else:
            return 'C'
    
    df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(classificar_abc)
    return df

def main():
    st.title('Análise Curva ABC')
    
    # Upload do arquivo
    uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        # Ler dados
        df = pd.read_excel(uploaded_file)
        
        # Criar curva ABC
        df_classificado = criar_curva_abc(df)
        
        # Mostrar resultados
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
            yaxis_title='Percentual Acumulado (%)',
            showlegend=True
        )
        
        st.plotly_chart(fig)
        
        # Download do resultado
        csv = df_classificado.to_csv(index=False)
        st.download_button(
            label="Download dados classificados",
            data=csv,
            file_name="curva_abc_classificada.csv",
            mime="text/csv"
        )

if __name__ == '__main__':
    main()

