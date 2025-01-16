import csv
import matplotlib.pyplot as plt
from openpyxl import Workbook
import io
import streamlit as st

def criar_curva_abc(dados):
    # Processar os dados da coluna TOTAL (coluna G = índice 6)
    for linha in dados:
        linha['TOTAL'] = float(linha['TOTAL'].replace('R$ ', '').replace('.', '').replace(',', '.'))
    
    # Ordenar os dados pelo TOTAL em ordem decrescente
    dados.sort(key=lambda x: x['TOTAL'], reverse=True)
    
    # Calcular percentuais e acumulados
    total_geral = sum(linha['TOTAL'] for linha in dados)
    acumulado = 0
    for linha in dados:
        linha['INCIDÊNCIA DO ITEM (%)'] = (linha['TOTAL'] / total_geral) * 100
        acumulado += linha['INCIDÊNCIA DO ITEM (%)']
        linha['INCIDÊNCIA ACUMULADA (%)'] = acumulado
        
        # Classificar em A, B ou C
        if acumulado <= 80:
            linha['CLASSIFICAÇÃO'] = 'A'
        elif acumulado <= 95:
            linha['CLASSIFICAÇÃO'] = 'B'
        else:
            linha['CLASSIFICAÇÃO'] = 'C'
    
    return dados

def main():
    st.title('Análise Curva ABC')

    uploaded_file = st.file_uploader("Carregue a planilha CSV", type=['csv'])
    
    if uploaded_file is not None:
        try:
            # Ler o arquivo CSV
            dados = []
            with io.TextIOWrapper(uploaded_file, encoding='utf-8') as f:
                leitor_csv = csv.DictReader(f, delimiter='\t')
                for linha in leitor_csv:
                    dados.append(linha)

            # Processar os dados e criar a curva ABC
            dados_classificados = criar_curva_abc(dados)
            
            # Exibir os resultados em uma tabela
            st.subheader('Dados Classificados')
            st.write(dados_classificados)
            
            # Criar gráfico da curva ABC
            incidencias_acumuladas = [linha['INCIDÊNCIA ACUMULADA (%)'] for linha in dados_classificados]
            plt.figure(figsize=(10, 6))
            plt.plot(incidencias_acumuladas, marker='o', label="Curva ABC")
            plt.title('Curva ABC')
            plt.xlabel('Itens')
            plt.ylabel('Incidência Acumulada (%)')
            plt.grid()
            plt.legend()
            
            # Exibir o gráfico no Streamlit
            st.pyplot(plt)
            
            # Exportar os resultados para Excel
            output = io.BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "CURVA ABC"
            
            # Adicionar cabeçalhos ao Excel
            cabecalhos = list(dados_classificados[0].keys())
            ws.append(cabecalhos)
            
            # Adicionar os dados classificados ao Excel
            for linha in dados_classificados:
                ws.append(list(linha.values()))
            
            wb.save(output)
            
            # Botão de download do arquivo Excel
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
