import pandas as pd
import openpyxl

def criar_curva_abc(arquivo_entrada, arquivo_saida):
    # Importar dados do Excel
    df = pd.read_excel(arquivo_entrada, sheet_name='CURVA ABC')
    
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
    
    # Aplicar classificação
    df['CLASSIFICAÇÃO'] = df['INCIDÊNCIA ACUMULADA (%)'].apply(classificar_abc)
    
    # Exportar para Excel
    df.to_excel(arquivo_saida, sheet_name='CURVA ABC', index=False)
    
    # Formatar planilha
    wb = openpyxl.load_workbook(arquivo_saida)
    ws = wb['CURVA ABC']
    
    # Ajustar largura das colunas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        ws.column_dimensions[column].width = max_length + 2
    
    wb.save(arquivo_saida)

# Uso do programa
arquivo_entrada = 'planilha_original.xlsx'
arquivo_saida = 'curva_abc_classificada.xlsx'

criar_curva_abc(arquivo_entrada, arquivo_saida)
