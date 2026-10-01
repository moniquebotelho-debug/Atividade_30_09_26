import openpyxl

# 1. Carrega o arquivo Excel
wb = openpyxl.load_workbook('c:\Users\Monique\Desktop\Monique\Atividade\estimativa_dou_2025.xls', data_only=True)

# 2. Seleciona a planilha ativa (ou use wb['Nome da Aba'])
aba = wb.active

# 3. Percorre as linhas e colunas exibindo os valores
for linha in aba.iter_rows(values_only=True):
    # Ignora linhas completamente vazias
    if any(linha):
        print(linha)
