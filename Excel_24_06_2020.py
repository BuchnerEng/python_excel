import pandas as pd
from openpyxl import load_workbook

# Lê a pasta de trabalho Excel e pega o nome das planilhas.
caminho = 'C:\CursoPython\Basic\Ex_CNPJ\Pasta1.xlsx'
sh_list = pd.read_excel(caminho, header=None, sheet_name=None)


# Verifica em cada planilha se ela não está vazia e captura os valores selecionados.
target_data = []
for key in sh_list:
    df = pd.read_excel(caminho, header=None, sheet_name=key)
    if not df.empty:
        # Para cada grupo de dados capturado deve-se alterar os valores do iloc.
        # [linha, coluna], respectivamente.
        target_data.append([
            key,  # Planilha de origem dos dados, remova a linha caso necessário.
            df.iloc[0, 0],  # Exemplo: Célula A1
            df.iloc[1, 0],  # Exemplo: Célula A2
        ])
print(target_data)


# Criando uma nova planilha com os dados, sem alterar os dados existentes
td = pd.DataFrame(target_data)

work_book = load_workbook(caminho)
new_sheet = pd.ExcelWriter(caminho)
new_sheet.book = work_book
td.to_excel(new_sheet, sheet_name='Resultado_Final')
new_sheet.save()
new_sheet.close()
