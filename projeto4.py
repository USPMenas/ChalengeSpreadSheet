import pandas as pd

# Carregar a planilha antiga
caminho_planilha_antiga = 'dataBase/banco_de_dados.xlsx'  # Substitua pelo caminho correto da planilha antiga
df_antiga = pd.read_excel(caminho_planilha_antiga)  # Ajuste o skiprows conforme necessário

# Carregar a planilha nova
caminho_planilha_nova = 'trashCan/spreedSheet.xlsx'  # Substitua pelo caminho correto da planilha nova
df_nova = pd.read_excel(caminho_planilha_nova, skiprows=13)

# Ajuste a seleção de colunas com base nos nomes corretos
colunas_desejadas = ['Data/hora', 'Local', 'Produto', 'Quantidade', 'Valor (R$)']  # Ajuste os nomes conforme necessário
df_antiga = df_antiga[colunas_desejadas]
df_nova = df_nova[colunas_desejadas]

# Remover ' un' da coluna 'Quantidade' e converter para inteiro em ambas as planilhas
df_nova['Quantidade'] = df_nova['Quantidade'].str.replace(' un', '').astype(int)


# Converter a coluna 'Data/hora' para datetime para garantir que a comparação seja feita corretamente
df_antiga['Data/hora'] = pd.to_datetime(df_antiga['Data/hora'])
df_nova['Data/hora'] = pd.to_datetime(df_nova['Data/hora'])

# Filtrar os dados da nova planilha que não estão na antiga (comparando apenas as datas)
novos_dados = df_nova[~df_nova['Data/hora'].isin(df_antiga['Data/hora'])]

# Adicionar os novos dados à planilha antiga
df_atualizado = pd.concat([novos_dados, df_antiga], ignore_index=True)

# Salvar a planilha antiga com os novos dados adicionados
novo_caminho_planilha_antiga = 'dataBase/banco_de_dados.xlsx'
df_atualizado.to_excel(novo_caminho_planilha_antiga, index=False)

print(f"Planilha atualizada!'")
