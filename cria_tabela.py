import pandas as pd
import random
from openpyxl import Workbook
from pathlib import Path

# Definindo as listas de valores para Loja, Item, Tamanho
lojas = [f'Loja {i}' for i in range(1, 11)]
itens = ['Camiseta', 'Calça', 'Shorts', 'Jaqueta', 'Vestido', 'Saia', 'Blusa', 'Moletom', 'Regata', 'Bermuda',
         'Pijama', 'Macacão', 'Blazer', 'Cachecol', 'Top', 'Legging', 'Sapato', 'Chinelo', 'Tênis', 'Sandália']
tamanhos = ['PP', 'P', 'M', 'G', 'GG', 'EGG']

# Criando DataFrame com 500 linhas
data = {'Loja': [], 'Item': [], 'Tamanho': [], 'Quantidade': [], 'Valor': []}

for _ in range(5000):
    data['Loja'].append(random.choice(lojas))
    data['Item'].append(random.choice(itens))
    data['Tamanho'].append(random.choice(tamanhos))
    data['Quantidade'].append(random.uniform(1, 10))
    data['Valor'].append(round(random.uniform(10, 100), 2))

df = pd.DataFrame(data)

# Salvando o DataFrame em um arquivo Excel na Área de Trabalho
desktop_path = str(Path.home() / 'C:\\Users\\nanni\\OneDrive\\Área de Trabalho')
excel_path = f'{desktop_path}/tabela_excel.xlsx'
df.to_excel(excel_path, index=False, engine='openpyxl')
print(f'Tabela salva em: {excel_path}')