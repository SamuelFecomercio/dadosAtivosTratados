from io import StringIO
from bs4 import BeautifulSoup
import pandas as pd

with open('Associados.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

soup = BeautifulSoup(html_content, 'html.parser')

table = soup.find('table')

df = pd.read_html(str(table))[0]

df_ativos = df[df['SITUAÇÃO'] == 'ATIVO']

df_ativos_filtered = df_ativos[['CÓDIGO', 'NOME', 'CPF']]

df_ativos_filtered.insert(0, 'PLANO', 32305)
df_ativos_filtered['LIMITE'] = 0

df_ativos_filtered['CPF'] = df_ativos_filtered['CPF'].apply(lambda x: '{:011d}'.format(x))

df_ativos_filtered.to_excel('dados_ativos.xlsx', index=False)

print("Dados salvos em 'dados_ativos.xlsx' com sucesso!")
