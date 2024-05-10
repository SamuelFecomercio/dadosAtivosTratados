from io import StringIO
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from datetime import datetime

with open('Associados.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

soup = BeautifulSoup(html_content, 'html.parser')

table = soup.find('table')

df = pd.read_html(str(table))[0]

df_ativos = df[df['SITUAÇÃO'] == 'ATIVO']

df_ativos_filtered = df_ativos[['CÓDIGO', 'NOME', 'CPF']]

# Renomeando a coluna "CÓDIGO" para "IDENTIFICACAO"
df_ativos_filtered.rename(columns={'CÓDIGO': 'IDENTIFICACAO'}, inplace=True)

df_ativos_filtered.insert(0, 'PLANO', 32305)
df_ativos_filtered['LIMITE'] = 0

# Adicionando um zero no início para CPFs com 10 dígitos
df_ativos_filtered['CPF'] = df_ativos_filtered['CPF'].apply(lambda x: '{:011}'.format(x))

# Removendo o negrito dos títulos das colunas
font = Font(bold=False)
df_ativos_filtered.style.set_caption('Dados Ativos').set_table_styles([{
    'selector': 'th',
    'props': [
        ('font-weight', 'normal'),
    ]
}])

# Salvando os dados em um arquivo Excel
current_date = datetime.now().strftime("%d-%m-%Y")  # Formatando a data como DD-MM-AAAA
file_name = f'dados_ativos_{current_date}.xlsx'

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df_ativos_filtered.to_excel(writer, index=False, sheet_name='Dados Ativos')
    workbook = writer.book
    worksheet = writer.sheets['Dados Ativos']
    for cell in worksheet['1:1']:
        cell.font = font
    
print(f"Dados salvos em '{file_name}' com sucesso!")
