import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel(r'C:\Users\ausna\VS CODE - CODS\Treinamento Python\projeto_1\vv_p1.xlsx')

print('-' * 25, 'Faturamento por loja', '-' * 25)

faturamento_loja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento_loja)

print('-' * 25, 'Quantidade por loja', '-' * 25)

quantidade_loja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_loja)

print('-' * 25, 'Ticket Médio por loja', '-' * 25)

ticket_medio = (faturamento_loja['Valor Final'] / quantidade_loja['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={
    0 : 'Ticket Médio'
})
print(ticket_medio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'depaulanascimentoj@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada loja:</p>

<p>Faturamento:</p>
{faturamento_loja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade:</p>
{quantidade_loja.to_html()}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição</p>

<p>Att, João</p>
'''

mail.Send()
print('E-mail enviado!')