import pandas as pd 
import win32com.client as win32 

# Importação da base de dados
tabela_venda = pd.read_excel('Vendas.xlsx')

# Visualização a base de dados
pd.set_option("display.max_columns",None)

# Faturamento por loja 
Faturamento_Loja = tabela_venda[['ID Loja','Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja 
Quantidade = tabela_venda[['ID Loja','Quantidade']].groupby('ID Loja').sum()

# Ticket médio por produto em cada loja 
Ticket_médio = (Faturamento_Loja['Valor Final'] / Quantidade['Quantidade']).to_frame()
Ticket_médio = Ticket_médio.rename(columns={0: ' Ticket Médio'})


# Enviar e-mail automaticamente 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'diegorp800@gmail.com '
mail.Subject = 'Relatório de vendas por loja '
mail.HTMLBody = f''' 
<p>Prezados, </p>
<p>segue o relatório de vendas por loja.</p>

<p>Faturamento:</p>
{Faturamento_Loja.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{Quantidade.to_html()}  

<p>Ticket Médio:</p>
{Ticket_médio.to_html(formatters={'Ticket_médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou a disposição</p>
'''
mail.Send()

print("-=-"*15)
print("E-mail, enviado com sucesso! ")
print("-=-"*15)