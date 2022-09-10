#!/usr/bin/env python
# coding: utf-8

# In[18]:


import pandas as pd
import win32com.client as win32
import pathlib

emails = pd.read_excel(r'C:\autoproject\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'C:\autoproject\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'C:\autoproject\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)


# In[19]:


vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


# In[20]:


dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha'])


# In[21]:


dia_indicador = vendas['Data'].max()
print(dia_indicador)


# In[22]:



caminho_backup = pathlib.Path(r'C:\autoproject\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)


# In[23]:


meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtproduto_dia = 4
meta_qtproduto_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# In[24]:


for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]
    #Faturamento

    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)

    #Diversidade de produtos

    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    #print(qtde_produtos_ano)
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #print(qtde_produtos_dia)

    #Ticket médio

    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    #print(ticket_medio_ano)

    #Ticket médio dia

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_medio_dia)
    
    #Enviar o e-mail
    
    outlook = win32.Dispatch("outlook.application")

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    # mail.Body = 'Texto do E-mail'

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtproduto_dia:
        cor_qtde_dia ='green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtproduto_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p>Bom Dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi: </p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color ="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
       <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtproduto_dia}</td>
        <td style="text-align: center"><font color ="{cor_qtde_dia}">◙</td>
      </tr>
      <tr>
       <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color ="{cor_ticket_dia}">◙</td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color ="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
       <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtproduto_ano}</td>
        <td style="text-align: center"><font color ="{cor_qtde_ano}">◙</td>
      </tr>
      <tr>
       <td>Ticket Médio</td>
        <td style="text-align: center">{ticket_medio_ano:.2f}</td>
        <td style="text-align: center">{meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color ="{cor_ticket_ano}">◙</td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para detalhes</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Lucas</p>
    '''
    #Anexos

    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo

    attachment = pathlib.Path.cwd() / local_arquivo
    mail.Attachments.Add(str(attachment))


    mail.Send()


# In[25]:


import win32com.client as win32
outlook = win32.Dispatch("outlook.application")

nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
# mail.Body = 'Texto do E-mail'

if faturamento_dia >= meta_faturamento_dia:
    cor_fat_dia = 'green'
else:
    cor_fat_dia = 'red'
if faturamento_ano >= meta_faturamento_ano:
    cor_fat_ano = 'green'
else:
    cor_fat_ano = 'red'
if qtde_produtos_dia >= meta_qtproduto_dia:
    cor_qtde_dia ='green'
else:
    cor_qtde_dia = 'red'
if qtde_produtos_ano >= meta_qtproduto_ano:
    cor_qtde_ano = 'green'
else:
    cor_qtde_ano = 'red'
if ticket_medio_dia >= meta_ticketmedio_dia:
    cor_ticket_dia = 'green'
else:
    cor_ticket_dia = 'red'
if ticket_medio_ano >= meta_ticketmedio_ano:
    cor_ticket_ano = 'green'
else:
    cor_ticket_ano = 'red'

mail.HTMLBody = f'''
<p>Bom Dia, {nome}</p>

<p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi: </p>

<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Dia</th>
    <th>Meta Dia</th>
    <th>Cenário Dia</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td style="text-align: center">R${faturamento_dia:.2f}</td>
    <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
    <td style="text-align: center"><font color ="{cor_fat_dia}">◙</font></td>
  </tr>
  <tr>
   <td>Diversidade de Produtos</td>
    <td style="text-align: center">{qtde_produtos_dia}</td>
    <td style="text-align: center">{meta_qtproduto_dia}</td>
    <td style="text-align: center"><font color ="{cor_qtde_dia}">◙</td>
  </tr>
  <tr>
   <td>Ticket Médio</td>
    <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
    <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
    <td style="text-align: center"><font color ="{cor_ticket_dia}">◙</td>
  </tr>
</table>
<br>
<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Ano</th>
    <th>Meta Ano</th>
    <th>Cenário Ano</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td style="text-align: center">R${faturamento_ano:.2f}</td>
    <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
    <td style="text-align: center"><font color ="{cor_fat_ano}">◙</font></td>
  </tr>
  <tr>
   <td>Diversidade de Produtos</td>
    <td style="text-align: center">{qtde_produtos_ano}</td>
    <td style="text-align: center">{meta_qtproduto_ano}</td>
    <td style="text-align: center"><font color ="{cor_qtde_ano}">◙</td>
  </tr>
  <tr>
   <td>Ticket Médio</td>
    <td style="text-align: center">{ticket_medio_ano:.2f}</td>
    <td style="text-align: center">{meta_ticketmedio_ano:.2f}</td>
    <td style="text-align: center"><font color ="{cor_ticket_ano}">◙</td>
  </tr>
</table>

<p>Segue em anexo a planilha com todos os dados para detalhes</p>

<p>Qualquer dúvida estou à disposição.</p>
<p>Att., Lucas</p>
'''
#Anexos

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
local_arquivo = caminho_backup / loja / nome_arquivo

attachment = pathlib.Path.cwd() / local_arquivo
mail.Attachments.Add(str(attachment))

mail.Send()


# In[26]:


faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(r'C:\autoproject\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
faturamento_lojas_dia.to_excel(r'C:\autoproject\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))


# In[29]:


outlook = win32.Dispatch("outlook.application")

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezada {nome}, bom dia!!


Melhor Loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f} 
Pior Loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1,0]:.2f}
    
Melhor Loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}
Pior Loja do Dia em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1,0]:.2f}
    
Segue em anexo os rankings do ano e do dia de todas as lojas.
    
Qualquer dúvida estou à disposição! 
Att., Lucas

'''


#Anexos

attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))


mail.Send()
print('E-mail da diretoria enviado! ')


# In[ ]:





# In[ ]:




