import pandas as pd
import pathlib
import win32com.client as win32

# Importing the data
emails = pd.read_excel(r'C:\Users\nicho\Downloads\Projeto1\Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'C:\Users\nicho\Downloads\Projeto1\Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'C:\Users\nicho\Downloads\Projeto1\Bases de Dados\Vendas.xlsx')

# Create a table for each shop and define the date of the indicator
vendas = vendas.merge(lojas, on='ID Loja')
dic_lojas = {}

# Definicao de metas
meta_fat_dia = 1000
meta_fat_ano = 1650000
meta_qntd_dia = 4
meta_qntd_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_ano = 500

# go to every shop
for loja in lojas['Loja']:
    dic_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]
# Last day and the last year on the Excel Table
dia_indicador = vendas['Data'].max()
# Save all of them in a backup
backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_backup = backup.iterdir()
lista_backup = []

for arquivo in arquivos_backup:
    lista_backup.append(arquivo.name)
# Criar uma pasta para cada loja no meu dic loja
for loja in dic_lojas:
    if loja not in lista_backup:
        nova_pasta = backup / loja
        nova_pasta.mkdir()
    # Nomeando o arquivo
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = backup / loja / nome_arquivo
    # Transformar em excel
    dic_lojas[loja].to_excel(local_arquivo)

# Calculate the indicator for the shops
for loja in dic_lojas:
    vendas_loja = dic_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    # Faturamento
    fat_ano = vendas_loja['Valor Final'].sum(numeric_only=True)
    fat_dia = vendas_loja_dia['Valor Final'].sum(numeric_only=True)

    # Diversidade de produtos
    qntd_ano = len(vendas_loja['Produto'].unique())
    qntd_dia = len(vendas_loja_dia['Produto'].unique())

    # Ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    # Ticket medio dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    # Send the email to each manager
    outlook = win32.Dispatch('outlook.application')
    nome = str(emails.loc[emails['Loja'] == loja, 'Gerente'].values[0])
    mail = outlook.CreateItem(0)
    mail.to = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    # Cores
    if fat_dia >= meta_fat_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if fat_ano >= meta_fat_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qntd_dia >= meta_qntd_dia:
        cor_qntd_dia = 'green'
    else:
        cor_qntd_dia = 'red'
    if qntd_ano >= meta_qntd_ano:
        cor_qntd_ano = 'green'
    else:
        cor_qntd_ano = 'red'

    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticket_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
<p>Bom dia, {nome}</p>
<p>Segue o resultado da loja {loja}, na data de <strong>{dia_indicador.day}/{dia_indicador.month}</strong>.</p>
<table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${fat_dia:.2f}</td>
        <td style="text-align: center">R${meta_fat_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_dia}</td>
        <td style="text-align: center">{meta_qntd_dia}</td>
        <td style="text-align: center"><font color="{cor_qntd_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
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
        <td style="text-align: center">R${fat_ano:.2f}</td>
        <td style="text-align: center">R${meta_fat_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_ano}</td>
        <td style="text-align: center">{meta_qntd_ano}</td>
        <td style="text-align: center"><font color="{cor_qntd_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticket_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
    </tr>
</table>

<p>Segue anexado a tabela em excel para mais detalhes.</p>
<p>Att., Nicholas</p>

    '''

    # Anexos
    attachment = pathlib.Path.cwd() / backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f'E-mail da loja {loja}, enviado com sucesso!')

# Create the ranking for the head business
faturamento_lojas_ano = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by='Valor Final', ascending=False)

# Nomeando o arquivo Ranking Anual
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx'
# Transformar em excel
faturamento_lojas_ano.to_excel(rf'Backup Arquivos Lojas\{nome_arquivo}')

vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

# Nomeando o arquivo Ranking Diario
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
# Transformar em excel
faturamento_lojas_dia.to_excel(rf'Backup Arquivos Lojas\{nome_arquivo}')

# Enviando E-mail para a diretoria
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}
Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Att.,
Nicholas
'''
# Anexos
attachment = pathlib.Path.cwd() / backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('Vasco da Gama')
