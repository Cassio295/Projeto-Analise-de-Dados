import pandas as pd
from pathlib import Path
import yagmail

vendas = pd.read_excel("Bases de Dados\Vendas.xlsx")
lojas = pd.read_csv("Bases de Dados\Lojas.csv", encoding= 'latin1', sep=';')
emails = pd.read_excel("Bases de Dados\Emails.xlsx")
# print(lojas)
# print(vendas)
# print(emails)


# Criando um dataframe com as vendas e o nome da loja
vendas_lojas = vendas.merge(lojas, on='ID Loja')
# print(vendas_lojas)

# Pegando a data mais recente de venda
data = vendas['Data'].max()
dia_indicador = data.day
mes_indicador = data.month
# print(mes_indicador)

# Criando um dicionario Com as vendas de  cada loja
dic_lojas = {}
for loja in lojas['Loja']:
    dic_lojas[loja] = vendas_lojas.loc[vendas_lojas['Loja'] == loja,:]

caminho = Path.cwd()

for loja in lojas['Loja']:
    Path(f'Backup Arquivos Lojas/{loja}').mkdir()



# Criando as pastas com o nome de cada loja
caminho = Path('Backup Arquivos Lojas')
arquivos = caminho.iterdir()
arquivos = [arquivo.name for arquivo in arquivos]


# Criando uma planilha com as vendas de cada loja
for loja, vendas_lojas in dic_lojas.items():
    for arquivo in arquivos:
        caminho_arquivo = caminho / arquivo
        if loja == arquivo:
            nome_arquivo = f'{dia_indicador}_{mes_indicador}_{loja}.xlsx'
            caminho_final = caminho_arquivo / nome_arquivo
            vendas_lojas.to_excel(caminho_final, index= False)
            print(f'O arquivos {nome_arquivo} foi gerado dentro da pasta {caminho_arquivo}')
        if caminho.exists:
            print(f'O arquivo  já existe')
            break


#definição das metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# Calculando faturamento,quantidades e ticket médio anual e diario
for loja in dic_lojas:
    vendas_lojas = dic_lojas[loja]
    vendas_diarias = vendas_lojas.loc[vendas_lojas['Data']==data,:]
    # Faturamento do Ano    
    faturamento = vendas_lojas['Valor Final'].sum()
    # print(faturamento)
    
    # Faturamento Diario
    faturamento_dia = vendas_diarias['Valor Final'].sum()
    # print(faturamento_dia)

    # Quantidade de Produtos 
    qtde_produtos = len(vendas_lojas['Produto'].unique())
    # print(qtde_produtos)

    # Quantidade de Produtos por dia
    qtde_produtos_dia = len(vendas_diarias['Produto'].unique())
    # print(qtde_produtos_dia)    
    
    # Ticket Medio
    valor_venda = vendas_lojas.groupby('Código Venda').sum(numeric_only = True)
    # print(valor_venda)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    # print(ticket_medio_ano)

    valor_venda_dia = vendas_diarias.groupby('Código Venda').sum(numeric_only = True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    # print(ticket_medio_dia)


# Verificando quais lojas bateram as metas
    if faturamento >= meta_faturamento_ano:
        cor_fat_ano = 'Green'
    else:
        cor_fat_ano = 'Red'
        
    if faturamento_dia >=meta_faturamento_dia:
        cor_fat_dia = 'Green'
    else:
        cor_fat_dia = 'Red'

    if qtde_produtos >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'Green'
    else:
        cor_qtde_ano = 'Red'
        
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'Green'
    else:
        cor_qtde_dia = 'Red'
        
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'Green'
    else:
        cor_ticket_dia = 'Red'

    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'Green'
    else:
        cor_ticket_ano = 'Red'

# Configurando o disparo de email

 # Definindo o usuario para o dispado do email
    usuario = yagmail.SMTP(user='cassiosilvaf2013@gmail.com', password='qtsgzwkmkbzxoxbh')
    email = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]

    html_body = f'''
    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador}/{mes_indicador})</strong> da <strong>Loja {loja}</strong> foi:</p>

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
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
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
        <td style="text-align: center">R${faturamento:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Lira</p>
    '''

    arquivo_envio =  f'{caminho_final}'
    usuario.send(to=email, subject= 'Envio de OnePage', contents=html_body, attachments=arquivo_envio)
    print(f'Email enviado com sucesso para o gerene {nome} da loja {loja}')


# Montando dataFrame para o diretor
email_diretor = emails.loc[emails['Loja'] =='Diretoria','E-mail'].values[0]
vendas_diretores = vendas
vendas_diretores = vendas_diretores.merge(lojas, on='ID Loja')
print(vendas_diretores)



# Montando tabela de faturamento anual para os diretores
faturamento_anual = vendas_diretores[['Loja', 'Valor Final']]
vendas_ano = faturamento_anual.groupby('Loja').sum()
melhor_venda = vendas_ano['Valor Final'].idxmax()
maior_venda = vendas_ano.loc[melhor_venda]
faturamento_ano = vendas_ano.to_excel('Vendas_anuais.xlsx')

# Montando tabela de faturamento por dia
vendas_por_dia = vendas_diretores[['Data','Loja','Valor Final']]
vendas_dia = vendas_por_dia.loc[vendas_diretores['Data'] == data]
faturamento_diario = vendas_dia.groupby('Loja').sum(numeric_only=True)
# Localizando a loja com a maior venda diaria
loja_dia = faturamento_diario['Valor Final'].max()
melhor_loja_dia = faturamento_diario.loc[faturamento_diario['Valor Final']==loja_dia]

relatorio_vendas_diario = faturamento_diario.to_excel('Vendas_diarias.xlsx')

corpo_texto = f"""

        <p>Boa tarde, {nome} tudo bem? Venho por meio deste email.</p> 
        
        <p>Te reppasar os relatorios anuais das lojas juntamento com o <b>ranking das lojas</b> </p>

        <p>Sendo que a melhor loja com vendas diarias foi a loja {melhor_loja_dia}</p>

        <p>É a melhor loja do ano foi a {maior_venda}</p>

        <p>Segue em anexo ambos os relatorios de vendas diarias e anuais</p>

"""

lista_arquivos = [Path.cwd() / 'Vendas_anuais.xlsx', Path.cwd() / 'Vendas_diarias.xlsx']
usuario.send(to=email_diretor,subject='Entrega de Relatorio anual', contents= corpo_texto, attachments=lista_arquivos)
print(f'E-mail enviado com sucesso para o diretor {nome} ')