import pandas as pd
import os
import yagmail

meta_ok = '🟢'
meta_fail = '🟠'
fat_meta_ano = 1650000
fat_meta_dia = 1000
prod_meta_ano = 120
prod_meta_dia = 4
tck_meta_ano = 500
tck_meta_dia = 500

#FAZENDO LOGIN NO SEU EMAIL
#---Para realizar esta etapa você precisa antes configurar o gmail para aceitar integrações com apps menos confiáveis
yag = yagmail.SMTP(user='SEU EMAIL AQUI',password='SUA SENHA AQUI')
conteudo = ''
assunto = ''
anexo = ''

#PUXANDO O BANCO DE DADOS
lojas_df = pd.read_csv('Bases de Dados/Lojas.csv',encoding="ISO-8859-1", sep=';')
email_df = pd.read_excel('Bases de Dados/Emails.xlsx')
vendas_df = pd.read_excel('Bases de Dados/Vendas.xlsx')

#UNIFICANDO TODAS AS PLANILHAS EM UMA SÓ
geral_df = pd.merge(vendas_df,lojas_df,how='left', on='ID Loja')
geral_df = pd.merge(geral_df,email_df,how='left', on="Loja")

#CRIA UMA PASTA PARA CADA LOJA
for loja in lojas_df['Loja']:
    dir = os.path.join('Backup Arquivos Lojas',loja)
    if os.path.exists(dir):
        pass
    else:
        os.mkdir(dir)

    #CRIAR PLANILHA PARA CADA LOJA DAS VENDAS DO ÚLTIMO DIA
    aux_anual_df = geral_df[geral_df['Loja'] == loja] #FILTRANDO DF GERAL PELO NOME DA LOJA

    ontem_data = vendas_df['Data'].drop_duplicates().nlargest(2).iloc[-1] #DEFININDO VARIAVEL COM A DATA DE ONTEM
    aux_df = aux_anual_df[aux_anual_df['Data'] == ontem_data] #CRIANDO DF SÓ COM OS DADOS DA DATA DE ONTEM

    #CRIA A PLANILHA XLSX E SALVA DENTRO DA PASTA COM O NOME DA LOJA
    aux_df.to_excel('Backup Arquivos Lojas/{}/{}_{}_{}.xlsx'.format(loja,ontem_data.day,ontem_data.month,loja))
    aux_anual_df.to_excel('Backup Arquivos Lojas/{}/Anual_{}.xlsx'.format(loja,loja))

    #CALCULANDO OS INDICADORES DIÁRIOS
    fat_diario = aux_df['Valor Final'].sum()

    group_cod_vendas_df = aux_df.groupby('Código Venda').sum() #AGRUPAR O DF PELO CÓDIGO DA VENDA
    tck_diario =  group_cod_vendas_df['Valor Final'].mean()

    div_produto_diario = aux_df['Produto'].nunique()

    # CALCULANDO OS INDICADORES ANUAIS
    fat_anual = aux_anual_df['Valor Final'].sum()

    group_cod_vendas_anual_df = aux_anual_df.groupby('Código Venda').sum()
    tck_anual = group_cod_vendas_anual_df['Valor Final'].mean()

    div_produto_anual = aux_anual_df['Produto'].nunique()

    conteudo= '### INDICADORES DIÁRIOS ###',\
        '\nLoja: {}'.format(loja),\
        '\nFaturamento/ Dia: R${}'.format(fat_diario),\
        '\nMeta Dia: R$1000 {}'.format(meta_ok if fat_diario >= fat_meta_dia else meta_fail),\
        '\nDiversidade de Produtos/ Dia: {}'.format(div_produto_diario),\
        '\nMeta Dia: 4 {}'.format(meta_ok if div_produto_diario >= prod_meta_dia else meta_fail),\
        '\nTicket Médio/ Dia: R${:,.2f}'.format(tck_diario),\
        '\nMeta Dia: R$500 {}'.format(meta_ok if tck_diario >= tck_meta_dia else meta_fail),\
        '\n=====================',\
        '\n                      ',\
        '\n### INDICADORES ANUAIS ###',\
        '\nFaturamento/ Anual: R${}'.format(fat_anual),\
        '\nMeta Ano: R$1.650.000 {}'.format(meta_ok if fat_anual >= fat_meta_ano else meta_fail),\
        '\nDiversidade de Produtos/ Ano: {}'.format(div_produto_anual),\
        '\nMeta Ano: 120 {}'.format(meta_ok if div_produto_anual >= prod_meta_ano else meta_fail),\
        '\nTicket Médio/ Ano: R${:,.2f}'.format(tck_anual),\
        '\nMeta Ano: R$500 {}'.format(meta_ok if tck_anual >= tck_meta_ano else meta_fail),\
        '\n====================='

    gerente = aux_anual_df['Gerente'].iloc[0]
    email_gerente = aux_anual_df['E-mail'].iloc[0]

    yag.send(to=email_gerente,subject="One Page dia {}. Loja {}".format(ontem_data,loja),contents=list(conteudo),attachments='Backup Arquivos Lojas/{}/{}_{}_{}.xlsx'.format(loja,ontem_data.day,ontem_data.month,loja))



##ENVIANDO RANKING PARA DIRETORIA

#CRIANDO O DATAFRAME DO RANKING ANUAL DE VENDAS DO MAIOR PARA MENOR
ranking_anual_df = geral_df.groupby('Loja')[['Loja','Valor Final']].sum()
ranking_anual_df = ranking_anual_df.sort_values(by='Valor Final',ascending=False)

#GERANDO A DATA DE HOJE
data_hoje = geral_df['Data'].max()

#CRIANDO O DATAFRAME FILTRADO PELA DATA DE HOJE
ranking_dia_df = geral_df[geral_df['Data']==data_hoje]
ranking_dia_df = ranking_dia_df.groupby('Loja')[['Loja','Valor Final']].sum()
ranking_dia_df = ranking_dia_df.sort_values(by='Valor Final', ascending=False)

#GERANDO O ARQUIVO XLSX DOS DOIS RANKINS: ANUAL E DIÁRIO
rank_dia_file = ranking_dia_df.to_excel('Backup Arquivos Lojas/{}_{}_Ranking_diario.xlsx'.format(data_hoje.month,data_hoje.day))
rank_ano_file = ranking_anual_df.to_excel('Backup Arquivos Lojas/{}_{}_Ranking_Anual.xlsx'.format(data_hoje.month,data_hoje.day))


#EXTRAINDO O EMAIL DA DIRETORIA
email_diretoria = email_df.loc[email_df['Loja']=='Diretoria','E-mail'].values[0]

#EXTRAINDO O NOME DA MELHOR E PIOR LOJA (ANUAL E DIÁRIO)
loja_top_dia = ranking_dia_df.index[0]
loja_pior_dia = ranking_dia_df.index[-1]
loja_top_ano = ranking_anual_df.index[0]
loja_pior_ano = ranking_anual_df.index[-1]

#FORMATANDO O EMAIL
assunto = 'Ranking Lojas'
corpo = f'''
Boa dia Diretoria,

A melhor loja do Ranking diário foi: {loja_top_dia} com faturamento de R${ranking_dia_df.iloc[0,0]}
A pior loja do Ranking diário foi : {loja_pior_dia}  com faturamento de R${ranking_dia_df.iloc[-1,0]}

A melhor loja do Ranking anual foi: {loja_top_ano} com faturamento de R${ranking_dia_df.iloc[0,0]}
A pior loja do Ranking anual foi : {loja_pior_ano}  com faturamento de R${ranking_dia_df.iloc[-1,0]}

Segue em anexos os arquivos em Excel.

Att..
'''

#ENVIANDO O EMAIL PARA DIRETORIA
yag.send(to=email_diretoria, subject=assunto,contents=corpo,attachments=['Backup Arquivos Lojas/{}_{}_Ranking_diario.xlsx'.format(data_hoje.month,data_hoje.day),'Backup Arquivos Lojas/{}_{}_Ranking_Anual.xlsx'.format(data_hoje.month,data_hoje.day)])





