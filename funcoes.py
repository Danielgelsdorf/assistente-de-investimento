#importando as bibliotecas.
from datetime import datetime
import win32com.client as win
import pandas as pd
from pandas_datareader import data as web
import time
import matplotlib.pyplot as plt
#criando a função de salvar o arquivo e tratar a tabela.
#Parâmetro: df é aquele que recebe o data frame.
def salvar(df):
    #Salvando os dados recebidos por parâmetros em um arquivo
    salva=pd.ExcelWriter('dados da rede.xlsx')
    df.to_excel(salva)
    salva.save()
    #Lendo, tratando e salvando o arquivo final, que será  enviado para o usuário.
    lendo=pd.read_excel('dados da rede.xlsx')
    lendo=lendo.rename(columns={'Date':'Data','High':'Maior preço','Low':'Menor preço','Adj Close':'preço do fim do dia'})
    lendo=lendo.drop(columns={'Volume','Open','Close'})
    arq=pd.ExcelWriter('Suas ações.xlsx')
    lendo.to_excel(arq)
    arq.save()
        #Recebendo o e-mail do usuário e chamando a função de enviar o e-mail.
    email=''
    email_enviar=(input('digite seu e-mail e precione enter.'))
    enviar(email_enviar)
#Definindo a função de buscar só mente uma ação
#Parâmetros: código, que é o código da ação, datai que é a data inicial das buscas e dataf que é a data que finaliza as buscas.
def buscar_acao(codigo,datai,dataf):
    #mostrando a hora que o usuário fez a busca.
    now = datetime.now()
    print('atualisado em ',now.day,now.month,now.year,'as',now.hour,now.minute)
    #buscando a ação com o pandas_datareader 
    grafico=web.DataReader(f'{codigo}.SA', data_source='yahoo', start=datai, end=dataf)
    print('ação:',codigo)
    #criando o gráfico, com o preço de fechamento.
    grafico["Adj Close"].plot(figsize=(15, 10))
    #mostrando o gráfico
    plt.show()
    salvar(grafico) #chamando a função salvar.
#Definindo a função que pega códigos de ações de um arquivo e busca as cotações.
#Parâmetros: datai, data inicial das buscas e dataf, data final.
def acoes_arquivo(datai,dataf):
    arq=pd.read_excel("Empresas.xlsx")#lendo os códigos das empresas.
    #mostrando a hora de atualização
    now = datetime.now()
    print('atualisado em ',now.day,now.month,now.year,'as',now.hour,now.minute)
    arquivo=pd.ExcelWriter('Suas ações.xlsx',engine='xlsxwriter')
    for i in arq["Empresas"]:
        acoes= web.DataReader(f'{i}.SA', data_source='yahoo', start=datai, end=dataf)
        print('ação:',i)
        acoes["Adj Close"].plot(figsize=(15, 10))
        plt.show()
        acoes=acoes.rename(columns={'Date':'Data','High':'Maior preço','Low':'Menor preço','Adj Close':'preço do fim do dia'})
        acoes=acoes.drop(columns={'Volume','Open','Close'})
        acoes.to_excel(arquivo,sheet_name=f'ação {i}')
    arquivo.save()
    email=''
    email_enviar=(input('digite seu e-mail e precione enter.'))
    enviar(email_enviar)
#Definindo a função de enviar e-mail
#Parâmetros: rec, recebe o e-mail digitado quando o programa termina de salvar o arquivo.
def enviar(rec):
    app=win.Dispatch('outlook.application')
    email=app.CreateItem(0)
    email.to=rec#e-mail do destinatário
    email.Subject="cotações de ações" #assunto.
    #criando o corpo do e-mail com html
    email.HTMLBody= f"""
    <P> olá, tudo bem</p>
    <p> segue em anexo seu arquivo com suas cotações de ações  </p>
    <p> abraços</p>
    """
    anexo="C://Users/Daniel/Desktop/trabalho/Suas ações.xlsx"#adicionando onde ta o arquivo
    email.Attachments.Add(anexo)
    email.Send()#enviando.
    print('e-mail enviado com sucesso.')
#definindo a função de agendar buscas
def agendar():
    print('digite em quantos minutos que deseja verificar suas ações')
    min=int(input('Depois de digitar precione enter'))
    while min <1:
        min=int(input('digite um tempo maior ou igual a um minuto'))
    min=min*60 #transformando os minutos em segundos
    datai=(input('digite a data inicial que deseja visualizar suas ações, a data precisa ser no formato MDA e separada por / ou -.'))
    verifica=len(datai)
    while((verifica>10)or (verifica <10)):
        datai=(input('digite uma data valida'))
        verifica=len(datai)   
    dataf=(input('digite a data final  que deseja visualizar suas ações, a data precisa ser no formato MDA e separada por / ou -.'))
    verifica=len(dataf)
    while((verifica>10)or (verifica <10)):
        dataf=(input('digite uma data valida'))
        verifica=len(dataf)
    x=0
    print('para visualizar a atualização das ações, abra o arquivo: Suas ações.xlsx')
    while x==0 :#fazendo um loop infinito
      programada(datai,dataf)
      print('para parar a verificação precione: ctrl+c')
      time.sleep(min)
def programada(datai,dataf):
    arq=pd.read_excel("Empresas.xlsx")#lendo os códigos das empresas.
    #mostrando a hora de atualização
    now = datetime.now()
    print('atualisado em ',now.day,now.month,now.year,'as',now.hour,now.minute)
    arquivo=pd.ExcelWriter('Suas ações.xlsx',engine='xlsxwriter')
    for i in arq["Empresas"]:
        acoes= web.DataReader(f'{i}.SA', data_source='yahoo', start=datai, end=dataf)
        acoes=acoes.rename(columns={'Date':'Data','High':'Maior preço','Low':'Menor preço','Adj Close':'preço do fim do dia'})
        acoes=acoes.drop(columns={'Volume','Open','Close'})
        acoes.to_excel(arquivo,sheet_name=f'ação {i}')
    arquivo.save()
