#importando as funções.
import funcoes
print('Bem vindo a a o seu assistente de investimentos.')
#definindo a função principal
def main():
  x=0
  while x<4:
    print('Digite 1: para pesquisar ações específicas \n Digite 2: para pesquisar várias definidas por você.\nDigite 3: para agendar execuções \n Digite 4: para sair.')
    x=int(input('escolha sua opção e precione enter.'))
    while x>4 or x<0:#verificando se a opção está correta.
      print('digite uma opção maior que 0 e menor ou igual a 4')
      x=int (input())
    if x == 1:
      codigo=(input('digite o código da ação igual no site yahoo e precione enter'))
      datai=(input('digite a data inicial no formato:MDA, separando por / ou -'))
      verifica=len(datai)
      while((verifica>10)or (verifica <10)):
        datai=(input('digite uma data valida'))
        verifica=len(datai)
      dataf=(input('digite a dada final, separando por / ou -, no formato MDA.'))
      verifica=len(dataf)
      while((verifica>10)or (verifica <10)):
        dataf=(input('digite uma data valida'))
        verifica=len(dataf)
      funcoes.buscar_acao(codigo,datai,dataf)#chamando a função de buscar ação individual.
    elif x==2:
      datai=(input('digite a data inicial no formato:MDA, separando por / ou -'))
      verifica=len(datai)
      while((verifica>10)or (verifica <10)):
        datai=(input('digite uma data valida'))
        verifica=len(datai)
      dataf=(input('digite a dada final, separando por / ou -, no formato MDA.'))
      verifica=len(dataf)
      while((verifica>10)or (verifica <10)):
        dataf=(input('digite uma data valida'))
        verifica=len(dataf)
      funcoes.acoes_arquivo(datai,dataf)#chamando a função de buscar vários códigos.
    elif x==3:
      #chamando a função de executar verificações programadas
      funcoes.agendar()
    elif x==4:
      #saindo do script
      print('Obrigado por usar este  script.')
      break
main()
