# importando os dados e a biblioteca
import dados_clientes
from openpyxl import *

# colocando os dados em uma varivael
tabela = dados_clientes.dados
# lendo a planilha em excel
wb = load_workbook(filename='ci_de_devolucao.xlsx')
# ativando a tabela que vou usar
planilha = wb.active
# numeração da ultima comanda de devolucao
numero_salvo = int(input("Digite o numero do ultimo codigo de devolucao que voce fez: "))

# funçao para o nivel 1
def nivel1():
    # primeira parte da CI de cima

    # entradas do usuario
    print('====================')
    codigo_cliente = input('Digite o codigo do cliente: ')
    notas_referentes = input('Digite a nota referente: ')
    laudo_nota_cliete = int(input('Digite o laudo ou nota do cliente: '))
    valor = float(input('Qual o valor a ser devolvido?: '))
    print('====================')

    # dados do dicionarios dados clietes
    novo_dict = tabela.get(codigo_cliente)
    cod_vendedor = novo_dict['vendedor']
    cod_nome = novo_dict['nome']

    planilha['E7'] = codigo_cliente + ' - ' + cod_nome
    planilha['C9'] = laudo_nota_cliete
    planilha['E9'] = valor
    planilha['D5'] = cod_vendedor
    planilha['D3'] = notas_referentes

    primeiro_cliente = cod_nome

    # segunda parte de baixo da CI

    # entredas do usuario
    print('====================')
    codigo_cliente = input('Digite o codigo do cliente: ')
    notas_referentes = input('Digite a nota referente: ')
    laudo_nota_cliete = int(input('Digite o laudo ou nota do cliente: '))
    valor = float(input('Qual o valor a ser devolvido?: '))
    print('====================')

    # dados do dicionarios dados clietes
    novo_dict = tabela.get(codigo_cliente)
    cod_vendedor = novo_dict['vendedor']
    cod_nome = novo_dict['nome']

    planilha['E28'] = codigo_cliente + ' - ' + cod_nome
    planilha['C30'] = laudo_nota_cliete
    planilha['E30'] = valor
    planilha['D26'] = cod_vendedor
    planilha['D24'] = notas_referentes

    wb.save(f'{numero_salvo} - {primeiro_cliente} - {cod_nome}.xlsx ')
    


# Nivel 2

def nivel2():

    # entradas do usuario
    print('====================')
    codigo_cliente = input('Digite o codigo do cliente: ')
    notas_referentes = input('Digite a nota referente: ')
    laudo_nota_cliete = int(input('Digite o laudo ou nota do cliente: '))
    valor = float(input('Qual o valor a ser devolvido?: '))
    print('====================')

    # dados do dicionarios dados clietes
    novo_dict = tabela.get(codigo_cliente)
    cod_vendedor = novo_dict['vendedor']
    cod_nome = novo_dict['nome']

    planilha['E7'] = codigo_cliente + ' - ' + cod_nome
    planilha['C9'] = laudo_nota_cliete
    planilha['E9'] = valor
    planilha['D5'] = cod_vendedor
    planilha['D3'] = notas_referentes

    planilha['E28'] = codigo_cliente + ' - ' + cod_nome
    planilha['C30'] = laudo_nota_cliete
    planilha['E30'] = valor
    planilha['D26'] = cod_vendedor
    planilha['D24'] = notas_referentes

    wb.save(f'{numero_salvo} - {cod_nome}.xlsx ')


def menu():
    while True:
        print("***************************************")
        print("*******Escolha o metodo de busca*******")
        print("***************************************")
        print("(1) nivel 1 (2) nivel 2 (3) Sair")
        print('')
        escolha = int(input("Qual e sua escolha? "))

        if (escolha == 1):
            nivel1()
        elif (escolha == 2):
            nivel2()
        elif (escolha == 3):
            break


menu()
