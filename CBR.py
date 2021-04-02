#!/usr/bin/python
# -*- coding: utf-8 -*-
#################################
#
# Código desenvolvido para o trabalho cientifíco: Idade Relativa / Relative Age
# Pesquisador responsável: Sr. Sabadini de Lima, José Paulo
# uma coleta sobre os dados dos atletas no site da CBF é realizado.
# proteção de dados fora previamente esclarecidos e autorizados.
#
# Desenvolvedor: Sr. Serra Pereira, Fabio
# E-Mail: fabio.serra.pereira@usp.br
# Data: 26 de março de 2021. 18:50 horário de Berlim, Alemanha.
#
# Orientador: Prof. Dr. Massa, Marcelo
#
##################################

#################################
#
# Bibliotecas
#
##################################

import html2text
from openpyxl import Workbook
from selenium import webdriver
import os

#################################
#
# diretórios
#
##################################

chromeDriver = os.path.abspath(__file__)[:-len('\CBF.py')] + os. sep + 'chromedriver.exe'

#################################
#
# Constantes
#
##################################

login = None # Precisa ser preenchido
password = None # Precisa ser preenchido
URL = "http://sistema.remobrasil.com/"
URL_Athlet = "http://sistema.remobrasil.com/athletes/"

###########

def coletaDados(driver):
    """
    Esta função será reponsável por avaliar o source do html e retornar os dados em forma de lista para serem gravados
    no EXCEL.
    :param driver: Interação com o Browser para ler o Source da página
    :return: List
    """

    # Aqui ele cria a o leitor de HTML e transforma em uma lista de caractéries
    Reader = html2text.HTML2Text()
    Reader.ignore_images = True
    Reader.ignore_links = True
    Reader.ignore_emphasis = True
    Reader = Reader.handle(driver.page_source)

    # print(Reader)

    # Cria uma lista de retorno
    retorno, contador = list(), 0

    # Aqui ele transforma caracteres em strings (palavras) e grava numa lista

    # Para o nome
    try:
        num_aux_inicio = Reader.find("Nome completo")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Nome completo") + 2
            num_aux_final = num_aux_inicio + 40
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Para o clube
    try:
        num_aux_inicio = Reader.find("Lista de Clubes")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Lista de Clubes") + 15
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "|":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Data de nascimento
    try:
        num_aux_inicio = Reader.find("Data de nascimento")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Data de nascimento") + 2
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Categoria
    try:
        num_aux_inicio = Reader.find("Nível")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Nível") + 2
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Altura
    try:
        num_aux_inicio = Reader.find("Altura")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Altura") + 2
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Peso
    try:
        num_aux_inicio = Reader.find("Peso")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Peso") + 2
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string[:-len(' Kg')])
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    # Sexo
    try:
        num_aux_inicio = Reader.find("Sexo")
        if num_aux_inicio == -1:
            retorno.append("Not informed")
        else:
            num_aux_inicio += len("Sexo") + 2
            num_aux_final = num_aux_inicio + 100
            string = ""
            for i in range(num_aux_inicio, num_aux_final):
                if Reader[i] != "\n":
                    string += Reader[i]
                else:
                    break
            retorno.append(string)
    except UnicodeDecodeError:
        retorno.append("Not informed")
        contador += 1

    if contador < 7:
        return retorno
    else:
        return False


def Excel_Grava(ws, dados, linha):
    """
    Esta função é resposável por gravar os dados em uma planilha de excel
    :param dados: Contém os dados no momento que precisam ser gravados na planilha
    :param linha: O número da linha onde será gravados os dados de dados
    :return: none
    """
    for i in range(len(dados)):
        ws.cell(row=linha,column=i+1).value = dados[i]

#################################
#
# Função Main
#
##################################

if __name__ == '__main__':
    """
    Trata-se da main do programa, ela é responsável por estabelecer uma comunicação com o Browser e executar a leitura
    do HTML e gravar no EXCEL.
    :return: none
    """
    # Primeiro eu inicializo o automatizador
    driver = webdriver.Chrome(chromeDriver)
    driver.get(URL)

    # Login e PassWord
    driver.find_element_by_id("session_email").send_keys(login)
    driver.find_element_by_id("session_password").send_keys(password)
    driver.find_element_by_name("button").click()


    # Criar a Planilha de Excel
    wb = Workbook()
    ws = wb.active

    # Existem 51396 (2066) atletas cadastrados no site....
    # Portanto como eu não consigo pegar somente os países que queria vou pegar todos e depois posso brincar no Excel
    # 36735 - 36823 corninho do programa rodou sem pegar direito
    for i in range(1, 2498):

        # Primeiro eu preciso pegar atualizar o meu URL com os atletas:
        URL = URL_Athlet + str(i)

        # Atualiza meu Link no automatizador
        driver.get(URL)

        # Print do atual contador
        print('Atual contador: ', i)

        # Pegar os dados
        Dados_Tratados = coletaDados(driver)
        print(Dados_Tratados)

        # Atualiza no Excel
        if Dados_Tratados != False:
            Excel_Grava(ws, Dados_Tratados, i)

            # Salva na Planilha
            wb.save('CBR_Coleta_de_dados_dos_atletas.xlsx')

    driver.close()


