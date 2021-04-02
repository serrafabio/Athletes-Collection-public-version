#!/usr/bin/python
# -*- coding: utf-8 -*-
#################################
#
# Código desenvolvido para o trabalho cientifíco: Idade Relativa / Relative Age
# Pesquisador responsável: Sr. Sabadini de Lima, José Paulo
# uma coleta sobre os dados dos atletas no site da FISA é realizado.
# proteção de dados fora previamente esclarecidos e autorizados.
#
# Desenvolvedor: Sr. Serra Pereira, Fabio
# E-Mail: fabio.serra.pereira@usp.br
# Data: 26 de março de 2021. 22:11 horário de Berlim, Alemanha.
#
# Orientador: Prof. Dr. Massa, Marcelo
#
##################################

#################################
#
# Bibliotecas
#
##################################

from openpyxl import Workbook
from selenium import webdriver
import os, time

#################################
#
# diretórios
#
##################################

chromeDriver = os.path.abspath(os.getcwd()) + os. sep + 'chromedriver.exe'

#################################
#
# Constantes
#
##################################

URL = "https://worldrowing.com/athletes/"
sexListener = ['/html/body/div/div/main/section/div/div[2]/div[1]/div[3]/ul/li[1]',
               '/html/body/div/div/main/section/div/div[2]/div[1]/div[3]/ul/li[2]']

###########

def coletaDados(driver, sex):
    """
    Esta função será reponsável por avaliar o source do html e retornar os dados em forma de lista para serem gravados
    no EXCEL.
    :param driver: Interação com o Browser para ler o Source da página
    :return: List
    """

    # Cria uma lista de retorno
    retorno, contador = list(), 0

    # Para o nome
    try:
        retorno.append(driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[1]/div/section/header').text)
    except:
        retorno.append("Not informed")
        contador += 1

    # Para o país
    try:
        retorno.append(
            driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[1]/div/section/ul/li[2]').text)
    except:
        retorno.append("Not informed")
        contador += 1

    # Data de nascimento
    try:
        retorno.append(driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[1]/div/section/ul/li[1]').text)
    except:
        retorno.append("Not informed")
        contador += 1

    # Categoria
    try:
        retorno.append(driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[2]/div/div[1]/div/table/tbody/tr[1]/td[1]').text)
    except:
        retorno.append("Not informed")
        contador += 1

    # Altura
    try:
        altura = driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[1]/div/section/ul/li[4]').text
        if altura[-2:] in ['cm', 'CM', 'Cm', 'cM']:
            altura = altura[:-2]
        retorno.append(altura)
    except:
        retorno.append("Not informed")
        contador += 1

    # Peso
    try:
        peso = driver.find_element_by_xpath('/html/body/div/div/main/div[1]/div[1]/div[1]/div/section/ul/li[3]').text
        if peso[-2:] in ['kg', 'KG', 'Kg', 'kG']:
            peso = peso[:-2]
        retorno.append(peso)
    except:
        retorno.append("Not informed")
        contador += 1

    # Sexo
    retorno.append(sex)

    if contador < 6:
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
    driver, driver2, sexCont = webdriver.Chrome(chromeDriver), webdriver.Chrome(chromeDriver), 0

    # Criar a Planilha de Excel
    wb, lin = Workbook(), 1
    ws = wb.active

    # Seleciona o sexo
    for sex in sexListener:
        # Recarregar página
        driver.get(URL), driver.fullscreen_window(), time.sleep(2)
        # Selecionar o sexo
        driver.find_element_by_xpath('/html/body/div/div/main/section/div/div[2]/div[1]/div[3]/button').click()
        driver.find_element_by_xpath(sex).click(), time.sleep(5)

        # Inicialização de coleta
        cont = 0
        for i in range(1,100000):
            try:
                # Apertar o botão de carregar mais
                if cont == 12:
                    element = driver.find_element_by_xpath(f'/html/body/div/div/main/section/div/div[2]/div[2]/div/div[{i}]/button')
                    driver.execute_script("arguments[0].click()", element), time.sleep(5)
                    cont = 0

                # Coletar link do atleta e carregar dados no segundo drive
                urlAthlete = driver.find_element_by_xpath(f'/html/body/div/div/main/section/div/div[2]/div[2]/div/div[{i}]/a').get_attribute('href')
                driver2.get(urlAthlete), driver2.fullscreen_window()
                cont += 1

                # Pegar os dados
                if sexCont == 0:
                    Dados_Tratados = coletaDados(driver2, 'Male')
                else:
                    Dados_Tratados = coletaDados(driver2, 'Female')
                print(Dados_Tratados)

                # Atualiza no Excel
                if Dados_Tratados != False:
                    Excel_Grava(ws, Dados_Tratados, lin)

                    # Salva na Planilha
                    wb.save('FISA_Coleta_de_dados_dos_atletas.xlsx')
                lin += 1
            except:
                break
        sexCont += 1

    driver.close()
    driver2.close()


