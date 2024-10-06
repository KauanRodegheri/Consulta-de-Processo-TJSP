import selenium
import openpyxl
from time import sleep
import selenium.webdriver
from selenium.webdriver.common.by import By

#ABRIR NAVEGADOR COM SELENIUM
driver = selenium.webdriver.Chrome()
driver.get('https://esaj.tjsp.jus.br/cpopg/open.do')
sleep(1)

#ABRIR ARQUIVO XLSX COM OPENPYXL
book = openpyxl.load_workbook('teste_tje/processos.xlsx')
pagina = book['Planilha1']

#PERCORRE AS LINHAS DESEJADAS NA PLANILHA

for linha in pagina.iter_rows(min_row=2):
    if not linha[0].value:
        print('\033[31mCOLUNA DE PROCESSOS NÃO POSSUI NENHUM VALOR\033[m')
        break
    processo_1 = linha[0].value
    status = linha[2].value
    sleep(1)
    
    #PASSO A PASSO PARA INSERIR INFORMAÇÕES NO SITE E CONFIRMAR
    click = driver.find_element(By.ID, 'radioNumeroAntigo').click()
    sleep(0.5)
    n_processo = driver.find_element(By.ID, 'nuProcessoAntigoFormatado')
    sleep(2)
    n_processo.clear()
    sleep(0.5)
    n_processo.send_keys(processo_1)
    sleep(2)
    driver.find_element(By.ID, 'botaoConsultarProcessos').click()
    sleep(2)

    #CAPTURANDO INFORMAÇÕES DOS PROCESSOS E SALVANDO NA PLANILHA
    data_processo = driver.find_element(By.XPATH, "//td[@class='dataMovimentacao']").text
    status_processo = driver.find_element(By.XPATH, "//td[@class='descricaoMovimentacao']").text.title()
    status_processo = status_processo.replace(' ','').split()[0]
    palavra = ''
    for letra in status_processo:
        if letra.isupper():
            letra = ' ' + letra
        palavra+=letra
    if linha[0].value == processo_1:
        linha[2].value = str(palavra)
        linha[3].value = str(data_processo)
    book.save('teste_tje/processos.xlsx')
    driver.get('https://esaj.tjsp.jus.br/cpopg/open.do')
    continue

    
    
    
