from selenium import webdriver
from time import sleep
from datetime import datetime 
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import PySimpleGUI as sg
import os
import pandas as pd

options = Options()
options.add_argument('--headless')
options.add_argument('window-size=1400,920')

def escolhe_planilha():
    SourceXlsx = None

    layout = [
       [sg.T("")],
       [sg.Text("Escolha a Planilha de SKU´s: "), sg.Input(key="-IN2-" ,change_submits=True), sg.FileBrowse(key="-IN-")],
       [sg.Submit(), sg.Cancel()]
    ]

    window = sg.Window('Lista para verificação', layout)
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event=="Exit" or event == "Cancel":
           print('É necessário escolher uma planilha')
           break
        elif event == "Submit":
           if len(values["-IN-"]) == 0:
              print('É necessário escolher uma planilha')
           else:
              SourceXlsx = values["-IN-"]
              break           
              
    window.close()
    
    return SourceXlsx

def page_scroll():
    # Pega a altura da página
    last_height = navegador.execute_script("return document.body.scrollHeight")

    while True:
        # Rola até o fim da página
        navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Espera a página carregar
        sleep(1)

        # Calcula a altura com a nova rolagem e compara com o tamanho da altura anterior
        new_height = navegador.execute_script("return document.body.scrollHeight")

        if new_height == last_height:
            break
        last_height = new_height

def insere_cep(produtos):
    for produto in produtos: 
        link = produto.find('a')['href']

        #Abrir a nova aba
        navegador.execute_script(f"window.open('{link}');")

        #Trocar pra nova aba
        navegador.switch_to.window(navegador.window_handles[1])
        try:
            element = WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[6]/app-freights/div/div/div/app-zipcode-form/div/form/div/div[1]/div/label/input')))
            element.send_keys('23040150')
            sleep(1)
            element.send_keys(Keys.ENTER)
        except:
            navegador.close()
            navegador.switch_to.window(navegador.window_handles[0])
            break
        sleep(2)
        try:
            confere_frete = navegador.find_element(by=By.XPATH, value= '//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[6]/app-freights/div/div/div[1]/div[2]/div/div/p/span[2]').text
            if '23040150' in confere_frete: 
                navegador.close()
                navegador.switch_to.window(navegador.window_handles[0])
                break
        except:
            navegador.close()
            navegador.switch_to.window(navegador.window_handles[0]) 

URL = 'https://www.fastshop.com.br/web/'

navegador = webdriver.Chrome(options=options)

navegador.get(URL)
sleep(1)

arquivos = pd.read_excel(escolhe_planilha())
arquivos = arquivos["SKU_FASTSHOP"].to_list()

variavel = navegador.find_element(by=By.XPATH, value= '//*[@id="search-box-id"]')
print(arquivos)
variavel.send_keys(str(arquivos))
sleep(1)

variavel.send_keys(Keys.ENTER)
sleep(1)

page_scroll()
sleep(0.5)

page_content = navegador.page_source
site = BeautifulSoup(page_content, 'html.parser')

produtos = site.findAll('app-product-item')
dados = []
contador = 1

insere_cep(produtos)

for produto in produtos: 
    link = produto.find('a')['href']
    bkp_sku = produto.find('a')['id']
    bkp_title = produto.find('a')['title']
    indisponivel = produto.find('p').text
    if 'Produto' in indisponivel:
        indisponivel = 'SKU indisponível para entrega'
    else:
        indisponivel = 'Disponível na tela'
    #Abrir a nova aba
    navegador.execute_script(f"window.open('{link}');")

    #Trocar pra nova aba
    navegador.switch_to.window(navegador.window_handles[1])
    sleep(3)

    #Pegar os dados da nova aba
    try:
        title = navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/h1').text
    except: 
        title = bkp_title
    try:    
        SKU = navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/span').text
    except:
        SKU = bkp_sku
    try:    
        voltagem = navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/app-sku-variations-grid/div/div/div[2]/button').text
    except:
        voltagem = 'SKU não possui voltagem ativa'
    try:    
        categoria = navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[1]/div/div[1]/app-breadcrumb/div/nav/ol/li[3]/a').text
    except:
        categoria = 'Categoria não encontrada'
    try:
        precopix = 'R$ ' + navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[2]/app-pdp-price/div/div[1]/span[2]').text + navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[2]/app-pdp-price/div/div[1]/span[3]').text
    except:
        try:
            precopix = 'R$ ' + navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[2]/app-pdp-price/div/div[2]/span[2]').text + navegador.find_element(by=By.XPATH, value='//*[@id="body"]/app-root/div[2]/div[4]/app-product/div[2]/div[1]/div[1]/div/div[2]/div/div[2]/app-pdp-price/div/div[2]/span[3]').text
        except:
            precopix = 'Preço não encontrado'
    try:
        precopadrao = navegador.find_element_by_class_name('text-bold').text
        check = precopadrao.split()
        if check[1] == str('de'):
            precopadrao = precopix
            precopix = 'Preço não encontrado'
    except:
        try:
            precopadrao = navegador.find_element_by_class_name('price-symbol') + ' ' + navegador.find_element_by_class_name('price-fraction') + navegador.find_element_by_class_name('price-cents')
        except:
            precopadrao = 'Preço não encontrado'
    try:
        precoprime = navegador.find_element_by_class_name('user-has-not').text
        precoprime = precoprime.split(' ')
        precoprime = precoprime[0] + ' ' + precoprime[1]
    except:
        precoprime = 'Preço não encontrado'        
    try:    
        imagem1 = produto.find('img')['src']
    except:
        imagem1 = 'Sem Imagem'
    try:    
        imagem2 = navegador.find_element(by=By.XPATH, value='//*[@id="scrollPriceBased"]/div/div[1]/div/swiper/div/div[1]/div[2]/div/div/img').get_attribute('src')
    except:
        imagem2 = 'Sem Imagem'

    if precopadrao == precopix and precopix == precoprime:
        frete = 'Produto fora do ar'
    else:
        try:
            frete = navegador.find_element(By.CLASS_NAME, 'error-message').text
        except:
            frete = 'Frete OK'
    try:
        page_scroll()
        sleep(0.5)
        descricao = navegador.find_element(by=By.CLASS_NAME, value='description-text').text
    except:
        descricao = 'Sem Descrição'

    contador = contador + 1

    dados.append([title, SKU, voltagem, categoria, precopadrao, precopix, precoprime, imagem1, imagem2, descricao, frete, indisponivel])

    navegador.close()
    navegador.switch_to.window(navegador.window_handles[0])

dataHora = datetime.now()
dataConvertida = dataHora.strftime('%d.%m_as_%Hh%M')
geraNomePlanilha = 'ConsultaSite_' + dataConvertida + '.xlsx'
dados_completos = pd.DataFrame(dados, columns=['Título', 'SKU_FASTSHOP', 'Voltagem', 'Categoria', 'Preço Normal', 'Preço PIX', 'Preço Prime', 'Primeira Imagem','Segunda Imagem', 'Prévia Descrição', 'Frete', 'Indisponível para entrega'])
dados_completos.to_excel(geraNomePlanilha, index=False)
sleep(1)

print('')
print("Foram encontrados " + str(len(arquivos)) + " SKUs no Excel. Dessa lista, " + str(len(produtos)) + " estavam no site")
print('')
print('Foi gerado o arquivo ' + geraNomePlanilha + ' em sua pasta')

navegador.quit()
print('Fim')

os.system("pause")

