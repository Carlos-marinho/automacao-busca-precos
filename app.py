import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from time import sleep

#PESQUISA DE PREÇOS AUTOMATIZADA
"""
SITES A SEREM VERIFICADOS:
-LEROY OK
-CHATUBA 
-AMOEDO
-OBRAMAX
-SEPA

FALTA IMPLEMENTAR:
-ABRIR PLANILHA COM OS NOMES DOS PRODUTOS (UTILIZAR A QUE O JONAS MANDOU)
-FAZER A PESQUISA COM TODOS OS NOMES LISTADOS NA PLANILHA E SALVÁ-LOS
"""


def processURLSearch(productNameList, store):
    
    if store.lower() == "leroy":        
        URL = f"https://www.leroymerlin.com.br/"
        productName = productNameList.replace(" ","%20")
        search = f"search?term={productName}&searchTerm={productName}&searchType=default"
        URL += search

    return URL

def setDriverSettings():
    try:
        options = Options()
        arguments = [
            '--block-new-web-contents',
            '--disable-notifications',
            '--no-default-browser-check',
            '--lang=pt-BR',
            '--window-position=36,68',
            '--window-size=1100,750',]

        for argument in arguments:
            options.add_argument(argument)

       # options.add_experimental_option("excludeSwitches", ["enable-logging"])

        driver = webdriver.Chrome(options=options)
       

        return driver

    except Exception as e:
        #logging.error(f'Erro na configuração do driver: {e}')
        print(f"Error: {e}")
        #return None

def getUrlData(URL, driver):
    try:
        driver.get(URL)
        WebDriverWait(driver, 30, poll_frequency=1).until(EC.element_to_be_clickable((By.CLASS_NAME, "new-product-thumb")))
        sleep(5)
        divMain = driver.find_elements(By.CLASS_NAME, "new-product-thumb")
    
    except Exception as e:
        print(f"Error: {e}")
        driver.quit()

    else:
        return divMain

def getProductDetails(elements, productName, productList):
    

    if len(elements) > 1:
        for element in elements:
            itemName = element.get_attribute("data-gtm-item-name")
            itemPrice = element.get_attribute("data-gtm-item-price")

            productNameList = productName.lower().split()
            itemNameList = itemName.lower().split()

            if all(word in itemNameList for word in productNameList[:2]): 
                print("Produto Encontrado")
                print(itemName)
                print(itemPrice)
                productList.append({"name": itemName, "price": itemPrice})
    else:
        itemName = elements[0].get_attribute("data-gtm-item-name")
        itemPrice = elements[0].get_attribute("data-gtm-item-price")

        productList.append({"name": itemName, "price": itemPrice})
    
    return productList
        

def main():
    driver = setDriverSettings()
    productList = []

    productName = input("Digite o nome do produto: ")
    #store = input("Digite a loja a ser pesquisada (leroy, chatuba): ")
    store = "leroy"

    processedURLSearch = processURLSearch(productName, store)

    dataList = getUrlData(processedURLSearch, driver)
    productList = getProductDetails(dataList, productName, productList)
        
        
    driver.quit()
    print(productList)

if __name__ == "__main__":
    main()