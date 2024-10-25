from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import openpyxl
import time
import pandas as pd
from datetime import datetime

#FALTA PARA FINALIZAR:
"""
-ARRUMAR CHECAGEM DE DESCRIÇÃO DOS PRODUTOS COM A DESCRIÇÃO DO PRODUTO NA PAGINA
-VERIFICAR ITENS PROMOCIONAIS SE NÃO OCORRE NENHUMA FALHA OU DIVERGENCIA NO PREÇO

"""


# Base URLs and XPath mappings for each site
SITE_INFO = {
    "leroymerlin": {
        "url": "https://www.leroymerlin.com.br/search?term={productName}&searchTerm={productName}&searchType=default",
        "price_xpath": "/html/body/div[7]/div[2]/div[2]/div/div/div[2]/div/div[4]/div/div/div/div[1]/div[2]/a/div[2]/div/span[1]",
        "description_xpath": "/html/body/div[7]/div[2]/div[2]/div/div/div[2]/div/div[4]/div/div/div/div[1]/div[1]/a/span"
    },
    "chatuba": {
        "url": "https://www.chatuba.com.br/{productName}?_q={productName}&map=ft",
        "price_xpath": "//*[@id='gallery-layout-container']/div/section/a/article/div[5]/div/div/div[1]/span[1]/span[1]/span",
        "description_xpath": "//*[@id='gallery-layout-container']/div/section/a/article/div[3]/h3/span"
    },
    "obramax": {
        "url": "https://www.obramax.com.br/{productName}?_q={productName}&map=ft",
        "price_xpath": "//*[@id='gallery-layout-container']/div[1]/section/a/article/div[4]/div/div[1]/div/div[3]/div/div/div[1]/div/span[2]",
        "description_xpath": "//*[@id='gallery-layout-container']/div[1]/section/a/article/div[3]/h3/span"
    },
    "amoedo": {
        "url": "https://www.amoedo.com.br/{productName}",
        "price_xpath": "//*[@id='gallery-layout-container']/div/section/a/article/div[4]/div/div/div/div[2]/span/span[1]/span",
        "description_xpath": "//*[@id='gallery-layout-container']/div[1]/section/a/article/div[4]/div/div[1]/div/h3/span"
    },
    "sepa": {
        "url": "https://www.sepaconstruirdecorar.com.br/{productName}",
        "price_xpath": "//*[@id='ResultItems_6358543']/div/ul/li/div/div[7]/div[1]/a/p/span[2]/strong",
        "description_xpath": "//*[@id='ResultItems_27269547']/div/ul/li[1]/div/div[7]/a/div"
    }
}

# Function to initialize the WebDriver
def initialize_driver():
    options = Options()
    arguments = [
            '--block-new-web-contents',
            '--disable-notifications',
            '--no-default-browser-check',
            '--lang=pt-BR',
            '--window-position=36,68',
            '--window-size=1100,750',
            '--headless']

    for argument in arguments:
        options.add_argument(argument)

    #service = Service(executable_path="/path/to/chromedriver")  # Update this path to your chromedriver
    #driver = webdriver.Chrome(service=service, options=options)

    driver = webdriver.Chrome(options=options)

    return driver

# Function to build the URL with the product name
def build_search_url(site_name, product_name):
    site_info = SITE_INFO[site_name]
    # Replace spaces in product name with '%20'
    formatted_product_name = product_name.replace(' ', '%20')
    return site_info["url"].format(productName=formatted_product_name)

# Function to check if the product description matches
def is_product_match(driver, site_name, product_name):
    try:
        site_info = SITE_INFO[site_name]
        description_element = driver.find_element(By.XPATH, site_info["description_xpath"])
        description_text = description_element.text.strip().lower()
        
        product_keywords_list = product_name.lower().split()
        
        checker = sum(1 for keyword in product_keywords_list if keyword in description_text)
        
        print(f'\nChecando se o produto {product_name} está na descrição {description_text}')
        print(f'Encontrados {checker} palavras na descrição no site {site_name}')
        # Check if product name is in the description text
        return checker >= 3
    except Exception as e:
        print(f"\nErro ao verificar a descrição do produto {product_name} no site {site_name}")
        return False
    
# Function to search for the product and retrieve the price
def get_product_price(driver, site_name, product_name):
    try:
        # Build the search URL
        search_url = build_search_url(site_name, product_name)
        driver.get(search_url)

        site_info = SITE_INFO[site_name]
        
        # Wait for the dynamic content to load
        WebDriverWait(driver, 30, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, site_info["price_xpath"])))
        time.sleep(5)
        
        # Verify if the product found matches the search
        if not is_product_match(driver, site_name, product_name):
            return "N/A"  # If no match, return N/A

        # Locate the price using the provided XPath
        price_element = driver.find_element(By.XPATH, site_info["price_xpath"])
        
        # Convert the price to a numerical format
        price = price_element.text.strip(' R$').replace(',','.')
        print(f'\nLoja: {site_name} - Produto: {product_name} - Preço: R${price}')
        
        return float(price) if price else "N/A"
    except Exception as e:
        print(f'Erro no site {site_name} produto {product_name} não encontrado')
        return "N/A"  # Return N/A if product not found or error occurs

# Function to process the entire list of products
def process_products(product_list):
    driver = initialize_driver()
    results = []

    for _, row in product_list.iterrows():
        product_code = row['Código']  # Adjusted to the column names in the new sheet
        product_name = row['Descrição']
        
        product_prices = {"Código": product_code, "Descrição": product_name}
        
        for site_name in SITE_INFO.keys():
            price = get_product_price(driver, site_name, product_name)
            product_prices[site_name] = price
        
        results.append(product_prices)


    """
    for product in product_list:
        
        product_prices = {"Descrição": product}
        
        for site_name in SITE_INFO.keys():
            price = get_product_price(driver, site_name, product)
            product_prices[site_name] = price
        
        results.append(product_prices)
    """
    driver.quit()
    return results


# Load the Excel file and prepare product list
file_path = './Tabela.xlsx'
spreadsheet = pd.ExcelFile(file_path)
product_list = pd.read_excel(spreadsheet, sheet_name='Dados', header=1 ,usecols=['Código','Descrição']).dropna(subset=['Código', 'Descrição'])


# Process the products to get current prices
price_results = process_products(product_list)

# Create DataFrame for the results
results_df = pd.DataFrame(price_results)

# Get the current date in DD/MM/YY format
current_date = datetime.now().strftime('%d/%m/%y')

# Add a column with the current date to the historical data
results_df['Data'] = current_date


# Save the results back to the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    # Save historical data with the current date
    results_df.to_excel(writer, sheet_name='Histórico', index=False)
    
    # Save current prices in a separate sheet
    current_prices_df = results_df.copy()
    #current_prices_df = results_df.copy().drop(columns=['Data'])  # Remove the 'Data' column for the current prices
    current_prices_df.to_excel(writer, sheet_name='Pesquisa de Preços', index=False)


#price_results = process_products(['calacata oro lux 90x90 biancogres'])
#print(price_results)
