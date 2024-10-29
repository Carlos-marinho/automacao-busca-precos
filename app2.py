from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import tkinter as tk
from tkinter import filedialog, scrolledtext
import openpyxl
import time
import pandas as pd
from datetime import datetime
import re

# Function to initialize the WebDriver
def initialize_driver(max_retries=3, timeout=60):
    options = Options()
    arguments = [
            '--block-new-web-contents',
            '--disable-notifications',
            '--no-default-browser-check',
            '--lang=pt-BR',
            '--window-position=36,68',
            '--window-size=1920,1080',
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-software-rasterizer',
            '--disable-extensions',
            '--disable-popup-blocking',
            '--disable-gpu',
            '--ignore-certificate-errors',
            '--blink-settings=imagesEnabled=false']

    for argument in arguments:
        options.add_argument(argument)

    #service = Service(executable_path="/path/to/chromedriver")  # Update this path to your chromedriver
    #driver = webdriver.Chrome(service=service, options=options)
    retries = 0

    while retries < max_retries:
        try:        
            driver = webdriver.Chrome(options=options)
            
            driver.set_page_load_timeout(timeout)

            print(f"Driver iniciado com sucesso na tentativa {retries + 1}")
            
            return driver
        except Exception as e:
            retries += 1 
            print(f"Falha ao iniciar o driver na tentativa {retries}: {str(e)}")

            # Se falhar o número máximo de vezes, feche o driver e lance um erro
            if retries == max_retries:
                if driver:
                    driver.quit()
                raise Exception("Falha ao iniciar o driver após várias tentativas")

            # Se falhar, aguarde alguns segundos antes de tentar novamente
            time.sleep(3)

# Function to build the URL with the product name
def build_search_url(site_name, product_name):
    site_info = SITE_INFO[site_name]
    # Replace spaces in product name with '%20'
    formatted_product_name = product_name.replace(' ', '%20')
    return site_info["url"].format(productName=formatted_product_name)

# Função para converter metros para centímetros, se aplicável
def convert_units(value, from_unit, to_unit='cm'):

    conversion_factors = {
        'm': 100,    # 1 metro = 100 cm
        'cm': 1,     # 1 centímetro = 1 cm
        'mm': 0.1    # 1 milímetro = 0.1 cm
    }

    return value * conversion_factors[from_unit]

# Função para extrair uma ou mais dimensões da descrição do produto
def extract_dimensions(text):
    text = text.lower().strip()

    dimension_patterns = [
        r'(\d+[\.,]?\d*)\s*(m|cm|mm|metros?|mts?)',  # Formatos de 1 dimensão com unidade (ex: "100m", "2,5mm", "100 metros")
        r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(m|cm|mm|metros?|mts?)',  # Formatos de 2 dimensões (ex: "30x40cm", "30 x 40 cm", "42,5x31cm")
        r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(m|cm|mm|metros?|mts?)'  # Formatos de 3 dimensões (ex: "30x40x50cm")
    ]

    for pattern in dimension_patterns:
        match = re.search(pattern, text)
        if match:
            # Extrair números e unidade correspondente
            numbers = match.groups()[:-1]  # Todas as dimensões numéricas
            unit = match.groups()[-1]  # Unidade da dimensão (m, cm, mm)

            # Normalizar números (substituir vírgula por ponto)
            normalized_numbers = [num.replace(',', '.') for num in numbers]

            # Unidades possíveis de conversão
            unit = unit.replace('metros', 'm').replace('mts', 'm')

            # Converter os valores para centímetros
            converted_values = [convert_units(float(num), unit, 'cm') for num in normalized_numbers]

            # Retornar as dimensões em centímetros
            return converted_values

    return None  # Caso não encontre dimensões na descrição

# Function to check if the product description matches
def is_product_match(driver, site_name, product_name):
    try:
        site_info = SITE_INFO[site_name]

        card_elements = driver.find_elements(By.XPATH, site_info["cards_xpath"])
        if len(card_elements) > 3:
            card_elements = card_elements[:3]
        
        matched_elements = []

        product_keywords_list = product_name.lower().split(" ")
        
        #Peso para palavra chave
        weights = [3, 2, 2] + [1] * (len(product_keywords_list) - 3)

        product_keyword_score = sum(weights[i] for i,keyword in enumerate(product_keywords_list)) 

        print(f'\nIniciando busca no site {site_name} para o produto {product_name}...\nScore: {product_keyword_score}')
        
        # Extrair possíveis dimensões da descrição do produto
        expected_dimensions = extract_dimensions(product_name.lower())

        # Iterate over each card element and check if the description matches
        for card in card_elements:
            try:
                # Find the description element within the card
                #Exception for leroymerlin because its cards classes' inconsistency 
                if site_name == 'leroymerlin':
                    description_element = card.get_attribute(site_info["description_xpath"])
                    description_text = description_element.strip().lower()
                    price = card.get_attribute(site_info["price_xpath"])
                else:
                    description_element = card.find_element(By.XPATH, site_info["description_xpath"])
                    description_text = description_element.text.strip().lower()
                    price = card.find_element(By.XPATH, site_info["price_xpath"]).text

                score = 0
                keyword_count = 0

                for i, keyword in enumerate(product_keywords_list):
                    if keyword in description_text:
                        score += weights[i]
                        keyword_count += 1
                
                # Verificar correspondência de dimensões
                found_dimensions = extract_dimensions(description_text)
                dimension_match = False

                if expected_dimensions and found_dimensions:
                    # Caso 1: Se houver apenas uma dimensão a ser comparada (ex: "3m" vs "300cm")
                    if len(expected_dimensions) == 1 and len(found_dimensions) == 1:
                        # Verifica se a única dimensão está dentro da tolerância de 5 cm
                        dimension_match = abs(expected_dimensions[0] - found_dimensions[0]) == 0
                    # Caso 2: Se houver múltiplas dimensões (ex: "1,20x1,20m" vs "120x120cm")
                    elif len(expected_dimensions) == len(found_dimensions):
                        # Verifica se todas as dimensões estão dentro da tolerância de 5 cm
                        dimension_match = all(abs(exp - found) == 0 for exp, found in zip(expected_dimensions, found_dimensions))

                    if dimension_match:
                        score += 3  # Adiciona peso adicional para correspondência de dimensões


                if keyword_count / len(product_keywords_list) >= 0.7 and score / product_keyword_score > 0.85:
                    matched_elements.append((card, score))

                print(f'\nChecando {description_text} - {keyword_count} palavras encontradas - Pontuação {score}')
                print(f'Dimensões esperadas: {expected_dimensions} / Dimensões encontradas: {found_dimensions}')
                print(f'Preço R$ {price.strip(' R$').replace(',','.')}')
            
            except Exception as e:
                print(f'Erro ao iterar card - Produto/preço indisponível {e}')
                continue

        if not matched_elements:
            return False
        
        max_count = max([count for _, count in matched_elements])  # Pega a maior contagem de palavras
        filtered_matches = [elem for elem, count in matched_elements if count == max_count]

        # Return the filtered matched elements or an empty list if no matches
        return filtered_matches

    except Exception as e:
        print(f"\nErro ao verificar a descrição do produto {product_name} no site {site_name}")
        return False
    
# Function to search for the product and retrieve the price
def get_product_price(driver, site_name, product_name):
    attempt = 0
    retries = 3
    wait_time = 10 

    while attempt < retries:
        try:
            # Build the search URL
            search_url = build_search_url(site_name, product_name)
            driver.get(search_url)

            site_info = SITE_INFO[site_name]
            
            # Wait for the dynamic content to load
            WebDriverWait(driver, 15, poll_frequency=0.5).until(EC.element_to_be_clickable((By.XPATH, site_info["cards_xpath"])))
            time.sleep(4)
            
            # Get the matching product elements (cards) based on the description
            matched_elements = is_product_match(driver, site_name, product_name)

            # Verify if the product found matches the search
            if not matched_elements or len(matched_elements) == 0:
                return "N/A"  # If no match, return N/A

            lowest_price = float("inf")

            print(f'\nChecando preços...')

            for element in matched_elements:
                try:
                    # Locate the price using the provided XPath
                    if site_name == 'leroymerlin':
                        price_element = element.get_attribute(site_info["price_xpath"])
                    else:
                        price_element = element.find_element(By.XPATH, site_info["price_xpath"]).text

                    price = price_element.strip(' R$').replace(',','.')
                    
                    if price:
                        price_value = float(price)
                        # Update the lowest price
                        if price_value < lowest_price:
                            lowest_price = price_value
                except Exception as e:
                    continue  # Skip if price extraction fails
            
            print(f'\nLoja: {site_name} - Produto: {product_name} - Menor Preço: R${lowest_price}')

            return lowest_price if lowest_price != float('inf') else "N/A"
        
        except NoSuchElementException as e:
            # Tratar especificamente o erro de elemento não encontrado
            print(f"Elemento não encontrado no site {site_name} com o produto {product_name}: {e}")
            return "N/A"
        
        except Exception as e:
            print(f'Erro no site {site_name} produto {product_name} não encontrado {e}')
            return "N/A"  # Return N/A if product not found or error occurs

# Function to process the entire list of products
def process_products(product_list, site_info):
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

    driver.quit()

    return results

def load_product_list(file_path, sheet_name='Dados', header_row=1, columns=['Código', 'Descrição']):
    spreadsheet = pd.ExcelFile(file_path)
    product_list = pd.read_excel(spreadsheet, sheet_name=sheet_name, header=header_row, usecols=columns)
    product_list = product_list.dropna(subset=columns)
    return product_list

def prepare_historical_data(results_df, site_info):
    historical_data = pd.DataFrame()
    
    for loja in site_info.keys():
        if loja in results_df.columns:
            loja_df = results_df[['Código', 'Descrição', loja, 'Data']].copy()
            loja_df['Loja'] = loja.capitalize()
            loja_df.rename(columns={loja: 'Preço'}, inplace=True)
            historical_data = pd.concat([historical_data, loja_df])
        else:
            print(f'Aviso: Loja {loja} não encontrada')
    
    # Remove valores nulos (linhas sem preço)
    historical_data = historical_data.replace('N/A', float('nan')).dropna(subset=['Preço'])
    return historical_data

def save_to_excel(file_path, data, sheet_name, mode='a', if_sheet_exists='overlay', index=False):
    with pd.ExcelWriter(file_path, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=index)

def start_search(file_path, site_info):
    start_time = time.time()
    
    # Carregar a lista de produtos
    print("Carregando lista de produtos...")
    product_list = load_product_list(file_path)

    print("Iniciando a busca de preços...")
    price_results = process_products(product_list, site_info)
    print(price_results)

    # Processar os produtos para obter os preços atuais
    results_df = pd.DataFrame(price_results)

    # Adicionar a data atual aos resultados
    current_date = datetime.now().strftime('%d/%m/%y')
    results_df['Data'] = current_date

    # Preparar os dados históricos
    historical_data = prepare_historical_data(results_df, site_info)

    # Salvar os dados históricos e os preços atuais de forma organizada
    print("Salvando resultados...")
    save_to_excel(file_path, results_df, sheet_name='Pesquisa de Preços', mode='a', if_sheet_exists='replace')
    save_to_excel(file_path, historical_data, sheet_name='Histórico', mode='a', if_sheet_exists='overlay')

    end_time = time.time()

    execution_time = end_time - start_time
    print(f"Tempo total de execução: {execution_time:.2f} segundos")

    print("Processo concluído com sucesso!")

# Função para abrir o seletor de arquivos e definir o caminho
def browse_file_path(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

# Função principal que cria a interface gráfica
def create_gui():
    # Configurações principais da janela
    root = tk.Tk()
    root.title("Automação de Busca de Preços")
    root.geometry("600x400")

    # Layout
    tk.Label(root, text="Selecione o arquivo Excel:").pack(pady=10)

    file_entry = tk.Entry(root, width=50)
    file_entry.insert(0, './TabelaTeste.xlsx')  # Caminho padrão
    file_entry.pack(pady=5)

    browse_button = tk.Button(root, text="Procurar", command=lambda: browse_file_path(file_entry))
    browse_button.pack(pady=5)

    start_button = tk.Button(root, text="Iniciar Busca", command=lambda: start_search(file_entry.get(), SITE_INFO, log_text))
    start_button.pack(pady=20)

    log_text = scrolledtext.ScrolledText(root, width=70, height=15)
    log_text.pack(pady=10)

    # Loop principal da interface gráfica
    root.mainloop()

def main(file_path, site_info):
    #create_gui()
    start_search(file_path, site_info)

if __name__ == "__main__":
    # Definir o caminho do arquivo e a estrutura de sites
    file_path = './Tabela.xlsx'
    # Base URLs and XPath mappings for each site
    SITE_INFO = {
        "leroymerlin": {
            "store_name": "Leroy Merlin",
            "url": "https://www.leroymerlin.com.br/search?term={productName}&searchTerm={productName}&searchType=default",
            "price_xpath": "data-gtm-item-price",
            "description_xpath": "data-gtm-item-name",
            "cards_xpath": "//div[contains(@class,'new-product-thumb') and contains(@class,'css-1snnu76-wrapper-wrapper--full-width')]"
        },
        "chatuba": {
            "store_name": "Chatuba",
            "url": "https://www.chatuba.com.br/{productName}?_q={productName}&map=ft",
            "price_xpath": ".//span[@class='vtex-product-price-1-x-currencyContainer vtex-product-price-1-x-currencyContainer--summary-shelf']",
            "description_xpath": ".//span[@class='vtex-product-summary-2-x-productBrand vtex-product-summary-2-x-productBrand--summary-shelf vtex-product-summary-2-x-brandName vtex-product-summary-2-x-brandName--summary-shelf t-body']",
            "cards_xpath": "//*[@id='gallery-layout-container']/div/section"
        },
        "obramax": {
            "store_name": "Obramax",
            "url": "https://www.obramax.com.br/{productName}?_q={productName}&map=ft",
            "price_xpath": ".//a/article/div[4]/div/div[1]/div/div[3]/div/div/div[1]/div/span[2]",
            "description_xpath": ".//a/article/div[3]/h3/span",
            "cards_xpath": "//*[@id='gallery-layout-container']/div/section"
        },
        "amoedo": {
            "store_name": "Amoedo",
            "url": "https://www.amoedo.com.br/{productName}",
            "price_xpath": ".//span[@class='vtex-product-price-1-x-sellingPriceValue vtex-product-price-1-x-sellingPriceValue--summary']/span[1]",
            "description_xpath": ".//span[@class='vtex-product-summary-2-x-productBrand vtex-product-summary-2-x-brandName t-body']",
            "cards_xpath": "//*[@id='gallery-layout-container']/div/section"
        },
        "sepa": {
            "store_name": "Sepa",
            "url": "https://www.sepaconstruirdecorar.com.br/{productName}",
            "price_xpath": ".//span[@class='valor-por']/strong",
            "description_xpath": ".//div[@class='product-name']",
            "cards_xpath": "//*[@class='prateleira']"
        }
    }

    # Executar a função principal
    main(file_path, SITE_INFO)



#CÓDIGO PARA TESTAR PRODUTOS SEM PLANILHA
#price_results = process_products(['porcelanato calacata oro lux 90x90 biancogres', 'cemento grigio 90x90 biancogres', 'porcelanato stratus grigio satin 90x90 biancogres'])
#print(price_results)    