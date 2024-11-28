from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from threading import Thread, Event
from openpyxl.utils import get_column_letter
import time
import pandas as pd
from datetime import datetime
import sys
import re
import os
from unidecode import unidecode


# Function to initialize the WebDriver
def initialize_driver(max_retries=3, timeout=30):
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
            '--ignore-certificate-errors',
            '--blink-settings=imagesEnabled=false']

    for argument in arguments:
        options.add_argument(argument)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)

    retries = 0

    while retries < max_retries:
        try:
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
def convert_units(value, from_unit):
    conversion_factors = {
        'm': 100,    # 1 metro = 100 cm
        'cm': 1,     # 1 centímetro = 1 cm
        'mm': 0.1,    # 1 milímetro = 0.1 cm
        'kg': 1,
        'l': 1
    }
    factor = conversion_factors.get(from_unit, 1)
    return round(value * factor,2)

# Função para extrair dimensões da descrição do produto
def extract_dimensions(text):
    text = text.lower().strip()

    # Substitui padrões como "1 5mm" por "1.5mm" ou "1,5mm"
    text = re.sub(r'(\d+)\s+(\d+)(mm|cm|m)', r'\1.\2\3', text)

    all_dimensions = []

    # Padrões para capturar dimensões de 3, 2 e 1 valores com e sem unidade
    patterns = [
        (r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?|kg|lts?|l)?', 3),  # 3 dimensões com ou sem unidade
        (r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?|kg|lts?|l)?', 2),  # 2 dimensões com ou sem unidade
        (r'(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?|kg|lts?|l)', 1)  # 1 dimensão com unidade
    ]

    # Processar cada padrão de forma sequencial
    for pattern, num_dimensions in patterns:
        while True:
            match = re.search(pattern, text)
            if not match:
                break  # Se não houver mais correspondências, saia do loop para este padrão

            # Extrair os valores e unidade de medida
            dimensions = [float(match.group(i).replace(',', '.')) for i in range(1, num_dimensions + 1)]
            unit = match.group(num_dimensions + 1) if match.group(num_dimensions + 1) else 'cm'
            unit = unit.replace('metros', 'm').replace('mts', 'm').replace('lts', 'l')
            
            # Converter para centímetros
            converted_dimensions = [convert_units(value, unit) for value in dimensions]

            # Evitar duplicações e garantir estrutura correta
            if len(converted_dimensions) == 1:
                if converted_dimensions[0] not in all_dimensions:
                    all_dimensions.append(converted_dimensions[0])
            else:
                if converted_dimensions not in all_dimensions:
                    all_dimensions.append(converted_dimensions)

            # Remover a parte do texto que foi capturada
            text = text[:match.start()] + text[match.end():]

    return all_dimensions if all_dimensions else None

# Function to check if the product description matches
def is_product_match(driver, site_name, product_name):
    try:
        site_info = SITE_INFO[site_name]

        card_elements = driver.find_elements(By.XPATH, site_info["cards_xpath"])
        if len(card_elements) > 8:
            card_elements = card_elements[:8]
        
        matched_elements = []

        product_keywords_list = product_name.lower().split(" ")
        
        #Peso para palavra chave
        weights = [2, 2, 2] + [1] * (len(product_keywords_list) - 4) + [3]

        product_keyword_score = sum(weights[i] for i,keyword in enumerate(product_keywords_list)) 
        
        print('')
        print('-'*60)
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
                    price = card.find_elements(By.XPATH, site_info["price_xpath"])

                    if price:
                        if len(price) > 1:
                            price = price[-1].text
                        else:
                            price = price[0].text

                score = 0
                keyword_count = 0

                for i, keyword in enumerate(product_keywords_list):
                    keyword = normalize_words(keyword)
                    #print(keyword)
                    if keyword in unidecode(description_text).replace(" ", ""):
                        #print(f'Checando na {unidecode(description_text).replace(" ", "")}')
                        score += weights[i]
                        keyword_count += 1
                
                # Verificar correspondência de dimensões
                found_dimensions = extract_dimensions(description_text)
                dimension_match = False

                if expected_dimensions and found_dimensions:
            
                    # Garante que ambas são listas, ou transforma valores únicos em listas para comparação
                    expected_dimensions = [[dim] if isinstance(dim, (int, float)) else dim for dim in expected_dimensions]
                    found_dimensions = [[dim] if isinstance(dim, (int, float)) else dim for dim in found_dimensions]

                    # Caso 1: Se ambas são listas de múltiplas dimensões (ex.: [[90, 180]] vs. [[90, 90, 180]])
                    if isinstance(expected_dimensions[0], list) and isinstance(found_dimensions[0], list):
                        # Ordenar sublistas com exatamente duas dimensões para ignorar a ordem
                        sorted_expected = [sorted(sublist) if len(sublist) == 2 else sublist for sublist in expected_dimensions]
                        sorted_found = [sorted(sublist) if len(sublist) == 2 else sublist for sublist in found_dimensions]

                        flattened_expected = [dim for sublist in sorted_expected for dim in sublist]
                        flattened_found = [dim for sublist in sorted_found for dim in sublist]
            
                        dimension_match = all(any(abs(exp - found) <= 3 for found in flattened_found) for exp in flattened_expected)

                    # Caso 2: Comparação direta de valores únicos (já transformados em listas)
                    elif len(expected_dimensions) == 1 and len(found_dimensions) == 1:
                        dimension_match = abs(expected_dimensions[0][0] - found_dimensions[0][0]) <= 3

                    #else:
                        # Lidar com casos mistos (uma lista e outra não, que pode ocorrer em dados inconsistentes)
                        #dimension_match = False


                    if dimension_match:
                        score += 5  # Aumenta o peso para correspondência de dimensões
                    else:
                        score -= 2  # Penaliza fortemente se as dimensões não corresponderem

                if len(product_keywords_list) > 3:
                    if keyword_count / len(product_keywords_list) >= 0.75 and score / product_keyword_score >= 0.75:
                        matched_elements.append((card, score))
                else:
                    if keyword_count / len(product_keywords_list) >= 0.65 and score / product_keyword_score >= 0.65:
                        matched_elements.append((card, score))

                print(f'\nEncontrado {description_text} - {keyword_count} palavras correspondentes - Pontuação {score}')
                print(f'Dimensões esperadas: {expected_dimensions}\nDimensões encontradas: {found_dimensions}')
                print(f'Preço R$ {price.strip(" R$ un m²").replace(",",".")}')
                
            
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

#Function to adjust product name if it's longer
def adjust_product_name(product_name):
    product_name = product_name.lower().split(" ")
    
    if len(product_name) > 5:
        new_product_name = product_name[:4] + [product_name[-1]]
    else:
        new_product_name = product_name

    product_name = " ".join(new_product_name)

    return product_name

# Function to search for the product and retrieve the price
def get_product_price(driver, site_name, product_name):
    attempt = 0
    retries = 2
    wait_time = 5 

    while attempt < retries:
        try:
            # Build the search URL
            if attempt > 0:
                new_product_name = adjust_product_name(product_name)
                search_url = build_search_url(site_name, new_product_name)
                print("Nome ajustado: ",new_product_name)
            else:
                search_url = build_search_url(site_name, product_name)
            
            driver.get(search_url)

            site_info = SITE_INFO[site_name]
            
            # Wait for the dynamic content to load
            #print('\nEsperando carregar conteudo do site...')
            WebDriverWait(driver, 30, poll_frequency=0.5).until(EC.presence_of_element_located((By.XPATH, site_info["cards_xpath"])))
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
                        price_element = element.find_elements(By.XPATH, site_info["price_xpath"])

                        if len(price_element) > 1:
                            price_element = price_element[-1].text
                        else:
                            price_element = price_element[0].text

                    price = price_element.strip(' R$ un m²').replace(',','.')
                    
                    if price:
                        price_value = float(price)
                        # Update the lowest price
                        if price_value < lowest_price:
                            lowest_price = price_value
                except Exception as e:
                    continue  # Skip if price extraction fails
            
            print(f'\nLoja: {site_name}\nProduto: {product_name}\nMenor Preço encontrado: R${lowest_price}')

            return lowest_price if lowest_price != float('inf') and lowest_price != 0 else "N/A"
        
        except NoSuchElementException as e:
            # Tratar especificamente o erro de elemento não encontrado
            print(f"Elemento não encontrado no site {site_name} com o produto {product_name}")

        except Exception as e:
            print(f'\nErro no site {site_name} produto {product_name} não encontrado')
        
        # Incrementa a tentativa e continua no loop
        attempt += 1
        if attempt < retries and len(product_name.split(" ")) > 5:
            print(f"Tentando buscar novamente...{attempt}\n")
            time.sleep(wait_time)  # Aguarde alguns segundos antes de tentar novamente
            continue
        else:
            return "N/A"  # Retorna "N/A" se todas as tentativas falharem

#Function to normalize words with gender
def normalize_words(word):
    if word[-1] == 'a' or word[-1] == 'o':
        return word[:-1]
    return word

# Function to process the entire list of products
def process_products(product_list, site_info, driver, stop_event, search_opt):
    results = []

    if search_opt == 'planilha':
        for _, row in product_list.iterrows():
            # Verifica se o evento de interrupção foi acionado
            if stop_event.is_set():
                print("Busca interrompida pelo usuário. Salvando dados até o momento e encerrando...")
                break

            product_code = row['Código']  # Adjusted to the column names in the new sheet
            product_name = row['Descrição']
            
            product_prices = {"Código": product_code, "Descrição": product_name}
            
            for site_name in site_info.keys():
                price = get_product_price(driver, site_name, product_name)
                product_prices[site_name] = price
            
            results.append(product_prices)

    elif search_opt == 'descricao':
        for product in product_list:
            # Verifica se o evento de interrupção foi acionado
            if stop_event.is_set():
                print("Busca interrompida pelo usuário. Salvando dados até o momento e encerrando...")
                break
            
            product_name = product.strip()

            if product_name == '':
                continue

            product_prices = {"Descrição": product_name}

            for site_name in site_info.keys():
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
    pd.set_option('future.no_silent_downcasting', True)
    historical_data = historical_data.replace('N/A', float('nan')).dropna(subset=['Preço']).infer_objects(copy=False)
    
    return historical_data

def save_to_excel(file_path, data, sheet_name, mode='a', if_sheet_exists='overlay', index=False):
    # Extrair a data dos novos dados
    if 'Data' in data.columns:
        new_data_date = data['Data'].iloc[0]
    else:
        new_data_date = pd.to_datetime('today').strftime('%d/%m/%y')
    
    try:
        # Carregar dados existentes na aba "Histórico"
        existing_data = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Filtrar os dados existentes para remover linhas com a mesma data que os novos dados
        existing_data = existing_data[existing_data['Data'] != new_data_date]

        # Concatenar os dados novos com os dados existentes filtrados
        combined_data = pd.concat([existing_data, data], ignore_index=True)
    except (FileNotFoundError, ValueError):
        # Se o arquivo ou a aba não existir, usa apenas os novos dados
        combined_data = data
    
    with pd.ExcelWriter(file_path, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=index)

        # Acessa a planilha após salvar para ajustar a largura da coluna "Descrição"
        worksheet = writer.sheets[sheet_name]
        
        # Encontra o comprimento máximo de todos os dados na coluna "Descrição"
        max_length = max([len(str(cell)) for cell in data['Descrição']])

        # Define o ajuste de largura com um pequeno espaço adicional
        col_letter = get_column_letter(data.columns.get_loc("Descrição") + 1)  # Obtém a letra da coluna "Descrição"
        worksheet.column_dimensions[col_letter].width = max_length + 2  # Ajusta a largura da coluna com espaço extra

def start_search(file_path, site_info, driver, stop_event, search_opt):
    start_time = time.time()
    
    # Carregar a lista de produtos
    print("Carregando lista de produtos...")

    if search_opt == 'planilha':
        product_list = load_product_list(file_path)
    
        print("Iniciando a busca de preços...")
        price_results = process_products(product_list, site_info, driver, stop_event, search_opt)
        # Processar os produtos para obter os preços atuais
        results_df = pd.DataFrame(price_results)

        # Adicionar a data atual aos resultados
        current_date = datetime.now().strftime('%d/%m/%y')
        results_df['Data'] = current_date

        # Preparar os dados históricos
        historical_data = prepare_historical_data(results_df, site_info)

        # Salvar os dados históricos e os preços atuais de forma organizada
        print("\nSalvando resultados...")
        save_to_excel(file_path, results_df, sheet_name='Pesquisa de Preços', mode='a', if_sheet_exists='replace')
        save_to_excel(file_path, historical_data, sheet_name='Histórico', mode='a', if_sheet_exists='overlay')
    
    elif search_opt == 'descricao':
        product_list = file_path
        print("Iniciando a busca de preços...") 
        price_results = process_products(product_list, site_info, driver, stop_event, search_opt)
        print("\n\nResultados encontrados:")
        for result in price_results:
            print(f'\nDescrição: {result["Descrição"].capitalize()}\n')
            print(f'Leroy Merlin: R$ {result["leroymerlin"]}')
            print(f'Chatuba: R$ {result["chatuba"]}')
            print(f'Obramax: R$ {result["obramax"]}')
            print(f'Amoedo: R$ {result["amoedo"]}')
            print(f'Sepa: R$ {result["sepa"]}\n')

    

    end_time = time.time()

    execution_time = end_time - start_time
    print(f"\nTempo total de execução: {execution_time:.2f} segundos")

    print("\nProcesso concluído com sucesso!")

def start_gui():
    stop_event = Event()  # Variável de controle para interromper a busca
    driver = None  # Variável global para o driver do Selenium

    # Função para selecionar o caminho do arquivo
    def select_file():
        file_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))
    
    # Função para atualizar o log na interface gráfica
    def log_message(message):
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)  # Scroll para o final do log
        root.update_idletasks()  # Atualiza a interface imediatamente

        # Escreve no arquivo de log
        write_log_to_file(message)
    
    # Função que executa a busca em uma nova thread
    def start_search_thread():
        global driver

        def run_search(path):
            global driver
            stop_event.clear()  # Reseta o evento de parada

            # Limpa o log antes de iniciar a nova busca
            log_text.delete(1.0, tk.END)    

            # Registra a data e hora da busca no log
            now = datetime.now().strftime("%d/%m/%Y %H:%M")
            log_message(f"Inicio da Busca: {now}")

            
            # Alterna visibilidade dos botões
            start_button.pack_forget()  # Oculta o botão "Iniciar Busca"
            file_name_input.config(state=tk.DISABLED)
            select_file_button.config(state=tk.DISABLED)
            description_input.config(state=tk.DISABLED)
            
            stop_button.config(state=tk.NORMAL)  # Reativa o botão "Interromper Busca"
            stop_button.pack(pady=10)  # Mostra o botão "Interromper Busca"
            
            try:
                # Redireciona os prints para o log da interface
                global print
                original_print = print
                print = lambda *args: log_message(" ".join(map(str, args)))

                # Inicia o driver e executa a função principal
                driver = initialize_driver()  # Inicia o driver separado
                main(file_path=path, site_info=SITE_INFO, stop_event=stop_event, driver=driver, search_opt=search_mode)  # Executa o código principal sem `stop_event`

                # Log de conclusão caso a busca finalize sem interrupção
                log_message("Busca concluída com sucesso!")
                #messagebox.showinfo("Concluído", "Processo de busca finalizado!")

            except Exception as e:
                log_message(f"Erro durante a execução: {e}")
                messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
            finally:
                # Restaura a função print, alterna visibilidade dos botões e encerra o driver se ativo
                print = original_print
                
                start_button.pack(pady=10)  # Reexibe o botão "Iniciar Busca"
                file_name_input.config(state=tk.NORMAL)
                select_file_button.config(state=tk.NORMAL)  # Reativa o botão "Selecionar Arquivo"
                description_input.config(state=tk.NORMAL)  # Reativa o campo de entrada de descrição
                
                stop_button.pack_forget()  # Oculta o botão "Interromper Busca"

                if driver:
                    driver.quit()  # Fecha o driver ao final ou em caso de interrupção
                    driver = None  # Reseta a variável driver para None

        # Escolhe entre buscar pela planilha ou pela descrição
        search_mode = search_mode_var.get()

        if search_mode == 'planilha':
            path = file_name_input.get()
        
            if path:
                # Inicia a thread de busca
                search_thread = Thread(target=run_search, args=(path,))
                search_thread.start()
            else:
                messagebox.showwarning("Aviso", "Por favor, selecione um arquivo Excel antes de iniciar.")
                start_button.config(state=tk.NORMAL)
        elif search_mode == 'descricao':
            descriptions = description_input.get()

            if ";" in descriptions:
                descriptions = descriptions.strip().split(';')
            else:
                descriptions = [descriptions.strip()]  # Se não houver ";", trata a entrada como um único item

            if descriptions:
                # Inicia a thread de busca
                search_thread = Thread(target=run_search, args=(descriptions,))
                search_thread.start()
            else:
                messagebox.showwarning("Aviso", "Por favor, insira ao menos uma descrição de produto antes de iniciar.")
                start_button.config(state=tk.NORMAL)
         
    # Função para interromper a busca
    def stop_search():
        global driver
        stop_event.set()  # Sinaliza para interromper a busca
        log_message("Interrupção solicitada. Salvando dados até o momento e encerrando...")
        stop_button.config(state=tk.DISABLED)  # Desativa o botão "Interromper Busca"
        try:
            if driver is not None:  # Verifica se o driver foi inicializado
                driver.quit()  # Fecha o driver imediatamente quando a interrupção é solicitada
                driver = None  # Reseta a variável driver para None
        except Exception as e:
            log_message(f"Erro ao encerrar o driver: {e}")

    # Configuração principal da janela
    root = tk.Tk()
    root.title("Automação Pesquisa de Preços")

    # Defina as dimensões da janela
    window_width = 700
    #window_height = 500
    window_height = 800

    # Calcula a posição para centralizar a janela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = (screen_width - window_width) // 2
    position_y = (screen_height - window_height) // 2
    
    # Define a posição centralizada da janela
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    # Variável para o caminho do arquivo
    file_path = tk.StringVar(value="./TabelaTeste.xlsx")

    # Variável de modo de busca
    search_mode_var = tk.StringVar(value="planilha")

     # Layout da interface
    tk.Label(root, text="Selecione o Modo de Busca:").pack(pady=10)
    
    # Frame para organizar as opções de busca na mesma linha
    mode_frame = tk.Frame(root)
    mode_frame.pack(pady=5)
    
    tk.Radiobutton(mode_frame, text="Buscar pela Planilha", variable=search_mode_var, value="planilha").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(mode_frame, text="Buscar pela Descrição", variable=search_mode_var, value="descricao").pack(side=tk.LEFT, padx=5)

    # Frame para organizar o label e o botão "Selecionar Arquivo" na mesma linha
    file_frame = tk.Frame(root)
    file_frame.pack(pady=10)
    tk.Label(file_frame, text="Arquivo Excel:").pack(side=tk.LEFT, padx=(0, 10))
    
    file_name_input = tk.Entry(file_frame, textvariable=file_path, width=30)
    file_name_input.pack(side=tk.LEFT)

    select_file_button = tk.Button(file_frame, text="Selecionar Arquivo", command=select_file)
    select_file_button.pack(side=tk.LEFT)

    tk.Label(root, text="Descrição do Produto(s) (separado por ';' para vários):").pack(pady=10)
    description_input = tk.Entry(root, width=50)
    description_input.pack(padx=10, pady=5)
    
    # Log de execução em uma área de texto rolável
    #log_text = ScrolledText(root, width=85, height=15)
    log_text = ScrolledText(root, width=85, height=40)
    log_text.pack(pady=10)
    log_text.insert(tk.END, "Log de execução:\n")

    # Botão para iniciar a busca
    start_button = tk.Button(root, text="Iniciar Busca", command=start_search_thread)
    start_button.pack(pady=10)

    # Botão para interromper a busca (inicialmente oculto)
    stop_button = tk.Button(root, text="Interromper Busca", command=stop_search)
    stop_button.pack_forget()  # Oculta o botão inicialmente
    
    root.mainloop()

def write_log_to_file(message, log_file="search_logs.txt"):
    # Lê as linhas atuais do arquivo de log, se ele já existir
    if os.path.exists(log_file):
        with open(log_file, 'r', encoding='utf-8') as file:
            logs = file.readlines()
    else:
        logs = []

    # Adiciona a nova mensagem de log
    logs.append(message + "\n")

    # Mantém somente os logs das últimas três buscas
    if logs.count("Inicio da Busca:") > 3:
        # Encontra o índice da quarta busca mais antiga e remove as anteriores
        start_indices = [i for i, line in enumerate(logs) if line.startswith("Inicio da Busca")]
        logs = logs[start_indices[1]:]  # Mantém os últimos três logs de busca

    # Escreve de volta no arquivo
    with open(log_file, 'w', encoding='utf-8') as file:
        file.writelines(logs)

def main(file_path, site_info, stop_event, driver, search_opt):
    #driver = initialize_driver() 
    try:
        expiration_date = "2024-11-29"
        current_date = datetime.now().strftime("%Y-%m-%d")
        if current_date > expiration_date:
            print("O período de avaliação expirou. Entre em contato para obter a versão completa.")
            sys.exit()
        else:
            start_search(file_path, site_info, driver, stop_event, search_opt)
    finally:
        driver.quit()  # Garante que o driver seja fechado no final, independentemente de interrupções 

if __name__ == "__main__":
    # Definir o caminho do arquivo e a estrutura de sites
    file_path = './TabelaTeste.xlsx'
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
            "price_xpath": ".//p[@class='amoedo-tema-4-x-pixValue']",
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
    #main(file_path, SITE_INFO)
    start_gui()
