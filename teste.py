import re

# Função para converter unidades para centímetros, se aplicável
def convert_units(value, from_unit):
    conversion_factors = {
        'm': 100,    # 1 metro = 100 cm
        'cm': 1,     # 1 centímetro = 1 cm
        'mm': 0.1    # 1 milímetro = 0.1 cm
    }
    return value * conversion_factors.get(from_unit, 1)

# Função para extrair dimensões da descrição do produto
def extract_dimensions(text):
    text = text.lower().strip()
    all_dimensions = []

    # Padrões para capturar dimensões
    patterns = [
        (r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?)?', 3),  # 3 dimensões com ou sem unidade
        (r'(\d+[\.,]?\d*)\s*x\s*(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?)?', 2),  # 2 dimensões com ou sem unidade
        (r'(\d+[\.,]?\d*)\s*(mm|cm|m|metros?|mts?)', 1)  # 1 dimensão com unidade
    ]

    for pattern, num_dimensions in patterns:
        while True:
            match = re.search(pattern, text)
            print(f'Match: {match}')
            if not match:
                break  # Se não houver mais correspondências, saia do loop para este padrão

            # Extrair os valores e unidade de medida
            dimensions = [float(match.group(i).replace(',', '.')) for i in range(1, num_dimensions + 1)]
            unit = match.group(num_dimensions + 1) if match.group(num_dimensions + 1) else 'cm'
            unit = unit.replace('metros', 'm').replace('mts', 'm')

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

def adjust_product_name(product_name):
    product_name = product_name.lower().split(" ")
    
    if len(product_name) > 5:
        new_product_name = product_name[:3] + [product_name[-1]]
    else:
        new_product_name = product_name

    product_name = " ".join(new_product_name)

    return product_name

# Testando a função com os produtos fornecidos
products = [
    "MANTA ASFÁLTICA VIAMANTA MULTIUSO 3MM VIAPOL",
    "FIO CABO FLEXIVEL 1,5MM BRANCO ROLO 100M CORFIO",
    "ivory retificado brilhante bege 61x61cm 1,87m",
    "CAIXA GORDURA 40 X 41 X 41 PRETA STAND C/CESTO METASUL",
]

results = [{'Descrição': 'porcelanato delta avorio 62x62', 'leroymerlin': 'N/A', 'chatuba': 58.5, 'obramax': 'N/A', 'amoedo': 69.9, 'sepa': 'N/A'}]


# for product in products:
#     #print(f'Dimensões para "{product}": {extract_dimensions(product)}\n')
#     print(f'Nome do produto ajustado: {adjust_product_name(product)}')


for result in results:
    print(result)
    print(f'\nDescrição: {result["Descrição"].capitalize()}\n')
    print(f'Leroy Merlin: {result["leroymerlin"]}')
    print(f'Chatuba: {result["chatuba"]}')
    print(f'Obramax: {result["obramax"]}')
    print(f'Amoedo: {result["amoedo"]}')
    print(f'Sepa: {result["sepa"]}\n')