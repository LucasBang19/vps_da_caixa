from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import os
import time
import xml.etree.ElementTree as ET
import re  # Importa o módulo para expressões regulares
from xml.dom import minidom  # Para formatar o XML
from datetime import datetime

def iniciar_driver():
    chrome_options = webdriver.ChromeOptions()
    arguments = ['--lang=pt-BR', '--window-size=1920,1080',
                '--disable-gpu', '--no-sandbox']
    for argument in arguments:
        chrome_options.add_argument(argument)

    prefs = {"download.default_directory": download_dir}
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 10)

    return driver, wait

# Configuração do WebDriver
download_dir = r"C:\Users\homeo\Desktop\FREELANCER\XML\arquivos"  # Caminho do diretório desejado
driver, wait = iniciar_driver()

try:
    # Acessar o site
    url = "https://venda-imoveis.caixa.gov.br/sistema/download-lista.asp"
    driver.get(url)

    # Verificar se o corpo da página foi completamente carregado
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Selecionar o estado MG no menu suspenso
    estado_select = wait.until(EC.presence_of_element_located((By.ID, "cmb_estado")))
    select = Select(estado_select)
    select.select_by_value("MG")  # Seleciona a opção com valor "MG"

    # Clicar no botão "Próximo" para iniciar o download
    btn_proximo = wait.until(EC.element_to_be_clickable((By.ID, "btn_next1")))
    btn_proximo.click()

    # Esperar o download concluir (ajustar tempo conforme necessário)
    time.sleep(3)

    print(f"Arquivo baixado na pasta: {download_dir}")

finally:
    # Fechar o navegador
    driver.quit()


def filtrar():
    # Caminho do arquivo baixado
    download_dir = os.path.expanduser(r"C:\Users\homeo\Desktop\FREELANCER\XML\arquivos")  # Pasta Downloads
    file_name = "Lista_imoveis_MG.csv"
    file_path = os.path.join(download_dir, file_name)

    # Verifica se o arquivo existe
    if not os.path.exists(file_path):
        print(f"Arquivo {file_name} não encontrado no diretório {download_dir}")
        exit()

    # Carregar o arquivo CSV no Pandas
    df = pd.read_csv(file_path, sep=";", encoding="latin1", engine="python")  # Configurar separador e encoding

    # Diagnóstico: Verificar as primeiras linhas do arquivo
    print("Prévia dos dados carregados:")
    print(df.head())

    # Diagnóstico: Verificar nomes das colunas
    print("Colunas disponíveis no arquivo CSV:")
    print(df.columns)

    # Caso o nome da coluna seja correto, aplicar o filtro
    try:
        filtro = df["Modalidade de venda"].isin(["Venda Direta Online", "Venda Online"])
        dados_filtrados = df[filtro]

        # Salvar os dados filtrados em formato .xlsx
        filtered_file = os.path.join(download_dir, "Imoveis_Venda_Direta_Online.xlsx")
        dados_filtrados.to_excel(filtered_file, index=False)

        print(f"Dados filtrados salvos em: {filtered_file}")
    except KeyError:
        print("A coluna 'Modalidade de venda' não foi encontrada. Verifique o nome exato ou tente usar o índice da coluna.")

    # Alternativa: Acessar pela posição da coluna
    if "Modalidade de venda" not in df.columns:
        print("Tentando acessar pela posição da coluna...")
        try:
            # Considerando que 'Modalidade de venda' está na coluna J (índice 9, pois começa em 0)
            coluna_modalidade = df.iloc[:, 9]
            filtro = coluna_modalidade.isin(["Venda Direta Online", "Venda Online"])
            dados_filtrados = df[filtro]

            # Salvar os dados filtrados em formato .xlsx
            filtered_file = os.path.join(download_dir, "Imoveis_Venda_Direta_Online.xlsx")
            dados_filtrados.to_excel(filtered_file, index=False)

            print(f"Dados filtrados salvos em: {filtered_file}")
        except IndexError:
            print("Não foi possível acessar a coluna pela posição. Verifique o arquivo CSV.")

filtrar()

time.sleep(1)

def criar_xml():
    # Caminho do arquivo Excel
    download_dir = os.path.expanduser(r"C:\Users\homeo\Desktop\FREELANCER\XML\arquivos")
    file_name = "Imoveis_Venda_Direta_Online.xlsx"
    caminho_xlsx = os.path.join(download_dir, file_name)

    # Carregar o arquivo .xlsx no Pandas
    df = pd.read_excel(caminho_xlsx)

    # Criação do elemento raiz do XML com namespaces
    ns = {
        "xmlns": "http://www.vivareal.com/schemas/1.0/VRSync",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:schemaLocation": "http://www.vivareal.com/schemas/1.0/VRSync http://xml.vivareal.com/vrsync.xsd"
    }
    root = ET.Element("ListingDataFeed", ns)

    current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    # Header com informações fornecidas
    header = ET.SubElement(root, "Header")
    ET.SubElement(header, "Provider").text = "Bangbot"
    ET.SubElement(header, "Email").text = "luisotavioshinkawa@hotmail.com"
    ET.SubElement(header, "ContactName").text = "Luis Otavio Marcolino Shinkawa "
    ET.SubElement(header, "PublishDate").text = current_time
    ET.SubElement(header, "Telephone").text = "+55 35 9710-0861"

    # Listings
    listings = ET.SubElement(root, "Listings")

    # Função para extrair o número de quartos da descrição
    def extract_bedrooms(description):
        match = re.search(r'(\d+)\s*qto\(s\)', description)
        if match:
            return match.group(1)
        return "0"  # Retorna 0 se não encontrar

    # Itera sobre cada linha do DataFrame e cria os elementos no formato correto
    for _, row in df.iterrows():
        listing = ET.SubElement(listings, "Listing")

        # Campos no XML
        ET.SubElement(listing, "ListingID").text = str(row.iloc[0])
        ET.SubElement(listing, "Title").text = (
            f"Oportunidade Única em {row.iloc[2]} - {row.iloc[1]} | "
            f"Tipo: Casa | Negociação: {row.iloc[9]}"
        )
        ET.SubElement(listing, "TransactionType").text = "For Sale"

        # Criar a URL da imagem com base no código do imóvel
        codigo_imovel = row.iloc[0]  # A primeira coluna (código do imóvel)

        if len(str(codigo_imovel)) <= 12:
            # Para imóveis com código curto, usa o padrão F00000
            img_url = f"https://venda-imoveis.caixa.gov.br/fotos/F00000{codigo_imovel}21.jpg"
        else:
            # Para imóveis com código longo, usa o padrão F + código do imóvel + 21
            img_url = f"https://venda-imoveis.caixa.gov.br/fotos/F{codigo_imovel}21.jpg"

        # Media
        media = ET.SubElement(listing, "Media")
        ET.SubElement(media, "Item", {"medium": "image", "caption": "img0", "primary": "true"}).text = img_url

        # Lista de imagens adicionais fixas
        default_images = [
            "caixa1.png",
            "caixa2.png",
            "caixa3.png",
            "azul.png",
            "caixa4.png",
            "imovelazul.png"
        ]

        # Adiciona as imagens adicionais na ordem especificada
        for idx, image in enumerate(default_images, start=1):
            ET.SubElement(media, "Item", {"medium": "image", "caption": f"img{idx}"}).text = os.path.join("imagens", image)

        # Details
        details = ET.SubElement(listing, "Details")
        ET.SubElement(details, "UsageType").text = "Residential"
        ET.SubElement(details, "PropertyType").text = "Residential / Home"
        ET.SubElement(details, "Description").text = "<![CDATA[ " + str(row.iloc[8]) + " ]]>"
        ET.SubElement(details, "ListPrice", {"currency": "BRL"}).text = str(row.iloc[5])
        ET.SubElement(details, "YearBuilt").text = ""
        ET.SubElement(details, "Bathrooms").text = "1"
        ET.SubElement(details, "Bedrooms").text = extract_bedrooms(row.iloc[8])
        ET.SubElement(details, "Garage", {"type": "Parking Space"}).text = "0"

        # Location
        location = ET.SubElement(listing, "Location", {"displayAddress": "All"})
        ET.SubElement(location, "Country", {"abbreviation": "BR"}).text = "Brasil"
        ET.SubElement(location, "State", {"abbreviation": row.iloc[1]}).text = row.iloc[1]
        ET.SubElement(location, "City").text = row.iloc[2]
        ET.SubElement(location, "Neighborhood").text = row.iloc[3]
        ET.SubElement(location, "Address").text = row.iloc[4]

        # ContactInfo
        contact_info = ET.SubElement(listing, "ContactInfo")
        ET.SubElement(contact_info, "Name").text = ""
        ET.SubElement(contact_info, "Email").text = ""
        ET.SubElement(contact_info, "Website").text = ""
        ET.SubElement(contact_info, "Telephone").text = ""

    # Salvando o XML formatado
    def pretty_print_xml(element):
        rough_string = ET.tostring(element, encoding="utf-8")
        parsed = minidom.parseString(rough_string)
        return parsed.toprettyxml(indent="    ")

    # Criar a pasta 'arquivos' se ela não existir
    if not os.path.exists("arquivos"):
        os.makedirs("arquivos")

    # Salvando o XML
    caminho_xml = os.path.join("arquivos", "Imoveis_Venda_Direta_Online.xml")
    with open(caminho_xml, "w", encoding="utf-8") as f:
        f.write(pretty_print_xml(root))

    print(f"Arquivo XML gerado com sucesso em: {caminho_xml}")

criar_xml()
