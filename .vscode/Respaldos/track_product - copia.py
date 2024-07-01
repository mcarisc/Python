import time
import requests
from bs4 import BeautifulSoup
#import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook

# Configuración del WebDriver
chrome_driver_path = 'ruta/al/chromedriver'  # Actualiza esta ruta
service = Service(chrome_driver_path)
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Ejecutar en modo headless si no necesitas ver el navegador
driver = webdriver.Chrome(service=service, options=options)

# URL del sitio web
url = 'https://www.saleoutlet.cl/camas-y-colchones/'

# Abrir la página web
driver.get(url)
time.sleep(3)  # Esperar a que la página se cargue

# Desplazarse hasta el final de la página para cargar todos los productos
last_height = driver.execute_script("return document.body.scrollHeight")

while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)  # Esperar a que se carguen los productos
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Obtener el contenido HTML de la página
html = driver.page_source
driver.quit()

# Cabeceras para la solicitud HTTP
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Realizar una solicitud GET al sitio web con las cabeceras
response = requests.get(url, headers=headers)

# Verificar si la solicitud fue exitosa
if response.status_code == 200:
    # Crear un objeto BeautifulSoup con el contenido HTML
    soup = BeautifulSoup(response.content, 'html.parser')

    # Encontrar todos los productos (ajustar las clases según la estructura del sitio web)
    productos = soup.find_all('div', class_='product-outer product-item__outer')

    # Lista para almacenar los datos de los productos
    lista_productos = []

    # Extraer información de cada producto
    for producto in productos:
        nombre = producto.find('h2', class_='woocommerce-loop-product__title').get_text(strip=True)
        precio = producto.find('span', class_='woocommerce-Price-amount amount').get_text(strip=True)
        #print(f'Nombre: {nombre}, Precio: {precio}')
        lista_productos.append([nombre, precio]);

      # Escribir los datos en un archivo CSV
    #with open('productos.csv', 'w', newline='', encoding='utf-8') as file:
    #    writer = csv.writer(file)
    #    writer.writerow(['Nombre', 'Precio'])
    #    writer.writerows(lista_productos)

    # Crear un nuevo libro de Excel y una hoja
    wb = Workbook()
    ws = wb.active
    ws.title = "Productos"

    # Escribir los encabezados en la primera fila
    ws.append(['Nombre', 'Precio'])

     # Escribir los datos de los productos en las filas siguientes
    for producto in lista_productos:
        ws.append(producto)

    # Guardar el archivo Excel
    wb.save('productos.xlsx')

    print('Los datos de los productos se han guardado en "productos.xlsx".')

else:
    print(f'No se pudo acceder al sitio web. Código de estado: {response.status_code}')
