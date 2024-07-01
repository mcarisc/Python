import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook

# Configuración del WebDriver
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')  # Ejecutar en modo headless si no necesitas ver el navegador
driver = webdriver.Chrome(options=options)
# URL del sitio web
url = 'https://www.saleoutlet.cl/camas-y-colchones/'
# Abrir la página web
driver.get(url)

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


# Crear un objeto BeautifulSoup con el contenido HTML
from bs4 import BeautifulSoup
soup = BeautifulSoup(html, 'html.parser')

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
