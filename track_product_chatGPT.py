import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import requests
from io import BytesIO
from PIL import Image as PILImage, UnidentifiedImageError

# Configuración del WebDriver
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')  # Ejecutar en modo headless si no necesitas ver el navegador
options.add_argument("--disable-gpu")  # Desactivar la aceleración por hardware
options.add_argument("--disable-extensions")  # Desactivar extensiones
options.add_argument("--disable-popup-blocking")  # Desactivar bloqueo de ventanas emergentes
options.add_argument("--ignore-certificate-errors")  # Ignorar errores de certificado
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
    imagen_tag = producto.find('img', class_='attachment-woocommerce_thumbnail size-woocommerce_thumbnail')
    imagen_url = imagen_tag['src'] if imagen_tag else ''
    lista_productos.append([nombre, precio, imagen_url])


# Crear un nuevo libro de Excel y una hoja
wb = Workbook()
ws = wb.active
ws.title = "Productos"

# Escribir los encabezados en la primera fila
ws.append(['Nombre', 'Precio', 'Imagen'])

    # Escribir los datos de los productos en las filas siguientes
for producto in lista_productos:
    ws.append(producto)

# Descargar y agregar las imágenes a la hoja de cálculo
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

for idx, producto in enumerate(lista_productos, start=2):
    imagen_url = producto[2]
    if imagen_url:
        try:
            response = requests.get(imagen_url, headers=headers)
            response.raise_for_status()

            img_format = imagen_url.split('.')[-1]
            img_content = BytesIO(response.content)
            pil_img = PILImage.open(img_content)

            if img_format.lower() == 'webp':
                pil_img = pil_img.convert("RGB")
                img_content = BytesIO()
                pil_img.save(img_content, format='PNG')
                img_content.seek(0)

            img = Image(img_content)
            img.height = 100  # Ajusta el tamaño de la imagen
            img.width = 100
            img_cell = f'D{idx}'
            ws.add_image(img, img_cell)
        except (requests.RequestException, UnidentifiedImageError) as e:
            print(f'Error al descargar la imagen de {imagen_url}: {e}')
            ws[f'D{idx}'] = imagen_url  # Escribir la URL de la imagen si hay un error

# Guardar el archivo Excel
wb.save('productos.xlsx')

print('Los datos de los productos se han guardado en "productos.xlsx".')
