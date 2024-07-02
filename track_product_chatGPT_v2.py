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

# Crear un objeto BeautifulSoup con el contenido HTML
from bs4 import BeautifulSoup
soup = BeautifulSoup(html, 'html.parser')

# Encontrar todos los productos (ajustar las clases según la estructura del sitio web)
productos = soup.find_all('div', class_='product-outer product-item__outer')

# Lista para almacenar los datos de los productos
lista_productos = []

# Extraer los enlaces de cada producto
for producto in productos:
    enlace_tag = producto.find('a', class_='woocommerce-LoopProduct-link woocommerce-loop-product__link')
    enlace = enlace_tag['href'] if enlace_tag else ''
    lista_productos.append(enlace)
    #print(enlace)

# Crear un nuevo libro de Excel y una hoja
wb = Workbook()
ws = wb.active
ws.title = "Productos"

# Escribir los encabezados en la primera fila
ws.append(['Nombre', 'Precio', 'Descripción', 'Imagen URL'])


# Descargar y extraer información de cada producto
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def obtener_informacion_producto(enlace):
    try:
        driver.get(enlace)

        # Esperar a que los elementos de los productos se carguen
        nombre = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'product_title entry-title'))
        )

        #time.sleep(3)  # Esperar a que se cargue la página del producto

        # Extraer la información del producto
        #nombre = driver.find_element(By.CSS_SELECTOR, 'product_title entry-title').text
        print(nombre)
        #precio = driver.find_element(By.CLASS_NAME, 'woocommerce-Price-amount amount').text
        #descripcion = driver.find_element(By.CLASS_NAME, 'electro-description clearfix').text
        descripcion = "prueba"
        # Extraer las imágenes del producto
        imagenes = driver.find_elements(By.CSS_SELECTOR, 'ol.flex-control-nav flex-control-thumbs')

        imagen_urls = [img.get_attribute('src') for img in imagenes]

        return nombre, precio, descripcion, imagen_urls
    except Exception as e:
        print(f"Error al procesar el producto en {enlace}: {e}")
        return None, None, None, None


for idx, enlace in enumerate(lista_productos, start=2):
    nombre, precio, descripcion, imagen_urls = obtener_informacion_producto(enlace)
    if nombre and precio and descripcion:
        # Escribir la información del producto en la hoja de cálculo
        ws.append([nombre, precio, descripcion, ", ".join(imagen_urls)])

        # Descargar y agregar las imágenes a la hoja de cálculo
        for img_idx, imagen_url in enumerate(imagen_urls):
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
                    img_cell = f'E{idx}'
                    ws.add_image(img, img_cell)
                except (requests.RequestException, UnidentifiedImageError) as e:
                    print(f'Error al descargar la imagen de {imagen_url}: {e}')
                    ws[f'E{idx}'] = imagen_url  # Escribir la URL de la imagen si hay un error

# Guardar el archivo Excel
wb.save('productos.xlsx')

print('Los datos de los productos se han guardado en "productos.xlsx".')

# Cerrar el navegador
driver.quit()