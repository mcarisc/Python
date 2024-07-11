import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import requests
from io import BytesIO
from PIL import Image as PILImage, UnidentifiedImageError
from bs4 import BeautifulSoup

# Configuración del WebDriver
chrome_driver_path = ''  # Actualiza esta ruta
service = Service(chrome_driver_path)
options = Options()
options.add_argument("--headless=new")  # Ejecutar en modo headless si no necesitas ver el navegador
options.add_argument("--disable-gpu")  # Desactivar la aceleración por hardware
options.add_argument("--disable-extensions")  # Desactivar extensiones
options.add_argument("--disable-popup-blocking")  # Desactivar bloqueo de ventanas emergentes
options.add_argument("--ignore-certificate-errors")  # Ignorar errores de certificado

driver = webdriver.Chrome(service=service, options=options)

# URL del sitio web
url = 'https://www.saleoutlet.cl/camas-y-colchones/'

# Headers para las solicitudes HTTP
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def abrir_pagina_principal(url):
    driver.get(url)
    #wait = WebDriverWait(driver, 3)
    #wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'product-item')))
    desplazarse_hasta_el_final()
    return driver.page_source

def desplazarse_hasta_el_final():
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # Esperar a que se carguen los productos
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def extraer_productos(html):
    soup = BeautifulSoup(html, 'html.parser')
    productos = soup.find_all('div', class_='product-outer product-item__outer')
    enlaces_productos = [producto.find('a', class_='woocommerce-LoopProduct-link woocommerce-loop-product__link')['href'] for producto in productos if producto.find('a', class_='woocommerce-LoopProduct-link')]
    return enlaces_productos

def obtener_informacion_producto(enlace):
    max_retries = 3
    for _ in range(max_retries):
        try:
            driver.get(enlace)
            time.sleep(3)  # Esperar a que se cargue la página del producto

            nombre = driver.find_element(By.CSS_SELECTOR, '.woocommerce-loop-product__title').text
            
            precio = driver.find_element(By.CSS_SELECTOR, '.woocommerce-Price-amount').text
            descripcion = driver.find_element(By.CSS_SELECTOR, '.woocommerce-product-details__short-description').text

            imagenes = driver.find_elements(By.CSS_SELECTOR, 'div.woocommerce-product-gallery__image img')
            imagen_urls = [img.get_attribute('src') for img in imagenes]

            return nombre, precio, descripcion, imagen_urls
        except Exception as e:
            print(f"Error al procesar el producto en {enlace}: {e}")
            time.sleep(5)  # Esperar antes de reintentar
    return None, None, None, None

def agregar_informacion_a_excel(wb, nombre, precio, descripcion, imagen_urls, fila):
    ws = wb.active
    ws.append([nombre, precio, descripcion, ", ".join(imagen_urls)])
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
                img_cell = f'E{fila}'
                ws.add_image(img, img_cell)
            except (requests.RequestException, UnidentifiedImageError) as e:
                print(f'Error al descargar la imagen de {imagen_url}: {e}')
                ws[f'E{fila}'] = imagen_url  # Escribir la URL de la imagen si hay un error

def main():
    html = abrir_pagina_principal(url)
    enlaces_productos = extraer_productos(html)

    wb = Workbook()
    ws = wb.active
    ws.title = "Productos"
    ws.append(['Nombre', 'Precio', 'Descripción', 'Imagen URL'])

    for idx, enlace in enumerate(enlaces_productos, start=2):
        nombre, precio, descripcion, imagen_urls = obtener_informacion_producto(enlace)
        if nombre and precio and descripcion:
            agregar_informacion_a_excel(wb, nombre, precio, descripcion, imagen_urls, idx)

    wb.save('productos.xlsx')
    print('Los datos de los productos se han guardado en "productos.xlsx".')

    driver.quit()

if __name__ == "__main__":
    main()
