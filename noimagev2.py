from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configura el navegador
firefox_options = Options()
firefox_options.add_argument("--headless")  # Ejecuta Firefox en modo headless si no quieres ver la ventana del navegador

# Reemplaza con la ruta correcta a tu GeckoDriver
service = Service('F:/Downloads/geckodriver-v0.35.0-win32/geckodriver.exe')

# Reemplaza con la ruta correcta a tu ejecutable de Firefox si está en una ubicación no estándar
firefox_binary = 'C:/Program Files/Mozilla Firefox/firefox.exe'

firefox_options.binary_location = firefox_binary

driver = webdriver.Firefox(service=service, options=firefox_options)

try:
    # Abre la página web
    driver.get('https://www.coes.org.pe/Portal/portalinformacion/MonitoreoSEIN')

    # Espera a que el iframe esté presente y cambia el contexto al iframe
    iframe = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='powerbi.com']"))
    )
    driver.switch_to.frame(iframe)

    # Espera a que los elementos <span> estén presentes dentro del iframe
    span_elements = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.textRun[style*='color: rgb(255, 255, 255);']"))
    )

    # Extrae el texto de cada elemento <span>
    for span_element in span_elements:
        extracted_text = span_element.text
        print(f"Texto extraído: {extracted_text}")

except Exception as e:
    print(f"Error: {e}")

finally:
    # Cierra el navegador
    driver.quit()
