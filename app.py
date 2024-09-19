from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os
import time
import pandas as pd  # Para crear el archivo Excel

# DOTENV
load_dotenv()

# Configura el navegador
firefox_options = Options()
firefox_options.add_argument("--headless")  # Ejecuta Firefox en modo headless si no quieres ver la ventana del navegador

# Reemplaza con la ruta correcta a tu GeckoDriver
gekko_driver_path = os.getenv("DRIVER")
print(f"[MSG] ------ Driver path:\t{gekko_driver_path}")
service = Service(gekko_driver_path)

# Reemplaza con la ruta correcta a tu ejecutable de Firefox si está en una ubicación no estándar
firefox_path = os.getenv("FIREFOX")
print(f"[MSG] ------ Firefox Path:\t{firefox_path}")
firefox_binary = firefox_path

firefox_options.binary_location = firefox_binary

driver = webdriver.Firefox(service=service, options=firefox_options)

try:
    # Abre la página web
    page_addres = os.getenv("PAGE")
    print(f"[MSG] ------ Page Address:\t{page_addres}")
    driver.get(page_addres)

    # Espera a que el iframe esté presente y cambia el contexto al iframe
    start_time = time.time()
    max_wait_time = 30
    interval = 1  # Cada cuanto tiempo actualizar el contador en segundos

    while True:
        elapsed_time = time.time() - start_time
        if elapsed_time > max_wait_time:
            raise TimeoutError("Se agotó el tiempo de espera para el iframe.")

        print(f"[MSG] ------ Esperando iframe...\n[MSG] ------ Tiempo transcurrido:\t{int(elapsed_time)}s", end="\r")
        
        try:
            iframe = WebDriverWait(driver, interval).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='powerbi.com']"))
            )
            print(f"[MSG] ------ Iframe encontrado después de:\t{int(elapsed_time)} segundos.")
            break
        except Exception as e:
            pass  # Continua esperando

    driver.switch_to.frame(iframe)

    # Espera a que los elementos <span> estén presentes dentro del iframe
    start_time = time.time()
    
    while True:
        elapsed_time = time.time() - start_time
        if elapsed_time > max_wait_time:
            raise TimeoutError("[MSG] ------ Se agotó el tiempo de espera para los elementos <span>.")

        print(f"[MSG] ------ Esperando elementos <span>...\n[MSG] ------ Tiempo transcurrido:\t{int(elapsed_time)}s", end="\r")
        
        try:
            span_elements = WebDriverWait(driver, interval).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.textRun[style*='color: rgb(255, 255, 255);']"))
            )
            print(f"[MSG] ------ Elementos <span> encontrados después de {int(elapsed_time)} segundos.")
            break
        except Exception as e:
            pass  # Continua esperando

    # Extrae el texto de cada elemento <span>
    extracted_data = []
    for span_element in span_elements:
        extracted_text = span_element.text
        print(f"[MSG] ------ Texto extraído: {extracted_text}")
        extracted_data.append(extracted_text)
    
    # Asignar los valores correspondientes a las columnas del Excel
    data = {
        "Fecha": [extracted_data[0]],  # Primer dato
        "Hora": [extracted_data[2]],  # Tercer dato
        "Hidroelectrica (GWh)": [extracted_data[3]],  # Cuarto dato
        "Termoeléctrica (GWh)": [extracted_data[6]],  # Séptimo dato
        "RER (GWh)": [extracted_data[9]],  # Décimo dato
        "Potencia Total Ejecutada (MW)": [extracted_data[12]]  # Decimotercer dato
    }

    # Crear un DataFrame con los datos
    df = pd.DataFrame(data)

    # Guardar el DataFrame en un archivo Excel
    output_path = "datos_extraidos.xlsx"
    df.to_excel(output_path, index=False)
    print(f"[MSG] ------ Datos guardados en {output_path}")

except Exception as e:
    print(f"Error: {e}")

finally:
    # Cierra el navegador
    driver.quit()
