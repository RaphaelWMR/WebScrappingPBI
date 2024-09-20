import tkinter as tk
from tkinter import filedialog, ttk
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
import threading
import json
from datetime import datetime

# Función para actualizar los mensajes en la interfaz gráfica
def update_output(message):
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    output_text.config(state=tk.NORMAL)
    output_text.insert(tk.END, f"[MSG {timestamp}] {message}\n")
    output_text.config(state=tk.DISABLED)
    output_text.yview(tk.END)  # Desplazar hacia abajo

# Función para actualizar la hora del sistema
def update_clock():
    current_time = datetime.now().strftime("%H:%M:%S")
    clock_label.config(text=f"Hora del sistema: {current_time}")
    root.after(1000, update_clock)

# Función para actualizar el temporizador para la próxima extracción
def update_timer(time_left):
    if time_left > 0:
        timer_label.config(text=f"Próxima extracción en: {time_left} segundos")
        root.after(1000, update_timer, time_left - 1)
    else:
        timer_label.config(text="Extrayendo datos ahora...")
        start_extraction()

# Función para seleccionar el archivo GeckoDriver
def select_driver():
    driver_path = filedialog.askopenfilename(title="Seleccionar GeckoDriver")
    driver_entry.delete(0, tk.END)
    driver_entry.insert(0, driver_path)

# Función para seleccionar el ejecutable de Firefox
def select_firefox():
    firefox_path = filedialog.askopenfilename(title="Seleccionar Firefox")
    firefox_entry.delete(0, tk.END)
    firefox_entry.insert(0, firefox_path)

# Función para ejecutar la extracción de datos
def extraer_datos(output_filename, driver_path, firefox_path, page_url):
    firefox_options = Options()
    firefox_options.add_argument("--headless")
    firefox_options.binary_location = firefox_path

    driver = webdriver.Firefox(service=Service(driver_path), options=firefox_options)
    try:
        update_output(f"Driver path: {driver_path}")
        update_output(f"Firefox Path: {firefox_path}")
        update_output(f"Page Address: {page_url}")

        driver.get(page_url)

        # Esperar por el iframe y cambiar de contexto
        start_time = time.time()
        max_wait_time = 30
        update_output("Esperando iframe...")

        while True:
            elapsed_time = time.time() - start_time
            if elapsed_time > max_wait_time:
                raise TimeoutError("Se agotó el tiempo de espera para el iframe.")

            update_output(f"Tiempo transcurrido: {int(elapsed_time)}s")

            try:
                iframe = WebDriverWait(driver, 1).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='powerbi.com']"))
                )
                update_output(f"Iframe encontrado después de {int(elapsed_time)} segundos.")
                break
            except Exception as e:
                pass

        driver.switch_to.frame(iframe)

        # Esperar por los elementos <span>
        start_time = time.time()
        update_output("Esperando elementos <span>...")

        while True:
            elapsed_time = time.time() - start_time
            if elapsed_time > max_wait_time:
                raise TimeoutError("Se agotó el tiempo de espera para los elementos <span>.")

            update_output(f"Tiempo transcurrido: {int(elapsed_time)}s")

            try:
                span_elements = WebDriverWait(driver, 1).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.textRun[style*='color: rgb(255, 255, 255);']"))
                )
                update_output(f"Elementos <span> encontrados después de {int(elapsed_time)} segundos.")
                break
            except Exception as e:
                pass

        # Extraer el texto de cada elemento <span>
        extracted_data = []
        for span_element in span_elements:
            extracted_text = span_element.text
            update_output(f"Texto extraído: {extracted_text}")
            extracted_data.append(extracted_text)

        # Asignar los valores correspondientes a las columnas del Excel
        data = {
            "Fecha": [extracted_data[0]],
            "Hora": [extracted_data[2]],
            "Hidroelectrica (GWh)": [extracted_data[3]],
            "Termoeléctrica (GWh)": [extracted_data[6]],
            "RER (GWh)": [extracted_data[9]],
            "Potencia Total Ejecutada (MW)": [extracted_data[12]]
        }

        # Crear un DataFrame con los datos
        df = pd.DataFrame(data)

        # Verificar si el archivo existe
        if os.path.exists(output_filename):
            existing_df = pd.read_excel(output_filename)
            df = pd.concat([existing_df, df], ignore_index=True)

        df.to_excel(output_filename, index=False)
        update_output(f"Datos agregados en {output_filename}")

    except Exception as e:
        update_output(f"Error: {e}")

    finally:
        driver.quit()

# Función para iniciar la extracción en un hilo separado
def start_extraction():
    output_filename = file_entry.get() + ".xlsx"
    driver_path = driver_entry.get()
    firefox_path = firefox_entry.get()
    page_url = page_entry.get()

    # Obtener minutos y segundos de las entradas
    try:
        minutes = int(minutes_entry.get())
        seconds = int(seconds_entry.get())
        total_time = minutes * 60 + seconds
    except ValueError:
        update_output("Error: Por favor, ingrese valores válidos para minutos y segundos.")
        return

    # Guardar los cambios en el archivo de configuración
    save_config(driver_path, firefox_path, page_url, minutes, seconds)

    threading.Thread(target=extraer_datos, args=(output_filename, driver_path, firefox_path, page_url)).start()

    # Iniciar el temporizador con el tiempo total
    update_timer(total_time)

# Función para guardar las configuraciones en un archivo json
def save_config(driver_path, firefox_path, page_url, minutes, seconds):
    config = {
        "driver_path": driver_path,
        "firefox_path": firefox_path,
        "page_url": page_url,
        "minutes": minutes,
        "seconds": seconds
    }
    with open("config.json", "w") as config_file:
        json.dump(config, config_file)

# Función para cargar la configuración desde el archivo json
def load_config():
    if os.path.exists("config.json"):
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
            driver_entry.insert(0, config["driver_path"])
            firefox_entry.insert(0, config["firefox_path"])
            page_entry.insert(0, config["page_url"])
            minutes_entry.delete(0, tk.END)
            minutes_entry.insert(0, config.get("minutes", "30"))  # Valor por defecto
            seconds_entry.delete(0, tk.END)
            seconds_entry.insert(0, config.get("seconds", "0"))  # Valor por defecto

# Interfaz gráfica principal
root = tk.Tk()
root.title("Extracción de Datos Automática")

# Reloj de sistema
clock_label = tk.Label(root, text="Hora del sistema: ")
clock_label.pack()

# Temporizador
timer_label = tk.Label(root, text="Próxima extracción en: ")
timer_label.pack()

# Ruta del GeckoDriver
tk.Label(root, text="Ruta del GeckoDriver:").pack()
driver_entry = tk.Entry(root, width=50)
driver_entry.pack()
tk.Button(root, text="Seleccionar GeckoDriver", command=select_driver).pack()

# Ruta del Firefox
tk.Label(root, text="Ruta del Firefox:").pack()
firefox_entry = tk.Entry(root, width=50)
firefox_entry.pack()
tk.Button(root, text="Seleccionar Firefox", command=select_firefox).pack()

# Página web
tk.Label(root, text="Página Web:").pack()
page_entry = tk.Entry(root, width=50)
page_entry.pack()

# Nombre del archivo Excel
tk.Label(root, text="Nombre del archivo Excel:").pack()
file_entry = tk.Entry(root, width=50)
file_entry.insert(0, "datos_extraidos")
file_entry.pack()

# Rango de tiempo para extracción
time_frame = tk.Frame(root)
time_frame.pack(pady=5)

tk.Label(time_frame, text="Rango de tiempo para extracción:").pack(side=tk.TOP)

tk.Label(time_frame, text="Minutos:").pack(side=tk.LEFT)
minutes_entry = tk.Entry(time_frame, width=5)
minutes_entry.insert(0, "30")  # Valor por defecto
minutes_entry.pack(side=tk.LEFT)

tk.Label(time_frame, text="Segundos:").pack(side=tk.LEFT)
seconds_entry = tk.Entry(time_frame, width=5)
seconds_entry.insert(0, "0")  # Valor por defecto
seconds_entry.pack(side=tk.LEFT)

# Cuadro de salida de mensajes con barra de desplazamiento
output_frame = tk.Frame(root)
output_frame.pack(pady=10)

output_text = tk.Text(output_frame, height=10, width=80)
output_text.config(state=tk.DISABLED)
output_text.pack(side=tk.LEFT)

scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
output_text.config(yscrollcommand=scrollbar.set)

# Botón para iniciar la extracción
tk.Button(root, text="Iniciar Extracción", command=start_extraction).pack()

# Cargar la configuración y actualizar la hora
load_config()
update_clock()

# Ejecutar la interfaz gráfica
root.mainloop()
