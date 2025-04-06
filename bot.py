import pyautogui
import pytesseract
from PIL import Image
import pandas as pd
import time
import os
import re
import cv2
import numpy as np
from datetime import datetime

# Configuración de Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Nombre del archivo Excel y log
file_name = "numeros.xlsx"
log_file = "log.txt"

# Verificar si el archivo Excel ya existe y cargar hoja "Datos"
if os.path.exists(file_name):
    xls = pd.ExcelFile(file_name)
    if "Datos" in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name="Datos")
    else:
        df = pd.DataFrame(columns=["Timestamp", "Usuario", "Precio", "Cantidad"])
else:
    df = pd.DataFrame(columns=["Timestamp", "Usuario", "Precio", "Cantidad"])

# Variable para almacenar el último conjunto de datos detectados
last_data = None

# Región de captura de pantalla
capture_region = (49, 439, 313, 235)  # Coordenadas actualizadas

def preprocess_image(image_path):
    image = cv2.imread(image_path)
    grayscale = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    equalized = cv2.equalizeHist(grayscale)
    processed = cv2.adaptiveThreshold(equalized, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                      cv2.THRESH_BINARY, 11, 2)
    cv2.imwrite("preprocessed_image.png", processed)
    return processed

def generate_hourly_summary(df):
    if df.empty:
        return pd.DataFrame()

    df["Timestamp"] = pd.to_datetime(df["Timestamp"])
    df["Hora"] = df["Timestamp"].dt.floor("H")

    resumen = []
    for (usuario, hora), data in df.groupby(["Usuario", "Hora"]):
        try:
            precios = [float(p.replace(",", "")) for p in data['Precio']]
            cantidades = [float(c) for c in data['Cantidad']]
            promedio_precio = sum(precios) / len(precios)
            max_disp = max(cantidades)
            min_disp = min(cantidades)
            comprada = max_disp - min_disp

            resumen.append({
                "Usuario": usuario,
                "Intervalo Horario": f"{hora.strftime('%H:%M')} - {(hora + pd.Timedelta(hours=1)).strftime('%H:%M')}",
                "Precio Promedio": round(promedio_precio, 2),
                "Cantidad Comprada": round(comprada, 2)
            })
        except:
            continue

    return pd.DataFrame(resumen)

while True:
    try:
        print("Capturando pantalla...")
        screenshot = pyautogui.screenshot(region=capture_region)
        screenshot.save("screenshot.png")
        preprocessed_image = preprocess_image("screenshot.png")
        text = pytesseract.image_to_string(preprocessed_image, config='--psm 6')
        text = text.strip().replace("\n", " ")

        # Limpiar texto eliminando todos los caracteres que no sean letras, números, puntos, comas o espacios
        text = re.sub(r"[^a-zA-Z0-9.,\s]", "", text)
        text = re.sub(r"\s+", " ", text).strip()
        print(f"Texto extraído limpio: {text}")

        # Expresión regular para extraer usuario, precio y cantidad
        match = re.search(r'([a-zA-Z]{3,}(?: [a-zA-Z]{2,}){0,5})\s+(\d{1,3}(?:,\d{3})*\.\d{2})\s*COP.*?Disponible[: ]\s*(\d{1,3}(?:,\d{3})*\.\d{2})\s*USDT', text, re.IGNORECASE)

        if match:
            usuario, precio, cantidad = match.groups()
            usuario = re.sub(r"[^a-zA-Z]", "", usuario)[:6]  # Solo letras, máximo 6 caracteres
            cantidad = cantidad.replace(",", "")
            current_data = (usuario.lower(), precio, cantidad)

            if last_data is None or current_data != last_data:
                last_data = current_data
                print(f"Datos actualizados: {usuario}, {precio} COP, {cantidad} USDT")
                new_row = pd.DataFrame({"Timestamp": [pd.Timestamp.now()], "Usuario": [usuario], "Precio": [precio], "Cantidad": [cantidad]})
                df = pd.concat([df, new_row], ignore_index=True)

                # Guardar en Excel (hoja Datos y hoja Resumen)
                resumen_df = generate_hourly_summary(df)
                with pd.ExcelWriter(file_name, engine="openpyxl", mode="w") as writer:
                    df.to_excel(writer, sheet_name="Datos", index=False)
                    resumen_df.to_excel(writer, sheet_name="Resumen", index=False)

                # Guardar en log
                with open(log_file, "a", encoding="utf-8") as log:
                    timestamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                    log.write(f"{timestamp} - Usuario: {usuario} | Precio: {precio} COP | Cantidad: {cantidad} USDT\n")
            else:
                print("No hay cambios en los datos.")
        else:
            print("No se detectaron los datos correctamente.")

        os.remove("screenshot.png")
        time.sleep(5)
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        break