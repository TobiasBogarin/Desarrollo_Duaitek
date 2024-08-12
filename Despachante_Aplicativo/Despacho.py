import csv
import glob
import os
import json
import requests
import re
import time
import traceback
from tkinter import Label, Entry, Button
from tkinter import Scrollbar
from datetime import datetime
from tkinter.ttk import Treeview, Style, Notebook
from tkinter import Listbox
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from tkinter import Button, Canvas, Entry, Frame, Scrollbar, Toplevel, messagebox, ttk

import customtkinter as ctk
import logging
import sys

# Configuración de logging
logging.basicConfig(filename='app_debug.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# IDs de hojas de cálculo de Google
SPREADSHEET_ID_3 = '1ObCWgChmjd0FHSw1QVHf68txS6zxADA3v4XBjG72wrY'  # flex test
SPREADSHEET_ID_4 = '15JFcSTheCBdDa4hIEzgduHqHJ5vkmyB4-_J12x4lIBo'  # colecta test
SPREADSHEET_ID_CADETES = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'
SPREADSHEET_ID_MULTIPLES = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'
SPREADSHEET_ID_DESPACHO_PENDIENTES = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'
SPREADSHEET_ID_FLEX = '1ObCWgChmjd0FHSw1QVHf68txS6zxADA3v4XBjG72wrY'
SPREADSHEET_ID_COLECTA = '15JFcSTheCBdDa4hIEzgduHqHJ5vkmyB4-_J12x4lIBo'
# Credenciales de servicio de Google
SERVICE_ACCOUNT_FILE = {
    "type": "service_account",
    "project_id": "api-prueba-387019",
    "private_key_id": "523a7ab379d3d59382de0b4dd5215f50f7fd1370",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDg6b66CIW2Q8/7\n6bAudowLEjxwHDvoSZluVTTbJx0oWkrjJTvWhrr595mNTl83S88lhIaIFHng8sTL\nEvsr/Zp1i6bAvgpiPJb8dm7g5IsH5ps+LL+5q0nJW6ZKb29W1/D1u3M9WiYLWH4P\njMehYxsmiHDB3pMRwoQfrJk06iF+xbO5HoWzUMvzuhmUeAbZ48/jbACrnoI+7Lou\n9risvi4AIjyBsxY162cVtyqAXmitXJ9QJAl4EoJoqSyAkl59aLcEi0UWUy0kuSD7\n7Pj3U+z3IyDWsOMld+lpv2cPdtij+/VtApaEYhklU8yYfzb8PG1CxkAFRrx2JP8x\n5IWiZSW3AgMBAAECggEAGfhOLOOdsjqzofONLlrt4hFhp9L6xTWftnJhlLSNf1f9\nsateShprLesocH042AUJjtwglJyYqMrKF7tsroWtTMlVSzLRDCAm5vvd4wXrWnwx\njMT/VmGwNsSTDPaVCpgLRgng9+1C21Mv1dJ8R+b4/u1ePQ8wOCrCYCM+hYegr9HO\nIF4cWMPxX5jF4y+MlsK+iy1s8lI/cmf5p5C0yd5ialtkycuSiMLwlz11UlcCGhf/\nTmose81uhl07F/EU4JELX/DGRJLUuLk4tBqiNdYfMztL2KvTsd1FVQ+PjS1Syc3D\nOEDm5P6eeIP6Fq2SPjUAiFTMmBpxRdlMSPq0Rqt6YQKBgQD3R4Pd2deByHviYfCJ\nW7Xu6z5QAmFxXgowhJPsqhqiHm7FqKKl4hqqEEFdRa50Az8stC7J4yEQmL9hzHTQ\nFWvVeSbozy6u0hyBtv4CZrzIFepZHMhW9KsBpm1sQOAliCZQa8vNr18C1nVLtKq8\nwFKcWt4lG1nmH5NyZb4ybHEXiQKBgQDo2E2Fw8SW+kxC1g7iS29OZk5jSj/qW7S6\n+23y31N+9MmSl9QZIDtS1dlU+6mxjeuaZsXbRDqBZzrpOSw163mUOMZqCOdrxs6k\nRrzRJ+kLWVhivA/eKLGZIbhCgDWtbq0JdKL0kOVKQSoO88OTnR6+4s9Judod46Qs\nnyr/BMvDPwKBgH+P1ObNSe8ZjU7rVzqEpQXrNOnxUHM7H+aHfgfIeJTJPjuZEs6g\nJUE1wYJsP+J5Ck31ZW2gTZ5SLeg1oMz3P/mP1hKjTmHA4hPIYqC6fwh4xbvSrUau\nUMk5IZmGnhq+cYVrFme04D6Gg1vah3l3fSZLee2KfoXIJDgPZF5+spiBAoGBALon\nLBs0PzhhFZUdo7qhinRIcIUK+Hx6IsyWdPmGOC+4rmrHfac00JjSJTW/GZS9HM5N\nOgOp0YhhKoUI02KsRoAMv/xH8BSHVe+aKhyhZrxPCs2tApafPBVsEu7/p2pnoGl9\n2UXjjZzG6kQX+JVMOSdtF0IfFtVsiHWwLuTBRdJrAoGAJ0BSCX/hNZxU7AqBi1gb\nqfHhotVoMVSimAyQT5vjHmdmyST8kFb3wn537jvrWc60PiCJ1nuk/IHXMLKjMByo\niYZ+MQsX2/59XJgx0Bpm75NshmeEkU4qh/4YbDCpYyiFZSXxsPRKwLeloOfX6oKT\n8VW0NaC0SXnsYYqamCTPnSA=\n-----END PRIVATE KEY-----\n",
    "client_email": "cerrador-1@api-prueba-387019.iam.gserviceaccount.com",
    "client_id": "117910477101184573911",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/cerrador-1%40api-prueba-387019.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# Scopes necesarios para leer y escribir en Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Carga las credenciales
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Crea el cliente de servicio para Google Sheets
service = build('sheets', 'v4', credentials=creds)


def obtener_shipment_id(id_interno, spreadsheet_id, rango):
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=rango).execute()
        values = result.get('values', [])

        if len(values) < 2:
            return None  # Asegurarse de que hay al menos una fila de datos

        primer_id = int(values[1][0])  # Obtener el primer ID en la segunda fila
        diferencia = primer_id - int(id_interno)  # Calcular la diferencia
        fila_objetivo = 1 + diferencia  # Calcular la fila objetivo

        if fila_objetivo < 1 or fila_objetivo >= len(values):
            return None  # Asegurarse de que la fila objetivo está dentro del rango válido

        # Print the row and its contents for debugging
        print(f"Fila objetivo: {fila_objetivo + 1}")
        print(f"Contenido de la fila objetivo: {values[fila_objetivo]}")

        return values[fila_objetivo][3]  # Devolver el shipment ID de la columna D (índice 3)
    except Exception as e:
        logging.error(f"Error al obtener shipment ID: {e}")
        return None


def obtener_tokens(clientes):
    url = "https://api.mercadolibre.com/oauth/token"
    tokens = {}
    for nombre, client in clientes.items():
        payload = {
            'grant_type': 'client_credentials',
            'client_id': client['client_id'],
            'client_secret': client['client_secret']
        }
        try:
            response = requests.post(url, data=payload)
            response.raise_for_status()
            tokens[nombre] = response.json()['access_token']
            print(f"Token obtenido con éxito para {nombre}: {tokens[nombre]}")
            logging.info(f"Token obtenido con éxito para {nombre}: {tokens[nombre]}")
        except requests.RequestException as e:
            print(f"Error al obtener token para {nombre}: {e}")
            logging.error(f"Error al obtener token para {nombre}: {e}")
    return tokens

# Inicializar tokens
def inicializar_tokens():
    clientes = {
        'DK': {
            'client_id': '8545932337071872',
            'client_secret': 'VtRNp52gxGKCDKLvR8YRfovPLL3ZBfIk'
        },
        'Tm': {
            'client_id': '8595972831111833',
            'client_secret': 'fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L'
        }
    }
    tokens = obtener_tokens(clientes)
    return tokens

# Funciones para obtener valores y datos de Google Sheets
def obtener_valores(sheet, spreadsheet_id, rango):
    while True:
        try:
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=rango).execute()
            return [item[0] for item in result.get('values', []) if item]
        except HttpError as error:
            if error.resp.status == 429:
                logging.warning("Se ha superado el límite de cuota, esperando 10 segundos para reintentar...")
                time.sleep(10)  # Espera 10 segundos antes de reintentar
            else:
                logging.error(f"Error al obtener valores: {error}")
                raise  # Vuelve a lanzar la excepción si no es un error de cuota

def obtener_datos(sheet, spreadsheet_id, rango):
    while True:
        try:
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=rango).execute()
            return result.get('values', [])
        except HttpError as error:
            if error.resp.status == 429:
                logging.warning("Se ha superado el límite de cuota, esperando 10 segundos para reintentar...")
                time.sleep(10)  # Espera 10 segundos antes de reintentar
            else:
                logging.error(f"Error al obtener datos: {error}")
                raise  # Vuelve a lanzar la excepción si no es un error de cuota

# Verificar el estado del shipment
def verificar_estado_shipment(tokens, shipment_id):
    url = f"https://api.mercadolibre.com/shipments/{shipment_id}"
    headers = {}
    
    for nombre, token in tokens.items():
        headers['Authorization'] = f'Bearer {token}'
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            shipment_data = response.json()
            return shipment_data
        except requests.RequestException as e:
            print(f"Error al verificar estado con el token de {nombre}: {e}")
            logging.error(f"Error al verificar estado con el token de {nombre}: {e}")
    
    messagebox.showerror("Error", "No se pudo verificar el estado del shipment con los tokens disponibles.")
    return None

def verificar_estado_token(shipment_id, tokens):
    url_template = "https://api.mercadolibre.com/shipments/{}"
    
    for account, token in tokens.items():
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url_template.format(shipment_id), headers=headers)
        
        if response.status_code == 200:
            shipment_data = response.json()
            status = shipment_data.get('status')
            if status == 'cancelled':
                messagebox.showwarning("Alerta", "No despachar este pedido porque está cancelado.")
            else:
                messagebox.showinfo("Verificación de Estado", "El pedido no se encuentra cancelado.")
                print(f"Status del pedido {shipment_id} para {account}: {status}")
                print(json.dumps(shipment_data, indent=2))  # Para imprimir el JSON formateado en la terminal
            return shipment_data  # Devuelve el JSON como objeto Python
        else:
            print(f"Error con la cuenta {account}: {response.status_code} - {response.text}")
            logging.error(f"Error con la cuenta {account}: {response.status_code} - {response.text}")

    print("No se pudo obtener el estado del shipment con los tokens proporcionados.")
    logging.error("No se pudo obtener el estado del shipment con los tokens proporcionados.")
    return None  # Devuelve None si no se puede obtener el estado del shipment
# Inicialización de la ventana
app = ctk.CTk()
app.title("Sistema de Escaneo")
app.state('zoomed')  # Maximiza la ventana

notebook = Notebook(app)
notebook.grid(row=0, column=0, columnspan=9, padx=40, pady=40, sticky="ew")

# Pestaña principal
tab_principal = ctk.CTkFrame(notebook)
notebook.add(tab_principal, text="Principal")

# Contadores de escaneo y listas de pedidos
contador_envios_flex = 0
contador_envios_colecta = 0
pedidos_envios_flex = []
pedidos_envios_colecta = []
estado_checkboxes = {"GESTIONAR": False, "GOFLEX": False, "TDP": False, "NESTOR": False}

def id_ya_registrado(id_escaneado):
    ya_registrado_flex = any(id_escaneado == pedido[0] for pedido in pedidos_envios_flex)
    ya_registrado_colecta = id_escaneado in pedidos_envios_colecta
    return ya_registrado_flex or ya_registrado_colecta

def confirmar_cambios(entries, tipo, ventana):
    global contador_envios_flex, contador_envios_colecta, pedidos_envios_flex, pedidos_envios_colecta

    try:
        if tipo == 'flex':
            nuevos_pedidos = []
            for i, entry in enumerate(entries):
                nuevo_id = entry.get().strip()
                if nuevo_id:  # Añadir solo si el ID no está vacío
                    nuevos_pedidos.append((nuevo_id, pedidos_envios_flex[i][1]))  # Asume que la segunda columna es algo relevante para Flex
            pedidos_envios_flex = nuevos_pedidos
            contador_envios_flex = len(pedidos_envios_flex)
            label_envios_flex.configure(text=f"Envíos Flex: {contador_envios_flex}")
        elif tipo == 'colecta':
            nuevos_pedidos = []
            for i, entry in enumerate(entries):
                nuevo_id = entry.get().strip()
                if nuevo_id:
                    nuevos_pedidos.append(nuevo_id)
            pedidos_envios_colecta = nuevos_pedidos
            contador_envios_colecta = len(pedidos_envios_colecta)
            label_envios_colecta.configure(text=f"Envíos Colecta: {contador_envios_colecta}")

        ventana.destroy()
        messagebox.showinfo("Actualización", "Los cambios se han guardado correctamente.")
    except Exception as e:
        logging.error(f"Error al confirmar cambios: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al guardar los cambios: {e}")

def mostrar_ids_almacenados():
    try:
        ids_ventana = Toplevel(app)
        ids_ventana.title("IDs Almacenados")
        ids_ventana.geometry("700x400")

        # Frame y Canvas para envíos Flex
        frame_flex_ext = Frame(ids_ventana, borderwidth=2, relief="groove")
        frame_flex_ext.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        canvas_flex = Canvas(frame_flex_ext)
        scrollbar_flex = Scrollbar(frame_flex_ext, orient="vertical", command=canvas_flex.yview)
        frame_flex_int = Frame(canvas_flex)

        def on_frame_flex_configure(event):
            canvas_flex.configure(scrollregion=canvas_flex.bbox("all"))

        frame_flex_int.bind("<Configure>", on_frame_flex_configure)

        canvas_flex.create_window((0, 0), window=frame_flex_int, anchor="nw")
        canvas_flex.configure(yscrollcommand=scrollbar_flex.set)

        canvas_flex.pack(side="left", fill="both", expand=True)
        scrollbar_flex.pack(side="right", fill="y")

        Label(frame_flex_int, text="Envíos Flex").pack(pady=5)

        entries_flex = []
        for i, pedido in enumerate(pedidos_envios_flex):
            entry = Entry(frame_flex_int)
            entry.insert(0, pedido[0])
            entry.pack(pady=2, fill='x', expand=True)
            entries_flex.append(entry)

        # Frame y Canvas para envíos Colecta
        frame_colecta_ext = Frame(ids_ventana, borderwidth=2, relief="groove")
        frame_colecta_ext.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        canvas_colecta = Canvas(frame_colecta_ext)
        scrollbar_colecta = Scrollbar(frame_colecta_ext, orient="vertical", command=canvas_colecta.yview)
        frame_colecta_int = Frame(canvas_colecta)

        def on_frame_colecta_configure(event):
            canvas_colecta.configure(scrollregion=canvas_colecta.bbox("all"))

        frame_colecta_int.bind("<Configure>", on_frame_colecta_configure)

        canvas_colecta.create_window((0, 0), window=frame_colecta_int, anchor="nw")
        canvas_colecta.configure(yscrollcommand=scrollbar_colecta.set)

        canvas_colecta.pack(side="left", fill="both", expand=True)
        scrollbar_colecta.pack(side="right", fill="y")

        Label(frame_colecta_int, text="Envíos Colecta").pack(pady=5)

        entries_colecta = []
        for i, pedido in enumerate(pedidos_envios_colecta):
            entry = Entry(frame_colecta_int)
            entry.insert(0, pedido)
            entry.pack(pady=2, fill='x', expand=True)
            entries_colecta.append(entry)

        Button(frame_flex_int, text="Confirmar Cambios Flex", command=lambda: confirmar_cambios(entries_flex, 'flex', ids_ventana)).pack(pady=10)
        Button(frame_colecta_int, text="Confirmar Cambios Colecta", command=lambda: confirmar_cambios(entries_colecta, 'colecta', ids_ventana)).pack(pady=10)
    except Exception as e:
        logging.error(f"Error al mostrar IDs almacenados: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al mostrar los IDs almacenados: {e}")

def actualizar_contador_envios_flex(pedidos, es_multiple=False):
    global contador_envios_flex
    try:
        contador_envios_flex += len(pedidos)
        # Captura la cadetería seleccionada en el momento del registro del pedido desde el Listbox
        if cadeteria_listbox.curselection():
            cadeteria_seleccionada = cadeterias[cadeteria_listbox.curselection()[0]]
        else:
            cadeteria_seleccionada = "No especificada"

        # Asocia los pedidos con la cadetería seleccionada y marca si son envíos múltiples
        pedidos_envios_flex.extend((pedido, cadeteria_seleccionada, "SI" if es_multiple else "NO") for pedido in pedidos)
        label_envios_flex.configure(text=f"Envíos Flex: {contador_envios_flex}")
    except Exception as e:
        logging.error(f"Error al actualizar contador de envíos Flex: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al actualizar el contador de envíos Flex: {e}")

def actualizar_contador_envios_colecta(pedidos):
    global contador_envios_colecta
    try:
        contador_envios_colecta += len(pedidos)
        pedidos_envios_colecta.extend(pedidos)
        label_envios_colecta.configure(text=f"Envíos Colecta: {contador_envios_colecta}")
    except Exception as e:
        logging.error(f"Error al actualizar contador de envíos Colecta: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al actualizar el contador de envíos Colecta: {e}")

def verificar_y_procesar_id(id_escaneado):
    try:
        if 20000000 <= int(id_escaneado) <= 29999999:
            # Procesar como un pedido de colecta
            if not id_ya_registrado(id_escaneado):  # Verifica si el ID ya está registrado
                procesar_id([id_escaneado])  # Procesa si no está duplicado
            else:
                messagebox.showerror("Pedido repetido", f"El ID {id_escaneado} ya ha sido escaneado.")
                mostrar_ids_almacenados()  # Muestra la interfaz de IDs repetidos
        else:
            try:
                # Buscar el ID en la hoja 'despacho_pendientes' y obtener las coordenadas
                despacho_pendientes = obtener_datos(service.spreadsheets(), SPREADSHEET_ID_DESPACHO_PENDIENTES, 'despacho_pendientes!Y:AE')
                ids_despacho = [fila[0] for fila in despacho_pendientes if fila]

                if id_escaneado in ids_despacho:
                    indice = ids_despacho.index(id_escaneado)
                    if len(despacho_pendientes[indice]) >= 7:  # Asegurarse de que la fila tenga al menos 7 columnas (Y:AE)
                        longitud = despacho_pendientes[indice][5]  # Columna AD es la sexta columna en el rango Y:AE
                        latitud = despacho_pendientes[indice][6]  # Columna AE es la séptima columna en el rango Y:AE

                        # Buscar otros IDs con las mismas coordenadas
                        ids_coincidentes = [
                            fila[0] for fila in despacho_pendientes
                            if len(fila) >= 7 and fila[5] == longitud and fila[6] == latitud and fila[0] != id_escaneado
                        ]

                        if ids_coincidentes:
                            titulo = "A los pedidos de la lista"
                            mostrar_sub_interfaz_escaneo(id_escaneado, ids_coincidentes, titulo)
                        else:
                            procesar_id([id_escaneado])
                    else:
                        messagebox.showerror("Error", f"El ID {id_escaneado} no tiene suficientes datos en la hoja 'despacho_pendientes'.")
                else:
                    messagebox.showerror("Error", f"El ID {id_escaneado} no se encontró en la hoja 'despacho_pendientes'.")
            except HttpError as error:
                logging.error(f"Error al intentar acceder a la hoja de cálculo: {error}")
                messagebox.showerror("Error", f"Error al intentar acceder a la hoja de cálculo: {error}")
    except Exception as e:
        logging.error(f"Error al verificar y procesar ID: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al verificar y procesar el ID: {e}")

def procesar_id(ids_escaneados, es_multiple=False):
    try:
        for id_escaneado in ids_escaneados:
            if id_escaneado.isdigit():  # Asegura que el ID sea numérico
                if id_ya_registrado(id_escaneado):
                    messagebox.showerror("Pedido repetido", f"El ID {id_escaneado} ya ha sido escaneado.")
                    mostrar_ids_almacenados()
                    return  # Termina la función si se encuentra un ID repetido
                else:
                    if 10000000 <= int(id_escaneado) <= 19999999:
                        actualizar_contador_envios_flex([id_escaneado], es_multiple)
                    elif 20000000 <= int(id_escaneado) <= 29999999:
                        actualizar_contador_envios_colecta([id_escaneado])
            entry_text.delete(0, 'end')  # Limpia el campo de texto después de procesar
    except Exception as e:
        logging.error(f"Error al procesar ID: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al procesar el ID: {e}")

def enviar_envios_multiples(ids):
    try:
        sheet = service.spreadsheets()
        
        # Obtener los datos existentes en la hoja ENVIOS_MULTIPLES
        existing_data = obtener_datos(sheet, SPREADSHEET_ID_MULTIPLES, 'ENVIOS_MULTIPLES!A:C')
        
        # Determinar la próxima fila vacía
        next_row = len(existing_data) + 1
        
        # Preparar los datos para enviar
        fecha_actual = datetime.now().strftime('%d/%m/%Y')
        data = []
        
        for row, id in enumerate(ids, start=next_row):
            # Insertar el ID en la columna A
            data.append({'range': f"ENVIOS_MULTIPLES!A{row}", 'values': [[id]]})
            
            # Insertar los IDs asociados en la columna B, excluyendo el ID actual
            ids_asociados = [i for i in ids if i != id]
            data.append({'range': f"ENVIOS_MULTIPLES!B{row}", 'values': [[", ".join(ids_asociados)]]})
            
            # Insertar la fecha en la columna C
            data.append({'range': f"ENVIOS_MULTIPLES!C{row}", 'values': [[fecha_actual]]})
        
        # Enviar los datos a Google Sheets
        body = {'valueInputOption': 'RAW', 'data': data}
        sheet.values().batchUpdate(spreadsheetId=SPREADSHEET_ID_MULTIPLES, body=body).execute()
        logging.info(f"Envíos múltiples actualizados en la hoja ENVIOS_MULTIPLES.")
    except HttpError as error:
        logging.error(f"Ocurrió un error al intentar enviar los envíos múltiples: {error}")
        messagebox.showerror("Error", "No se pudo actualizar la lista de envíos múltiples en Google Sheets.")
    except Exception as e:
        logging.error(f"Error inesperado al enviar envíos múltiples: {e}")
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")

def mostrar_sub_interfaz_escaneo(id_escaneado, ids_asociados, titulo):
    try:
        sub_interfaz = Toplevel(app)
        sub_interfaz.title(f"DEBE SALIR CON {titulo}")
        sub_interfaz.grab_set()  # Asegura que la sub interfaz esté activa

        style = Style()
        style.configure("Treeview", font=("Helvetica", 12))
        style.configure("Treeview.Heading", font=("Helvetica", 14, "bold"))

        tree = Treeview(sub_interfaz, columns=("ID",), show="headings", height=10)
        tree.heading("ID", text="ID")
        tree.pack(pady=10, padx=10, fill="both", expand=True)

        for id_asociado in ids_asociados:
            tree.insert("", "end", values=(id_asociado,))

        ids_verificados = set()

        def on_sub_interfaz_enter(event):
            id_verificacion = entry_verificacion.get().strip()
            for item in tree.get_children():
                item_id = tree.item(item, "values")[0]
                if id_verificacion == item_id:
                    tree.item(item, tags=("verified",))
                    ids_verificados.add(id_verificacion)
                    style.configure("Treeview", background="white", foreground="black")
                    style.map("Treeview", background=[("selected", "green")])
                    tree.tag_configure("verified", background="green", foreground="white")
                    entry_verificacion.delete(0, 'end')
                    if ids_verificados == set(ids_asociados):
                        ids_totales = [id_escaneado] + list(ids_verificados)
                        procesar_id(ids_totales, es_multiple=True)
                        enviar_envios_multiples(ids_totales)  # Enviar los IDs a la hoja de Google Sheets
                        sub_interfaz.destroy()
                        entry_text.configure(state='normal')
                        entry_text.delete(0, 'end')
                    break
            else:
                messagebox.showerror("Error", "El ID escaneado no coincide con los pedidos asociados.")
                # Mantener el enfoque en la sub interfaz y limpiar el campo de entrada
                sub_interfaz.lift()  
                entry_verificacion.delete(0, 'end')

        def no_almacenar():
            sub_interfaz.destroy()
            entry_text.configure(state='normal')

        label_titulo = Label(sub_interfaz, text=f"DEBE SALIR CON {titulo}", font=("Helvetica", 16, "bold"))
        label_titulo.pack(pady=10)

        entry_verificacion = Entry(sub_interfaz)
        entry_verificacion.pack(pady=10)
        entry_verificacion.bind("<Return>", on_sub_interfaz_enter)

        boton_no_almacenar = Button(sub_interfaz, text="No Almacenar", command=no_almacenar)
        boton_no_almacenar.pack(pady=10)

        # Deshabilitar la entrada principal hasta que se cierre la sub interfaz
        entry_text.configure(state='disabled')
    except Exception as e:
        logging.error(f"Error al mostrar sub interfaz de escaneo: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al mostrar la sub interfaz de escaneo: {e}")

def on_enter(event):
    if entry_text.cget('state') == 'normal':  # Asegura que el campo de entrada esté activo
        id_escaneado = entry_text.get().strip()
        verificar_y_procesar_id(id_escaneado)  # Verifica y procesa el ID

def alphanum_key(s):
    """Divide la cadena en una lista de cadenas y números."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

# Funciones para generar archivos CSV y resetear contadores
def generar_csv_envios_flex():
    global contador_envios_flex
    try:
        if not pedidos_envios_flex:
            logging.warning("No hay envíos flex para guardar en el CSV.")
            messagebox.showwarning("Advertencia", "No hay envíos flex para guardar en el CSV.")
            return

        filename = f"envios_flex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        # Obtiene la cadetería seleccionada en el momento de finalizar el despacho desde el Listbox
        if cadeteria_listbox.curselection():
            cadeteria_seleccionada = cadeterias[cadeteria_listbox.curselection()[0]]
        else:
            cadeteria_seleccionada = "No especificada"

        # Ordena los pedidos por ID antes de escribirlos
        pedidos_envios_flex.sort(key=lambda x: alphanum_key(x[0] if len(x) > 0 else ''))

        # Guardar una copia de seguridad de los pedidos actuales antes de borrarlos
        with open('backup_flex_ids.csv', 'w', newline='', encoding='utf-8') as backup_file:
            backup_writer = csv.writer(backup_file)
            backup_writer.writerow(["ID", "Cadetería", "Envío Múltiple"])
            for pedido in pedidos_envios_flex:
                if len(pedido) != 3:
                    pedido = (pedido[0], pedido[1], "No") if len(pedido) == 2 else (pedido[0], "No especificada", "No")
                backup_writer.writerow(pedido)

        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["ID", "Cadetería", "Envío Múltiple"])
            for pedido in pedidos_envios_flex:
                if len(pedido) != 3:
                    pedido = (pedido[0], pedido[1], "No") if len(pedido) == 2 else (pedido[0], "No especificada", "No")
                writer.writerow(pedido)
                
        logging.info(f"Archivo CSV {filename} generado exitosamente.")
        pedidos_envios_flex.clear()
        contador_envios_flex = 0
        label_envios_flex.configure(text="Envíos Flex: 0")
        
        generar_pdf_desde_csv(filename, cadeteria_seleccionada)
        messagebox.showinfo("Éxito", f"Archivo CSV {filename} generado y guardado exitosamente.")
    except Exception as e:
        logging.error(f"Error al generar CSV de envíos Flex: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al generar el CSV de envíos Flex: {e}")


def generar_csv_envios_colecta():
    global contador_envios_colecta
    try:
        if not pedidos_envios_colecta:
            logging.warning("No hay envíos de colecta para guardar en el CSV.")
            messagebox.showwarning("Advertencia", "No hay envíos de colecta para guardar en el CSV.")
            return

        filename = f"envios_colecta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        # Guardar una copia de seguridad de los pedidos actuales antes de borrarlos
        with open('backup_colecta_ids.csv', 'w', newline='', encoding='utf-8') as backup_file:
            backup_writer = csv.writer(backup_file)
            backup_writer.writerow(["ID"])
            for pedido in pedidos_envios_colecta:
                backup_writer.writerow([pedido])

        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["ID"])
            for pedido in pedidos_envios_colecta:
                writer.writerow([pedido])
        logging.info(f"Archivo CSV {filename} generado exitosamente.")
        pedidos_envios_colecta.clear()
        contador_envios_colecta = 0
        label_envios_colecta.configure(text="Envíos Colecta: 0")
        messagebox.showinfo("Éxito", f"Archivo CSV {filename} generado y guardado exitosamente.")
    except Exception as e:
        logging.error(f"Error al generar CSV de envíos Colecta: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al generar el CSV de envíos Colecta: {e}")

def asignar_a_cadeteria():
    try:
        lista_de_archivos = glob.glob('*.csv')
        if not lista_de_archivos:
            logging.warning("No se encontraron archivos CSV en el directorio.")
            messagebox.showerror("Error", "No se encontraron archivos CSV en el directorio.")
            return

        ultimo_archivo = max(lista_de_archivos, key=os.path.getctime)
        df = pd.read_csv(ultimo_archivo)
        ids_envios = df.iloc[:, 0].astype(str).tolist()
        cadeterias = df.iloc[:, 1].tolist()

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Obtener todos los IDs de la columna A en la hoja 'Despacho'
        ids_hoja = obtener_valores(sheet, SPREADSHEET_ID_3, "Despacho!A:A")

        MAX_UPDATES_PER_BATCH = 99
        batches = []
        current_batch = []

        for i, id_envio in enumerate(ids_envios):
            try:
                fila = ids_hoja.index(id_envio) + 1  # +1 porque los índices de la hoja comienzan en 1
                current_batch.extend([
                    {'range': f"Despacho!K{fila}", 'values': [[cadeterias[i]]]},
                    {'range': f"Despacho!M{fila}", 'values': [["ENVIADO"]]},
                    {'range': f"Despacho!O{fila}", 'values': [[datetime.now().strftime("%d/%m/%Y")]]}
                ])
            except ValueError:
                logging.warning(f"ID {id_envio} no encontrado en la hoja 'Despacho'.")

            if len(current_batch) >= MAX_UPDATES_PER_BATCH * 3:
                batches.append(current_batch)
                current_batch = []

        if current_batch:
            batches.append(current_batch)

        for batch in batches:
            body = {'valueInputOption': 'RAW', 'data': batch}
            sheet.values().batchUpdate(spreadsheetId=SPREADSHEET_ID_3, body=body).execute()

        logging.info(f"Datos de cadetería actualizados para los IDs en {ultimo_archivo}.")
        messagebox.showinfo("Éxito", f"Datos de cadetería actualizados para los IDs en {ultimo_archivo}.")
    except Exception as e:
        logging.error(f"Error al asignar cadetería: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al asignar cadetería: {e}")

def asignar_colecta():
    try:
        lista_de_archivos = glob.glob('*envios_colecta*.csv')
        if not lista_de_archivos:
            logging.warning("No se encontraron archivos CSV de colecta en el directorio.")
            messagebox.showerror("Error", "No se encontraron archivos CSV de colecta en el directorio.")
            return

        ultimo_archivo = max(lista_de_archivos, key=os.path.getctime)
        df = pd.read_csv(ultimo_archivo)
        ids_envios = df.iloc[:, 0].astype(str).tolist()

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Obtener todos los IDs de la columna A en la hoja 'Despacho'
        ids_hoja = obtener_valores(sheet, SPREADSHEET_ID_4, "Despacho!A:A")

        MAX_UPDATES_PER_BATCH = 99
        batches = []
        current_batch = []

        for id_envio in ids_envios:
            try:
                fila = ids_hoja.index(id_envio) + 1  # +1 porque los índices de la hoja comienzan en 1
                update_data = [
                    {'range': f"Despacho!K{fila}", 'values': [["COLECTA"]]},
                    {'range': f"Despacho!N{fila}", 'values': [["ENVIADO"]]},
                    {'range': f"Despacho!L{fila}", 'values': [[datetime.now().strftime("%d/%m/%Y")]]}
                ]
                current_batch.extend(update_data)
            except ValueError:
                logging.warning(f"ID {id_envio} no encontrado en la hoja 'Despacho'.")

            if len(current_batch) >= MAX_UPDATES_PER_BATCH * 3:
                batches.append(current_batch)
                current_batch = []

        if current_batch:
            batches.append(current_batch)

        for batch in batches:
            body = {'valueInputOption': 'RAW', 'data': batch}
            sheet.values().batchUpdate(spreadsheetId=SPREADSHEET_ID_4, body=body).execute()

        logging.info(f"Datos de colecta actualizados para los IDs en {ultimo_archivo}.")
        messagebox.showinfo("Éxito", f"Datos de colecta actualizados para los IDs en {ultimo_archivo}.")
    except Exception as e:
        logging.error(f"Error al asignar colecta: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al asignar colecta: {e}")

# Función para abrir una nueva ventana y almacenar el pedido
def almacenar_pedido():
    def guardar_pedido():
        try:
            pedido = entry_pedido.get().strip()
            if pedido.isdigit():  # Verifica que el ID sea numérico
                id_escaneado = int(pedido)
                if 10000000 <= id_escaneado <= 19999999:
                    actualizar_contador_envios_flex([pedido])
                elif 20000000 <= int(id_escaneado) <= 29999999:
                    actualizar_contador_envios_colecta([pedido])
                almacenar_pedido_ventana.destroy()
            else:
                messagebox.showerror("Error", "El ID debe ser un número.")
        except Exception as e:
            logging.error(f"Error al almacenar pedido: {e}")
            messagebox.showerror("Error", f"Ocurrió un error al almacenar el pedido: {e}")

    try:
        almacenar_pedido_ventana = ctk.CTkToplevel(app)
        almacenar_pedido_ventana.title("Almacenar Pedido")
        almacenar_pedido_ventana.focus_force()
        almacenar_pedido_ventana.grab_set()

        label = ctk.CTkLabel(almacenar_pedido_ventana, text="Ingrese el número de pedido:")
        label.pack(pady=5)

        entry_pedido = ctk.CTkEntry(almacenar_pedido_ventana)
        entry_pedido.pack(pady=5)

        boton_guardar = ctk.CTkButton(almacenar_pedido_ventana, text="Guardar", command=guardar_pedido)
        boton_guardar.pack(pady=5)
    except Exception as e:
        logging.error(f"Error al abrir ventana para almacenar pedido: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al abrir la ventana para almacenar pedido: {e}")

# Función para cargar las cadeterías desde Google Sheets
def cargar_cadeterias():
    try:
        sheet = service.spreadsheets()
        cadeterias = obtener_valores(sheet, SPREADSHEET_ID_CADETES, 'CADETES!A:A')
        return cadeterias
    except HttpError as error:
        logging.error(f"Ocurrió un error al intentar recuperar las cadeterías: {error}")
        messagebox.showerror("Error", "No se pudo cargar la lista de cadeterías desde Google Sheets.")
        return []
    except Exception as e:
        logging.error(f"Error inesperado al cargar cadeterías: {e}")
        messagebox.showerror("Error", f"Ocurrió un error inesperado al cargar cadeterías: {e}")
        return []

# Widgets de la pestaña principal
entry_text = ctk.CTkEntry(tab_principal, placeholder_text="Escanea o ingresa un ID")
entry_text.grid(row=0, column=0, columnspan=9, padx=40, pady=40, sticky="ew")
entry_text.bind("<Return>", on_enter)

label_envios_flex = ctk.CTkLabel(tab_principal, text="Envíos Flex: 0", font=('Roboto', 18))
label_envios_flex.grid(row=2, column=0, columnspan=9, padx=5, pady=5, sticky="ew")

label_envios_colecta = ctk.CTkLabel(tab_principal, text="Envíos Colecta: 0", font=('Roboto', 18))
label_envios_colecta.grid(row=4, column=0, columnspan=9, padx=5, pady=5, sticky="ew")

font_size = 20  # Tamaño de la fuente triplicado
button_width = 300  # Establece el ancho deseado
button_height = 11  # Establece el alto deseado

boton_almacenar_pedido = ctk.CTkButton(tab_principal, text="ALMACENAR PEDIDO", command=almacenar_pedido, corner_radius=20, font=('Roboto', font_size), width=button_width, height=button_height)
boton_asignar_cadeteria = ctk.CTkButton(tab_principal, text="ASIGNAR CADETERIA", command=asignar_a_cadeteria, corner_radius=20, font=('Roboto', font_size), width=button_width, height=button_height)
boton_asignar_colecta = ctk.CTkButton(tab_principal, text="ASIGNAR COLECTA", command=asignar_colecta, corner_radius=20, font=('Roboto', font_size), width=button_width, height=button_height)
boton_terminar_flex = ctk.CTkButton(tab_principal, text="Terminar Despacho Flex", command=generar_csv_envios_flex, corner_radius=20, font=('Roboto', font_size), width=button_width, height=button_height)
boton_terminar_colecta = ctk.CTkButton(tab_principal, text="Terminar Despacho Colecta", command=generar_csv_envios_colecta, corner_radius=20, font=('Roboto', font_size), width=button_width, height=button_height)

boton_almacenar_pedido.grid(row=6, column=0, pady=20, padx=20, sticky="ew")
boton_asignar_cadeteria.grid(row=6, column=2, pady=20, padx=20, sticky="ew")
boton_asignar_colecta.grid(row=6, column=4, pady=20, padx=20, sticky="ew")
boton_terminar_flex.grid(row=8, column=0, pady=20, padx=20, columnspan=2, sticky="ew")
boton_terminar_colecta.grid(row=8, column=4, pady=20, padx=20, sticky="ew")

font_size = 16  # Tamaño de la fuente, el mismo que los botones
listbox_width = 15  # Ajusta este valor según sea necesario para tu contenido

cadeteria_listbox = Listbox(tab_principal, height=8, width=listbox_width, exportselection=False, font=('Roboto', font_size))
cadeteria_listbox.grid(row=10, column=0, columnspan=6, padx=20, pady=20, sticky="nsew")

# Creación del Scrollbar
scrollbar = Scrollbar(tab_principal, orient="vertical")
scrollbar.grid(row=10, column=5, sticky='ns', padx=(0, 20))  # Asegura que el scrollbar se adhiera al norte y sur (top and bottom)

# Configuración del Scrollbar para controlar el yview del Listbox
scrollbar.config(command=cadeteria_listbox.yview)
# Configuración del Listbox para comunicarse con el Scrollbar
cadeteria_listbox.config(yscrollcommand=scrollbar.set)

# Cargar las cadeterías desde Google Sheets
cadeterias = cargar_cadeterias()

# Agregar elementos al Listbox
for cadeteria in cadeterias:
    cadeteria_listbox.insert('end', cadeteria)

def generar_pdf_desde_csv(nombre_archivo_csv, cadeteria):
    try:
        datos = pd.read_csv(nombre_archivo_csv, encoding='utf-8')
    except UnicodeDecodeError:
        datos = pd.read_csv(nombre_archivo_csv, encoding='latin-1')  # Intenta con otra codificación si falla utf-8

    ahora = datetime.now()
    titulo_pdf = f"Remito {cadeteria} {ahora.strftime('%d/%m/%Y (%H:%M)')}"
    nombre_archivo_pdf = f"{titulo_pdf}.pdf".replace('/', '_').replace(':', '_').replace(' ', '_')

    pdf = SimpleDocTemplate(nombre_archivo_pdf, pagesize=letter)
    elementos = []

    estilos = getSampleStyleSheet()
    estiloTitulo = estilos['Normal']
    estiloTitulo.alignment = 1
    estiloTitulo.fontSize = 14
    estiloTitulo.spaceAfter = 20

    parrafo_titulo = Paragraph(titulo_pdf, estiloTitulo)
    elementos.append(parrafo_titulo)

    data = [datos.columns.to_list()] + datos.fillna("").values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,0), (-1,0), colors.gray),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elementos.append(table)

    total_pedidos = len(data) - 1
    estiloTotal = ParagraphStyle('total', parent=estilos['Normal'], fontSize=32, alignment=1, spaceBefore=20)
    parrafo_total = Paragraph(f"<b>Total de pedidos: {total_pedidos}</b>", estiloTotal)
    elementos.append(parrafo_total)

    pdf.build(elementos)
    logging.info(f"Archivo PDF {nombre_archivo_pdf} generado exitosamente.")

# Función para manejar excepciones no capturadas y guardarlas en un archivo de texto
def log_uncaught_exceptions(ex_cls, ex, tb):
    with open('crash_report.txt', 'w') as f:
        f.write("".join(traceback.format_tb(tb)))
        f.write(f"{ex_cls.__name__}: {ex}")
    logging.error(f"{ex_cls.__name__}: {ex}")
    sys.exit(1)

# Configurar el manejador de excepciones no capturadas
sys.excepthook = log_uncaught_exceptions
def procesar_pedido(id_pedido, cadeteria, es_multiple=None):
    if es_multiple is None:
        es_multiple = "No"  # Asumiendo que el valor predeterminado es "No"
    pedidos_envios_flex.append((id_pedido, cadeteria, es_multiple))
def crear_tab_verificar_estado(notebook, tokens):
    tab_verificar_estado = ctk.CTkFrame(notebook)

    label_verificar = ctk.CTkLabel(tab_verificar_estado, text="Verificar Estado", font=('Roboto', 18))
    label_verificar.pack(pady=10)

    entry_verificar = ctk.CTkEntry(tab_verificar_estado, placeholder_text="Escanea o ingresa un ID")
    entry_verificar.pack(pady=10)

    def on_verificar(event=None):
        id_interno = entry_verificar.get().strip()
        if id_interno.isdigit():
            id_interno = int(id_interno)
            if 10000000 <= id_interno <= 19999999:
                # Buscar en la hoja de Flex
                shipment_id = obtener_shipment_id(str(id_interno), SPREADSHEET_ID_FLEX, 'A:D')
            elif 20000000 <= id_interno <= 29999999:
                # Buscar en la hoja de Colecta
                shipment_id = obtener_shipment_id(str(id_interno), SPREADSHEET_ID_COLECTA, 'A:D')
            else:
                shipment_id = None

            if shipment_id:
                shipment_data = verificar_estado_token(shipment_id, tokens)
                if shipment_data:
                    # Aquí puedes manipular el objeto shipment_data si es necesario
                    pass
            else:
                messagebox.showerror("Error", "No se encontró el shipment ID para el ID interno proporcionado.")
        else:
            messagebox.showerror("Error", "Por favor, ingrese un ID numérico para verificar.")

    boton_verificar = ctk.CTkButton(tab_verificar_estado, text="Verificar", command=on_verificar)
    boton_verificar.pack(pady=10)

    entry_verificar.bind("<Return>", on_verificar)

    notebook.add(tab_verificar_estado, text="Verificar Estado")



# Inicio de la aplicación
if __name__ == '__main__':
    tokens = inicializar_tokens()
    crear_tab_verificar_estado(notebook, tokens)
    app.mainloop()
