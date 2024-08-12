import os
import re
import socket
import threading
import time
from collections import Counter, defaultdict
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient import errors
import psutil
import webbrowser
from tkinter import ttk, messagebox
import tkinter as tk
from tkinter import scrolledtext
from ttkthemes import ThemedTk
import customtkinter as ctk

# Variable global para el cronómetro
start_time = None
cronometro_corriendo = False

# Ruta al archivo JSON con las credenciales de la cuenta de servicio
SERVICE_ACCOUNT_FILE = {
    "type": "service_account",
    "project_id": "api-prueba-387019",
    "private_key_id": "0476618b60b9898097587cded561d722cacec08b",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCePl1VxFnEUJnL\nY6pL3wc7f57srPsMN4Ny35bir34nyQblJ85xReCYDH4yx9iNfJLVOMPZ+kjpBDln\nvOjl2Yx57U9KVIE9x95XPG1cVixW87NgzahWgl9EEnqi92Ek2FqQ3Zrb8wfHqtOv\noTsyZUrVpC5pcuNlpUziEphmqyYjM5Ts3BiAGtANNxJnu/dQvDjGTisUgP50xcYu\nx8jveS8ZzvIa8OQOnDUy/xaunHONpcCGBYBxwB6CWRwfzs6dwzYgqio5cTXBDHEZ\nDMnWe038LO6gpf45Quza9SBBf5HqDAoGiGyxJOrCALHDrXN7StJJ2eZAMAUmIXJU\nR4/dzM8hAgMBAAECggEAAU7fKyUVH+IRr2R6xSdKEkQXDnpMlfiEgefWwbd1KcOH\n5s+KLZYICgG6blWIFP1VYmVWeyAOR7zU86ZUX4gJmP1s9R6XNelHzX1Oox+EPNDF\naBSzIknwGyoy8NNuLmWKdB2c2Nocs5oWJ2LEEEHVmUgC6TS9q/qulMTWFS0vgsV4\nTnGf2Wxn8Z2i6CH716I1exYEr1RvO7MB2XllXZD4A92h/cPWP4WGuxJSW01FYFAJ\nHOQSIEOCdOiyBvJXn+W1FkeA/2n3vXMZ1JTg/mXrFVr2GsNe1qTS9hKnTrn1OiSg\noaBPydcb3zm7W/Gh2V4rgtjZ87OHPDNmKpu2NIXWFQKBgQDRIxbRxiLfqPFgaEyL\nqpMWJmzZz46Ks5U09idBUr7A8XAKHIC1by17R5vNfLJ04DjkYk11LDR79xEOj8ej\nw73hb16JoW9fpanrLgUnFbEpLn5l9hLQ0XTytaD9HezIBNbVBguBTJX7zVh5CoE+\n/aDSO5ZT82S6nKva0tpVLG1bvQKBgQDBs9R+0XmYuYc//ONSyPyLRYG8XzwjmG/7\nq5NBk1sObYdCD8RoGoSyyPxON2f2IdSeRWs/cpvzn6WliTzirPjvGigsgVdpdhHw\nvqFLcCT/spuIcSw7sXHIGuDks7ENqFGMTWvUhVSXnRHLe6DI3oa8HqGz01wn0KG4\nG4h9F6ilNQKBgCjkIXwWr47nq3pcKktMqMG19aaIzKgWAOCAGKDPGvkCvBNeC1Ru\nTohCMhZ1gTcU+Eq/ZcG5SPHMFwuv1SsSBLhlXCiw5ePfsl+Hef5wYVlZ++DXwkoO\nCj3CgZ28EiXxvEeejPVVEiZPujLAeRDl6mHwwK3dcx6Br6BzhbkNnLxJAoGBAIDo\ny+9GOzVPq1oDU8oBZkjNjiSuAJjz1vrXWZ3Xkk7Mx6pIUf6yP0O4jGkFaJOeOiDG\ntDQ6ct00jwnWhoWYdeayMrBDt8fAwvnD9BX/YTAmTmoEX1quobl5YdfyWk9riDHm\nSG/b98saxf/0m7FynppNjHHfpMH+j9G1ySTu/NZNAoGAR8TiBRywXFXE11gIypIv\nsyh6x8uKe+R44mnKwQuK1+fuesjS2Kq7b2c7ZPKNXOohaV6Q1el2tWDMOB7lS6gz\nb8AXtTMIvsu97i/zrX9nNSK9cXzUm/rUVUANDvFfE/FBn0+4aNkCOsw8kjvbyxKg\nOKR8GQLv+V/sUAQgqSO+n7s=\n-----END PRIVATE KEY-----\n",
    "client_email": "cerrador@api-prueba-387019.iam.gserviceaccount.com",
    "client_id": "107808667037270710654",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/cerrador%40api-prueba-387019.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}


# Scopes necesarios para leer y escribir en Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Carga las credenciales
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# IDs de las hojas de cálculo
SPREADSHEET_ID_1 = '1ObCWgChmjd0FHSw1QVHf68txS6zxADA3v4XBjG72wrY'  # ID FLEX
SPREADSHEET_ID_2 = '1MxqTFSBUL8UA0HeABTMI-6Sk8tH5WgK-HvDfqGimll8'  # ID de la segunda hoja
SPREADSHEET_ID_3 = '1c_A8QSjVRKfUnQX_NtM-8eVt_DKwn-ZtpP_6GD1CAZI'  # ID de la tercera hoja
SPREADSHEET_ID_4 = '15JFcSTheCBdDa4hIEzgduHqHJ5vkmyB4-_J12x4lIBo'  # ID COLECTA 
SPREADSHEET_ID_REGISTRO_CERRADORES = '1RRKrOiq6VuKtNx2MMWkBMBQ5H8KBuBs9MtKN-o1zH5c'  # ID RegistroCerradores

# Variable global para almacenar los conceptos
conceptos = {}
concepto_actual = None
id_concepto_actual = None

def cargar_conceptos():
    global conceptos
    try:
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_REGISTRO_CERRADORES, range='Conceptos!A:B').execute()
        values = result.get('values', [])
        # Store concepts in a dictionary with ID as key and concept as value
        conceptos = {row[0]: row[1] for row in values if len(row) >= 2}
        print("Conceptos cargados exitosamente.")
    except Exception as e:
        print(f"Error al cargar conceptos: {e}")

def obtener_usuarios():
    try:
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_REGISTRO_CERRADORES, range='USERS!A:A').execute()
        values = result.get('values', [])
        usuarios = [row[0] for row in values if row]
        return usuarios
    except Exception as e:
        print(f"Error al obtener la lista de usuarios: {e}")
        return []

# Conexión a la API de Sheets
service = build('sheets', 'v4', credentials=creds)

# Función para actualizar el cronómetro
def actualizar_tiempo():
    global start_time, cronometro_corriendo
    if cronometro_corriendo:
        elapsed_time = int(time.time() - start_time)
        horas, resto = divmod(elapsed_time, 3600)
        minutos, segundos = divmod(resto, 60)
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"
        if cronometro_label.winfo_exists():
            cronometro_label.config(text=tiempo_formateado)
            ventana_detalle.after(1000, actualizar_tiempo)

# Función para iniciar el cronómetro
def iniciar_cronometro():
    global start_time, cronometro_corriendo
    start_time = time.time()
    cronometro_corriendo = True
    actualizar_tiempo()

# Función para detener el cronómetro
def detener_cronometro():
    global cronometro_corriendo
    cronometro_corriendo = False

# Función para obtener el tiempo transcurrido
def obtener_tiempo_transcurrido():
    global start_time
    if start_time:
        elapsed_time = int(time.time() - start_time)
        return elapsed_time
    return 0

ventana_seleccion_concepto = None
ventana_mostrada = False
def actualizar_tiempo_principal():
    global start_time_principal, cronometro_principal_corriendo, ventana_seleccion_concepto, concepto_actual, ventana_mostrada

    if cronometro_principal_corriendo:
        elapsed_time = int(time.time() - start_time_principal)
        horas, resto = divmod(elapsed_time, 3600)
        minutos, segundos = divmod(resto, 60)
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"
        
        if cronometro_principal_label.winfo_exists():
            cronometro_principal_label.configure(text=tiempo_formateado)
            root.after(1000, actualizar_tiempo_principal)

        # Mensaje de depuración
        print(f"Tiempo transcurrido: {elapsed_time} segundos. Concepto actual: {concepto_actual}")

        # Mostrar la ventana de selección de concepto si se llega a 4 minutos y el concepto actual es "ESCANEANDO"
        if elapsed_time >= 240 and concepto_actual == "ESCANEANDO" and not ventana_mostrada:
            print("Condiciones para mostrar ventana de selección de concepto cumplidas.")
            ventana_mostrada = True
            mostrar_ventana_seleccion_concepto()




def iniciar_cronometro_principal():
    global start_time_principal, cronometro_principal_corriendo, ventana_mostrada
    start_time_principal = time.time()
    cronometro_principal_corriendo = True
    ventana_mostrada = False  # Resetear la bandera cada vez que se inicia el cronómetro
    actualizar_tiempo_principal()



def detener_cronometro_principal():
    global cronometro_principal_corriendo
    cronometro_principal_corriendo = False

def obtener_tiempo_transcurrido_principal():
    global start_time_principal
    if start_time_principal:
        elapsed_time = int(time.time() - start_time_principal)
        return elapsed_time
    return 0



def handle_google_api_error(error):
    """ Maneja errores comunes de la API de Google Sheets. """
    try:
        if isinstance(error, errors.HttpError):
            if error.resp.status in [500, 502, 503, 504]:
                print("Error del servidor, por favor intente nuevamente más tarde.")
            else:
                print(f"Error HTTP: {error.resp.status} - {error.resp.reason}")
        else:
            print(f"Error no HTTP: {error}")
    except AttributeError:
        print("Ocurrió un error desconocido con la API de Google.")

def set_high_priority():
    try:
        p = psutil.Process(os.getpid())
        p.nice(psutil.HIGH_PRIORITY_CLASS)
    except Exception as e:
        print(f"No se pudo establecer la prioridad alta: {e}")

set_high_priority()

def extraer_numero(texto):
    # Patrón para capturar cualquier secuencia de 8 dígitos
    patron = r'(\d{8})'
    
    # Busca en el texto utilizando el patrón para 8 dígitos
    match = re.search(patron, texto)
    if match:
        return match.group(1)  # Retorna la coincidencia encontrada

    return None  # Retorna None si no hay coincidencias

def convertir_a_segundos(tiempo):
    horas, minutos, segundos = map(int, tiempo.split(':'))
    return horas * 3600 + minutos * 60 + segundos

def extraer_datos(texto_escaneado):
    numero = None
    cuenta = None
    # Attempt to extract a number with the first pattern
    match_numero = re.search(r'\[id\[Ñ\[(\d+)', texto_escaneado)
    if match_numero:
        numero = match_numero.group(1)

    if not numero:
        match = re.search(r'(\d{8})', texto_escaneado)
        if match:
            numero = match.group(1)

    # Attempt to extract account based on known identifiers
    if "id[Ñ264490310" in texto_escaneado:
        cuenta = "DUAITEK"
    elif "id[Ñ39384567" in texto_escaneado:
        cuenta = "TODOMICRO"

    return numero, cuenta

def buscar_enlace(numero):
    numero_int = int(numero)  # Convierte el número a entero para la comparación

    # Decide en qué hoja de cálculo buscar según el rango del número
    if 10000000 <= numero_int <= 19999999:
        spreadsheet_id = SPREADSHEET_ID_1
        hoja = 'Presentacion y cierre'
    elif 20000000 <= numero_int <= 29999999:
        spreadsheet_id = SPREADSHEET_ID_4
        hoja = 'Presentacion y cierre'
    else:
        return None  # Retorna None si el número no está en ninguno de los rangos esperados

    try:
        # Obtiene el ID de la fila 2 para calcular la diferencia
        primer_id_range = f'{hoja}!A2'
        primer_id_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=primer_id_range).execute()
        primer_id_value = primer_id_result.get('values', [])

        if primer_id_value and primer_id_value[0]:
            primer_id = int(primer_id_value[0][0])
            id_diff = primer_id - numero_int  # Calcula la diferencia entre el primer ID y el ID escaneado
            fila_deseada = 2 + id_diff

            # Asegura que la fila deseada esté dentro de un rango 
            if 2 <= fila_deseada:
                enlace_range = f'{hoja}!F{fila_deseada}'
                enlace_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=enlace_range).execute()
                enlace_value = enlace_result.get('values', [])
                if enlace_value and enlace_value[0]:
                    return enlace_value[0][0]
    except Exception as e:
        print(f"Error al obtener datos de Google Sheets: {e}")

    return None

fila_pedido_global = None

def obtener_detalle_pedido(numero):
    global fila_pedido_global  # Usamos la variable global para almacenar la fila del pedido

    try:
        # Define el rango de búsqueda de una sola vez
        range_busqueda = 'Cerrador 1!A:E'
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_2, range=range_busqueda).execute()
        values = result.get('values', [])

        # Busca en los valores obtenidos para encontrar el número de pedido
        for i, row in enumerate(values):
            if len(row) > 0 and row[0] == numero:
                if len(row) > 4:
                    fila_pedido_global = i + 1  # Guardamos la fila donde se encuentra el pedido
                    return row[4]  # Retorna el detalle del pedido, asumiendo que está en la columna E
        return None
    except Exception as e:
        print(f"Error al obtener detalle del pedido: {e}")
        return None


def mostrar_detalle_sku(skus):
    ventana_detalle_sku = tk.Toplevel()
    ventana_detalle_sku.title("Detalle de Artículos")

    texto_detalle = scrolledtext.ScrolledText(ventana_detalle_sku, wrap='word', width=50, height=20)
    texto_detalle.pack(padx=10, pady=10)

    # Aquí almacenamos los enlaces a las imágenes para usarlos más tarde
    enlaces_imagenes = {}

    for i, sku in enumerate(skus, start=1):
        range_sku = f'REFERENCIAS!A:C'  # Ahora incluimos la columna C en el rango
        result_sku = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_3, range=range_sku).execute()
        valores_sku = result_sku.get('values', [])

        for valor in valores_sku:
            if valor[0] == sku:  # Compara el SKU con el valor en la columna A
                detalle = valor[1] if len(valor) > 1 else "No hay detalles disponibles"
                enlace_imagen = valor[2] if len(valor) > 2 else None  # Obtiene el enlace de la imagen si existe
                enlaces_imagenes[sku] = enlace_imagen  # Almacena el enlace de la imagen
                texto_detalle.insert('end', f"{i}- {sku}\n{detalle}\n\n")
                break

    texto_detalle.config(state='disabled')

    # Función para abrir las imágenes en el navegador
    def abrir_imagenes():
        for enlace in enlaces_imagenes.values():
            if enlace:  # Verifica que el enlace exista antes de intentar abrirlo
                webbrowser.open(enlace)

    # Botón para abrir las imágenes de los SKU
    boton_abrir_imagenes = tk.Button(ventana_detalle_sku, text="Abrir Imágenes de Artículos", command=abrir_imagenes)
    boton_abrir_imagenes.pack(pady=10)

# Marcar que el detalle de los SKU ha sido revisado
detalle_revisado = True

def actualizar_lista_skus(lista_skus, detalle):
    for item in lista_skus.get_children():
        lista_skus.delete(item)  # Limpia la lista antes de agregar nuevos elementos

    skus_requeridos = procesar_detalle(detalle)
    skus_contados = contar_skus_almacenados()

    for sku, info in skus_requeridos.items():
        color = info['color']
        cantidad_requerida = info['unidades']
        cantidad_contada = skus_contados.get(sku, 0)

        texto_color = color if color != "-" else "Sin color asignado"

        estilo_fila = "white.TLabel"  # Estilo por defecto en caso de discrepancia

        if cantidad_contada == cantidad_requerida:
            estilo_fila = "Green.TLabel"  # Aplica el estilo verde solo si las cantidades concuerdan exactamente

        lista_skus.insert('', 'end', values=(sku, texto_color, cantidad_contada if color != "-" else ""), tags=(estilo_fila,))

    lista_skus.tag_configure("Green.TLabel", background="green")
    lista_skus.tag_configure("white.TLabel", background="white")

def abrir_enlace_pedido():
    global fila_pedido_global, ultimo_numero

    if fila_pedido_global is not None:
        # Construye el rango para obtener el enlace directamente de la fila conocida
        range_b = f'Cerrador 1!B{fila_pedido_global}'
        result_b = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_2, range=range_b).execute()
        values_b = result_b.get('values', [])

        # Comprueba si existe un enlace en la posición esperada y ábrelo
        if values_b and values_b[0]:
            enlace = values_b[0][0]
            webbrowser.open(enlace)  # Abre el enlace en el navegador web predeterminado
            return
        else:
            messagebox.showerror("Error", "Enlace no encontrado para el número de pedido.")
    else:
        messagebox.showerror("Error", "No se ha identificado la fila del pedido actual.")

def preconfigurar_estilos():
    style = ttk.Style()
    style.configure("mystyle.Treeview", font=('Arial', 14))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 14, 'bold'))

# Función para abrir la interfaz del detalle del pedido
def abrir_interfaz_detalle_pedido(detalle):
    global ultimo_numero, skus_almacenados, cronometro_label, ventana_detalle, menu_conceptos

    print(f"Abriendo detalle del pedido para: {ultimo_numero}")
    texto_titulo = f"Detalle del Pedido - {ultimo_numero}" if ultimo_numero is not None else "Detalle del Pedido"

    ventana_detalle = ctk.CTkToplevel()  # Asegúrate que ctk soporta CTkToplevel
    ventana_detalle.title(texto_titulo)


    # Obtener dimensiones de la pantalla del usuario
    pantalla_ancho = ventana_detalle.winfo_screenwidth()
    pantalla_alto = ventana_detalle.winfo_screenheight()

    # Tamaño de la ventana
    ventana_ancho = 1000
    ventana_alto = 710

    # Calcular posición x e y para centrar la ventana
    x_pos = int((pantalla_ancho / 2) - (ventana_ancho / 2))
    y_pos = int((pantalla_alto / 2) - (ventana_alto / 2)) -50

    ventana_detalle.geometry(f"{ventana_ancho}x{ventana_alto}+{x_pos}+{y_pos}")

    titulo_label = tk.Label(ventana_detalle, text=texto_titulo, font=("Arial", 24))
    titulo_label.pack(pady=10)

    boton_abrir_enlace = tk.Button(ventana_detalle, text="Abrir Enlace del Pedido", font=("Arial", 14), command=abrir_enlace_pedido)
    boton_abrir_enlace.pack(pady=5)

    cronometro_label = tk.Label(ventana_detalle, text="00:00:00", font=("Arial", 18))
    cronometro_label.pack(pady=5)

    iniciar_cronometro()

    style = ttk.Style()
    style.configure("Custom.Treeview.Heading", font=("Arial", 16))
    style.configure("Custom.Treeview", font=("Arial", 14))

    columnas = ('SKU', 'COLOR', 'CANTIDAD ESCAN')
    lista_skus = ttk.Treeview(ventana_detalle, columns=columnas, show='headings', style="Custom.Treeview")
    for col in columnas:
        lista_skus.heading(col, text=col)
        lista_skus.column(col, width=136, anchor='center')
    
    lista_skus.pack(expand=True, fill='both', padx=5, pady=5)




    # Llenar la lista de SKUs de manera eficiente
    actualizar_lista_skus(lista_skus, detalle)

    sku_o_id_input = tk.Entry(ventana_detalle, width=30, font=("Arial", 14))
    sku_o_id_input.pack(pady=5)
    
    # Usar after para asegurar que el foco se establezca después de que todos los widgets estén cargados
    ventana_detalle.after(100, lambda: sku_o_id_input.focus_set())

    almacenar_button = tk.Button(ventana_detalle, text="Almacenar SKU / ID", font=("Arial", 14), command=lambda: almacenar_sku_o_id(sku_o_id_input, lista_skus, detalle))
    almacenar_button.pack(pady=5)
    sku_o_id_input.bind('<Return>', lambda event: almacenar_sku_o_id(sku_o_id_input, lista_skus, detalle))




    def copiar_sku(event):
        seleccionado = lista_skus.selection()
        if seleccionado:
            item = lista_skus.item(seleccionado)
            sku = item['values'][0]
            ventana_detalle.clipboard_clear()
            ventana_detalle.clipboard_append(sku)
            messagebox.showinfo("Información", f"SKU {sku} copiado al portapapeles.")

    lista_skus.bind('<Double-1>', copiar_sku)

    editar_input = tk.Entry(ventana_detalle, width=30, font=("Arial", 14))
    editar_input.pack(pady=5)
    
    editar_button = tk.Button(ventana_detalle, text="Editar Cantidad", font=("Arial", 14), command=lambda: editar_cantidad(editar_input, lista_skus, detalle))
    editar_button.pack(pady=5)

    def obtener_skus_seleccionados(lista_skus):
        seleccionados = lista_skus.selection()
        if seleccionados:
            return [lista_skus.item(sku)['values'][0] for sku in seleccionados]
        else:
            return [lista_skus.item(sku)['values'][0] for sku in lista_skus.get_children()]

    # Crear un Frame para los botones
    button_frame = tk.Frame(ventana_detalle)
    button_frame.pack(pady=20)

    boton_detalle_sku = tk.Button(button_frame, text="DETALLE DE ARTÍCULOS", font=("Arial", 18), command=lambda: mostrar_detalle_sku(obtener_skus_seleccionados(lista_skus)))  # Aumentar tamaño de fuente
    boton_detalle_sku.pack(side=tk.LEFT, padx=5)

    global boton_cerrar_pedido
    boton_cerrar_pedido = tk.Button(button_frame, text="EMPAQUETAR", font=("Arial", 18), command=lambda: cerrar_pedido(ventana_detalle), state=tk.DISABLED)  # Aumentar tamaño de fuente
    boton_cerrar_pedido.pack(side=tk.LEFT, padx=5)

    verificar_button = tk.Button(button_frame, text="Verificar Pedido", font=("Arial", 18), command=lambda: verificar_pedido(lista_skus, detalle, ventana_detalle, boton_cerrar_pedido))  # Aumentar tamaño de fuente
    verificar_button.pack(side=tk.LEFT, padx=5)


    # Reset the dropdown menu
    menu_conceptos.set('')

# Make sure this global variable is properly defined and used where appropriate
menu_conceptos = None
skus_almacenados = []

def editar_cantidad(editar_input, lista_skus, detalle):
    global skus_almacenados
    entrada = editar_input.get()
    if ',' in entrada:
        sku, cantidad = entrada.split(',', 1)
        cantidad = int(cantidad)
        skus_almacenados = [s for s in skus_almacenados if s != sku]
        skus_almacenados.extend([sku] * cantidad)
        actualizar_lista_skus(lista_skus, detalle)
        editar_input.delete(0, tk.END)

# Variable global para almacenar los datos de IDs y SKUs
ids_y_skus = {}

def cargar_ids_y_skus():
    global ids_y_skus
    range_id = 'REFERENCIAS!D:E'
    result_id = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_3, range=range_id).execute()
    valores_id = result_id.get('values', [])
    
    # Filtra y convierte solo los valores que son dígitos
    ids_y_skus = {int(valor[0]): valor[1] for valor in valores_id if len(valor) >= 2 and valor[0].isdigit()}

def almacenar_sku_o_id(input_field, lista_skus, detalle):
    global skus_almacenados, ids_y_skus
    texto_entrada = input_field.get().strip()

    # Intentar manejar la entrada directamente como un SKU
    skus_requeridos = procesar_detalle(detalle)
    if texto_entrada in skus_requeridos:
        sku = texto_entrada
    else:
        # Intentar interpretar la entrada como un ID
        id_escaneado = texto_entrada[-6:-1]
        try:
            id_escaneado_int = int(id_escaneado)  # Convertir a entero para eliminar ceros a la izquierda
            sku = ids_y_skus.get(id_escaneado_int)
            if not sku:
                raise ValueError("ID no encontrado")
        except ValueError:
            messagebox.showerror("Error", "Formato de ID incorrecto o ID no encontrado.")
            input_field.delete(0, tk.END)  # Limpiar el campo de entrada
            return

    # Proceder con el SKU encontrado o ingresado directamente
    if sku in skus_requeridos:
        skus_almacenados.append(sku)
        actualizar_lista_skus(lista_skus, detalle)
    else:
        messagebox.showerror("Error", "ARTICULO INCORRECTO")

    input_field.delete(0, tk.END)  # Limpiar el campo de entrada después de procesar
    verificar_estado_pedido(lista_skus, detalle, boton_cerrar_pedido, False)
    
    input_field.focus_set()
# Llamar a cargar_ids_y_skus al inicio o cuando se cargue la ventana o la hoja
cargar_ids_y_skus()

def mostrar_detalle_faltantes_excedentes(faltantes, excedentes):
    mensaje = "Detalles de la verificación:\n\n"
    if faltantes:
        mensaje += "Faltantes:\n"
        for sku, info in faltantes.items():
            mensaje += f"- SKU {sku} (Color: {info['color']}): Faltan {info['faltantes']} unidades\n"
    if excedentes:
        mensaje += "\nExcedentes:\n"
        for sku, info in excedentes.items():
            mensaje += f"- SKU {sku} (Color: {info['color']}): {info['excedentes']} unidades de más\n"
    
    # Mostrar un solo mensaje con todos los detalles
    messagebox.showinfo("Resultado de la Verificación", mensaje)

def verificar_estado_pedido(lista_skus, detalle, boton_cerrar_pedido, es_verificacion_manual):
    faltantes, excedentes = skus_faltantes(detalle)
    
    # Solo habilitar el botón de cerrar pedido si no hay faltantes ni excedentes
    if not faltantes and not excedentes:
        boton_cerrar_pedido['state'] = tk.NORMAL
    else:
        boton_cerrar_pedido['state'] = tk.DISABLED

    # Si es una verificación manual, mostrar detalles sobre los faltantes o excedentes
    if es_verificacion_manual:
        mostrar_detalle_faltantes_excedentes(faltantes, excedentes)

def actualizar_lista_skus(lista_skus, detalle):
    skus_requeridos = procesar_detalle(detalle)
    skus_contados = contar_skus_almacenados()

    # Preparar todos los elementos nuevos
    items_para_insertar = [
        (
            sku,
            "Sin color asignado" if info['color'] == "-" else info['color'],
            "" if info['color'] == "-" else str(skus_contados.get(sku, 0)),
            "Green.TLabel" if skus_contados.get(sku, 0) == info['unidades'] else "white.TLabel"
        )
        for sku, info in skus_requeridos.items()
    ]

    # Redefinir elementos en la UI en un solo paso
    lista_skus.delete(*lista_skus.get_children())
    for item in items_para_insertar:
        lista_skus.insert('', 'end', values=item[:3], tags=(item[3],))

    # Configurar estilos una sola vez
    lista_skus.tag_configure("Green.TLabel", background="green")
    lista_skus.tag_configure("white.TLabel", background="white")

def contar_skus_almacenados():
    skus_contados = {}
    for sku in skus_almacenados:
        skus_contados[sku] = skus_contados.get(sku, 0) + 1
    return skus_contados

def verificar_pedido(lista_skus, detalle, ventana_detalle, boton_cerrar_pedido):
    faltantes, excedentes = skus_faltantes(detalle)

    # Si no hay faltantes ni excedentes, el pedido está completo
    if not faltantes and not excedentes:
        messagebox.showinfo("Confirmación", "PEDIDO COMPLETO!")
        boton_cerrar_pedido['state'] = tk.NORMAL
    else:
        # Si hay faltantes o excedentes, muestra los detalles utilizando la función existente
        mostrar_detalle_faltantes_excedentes(faltantes, excedentes)

def concuerdan_skus(detalle):
    skus_requeridos = procesar_detalle(detalle)
    skus_contados = contar_skus_almacenados()

    for sku, cantidad_requerida in skus_requeridos.items():
        if skus_contados.get(sku, 0) != cantidad_requerida:
            return False

    return True

def procesar_detalle(detalle):
    detalle_procesado = defaultdict(lambda: {'unidades': 0, 'colores': Counter()})
    pattern = re.compile(r"(\S+), UNIDADES= (\d+), COLOR= ([^\n]+)")
    
    for match in pattern.finditer(detalle):
        sku_info, unidades_str, color = match.groups()
        unidades = int(unidades_str)
        detalle_procesado[sku_info]['unidades'] += unidades
        if color != "-":
            detalle_procesado[sku_info]['colores'][color] += unidades
    
    # Formatear la información de color para cada SKU
    for sku, info in detalle_procesado.items():
        colores_detalle = ', '.join(f"{color} ({unidades}U)" for color, unidades in info['colores'].items())
        info['color'] = colores_detalle if colores_detalle else "Sin color asignado"

    return dict(detalle_procesado)


def es_pedido_repetido(numero):
    # Verificar si el número ya está en la hoja de cálculo
    range_a = 'Cerrador 1!A:A'
    result_a = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_2, range=range_a).execute()
    values = result_a.get('values', [])
    
    # Compara el número con cada valor en la columna A
    for value in values:
        if value and value[0].strip() == numero.strip():  # .strip() elimina los espacios en blanco al inicio y al final
            return True  # Encuentra una coincidencia exacta

    return False  # No encuentra una coincidencia

def on_revisar():
    global pedido_pendiente, ultimo_numero, skus_almacenados, id_concepto_actual, concepto_actual

    texto_escaneado = input_field.get()
    actualizar_label_pedido_pendiente()
    numero, cuenta = extraer_datos(texto_escaneado)

    # Verificar la longitud del número extraído
    if len(numero) != 8:
        messagebox.showerror("Error", "El ID del pedido debe contener exactamente 8 dígitos.")
        return

    if es_pedido_repetido(numero):
        messagebox.showwarning("Advertencia", "PEDIDO REPETIDO!! AVISAR AL ETIQUETADOR")
        return  # Salir si el pedido es repetido

    if numero:
        # Detener y registrar el tiempo del concepto anterior
        tiempo_transcurrido_principal = obtener_tiempo_transcurrido_principal()
        detener_cronometro_principal()
        enviar_tiempo_principal(tiempo_transcurrido_principal, ultimo_numero if ultimo_numero else "None")  # Enviar el tiempo con el concepto anterior

        # Cambiar al concepto "Tiempo_entre_paquetes"
        concepto_actual = "TIEMPO_ENTRE_PAQUETES"
        id_concepto_actual = 1  # ID de concepto para Tiempo_entre_paquetes
        iniciar_cronometro_principal()

        if not es_pedido_repetido(numero):
            ultimo_numero = numero  # Actualizar ultimo_numero solo si el número no está repetido en la hoja
            enlace = buscar_enlace(numero)
            if enlace:
                enviar_datos_iniciales(numero, enlace, cuenta)
                detalle_pedido = obtener_detalle_pedido(numero)
                if detalle_pedido:
                    skus_almacenados.clear()

                    # Cambiar al concepto "REVISANDO_PEDIDO"
                    concepto_actual = "REVISANDO_PEDIDO"
                    id_concepto_actual = 2  # ID de concepto para revisando pedido
                    iniciar_cronometro_principal()

                    abrir_interfaz_detalle_pedido(detalle_pedido)
                else:
                    messagebox.showerror("Error", "No se encontró detalle para el pedido.")
                input_field.delete(0, tk.END)
            else:
                messagebox.showerror("Error", "Enlace no encontrado para el número proporcionado.")
        else:
            messagebox.showwarning("Advertencia", "PEDIDO REPETIDO!! AVISAR AL ETIQUETADOR")
    else:
        messagebox.showerror("Error", "Formato de texto escaneado no válido.")




def enviar_tiempo_principal(tiempo_en_segundos, paquete_id=None):
    global id_concepto_actual, concepto_actual, ultimo_numero
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tiempo_en_dias = tiempo_en_segundos / (24 * 3600)  # Conversión a fracción de días

    # Y asegurar que para los conceptos 2 y 3 siempre se utilice 'ultimo_numero'
    if id_concepto_actual in [2, 3] or paquete_id is None:
        paquete_id = ultimo_numero  # Usa 'ultimo_numero' para los conceptos 2 y 3 o si no se proporciona un ID

    try:
        paquete_id = int(paquete_id)  # Intenta convertir a entero
    except ValueError:
        try:
            paquete_id = float(paquete_id)  # Intenta convertir a flotante
        except ValueError:
            pass  # Mantén el ID del paquete como texto si no se puede convertir

    nueva_fila = [
        [id_concepto_actual, paquete_id, nombre_usuario, fecha_actual, tiempo_en_dias]
    ]

    body = {
        'values': nueva_fila
    }

    try:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID_REGISTRO_CERRADORES,
            range='RegistroCerradores!A:E',
            valueInputOption='RAW',
            body=body
        ).execute()
    except HttpError as error:
        handle_google_api_error(error)



def reabrir_ultimo_pedido():
    global ultimo_numero, pedido_pendiente
    actualizar_label_pedido_pendiente()
    if ultimo_numero:
        detalle_pedido = obtener_detalle_pedido(ultimo_numero)
        if detalle_pedido:
            # Si se encuentra el detalle, abrir el detalle del pedido
            abrir_interfaz_detalle_pedido(detalle_pedido)
        else:
            # Intentar obtener el enlace directamente con buscar_enlace
            enlace = buscar_enlace(ultimo_numero)
            if enlace:
                # Si se encuentra el enlace, usarlo para reenviar el pedido
                enviar_datos_iniciales(ultimo_numero, enlace, "Cuenta no disponible")
                abrir_interfaz_detalle_pedido(obtener_detalle_pedido(ultimo_numero))
            else:
                # Si no se encuentra el enlace, proceder como antes
                enviar_datos_iniciales(ultimo_numero, "Enlace no disponible", "Cuenta no disponible")
                abrir_interfaz_detalle_pedido(obtener_detalle_pedido(ultimo_numero))
    else:
        messagebox.showinfo("Información", "No hay pedido pendiente de cierre.")

def enviar_datos_iniciales(numero, enlace, cuenta):
    fila_vacia = encontrar_fila_vacia(service.spreadsheets(), SPREADSHEET_ID_2, 'Cerrador 1!A:A')

    # Asegúrate de que 'numero' sea tratado como un número si no inicia con caracteres especiales
    try:
        numero = int(numero)  # Intenta convertir a entero
    except ValueError:
        try:
            numero = float(numero)  # Intenta convertir a flotante
        except ValueError:
            pass  # Mantén el número como texto si no se puede convertir

    body = {
        "valueInputOption": "RAW",
        "data": [
            {"range": f'Cerrador 1!A{fila_vacia}', "values": [[numero]]},
            {"range": f'Cerrador 1!B{fila_vacia}', "values": [[enlace]]}
        ]
    }

    if cuenta:
        body["data"].append({"range": f'Cerrador 1!C{fila_vacia}', "values": [[cuenta]]})

    # Envía el número, enlace y cuenta a la hoja de cálculo en una sola llamada en lote
    service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID_2, body=body).execute()

def encontrar_fila_vacia(sheet_service, spreadsheet_id, range_name):
    """ Encuentra la primera fila vacía en una columna dada """
    result = sheet_service.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    # Retorna el número de la fila siguiente a la última con datos
    return len(values) + 1

def escribir_datos_en_sheet2(numero, enlace, fila, nombre_usuario):
    range_a = f'Cerrador 1!A{fila}'
    range_b = f'Cerrador 1!B{fila}'
    range_d = f'Cerrador 1!D{fila}'  # Columna para el nombre del usuario

    body = {
        "valueInputOption": "RAW",
        "data": [
            {"range": range_a, "values": [[numero]]},
            {"range": range_b, "values": [[enlace]]},
            {"range": range_d, "values": [[nombre_usuario]]}
        ]
    }

    service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID_2, body=body).execute()
  # Actualiza la hoja con el nombre del usuario

ultimo_numero = None
ultimo_enlace = None

def on_submit():
    global ultimo_numero, ultimo_enlace
    texto_escaneado = input_field.get()
    numero = extraer_numero(texto_escaneado)

    if not numero:
        messagebox.showerror("Error", "Formato de texto escaneado no válido.")
        return

    if ultimo_numero is not None:
        messagebox.showwarning("Advertencia", "Por favor, cierra el pedido anterior antes de escanear uno nuevo.")
        return

    # Combinar la verificación del número y la obtención del enlace en una sola operación
    enlace, numero_repetido = buscar_enlace_y_verificar_repetido(numero)

    if numero_repetido:
        messagebox.showwarning("Advertencia", "PEDIDO REPETIDO!! AVISAR AL ETIQUETADOR")
        ultimo_numero = None  # Resetear el último número ya que el pedido es repetido
        ultimo_enlace = None
    elif enlace:
        webbrowser.open(enlace)
        ultimo_numero = numero
        ultimo_enlace = enlace
        input_field.delete(0, tk.END)
    else:
        messagebox.showerror("Error", "Número no encontrado en la hoja de cálculo.")

def buscar_enlace_y_verificar_repetido(numero):
    # Define el rango para incluir la columna de enlaces además de los números de pedido
    range_a = 'Cerrador 1!A:B'
    result_a = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_2, range=range_a).execute()
    values_a = result_a.get('values', [])
    
    # Itera sobre las filas devueltas, verificando que no estén vacías
    for value in values_a:
        if value and len(value) > 1 and value[0] == numero:  # Verifica que la fila tenga al menos dos elementos y el primero sea el número buscado
            return value[1], True  # Devuelve el enlace y True si el número es repetido
    
    return None, False  # Devuelve None y False si no se encuentra el número

# Inicializa el contador de paquetes cerrados
paquetes_cerrados = 0

# Crear la ventana principal
root = ThemedTk(theme="Adapta")
root.title("CERRADOR 1")

# Configuración inicial de la ventana principal (se oculta hasta que el usuario inicie sesión)
root.withdraw()

# Configuración inicial
root = ThemedTk(theme="Adapta")
root.withdraw()  # Ocultar inicialmente la ventana principal hasta que el usuario inicie sesión

# Variable global para almacenar el widget Label del título
label_titulo = None

# Variable global para el nombre del usuario
nombre_usuario = ""  # Inicialmente vacío hasta que el usuario inicie sesión

def mostrar_ventana_seleccion_concepto_inicial():
    global conceptos  # Declara conceptos como global para acceder a la variable fuera de esta función
    ventana_seleccion_concepto = ctk.CTkToplevel()
    ventana_seleccion_concepto.title("Seleccione el Concepto Inicial")
    ventana_seleccion_concepto.geometry("400x200")

    label = ctk.CTkLabel(ventana_seleccion_concepto, text="Seleccione el concepto inicial:", font=("Arial", 14))
    label.pack(pady=10)

    # Ahora la variable conceptos es reconocida correctamente como la global
    conceptos_disponibles = list(conceptos.values())
    concepto_seleccionado = tk.StringVar()
    menu_conceptos = ctk.CTkOptionMenu(ventana_seleccion_concepto, variable=concepto_seleccionado, values=conceptos_disponibles)
    menu_conceptos.pack(pady=10)

    def on_concepto_seleccionado():
        global concepto_actual, id_concepto_actual
        concepto_actual = concepto_seleccionado.get()
        id_concepto_actual = next(key for key, value in conceptos.items() if value == concepto_actual)
        root.deiconify()  # Mostrar la ventana principal
        iniciar_cronometro_principal()  # Iniciar el cronómetro principal con el concepto seleccionado
        ventana_seleccion_concepto.destroy()

    boton_seleccionar = ctk.CTkButton(ventana_seleccion_concepto, text="Seleccionar", command=on_concepto_seleccionado)
    boton_seleccionar.pack(pady=10)

    ventana_seleccion_concepto.mainloop()



def mostrar_ventana_inicio_sesion():
    # Configura el tema de customtkinter
    ctk.set_appearance_mode("Dark")  # "Light" (Predeterminado), "Dark", "System"
    ctk.set_default_color_theme("dark-blue")  # "blue" (Predeterminado), "green", "dark-blue"

    usuarios = obtener_usuarios()
    
    def iniciar_sesion():
        global nombre_usuario
        nombre_usuario = usuario_seleccionado.get()
        if nombre_usuario:
            ventana_inicio_sesion.destroy()
            label_titulo.configure(text=f"{nombre_usuario}")  # Actualizar el texto del título con el nombre del usuario
            mostrar_ventana_seleccion_concepto_inicial()  # Abrir la ventana para seleccionar el concepto inicial
        else:
            ctk.CTkMessageBox.show_error(title="Error", message="Por favor, seleccione su nombre.")

    # Crear la ventana de inicio de sesión
    ventana_inicio_sesion = ctk.CTk()
    ventana_inicio_sesion.title("Inicio de Sesión")

    # Tamaño y posición de la ventana
    ancho_ventana = 500
    alto_ventana = 300
    pantalla_ancho = ventana_inicio_sesion.winfo_screenwidth()  # Ancho de la pantalla
    pantalla_alto = ventana_inicio_sesion.winfo_screenheight()  # Alto de la pantalla
    x_pos = int((pantalla_ancho / 2) - (ancho_ventana / 2))
    y_pos = int((pantalla_alto / 2) - (alto_ventana / 2))
    ventana_inicio_sesion.geometry(f"{ancho_ventana}x{alto_ventana}+{x_pos}+{y_pos}")  # Ajusta el tamaño y la posición

    # Agregar un label
    label = ctk.CTkLabel(ventana_inicio_sesion, text="Seleccione su nombre:", text_color="orange")
    label.pack(pady=20)

    # Agregar un menú desplegable
    usuario_seleccionado = tk.StringVar(ventana_inicio_sesion)
    menu_desplegable = ctk.CTkOptionMenu(ventana_inicio_sesion, variable=usuario_seleccionado, values=usuarios)
    menu_desplegable.pack(pady=10)

    # Agregar un botón de inicio de sesión
    boton_iniciar_sesion = ctk.CTkButton(ventana_inicio_sesion, text="Iniciar Sesión", command=iniciar_sesion, fg_color="orange", hover_color="orange4")
    boton_iniciar_sesion.pack(pady=20)

    ventana_inicio_sesion.mainloop()


pedido_pendiente = False

# Función para cerrar el pedido
def cerrar_pedido(ventana_detalle):
    global ultimo_numero, nombre_usuario, fila_pedido_global, start_time, pedido_pendiente, paquetes_cerrados, id_concepto_actual

    detener_cronometro()
    pedido_pendiente = False
    actualizar_label_pedido_pendiente()

    if ultimo_numero and fila_pedido_global:
        faltantes, excedentes = skus_faltantes(obtener_detalle_pedido(ultimo_numero))
        if not faltantes and not excedentes:
            try:
                range_d = f'Cerrador 1!D{fila_pedido_global}'
                body = {
                    'valueInputOption': 'RAW',
                    'data': [{'range': range_d, 'values': [[nombre_usuario]]}]
                }

                service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID_2, body=body).execute()

                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                tiempo_en_segundos = obtener_tiempo_transcurrido()
                tiempo_en_dias = tiempo_en_segundos / (24 * 3600)  # Conversión a fracción de días

                nueva_fila = [
                    [2, int(ultimo_numero), nombre_usuario, fecha_actual, tiempo_en_dias]
                ]

                body = {
                    'values': nueva_fila
                }

                service.spreadsheets().values().append(
                    spreadsheetId=SPREADSHEET_ID_REGISTRO_CERRADORES,
                    range='RegistroCerradores!A:E',
                    valueInputOption='RAW',
                    body=body
                ).execute()

                # Incrementar el contador de paquetes cerrados
                paquetes_cerrados += 1
                contador_label.configure(text=f"Paquetes Cerrados: {paquetes_cerrados}")

                # Resetear los valores globales para evitar que el pedido quede pendiente
                fila_pedido_global = None
                start_time = None
                id_concepto_actual = None

                # Cerrar la ventana de detalle
                ventana_detalle.destroy()

                # Abrir la interfaz de empaquetado
                abrir_interfaz_empaquetado(ultimo_numero)  # ultimo_numero aún contiene el ID del pedido

                messagebox.showinfo("Confirmación", "Pedido cerrado correctamente")
            except HttpError as error:
                handle_google_api_error(error)
            except Exception as e:
                print(f"Ocurrió un error inesperado: {e}")
        else:
            mensaje = "No se puede cerrar el pedido debido a discrepancias.\n"
            if faltantes:
                mensaje += "Faltan SKUs:\n"
                for sku, info in faltantes.items():
                    mensaje += f"- {sku}: Faltan {info['faltantes']} unidades\n"
            if excedentes:
                mensaje += "SKUs con exceso:\n"
                for sku, info in excedentes.items():
                    mensaje += f"- {sku}: {info['excedentes']} unidades de más\n"
            messagebox.showerror("Error", mensaje)
    else:
        messagebox.showwarning("Advertencia", "No hay un pedido para cerrar.")

ventana_empaquetado = None
def abrir_interfaz_empaquetado(id_pedido):
    global cronometro_principal_corriendo, start_time_principal, concepto_actual, id_concepto_actual, ventana_empaquetado

    if ventana_empaquetado and ventana_empaquetado.winfo_exists():
        ventana_empaquetado.lift()
        return

    # Detener el cronómetro principal actual antes de cambiar el concepto
    detener_cronometro_principal()

    # Configurar el nuevo concepto como EMPAQUETANDO_PEDIDO
    concepto_actual = "EMPAQUETANDO_PEDIDO"
    id_concepto_actual = 3  # ID de concepto para empaquetado
    iniciar_cronometro_principal()

    # Deshabilitar el campo de entrada de escaneo de paquetes
    input_field.configure(state="disabled")

    # Crear la ventana de empaquetado
    ventana_empaquetado = ctk.CTkToplevel()
    ventana_empaquetado.title(f"Empaquetado del Pedido {id_pedido}")

    # Configurar el tamaño y la posición de la ventana
    ventana_empaquetado.geometry("600x400")

    # Mantener la ventana de empaquetado en primer plano
    ventana_empaquetado.attributes("-topmost", True)

    # Título con el ID del pedido
    label_titulo = ctk.CTkLabel(ventana_empaquetado, text=f"Empaquetado del Pedido {id_pedido}", font=("Arial", 24))
    label_titulo.pack(pady=20)

    # Campo de entrada para escanear el ID del pedido
    input_empaquetado = ctk.CTkEntry(ventana_empaquetado, width=400, font=("Arial", 20))
    input_empaquetado.pack(pady=20)

    def finalizar_empaquetado():
        global concepto_actual, id_concepto_actual, ultimo_numero, ventana_empaquetado

        # Detener el cronómetro de empaquetado
        detener_cronometro_principal()
        tiempo_empaquetado = obtener_tiempo_transcurrido_principal()

        # Registrar el tiempo de empaquetado
        enviar_tiempo_principal(tiempo_empaquetado, ultimo_numero)

        # Reiniciar el cronómetro con el concepto de Tiempo_entre_paquetes
        concepto_actual = "TIEMPO_ENTRE_PAQUETES"
        id_concepto_actual = 1# ID de concepto para Tiempo_entre_paquetes
        iniciar_cronometro_principal()

        # Habilitar el campo de entrada de escaneo de paquetes
        input_field.configure(state="normal")

        # Resetear el ultimo_numero para evitar conflictos futuros
        ultimo_numero = None

        # Cerrar la ventana de empaquetado
        ventana_empaquetado.destroy()
        ventana_empaquetado = None

    def validar_id_escaneado(event):
        id_escaneado = input_empaquetado.get()
        if id_escaneado == str(id_pedido):
            finalizar_empaquetado()
        else:
            messagebox.showerror("Error", "El ID escaneado no coincide con el ID del pedido.")

    # Bind the validation function to the input field
    input_empaquetado.bind('<Return>', validar_id_escaneado)

    # Capturar el evento de cierre de la ventana
    def on_close():
        messagebox.showwarning("Advertencia", "Debe finalizar el empaquetado para cerrar esta ventana.")
        ventana_empaquetado.lift()

    ventana_empaquetado.protocol("WM_DELETE_WINDOW", on_close)

    ventana_empaquetado.mainloop()


def skus_faltantes(detalle):
    skus_requeridos = procesar_detalle(detalle)
    skus_contados = contar_skus_almacenados()
    faltantes = {}
    excedentes = {}

    for sku, info in skus_requeridos.items():
        cantidad_requerida = info['unidades']  # Asume que 'unidades' es la clave para la cantidad requerida
        cantidad_contada = skus_contados.get(sku, 0)

        if cantidad_contada < cantidad_requerida:
            faltantes[sku] = {'faltantes': cantidad_requerida - cantidad_contada, 'color': info['color']}
        elif cantidad_contada > cantidad_requerida:
            excedentes[sku] = {'excedentes': cantidad_contada - cantidad_requerida, 'color': info['color']}

    return faltantes, excedentes  # Devuelve tanto los faltantes como los excedentes

def centrar_ventana(ventana, ancho, alto):
    # Obtener las dimensiones de la pantalla
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()

    # Calcular la posición x e y para centrar la ventana
    x = (pantalla_ancho / 2) - (ancho / 2)
    y = (pantalla_alto / 2) - (alto / 2)

    ventana.geometry('%dx%d+%d+%d' % (ancho, alto, x, y))

def mostrar_detalle_articulo(detalle):
    ventana_detalle = ctk.CTkToplevel()  # Usa CTkToplevel para la ventana de detalles
    ventana_detalle.title("Detalle del Artículo")
    centrar_ventana(ventana_detalle, 400, 200)

    # Usa CTkLabel para el texto del detalle
    label_detalle = ctk.CTkLabel(ventana_detalle, text=detalle, font=("Arial", 22), text_color="white")
    label_detalle.pack(pady=10, padx=10)

root = ctk.CTk()  # Inicializa la ventana principal con CTk
root.title("CERRADOR 1")
root.geometry("2560x1400") 

label_titulo = ctk.CTkLabel(root, text="CERRADOR 1", font=("Arial", 43), text_color="white", anchor="center")
label_titulo.pack(pady=20)


label_pedido_pendiente = ctk.CTkLabel(root, text="No hay pedidos pendientes", font=("Arial", 14), text_color="white")
label_pedido_pendiente.pack(pady=10)

# Function to update the pending order label
def actualizar_label_pedido_pendiente():
    global ultimo_numero, pedido_pendiente
    if pedido_pendiente and ultimo_numero:
        label_pedido_pendiente.configure(text=f"Tienes pendiente el pedido {ultimo_numero}")
    else:
        label_pedido_pendiente.configure(text="No hay pedidos pendientes")

# Label to show the number of closed packages
contador_label = ctk.CTkLabel(root, text="Paquetes Cerrados: 0", font=("Arial", 36), text_color="white", anchor="center")
contador_label.pack(pady=10)

# Variable of control linked to the input field
entrada_var = tk.StringVar()

# Maintain only this input_field which uses the control variable entrada_var
input_field = ctk.CTkEntry(root, textvariable=entrada_var, width=600, corner_radius=20, font=("Arial", 20), justify="center")
input_field.pack(pady=10)
input_field.configure(state="normal")  # Always keep it enabled

# Function that is called when there is a change in the input field
def on_cambio(*args):
    texto = entrada_var.get()
    if texto:  # If there is text, assume something was scanned or entered
        on_revisar()

# Bind the on_cambio function to the control variable
entrada_var.trace("w", on_cambio)

# Bind the <Return> event to the input field to execute on_revisar when Enter is pressed
input_field.bind('<Return>', lambda event: on_revisar())

def registrar_cierre_sesion():
    global nombre_usuario, id_concepto_actual, concepto_actual
    if nombre_usuario:
        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tiempo_en_segundos = obtener_tiempo_transcurrido_principal()
        tiempo_en_dias = tiempo_en_segundos / (24 * 3600)  # Convertir a fracción de días

        # Registro con el concepto actual
        nueva_fila_concepto_actual = [
            [id_concepto_actual, "None" if not ultimo_numero else ultimo_numero, nombre_usuario, fecha_actual, tiempo_en_dias]
        ]

        # Registro de LOGOUT
        nueva_fila_logout = [
            [19, "None", nombre_usuario, fecha_actual, None]
        ]

        body = {
            'values': nueva_fila_concepto_actual + nueva_fila_logout
        }

        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID_REGISTRO_CERRADORES,
            range='RegistroCerradores!A:E',
            valueInputOption='RAW',
            body=body
        ).execute()


def on_closing():
    registrar_cierre_sesion()
    root.destroy()

# Modify the logout button to call the new function
logout_button = ctk.CTkButton(root, text="Cerrar Sesión", command=on_closing)
logout_button.pack(pady=20)

# Bind the window close event to the new function
root.protocol("WM_DELETE_WINDOW", on_closing)

# Button to reopen the last pending order
boton_reabrir_pedido = ctk.CTkButton(root, text="Reabrir Último Pedido Pendiente", command=reabrir_ultimo_pedido)
boton_reabrir_pedido.pack(pady=20)

def actualizar_concepto_en_interfaz():
    global menu_conceptos, concepto_actual
    if menu_conceptos and concepto_actual:
        menu_conceptos.set(concepto_actual)
        print(f"Interfaz actualizada al concepto: {concepto_actual}")

# Function to create and manage the concept menu
def crear_menu_conceptos():
    global menu_conceptos, concepto_actual
    conceptos_list = list(conceptos.values())  # List of concept names
    menu_conceptos = ttk.Combobox(root, values=conceptos_list, font=("Arial", 20))
    menu_conceptos.pack(pady=10)
    if concepto_actual in conceptos_list:
        menu_conceptos.set(concepto_actual)  # Set the initial value if it's in the list
    else:
        menu_conceptos.set(conceptos_list[0] if conceptos_list else '')  # Set to first item or empty if list is empty
    menu_conceptos.bind("<<ComboboxSelected>>", almacenar_concepto_seleccionado)

def almacenar_concepto_seleccionado(event):
    global concepto_actual, id_concepto_actual, menu_conceptos
    nuevo_concepto = event.widget.get()
    nuevo_id_concepto = None

    # Buscar el ID del concepto seleccionado
    for key, value in conceptos.items():
        if value == nuevo_concepto:
            nuevo_id_concepto = key
            break

    # Si se encontró un nuevo concepto y es diferente al actual
    if nuevo_id_concepto is not None and concepto_actual != nuevo_concepto:
        # Si se cambia de conceptos críticos, manejar el tiempo
        if concepto_actual not in ["ESCANEANDO", "REVISANDO_PEDIDO"]:
            tiempo_transcurrido_principal = obtener_tiempo_transcurrido_principal()
            enviar_tiempo_principal(tiempo_transcurrido_principal, "None")
            iniciar_cronometro_principal()  # Reiniciar el cronómetro principal

        # Actualizar el concepto actual y su ID
        concepto_actual = nuevo_concepto
        id_concepto_actual = nuevo_id_concepto

        # Informar al usuario del cambio de concepto
        messagebox.showinfo("Información", f"Concepto seleccionado: {concepto_actual}")

        # Actualizar dependencias en la interfaz
        actualizar_concepto_en_interfaz()

# Main timer label
cronometro_principal_label = ctk.CTkLabel(root, text="00:00:00", font=("Arial", 20))
cronometro_principal_label.pack(pady=10)

# Function to show the concept selection window
def mostrar_ventana_seleccion_concepto():
    global ventana_seleccion_concepto

    if ventana_seleccion_concepto is not None and ventana_seleccion_concepto.winfo_exists():
        print("La ventana de selección de concepto ya está abierta.")
        return  # No hacer nada si la ventana ya está abierta

    ventana_seleccion_concepto = ctk.CTkToplevel()
    ventana_seleccion_concepto.title("Seleccione un Concepto")
    ventana_seleccion_concepto.geometry("400x200")

    label = ctk.CTkLabel(ventana_seleccion_concepto, text="Seleccione un concepto para justificar el tiempo:", font=("Arial", 14))
    label.pack(pady=10)

    menu_conceptos = ctk.CTkOptionMenu(ventana_seleccion_concepto, values=list(conceptos.values()), font=("Arial", 14))
    menu_conceptos.pack(pady=10)

    def on_concepto_seleccionado():
        global id_concepto_actual, concepto_actual
        nuevo_concepto = menu_conceptos.get()
        for key, value in conceptos.items():
            if value == nuevo_concepto:
                id_concepto_actual = key
                concepto_actual = value
                break

        # Enviar el tiempo acumulado hasta ahora con el concepto anterior
        tiempo_transcurrido_principal = obtener_tiempo_transcurrido_principal()
        enviar_tiempo_principal(tiempo_transcurrido_principal, "None")  # Enviar el tiempo con el concepto anterior y paquete None

        # Reiniciar el cronómetro principal con el nuevo concepto
        iniciar_cronometro_principal()

        ventana_seleccion_concepto.destroy()

    boton_seleccionar = ctk.CTkButton(ventana_seleccion_concepto, text="Seleccionar", command=on_concepto_seleccionado)
    boton_seleccionar.pack(pady=10)

    ventana_seleccion_concepto.mainloop()



# Initialize main timer
iniciar_cronometro_principal()

# Load concepts from the spreadsheet
cargar_conceptos()

# Call the function to create the concept menu
crear_menu_conceptos()

# Show the login window at the start
mostrar_ventana_inicio_sesion()

root.mainloop()
