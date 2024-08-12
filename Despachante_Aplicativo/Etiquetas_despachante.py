import os
import socket
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import win32print
import subprocess

# Identificadores de las planillas de Google Sheets
SPREADSHEET_ID_1 = '1ObCWgChmjd0FHSw1QVHf68txS6zxADA3v4XBjG72wrY'
SPREADSHEET_ID_2 = '15JFcSTheCBdDa4hIEzgduHqHJ5vkmyB4-_J12x4lIBo'

# Credenciales de la cuenta de servicio
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

envios_flex_mensajeria = []
envios_colecta = []
contador_flex_mensajeria = 0
contador_colecta = 0

def login():
    global stored_username  # Define una variable global para almacenar el nombre de usuario
    username = entry_username.get()
    if username:  # Asegura que el nombre de usuario no esté vacío
        stored_username = username
        window.destroy()  # Cierra la ventana de inicio de sesión
        open_despachante_interface()  # Abre la interfaz del despachante
    else:
        messagebox.showerror("Error", "Por favor, ingrese un nombre de usuario")

def obtener_prioridades():
    # Esta función consultará la hoja de cálculo para obtener las prioridades basadas en la fecha
    fecha_consulta = fecha_entry.get()
    if not fecha_consulta:
        messagebox.showerror("Error", "Por favor, ingrese una fecha")
        return

    try:
        # Obtener datos de la hoja "presentacion y cierre"
        range_name = 'presentacion y cierre!A2:V'
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_1, range=range_name).execute()
        valores_presentacion = result.get('values', [])

        # Obtener datos de la hoja "DESPACHO"
        range_name_despacho = 'DESPACHO!A2:M'
        result_despacho = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID_1, range=range_name_despacho).execute()
        valores_despacho = result_despacho.get('values', [])
        despacho_dict = {row[0]: row[12] for row in valores_despacho if len(row) > 12}

        prioridades = []

        for row in valores_presentacion:
            if len(row) > 21 and row[1] == fecha_consulta and row[21]:  # Verifica la fecha en la columna B y el valor en la columna V
                id_pedido = row[0]
                if id_pedido in despacho_dict and despacho_dict[id_pedido] == "PENDIENTE":  # Verifica que la columna M diga "PENDIENTE"
                    prioridades.append((id_pedido, row[21]))  # Guarda el ID y el valor de la columna V

        # Limpiar la tabla
        for item in tabla_prioridades.get_children():
            tabla_prioridades.delete(item)

        # Insertar nuevas prioridades en la tabla
        for prioridad in prioridades:
            tabla_prioridades.insert("", tk.END, values=prioridad)

        if not prioridades:
            messagebox.showinfo("Sin resultados", "No se encontraron prioridades para la fecha ingresada")

    except Exception as e:
        messagebox.showerror("Error", f"Error al obtener prioridades: {e}")



def find_zpl_and_comuna_in_sheet(service, origin_number):
    global SPREADSHEET_ID_1, SPREADSHEET_ID_2

    def find_in_single_sheet(spreadsheet_id, range_start):
        range_name = f'presentacion y cierre!{range_start}:V'  # Ajustar a V para incluir la columna de cadetería
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        for row in values:
            if row and len(row) >= 8 and str(row[0]) == str(origin_number):
                zpl_code = row[7] if len(row) > 7 else None
                comuna = row[20] if len(row) > 20 else None
                cadeteria = row[21] if len(row) > 21 and row[21].strip() != "" else None  # Asumir que cadetería está en columna V, manejar vacíos
                return zpl_code, comuna, cadeteria
        return None, None, None

    origin_number_int = int(origin_number)
    if 10000000 <= origin_number_int <= 19999999:
        zpl_code, comuna, cadeteria = find_in_single_sheet(SPREADSHEET_ID_1, 'A2')
        if zpl_code:
            return zpl_code, comuna, cadeteria, "NEXUS"
    elif 20000000 <= origin_number_int <= 29999999:
        zpl_code, comuna, cadeteria = find_in_single_sheet(SPREADSHEET_ID_2, 'A2')
        if zpl_code:
            return zpl_code, comuna, cadeteria, "NEXUS2"

    return None, None, None, None

def send_data_to_printer(ip_address, port, data):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.connect((ip_address, port))
            sock.sendall(data.encode('utf-8'))
            print("Datos enviados correctamente a la impresora.")
        except Exception as e:
            print(f"Error al enviar datos a la impresora: {e}")

def send_zpl_to_printer(printer_name, zpl_data):
    # Asegúrate de que esta función envíe los datos a la impresora correcta
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        job = win32print.StartDocPrinter(hPrinter, 1, ("ZPL Print Job", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, zpl_data.encode('utf-8'))
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

def procesar_pedido(scanned_value, zpl_code, printer_name):
    global envios_flex_mensajeria, contador_flex_mensajeria, envios_colecta, contador_colecta
    if zpl_code:
        # Añade el pedido a la lista correspondiente y actualiza el contador
        if printer_name == "NEXUS":
            envios_flex_mensajeria.append(scanned_value)
            contador_flex_mensajeria += 1
        elif printer_name == "NEXUS2":
            envios_colecta.append(scanned_value)
            contador_colecta += 1
        
        send_zpl_to_printer(printer_name, zpl_code)  # Envía ZPL a la impresora
        actualizar_contadores()  # Actualiza los contadores en la interfaz
    else:
        print("Código ZPL no proporcionado.")

def obtener_zonas_de_advertencia():
    spreadsheet_id = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'  # ID de tu hoja de cálculo
    range_name = 'EXCEPCION_ZONAS!A:A'  # Rango que contiene las zonas de advertencia

    try:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        zonas_advertencia = [item[0] for item in values if item]  # Crea una lista de zonas, excluyendo celdas vacías
        return zonas_advertencia
    except Exception as e:
        print(f"Error al obtener zonas de advertencia: {e}")
        return []

zonas_advertencia = obtener_zonas_de_advertencia()

def on_scan(event=None):
    global envios_flex_mensajeria, envios_colecta, contador_flex_mensajeria, contador_colecta, scanned_value, zonas_advertencia

    scanned_value = event.widget.get().strip()
    scanned_value = scanned_value.replace(" ", "")
    event.widget.delete(0, 'end')

    def procesar_y_mostrar_alerta():
        global scanned_value, zonas_advertencia
        zpl_code, comuna_o_zona, cadeteria, printer_name = find_zpl_and_comuna_in_sheet(service, scanned_value)
        if zpl_code and printer_name:
            # Verificar la asignación de cadetería
            if cadeteria:
                respuesta = messagebox.askokcancel("Pedido Asignado", f"El pedido {scanned_value} está asignado a la cadetería {cadeteria}. ¿Desea procesarlo?")
                if not respuesta:
                    return  # No procesar si el usuario cancela
            # Manejo de zonas de advertencia
            if comuna_o_zona in zonas_advertencia:
                respuesta = messagebox.askokcancel("Zona Específica Detectada", f"El pedido {scanned_value} corresponde a {comuna_o_zona}. ¿Desea procesarlo?")
                if not respuesta:
                    return  # No procesar si el usuario cancela
            # Manejo estándar por printer
            if printer_name == "NEXUS" or printer_name == "NEXUS2":
                procesar_pedido(scanned_value, zpl_code, printer_name)
        else:
            print("Código ZPL no encontrado o impresora no determinada.")

    if any(scanned_value == envio for envio in envios_flex_mensajeria + envios_colecta):
        global alert_window
        alert_window = tk.Toplevel()
        alert_window.title("Pedido Repetido")

        def eliminar_pedido():
            global envios_flex_mensajeria, envios_colecta, contador_flex_mensajeria, contador_colecta, scanned_value
            if scanned_value in envios_flex_mensajeria:
                envios_flex_mensajeria.remove(scanned_value)
                contador_flex_mensajeria -= 1
            elif scanned_value in envios_colecta:
                envios_colecta.remove(scanned_value)
                contador_colecta -= 1
            actualizar_contadores()
            alert_window.destroy()

        def reemplazar_pedido():
            eliminar_pedido()
            procesar_y_mostrar_alerta()
            alert_window.destroy()

        tk.Label(alert_window, text=f"El pedido {scanned_value} ya fue escaneado.").pack(pady=10)
        tk.Button(alert_window, text="Eliminar", command=eliminar_pedido).pack(side='left', padx=10, pady=10)
        tk.Button(alert_window, text="Reemplazar", command=reemplazar_pedido).pack(side='right', padx=10, pady=10)
    else:
        procesar_y_mostrar_alerta()

def excluir_ids(ids_a_excluir):
    ids = ids_a_excluir.split(',')  # Asumiendo que los IDs están separados por comas
    global envios_flex_mensajeria, envios_colecta, contador_flex_mensajeria, contador_colecta
    
    # Calcula la longitud inicial de las listas
    longitud_inicial_flex = len(envios_flex_mensajeria)
    longitud_inicial_colecta = len(envios_colecta)
    
    # Filtra los envíos de FLEX/MENSAJERÍA
    envios_flex_mensajeria = [envio for envio in envios_flex_mensajeria if envio[0] not in ids]
    
    # Filtra los envíos de COLECTA
    envios_colecta = [envio for envio in envios_colecta if envio[0] not in ids]
    
    # Calcula cuántos IDs se han excluido
    ids_excluidos_flex = longitud_inicial_flex - len(envios_flex_mensajeria)
    ids_excluidos_colecta = longitud_inicial_colecta - len(envios_colecta)
    
    # Actualiza los contadores
    contador_flex_mensajeria -= ids_excluidos_flex
    contador_colecta -= ids_excluidos_colecta
    
    # Asegúrate de que los contadores no sean negativos
    contador_flex_mensajeria = max(contador_flex_mensajeria, 0)
    contador_colecta = max(contador_colecta, 0)
    
    actualizar_contadores()  # Actualiza los contadores en la interfaz

def detener_y_limpiar_cola(impresora):
    try:
        # Detener el servicio de cola de impresión
        subprocess.run(["net", "stop", "spooler"], check=True)

        # Eliminar trabajos de impresión en cola
        subprocess.run(["del", "/Q", "/F", "/S", r"%windir%\System32\spool\PRINTERS\*.*"], check=True, shell=True)

        # Reiniciar el servicio de cola de impresión
        subprocess.run(["net", "start", "spooler"], check=True)

        # Mostrar mensaje de confirmación
        messagebox.showinfo("Reinicio de Cola de Impresión", f"La cola de impresión para {impresora} ha sido reiniciada exitosamente.")
    except subprocess.CalledProcessError as e:
        # Mostrar mensaje de error
        messagebox.showerror("Error", f"Error al reiniciar la cola de impresión para {impresora}: {e}")

def open_despachante_interface():
    global label_contador_flex_mensajeria, label_contador_colecta, lista_comunas, fecha_entry, tabla_prioridades
    despachante_window = ctk.CTk()
    despachante_window.title(f"DESPACHANTE ({stored_username})")

    user_label = ctk.CTkLabel(despachante_window, text=f"Despachante: {stored_username}", font=("default", 10, "bold"))
    user_label.pack(pady=2)

    # Crear un contenedor de pestañas
    tab_control = ttk.Notebook(despachante_window)
    tab_control.pack(expand=1, fill="both")

    # Pestaña de Envíos
    tab_envios = ctk.CTkFrame(tab_control)
    tab_control.add(tab_envios, text="Envíos")

    horarios_despacho = """NESTOR: 10:00 - 10:15HS
GESTIONAR: 10HS - 10:30HS / 14:30-15:00 HS
TDP: 15:00 - 16:00 (Comunas 12,13,14 y 15 no se lleva)
HORARIOS COLECTA: 10:40-11:00HS / 12:00-13:30 HS / 15:00-16:00HS"""

    label_horarios = ctk.CTkLabel(tab_envios, text=horarios_despacho, 
                                font=("default", 15, "bold"), 
                                fg_color=("gray75"), 
                                text_color="white", 
                                width=200, height=100, 
                                corner_radius=10)
    label_horarios.pack(pady=20)

    comunas_frame = ctk.CTkFrame(tab_envios)
    comunas_frame.pack(pady=(10, 0))

    label_comunas = ctk.CTkLabel(comunas_frame, text="Seleccionar Comunas:")
    label_comunas.pack(side="left")

    # Crear un frame para la Listbox y la Scrollbar
    listbox_frame = tk.Frame(comunas_frame)
    listbox_frame.pack(side="left", padx=5)

    # Configurar la Listbox
    lista_comunas = tk.Listbox(listbox_frame, selectmode='multiple', exportselection=0, height=6)
    for i in range(1, 16):
        lista_comunas.insert(tk.END, f"Comuna {i}")
    lista_comunas.pack(side="left", fill="y")

    # Configurar la Scrollbar
    scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")
    scrollbar.config(command=lista_comunas.yview)
    scrollbar.pack(side="right", fill="y")
    lista_comunas.config(yscrollcommand=scrollbar.set)

    def seleccionar_todas_comunas():
        lista_comunas.select_set(0, tk.END)

    def deseleccionar_todas_comunas():
        lista_comunas.selection_clear(0, tk.END)

    boton_seleccionar_todo = ctk.CTkButton(comunas_frame, text="Seleccionar Todo", command=seleccionar_todas_comunas)
    boton_seleccionar_todo.pack(side="left", padx=5)

    boton_deseleccionar_todo = ctk.CTkButton(comunas_frame, text="Deseleccionar Todo", command=deseleccionar_todas_comunas)
    boton_deseleccionar_todo.pack(side="left", padx=5)

    label_contador_flex_mensajeria = ctk.CTkLabel(tab_envios, text="Envíos FLEX/MENSAJERÍA: 0")
    label_contador_flex_mensajeria.pack(pady=10)

    label_contador_colecta = ctk.CTkLabel(tab_envios, text="Envíos COLECTA: 0")
    label_contador_colecta.pack(pady=10)

    scan_entry = ctk.CTkEntry(tab_envios)
    scan_entry.pack(pady=20)
    scan_entry.focus()
    scan_entry.bind('<Return>', on_scan)

    # Crea un contenedor para los botones
    botones_container = ctk.CTkFrame(tab_envios)
    botones_container.pack(pady=10)

    botones_frame = ctk.CTkFrame(tab_envios)
    botones_frame.pack(pady=20)

    # Botón para reiniciar la cola de la impresora Nexus
    boton_reiniciar_cola_nexus = ctk.CTkButton(botones_frame, text="Reiniciar Cola Nexus",
                                                command=lambda: detener_y_limpiar_cola("Nexus"))
    boton_reiniciar_cola_nexus.grid(row=0, column=0, pady=10, padx=10)

    # Botón para reiniciar la cola de la impresora Nexus2
    boton_reiniciar_cola_nexus2 = ctk.CTkButton(botones_frame, text="Reiniciar Cola Nexus2",
                                                command=lambda: detener_y_limpiar_cola("Nexus2"))
    boton_reiniciar_cola_nexus2.grid(row=0, column=1, pady=10, padx=10)

    # Botón para cerrar sesión
    cerrar_sesion_button = ctk.CTkButton(botones_frame, text="Cerrar Sesión", command=despachante_window.destroy)
    cerrar_sesion_button.grid(row=1, column=0, pady=10, padx=10)

    def terminar_despacho_flex():
        global envios_flex_mensajeria, contador_flex_mensajeria
        envios_flex_mensajeria.clear()  # Limpia la lista de envíos FLEX/MENSAJERÍA
        contador_flex_mensajeria = 0  # Reinicia el contador a 0
        actualizar_contadores()  # Actualiza los contadores en la interfaz

    def terminar_despacho_colecta():
        global envios_colecta, contador_colecta
        envios_colecta.clear()  # Limpia la lista de envíos COLECTA
        contador_colecta = 0  # Reinicia el contador a 0
        actualizar_contadores()  # Actualiza los contadores en la interfaz

    # Luego, modifica la creación de los botones para usar estas nuevas funciones
    terminar_despacho_flex_button = ctk.CTkButton(botones_container, text="TERMINAR DESPACHO FLEX",
                                                command=terminar_despacho_flex,
                                                fg_color="#FFFF99",
                                                text_color="black",
                                                font=("default", 10, "bold"))
    terminar_despacho_flex_button.pack(side='left', padx=5)

    terminar_despacho_colecta_button = ctk.CTkButton(botones_container, text="TERMINAR DESPACHO COLECTA",
                                                    command=terminar_despacho_colecta,
                                                    fg_color="#99FF99",
                                                    text_color="black",
                                                    font=("default", 10, "bold"))
    terminar_despacho_colecta_button.pack(side='left', padx=5)

    # Pestaña de Prioridades
    tab_prioridades = ctk.CTkFrame(tab_control)
    tab_control.add(tab_prioridades, text="Prioridades")

    label_fecha = ctk.CTkLabel(tab_prioridades, text="Fecha (DD/MM/AAAA):")
    label_fecha.pack(pady=(20, 5))

    fecha_entry = ctk.CTkEntry(tab_prioridades)
    fecha_entry.pack(pady=(0, 20))

    obtener_prioridades_button = ctk.CTkButton(tab_prioridades, text="Obtener Prioridades", command=obtener_prioridades)
    obtener_prioridades_button.pack(pady=10)

    # Crear tabla para mostrar prioridades
    tabla_prioridades = ttk.Treeview(tab_prioridades, columns=("ID", "Prioridad"), show="headings")
    tabla_prioridades.heading("ID", text="ID")
    tabla_prioridades.heading("Prioridad", text="Prioridad")
    tabla_prioridades.pack(pady=20, fill="both", expand=True)

    despachante_window.mainloop()

def actualizar_contadores():
    global contador_flex_mensajeria, contador_colecta, label_contador_flex_mensajeria, label_contador_colecta
    label_contador_flex_mensajeria.configure(text=f"Envíos FLEX/MENSAJERÍA: {contador_flex_mensajeria}")
    label_contador_colecta.configure(text=f"Envíos COLECTA: {contador_colecta}")

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    center_x = int(screen_width / 2 - width / 2)
    center_y = int(screen_height / 2 - height / 2)
    window.geometry(f'{width}x{height}+{center_x}+{center_y}')

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

window = ctk.CTk()
window.title("Inicio de Sesión")
center_window(window, 500, 250)

label_username = ctk.CTkLabel(window, text="Nombre de Usuario:")
label_username.pack(pady=(20, 5))

entry_username = ctk.CTkEntry(window)
entry_username.pack(pady=(0, 20))

button_login = ctk.CTkButton(window, text="Iniciar Sesión", command=login)
button_login.pack()

window.mainloop()
