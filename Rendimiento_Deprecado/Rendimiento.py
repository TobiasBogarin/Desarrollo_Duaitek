import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime
from plyer import notification
import screeninfo
import time

# Configuración de las credenciales de la cuenta de servicio para Google Sheets
SERVICE_ACCOUNT_FILE = {
    "type": "service_account",
    "project_id": "api-prueba-387019",
    "private_key_id": "523a7ab379d3d59382de0b4dd5215f50f7fd1370",
    "private_key": """-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDg6b66CIW2Q8/7
    6bAudowLEjxwHDvoSZluVTTbJx0oWkrjJTvWhrr595mNTl83S88lhIaIFHng8sTL
    Evsr/Zp1i6bAvgpiPJb8dm7g5IsH5ps+LL+5q0nJW6ZKb29W1/D1u3M9WiYLWH4P
    jMehYxsmiHDB3pMRwoQfrJk06iF+xbO5HoWzUMvzuhmUeAbZ48/jbACrnoI+7Lou
    9risvi4AIjyBsxY162cVtyqAXmitXJ9QJAl4EoJoqSyAkl59aLcEi0UWUy0kuSD7
    7Pj3U+z3IyDWsOMld+lpv2cPdtij+/VtApaEYhklU8yYfzb8PG1CxkAFRrx2JP8x
    5IWiZSW3AgMBAAECggEAGfhOLOOdsjqzofONLlrt4hFhp9L6xTWftnJhlLSNf1f9
    sateShprLesocH042AUJjtwglJyYqMrKF7tsroWtTMlVSzLRDCAm5vvd4wXrWnwx
    jMT/VmGwNsSTDPaVCpgLRgng9+1C21Mv1dJ8R+b4/u1ePQ8wOCrCYCM+hYegr9HO
    IF4cWMPxX5jF4y+MlsK+iy1s8lI/cmf5p5C0yd5ialtkycuSiMLwlz11UlcCGhf/
    Tmose81uhl07F/EU4JELX/DGRJLUuLk4tBqiNdYfMztL2KvTsd1FVQ+PjS1Syc3D
    OEDm5P6eeIP6Fq2SPjUAiFTMmBpxRdlMSPq0Rqt6YQKBgQD3R4Pd2deByHviYfCJ
    W7Xu6z5QAmFxXgowhJPsqhqiHm7FqKKl4hqqEEFdRa50Az8stC7J4yEQmL9hzHTQ
    FWvVeSbozy6u0hyBtv4CZrzIFepZHMhW9KsBpm1sQOAliCZQa8vNr18C1nVLtKq8
    wFKcWt4lG1nmH5NyZb4ybHEXiQKBgQDo2E2Fw8SW+kxC1g7iS29OZk5jSj/qW7S6
    +23y31N+9MmSl9QZIDtS1dlU+6mxjeuaZsXbRDqBZzrpOSw163mUOMZqCOdrxs6k
    RrzRJ+kLWVhivA/eKLGZIbhCgDWtbq0JdKL0kOVKQSoO88OTnR6+4s9Judod46Qs
    nyr/BMvDPwKBgH+P1ObNSe8ZjU7rVzqEpQXrNOnxUHM7H+aHfgfIeJTJPjuZEs6g
    JUE1wYJsP+J5Ck31ZW2gTZ5SLeg1oMz3P/mP1hKjTmHA4hPIYqC6fwh4xbvSrUau
    UMk5IZmGnhq+cYVrFme04D6Gg1vah3l3fSZLee2KfoXIJDgPZF5+spiBAoGBALon
    LBs0PzhhFZUdo7qhinRIcIUK+Hx6IsyWdPmGOC+4rmrHfac00JjSJTW/GZS9HM5N
    OgOp0YhhKoUI02KsRoAMv/xH8BSHVe+aKhyhZrxPCs2tApafPBVsEu7/p2pnoGl9
    2UXjjZzG6kQX+JVMOSdtF0IfFtVsiHWwLuTBRdJrAoGAJ0BSCX/hNZxU7AqBi1gb
    qfHhotVoMVSimAyQT5vjHmdmyST8kFb3wn537jvrWc60PiCJ1nuk/IHXMLKjMByo
    iYZ+MQsX2/59XJgx0Bpm75NshmeEkU4qh/4YbDCpYyiFZSXxsPRKwLeloOfX6oKT
    8VW0NaC0SXnsYYqamCTPnSA=\n-----END PRIVATE KEY-----""",
    "client_email": "cerrador-1@api-prueba-387019.iam.gserviceaccount.com",
    "client_id": "117910477101184573911",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/cerrador-1%40api-prueba-387019.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}
pending_order_ids = set()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

def move_to_screen2(window):
    # Intentar mover la ventana a la pantalla 2 hasta que se detecte
    for _ in range(10):  # Intentar 10 veces con un retraso entre intentos
        screens = screeninfo.get_monitors()
        if len(screens) > 1:
            screen2 = screens[1]  # Asumiendo que la pantalla 2 es la segunda en la lista
            window.geometry(f"1024x600+{screen2.x}+{screen2.y}")
            return
        else:
            time.sleep(3)  # Esperar 3 segundos antes de volver a intentar

    print("No se encontró una segunda pantalla.")

def fetch_data():
    SPREADSHEET_ID = '1MxqTFSBUL8UA0HeABTMI-6Sk8tH5WgK-HvDfqGimll8'
    RANGE_NAME = 'CERRADOR 1!A2:G'
    
    # Realizar solicitud batchGet
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])
    
    user_data = {}
    total_closed = 0

    # Procesamiento de los valores obtenidos
    for row in values:
        if len(row) >= 7 and row[3].strip():
            user = row[3].strip()
            close_time = datetime.datetime.strptime(row[6], "%d/%m/%Y %H:%M:%S")
            hour = close_time.hour

            if user not in user_data:
                user_data[user] = {'times': [], 'hourly': {}, 'count': 0}

            user_data[user]['times'].append(close_time)
            user_data[user]['count'] += 1
            total_closed += 1

            if hour in user_data[user]['hourly']:
                user_data[user]['hourly'][hour] += 1
            else:
                user_data[user]['hourly'][hour] = 1

    # Calculo del tiempo promedio
    for user, data in user_data.items():
        if data['times']:
            min_time = min(data['times'])
            max_time = max(data['times'])
            total_duration = (max_time - min_time).total_seconds()

            # Calcular las horas laborales entre las 8:00 y las 18:00
            work_start = datetime.datetime.combine(min_time.date(), datetime.time(8, 0))
            work_end = datetime.datetime.combine(max_time.date(), datetime.time(18, 0))
            actual_start = max(min_time, work_start)
            actual_end = min(max_time, work_end)

            if actual_start < actual_end:
                work_duration = (actual_end - actual_start).total_seconds()
                lunch_deduction = (work_duration / 3600) * 6 * 60  # Descuento de 6 minutos por hora
                adjusted_duration = max(0, work_duration - lunch_deduction)
                data['average_time'] = adjusted_duration / len(data['times']) / 60 if len(data['times']) > 0 else 0
            else:
                data['average_time'] = 0
        else:
            data['average_time'] = 0

    return user_data, total_closed



def fetch_pending_orders():
    SPREADSHEET_ID = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
    RANGE_NAME = 'PEDIDOS_PENDIENTES!A:E'  # Ajuste para incluir hasta la columna E
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])
    
    # Filtrar filas donde A, B, C y D estén vacías
    filtered_values = [row for row in values if len(row) >= 4 and all(row[i].strip() for i in range(4))]
    
    return filtered_values



def update_order_status(row_index, new_status):
    SPREADSHEET_ID = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
    range_name = f'PEDIDOS_PENDIENTES!E{row_index + 1}'  # Sumamos 1 porque las filas en Google Sheets son 1-indexadas
    value_input_option = "RAW"
    value_range_body = {
        "values": [
            [new_status]
        ]
    }
    request = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=range_name,
        valueInputOption=value_input_option, body=value_range_body)
    response = request.execute()
    return response


def create_pending_orders_tab(tab_control):
    pending_tab = ctk.CTkFrame(tab_control)
    tab_control.add(pending_tab, text='PEDIDOS PENDIENTES')

    # Aumentar tamaño de fuente y ajustar el tamaño de las columnas y filas
    style = ttk.Style()
    style.configure("Treeview", font=("Helvetica", 14), rowheight=30)  # Ajusta la altura de las filas y la fuente
    style.configure("Treeview.Heading", font=("Helvetica", 16, "bold"))  # Fuente para los encabezados

    columns = ('ID', 'MOTIVO', 'USUARIO', 'SKU', 'ESTADO')
    tree = ttk.Treeview(pending_tab, columns=columns, show='headings', style="Treeview")
    tree.heading('ID', text='ID')
    tree.heading('MOTIVO', text='MOTIVO')
    tree.heading('USUARIO', text='USUARIO')
    tree.heading('SKU', text='SKU')
    tree.heading('ESTADO', text='ESTADO')
    tree.pack(expand=True, fill='both')

    # Lógica para seleccionar y actualizar entradas como antes

    return tree



def create_pie_chart(data, frame):
    for widget in frame.winfo_children():
        widget.destroy()

    labels = [user for user in data.keys()]
    sizes = [details['count'] for user, details in data.items()]
    
    fig, ax = plt.subplots()
    # Aumentar tamaño de las etiquetas del gráfico
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', textprops={'fontsize': 22})
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    plt.close(fig)  # Cerrar la figura después de dibujarla




def populate_treeview(tree, data):
    tree.delete(*tree.get_children())
    for user, details in data.items():
        tree.insert('', 'end', values=(user, details['count'], f"{details['average_time']:.2f} minutos"))


def update_user_tabs(user_data, tab_control):
    global user_tabs
    current_users = set(user_data.keys())
    existing_tabs = set(user_tabs.keys())

    # Eliminar pestañas para usuarios que ya no existen
    for user in existing_tabs - current_users:
        tab_to_remove = user_tabs.pop(user)
        tab_control.forget(tab_to_remove)
        tab_to_remove.destroy()

    # Crear o actualizar pestañas para los usuarios actuales
    for user, details in user_data.items():
        if user not in user_tabs:
            # Crear una nueva pestaña si no existe
            user_tab = ctk.CTkFrame(tab_control)
            tab_control.add(user_tab, text=user)
            user_tabs[user] = user_tab
        # Siempre actualiza el gráfico, independientemente de si es nuevo o existente
        create_hourly_chart(details, user_tabs[user])

def count_pending_orders():
    spreadsheet_flex_id = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'
    spreadsheet_colecta_id = '1hJIk3bpR5zLzEux5lEb1eUevh21Z2cVhB4ySpWEKZNU'
    range_name = 'A2:K'

    result_flex = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_flex_id, 
        range=range_name
    ).execute()
    values_flex = result_flex.get('values', [])
    pending_flex_services = {}

    for row in values_flex:
        if len(row) > 10 and row[3] and row[10] == "PENDIENTE":
            service_type = row[2]
            pending_flex_services[service_type] = pending_flex_services.get(service_type, 0) + 1

    result_colecta = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_colecta_id, 
        range=range_name
    ).execute()
    values_colecta = result_colecta.get('values', [])
    pending_colecta_services = {}

    for row in values_colecta:
        if len(row) > 10 and row[3] and row[10] == "PENDIENTE":
            service_type = row[2]
            pending_colecta_services[service_type] = pending_colecta_services.get(service_type, 0) + 1

    return pending_flex_services, pending_colecta_services

def calculate_estimated_times(pending_services, num_employees=1):
    estimated_times = {}
    avg_time_per_package = 2.20  # tiempo promedio para cerrar un paquete en minutos
    adjusted_time_per_package = avg_time_per_package / max(1, num_employees)  # evita división por cero
    for service, count in pending_services.items():
        estimated_times[service] = count * adjusted_time_per_package
    return estimated_times

def display_estimated_times(estimated_times):
    """
    Muestra los tiempos estimados en la interfaz de usuario.
    :param estimated_times: Diccionario con los tiempos estimados por servicio.
    """
    for service, time in estimated_times.items():
        print(f"Servicio: {service}, Tiempo Estimado: {time} minutos")

def add_estimated_time_tab(tab_control, estimated_times):
    """
    Crea una nueva pestaña que muestra los tiempos estimados de cierre por servicio en horas y minutos si es mayor que una hora,
    y en minutos si es menos de una hora. Muestra el tiempo total estimado en el mismo formato.
    :param tab_control: Controlador de pestañas de la aplicación.
    :param estimated_times: Diccionario con los tiempos estimados por servicio.
    """
    # Busca si la pestaña ya existe y la borra para crear una nueva con datos actualizados
    for tab in tab_control.tabs():
        if tab_control.tab(tab, "text") == "Tiempo estimado":
            tab_control.forget(tab)
            break

    estimated_tab = ctk.CTkFrame(tab_control)
    tab_control.add(estimated_tab, text="Tiempo estimado")

    # Agregar los datos de tiempos estimados a la pestaña
    total_estimated_time = 0
    for service, time in estimated_times.items():
        total_estimated_time += time
        if time < 60:
            time_text = f"{time:.2f} minutos"
        else:
            hours = int(time // 60)
            minutes = time % 60
            time_text = f"{hours}h {minutes:.2f}m"
        label = ctk.CTkLabel(estimated_tab, text=f"Servicio: {service}, Tiempo Estimado: {time_text}")
        label.pack(pady=5, padx=10, anchor='w')
    
    # Agregar etiqueta para el tiempo total estimado en horas y minutos o solo minutos
    if total_estimated_time < 60:
        total_time_text = f"{total_estimated_time:.2f} minutos"
    else:
        total_hours = int(total_estimated_time // 60)
        total_minutes = total_estimated_time % 60
        total_time_text = f"{total_hours}h {total_minutes:.2f}m"

    total_time_label = ctk.CTkLabel(estimated_tab, text=f"Tiempo total estimado de trabajo para completar todos los servicios: {total_time_text}")
    total_time_label.pack(pady=10, padx=10, anchor='w')

def update_data(tree,user_cards_tab, tab1, total_label, tab_control, pending_flex_label, pending_colecta_label, num_employees, pending_orders_tree, auto_refresh=False):
    global pending_order_ids
    try:
        user_data, total_closed = fetch_data()
        populate_treeview(tree, user_data)
        create_pie_chart(user_data, tab1)
        total_label.configure(text=f"Total de paquetes cerrados: {total_closed}")
        update_user_tabs(user_data, tab_control)
        
        pending_flex, pending_colecta = count_pending_orders()
        pending_flex_text = "\n".join(f"{service}: {count}" for service, count in pending_flex.items())
        pending_colecta_text = "\n".join(f"{service}: {count}" for service, count in pending_colecta.items())
        
        pending_flex_label.configure(text=f"PENDIENTES FLEX:\n{pending_flex_text}", text_color="YELLOW", fg_color="black", font=("Arial", 16, "bold"))
        pending_colecta_label.configure(text=f"PENDIENTES COLECTA:\n{pending_colecta_text}", text_color="red", fg_color="black", font=("Arial", 16, "bold"))

        num_employees = int(num_employees)
        estimated_times_flex = calculate_estimated_times(pending_flex, num_employees)
        estimated_times_colecta = calculate_estimated_times(pending_colecta, num_employees)
        add_estimated_time_tab(tab_control, {**estimated_times_flex, **estimated_times_colecta})

        pending_orders = fetch_pending_orders()
        new_order_ids = {row[0] for row in pending_orders if len(row) >= 1}

        # Detect new orders and count unique IDs
        new_pending_ids = new_order_ids - pending_order_ids
        total_unique_pending_ids = len(new_order_ids)

        if new_pending_ids:
            unique_new_pending_ids = len(new_pending_ids)
            notification_message = f"Hay {unique_new_pending_ids} nuevos pedidos por resolver. Total de pendientes: {total_unique_pending_ids}."
        else:
            notification_message = f"No hay nuevos pedidos pendientes. Total pendientes: {total_unique_pending_ids}."

        notification.notify(
            title='Actualización de Pedidos Pendientes',
            message=notification_message,
            timeout=10
        )

        pending_order_ids = new_order_ids

        pending_orders_tree.delete(*pending_orders_tree.get_children())
        for index, row in enumerate(pending_orders):
            if len(row) >= 5:  # Asegurarse de que la fila tenga al menos 5 columnas
                pending_orders_tree.insert('', 'end', values=(row[0], row[1], row[2], row[3], row[4]), iid=index)

        if auto_refresh:
            root.after(15000, lambda: update_data(tree,user_cards_tab, tab1, total_label, tab_control, pending_flex_label, pending_colecta_label, num_employees, pending_orders_tree, True))
    except Exception as e:
        print(f"Error updating data: {e}")



def create_hourly_chart(details, frame):
    for widget in frame.winfo_children():
        widget.destroy()
    
    hours = sorted(details['hourly'].keys())
    counts = [details['hourly'][hour] for hour in hours]
    fig, ax = plt.subplots()
    ax.bar(hours, counts, color='blue')
    ax.set_title("Paquetes por hora")
    ax.set_xlabel("Hora del día")
    ax.set_ylabel("Paquetes cerrados")
    ax.set_xticks(hours)
    ax.set_xticklabels([f"{hour}:00" for hour in hours])

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=ctk.BOTH, expand=True)
    plt.close(fig)  # Cerrar la figura después de dibujarla

def create_user_cards_tab(tab_control, user_data):
    user_cards_tab = ctk.CTkFrame(tab_control)
    tab_control.add(user_cards_tab, text="Resumen de Usuarios")

    # Frame contenedor que permitirá un "wrapping" automático
    cards_frame = ctk.CTkFrame(user_cards_tab, fg_color="white")
    cards_frame.pack(fill='both', expand=True, padx=20, pady=20)

    total_operations = sum(details['count'] for details in user_data.values())

    # Usamos un contenedor adicional que permite ajustar automáticamente las cards
    wrapping_frame = tk.Frame(cards_frame, bg="white")  # Asegurándose de que tenga el mismo color de fondo
    wrapping_frame.pack(fill='both', expand=True)

    # Creación de cards dentro del contenedor que permite "wrapping"
    for user, details in user_data.items():
        card = ctk.CTkFrame(wrapping_frame, corner_radius=10, fg_color="#E0F7FA", width=200, height=150)  # Tamaño fijo para uniformidad
        card.pack(side='left', padx=10, pady=10, fill='both', expand=False)

        user_label = ctk.CTkLabel(card, text=user, font=("Arial", 32, "bold"), fg_color="black")
        user_label.pack(pady=(10, 2), padx=10, anchor='w')

        count_label = ctk.CTkLabel(card, text=f"Paquetes_cerrados: {details['count']}", font=("Arial", 32, "bold"), fg_color="black")
        count_label.pack(pady=2, padx=10, anchor='w')

        percentage = (details['count'] / total_operations * 100) if total_operations > 0 else 0
        percentage_label = ctk.CTkLabel(card, text=f"Porcentaje del total: {percentage:.2f}%", font=("Arial", 21, "bold"), fg_color="black")
        percentage_label.pack(pady=2, padx=10, anchor='w')

        if 'average_time' in details:
            time_label = ctk.CTkLabel(card, text=f"Tiempo promedio: {details['average_time']:.2f} minutos", font=("Arial", 20, "bold"), fg_color="black")
            time_label.pack(pady=(2, 10), padx=10, anchor='w')

    return user_cards_tab

def main():
    global user_tabs, root, pending_order_ids,user_cards_tab
    user_tabs = {}
    pending_order_ids = set()

    root = ctk.CTk()
    root.title("Rendimientos")
    
    move_to_screen2(root)  # Mueve la ventana a la pantalla 2

    tab_control = ttk.Notebook(root)
    tab1 = ctk.CTkFrame(tab_control)
    tab2 = ctk.CTkFrame(tab_control)
    tab_control.add(tab1, text='Pie Chart')
    tab_control.add(tab2, text='Data')
    
    # Añadir la nueva pestaña para las cards de usuario
    user_data, _ = fetch_data()  # Suponiendo que esto recupera tus datos de usuario
    create_user_cards_tab(tab_control, user_data)
    
    tab_control.pack(expand=True, fill='both', padx=20, pady=20)


    columns = ('USUARIO', 'TOTAL DE PAQUETES CERRADOS', 'PAQUETES X (min)')
    tree = ttk.Treeview(tab2, columns=columns, show='headings')
    tree.heading('USUARIO', text='USUARIO')
    tree.heading('TOTAL DE PAQUETES CERRADOS', text='TOTAL DE PAQUETES CERRADOS')
    tree.heading('PAQUETES X (min)', text='Promedio de tiempo (min)')
    tree.pack(expand=True, fill='both')

    total_label = ctk.CTkLabel(root, text="", text_color="black", fg_color="#00FF00", font=("Arial", 14, "bold"))
    total_label.pack(side=tk.BOTTOM, fill=tk.X)

    pending_flex_label = ctk.CTkLabel(root, text="", text_color="red", fg_color="black", font=("Arial", 16, "bold"))
    pending_flex_label.pack(side=tk.BOTTOM, fill=tk.X)

    pending_colecta_label = ctk.CTkLabel(root, text="", text_color="red", fg_color="black", font=("Arial", 16, "bold"))
    pending_colecta_label.pack(side=tk.BOTTOM, fill=tk.X)

    employee_frame = ctk.CTkFrame(root)
    employee_frame.pack(fill='x', padx=20, pady=10)
    employee_label = ctk.CTkLabel(employee_frame, text="Número de empleados:")
    employee_label.pack(side='left')
    num_employees = tk.StringVar(value="1")
    employee_dropdown = ttk.Combobox(employee_frame, textvariable=num_employees, values=[str(i) for i in range(1, 6)])
    employee_dropdown.pack(side='left', padx=10)

    update_button = ctk.CTkButton(root, text="Actualizar", command=lambda: update_data(tree, tab1, total_label, tab_control, pending_flex_label, pending_colecta_label, num_employees.get(), pending_orders_tree))
    update_button.pack(pady=10)

    pending_orders_tree = create_pending_orders_tab(tab_control)

    update_data(tree,user_cards_tab, tab1, total_label, tab_control, pending_flex_label, pending_colecta_label, num_employees.get(), pending_orders_tree, True)

    root.mainloop()

if __name__ == "__main__":
    main()



""" #Traceback (most recent call last):
File "c:\Users\todom\OneDrive\Escritorio\CERRADOR\Rendimiento.py", line 478, in <module>
main()
File "c:\Users\todom\OneDrive\Escritorio\CERRADOR\Rendimiento.py", line 473, in main
    update_data(tree,user_cards_tab, tab1, total_label, tab_control, pending_flex_label, pending_colecta_label, num_employees.get(), pending_orders_tree, True)
                    ^^^^^^^^^^^^^^
NameError: name 'user_cards_tab' is not defined """
#Falta definir en fetch data and update data y que no rompa otra funcion revisar