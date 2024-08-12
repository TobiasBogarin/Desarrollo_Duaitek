import logging
import sys
import traceback
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.popup import Popup
from kivy.clock import Clock
from kivy.uix.spinner import Spinner
from kivy.uix.checkbox import CheckBox
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import concurrent.futures
import time
from kivy.uix.scrollview import ScrollView
from kivy.uix.image import Image
from kivy.uix.anchorlayout import AnchorLayout
from kivy.core.window import Window
from kivy.uix.behaviors import ButtonBehavior

# Asegurarse de que la ventana esté en pantalla completa y tenga la resolución deseada
Window.size = (1600, 900)
Window.fullscreen = True

# Configuración del logging para registrar errores en un archivo de texto
logging.basicConfig(level=logging.INFO, filename='crash_report.txt', filemode='w', format='%(asctime)s - %(levelname)s - %(message)s')

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    with open('crash_report.txt', 'a') as f:
        f.write("Uncaught exception:\n")
        traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)

sys.excepthook = handle_exception

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
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

SHEET_ID_1 = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'
SHEET_ID_2 = '1hJIk3bpR5zLzEux5lEb1eUevh21Z2cVhB4ySpWEKZNU'
SHEET_ID_TIEMPO = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
RANGE_NAME = 'Presentacion y cierre!A:D'
RANGE_NAME_TIEMPO = 'Tiempo de presentado!A:D'
RANGE_NAME_DETALLE = 'Detalle de pedidos!A:C'
RANGE_NAME_USUARIOS = 'USUARIOS!A:A'
RANGE_NAME_UBICACIONES = 'UBICACIONES!A:C'
RANGE_NAME_IDSKU = 'IDSKU!A:B'

ubicaciones_sku = {}
ids_sku = {}

def cargar_ubicaciones_sku():
    result = sheet.values().get(spreadsheetId=SHEET_ID_1, range=RANGE_NAME_UBICACIONES).execute()
    values = result.get('values', [])

    for row in values:
        if len(row) >= 3:
            sku = row[0]
            ubi1 = row[1]
            ubi2 = row[2]
            ubicaciones_sku[sku] = (ubi1, ubi2)

def cargar_ids_sku():
    result = sheet.values().get(spreadsheetId=SHEET_ID_TIEMPO, range=RANGE_NAME_IDSKU).execute()
    values = result.get('values', [])

    for row in values:
        if len(row) >= 2:
            sku = row[0]
            id_sku = row[1]
            ids_sku[sku] = id_sku

def obtener_usuarios():
    result = sheet.values().get(spreadsheetId=SHEET_ID_TIEMPO, range=RANGE_NAME_USUARIOS).execute()
    values = result.get('values', [])
    return [row[0] for row in values if row]

def obtener_shipment_id(id_interno):
    if 10000000 <= id_interno <= 19999999:
        sheet_id = SHEET_ID_1
    elif 20000000 <= id_interno <= 29999999:
        sheet_id = SHEET_ID_2
    else:
        return None

    result = sheet.values().get(spreadsheetId=sheet_id, range=RANGE_NAME).execute()
    values = result.get('values', [])

    for row in values:
        if len(row) > 3 and row[0] == str(id_interno):
            return row[3]
    return None

def obtener_tokens(clientes):
    url = "https://api.mercadolibre.com/oauth/token"
    tokens = []
    for client in clientes:
        payload = {
            'grant_type': 'client_credentials',
            'client_id': client['client_id'],
            'client_secret': client['client_secret']
        }
        try:
            response = requests.post(url, data=payload)
            response.raise_for_status()
            tokens.append(response.json()['access_token'])
            print(f"Token obtenido con éxito para client_id {client['client_id']}")
        except requests.RequestException as e:
            print(f"Error al obtener token con client_id {client['client_id']}: {e}")
    return tokens

clientes = [
    {'client_id': '8595972831111833', 'client_secret': 'fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L'},
    {'client_id': '8545932337071872', 'client_secret': 'VtRNp52gxGKCDKLvR8YRfovPLL3ZBfIk'}
]

tokens = obtener_tokens(clientes)

def obtener_skus(access_token, item_ids):
    headers = {'Authorization': f'Bearer {access_token}'}
    skus = {}
    
    def fetch_sku(item_id):
        url = f"https://api.mercadolibre.com/items/{item_id}?include_attributes=all"
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()
            sku = None

            attributes = data.get('attributes', [])
            for attribute in attributes:
                if attribute.get('id') == 'SELLER_SKU':
                    sku = attribute.get('value_name')
                    if sku:
                        break

            if not sku and 'variations' in data:
                for variation in data['variations']:
                    var_attributes = variation.get('attributes', [])
                    for attribute in var_attributes:
                        if attribute.get('id') == 'SELLER_SKU':
                            sku = attribute.get('value_name')
                            if sku:
                                break

            if not sku:
                print(f"SKU no encontrado para {item_id}. Utilizando valor por defecto.")
                sku = 'SKU no disponible'
                
            return item_id, sku
        except requests.RequestException as e:
            print(f"Error al obtener información para {item_id}: {e}")
            return item_id, 'Error al obtener datos'
    
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(fetch_sku, item_ids)
    
    for item_id, sku in results:
        skus[item_id] = sku
    
    return skus

def obtener_datos_de_envio(access_token, shipment_id, id_interno):
    url = f"https://api.mercadolibre.com/shipments/{shipment_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        datos = response.json()
        print("Datos JSON obtenidos:", datos)  # Imprimir los datos JSON en la terminal

        order_id = datos.get('order_id')
        variaciones = []
        if order_id:
            variaciones = obtener_datos_de_orden(access_token, order_id)

        # Obtener los IDs de los artículos del envío
        item_ids = [item['id'] for item in datos.get('shipping_items', [])]

        # Obtener los SKUs asociados a los IDs de los artículos
        skus = obtener_skus(access_token, item_ids)

        # Asociar cada artículo con sus ubicaciones y variaciones
        items_with_locations = [
            (item['id'], skus.get(item['id'], 'SKU no encontrado'), item['quantity'], ubicaciones_sku.get(skus.get(item['id'], 'SKU no encontrado'), ('NINGUNA', 'NINGUNA')), variaciones)
            for item in datos.get('shipping_items', [])
        ]

        # Ordenar los artículos por su primera ubicación
        sorted_items = sorted(items_with_locations, key=lambda x: ('999' if x[3][0] == "NINGUNA" else x[3][0]))

        return id_interno, sorted_items, variaciones

    except requests.HTTPError as http_err:
        print(f"HTTP error occurred while fetching shipment data for {shipment_id}: {http_err}")
    except requests.RequestException as req_err:
        print(f"Error occurred while fetching shipment data for {shipment_id}: {req_err}")
    except KeyError as key_err:
        print(f"Key error: {key_err} - possibly missing data in the response for shipment {shipment_id}")
    except Exception as err:
        print(f"An unexpected error occurred while fetching shipment data for {shipment_id}: {err}")

    return id_interno, [], []

def obtener_variaciones_color(access_token, order_ids):
    headers = {'Authorization': f'Bearer {access_token}'}
    variaciones = []
    
    for order_id in order_ids:
        url = f"https://api.mercadolibre.com/orders/{order_id}"
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            datos_orden = response.json()
            
            for item in datos_orden.get('order_items', []):
                color_variation = "-"
                sku = item['item'].get('seller_sku', 'SKU no disponible')
                for attribute in item['item'].get('variation_attributes', []):
                    if attribute.get('id') == 'COLOR':
                        color_variation = attribute.get('value_name', '-')
                        break
                variaciones.append({'sku': sku, 'color': color_variation})

        except requests.RequestException as e:
            print(f"Error al obtener información de la orden {order_id}: {e}")
            variaciones.append({'sku': 'SKU no disponible', 'color': '-'})
    
    return variaciones




def obtener_datos_de_orden(access_token, order_id):
    url = f"https://api.mercadolibre.com/orders/{order_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        datos_orden = response.json()
        print("Datos de la orden obtenidos:", datos_orden)

        pack_id = datos_orden.get('pack_id')
        if pack_id:
            datos_carrito = obtener_datos_de_carrito(access_token, pack_id)
            order_ids = [order['id'] for order in datos_carrito.get('orders', [])]
            variaciones = obtener_variaciones_color(access_token, order_ids)
            print("Variaciones de color obtenidas:", variaciones)
            return variaciones

        variaciones = []
        for item in datos_orden.get('order_items', []):
            color_variation = "-"
            for attribute in item['item'].get('variation_attributes', []):
                if attribute.get('id') == 'COLOR':
                    color_variation = attribute.get('value_name', '-')
                    break
            variaciones.append(color_variation)
        
        print("Variaciones obtenidas:", variaciones)
        return variaciones

    except requests.HTTPError as http_err:
        print(f"HTTP error occurred while fetching order data for {order_id}: {http_err}")
    except requests.RequestException as req_err:
        print(f"Error occurred while fetching order data for {order_id}: {req_err}")
    except Exception as err:
        print(f"An unexpected error occurred while fetching order data for {order_id}: {err}")
        return []

def consultar_carritos(access_token, pack_id):
    url = f"https://api.mercadolibre.com/packs/{pack_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    variaciones_totales = []
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        datos_carrito = response.json()
        print("Datos del carrito obtenidos:", datos_carrito)

        for order in datos_carrito.get('orders', []):
            order_id = order['id']
            variaciones = obtener_datos_de_orden(access_token, order_id)
            variaciones_totales.extend(variaciones)
        
        return datos_carrito, variaciones_totales
    except requests.HTTPError as http_err:
        print(f"HTTP error occurred while fetching cart data for {pack_id}: {http_err}")
    except requests.RequestException as req_err:
        print(f"Error occurred while fetching cart data for {pack_id}: {req_err}")
    except Exception as err:
        print(f"An unexpected error occurred while fetching cart data for {pack_id}: {err}")
        return {}, []

def obtener_datos_de_carrito(access_token, pack_id):
    url = f"https://api.mercadolibre.com/packs/{pack_id}"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        datos_carrito = response.json()
        print("Datos del carrito obtenidos:", datos_carrito)
        return datos_carrito
    except requests.HTTPError as http_err:
        print(f"HTTP error occurred while fetching cart data for {pack_id}: {http_err}")
    except requests.RequestException as req_err:
        print(f"Error occurred while fetching cart data for {pack_id}: {req_err}")
    except Exception as err:
        print(f"An unexpected error occurred while fetching cart data for {pack_id}: {err}")
        return {}

def format_time(seconds):
    minutes = int(seconds // 60)
    remaining_seconds = seconds % 60
    formatted_time = f"{minutes:02}:{remaining_seconds:05.2f}"
    return formatted_time

def enviar_tiempos_a_sheets(tiempos, nombre_tanda, usuario):
    values = [[str(id_interno), format_time(tiempo), usuario, nombre_tanda] for id_interno, tiempo in tiempos]
    body = {'values': values}
    result = sheet.values().append(
        spreadsheetId=SHEET_ID_TIEMPO,
        range=RANGE_NAME_TIEMPO,
        valueInputOption='RAW',
        body=body
    ).execute()
    print(f'{result.get("updates").get("updatedCells")} celdas actualizadas.')


def enviar_detalle_pedidos_a_sheets(pedidos_data):
    values = []
    for id_interno, items in pedidos_data:
        for item in items:
            item_id, sku, cantidad, ubicaciones = item[:4]
            values.append([id_interno, sku, cantidad])
    body = {'values': values}
    result = sheet.values().append(
        spreadsheetId=SHEET_ID_TIEMPO,
        range=RANGE_NAME_DETALLE,
        valueInputOption='RAW',
        body=body
    ).execute()
    print(f'{result.get("updates").get("updatedCells")} celdas actualizadas.')

class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        
        layout = BoxLayout(orientation='vertical', padding=[40, 100, 40, 40], spacing=20)
        
        header_layout = AnchorLayout(size_hint=(1, 0.3))
        header_label = Label(text='Selecciona un usuario', font_size='32sp', bold=True, color=(0.1, 0.5, 0.8, 1))
        header_layout.add_widget(header_label)
        layout.add_widget(header_layout)
        
        self.spinner = Spinner(
            text='Seleccione Usuario',
            values=obtener_usuarios(),
            size_hint=(1, 0.2),
            font_size='20sp',
            background_color=(0.3, 0.3, 0.3, 1),
            color=(1, 1, 1, 1)
        )
        layout.add_widget(self.spinner)
        
        self.login_button = Button(
            text='Login',
            size_hint=(1, 0.2),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.login_button.bind(on_press=self.login)
        layout.add_widget(self.login_button)
        
        self.add_widget(layout)
    
    def login(self, instance):
        if self.spinner.text == 'Seleccione Usuario':
            self.mostrar_popup('Error', 'Debe seleccionar un usuario')
        else:
            app = App.get_running_app()
            app.usuario = self.spinner.text
            app.screen_manager.current = 'scan_pedidos'
    
    def mostrar_popup(self, title, message):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text=message, font_size='18sp', halign='center', color=(1, 1, 1, 1)))
        close_button = Button(text='Cerrar', size_hint=(1, 0.2), font_size='18sp')
        content.add_widget(close_button)
        popup = Popup(title=title, content=content, size_hint=(0.75, 0.5), background_color=(0.3, 0.3, 0.3, 1))
        close_button.bind(on_press=popup.dismiss)
        popup.open()


class CustomTextInput(TextInput):
    def insert_text(self, substring, from_undo=False):
        # Filtrar espacios y caracteres no deseados
        s = ''.join(char for char in substring if char not in ' ')
        return super().insert_text(s, from_undo=from_undo)

class CustomButton(Button):
    pass


class ScanPedidosScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.pedidos = []

        layout = BoxLayout(orientation='vertical', padding=[40, 100, 40, 40], spacing=20)

        header_layout = AnchorLayout(size_hint=(1, 0.3))
        header_label = Label(text='Escaneo de Pedidos', font_size='32sp', bold=True, color=(0.1, 0.5, 0.8, 1))
        header_layout.add_widget(header_label)
        layout.add_widget(header_layout)

        self.input = CustomTextInput(
            hint_text='Ingrese el ID Interno',
            size_hint=(1, 0.1),
            multiline=False,
            font_size='20sp',
            foreground_color=(0, 0, 0, 1),  # Color del texto (negro)
            background_color=(1, 1, 1, 1),  # Color de fondo (blanco)
            hint_text_color=(0.5, 0.5, 0.5, 1)  # Color del texto del hint
        )
        self.input.bind(on_text_validate=self.agregar_pedido)
        layout.add_widget(self.input)

        self.add_button = Button(
            text='Agregar Pedido',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.add_button.bind(on_press=self.agregar_pedido)
        layout.add_widget(self.add_button)

        self.counter_label = Label(
            text='Pedidos Escaneados: 0',
            size_hint=(1, 0.1),
            font_size='20sp',
            color=(0.1, 0.5, 0.8, 1)
        )
        layout.add_widget(self.counter_label)

        self.start_button = Button(
            text='Empezar Presentación',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.start_button.bind(on_press=self.empezar_presentacion)
        layout.add_widget(self.start_button)

        self.add_widget(layout)
        Clock.schedule_once(self.focus_input, 0.1)

    def on_enter(self):
        super().on_enter()
        self.input.focus = True
        Clock.schedule_interval(self.mantener_foco, 0.5)

    def on_leave(self):
        super().on_leave()
        Clock.unschedule(self.mantener_foco)

    def focus_input(self, dt):
        self.input.focus = True

    def mantener_foco(self, dt):
        if not self.input.focus:
            self.input.focus = True

    def agregar_pedido(self, instance):
        try:
            id_interno = int(self.input.text)
            if id_interno in self.pedidos:
                self.mostrar_popup('Error', 'El ID ya ha sido agregado')
            else:
                self.pedidos.append(id_interno)
                self.input.text = ''
                self.counter_label.text = f'Pedidos Escaneados: {len(self.pedidos)}'
            Clock.schedule_once(self.focus_input, 0.1)
        except ValueError:
            self.mostrar_popup('Error', 'Por favor, ingrese un ID Interno válido')
            Clock.schedule_once(self.focus_input, 0.1)

    def empezar_presentacion(self, instance):
        if not self.pedidos:
            self.mostrar_popup('Error', 'No se han agregado pedidos')
            return
        app = App.get_running_app()
        app.presentacion_screen.set_pedidos(self.pedidos)
        app.screen_manager.current = 'presentacion'

    def reset(self):
        self.pedidos = []
        self.counter_label.text = 'Pedidos Escaneados: 0'

    def mostrar_popup(self, title, message):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text=message, font_size='18sp', halign='center'))
        close_button = Button(text='Cerrar', size_hint=(1, 0.2), font_size='18sp')
        content.add_widget(close_button)
        popup = Popup(title=title, content=content, size_hint=(0.75, 0.5))
        close_button.bind(on_press=popup.dismiss)
        popup.open()



logging.basicConfig(level=logging.INFO)

class PresentacionScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.layout = BoxLayout(orientation='vertical', padding=[40, 60, 40, 40], spacing=20)
        self.pedidos = []
        self.pedidos_data = []
        self.current_pedido_index = 0
        self.tiempos = []
        self.start_time = None

        header_layout = AnchorLayout(size_hint=(1, 0.3))
        header_label = Label(text='Presentación de Pedidos', font_size='32sp', bold=True, color=(1, 1, 1, 1))
        header_layout.add_widget(header_label)
        self.layout.add_widget(header_layout)

        grid_header = GridLayout(cols=6, size_hint=(1, 0.1))
        grid_header.add_widget(Label(text='ID Interno', size_hint_x=0.1, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='SKU', size_hint_x=0.2, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='Cantidad', size_hint_x=0.1, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='Ubicación 1', size_hint_x=0.2, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='Ubicación 2', size_hint_x=0.2, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='Verificación', size_hint_x=0.1, font_size='20sp', bold=True, color=(1, 1, 1, 1)))
        self.layout.add_widget(grid_header)

        self.resultados = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.resultados.bind(minimum_height=self.resultados.setter('height'))
        scroll_view = ScrollView(size_hint=(1, 0.6))
        scroll_view.add_widget(self.resultados)
        self.layout.add_widget(scroll_view)

        self.counter_label = Label(
            text='Pedidos Restantes: 0',
            size_hint=(1, 0.1),
            font_size='20sp',
            color=(1, 1, 1, 1)
        )
        self.layout.add_widget(self.counter_label)

        self.scan_input = CustomTextInput(
            hint_text='Escanear articulo',
            size_hint=(1, 0.1),
            multiline=False,
            font_size='24sp',  # Tamaño de fuente aumentado
            foreground_color=(0, 0, 0, 1),  # Color del texto (negro)
            background_color=(1, 1, 0, 1),  # Color de fondo (amarillo claro)
            hint_text_color=(0.5, 0.5, 0.5, 1),  # Color del texto del hint (gris)
            padding=(8, 8)  # Padding para más visibilidad
        )
        self.scan_input.bind(on_text_validate=self.validar_articulo)
        self.layout.add_widget(self.scan_input)

        self.next_button = CustomButton(
            text='Siguiente Pedido',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.next_button.bind(on_press=self.siguiente_pedido)
        self.layout.add_widget(self.next_button)
        self.next_button.disabled = True

        self.complete_button = CustomButton(
            text='Tanda Completada',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.complete_button.bind(on_press=self.tanda_completada)
        self.layout.add_widget(self.complete_button)
        self.complete_button.disabled = True

        self.pending_button = CustomButton(
            text='Marcar como Pendiente',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.9, 0.3, 0.3, 1),
            color=(1, 1, 1, 1)
        )
        self.pending_button.bind(on_press=self.abrir_menu_pendiente)
        self.layout.add_widget(self.pending_button)

        self.show_color_button = CustomButton(
            text='Mostrar Color',
            size_hint=(1, 0.1),
            font_size='20sp',
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.show_color_button.bind(on_press=self.mostrar_color)
        self.layout.add_widget(self.show_color_button)

        self.add_widget(self.layout)

        self.articulos_escaneados = {}
        self.colores_pedidos = []
        Clock.schedule_once(self.focus_input, 0.1)
        Window.bind(mouse_pos=self.on_mouse_pos)

    def on_enter(self):
        super().on_enter()
        self.scan_input.focus = True
        Clock.schedule_interval(self.mantener_foco, 0.5)

    def on_leave(self):
        super().on_leave()
        Clock.unschedule(self.mantener_foco)

    def mantener_foco(self, dt):
        if not self.scan_input.focus:
            self.scan_input.focus = True

    def on_mouse_pos(self, window, pos):
        if not self.collide_point(*pos):
            return

        focus_widget = self._get_focus_widget()
        if not isinstance(focus_widget, CustomButton):
            self.focus_input(0)

    def _get_focus_widget(self):
        from kivy.core.window import Window
        return Window.children[0]

    def set_pedidos(self, pedidos):
        self.pedidos = sorted(pedidos)
        self.pedidos_data = []
        self.tiempos = []
        self.current_pedido_index = 0
        self.articulos_escaneados = {}
        self.colores_pedidos = []
        self.complete_button.disabled = True
        self.start_time = time.time()

        # Obtener y ordenar los datos del envío por ubicación
        for id_interno in self.pedidos:
            shipment_id = obtener_shipment_id(id_interno)
            if shipment_id:
                for token in tokens:
                    id_interno, items, colores = obtener_datos_de_envio(token, shipment_id, id_interno)
                    if items:
                        self.pedidos_data.append((id_interno, items))
                        # Almacenar los colores de las variaciones con el ID del pedido
                        self.colores_pedidos.append((id_interno, colores))
                        break

        # Ordenar los pedidos en base a la ubicación de sus SKUs
        self.pedidos_data.sort(key=lambda x: (
            'z' if len(x[1]) > 1 else '999' if any(item[3][0] == "NINGUNA" for item in x[1]) else x[1][0][3][0]
        ))
        self.counter_label.text = f'Pedidos Restantes: {len(self.pedidos_data) - self.current_pedido_index}'
        self.mostrar_pedido()

    def mostrar_pedido(self):
        self.resultados.clear_widgets()
        if self.current_pedido_index < len(self.pedidos_data):
            id_interno, items = self.pedidos_data[self.current_pedido_index]
            for item_id, sku, quantity, (ubi1, ubi2), variaciones in items:
                ubi1, ubi2 = ubicaciones_sku.get(sku, ('No disponible', 'No disponible'))
                self.articulos_escaneados[sku] = 0
                item_layout = GridLayout(cols=6, spacing=10, size_hint_y=None, height=50)
                item_layout.add_widget(Label(text=f'{id_interno}', size_hint_x=0.1, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{sku}', size_hint_x=0.2, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{quantity}', size_hint_x=0.1, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{ubi1}', size_hint_x=0.2, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{ubi2}', size_hint_x=0.2, font_size='24sp', color=(1, 1, 1, 1)))
                checkbox = CheckBox(active=False, size_hint_x=0.1)
                item_layout.add_widget(checkbox)
                self.resultados.add_widget(item_layout)
                self.articulos_escaneados[sku] = {
                    'checkbox': checkbox,
                    'cantidad': quantity,
                    'contador': 0
                }
        else:
            self.complete_button.disabled = False

    def mostrar_color(self, instance):
        if self.current_pedido_index < len(self.colores_pedidos):
            id_interno_actual, _ = self.pedidos_data[self.current_pedido_index]

            for id_interno, colores in self.colores_pedidos:
                if id_interno == id_interno_actual:
                    # Crear un diccionario de SKU a colores
                    sku_color_dict = {}
                    for color_info in colores:
                        if isinstance(color_info, dict):
                            sku = color_info.get('sku', id_interno_actual)
                            color_name = color_info.get('color', '-')
                        else:
                            # Si color_info es una cadena, asumimos que es solo el color y no hay SKU disponible
                            sku = id_interno_actual
                            color_name = color_info

                        if sku in sku_color_dict:
                            sku_color_dict[sku].add(color_name)
                        else:
                            sku_color_dict[sku] = {color_name}

                    # Formatear la información de color y SKU para mostrarla en el popup
                    color_info_str = "\n".join([f"{sku}: {', '.join(colors)}" for sku, colors in sku_color_dict.items()])
                    color_info_str = color_info_str if color_info_str else "Color no disponible"
                    self.mostrar_popup(f'Colores del Pedido {id_interno}', color_info_str)
                    return

            self.mostrar_popup('Error', 'No hay información de color para mostrar.')
        else:
            self.mostrar_popup('Error', 'No hay información de color para mostrar.')

    def validar_articulo(self, instance):
        articulo_id = self.scan_input.text.strip()
        self.scan_input.text = ''
        Clock.schedule_once(self.focus_input, 0.1)

        try:
            id_escaneado = articulo_id[-5:-1]
            id_escaneado_int = int(id_escaneado)
            id_escaneado = str(id_escaneado_int)
            logging.info(f'ID escaneado: {id_escaneado}')
            sku = None
            for stored_sku, stored_id in ids_sku.items():
                if stored_id == id_escaneado:
                    sku = stored_sku
                    break
            if not sku:
                raise ValueError("ID no encontrado")
        except ValueError:
            logging.error(f'Error: Formato de ID incorrecto o ID no encontrado. Entrada: {articulo_id}')
            self.mostrar_popup('Error', 'Formato de ID incorrecto o ID no encontrado.')
            return

        if self.current_pedido_index < len(self.pedidos_data):
            _, items = self.pedidos_data[self.current_pedido_index]
            for item_id, item_sku, quantity, (ubi1, ubi2), _ in items:
                if sku == item_sku:
                    self.articulos_escaneados[item_sku]['contador'] += 1
                    if self.articulos_escaneados[item_sku]['contador'] == self.articulos_escaneados[item_sku]['cantidad']:
                        self.articulos_escaneados[item_sku]['checkbox'].active = True
                    elif self.articulos_escaneados[item_sku]['contador'] > self.articulos_escaneados[item_sku]['cantidad']:
                        self.mostrar_popup('Excedente', f'El SKU {item_sku} ha sido escaneado más veces de las requeridas.')
                    break
            else:
                self.mostrar_popup('Error', 'El artículo escaneado no coincide con ningún SKU en la lista.')

            if all(self.articulos_escaneados[item_sku]['contador'] == self.articulos_escaneados[item_sku]['cantidad'] for item_id, item_sku, quantity, _, _ in items):
                self.next_button.disabled = False

    def siguiente_pedido(self, instance):
        if self.start_time is None:
            self.mostrar_popup('Error', 'Tiempo de inicio no establecido')
            return

        end_time = time.time()
        elapsed_time = end_time - self.start_time
        id_interno, _ = self.pedidos_data[self.current_pedido_index]
        self.tiempos.append((id_interno, elapsed_time))

        self.current_pedido_index += 1
        self.start_time = time.time()

        self.counter_label.text = f'Pedidos Restantes: {len(self.pedidos_data) - self.current_pedido_index}'
        self.mostrar_pedido()
        self.next_button.disabled = True

        if self.current_pedido_index >= len(self.pedidos_data):
            self.complete_button.disabled = False

    def tanda_completada(self, instance):
        if self.current_pedido_index < len(self.pedidos_data):
            self.mostrar_popup('Error', 'Aún hay pedidos pendientes')
            return

        self.mostrar_popup('Tanda Completada', 'Todos los pedidos han sido presentados')

        if self.pedidos:
            nombre_tanda = f'TANDA {min(self.pedidos)} - {max(self.pedidos)} ({sum(t[1] for t in self.tiempos):.2f} segundos)'
            app = App.get_running_app()
            enviar_tiempos_a_sheets(self.tiempos, nombre_tanda, app.usuario)
            enviar_detalle_pedidos_a_sheets(self.pedidos_data)

        app.screen_manager.current = 'view_pedidos'
        app.view_pedidos_screen.mostrar_pedido(self.pedidos_data)
        self.reset()

    def abrir_menu_pendiente(self, instance):
        if self.current_pedido_index < len(self.pedidos_data):
            _, items = self.pedidos_data[self.current_pedido_index]

            content = BoxLayout(orientation='vertical', padding=10, spacing=10)

            # Spinner para seleccionar el motivo
            dropdown = Spinner(
                text='Seleccione el Motivo',
                values=('Stock insuficiente en la sucursal solicitar traslado', 'No stock en primera ubicación reponer stock del sotano', 'Pendiente esperando traslado de tecnica'),
                size_hint=(1, 0.3),
                font_size='18sp'
            )
            content.add_widget(dropdown)

            # Layout para los SKU seleccionables
            sku_layout = BoxLayout(orientation='vertical', size_hint=(1, None))
            sku_layout.bind(minimum_height=sku_layout.setter('height'))
            self.sku_checkboxes = {}
            for item_id, sku, quantity, (ubi1, ubi2), _ in items:
                sku_checkbox = CheckBox()
                sku_label = Label(text=f"{sku} (Cantidad: {quantity})", font_size='18sp')
                box = BoxLayout(size_hint_y=None, height=50)
                box.add_widget(sku_checkbox)
                box.add_widget(sku_label)
                sku_layout.add_widget(box)
                self.sku_checkboxes[sku] = (sku_checkbox, quantity)

            scroll_view = ScrollView(size_hint=(1, 0.5))
            scroll_view.add_widget(sku_layout)
            content.add_widget(scroll_view)

            submit_button = Button(text='Confirmar', size_hint=(1, 0.2), font_size='18sp')
            submit_button.bind(on_press=lambda x: self.registrar_pendiente(dropdown.text))
            content.add_widget(submit_button)

            self.popup = Popup(title='Motivo del Estado Pendiente', content=content, size_hint=(0.75, 0.75))
            self.popup.open()

    def registrar_pendiente(self, motivo):
        if motivo.startswith('Seleccione'):
            self.mostrar_popup('Error', 'Debe seleccionar un motivo')
        else:
            skus_afectados = [(sku, quantity) for sku, (checkbox, quantity) in self.sku_checkboxes.items() if checkbox.active]
            if not skus_afectados:
                self.mostrar_popup('Error', 'Debe seleccionar al menos un SKU')
                return
            self.popup.dismiss()
            if self.current_pedido_index < len(self.pedidos_data):
                id_interno, _ = self.pedidos_data[self.current_pedido_index]
                self.enviar_datos_pendientes(id_interno, motivo, App.get_running_app().usuario, skus_afectados)
                self.siguiente_pedido(None)
            else:
                self.mostrar_popup('Error', 'Índice fuera de rango al registrar pendiente')

    def enviar_datos_pendientes(self, id_interno, motivo, usuario, skus_afectados):
        sheet_id = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
        range_name = 'PEDIDOS_PENDIENTES!A:D'

        # Obtener las filas actuales de las columnas A, B, C, y D
        result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get('values', [])

        # Calcular la primera fila vacía en las columnas A, B, C, y D
        first_empty_row = len(values) + 1  # Calcula la primera fila vacía

        # Preparar los valores para ser enviados
        data = [[id_interno, motivo, usuario, f"{sku} ({quantity} u)"] for sku, quantity in skus_afectados]
        body = {'values': data}

        # Especificar el rango para agregar los datos en la primera fila vacía
        range_to_append = f'PEDIDOS_PENDIENTES!A{first_empty_row}:D{first_empty_row + len(data) - 1}'
        result = sheet.values().append(
            spreadsheetId=sheet_id,
            range=range_to_append,
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f'{result.get("updates").get("updatedCells")} celdas actualizadas en PEDIDOS_PENDIENTES.')

    def mostrar_popup(self, title, message):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text=message, font_size='18sp', halign='center', color=(1, 1, 1, 1)))
        close_button = Button(text='Cerrar', size_hint=(1, 0.2), font_size='18sp')
        content.add_widget(close_button)
        popup = Popup(title=title, content=content, size_hint=(0.75, 0.5), background_color=(0.3, 0.3, 0.3, 1))
        close_button.bind(on_press=popup.dismiss)
        popup.open()

    def reset(self):
        self.pedidos = []
        self.pedidos_data = []
        self.current_pedido_index = 0
        self.tiempos = []
        self.start_time = None
        self.articulos_escaneados = {}
        self.colores_pedidos = []
        self.next_button.disabled = True
        self.complete_button.disabled = True

    def focus_input(self, dt):
        self.scan_input.focus = True

    def mantener_foco(self, dt):
        if not self.scan_input.focus:
            self.scan_input.focus = True




class ViewPedidosScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.layout = BoxLayout(orientation='vertical', padding=[40, 60, 40, 40], spacing=20)
        
        header_layout = AnchorLayout(size_hint=(1, 0.2))
        header_label = Label(text='Visualización de Pedidos', font_size='32sp', bold=True, color=(0.1, 0.5, 0.8, 1))
        header_layout.add_widget(header_label)
        self.layout.add_widget(header_layout)
        
        grid_header = GridLayout(cols=3, size_hint=(1, None), height=50)
        grid_header.add_widget(Label(text='ID Interno', size_hint_x=0.3, font_size='24sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='SKU', size_hint_x=0.4, font_size='24sp', bold=True, color=(1, 1, 1, 1)))
        grid_header.add_widget(Label(text='Cantidad', size_hint_x=0.3, font_size='24sp', bold=True, color=(1, 1, 1, 1)))
        self.layout.add_widget(grid_header)
        
        self.resultados = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.resultados.bind(minimum_height=self.resultados.setter('height'))
        scroll_view = ScrollView(size_hint=(1, 0.6))
        scroll_view.add_widget(self.resultados)
        self.layout.add_widget(scroll_view)
        
        self.close_button = Button(
            text='Cerrar Visualización',
            size_hint=(1, 0.1),
            font_size='22sp',
            bold=True,
            background_color=(0.1, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.close_button.bind(on_press=self.cerrar_visualizacion)
        self.layout.add_widget(self.close_button)
        
        self.add_widget(self.layout)
    
    def mostrar_pedido(self, pedidos_data):
        self.resultados.clear_widgets()
        for id_interno, items in pedidos_data:
            for item in items:
                item_id, sku, quantity, ubicaciones = item[:4]
                item_layout = GridLayout(cols=3, spacing=10, size_hint_y=None, height=50)
                item_layout.add_widget(Label(text=f'{id_interno}', size_hint_x=0.3, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{sku}', size_hint_x=0.4, font_size='24sp', color=(1, 1, 1, 1)))
                item_layout.add_widget(Label(text=f'{quantity}', size_hint_x=0.3, font_size='24sp', color=(1, 1, 1, 1)))
                self.resultados.add_widget(item_layout)

    def cerrar_visualizacion(self, instance):
        self.resultados.clear_widgets()
        app = App.get_running_app()
        app.screen_manager.current = 'scan_pedidos'
        app.scan_pedidos_screen.reset()

class MyApp(App):
    usuario = None
    
    def build(self):
        cargar_ubicaciones_sku()
        cargar_ids_sku()
        
        self.screen_manager = ScreenManager()
        
        self.login_screen = LoginScreen(name='login')
        self.scan_pedidos_screen = ScanPedidosScreen(name='scan_pedidos')
        self.presentacion_screen = PresentacionScreen(name='presentacion')
        self.view_pedidos_screen = ViewPedidosScreen(name='view_pedidos')
        
        self.screen_manager.add_widget(self.login_screen)
        self.screen_manager.add_widget(self.scan_pedidos_screen)
        self.screen_manager.add_widget(self.presentacion_screen)
        self.screen_manager.add_widget(self.view_pedidos_screen)
        
        return self.screen_manager

if __name__ == '__main__':
    MyApp().run()
