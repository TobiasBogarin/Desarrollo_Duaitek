import tkinter as tk
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time

class Credenciales:
    def __init__(self):
        self.credentials = None
        self.service = None
        self.service_account_info = {
            "type": "service_account",
            "project_id": "api-prueba-387019",
            "private_key_id": "523a7ab379d3d59382de0b4dd5215f50f7fd1370",
            "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDg6b66CIW2Q8/7\n6bAudowLEjxwHDvoSZluVTTbJx0oWkrjJTvWhrr595mNTl83S88lhIaIFHng8sTL\nEvsr/Zp1i6bAvgpiPJb8dm7g5IsH5ps+LL+5q0nJW6ZKb29W1/D1u3M9WiYLWH4P\njMehYxsmiHDB3pMRwoQfrJk06iF+xbO5HoWzUMvzuhmUeAbZ48/jbACrnoI+7Lou\n9risvi4AIjyBsxY162cVtyqAXmitXJ9QJAl4EoJoqSyAkl59aLcEi0UWUy0kuSD7\n7Pj3U+z3IyDWsOMld+lpv2cPdtij+/VtApaEYhklU8yYfzb8PG1CxkAFRrx2JP8x\n5IWiZSW3AgMBAAECggEAGfhOLOOdsjqzofONLlrt4hFhp9L6xTWftnJhlLSNf1f9\nsateShprLesocH042AUJjtwglJyYqMrKF7tsroWtTMlVSzLRDCAm5vvd4wXrWnwx\njMT/VmGwNsSTDPaVCpgLRgng9+1C21Mv1dJ8R+b4/u1ePQ8wOCrCYCM+hYegr9HO\nIF4cWMPxX5jF4y+MlsK+iy1s8lI/cmf5p5C0yd5ialtkycuSiMLwlz11UlcCGhf/\nTmose81uhl07F/EU4JELX/DGRJLUuLk4tBqiNdYfMztL2KvTsd1FVQ+PjS1Syc3D\nOEDm5P6eeIP6Fq2SPjUAiFTMmBpxRdlMSPq0Rqt6YQKBgQD3R4Pd2deByHviYfCJ\nW7Xu6z5QAmFxXgowhJPsqhqiHm7FqKKl4hqqEEFdRa50Az8stC7J4yEQmL9hzHTQ\nFWvVeSbozy6u0hyBtv4CZrzIFepZHMhW9KsBpm1sQOAliCZQa8vNr18C1nVLtKq8\nwFKcWt4lG1nmH5NyZb4ybHEXiQKBgQDo2E2Fw8SW+kxC1g7iS29OZk5jSj/qW7S6\n+23y31N+9MmSl9QZIDtS1dlU+6mxjeuaZsXbRDqBZzrpOSw163mUOMZqCOdrxs6k\nRrzRJ+kLWVhivA/eKLGZIbhCgDWtbq0JdKL0kOVKQSoO88OTnR6+4s9Judod46Qs\nnyr/BMvDPwKBgH+P1ObNSe8ZjU7rVzqEpQXrNOnxUHM7H+aHfgfIeJTJPjuZEs6g\nJUE1wYJsP+J5Ck31ZW2gTZ5SLeg1oMz3P/mP1hKjTmHA4hPIYqC6fwh4xbvSrUau\nUMk5IZmGnhq+cYVrFme04D6Gg1vah3l3fSZLee2KfoXIJDgPZF5+spiBAoGBALon\nLBs0PzhhFZUdo7qhinRIcIUK+Hx6IsyWdPmGOC+4rmrHfac00JjSJTW/GZS9HM5N\nOgOp0YhhKoUI02KsRoAMv/xH8BSHVe+aKhyhZrxPCs2tApafPBVsEu7/p2pnoGl9\n2UXjjZzG6kQX+JVMOSdtF0IfFtVsiHWwLuTBRdJrAoGAJ0BSCX/hNZxU7AqBi1gb\nqfHhotVoMVSimAyQT5vjHmdmyST8kFb3wn537jvrWc60PiCJ1nuk/IHXMLKjMByo\niYZ+MQsX2/59XJgx0Bpm75NshmeEkU4qh/4YbDCpYyiFZSXxsPRKwLeloOfX6oKT\n8VW0NaC0SXnsYYqamCTPnSA=\n-----END PRIVATE KEY-----\n",
            "client_email": "cerrador-1@api-prueba-387019.iam.gserviceaccount.com",
            "client_id": "117910477101184573911",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/v1/certs",
            "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/cerrador-1%40api-prueba-387019.iam.gserviceaccount.com",
            "universe_domain": "googleapis.com"
        }
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']

    def autenticar(self):
        try:
            self.credentials = service_account.Credentials.from_service_account_info(
                self.service_account_info,
                scopes=self.scopes
            )
            self.service = build('sheets', 'v4', credentials=self.credentials)
            print("Autenticación exitosa.")
        except Exception as e:
            print(f"Error de autenticación: {str(e)}")

class SpreadsheetDatos:
    def __init__(self, credenciales, spreadsheet_id):
        self.credenciales = credenciales
        self.spreadsheet_id = spreadsheet_id

    def cargar_usuarios(self):
        if not self.credenciales.service:
            print("No está autenticado.")
            return []

        range_name = 'usuarios!A:A'
        result = self.credenciales.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])
        if not values:
            print("No se encontraron usuarios en la hoja de cálculo.")
            return []
        else:
            usuarios = [row[0] for row in values if row]
            print("Usuarios cargados:")
            for usuario in usuarios:
                print(usuario)
            return usuarios

    def cargar_conceptos(self):
        if not self.credenciales.service:
            print("No está autenticado.")
            return []

        range_name = 'conceptos!A:B'
        result = self.credenciales.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])
        if not values:
            print("No se encontraron conceptos en la hoja de cálculo.")
            return []
        else:
            print("Conceptos cargados:")
            for id, concept in values:
                print(f"ID: {id}, Concepto: {concept}")
            return values

class LoginScreen:
    def __init__(self, root, usuarios, on_login):
        self.root = root
        self.usuarios = usuarios
        self.on_login = on_login

        self.root.title("Pantalla de Inicio de Sesión")
        self.setup_ui()

    def setup_ui(self):
        window_width = 300
        window_height = 200
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_position = (screen_width // 2) - (window_width // 2)
        y_position = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        self.label = tk.Label(self.root, text="Selecciona un usuario:")
        self.label.pack(pady=10)

        self.selected_user = tk.StringVar(self.root)
        self.selected_user.set(self.usuarios[0] if self.usuarios else "No hay usuarios")
        self.dropdown = tk.OptionMenu(self.root, self.selected_user, *self.usuarios)
        self.dropdown.pack(pady=10)

        self.login_button = tk.Button(self.root, text="Iniciar sesión", command=self.iniciar_sesion)
        self.login_button.pack(pady=20)

    def iniciar_sesion(self):
        usuario_seleccionado = self.selected_user.get()
        if usuario_seleccionado != "No hay usuarios":
            self.on_login(usuario_seleccionado)
            self.root.withdraw()
        else:
            print("No hay usuarios disponibles para iniciar sesión.")

class DropdownWithScrollbar(tk.Frame):
    def __init__(self, parent, options, on_select_callback, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.selected_option = tk.StringVar(value=options[0] if options else "No hay conceptos")
        self.options = options
        self.on_select_callback = on_select_callback

        self.button = tk.Button(self, textvariable=self.selected_option, command=self.toggle_listbox)
        self.button.pack(fill=tk.X)

        self.listbox_frame = tk.Frame(self)
        self.scrollbar = tk.Scrollbar(self.listbox_frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(self.listbox_frame, yscrollcommand=self.scrollbar.set, height=8, width=30)
        self.scrollbar.config(command=self.listbox.yview)

        for option in options:
            self.listbox.insert(tk.END, option)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_select)

    def toggle_listbox(self):
        if self.listbox_frame.winfo_ismapped():
            self.listbox_frame.pack_forget()
        else:
            self.listbox_frame.pack(fill=tk.X)

    def on_select(self, event):
        selected_index = self.listbox.curselection()
        if selected_index:
            self.selected_option.set(self.listbox.get(selected_index))
            self.toggle_listbox()
            self.on_select_callback()
            print(f"Concepto seleccionado: {self.selected_option.get()}")

class TIEMPOCONCEPTOS:
    def __init__(self):
        self.tiempos = []

    def registrar_tiempo(self, concepto, tiempo):
        self.tiempos.append((concepto, tiempo))
        print(f"Registrado: Concepto: {concepto}, Tiempo: {tiempo:.2f} segundos")

    def imprimir_tiempos(self):
        print("Tiempos registrados:")
        for concepto, tiempo in self.tiempos:
            print(f"Concepto: {concepto}, Tiempo: {tiempo:.2f} segundos")

class ESCANEO:
    def __init__(self, root, conceptos, usuario):
        self.root = root
        self.root.title("Interfaz de Escaneo")
        self.conceptos = conceptos
        self.tiempo_conceptos = TIEMPOCONCEPTOS()
        self.current_concept = None
        self.usuario = usuario
        self.setup_ui()
        self.start_time = time.time()
        self.ids_escaneados = []
        self.update_clock()

    def setup_ui(self):
        window_width = 400
        window_height = 300
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_position = (screen_width // 2) - (window_width // 2)
        y_position = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        self.timer_label = tk.Label(self.root, text="")
        self.timer_label.pack()

        self.entry = tk.Entry(self.root)
        self.entry.bind("<Return>", self.handle_scan)
        self.entry.pack()
        self.entry.focus()

        self.count_label = tk.Label(self.root, text="IDs escaneados: 0")
        self.count_label.pack()

        concept_titles = [concept for _, concept in self.conceptos]
        self.dropdown = DropdownWithScrollbar(self.root, concept_titles, self.on_concept_select)
        self.dropdown.pack()
        
        for index, (id, concept) in enumerate(self.conceptos):
            if id == '1':
                self.dropdown.listbox.selection_set(index)
                self.dropdown.selected_option.set(concept)
                self.current_concept = concept
                break
        self.start_time = time.time()  

        self.upload_button = tk.Button(self.root, text="Subir Datos", command=self.subir_datos)
        self.upload_button.pack(pady=10)

    def handle_scan(self, event):
        scanned_id = self.entry.get()
        if len(scanned_id) == 8 and scanned_id.isdigit():
            if self.current_concept != "PRESENTANDO":
                self.registrar_tiempo_actual()
            if self.dropdown.selected_option.get() != "PRESENTANDO":
                self.set_concept_to_presenting()
            self.registrar_id(scanned_id)
        else:
            print("El ID debe ser un número de 8 dígitos.")
        self.entry.delete(0, tk.END)
        self.entry.focus()


    def registrar_tiempo_actual(self):
        elapsed_time = time.time() - self.start_time
        self.tiempo_conceptos.registrar_tiempo(self.current_concept, elapsed_time)
        self.subir_estado(self.current_concept, elapsed_time)
        self.start_time = time.time()

    def subir_estado(self, concepto, tiempo):
        credenciales = Credenciales()
        credenciales.autenticar()

        if not credenciales.service:
            print("No se puede subir datos: no está autenticado.")
            return

        spreadsheet_id = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
        range_name = 'Temporal!A:D'

        concept_id = self.get_concept_id(concepto)
        formatted_time = time.strftime('%H:%M:%S', time.gmtime(tiempo))

        if formatted_time != "00:00:00":
            data = [['', concept_id, self.usuario, formatted_time]]
            body = {'values': data}

            try:
                response = credenciales.service.spreadsheets().values().append(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption='USER_ENTERED',
                    body=body
                ).execute()
                print(f"Estado subido: Concepto {concepto}, Tiempo: {formatted_time}.")
            except Exception as e:
                print(f"Error al subir estado: {str(e)}")
        else:
            print("No se suben datos: Tiempo registrado es 00:00:00.")

    def encontrar_primera_fila_vacia(self, credenciales, spreadsheet_id, range_name):
        try:
            result = credenciales.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            values = result.get('values', [])

            if not values:
                return 1  

            fila_vacia = len(values) + 1  
            for i, row in enumerate(values, start=1):
                if not any(row):  
                    fila_vacia = i
                    break
            return fila_vacia
        except Exception as e:
            print(f"Error al buscar fila vacía: {str(e)}")
            return None


    def set_concept_to_presenting(self):
        for index, (id, concept) in enumerate(self.conceptos):
            if concept == "PRESENTANDO":
                self.dropdown.listbox.selection_clear(0, tk.END)
                self.dropdown.listbox.selection_set(index)
                self.dropdown.selected_option.set(concept)
                self.on_concept_select()
                break

    def on_concept_select(self):
        new_concept = self.dropdown.selected_option.get()
        if new_concept != self.current_concept:
            self.change_concept(new_concept)

    def change_concept(self, new_concept):
        elapsed_time = time.time() - self.start_time
        if not self.ids_escaneados:
            self.subir_estado(self.current_concept, elapsed_time)
        else:
            self.subir_datos()
        self.current_concept = new_concept
        self.start_time = time.time()


    def registrar_id(self, id):
        current_concept = self.dropdown.selected_option.get()
        if id not in self.ids_escaneados:
            self.ids_escaneados.append(id)
            self.count_label.config(text=f"IDs escaneados: {len(self.ids_escaneados)}")
            print(f"ID registrado: {id} con el concepto: {current_concept}")
        else:
            print(f"ID duplicado (no registrado): {id}")

    def update_clock(self):
        elapsed_time = time.time() - self.start_time
        formatted_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
        self.timer_label.config(text=formatted_time)
        self.root.after(1000, self.update_clock)



    def subir_datos(self):
        credenciales = Credenciales()
        credenciales.autenticar()

        if not credenciales.service:
            print("No se puede subir datos: no está autenticado.")
            return

        spreadsheet_id = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
        range_name = 'Temporal'

        fila_vacia = self.encontrar_primera_fila_vacia(credenciales, spreadsheet_id, range_name + '!A:D')

        if fila_vacia is None:
            print("No se pudo encontrar una fila vacía adecuada.")
            return

        elapsed_time = format_elapsed_time(time.time() - self.start_time)
        
        # Crear la lista de datos a subir, aplicando tiempo solo al primer ID
        data = [
            [id_escaneado, self.get_concept_id(self.current_concept), self.usuario,
            elapsed_time if i == 0 else '00:00:00']
            for i, id_escaneado in enumerate(self.ids_escaneados)
        ]

        if not data:
            print("No hay datos para subir.")
            return

        body = {'values': data}
        try:
            # Ajustar el rango para el tamaño del data
            end_row = fila_vacia + len(data) - 1
            update_range = f"{range_name}!A{fila_vacia}:D{end_row}"

            response = credenciales.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=update_range,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Datos subidos exitosamente a las filas: {fila_vacia} a {end_row}")

            # Limpiar IDs escaneados y reiniciar contadores
            self.ids_escaneados.clear()
            self.count_label.config(text="IDs escaneados: 0")
            self.start_time = time.time()

            # Cambiar automáticamente al concepto con ID 26
            self.set_concept_by_id('26')

        except Exception as e:
            print(f"Error al subir datos: {str(e)}")



    def set_concept_by_id(self, concept_id):
        for index, (id, concept) in enumerate(self.conceptos):
            if id == concept_id:
                self.dropdown.listbox.selection_clear(0, tk.END)
                self.dropdown.listbox.selection_set(index)
                self.dropdown.selected_option.set(concept)
                self.on_concept_select()
                break



    def get_concept_id(self, concept_name):
        for id, concept in self.conceptos:
            if concept == concept_name:
                return id
        return None

def on_user_login(usuario, conceptos):
    global usuario_actual
    usuario_actual = usuario
    print(f"Usuario {usuario} ha iniciado sesión.")

    ventana_escaneo = tk.Toplevel()
    app_escaneo = ESCANEO(ventana_escaneo, conceptos, usuario)

def format_elapsed_time(elapsed_seconds):
    hours = int(elapsed_seconds // 3600)
    remaining_seconds = int(elapsed_seconds % 3600)
    minutes = int(remaining_seconds // 60)
    seconds = int(remaining_seconds % 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def main():
    credenciales = Credenciales()
    credenciales.autenticar()

    usuarios_spreadsheet_id = '1eKfJnr-R1H2EMJJh4NVsLCzp7umVguzetVradfvKASg'
    usuarios_datos = SpreadsheetDatos(credenciales, usuarios_spreadsheet_id)
    usuarios = usuarios_datos.cargar_usuarios()

    conceptos_spreadsheet_id = '1RRKrOiq6VuKtNx2MMWkBMBQ5H8KBuBs9MtKN-o1zH5c'
    conceptos_datos = SpreadsheetDatos(credenciales, conceptos_spreadsheet_id)
    conceptos = conceptos_datos.cargar_conceptos()

    root = tk.Tk()
    login_screen = LoginScreen(root, usuarios, lambda usuario: on_user_login(usuario, conceptos))
    root.mainloop()

if __name__ == "__main__":
    main()
