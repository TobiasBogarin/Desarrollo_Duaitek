import customtkinter as ctk
from google.oauth2 import service_account
from googleapiclient.discovery import build

# IDs de las hojas de cálculo
SPREADSHEET_ID_1 = '1ObCWgChmjd0FHSw1QVHf68txS6zxADA3v4XBjG72wrY'
SPREADSHEET_ID_2 = '15JFcSTheCBdDa4hIEzgduHqHJ5vkmyB4-_J12x4lIBo'
# Credenciales de la cuenta de servicio (asegúrate de tener la clave privada completa)
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
credentials = service_account.Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE)
service = build('sheets', 'v4', credentials=credentials)
# Inicializar la aplicación de Tkinter
ctk.set_appearance_mode("DARK")  # Set theme to match system theme, can be "light" or "dark"
ctk.set_default_color_theme("blue")  # Set color theme, you can choose others or create a custom one

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Generación de Etiquetas')
        self.geometry('600x400')

        # Configuración de la UI
        label = ctk.CTkLabel(self, text="Introduce los IDs (separados por comas):")
        label.pack(pady=10)
        self.entry = ctk.CTkEntry(self, width=300, placeholder_text="ID1, ID2, ID3, ...")
        self.entry.pack(pady=10)
        submit_button = ctk.CTkButton(self, text="Generar Individual", command=self.submit)
        submit_button.pack(pady=5)
        generate_massive_button = ctk.CTkButton(self, text="Generar Masivo", command=self.generate_massive)
        generate_massive_button.pack(pady=5)

    def submit(self):
        user_id = self.entry.get().strip()
        etiqueta_zpl = self.process_id(user_id)
        if etiqueta_zpl:
            with open(f"{user_id}.txt", "w") as archivo:
                archivo.write(etiqueta_zpl)
            print(f"Etiqueta para {user_id} guardada.")

    def generate_massive(self):
        ids = self.entry.get().split(',')
        with open("etiquetas_masivas.txt", "w") as archivo_masivo:
            for user_id in ids:
                user_id = user_id.strip()
                etiqueta_zpl = self.process_id(user_id)
                if etiqueta_zpl:
                    archivo_masivo.write(etiqueta_zpl + "\n\n")
            print("Todas las etiquetas se han guardado en 'etiquetas_masivas.txt'.")

    def process_id(self, user_id):
        print(f"Procesando ID: {user_id}")


        hojas_busqueda = ['presentacion y cierre', 'Despacho']
        nro_pedido, destinatario = self.buscar_en_hoja(SPREADSHEET_ID_2, hojas_busqueda, user_id)

        if not nro_pedido:
            nro_pedido, destinatario = self.buscar_en_hoja(SPREADSHEET_ID_1, hojas_busqueda, user_id)

        if nro_pedido:
            print(f"Nro Pedido: {nro_pedido}, Destinatario: {destinatario}")
            direccion, cp, zona, telefono, referencia, tipo_envio = self.buscar_despacho(SPREADSHEET_ID_2 if nro_pedido else SPREADSHEET_ID_1, user_id)

            if direccion:
                print(f"Dirección: {direccion}, CP: {cp}, Zona: {zona}, Teléfono: {telefono}, Referencia: {referencia}, Tipo de Envío: {tipo_envio}")
                self.generar_y_guardar_etiqueta(nro_pedido, direccion, cp, zona, user_id, referencia, destinatario, tipo_envio)
            else:
                print("Pedido no encontrado en la hoja de despacho.")
        else:
            print("ID no encontrado en ninguna hoja de cálculo.")

    def buscar_en_hoja(self, spreadsheet_id, hojas, user_id):
        for hoja in hojas:
            range_name = f'{hoja}!A:L'
            result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            for row in values:
                if row and row[0] == user_id:
                    nro_pedido = row[3]
                    destinatario = row[11]
                    return nro_pedido, destinatario
        return None, None

    def buscar_despacho(self, spreadsheet_id, user_id):
        range_name = 'despacho!A:K'
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        for row in values:
            if row and row[0] == user_id:
                direccion = row[3] if len(row) > 3 else "No adjuntó dirección"
                cp = row[4] if len(row) > 4 else "Sin CP"
                zona = row[6] if len(row) > 6 else "Sin zona"
                telefono = row[7] if len(row) > 7 and row[7] else "No adjuntó un teléfono"
                referencia = row[9] if len(row) > 9 else "Sin referencia"
                tipo_envio = row[10] if len(row) > 10 else "Tipo de envío no especificado"
                return direccion, cp, zona, telefono, referencia, tipo_envio
        return None, None, None, None, None, None

    def generar_y_guardar_etiqueta(self, nro_pedido, direccion, cp, zona, user_id, referencia, destinatario, tipo_envio):
        etiqueta_zpl = f"""
^XA
^FO270,25^A0N,44,44^FB660,2,0,L^FH^FDTODOMICRO^FS
^FO30,65^A0N,24,24^FB660,2,0,L^FH^FDSAN BLAS 2646^FS
^FO30,100^A0N,24,24^FB660,2,0,L^FH^FD CP 1416^FS
^FO30,135^A0N,44,44^FDVenta: {nro_pedido}^FS
^FX 1 Horizontal Line ^FS
^FO0,200^GB850,0,2^FS
^FO0,215^A0N,48,48^FB800,1,0,C^FR^FH^FDCORREO ARGENTINO^FS
^FX 2 Horizontal Line ^FS
^FO0,265^GB850,0,2^FS
^FO0,280^A0N,48,48^FB370,1,0,C^FDID:^FS
^FO350,265^GB2,67,2^FS
^FO369,280^A0N,43,43^FB420,1,0,C^FD{user_id}^FS
^FO0,330^GB850,0,2^FS
^FO50,431^BC^FD{user_id}^FS
^FO390,435^A0N,33,33^FB400,1,0,C^FH^FDCP: {cp}^FS
^FO389,485^A0N,48,48^FB400,2,0,C^FH^FD{zona}^FS
^FX 3 Horizontal Line ^FS
^FO0,700^GB850,0,2^FS
^FO0,715^A0N,48,48^FB800,1,0,C^FH^FD{tipo_envio}^FS  # Incluir "tipo de envío" en la etiqueta
^FX 3 Horizontal Line ^FS
^FO0,765^GB850,0,2^FS
^FO40,800^A0N,28,28^FB710,2,0,L^FH^FDDireccion: {direccion}^FS
^FO40,890^A0N,28,28^FB710,3,0,L^FH^FDReferencia: {referencia}^FS
^FO40,1010^A0N,28,28^FH^FDBarrio: {zona}^FS
^FO40,1090^A0N,28,28^FB720,2,0,L^FH^FDDestinatario: {destinatario}^FS
^XZ
"""
        with open("etiquetas_masivas.txt", "a") as archivo:
            archivo.write(etiqueta_zpl + "\n\n")

        print(f"Etiqueta para {user_id} agregada al archivo 'etiquetas_masivas.txt'.")
        
if __name__ == "__main__":
    app = App()
    app.mainloop()
