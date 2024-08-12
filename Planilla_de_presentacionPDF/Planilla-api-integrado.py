import os
import PyPDF2
import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configura las claves de API de Google
json_creds = {
    "type": "service_account",
    "project_id": "api-prueba-387019",
    "private_key_id": "664f76f73954d810e0767b5b75e536f0a0923e64",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDynla6I6lPUTot\nyZANW6W1yUKXDVeJuDvDfFeOFNdfxpz/iko5MfM1thVdof7/+FfebSamJhgbKGXH\nB7o3jjnpUF8E0f8DfDuEf7o1cIXwZrpJwWVRaS2nd9yC6y5o0eownTfSsiv7pn32\nkanFlbGjqpBEqRiqIeDrlLnNh57ZEMVh/jRu/LgOrPqdtrddHBoijd+biczw7aJH\n4U8pDjFdk62kCoQ9JBYBquHaTHgpc59TSTAVmf4RCflsLe10/iMeZ8mfpordKfb7\nBHh8NhIR2qeaUR9rVqU2UfB8SvR+PpMWIs2jjgPP6c2DB3mm80rpF42zfAmrRV2X\nBhxj/39bAgMBAAECggEABb8rDg+oLqvINVYakREZqhD8AHidbogBn7jDQ3yfjrI1\nUnbjw9pjXg3+ZZBfv4+8rRUSyZ60tXf4foViq3bf958iu7TUVdshzG/7aNFN1Yrp\nYLPcWC2l74XkSumSg4MYCVIVWhVEUMVPIyxXbsmul9IuG/eRZlOVG0oTje28x9Fq\nZ6+ZSQbFTxsFgplOSayxJcvTQfIzt8/PPLmUBty1s6Q68Q0fYFD/qjuHetQ1GKlv\n83mSTM2GotfATpDRILZ/n5v9FrAII7vlYaLYdhTETzEydjgISvlxlAvhAhCUeYC2\nNbvMTHesIdTTo78wActTfFIa50S+UMq2nPKTNMWAoQKBgQD6Ilf/sNg6WEFvTkIp\nKeMf30R4l3JEnDTFXEZDJpYqt28eiPm2r+wBBcIM0Djpahe2ITKWzqS/K1o52MM7\ngfyX9jgeIvC45OwKmksqqmCsshLLKeatM/RUPCAjt21xJSHtbz6o04xXhyn3mtv0\nlvCaZgNy/GNjWqBbUwexTm2JMQKBgQD4TuAmdQWY1WyY10OYUWetY0jJ8lDB/kaR\n9jDGH7OvvKMdtYYKYxq7A2iTyOCcnyr+3mQlHd5OLF7Oa7URBEYJdmtA14CjAhNp\nZ3Kmce0wY4uPWyujzJBYiZftUjVuxnk17q6YwfBQLfw53AzGz5WixcVZUhAWekC6\n2/UoZvmuSwKBgQDtJ1K3ojvgVXz0wwYHcSdeOJj6nNxCILgHxwz27cbCiVhZYxUf\nGHxyG7t32pOa+nOwwpjsUs/wUHIjFllEOmH60f8y033YT3NcOh26Pf+avNsEtJ14\n6iFlG/x84JRrCgG41BhciPYupoArui+BHvrP6JislI7GzE3tSDOq7+j6gQKBgQC3\nuFRs3+S2QiNJquxehMy7I1y13s4V2veIA6nuzYH7owzlbGuyv8UFXe5Aej6GY9ZC\nIXXjaIgVOwsim9qqrojLc4zDuy94bI7ETEAuGtkuFlkqRoCxfyfF+ngopczXG46P\ncvxIFiaijIO0o7XoW6sRdlcgUXGJ0AaYuypXLGnMpQKBgC5KptrIYO8g2Wwil/Hc\nQt22kKnvWxitq//RkqAID67jEQh5178VMoJF6juwflP0uyL1oeB0OkVWFhbf+8WA\nI8Kmke5AaVbzJDF28pAOMu1XHaejbyLo92c/PpiBFhFKShbdurI0ZwTaYuIDD4FQ\nX23X0Mr6e/gUAlO+DR8dspT2\n-----END PRIVATE KEY-----\n",
    "client_email": "planilla@api-prueba-387019.iam.gserviceaccount.com",
    "client_email": "107956038313503016079",
    "client_id": "your-client-id",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/planilla%40api-prueba-387019.iam.gserviceaccount.com"
}

# Configura el ámbito y las credenciales
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(json_creds, scope)

# Autentica el cliente de gspread
client = gspread.authorize(creds)

# Autentica el cliente de gspread
client = gspread.authorize(creds)

def generar_excel():
    try:
        # Abre la hoja de cálculo por su nombre
        spreadsheet = client.open('planilla')
        # Abre una hoja de trabajo específica
        worksheet = spreadsheet.worksheet('log5')

        # Lee el archivo PDF y extrae el texto
        pdf_file = open('CONTROL.pdf', 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        pdf_file.close()

        # Procesa el texto para obtener los datos de ventas
        ventas = []
        for line in text.split("\n"):
            if line.strip().startswith("Venta"):
                venta = line.strip().split()[-1]
            elif line.strip().startswith("SKU"):
                sku = line.strip().split()[-1]
            elif line.strip().startswith("Cantidad"):
                cantidad_text = line.strip().split()[-1]
                if re.match(r'\d+', cantidad_text):
                    cantidad = int(cantidad_text)
                else:
                    numbers = re.findall(r'\d+', line)
                    cantidad = int(numbers[-1]) if numbers else 0
                ventas.append([venta, sku, cantidad, "-"])
            elif "Color" in line.strip():
                color = line.strip().split(":")[-1].strip()
                if ventas:
                    ventas[-1][-1] = color

        # Define el ID inicial basado en la entrada del usuario
        id_inicial = int(entry.get())

        # Genera IDs para cada venta
        id_actual = id_inicial
        id_venta = {}
        for venta in ventas:
            if venta[0] not in id_venta:
                id_venta[venta[0]] = id_actual
                id_actual += 1
            venta.insert(0, id_venta[venta[0]])

        # Crea un DataFrame con los datos extraídos
        df_ventas = pd.DataFrame(ventas, columns=["ID", "Venta", "SKU", "Cantidad", "Color"])

        # Genera la columna 'URL' considerando carritos con el mismo número de venta
        base_url = "https://www.mercadolibre.com.ar/ventas/{}/detalle?callbackUrl=https%3A%2F%2Fwww.mercadolibre.com.ar%2Fventas%2Fomni%2Flistado%3Fplatform.id%3DML%26channel%3Dmarketplace%26filters%3D%26sort%3D%26page%3D1%26search%3D43160654561%26startPeriod%3DWITH_DATE_CLOSED_6M_OLD%26toCurrent%3D%26fromCurrent%3D"
        df_ventas['URL'] = df_ventas['Venta'].apply(lambda x: base_url.format(x))

        # Verifica y elimina URLs duplicados, dejando la primera ocurrencia
        df_ventas['URL'] = df_ventas['URL'].mask(df_ventas['URL'].duplicated(), '')


        # Invierte el orden de las filas en el DataFrame
        df_ventas = df_ventas.iloc[::-1].reset_index(drop=True)

        # Guarda el DataFrame en un archivo Excel
        nombre_xlsx = f"{id_inicial}_CONTROL.xlsx"
        df_ventas.to_excel(nombre_xlsx, index=False)

        # Actualiza la hoja de cálculo con los nuevos datos
        values = df_ventas.values.tolist()
        num_rows, num_cols = df_ventas.shape
        cell_range = f'A2:{chr(64 + num_cols)}{num_rows + 1}'
        worksheet.update(cell_range, values, value_input_option='USER_ENTERED')


        # Muestra un mensaje de éxito
        messagebox.showinfo("Excel generado", f"Archivo {nombre_xlsx} creado.")
        ventana.destroy()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Inicia la interfaz gráfica
ventana = tk.Tk()
ventana.geometry("880x440")
ventana.title("Generador Planilla de presentación")

label = tk.Label(ventana, text="Ingrese el ID inicial:")
label.config(font=("Arial", 24))
label.pack(pady=50)

entry = tk.Entry(ventana)
entry.pack()

boton_generar = tk.Button(ventana, text="Generar Planilla de Presentación", command=generar_excel)
boton_generar.pack(pady=20)

ventana.mainloop()
