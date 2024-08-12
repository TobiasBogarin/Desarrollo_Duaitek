import re
import csv
import requests
import os
import xlwt
from urllib.parse import unquote
import datetime
from tkinter import Tk, filedialog

def obtener_credenciales():
    print("Seleccione la cuenta con la que desea trabajar:")
    print("1: DK")
    print("2: TM")
    opcion = input("Ingrese el número de la cuenta: ")
    if opcion == "1":
        return ('8545932337071872', 'VtRNp52gxGKCDKLvR8YRfovPLL3ZBfIk')
    elif opcion == "2":
        return ('8595972831111833', 'fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L')
    else:
        print("Opción no válida. Usando credenciales por defecto TM.")
        return ('8595972831111833', 'fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L')


def decode_address(encoded_address):
    """Decodifica la dirección codificada en UTF-8."""
    decoded_address = unquote(encoded_address.replace('_', '%'))
    return decoded_address

def extract_info(content):
    lines = content.split("\n")
    info_list = []
    info = {}

    for line in lines:
        line = line.strip()

        if line.startswith("^XA"):
            if info:
                info_list.append(info)
                info = {}
        elif line.startswith("^FO"):
            parts = line.split("^FD")

            if len(parts) == 2:
                key = parts[0].strip("^FO").strip()
                value = parts[1].strip("^FS").strip()
                info[key] = value

    if info:
        info_list.append(info)

    return info_list

def save_to_excel(info_list):
    workbook = xlwt.Workbook()

    for index, info in enumerate(info_list):
        sheet = workbook.add_sheet(f"Etiqueta {index + 1}")

        row = 0
        for key, value in info.items():
            sheet.write(row, 0, key)
            sheet.write(row, 1, value)
            row += 1

    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"etiqueta_{current_datetime}.xls"
    workbook.save(file_name)

    print(f"La información se ha extraído y guardado en '{file_name}'.")

def modify_labels(content, initial_id):
    modified_labels = ""
    label_list = content.replace("^XA^MCY^XZ", "").split("^XZ")

    id_counter = int(initial_id)
    
    for i, label in enumerate(label_list):
        if label.strip() == '':
            continue

#modifica las etiquetas
    for label in label_list:
        modified_label = label.replace(
            "^FO120,22^A0N,24,24^FH", "^FO120,22^A0N,24,24^FH"
        ).replace(
            "^FO120,65^A0N,24,24^FB660,2,0,L^FH", ""
        ).replace(
            "^FO120,100^A0N,24,24^FB660,2,0,L^FH", ""
        ).replace(
            "^FO120,135^A0N,24,24^", ""
        ).replace(
            "^FO272,132^A0N,27,27", ""
        ).replace(
            "^FO250,132^A0N,27,27^FD", ""
        ).replace(
            "^FO500,135^A0N,24,24^FH", "^FO60,105^A0N,24,24^FH"
        ).replace(
            "^FO652,132^A0N,27,27", "^FO209,105^A0N,24,24"
        ).replace(
            "^FO0,215^A0N,48,48^FB800,1,0,C^FR^FH", ""
        ).replace(
            "^FO0,280^A0N,48,48^FB370,1,0,C", ""
        ).replace(
            "^FO369,280^A0N,43,43^FB420,1,0,C", ""
        ).replace(
            "^FO370,281^A0N,43,43^FB420,1,0,C", ""
        ).replace(
            "^FO80,370^BY4,4,0^BQN,2,7", "^FO60,405^BY4,4,0^BQN,2,7"
        ).replace(
            "^FO390,435^A0N,33,33^FB400,1,0,C^FH", "^FO390,435^A0N,33,33^FB400,1,0,C^FH"
        ).replace(
            "^FO389,485^A0N,48,48^FB400,2,0,C^FH", "^FO389,485^A0N,48,48^FB400,2,0,C^FH"
        ).replace(
            "^FO390,486^A0N,48,48^FB400,2,0,C^FH", "^FO390,486^A0N,48,48^FB400,2,0,C^FH"
        ).replace(
            "^FO389,580^A0N,45,45^FB400,2,0,C^FH", "^FO389,580^A0N,45,45^FB400,2,0,C^FH"
        ).replace(
            "^FO390,581^A0N,45,45^FB400,2,0,C^FH", "^FO390,581^A0N,45,45^FB400,2,0,C^FH"
        ).replace(
            "^FO0,715^A0N,48,48^FB800,1,0,C^FH", ""
        ).replace(
            "^FO40,800^A0N,28,28^FB710,2,0,L^FH", "^FO60,140^A0N,24,24^FB710,2,0,L^FH"
        ).replace(
            "^FO39,800^A0N,28,28^FH", ""
        ).replace(
            "^FO40,890^A0N,28,28^FB710,3,0,L^FH", "^FO60,176^A0N,24,24^FB710,3,0,L^FH"
        ).replace(
            "^FO39,890^A0N,28,28^FH", ""
        ).replace(
            "^FO40,1010^A0N,28,28^FH", ""
        ).replace(
            "^FO39,1010^A0N,28,28^FH", ""
        ).replace(
            "^FO40,1090^A0N,28,28^FB720,2,0,L^FH", "^FO60,300^A0N,24,24^FB720,2,0,L^FH"
        ).replace(
            "^FO39,1090^A0N,28,28^FH", ""
        )

        modified_label_lines = modified_label.split("\n")
        new_label = ""
        skip_next_line = False
        for line in modified_label_lines:
            if "Horizontal Line" not in line:
                if not skip_next_line:
                    new_label += line + "\n"
                skip_next_line = False
            else:
                skip_next_line = True

        # Añade el ID solo si la etiqueta no está vacía y es una etiqueta válida
        if new_label.strip() != '':
            start_index = new_label.find("^XA") + len("^XA")
            new_label = new_label[:start_index] + "\n^FO445,715^A0N,54,54^FDID: " + str(id_counter) + "^FS" + new_label[start_index:]
            id_counter += 1  # Incrementa el ID solo si se ha añadido a una etiqueta

        modified_labels += new_label + "^XZ"

    # Verificación final para eliminar un posible ID suelto al final
    if modified_labels.endswith("^FO445,715^A0N,54,54^FDID: " + str(id_counter) + "^FS^XZ"):
        modified_labels = modified_labels.rsplit("^FO445,715^A0N,54,54^FDID: " + str(id_counter) + "^FS", 1)[0] + "^XZ"

    # Elimina el ^XZ adicional al final si existe
    if modified_labels.endswith("^XZ"):
        modified_labels = modified_labels[:-3]

    return modified_labels

def select_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Archivos de texto", "*.txt")])
    return file_path

# Definición de la variable global en el nivel más alto del script
global trackings
trackings = []



# Funciones para interactuar con la API de Mercado Libre
def obtener_token(client_id, client_secret):
    url = "https://api.mercadolibre.com/oauth/token"
    payload = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret
    }
    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        return response.json()['access_token']
    except requests.RequestException as e:
        print(f"Error al obtener token: {e}")
        return None


def dividir_en_lotes(lista, tamano_lote):
    """ Divide la lista dada en sub-listas de tamaño especificado. """
    for i in range(0, len(lista), tamano_lote):
        yield lista[i:i + tamano_lote]


def obtener_datos_de_envios(access_token, shipment_ids):
    datos_envio = {}
    lotes = dividir_en_lotes(shipment_ids, 50)
    for lote in lotes:
        for shipment_id in lote:
            url = f"https://api.mercadolibre.com/shipments/{shipment_id}"
            headers = {'Authorization': f'Bearer {access_token}'}
            try:
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                datos = response.json()
                receiver_address = datos.get('receiver_address', {})
                datos_envio[shipment_id] = (receiver_address.get('latitude', ''), receiver_address.get('longitude', ''))
            except requests.RequestException as e:
                print(f"Error al obtener datos de envío para {shipment_id}: {e}")
                datos_envio[shipment_id] = ('', '')  # Asignar valores vacíos en caso de error
    return datos_envio

def main():
    global trackings  # Declaración para utilizar la variable global dentro de esta función

    # Obtiene las credenciales según la elección del usuario
    client_id, client_secret = obtener_credenciales()
    
    file_name = "Etiqueta de envio.txt"
    file_path = os.path.join(os.getcwd(), file_name)

    if not os.path.isfile(file_path):
        print(f"No se encontró el archivo '{file_name}'.")
        return

    with open(file_path, "r") as file:
        content = file.read()

    initial_id = int(input("Ingrese el ID inicial para las etiquetas: "))

    # Modifica las etiquetas y obtiene las etiquetas modificadas
    modified_labels = modify_labels(content, initial_id)

    # Calcula el ID final basado en el número de etiquetas procesadas
    num_labels = len(modified_labels.split("^XZ")) - 1  # Asume que cada etiqueta termina con "^XZ"
    final_id = initial_id + num_labels - 1

    # Define el nombre del archivo usando el ID inicial y el ID final
    modified_file_name = f"{initial_id}-{final_id}.txt"
    
    # Guarda las etiquetas modificadas en el nuevo archivo
    with open(modified_file_name, 'w') as modified_file:
        modified_file.write(modified_labels)
    
    print(f"Etiqueta modificada guardada en el archivo: {modified_file_name}")

    # Extracción de los códigos de seguimiento y otros datos
    tracking_regex = r'\^FDLA,\{"id":"(\d+)"'
    trackings = re.findall(tracking_regex, content)

    destinatario_regex = r'\^FDDestinatario: (.*?)\^FS'
    direccion_regex = r'\^FDDireccion: (.*?)\^FS'
    cp_regex = r'\^FDCP: (.*?)\^FS'
    referencia_regex = r'\^FDReferencia: (.*?)\^FS'

    destinatarios = re.findall(destinatario_regex, content)
    direcciones = [decode_address(addr) for addr in re.findall(direccion_regex, content)]
    cps = re.findall(cp_regex, content)
    referencias = re.findall(referencia_regex, content)

    zpl_codes = modified_labels.split("^XZ")
    if not zpl_codes[-1].strip():
        zpl_codes.pop()

    # Obtiene token y realiza la llamada a la API para obtener datos en lotes
    access_token = obtener_token(client_id, client_secret)
    if access_token:
        datos_de_envios = obtener_datos_de_envios(access_token, trackings)

        csv_file_name = modified_file_name.replace(".txt", ".csv")
        with open(csv_file_name, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["ZPL", "ID", "TRACKING", "Destinatarios", "Direcciones", "CPs", "Referencias", "Latitud", "Longitud", "Tipo de Envío"])

            num_rows = len(trackings)
            id_counter = initial_id + num_rows - 1

            for i in range(num_rows):
                current_id = id_counter - i
                zpl_code = zpl_codes[num_rows - 1 - i] + "^XZ"
                zpl_code = zpl_code.strip()
                zpl_code = zpl_code.replace("\n", " ")

                tracking = trackings[num_rows - 1 - i] if i < len(trackings) else ""
                destinatario = destinatarios[num_rows - 1 - i] if i < len(destinatarios) else ""
                direccion = direcciones[num_rows - 1 - i] if i < len(direcciones) else ""
                cp = cps[num_rows - 1 - i] if i < len(cps) else ""
                referencia = referencias[num_rows - 1 - i] if i < len(referencias) else ""
                latitud, longitud = datos_de_envios.get(tracking, ('', ''))

                # Determinar el tipo de envío basado en la presencia de un marcador en la etiqueta
                tipo_envio = "COMERCIAL" if "^FDCOMERCIAL^FS" in zpl_code else "RESIDENCIAL"

                writer.writerow([zpl_code, current_id, tracking, destinatario, direccion, cp, referencia, latitud, longitud, tipo_envio])

        print(f"Información extraída y guardada en '{csv_file_name}'.")
    else:
        print("No se pudo obtener el token de acceso.")



    # Genera las etiquetas paralelas
    parallel_labels = create_parallel_labels(modified_labels)

    # Calcula el último ID utilizado
    final_id = initial_id + len(modified_labels.split("^XZ")) - 2  # Ajusta según cómo estés incrementando id_counter

    # Construye el nombre del archivo de etiquetas paralelas
    parallel_file_name = f"modificada_{initial_id}-{final_id}.txt"

    # Guarda las etiquetas paralelas en el nuevo archivo
    with open(parallel_file_name, 'w') as parallel_file:
        parallel_file.write(parallel_labels)

    print(f"Etiquetas paralelas guardadas en el archivo: {parallel_file_name}")


def create_parallel_labels(modified_labels):
    parallel_labels = ""
    label_list = [label for label in modified_labels.split("^XZ") if label.strip() != '']

    for index, label in enumerate(label_list):
        parallel_labels += "^XA\n"  # Comienza una nueva etiqueta

        id_match = re.search(r"\^FDID: (\d+)\^FS", label)
        tipo_envio = "COMERCIAL" if "^FDCOMERCIAL^FS" in label else "RESIDENCIAL"
        valor_a_extraer_match = re.search(r"\^FO389,485\^A0N,48,48\^FB400,2,0,C\^FH\^FD(.+?)\^FS", label)

        if id_match and valor_a_extraer_match:
            id_value = id_match.group(1)
            valor_a_extraer = valor_a_extraer_match.group(1)

            # Construye la etiqueta con el valor extraído
            parallel_label = f"^FO35,25^A0N,34,34^FDID: {id_value}^FS\n" \
                            f"^BY2,3,55\n" \
                            f"^FO280,52^BC^FD{id_value}^FS\n" \
                            f"^FO30,80^A0N,25,25^FB200,2,0,C^FH^FD{tipo_envio}^FS\n" \
                            f"^FO30,116^A0N,25,25^FB200,2,0,C^FH^FD{valor_a_extraer}^FS\n" \
                            f"^FO30,150^GB750,1,3^FS\n" \
                            "^XZ\n"  # Finaliza la etiqueta actual

        parallel_labels += parallel_label

    return parallel_labels.strip()



if __name__ == "__main__":
    main()
# Eliminar el archivo de Excel generado
for file_name in os.listdir():
    if file_name.startswith("etiqueta_") and file_name.endswith(".xls"):
        os.remove(file_name)

#En caso de seleccionar DK debe utilizar esta credencial
#Client_id:8545932337071872
#Cliente_secret:VtRNp52gxGKCDKLvR8YRfovPLL3ZBfIk

#en caso de seleccionar Tm debe utilizar esta credencial
#Client_id:8595972831111833
#Cliente_secret:fKyvOSCw3vhI2fDKqroEY8zdcd77ZN5L