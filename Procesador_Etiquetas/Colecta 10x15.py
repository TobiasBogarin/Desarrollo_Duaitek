import re
import csv
import os
import xlwt
from urllib.parse import unquote
import datetime
from tkinter import Tk, filedialog


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
            "^FO120,65^A0N,24,24^FB660,2,0,L^FH", "FO120,65^A0N,24,24^FB660,2,0,L^FH"
        ).replace(
            "^FO120,100^A0N,24,24^FB660,2,0,L^FH", "^FO120,100^A0N,24,24^FB660,2,0,L^FH"
        ).replace(
            "^FO120,135^A0N,24,24^", "^FO120,135^A0N,24,24^"
        ).replace(
            "^FO272,132^A0N,27,27", "^FO272,132^A0N,27,27"
        ).replace(
            "^FO250,132^A0N,27,27^FD", "^FO250,132^A0N,27,27^FD"
        ).replace(
            "^FO500,135^A0N,24,24^FH", "^FO60,105^A0N,24,24^FH"
        ).replace(
            "^FO652,132^A0N,27,27", "^FO209,105^A0N,24,24"
        ).replace(
            "^FO0,215^A0N,48,48^FB800,1,0,C^FR^FH", "FO0,215^A0N,48,48^FB800,1,0,C^FR^FH"
        ).replace(
            "^FO0,280^A0N,48,48^FB370,1,0,C", "FO0,280^A0N,48,48^FB370,1,0,C"
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
            new_label = new_label[:start_index] + "\n^FO505,115^A0N,54,54^FDID: " + str(id_counter) + "^FS" + new_label[start_index:]
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

def obtener_valor_lista(lista, indice, valor_default="Desconocido"):
    """Devuelve el valor en el índice dado de la lista o un valor predeterminado si el índice está fuera de rango."""
    try:
        return lista[indice]
    except IndexError:
        return valor_default


def main():
    file_name = "Etiqueta de envio.txt"
    file_path = os.path.join(os.getcwd(), file_name)

    if not os.path.isfile(file_path):
        print(f"No se encontró el archivo '{file_name}'.")
        return

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
    except UnicodeDecodeError:
        # Fallback a latin-1 si utf-8 falla
        with open(file_path, "r", encoding="latin-1") as file:
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

    # Continúa con la extracción de los códigos de seguimiento y otros datos...
    tracking_regex = r'\^FDLA,\{.*?"id":"(\d+)".*?\}\^FS'
    trackings = re.findall(tracking_regex, content)

    # Continúa con las expresiones regulares para extraer destinatario, dirección, CP y referencia...
    destinatario_regex = r'\^FO30,970\^A0N,26,26\^FB600,2,0,L\^FH\^FD[^)]+\(([^)]+)\)\^FS'

    destinatarios = re.findall(destinatario_regex, content)


        # Divide el contenido modificado para obtener los códigos ZPL individuales
    zpl_codes = modified_labels.split("^XZ")
    # Elimina el último elemento si está vacío (debido al split)
    if not zpl_codes[-1].strip():
        zpl_codes.pop()

    csv_file_name = modified_file_name.replace(".txt", ".csv")
    with open(csv_file_name, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["ZPL", "ID", "TRACKING", "Destinatarios"])

        num_rows = len(trackings)
        id_counter = initial_id + num_rows - 1

        for i in range(num_rows):
            current_id = id_counter - i
            zpl_code = zpl_codes[num_rows - 1 - i] + "^XZ"  # Añade ^XZ al final de cada código ZPL
            zpl_code = zpl_code.strip().replace("\n", " ")  # Formatea el código ZPL

            # Utiliza la función auxiliar para obtener el tracking y el destinatario de forma segura
            tracking = obtener_valor_lista(trackings, num_rows - 1 - i, "Sin Tracking")
            destinatario = obtener_valor_lista(destinatarios, num_rows - 1 - i, "Desconocido")

            writer.writerow([zpl_code, current_id, tracking, destinatario])

            print(f"Información extraída y guardada en '{csv_file_name}'.")

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
    label_list = [label for label in modified_labels.split("^XZ") if label.strip()]

    for label in label_list:
        parallel_labels += "^XA\n"  # Comienza una nueva etiqueta

        id_match = re.search(r"\^FDID: (\d+)\^FS", label)

        if id_match:
            id_value = id_match.group(1)

            # Construye la etiqueta con las nuevas posiciones y el tipo de envío identificado
            parallel_label = f"^FO30,15^A0N,34,34^FDID: {id_value}^FS\n" \
                            f"^FO178,132^A0N,32,32^FB460,1,0,C^FR^FDDESPACHAR COLECTA^FS\n" \
                            f"^BY2,3,50\n^FO285,51^BC^FD{id_value}^FS\n" \
                            f"^FO30,165^GB750,3,3^FS\n"  # Agrega la línea divisoria

            parallel_labels += parallel_label
            parallel_labels += "^XZ\n"  # Finaliza la etiqueta actual

    return parallel_labels.strip()

if __name__ == "__main__":
    main()

# Eliminar el archivo de Excel generado
for file_name in os.listdir():
    if file_name.startswith("etiqueta_") and file_name.endswith(".xls"):
        os.remove(file_name)
