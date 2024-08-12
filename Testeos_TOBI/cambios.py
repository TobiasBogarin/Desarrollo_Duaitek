import win32print
import win32ui
from tkinter import Tk, Label, Entry, Button, Text, Checkbutton, IntVar, Frame
import datetime

def create_single_label(id_value, tipo_envio):
    y_offset = 0

    parallel_label = f"^XA\n" \
                    f"^FO35,{25 + y_offset}^A0N,34,34^FDID: {id_value}^FS\n" \
                    f"^BY2,3,55\n" \
                    f"^FO280,{52 + y_offset}^BC^FD{id_value}^FS\n" \
                    f"^FO30,{80 + y_offset}^A0N,25,25^FB200,2,0,C^FH^FD{tipo_envio}^FS\n" \
                    f"^FO30,{150 + y_offset}^GB750,1,3^FS\n" \
                    f"^XZ\n"
    return parallel_label.strip()

def generate_and_store_label():
    global all_labels, label_count
    id_value = entry.get()
    if id_value:
        tipo_envio = "CAMBIO" if tipo_envio_var.get() == 1 else "ENTREGA FALTANTE"
        label = create_single_label(id_value, tipo_envio)
        all_labels += label + '\n'
        label_count += 1
        text.delete(1.0, "end")
        text.insert("end", all_labels)
        counter_label.config(text=str(label_count))
        entry.delete(0, "end")

def print_and_clear_labels():
    global all_labels, label_count
    # Guardar en un archivo
    filename = f"etiquetas_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(filename, 'w') as f:
        f.write(all_labels)

    # Enviar a la impresora específica
    printer_name = "honeywell"
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Etiquetas", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, all_labels.encode('utf-8'))
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

    # Limpia las etiquetas y el contador
    all_labels = ""
    label_count = 0
    text.delete(1.0, "end")
    counter_label.config(text=str(label_count))

root = Tk()
root.title("Generador de Etiquetas")

Label(root, text="Ingrese el ID:").pack()
Label(root, text="Total de etiquetas generadas:").pack()

counter_label = Label(root, text="0")
counter_label.pack()

entry = Entry(root)
entry.pack()

# Frame para contener los checkboxes
checkbox_frame = Frame(root)
checkbox_frame.pack()

tipo_envio_var = IntVar(value=1)  # Predeterminado a CAMBIO
Checkbutton(checkbox_frame, text="CAMBIO", variable=tipo_envio_var, onvalue=1, offvalue=0).pack(side="left")
Checkbutton(checkbox_frame, text="ENTREGA FALTANTE", variable=tipo_envio_var, onvalue=0, offvalue=1).pack(side="left")

all_labels = ""
label_count = 0

text = Text(root, height=15, width=80)
text.pack()

entry.bind("<Return>", lambda event: generate_and_store_label())

Button(root, text="Generar Etiqueta", command=generate_and_store_label).pack()
Button(root, text="Imprimir Etiquetas", command=print_and_clear_labels).pack()

root.mainloop()
