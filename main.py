import os
import csv
import pypandoc
from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askdirectory, asksaveasfilename
from tkinter import ttk

def main():
    root = Tk()
    root.withdraw()

    carpeta = askdirectory(title="Seleccione la carpeta con archivos .doc/.docx")

    if not carpeta:
        return

    mnum = simpledialog.askstring("Pregunta", "Cuál es el M.**? (Ejemplo: M.23)", parent=root)
    if not mnum:
        return

    # Elegir dónde guardar los archivos de salida
    csv_path = asksaveasfilename(
        title="Guardar CSV de salida",
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        initialfile="output.csv",
    )
    if not csv_path:
        return

    error_path = asksaveasfilename(
        title="Guardar archivo de errores",
        defaultextension=".txt",
        filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
        initialfile="errores.txt",
    )
    if not error_path:
        return

    rows = [
        [
            f'Informe No {mnum}.',
            'Fecha de Recibo',
            'Documento de identidad N°',
            'Paciente',
            'MUESTRA',
            'Edad',
            'Sexo',
            'Teléfono',
            'Institución',
            'DESCRIPCIÓN MACROSCÓPICA',
            'DESCRIPCIÓN MICROSCÓPICA',
            'DIAGNÓSTICO'
        ]
    ]
    order = [
        f'Informe No {mnum}.',
        'Fecha de Recibo:',
        'Documento de identidad N°:',
        'Paciente:',
        'MUESTRA:',
        'Edad:',
        'Sexo:',
        'Teléfono:',
        'Institución:',
        'DESCRIPCIÓN MACROSCÓPICA',
        'DESCRIPCIÓN MICROSCÓPICA',
        'DIAGNÓSTICO:'
    ]

    error_de_archivo = []

    # Preparar lista de archivos a procesar y barra de progreso
    archivos = [f for f in os.listdir(carpeta) if '.doc' in f]

    # Mostrar ventana con barra de progreso
    root.deiconify()
    root.title("Progreso de conversión Word → CSV")
    status_var = ttk.Label(root, text=f"0 / {len(archivos)} procesados")
    status_var.pack(padx=12, pady=(12, 6))
    file_var = ttk.Label(root, text="Listo")
    file_var.pack(padx=12, pady=(0, 6))
    progress = ttk.Progressbar(root, orient='horizontal', mode='determinate', maximum=max(len(archivos), 1))
    progress.pack(fill='x', padx=12, pady=(0, 12))
    root.update_idletasks()

    procesados = 0

    for file_name in archivos:
        file_var.config(text=f"Procesando: {file_name}")
        root.update_idletasks()

        try:
            output = pypandoc.convert_file(os.path.join(carpeta, file_name), 'plain', outputfile="tmp.txt")
        except Exception:
            error_de_archivo.append(file_name)
            procesados += 1
            progress['value'] = procesados
            status_var.config(text=f"{procesados} / {len(archivos)} procesados")
            root.update_idletasks()
            continue

        with open("tmp.txt", "r") as file:
            document = ""
            tokens = [f'Informe No {mnum}.', 'Fecha de Recibo:', 'Documento de identidad N°:', 'Paciente:', 'MUESTRA:']
            compound_tokens = [('Edad:', 'Sexo:', 'Teléfono:'), ('Institución:', 'Servicio:')]
            next_line_tokens = [
                ('DESCRIPCIÓN MACROSCÓPICA', 'DESCRIPCIÓN MICROSCÓPICA'),
                ('DESCRIPCIÓN MICROSCÓPICA', 'MÉTODO UTILIZADO:'),
                ('DIAGNÓSTICO:', '_____________________________________')
            ]

            row = {}
            text = ""

            def add_token(token, result):
                row[token] = result

            for line in file.readlines():
                text += line
                for token in tokens:
                    ind = line.find(token)
                    if ind != -1:
                        result = line[ind+len(token):].strip()
                        if token == f'Informe No {mnum}.' or token == 'Fecha de Recibo:':
                            result = result.replace('|', '').strip()
                        add_token(token, result)
                for token_list in compound_tokens:
                    token_ind = []
                    for token in token_list:
                        token_ind.append(line.find(token))
                    if -1 not in token_ind:
                        token_ind.append(len(line) + 1)
                        for i in range(len(token_ind) - 1):
                            token = token_list[i]
                            ind = token_ind[i]
                            nind = token_ind[i + 1]
                            result = line[ind + len(token):nind].strip()
                            add_token(token, result)
            for start_token, end_token in next_line_tokens:
                start_ind = text.find(start_token)
                end_ind = text.find(end_token, start_ind)
                result = text[start_ind + len(start_token):end_ind].strip().replace('\n', '')
                add_token(start_token, result)
            if "Servicio:" in row:
                del row["Servicio:"]
            cur_row = []
            try:
                for key in order:
                    cur_row.append(row[key])
                rows.append(cur_row)
            except Exception as e:
                error_de_archivo.append(file_name)

        # Actualizar progreso tras cada archivo
        procesados += 1
        progress['value'] = procesados
        status_var.config(text=f"{procesados} / {len(archivos)} procesados")
        root.update_idletasks()

    with open(csv_path, 'w', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter='|')
        writer.writerows(rows)

    with open(error_path, 'w') as file:
        file.write("\n".join(error_de_archivo))

    file_var.config(text="Finalizado")
    status_var.config(text=f"{procesados} / {len(archivos)} procesados")
    progress['value'] = progress['maximum']
    root.update_idletasks()

    messagebox.showinfo(
        "Conversión completada",
        f"CSV guardado en:\n{csv_path}\n\nErrores guardados en:\n{error_path}\n\nArchivos procesados: {procesados}",
        parent=root,
    )

    root.destroy()

main()
