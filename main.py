import os
import sys
import csv
import time
import traceback
from tkinter import Tk, simpledialog, messagebox, StringVar
from tkinter.filedialog import askdirectory, asksaveasfilename
from tkinter import ttk

def set_pandoc_env():
    if os.environ.get("PYPANDOC_PANDOC"):
        return
    exe = "pandoc.exe" if os.name == "nt" else "pandoc"
    candidates = []

    # PyInstaller onefile unpack dir
    base = getattr(sys, "_MEIPASS", None)
    if base:
        candidates.append(os.path.join(base, "pandoc", exe))

    # PyInstaller onedir / macOS .app layout
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)  # .../Contents/MacOS
        candidates.append(os.path.join(exe_dir, "pandoc", exe))
        # Sometimes resources go here:
        candidates.append(os.path.join(os.path.dirname(exe_dir), "Resources", "pandoc", exe))

    # Dev fallbacks (don’t rely on these in production, but handy)
    candidates += ["/opt/homebrew/bin/pandoc", "/usr/local/bin/pandoc", "/usr/bin/pandoc"]

    for c in candidates:
        if os.path.exists(c) and os.access(c, os.X_OK):
            os.environ["PYPANDOC_PANDOC"] = c
            break

def _ellipsize_middle(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    if max_chars <= 3:
        return text[:max_chars]
    left = (max_chars - 1) // 2
    right = max_chars - 1 - left
    return text[:left] + '…' + text[-right:]


def main():
    set_pandoc_env()
    import pypandoc  # import after env is set

    root = Tk()
    root.withdraw()

    carpeta = askdirectory(title="Seleccione la carpeta con archivos .doc/.docx")
    if not carpeta:
        return

    mnum = simpledialog.askstring("Pregunta", "Cuál es el M.**? (Ejemplo: M.23)", parent=root)
    if not mnum:
        return

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

    rows = [[
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
    ]]
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

    archivos = [f for f in os.listdir(carpeta) if '.doc' in f]  # matches .doc and .docx

    # UI progreso
    root.deiconify()
    root.title("Progreso de conversión Word → CSV")
    status_text = StringVar(value=f"0 / {len(archivos)} procesados")
    status_var = ttk.Label(root, textvariable=status_text, width=30, anchor='w')
    status_var.pack(padx=12, pady=(12, 6))
    file_text = StringVar(value="Listo")
    file_var = ttk.Label(root, textvariable=file_text, width=60, anchor='w')
    file_var.pack(padx=12, pady=(0, 6))
    progress = ttk.Progressbar(root, orient='horizontal', mode='determinate', maximum=max(len(archivos), 1))
    progress.pack(fill='x', padx=12, pady=(0, 12))
    root.update_idletasks()
    # Freeze window size to prevent geometry jumps due to label text changes
    root.minsize(root.winfo_width(), root.winfo_height())
    root.resizable(False, False)

    procesados = 0
    last_ui = time.monotonic()

    for file_name in archivos:
        display_name = _ellipsize_middle(file_name, 56)
        file_text.set(f"Procesando: {display_name}")
        now = time.monotonic()
        if now - last_ui > 0.12:
            root.update_idletasks()
            last_ui = now

        path_in = os.path.join(carpeta, file_name)
        try:
            # Get plain text directly (avoid temp files / permissions)
            text = pypandoc.convert_file(path_in, 'plain')
        except Exception as e:
            err = traceback.format_exc()
            error_de_archivo.append(f"{file_name} :: conversión fallida\n{err}")
            procesados += 1
            progress['value'] = procesados
            status_text.set(f"{procesados} / {len(archivos)} procesados")
            now = time.monotonic()
            if now - last_ui > 0.12:
                root.update_idletasks()
                last_ui = now
            continue

        tokens = [f'Informe No {mnum}.', 'Fecha de Recibo:', 'Documento de identidad N°:', 'Paciente:', 'MUESTRA:']
        compound_tokens = [('Edad:', 'Sexo:', 'Teléfono:'), ('Institución:', 'Servicio:')]
        next_line_tokens = [
            ('DESCRIPCIÓN MACROSCÓPICA', 'DESCRIPCIÓN MICROSCÓPICA'),
            ('DESCRIPCIÓN MICROSCÓPICA', 'MÉTODO UTILIZADO:'),
            ('DIAGNÓSTICO:', '_____________________________________')
        ]

        row = {}
        text_acc = ""

        def add_token(token, result):
            row[token] = result

        for line in text.splitlines(True):  # keep newlines
            text_acc += line
            for token in tokens:
                ind = line.find(token)
                if ind != -1:
                    result = line[ind+len(token):].strip()
                    if token in (f'Informe No {mnum}.', 'Fecha de Recibo:'):
                        result = result.replace('|', '').strip()
                    add_token(token, result)

            for token_list in compound_tokens:
                inds = [line.find(tk) for tk in token_list]
                if -1 not in inds:
                    inds.append(len(line) + 1)
                    for i in range(len(inds) - 1):
                        tk = token_list[i]
                        ind = inds[i]
                        nind = inds[i + 1]
                        result = line[ind + len(tk):nind].strip()
                        add_token(tk, result)

        for start_token, end_token in next_line_tokens:
            start_ind = text_acc.find(start_token)
            end_ind = text_acc.find(end_token, start_ind)
            if start_ind != -1 and end_ind != -1:
                result = text_acc[start_ind + len(start_token):end_ind].strip().replace('\r\n', '\n').replace('\n', '')
                add_token(start_token, result)

        if "Servicio:" in row:
            del row["Servicio:"]

        try:
            rows.append([row[key] for key in order])
        except Exception:
            err = traceback.format_exc()
            error_de_archivo.append(f"{file_name} :: extracción de campos fallida\n{err}")

        # progreso
        procesados += 1
        progress['value'] = procesados
        status_text.set(f"{procesados} / {len(archivos)} procesados")
        now = time.monotonic()
        if now - last_ui > 0.12:
            root.update_idletasks()
            last_ui = now

    # Guardar salidas
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter='|')
        writer.writerows(rows)

    with open(error_path, 'w', encoding='utf-8') as f:
        if error_de_archivo:
            f.write(f"Archivos con error ({len(error_de_archivo)})\n\n")
            f.write("\n\n".join(error_de_archivo))
        else:
            f.write("Sin errores")

    file_text.set("Finalizado")
    status_text.set(f"{procesados} / {len(archivos)} procesados")
    progress['value'] = progress['maximum']
    root.update_idletasks()

    messagebox.showinfo(
        "Conversión completada",
        f"CSV guardado en:\n{csv_path}\n\nErrores guardados en:\n{error_path}\n\nArchivos procesados: {procesados}",
        parent=root,
    )
    root.destroy()

if __name__ == "__main__":
    main()
