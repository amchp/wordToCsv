import os
import csv
import pypandoc
from tkinter import Tk
from tkinter.filedialog import askdirectory
from tkinter import simpledialog

Tk().withdraw()

carpeta = askdirectory()

mnum = simpledialog.askstring("Pregunta", "Cuál es el M.**? (Ejemplo: M.23)")

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

for file_name in os.listdir(carpeta):

    print(file_name, "procesado")
    if '.doc' not in file_name:
        continue
    try:
        output = pypandoc.convert_file(carpeta + '/' + file_name, 'plain', outputfile="tmp.txt")
    except: 
        error_de_archivo.append(file_name)
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

with open('output.csv', 'w', encoding='utf-8-sig') as file:
    writer = csv.writer(file, delimiter='|')
    writer.writerows(rows)

with open('error.txt', 'w') as file:
    file.write("\n".join(error_de_archivo))
