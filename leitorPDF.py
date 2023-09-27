import PyPDF2
import re
import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para extrair horários de entrada e saída de um texto
def extrair_horarios(texto):
    padrao_horario = r"\d{2}:\d{2}"
    horarios = re.findall(padrao_horario, texto)
    return horarios

# Função para extrair todas as datas de uma página
def extrair_datas_da_pagina(page_text):
    padrao_data = r"\d{2}/\d{2}/\d{4}"
    datas = re.findall(padrao_data, page_text)
    return datas

# Função para ler o PDF e extrair os horários
def extrair_horarios_de_pdf(pdf_path):
    horarios_por_dia = {}
    
    pdf_file = open(pdf_path, "rb")
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    data = []
    horarios1 = []
    horarios2 = []
    
    for page_num in range(len(pdf_reader.pages)):
        dif = 0
        page = pdf_reader.pages[page_num]
        page_text = page.extract_text()
        
        datas_da_pagina = extrair_datas_da_pagina(page_text)
        horarios_da_pagina = extrair_horarios(page_text)

        dif = int(len(datas_da_pagina) - len(horarios_da_pagina) / 4)
        if dif > 0:
            for _ in range(dif):
                horarios1.append(("", ""))  # Adicione um par de horários em branco
                horarios2.append(("", ""))
            
        for data in datas_da_pagina:
            if horarios_da_pagina:  # Verifique se há horários disponíveis
                horarios1 = horarios_da_pagina[:2]  # Captura os 2 primeiros horários
                horarios_da_pagina = horarios_da_pagina[2:]  # Remove os horários capturados
                horarios2 = horarios_da_pagina[:2]  # Captura os 2 horários seguintes ou vazios
                horarios_da_pagina = horarios_da_pagina[2:]  # Remove os horários capturados ou mantém vazio

                if data not in horarios_por_dia:
                    horarios_por_dia[data] = [] 
                
                if horarios1[0] and horarios1[1] and horarios2[0] and horarios2[1]:
                    horarios_por_dia[data].append((horarios1[0], horarios1[1], horarios2[0], horarios2[1]))
                elif horarios1[0] and horarios1[1]:
                    horarios_por_dia[data].append((horarios1[0], horarios1[1], "", ""))
                elif horarios2[0] and horarios2[1]:
                    horarios_por_dia[data].append(("", "", horarios2[0], horarios2[1]))
                else: 
                    horarios_por_dia[data].append(("", "", "", ""))
    
    pdf_file.close()
    
    return horarios_por_dia

# Função para salvar os horários em um arquivo CSV
def salvar_horarios_em_csv(horarios, csv_path):
    with open(csv_path, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Data", "Entrada 1", "Saída 1", "Entrada 2", "Saída 2"])
        
        for data, lista_horarios in horarios.items():
            for entrada1, saida1, entrada2, saida2 in lista_horarios:
                writer.writerow([data, entrada1, saida1, entrada2, saida2])

# Função para selecionar o arquivo PDF de entrada
def selecionar_arquivo_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        return pdf_path
    else:
        return None

# Função para encontrar um nome de arquivo único
def encontrar_nome_unico(diretorio, nome_base, extensao):
    contador = 1
    while True:
        novo_nome = f"{nome_base}_{contador}.{extensao}"
        if not os.path.exists(f"{diretorio}\{novo_nome}"):
            return novo_nome
        contador += 1

# Função para selecionar o diretório de saída
def selecionar_diretorio_saida():
    diretorio_saida = filedialog.askdirectory()
    if diretorio_saida:
        return diretorio_saida
    else:
        return None

# Função principal para processar o PDF
def processar_pdf():
    pdf_path = selecionar_arquivo_pdf()
    
    if pdf_path:
        csv_path_base = "horarios"  # Nome base do arquivo CSV de saída
        extensao_csv = "csv"  # Extensão do arquivo CSV
        
        diretorio_saida = selecionar_diretorio_saida()
        if not diretorio_saida:
            return

        if os.path.exists(pdf_path):
            messagebox.showinfo("Sucesso", "Leitura bem sucedida!")
        else:
            messagebox.showerror("Erro", "Erro ao abrir o arquivo PDF.")
            return

        horarios = extrair_horarios_de_pdf(pdf_path)
        csv_path = os.path.join(diretorio_saida, encontrar_nome_unico(diretorio_saida, csv_path_base, extensao_csv))

        salvar_horarios_em_csv(horarios, csv_path)

        if os.path.exists(csv_path):
            messagebox.showinfo("Sucesso", "CSV criado com sucesso em: " + csv_path)
        else:
            messagebox.showerror("Erro", "Erro, não foi possível criar o CSV.")


root = tk.Tk()
root.title("Extrair Horários de PDF")
root.geometry("380x135+562+320")  # Centralize a janela e defina a posição
root.resizable(False, False)  # Bloquear altura e largura da janela

header_label = tk.Label(root, text="Extrair Horários de PDF", font=("Arial", 16))
header_label.pack(pady=10)

frame = tk.Frame(root)
frame.pack(padx=20, pady=10)

select_pdf_button = tk.Button(frame, text="Selecionar PDF", command=processar_pdf)
select_pdf_button.pack(side="left")

exit_button = tk.Button(frame, text="Sair", command=root.quit)
exit_button.pack(side="right")

root.mainloop()