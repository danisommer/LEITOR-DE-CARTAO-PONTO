import PyPDF2
import csv
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para extrair horários e datas de um pdf
def extrair_datas_horarios(pdf_path):
    dados = []
    pdf = PyPDF2.PdfReader(open(pdf_path, "rb"))

    i = 0
    while i < len(pdf.pages):
        page = pdf.pages[i]
        page_text = page.extract_text()
        lines = page_text.split('\n')

        for line in lines:
            data = None
            horarios = []
            
            data = re.findall(r'\b\d{2}[/]\d{2}[/]\d{4}\b', line)
            
            if data:
                # Use re.findall para encontrar todos os horários da linha
                horarios_match = re.findall(r'\b\d{2}:\d{2}\b', line)
                if horarios_match:
                    horarios.extend(horarios_match)

                while len(horarios) < 4:
                    horarios.append('')

                dados.append(data + horarios)

        i += 1

    return dados

# Função para salvar os horários em um arquivo CSV
def salvar_csv(dados, csv_path):
    with open(csv_path, mode='w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["DATA", "ENTRADA 1", "SAIDA 1", "ENTRADA 2", "SAIDA 2"])
        writer.writerows(dados)

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

        dados = extrair_datas_horarios(pdf_path)
        csv_path = os.path.join(diretorio_saida, encontrar_nome_unico(diretorio_saida, csv_path_base, extensao_csv))

        salvar_csv(dados, csv_path)

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