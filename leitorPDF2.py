import PyPDF2
import csv
import re

def extrair_datas_horarios(pdf_path):
    dados = []
    pdf = PyPDF2.PdfReader(open(pdf_path, "rb"))

    i = 0
    while i < len(pdf.pages):
        page = pdf.pages[i]
        page_text = page.extract_text()
        lines = page_text.split('\n')

        data = None
        horarios = []

        for line in lines:
            # Use re.search corretamente para encontrar datas
            date_match = re.search(r'\d{2}[-/]\d{2}[-/]\d{4}', line)
            if date_match:
                if data is not None:
                    break  # Já encontramos outra data, paramos de buscar horários
                data = date_match.group().strip()
            elif re.search(r'\b\d{2}:\d{2}\b', line) and data is not None:
                # Adicione um loop para buscar todos os horários da linha
                for horario_match in re.finditer(r'\b\d{2}:\d{2}\b', line):
                    horarios.append(horario_match.group())

        if data is not None:
            # Preencha com espaços em branco caso haja menos de 4 horários
            while len(horarios) < 4:
                horarios.append(' ')

            dados.append([data] + horarios)

        i += 1

    return dados

def salvar_csv(dados, csv_path):
    with open(csv_path, mode='w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["DATA", "ENTRADA 1", "SAIDA 1", "ENTRADA 2", "SAIDA 2"])
        writer.writerows(dados)

if __name__ == "__main__":
    pdf_path = "cartao_completo.pdf"
    csv_path = "saida.csv"

    dados = extrair_datas_horarios(pdf_path)
    salvar_csv(dados, csv_path)

    print("Dados extraídos e salvos em CSV com sucesso!")