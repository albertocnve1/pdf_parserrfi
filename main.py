import os
import re
import pdfplumber
from openpyxl import Workbook

def extract_specific_lines_to_excel(pdf_folder, excel_path):
    # Crea un nuovo foglio di lavoro Excel
    wb = Workbook()

    # Elenco dei file PDF nella cartella specificata
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    # Itera attraverso ogni file PDF nella cartella
    for pdf_file in pdf_files:
        # Estrai il nome del file senza l'estensione
        file_name = os.path.splitext(pdf_file)[0]

        # Crea un nuovo foglio di lavoro Excel per il file corrente
        ws = wb.create_sheet(title=file_name)

        # Path completo del file PDF
        pdf_path = os.path.join(pdf_folder, pdf_file)

        # Apre il file PDF con pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            # Itera attraverso tutte le pagine del PDF
            for page_num in range(len(pdf.pages)):
                page = pdf.pages[page_num]
                text = page.extract_text()

                # Lista per memorizzare le righe estratte per la pagina corrente
                extracted_lines = []

                # Itera attraverso tutte le righe del testo
                for line in text.split('\n'):
                    # Utilizza espressioni regolari per trovare le righe con i codici desiderati
                    match = re.match(r'^(0969|0970|0991|0992|0AD0|0AD1)\s', line)
                    if match:
                        # Separa la riga in codice, descrizione e valore
                        cells = line.split()
                        if len(cells) >= 3:
                            extracted_lines.append([cells[0], ' '.join(cells[1:-1]), cells[-1]])

                # Aggiungi le righe estratte al foglio di lavoro
                for row in extracted_lines:
                    ws.append(row)

    # Rimuovi il foglio di lavoro predefinito
    wb.remove(wb['Sheet'])

    # Salva il foglio di lavoro Excel
    wb.save(excel_path)

# Cartella contenente i file PDF delle buste paga
pdf_folder = '/Users/albertocanavese/Desktop/test_python/Buste paga'

# Path del file Excel in cui verranno scritte le informazioni
excel_path = '/Users/albertocanavese/Desktop/test_python/output.xlsx'

# Chiama la funzione per estrarre le righe specifiche dai PDF e scriverle in un foglio Excel
extract_specific_lines_to_excel(pdf_folder, excel_path)

print("Estrazione completata!")