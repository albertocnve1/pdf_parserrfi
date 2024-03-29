import PyPDF2
import re
from openpyxl import Workbook

def extract_specific_lines_to_excel(pdf_path, excel_path):
    # Crea un nuovo foglio di lavoro Excel
    wb = Workbook()
    
    # Apre il file PDF in modalità lettura binaria
    with open(pdf_path, 'rb') as pdf_file:
        # Crea un lettore PDF
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Itera attraverso tutte le pagine del PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            lines = text.split('\n')

            # Lista per memorizzare le righe estratte per la pagina corrente
            extracted_lines = []

            # Itera attraverso tutte le righe del testo
            for line in lines:
                # Utilizza espressioni regolari per trovare le righe con i codici desiderati
                match = re.match(r'^(0969|0970|0991|0992|0AD0|0AD1)', line)
                if match:
                    # Separa la riga in codice, descrizione e valore
                    cells = line.split()
                    # Aggiungi le informazioni estratte alla lista
                    extracted_lines.append([cells[0], ' '.join(cells[1:-1]), cells[-1]])

            # Se ci sono righe estratte per questa pagina, crea un nuovo foglio di lavoro Excel per la pagina corrente
            if extracted_lines:
                # Crea un nuovo foglio di lavoro per la pagina corrente
                ws = wb.create_sheet(title=f'Pagina {page_num + 1}')
                ws.append(['Codice', 'Descrizione', 'Valore'])  # Aggiungi l'intestazione
                # Aggiungi le righe estratte al foglio di lavoro
                for row in extracted_lines:
                    ws.append(row)
                # Inserisci una riga vuota alla fine del foglio di lavoro
                ws.append([])

    # Salva il foglio di lavoro Excel
    wb.save(excel_path)

# Path del file PDF da convertire
pdf_path = '/Users/albertocanavese/Desktop/test_python/anno 2014.pdf'

# Path del file Excel in cui verranno scritte le informazioni
excel_path = '/Users/albertocanavese/Desktop/test_python/output.xlsx'

# Chiama la funzione per estrarre le righe specifiche dal PDF e scriverle in un foglio Excel
extract_specific_lines_to_excel(pdf_path, excel_path)

print("Estrazione completata!")
