import os
import re
import pdfplumber
from openpyxl import Workbook

def extract_specific_lines_to_excel(pdf_folder):
    # Elenco delle cartelle nella cartella specificata
    folders = [f.path for f in os.scandir(pdf_folder) if f.is_dir()]

    # Itera attraverso ogni cartella
    for folder in folders:
        # Estrai il nome della cartella
        folder_name = os.path.basename(folder)
        
        # Crea un nuovo foglio di lavoro Excel per il dipendente corrente
        wb = Workbook()
        
        # Dizionario per memorizzare le somme delle righe per ogni legenda
        summary = {}
        
        # Elenco dei file PDF nella cartella del dipendente corrente
        pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
        
        # Itera attraverso ogni file PDF nella cartella del dipendente corrente
        for pdf_file in pdf_files:
            # Estrai il nome del file senza l'estensione
            file_name = os.path.splitext(pdf_file)[0]
            
            # Crea un nuovo foglio di lavoro Excel per il file corrente
            ws = wb.create_sheet(title=file_name)
            
            # Path completo del file PDF
            pdf_path = os.path.join(folder, pdf_file)
            
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
                                # Sostituisci la virgola con un punto per consentire la conversione in float
                                value = cells[-1].replace(',', '.')
                                extracted_lines.append([cells[0], ' '.join(cells[1:-1]), float(value)])
                    
                    # Aggiungi le righe estratte al foglio di lavoro
                    for row in extracted_lines:
                        ws.append(row)
                        
                        # Aggiorna il dizionario di riepilogo
                        legend = row[0]
                        value = row[2]
                        if legend in summary:
                            summary[legend] += value
                        else:
                            summary[legend] = value
        
            # Aggiungi legenda sopra i vari codici e risultati anche su questo foglio
            ws.insert_rows(1)
            ws['A1'] = "Codice"
            ws['B1'] = "Descrizione"
            ws['C1'] = "Valore"
        
        # Rimuovi il foglio di lavoro predefinito
        wb.remove(wb['Sheet'])
        
        # Crea un nuovo foglio di lavoro per il riepilogo
        summary_ws = wb.create_sheet(title="RIEPILOGO")
        
        # Aggiungi legenda sopra i vari codici e risultati
        summary_ws.append(["Codice", "Descrizione", "Valore"])
        
        # Scrivi le somme nel foglio di riepilogo
        for legend, total in summary.items():
            summary_ws.append([legend, '', total])  # aggiungi una riga vuota tra ogni codice e il suo totale
        
        # Path del file Excel per il dipendente corrente
        excel_path = os.path.join(folder, f'{folder_name}.xlsx')
        
        # Salva il foglio di lavoro Excel per il dipendente corrente
        wb.save(excel_path)
        
        print(f"Estrazione completata per il dipendente {folder_name}.")

# Cartella contenente le cartelle dei dipendenti con le buste paga
customers_folder = '/Users/albertocanavese/Desktop/test_python/Buste paga'

# Chiama la funzione per estrarre le righe specifiche dai PDF e scriverle in fogli Excel per ogni dipendente
extract_specific_lines_to_excel(customers_folder)
