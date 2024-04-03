import os
import re
import pdfplumber
from openpyxl import Workbook
from datetime import datetime

def italian_to_english_month(month):
    # Dizionario per la conversione dei nomi dei mesi italiani in inglesi
    months = {
        "Gennaio": "January",
        "Febbraio": "February",
        "Marzo": "March",
        "Aprile": "April",
        "Maggio": "May",
        "Giugno": "June",
        "Luglio": "July",
        "Agosto": "August",
        "Settembre": "September",
        "Ottobre": "October",
        "Novembre": "November",
        "Dicembre": "December"
    }
    return months.get(month, '')

def extract_specific_lines_to_excel(pdf_folder):
    # Elenco delle cartelle dei dipendenti nella cartella specificata
    employees = [f.path for f in os.scandir(pdf_folder) if f.is_dir()]
    
    # Itera attraverso ogni cartella (dipendente)
    for employee_folder in employees:
        # Estrai il nome del dipendente dalla cartella
        employee_name = os.path.basename(employee_folder)
        
        # Crea un nuovo foglio di lavoro Excel per il dipendente corrente
        wb = Workbook()
        
        # Dizionario per memorizzare le somme delle righe per ogni legenda
        summary = {}
        
        # Elenco delle cartelle (anni) per il dipendente corrente
        years = [f.path for f in os.scandir(employee_folder) if f.is_dir()]
        
        # Itera attraverso ogni cartella (anno) per il dipendente corrente
        for year_folder in years:
            # Estrai l'anno dalla cartella
            year = os.path.basename(year_folder)
            
            # Crea un nuovo foglio di lavoro Excel per l'anno corrente
            ws = wb.create_sheet(title=year)
            
            # Elenco dei file PDF nelle cartelle del dipendente e dell'anno corrente
            pdf_files = [f for f in os.listdir(year_folder) if f.endswith('.pdf')]
            
            # Dizionario per memorizzare i valori estratti per ogni codice e mese
            data = {}
            
            # Itera attraverso ogni file PDF per il dipendente e l'anno corrente
            for pdf_file in pdf_files:
                # Estrai il mese e l'anno dal nome del file PDF
                filename = os.path.splitext(pdf_file)[0]
                month, year = filename.split(' ', 1)
                month = italian_to_english_month(month)
                
                # Inizializza un dizionario per il mese corrente
                if month not in data:
                    data[month] = {}
                
                # Path completo del file PDF
                pdf_path = os.path.join(year_folder, pdf_file)
                
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
                                    extracted_lines.append([cells[0], float(value)])
                        
                        # Aggiungi le righe estratte al dizionario dei dati
                        for row in extracted_lines:
                            code = row[0]
                            value = row[1]
                            if code not in data[month]:
                                data[month][code] = value
                            else:
                                data[month][code] += value
            
            # Ordina i mesi in ordine cronologico
            months_sorted = sorted(data.keys(), key=lambda x: datetime.strptime(x, '%B').month)
            
            # Scrittura dei dati nel foglio di lavoro
            ws.append(['Codice'] + months_sorted)  # Intestazioni delle colonne con i mesi
            for code in sorted(data[months_sorted[0]].keys()):
                row = [code]
                for month in months_sorted:
                    row.append(data[month].get(code, ''))
                ws.append(row)
            
            # Modifica il nome del foglio di lavoro con il nome del dipendente e l'anno lavorativo
            ws.title = f'Anno {year}'
        
        # Modifica il titolo della cella A1 con il nome del dipendente e l'anno lavorativo
        first_sheet = wb[wb.sheetnames[0]]
        first_sheet['A1'] = f'Dipendente: {employee_name}'
        
        # Path del file Excel per il dipendente corrente
        excel_path = os.path.join(pdf_folder, f'{employee_name}.xlsx')
        
        # Salva il foglio di lavoro Excel per il dipendente corrente
        wb.save(excel_path)
        
        print(f"Estrazione completata per il dipendente {employee_name}.")

# Chiedi all'utente di inserire il percorso della cartella dei dipendenti con le buste paga
customers_folder = input("Inserisci il percorso della cartella dei dipendenti con le buste paga: ")

# Chiama la funzione per estrarre le righe specifiche dai PDF e scriverle in fogli Excel per ogni dipendente e anno
extract_specific_lines_to_excel(customers_folder)
