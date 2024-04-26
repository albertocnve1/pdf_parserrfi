import os
import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
from datetime import datetime
import tempfile

# Creazione di uno stile per il formato delle celle in euro
euro_style = NamedStyle(name="euro_style")
euro_style.number_format = '€ #,##0.00'  # Formato valuta in euro

def italian_month_to_number(month):
    # Dizionario per la conversione dei nomi dei mesi italiani in numeri
    months = {
        "Gennaio": 1,
        "Febbraio": 2,
        "Marzo": 3,
        "Aprile": 4,
        "Maggio": 5,
        "Giugno": 6,
        "Luglio": 7,
        "Agosto": 8,
        "Settembre": 9,
        "Ottobre": 10,
        "Novembre": 11,
        "Dicembre": 12
    }
    return months.get(month, 0)  # Restituisce 0 se il mese non è presente nel dizionario

def nome_mese(mese):
    mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
    return mesi[mese - 1] if 1 <= mese <= 12 else None

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
                # Crea un file di testo temporaneo
                temp_text_file = tempfile.NamedTemporaryFile(mode='w', delete=False)
                
                # Estrai il mese e l'anno dal nome del file PDF
                filename, file_extension = os.path.splitext(pdf_file)
                month, year = filename.split(' ', 1)
                
                # Path completo del file PDF
                pdf_path = os.path.join(year_folder, pdf_file)
                
                # Apre il file PDF con pdfplumber
                with pdfplumber.open(pdf_path) as pdf:
                    # Itera attraverso tutte le pagine del PDF
                    for page_num in range(len(pdf.pages)):
                        page = pdf.pages[page_num]
                        text = page.extract_text()
                        
                        # Itera attraverso tutte le righe del testo
                        for line in text.split('\n'):
                            # Utilizza espressioni regolari per trovare le righe con i codici desiderati
                            match = re.match(r'^(0969|0970|0991|0992|0AD0|0AD1|0421|0457|0131|0576|0376|0377|0169|0170|0965|0966|0967|0987|0988|0790|0076)\s', line)
                            if match:
                                # Scrivi la riga nel file di testo temporaneo
                                temp_text_file.write(line + '\n')
                
                # Chiudi il file di testo temporaneo
                temp_text_file.close()
                
                # Analizza il file di testo temporaneo e aggiorna i dati
                with open(temp_text_file.name, 'r') as temp_file:
                    for line in temp_file:
                        # Separa la riga in codice, descrizione e valore
                        cells = line.split()
                        if len(cells) >= 3:
                            # Sostituisci la virgola con un punto per consentire la conversione in float
                            value = cells[-1].replace(',', '.')
                            if month not in data:
                                data[month] = {}
                            if cells[0] not in data[month]:
                                data[month][cells[0]] = float(value)
                            else:
                                data[month][cells[0]] += float(value)
                
                # Elimina il file di testo temporaneo
                os.remove(temp_text_file.name)
            
            # Ordina i mesi in ordine cronologico utilizzando il numero del mese
            months_sorted = sorted(data.keys(), key=lambda x: italian_month_to_number(x))
            
            # Scrittura dei dati nel foglio di lavoro
            ws.append(['Codice'] + months_sorted)  # Intestazioni delle colonne con i mesi
            for code in sorted(data[months_sorted[0]].keys()):
                row = [code]
                for month in months_sorted:
                    row.append(data[month].get(code, ''))
                ws.append(row)
            
            # Modifica il nome del foglio di lavoro con il nome del dipendente e l'anno lavorativo
            ws.title = f'Anno {year}'
            
            # Somma di tutte le celle da B alla colonna corrente e scrittura nella cella corrispondente alla riga Totale (riga 30)
            last_row = 30  # Righe del totale
            ws[f'O{last_row}'] = f"=SUM(B{last_row}:M{last_row})"
            
            # Scrittura di "TOTALE" nella cella A corrispondente alla riga Totale (riga 30)
            ws[f'A{last_row}'] = "TOTALE"
            
            # Applica lo stile in grassetto alla riga del totale
            ws[f'A{last_row}'].font = Font(bold=True)
            for col in range(2, 14):  # Modifica il range fino alla colonna M
                ws[get_column_letter(col) + str(last_row)].font = Font(bold=True)
                
                # Somma delle righe dalla riga 4 alla riga 30 e scrittura nella riga 30 per ogni colonna
                ws[get_column_letter(col) + '30'] = f"=SUM({get_column_letter(col)}4:{get_column_letter(col)}{last_row-1})"
            
            # Aggiungi le scritte richieste nelle celle specifiche
            ws['A31'] = "Presenze"
            ws['A32'] = "Ferie"
            ws['A33'] = "Retribuzione mensile"
            ws['A34'] = "VALORE MEDIO GIORNALIERO VOCI CONTRATTUALI ACCESSORIE"
            ws['A35'] = "RETRIBUZIONE GIORNALIERA (1/26 ART.68, COMM.6 DEL CCNL)"
            ws['A36'] = "INCIDENZA"
            
            # Applica lo stile in grassetto alla riga 34
            ws['A34'].font = Font(bold=True)
            ws['A35'].font = Font(bold=True)
            ws['A36'].font = Font(bold=True)
            # Calcolo dei valori nelle celle specificate
            ws['B34'] = f"=O30/B31"
            ws['B35'] = f"=B33/26"
            ws['B36'] = f"=(B34/100)*B35"
            ws['C36'] = f"%"
            # Applicazione dello stile alle celle specificate
            for cell in ['B34', 'B35']:
                ws[cell].style = euro_style

            # Applicazione dello stile alle celle della riga 30 tranne A30
            for cell in ws[30]:
                if cell.column_letter != 'A':
                    cell.style = euro_style

            for row in range(2, 30):
                for col in range(2, 14):
                    cell = ws.cell(row=row, column=col)
                    cell.style = euro_style

            # Allinea il testo delle celle
            for row in ws.iter_rows(min_row=31, max_row=36, min_col=1, max_col=1):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left')
            
            # Path del file Excel per il dipendente corrente
            excel_path = os.path.join(pdf_folder, f'{employee_name}.xlsx')
            
            # Salva il foglio di lavoro Excel per il dipendente corrente
            wb.save(excel_path)
            
            print(f"Estrazione completata per il dipendente {employee_name} anno {year}.")

# Funzione per preparare i file PDF
def prepare_and_extract(pdf_folder):
    # Attraversa ricorsivamente tutte le sottocartelle e trova i file PDF
    for dirpath, _, filenames in os.walk(pdf_folder):
        for filename in filenames:
            if filename.endswith('.PDF') or filename.endswith('.pdf'):
                file_path = os.path.join(dirpath, filename)
                try:
                    # Estrapola l'anno e il mese dal nome del file
                    anno = int(filename[:4])
                    mese = int(filename[5:7])
                except ValueError:
                    print(f"Il file '{filename}' non segue il formato YYYY_MM.pdf. Ignorato.")
                    continue

                # Ottieni il nome del mese in italiano
                nome_mese_it = nome_mese(mese)

                if nome_mese_it:
                    # Rimuovi spazi extra dal nome del mese
                    nome_mese_it = nome_mese_it.strip()

                    # Rinomina il file con il formato desiderato
                    nuovo_nome = f"{nome_mese_it} {anno}.pdf"
                    # Rinomina il file
                    nuovo_path = os.path.join(dirpath, nuovo_nome)
                    os.rename(file_path, nuovo_path)
                    print(f"File rinominato: {file_path} -> {nuovo_path}")
                else:
                    print(f"Errore: Mese non valido nel file {filename}")
    
    # Dopo la preparazione dei PDF, esegui l'estrazione dei dati
    extract_specific_lines_to_excel(pdf_folder)

# Chiedi all'utente di inserire il percorso della cartella dei dipendenti con le buste paga
customers_folder = input("Inserisci il percorso della cartella dei dipendenti con le buste paga: ")

# Chiama la funzione per preparare i file PDF e poi estrarre le righe specifiche dai PDF e scriverle in fogli Excel per ogni dipendente e anno
prepare_and_extract(customers_folder)
