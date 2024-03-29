import PyPDF2
import re

def extract_specific_lines(pdf_path, text_path):
    # Apre il file PDF in modalità lettura binaria
    with open(pdf_path, 'rb') as pdf_file:
        # Crea un lettore PDF
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        
        # Inizializza una lista vuota per contenere le righe estratte
        extracted_lines = []

        # Estrae il testo da ciascuna pagina del PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            lines = text.split('\n')
            # Itera su ciascuna riga del testo
            for line in lines:
                # Utilizza espressioni regolari per trovare le righe con i codici desiderati
                match = re.match(r'^(0969|0970|0991|0992|0AD0|0AD1)', line)
                if match:
                    extracted_lines.append(line)
            # Aggiungi un salto di riga alla fine della pagina
            extracted_lines.append('\n')

        # Scrive le righe estratte su un file di testo
        with open(text_path, 'w', encoding='utf-8') as text_file:
            for line in extracted_lines:
                text_file.write(line + '\n')

# Path del file PDF da convertire
pdf_path = '/Users/albertocanavese/Desktop/test_python/anno 2014.pdf'

# Path del file di testo in cui verrà scritto il testo estratto
text_path = '/Users/albertocanavese/Desktop/test_python/output.txt'

# Chiama la funzione per estrarre le righe specifiche dal PDF e scriverle su un file di testo
extract_specific_lines(pdf_path, text_path)

print("Estrazione completata!")
