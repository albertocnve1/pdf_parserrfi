import PyPDF2

def convert_pdf_to_text(pdf_path, text_path):
    # Apre il file PDF in modalità lettura binaria
    with open(pdf_path, 'rb') as pdf_file:
        # Crea un lettore PDF
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        
        # Inizializza una stringa vuota per contenere il testo estratto
        text = ''
        
        # Estrae il testo da ciascuna pagina del PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        
        # Scrive il testo estratto su un file di testo
        with open(text_path, 'w', encoding='utf-8') as text_file:
            text_file.write(text)

# Path del file PDF da convertire
pdf_path = '/Users/albertocanavese/Desktop/test_python/anno 2014.pdf'

# Path del file di testo in cui verrà scritto il testo estratto
text_path = '/Users/albertocanavese/Desktop/test_python/output.txt'

# Chiama la funzione per convertire il PDF in testo
convert_pdf_to_text(pdf_path, text_path)

print("Conversione completata!")
