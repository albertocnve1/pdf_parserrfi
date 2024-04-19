import os
from datetime import datetime
import glob

# Definisci una funzione per ottenere il nome del mese in italiano
def nome_mese(mese):
    mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
    return mesi[mese - 1] if 1 <= mese <= 12 else None

# Definisci la cartella radice da cui iniziare la scansione
cartella_radice = input("Inserisci il percorso della cartella radice: ")

# Attraversa ricorsivamente tutte le sottocartelle e trova i file PDF
for dirpath, _, filenames in os.walk(cartella_radice):
    for filename in filenames:
        if filename.endswith('.PDF') or filename.endswith('.pdf'):  # Modificato per considerare anche estensioni in maiuscolo
            file_path = os.path.join(dirpath, filename)

            # Estrapola l'anno e il mese dal nome del file
            try:
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
