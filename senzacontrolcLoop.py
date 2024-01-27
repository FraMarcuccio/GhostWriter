import time
import pyperclip
import pyautogui
import os
import sqlite3
import docx2txt

def estrai_domande_risposte(testo):
    linee = testo.split('\n')

    domande_risposte = []
    domanda_attuale = None
    risposta_attuale = None

    for linea in linee:
        if linea.strip().startswith(('1', '2', '3', '4', '5', '6', '7', '8', '9')):
            if domanda_attuale and risposta_attuale:
                domande_risposte.append((domanda_attuale, risposta_attuale))
            domanda_attuale = linea.strip()
            risposta_attuale = ""
        elif domanda_attuale is not None:
            risposta_attuale += linea.strip() + '\n'

    if domanda_attuale and risposta_attuale:
        domande_risposte.append((domanda_attuale, risposta_attuale))

    return domande_risposte

def crea_o_controlla_db(nome_db):
    connessione = sqlite3.connect(nome_db)
    cursore = connessione.cursor()

    # Crea la tabella se non esiste
    cursore.execute('''
        CREATE TABLE IF NOT EXISTS domande_risposte (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            domanda TEXT,
            risposta TEXT
        )
    ''')

    connessione.commit()
    connessione.close()

def inserisci_su_db_se_necessario(domande_risposte, nome_db):
    connessione = sqlite3.connect(nome_db)
    cursore = connessione.cursor()

    # Seleziona tutte le domande e risposte dal database
    cursore.execute('SELECT domanda, risposta FROM domande_risposte')
    domande_risposte_db = cursore.fetchall()

    # Confronta le domande nel documento con quelle nel database
    if set(domande_risposte) == set(domande_risposte_db):
        print("Le domande nel documento sono identiche a quelle nel database. Nessuna azione necessaria.")
    else:
        # Cancella tutte le domande nel database
        cursore.execute('DELETE FROM domande_risposte')

        # Inserisci le nuove domande nel database
        for domanda, risposta in domande_risposte:
            cursore.execute('INSERT INTO domande_risposte (domanda, risposta) VALUES (?, ?)', (domanda, risposta))

        connessione.commit()
        print("Domande e risposte nel database aggiornate.")

    # Chiudi la connessione
    connessione.close()

def cerca_miglior_match(input_utente, domande_risposte):
    parole_input = input_utente.lower().split()
    miglior_match_domanda = None
    miglior_match_risposta = None
    max_corrispondenze = 0

    for domanda, risposta in domande_risposte:
        corrispondenze = sum(parola in domanda.lower() for parola in parole_input)
        if corrispondenze > max_corrispondenze:
            max_corrispondenze = corrispondenze
            miglior_match_domanda = domanda
            miglior_match_risposta = risposta

    return miglior_match_domanda, miglior_match_risposta

def leggi_testo_selezionato():
    # Simula la combinazione di tasti Ctrl+C
    pyautogui.hotkey('ctrl', 'c')
    # Ottieni il testo dalla clipboard
    testo_selezionato = pyperclip.paste()
    return testo_selezionato

def main():
  
    # Ottieni il testo selezionato con il cursore del mouse
    input_utente = leggi_testo_selezionato()

    # Se il testo selezionato non è vuoto, procedi
    if input_utente:
        print("Testo selezionato:", input_utente)

        # Estrai domande e risposte dal documento Word
        nome_file_word = 'EHH.docx'
        percorso_file_word = os.path.join(os.path.dirname(__file__), nome_file_word)
        testo_documento = docx2txt.process(percorso_file_word)
        domande_risposte = estrai_domande_risposte(testo_documento)

        # Nome del database SQLite
        nome_db = 'domande_risposte.db'

        # Crea o controlla il database
        crea_o_controlla_db(nome_db)

        # Inserisci nel database SQLite solo se necessario
        inserisci_su_db_se_necessario(domande_risposte, nome_db)

        # Cerca la migliore corrispondenza nel database
        domanda_migliore, risposta_migliore = cerca_miglior_match(input_utente, domande_risposte)

        if domanda_migliore and risposta_migliore:
            print("Domanda migliore:", domanda_migliore)
            print("Risposta migliore:", risposta_migliore)
        else:
            print("Nessuna corrispondenza trovata nel database.")
    else:
        print("Nessun testo selezionato.")

# Loop infinito
while True:
    try:
        # Esegui la funzione principale
        main()

        # Attendi prima di eseguire nuovamente
        time.sleep(2.5)
    except Exception as e:
        # Ignora le eccezioni e continua il loop
        pass
