import os
import sqlite3
import docx2txt
import pyperclip
import keyboard
import time
import pyautogui

#Seleziona testo, control+c, poi premi altgr e verrà scritta la risposta, premi spazio per interrompere

def move_cursor_to_textbox(x, y):
    # Sposta il cursore del mouse nella posizione desiderata
    pyautogui.moveTo(x, y)

def write_to_textbox(text):
    # Simula la digitazione del testo nella textbox
    pyautogui.typewrite(text)

def move_output_window():
    # Ottieni le coordinate correnti del cursore
    x, y = pyautogui.position()

    # Muovi la finestra di output vicino al cursore
    #pyautogui.moveTo(x + 20, y + 20)

def scrivi_dove_punta(parola):
    # Ottieni le coordinate del mouse
    x, y = pyautogui.position()

    # Sposta il cursore del mouse alle coordinate ottenute
    pyautogui.moveTo(x, y)

    # Inizializza la variabile di controllo
    interruzione_spazio = False

    for carattere in parola:
        if keyboard.is_pressed('space'):
            # Interrompi la scrittura se il tasto "Spazio" è premuto
            interruzione_spazio = True
            break
        else:
            # Scrivi il carattere
            pyautogui.typewrite(carattere)

    # Dopo l'interruzione, rimuovi tutti gli elementi dalla clipboard
    pyperclip.copy('')

def estrai_domande_risposte_da_testo(testo):
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
    # Funzione per ottenere il testo selezionato con il cursore del mouse
    return pyperclip.paste()

def main():
    # Memorizza il contenuto della clipboard prima della pressione di AltGr
    clipboard_prima = pyperclip.paste()

    # Aspetta che venga premuto il tasto 'AltGr'
    keyboard.wait("AltGr")

    # Aspetta un breve periodo di tempo per assicurarsi che il testo sia stato copiato negli appunti
    time.sleep(1.0)

    # Ottieni il testo selezionato con il cursore del mouse
    input_utente = leggi_testo_selezionato()

    # Confronta con la clipboard prima della pressione di AltGr
    if input_utente and clipboard_prima != input_utente:
        print("Testo selezionato:", input_utente)

        # Estrai domande e risposte dal documento Word
        nome_file_word = 'EHH.docx'
        percorso_file_word = os.path.join(os.path.dirname(__file__), nome_file_word)
        testo_documento = docx2txt.process(percorso_file_word)
        domande_risposte = estrai_domande_risposte_da_testo(testo_documento)

        # Nome del database SQLite
        nome_db = 'domande_risposte.db'

        # Crea o controlla il database
        crea_o_controlla_db(nome_db)

        # Inserisci nel database SQLite solo se necessario
        inserisci_su_db_se_necessario(domande_risposte, nome_db)

        # Cerca la migliore corrispondenza nel database
        domanda_migliore, risposta_migliore = cerca_miglior_match(input_utente, domande_risposte)

        if domanda_migliore and risposta_migliore:
            # Muovi il cursore nella textbox
            #move_cursor_to_textbox()
            #move_output_window()
            #move_cursor_to_textbox(x, y)

            # Scrivi il risultato nella textbox
            #write_to_textbox(f"Domanda migliore: {domanda_migliore}\nRisposta migliore: {risposta_migliore}\n")
            scrivi_dove_punta(risposta_migliore)
        else:
            # Scrivi un messaggio nel caso non ci siano corrispondenze nel database
            print("Nessuna corrispondenza trovata nel database.")

    # Muovi la finestra di output vicino al cursore
    #move_output_window()

# Loop infinito
while True:
    try:
        # Esegui la funzione principale
        main()

        # Attendi prima di eseguire nuovamente
        time.sleep(0,5)        
        
    except Exception as e:
        # Ignora le eccezioni e continua il loop
        pass
