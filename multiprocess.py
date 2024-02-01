import os
import sqlite3
import docx2txt
import pyperclip
import keyboard
import time
import pyautogui
import multiprocessing

#Seleziona testo, control+c, poi premi altgr e verrà scritta la risposta

def scrivi_dove_punta(parola):
    # Ottieni le coordinate del mouse
    x, y = pyautogui.position()

    # Sposta il cursore del mouse alle coordinate ottenute
    pyautogui.moveTo(x, y)

    # Scrivi la parola
    pyautogui.typewrite(parola, interval=0.005)

def scrivi_dove_punta2(parola, pause_event, exit_event):
    x, y = pyautogui.position()
    pyautogui.moveTo(x, y)

    while not exit_event.is_set():
        if not pause_event.is_set():
            for carattere in parola:
                keyboard.press_and_release(carattere)
                time.sleep(0.005)  # Aggiungi una breve pausa tra i caratteri

def pause_process(pause_event, exit_event):
    print("Premi - per mettere in pausa e + per riprendere. Premi ESC per terminare.")
    while not exit_event.is_set():
        if keyboard.is_pressed('-'):
            pause_event.set()
            time.sleep(0.1)  # Aggiungi una breve pausa per evitare il rilevamento multiplo del tasto
        elif keyboard.is_pressed('+'):
            pause_event.clear()
            time.sleep(0.1)  # Aggiungi una breve pausa per evitare il rilevamento multiplo del tasto
        elif keyboard.is_pressed('esc'):
            exit_event.set()
            break
        time.sleep(0.1)

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

    # Creazione degli eventi condivisi tra i processi
    exit_event = multiprocessing.Event()
    pause_event = multiprocessing.Event()

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
            # Creazione dei processi e passaggio degli eventi
            write_process = multiprocessing.Process(target=scrivi_dove_punta2, args=(risposta_migliore, pause_event, exit_event))
            wait_process = multiprocessing.Process(target=pause_process, args=(pause_event, exit_event))
            # Avvio dei processi
            write_process.start()
            wait_process.start()

            # Attendere che il processo di scrittura termini
            write_process.join()

            # Attendere che il processo di attesa termini
            wait_process.join()
        else:
            print("Nessuna corrispondenza trovata nel database.")

    # Muovi la finestra di output vicino al cursore
    #move_output_window()

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
