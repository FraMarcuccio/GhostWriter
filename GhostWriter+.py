import os
import sqlite3
import docx2txt
import pyperclip
import keyboard
import time
import pyautogui
import difflib

#Seleziona testo, control+c, poi premi altgr e verrà scritta la risposta, per interromepre cursore all'angolo dello schermo

pyautogui.FAILSAFE = True # pyautogui ha un valore failsafe impostato a coordinate X=0 e Y=0 del mouse, se il cursore quindi va in un angolo dello schermo, si blocca la scrittura

def scrivi_dove_punta(parola):
    lunghezza_parola = len(parola)
    lunghezza_parola_scritta = 0

    try:
        while lunghezza_parola_scritta < lunghezza_parola:
            # Ottieni le coordinate del mouse
            x, y = pyautogui.position()

            # Sposta il cursore del mouse alle coordinate ottenute
            pyautogui.moveTo(x, y)

            # Scrivi la porzione mancante della parola
            caratteri_da_scrivere = parola[lunghezza_parola_scritta:]
            pyautogui.write(caratteri_da_scrivere, interval=0.005)

            # Aggiorna la lunghezza della parola scritta
            lunghezza_parola_scritta = len(parola)

            # Dopo l'interruzione, rimuovi tutti gli elementi dalla clipboard
            pyperclip.copy('')

    except pyautogui.FailSafeException:
        print("Failsafe attivato! PyAutoGUI è stato interrotto.")


def cerca_miglior_match(input_utente, domande_risposte):
    parole_input = input_utente.lower().split()
    miglior_match_domanda = None
    miglior_match_risposta = None
    max_corrispondenze = 0

    for domanda, risposta in domande_risposte:
        corrispondenze = 0
        for parola_input in parole_input:
            # Calcola il rapporto di similarità con ogni parola nella domanda
            similarita_max = max(difflib.SequenceMatcher(None, parola_input, parola_domanda).ratio()
                                for parola_domanda in domanda.lower().split())
            
            # Se il rapporto supera una certa soglia, considera la parola come corrispondente
            if similarita_max > 0.8:
                corrispondenze += 1

        if corrispondenze > max_corrispondenze:
            max_corrispondenze = corrispondenze
            miglior_match_domanda = domanda
            miglior_match_risposta = risposta

    return miglior_match_domanda, miglior_match_risposta


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

def main():
    # Aspetta che venga premuto il tasto 'AltGr'
    keyboard.wait("AltGr")

    # Aspetta un breve periodo di tempo per assicurarsi che il testo sia stato copiato negli appunti
    time.sleep(1.0)

    # Ottieni il testo selezionato con il cursore del mouse
    input_utente = pyperclip.paste()

    print("Testo selezionato:", input_utente)

    # Estrai domande e risposte dal documento Word
    nome_file_word = 'EHH.docx'
    percorso_file_word = os.path.join(os.path.dirname(__file__), nome_file_word)
    testo_documento = docx2txt.process(percorso_file_word)
    domande_risposte = estrai_domande_risposte_da_testo(testo_documento)

    # Cerca la migliore corrispondenza nel database
    domanda_migliore, risposta_migliore = cerca_miglior_match(input_utente, domande_risposte)

    if domanda_migliore and risposta_migliore:
        print(len(risposta_migliore))
        scrivi_dove_punta(risposta_migliore)
    else:
        # Scrivi un messaggio nel caso non ci siano corrispondenze nel database
        print("Nessuna corrispondenza trovata nel database.")

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
