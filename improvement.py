import multiprocessing
import time
import keyboard

def write_text_process(exit_event, pause_event):
    while not exit_event.is_set():
        if not pause_event.is_set():
            print("Testo a piacere...")
        time.sleep(1)

def wait_for_keypress_process(exit_event, pause_event):
    print("Premi - per mettere in pausa e + per riprendere. Premi ESC per terminare.")
    while True:
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

if __name__ == "__main__":
    # Creazione degli eventi condivisi tra i processi
    exit_event = multiprocessing.Event()
    pause_event = multiprocessing.Event()

    # Creazione dei processi e passaggio degli eventi
    write_process = multiprocessing.Process(target=write_text_process, args=(exit_event, pause_event))
    wait_process = multiprocessing.Process(target=wait_for_keypress_process, args=(exit_event, pause_event))

    # Avvio dei processi
    write_process.start()
    wait_process.start()

    try:
        # Attendere che il processo di scrittura termini
        write_process.join()
    except KeyboardInterrupt:
        # Gestire l'interruzione manuale del processo
        pass

    # Attendere che il processo di attesa termini
    wait_process.join()

    print("Programma terminato.")
