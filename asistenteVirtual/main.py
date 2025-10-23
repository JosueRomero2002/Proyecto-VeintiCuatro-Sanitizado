from whisper_listener import WhisperListener
from openai_responder import OpenAIResponder
from sentinela import Sentinela
import threading
import time
import os


# Configuración inicial
openai_api_key = os.getenv("")

# Inicializar componentes
whisper_listener = WhisperListener(openai_api_key)
openai_responder = OpenAIResponder(openai_api_key)

# Función para manejar la transcripción de voz
def handle_transcription(text):
    print(f"Usuario dijo: {text}")
    if "stop" in text.lower():
        openai_responder.stop()
    else:
        threading.Thread(target=openai_responder.respond, args=(text, print)).start()

# Función para el trigger de la sentinela
def sentinela_trigger():
    print("¡Trigger de la sentinela activado!")

# Iniciar la sentinela
sentinela = Sentinela(interval=10, callback=sentinela_trigger)
sentinela.start()

# Iniciar la escucha de voz en un hilo separado
whisper_thread = threading.Thread(target=whisper_listener.listen, args=(handle_transcription,))
whisper_thread.start()

# Manejar la entrada de texto desde la terminal
def listen_to_console():
    while True:
        user_input = input("Tú (Texto): ")
        if user_input.lower() == "stop":
            openai_responder.stop()
        else:
            threading.Thread(target=openai_responder.respond, args=(user_input, print)).start()

# Iniciar la escucha de la terminal en un hilo separado
console_thread = threading.Thread(target=listen_to_console)
console_thread.start()

# Mantener el programa en ejecución
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("Deteniendo el asistente...")
    whisper_listener.stop()
    sentinela.stop()
    whisper_thread.join()
    console_thread.join()