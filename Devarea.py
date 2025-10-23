import time
import os
import threading
import openai
import subprocess
import keyboard
from pydub import AudioSegment
from pydub.playback import play
from tempfile import NamedTemporaryFile
from Unidad_Accion.Accion_NoInvasiva.Toasts import ToastifyWindows

from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth
from APIs.autoGPT import AutoGPT

# Inicialización de APIs y utilidades
ToastifyWindows = ToastifyWindows()
AutoGPT = AutoGPT()
openai.api_key = ""

# Variables de control
stop_response = False
voice_mode_active = False
voice_thread = None
lista_guardada = None
execution_lock = threading.Lock()

# Función para grabar audio con pydub y transcribirlo usando Whisper
def grabar_y_transcribir():
    from pydub import AudioSegment
    from pydub.playback import play
    import sounddevice as sd
    import soundfile as sf

    print("Grabando... (Di algo y espera unos segundos)")
    fs = 44100  # Frecuencia de muestreo
    duration = 5  # Duración en segundos
    filename = "audio.wav"

    # Grabación del audio
    myrecording = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype='int16')
    sd.wait()
    sf.write(filename, myrecording, fs)

    print("Grabación completa. Procesando...")

    # Transcripción con Whisper
    with open(filename, "rb") as audio_file:
        response = openai.Audio.transcribe("whisper-1", audio_file)
        text = response['text']
        return text

# Resto del flujo principal del código
def process_response(prompt):
    global stop_response
    response = "Not Yet Implemented"
    # response = AutoGPT.get_response(prompt)
    for chunk in response.split('. '):
        if stop_response:
            print("Respuesta interrumpida.")
            ToastifyWindows.sendMessagetoToast("Respuesta interrumpida.")
            break
        print(chunk + '.')
        time.sleep(2)

def listen_to_voice():
    global stop_response, voice_mode_active
    while True:
        if not voice_mode_active:
            time.sleep(1)
            continue

        try:
            text = grabar_y_transcribir()
            print("Voz detectada: " + text)

            if text.lower() == "permitir":
                ejecutar_contactar_tutorias()
            elif text.lower() == "stop":
                stop_response = True
                print("Interrupción por voz.")
        except Exception as e:
            print(f"Error al procesar el audio: {e}")

def listen_to_console():
    global stop_response, voice_mode_active, voice_thread

    while True:
        user_input = input("Tú (Texto): ")

        if user_input.lower() == "permitirc":
            ejecutar_contactar_tutorias()

        elif user_input.lower() == "stop":
            stop_response = True
            print("Interrupción por consola.")

        elif user_input.lower() == "modo silencioso":
            voice_mode_active = False
            print("Modo silencioso activado. El reconocimiento de voz está pausado.")

        elif user_input.lower() == "modo de voz":
            if not voice_mode_active:
                voice_mode_active = True
                print("Modo de voz activado. El reconocimiento de voz se reanuda.")
            else:
                print("El modo de voz ya está activado.")
        else:
            stop_response = False
            response_thread = threading.Thread(target=process_response, args=(user_input,))
            response_thread.start()

def ejecutar_contactar_tutorias():
    ruta_tests = os.path.join(os.path.dirname(__file__), 'Tests')

    with execution_lock:
        print("Ejecutando ContactarTutorias.py...")
        subprocess.run(["python", os.path.join(ruta_tests, "ContactarTutorias.py")])

def ejecutar_algoritmo_tutorias():
    ruta_tests = os.path.join(os.path.dirname(__file__), 'Tests')

    with execution_lock:
        print("Ejecutando AlgoritmoTutorias.py...")
        subprocess.run(["python", os.path.join(ruta_tests, "AlgoritmoTutorias.py")], check=True)

def on_key_event(event):
    global stop_response
    if event.name == 'f12':
        stop_response = True
        print("Interrupción por tecla F12.")

keyboard.on_press(on_key_event)

def main():
    # ejecutar_algoritmo_tutorias()
    global voice_thread

    # sharepoint_thread = threading.Thread(target=detectar_cambios)
    voice_thread = threading.Thread(target=listen_to_voice)
    console_thread = threading.Thread(target=listen_to_console)

    # sharepoint_thread.start()
    voice_thread.start()
    console_thread.start()

    # sharepoint_thread.join()
    voice_thread.join()
    console_thread.join()

main()
