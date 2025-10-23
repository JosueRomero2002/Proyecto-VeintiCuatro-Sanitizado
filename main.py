import time
import os
import threading
import speech_recognition as sr
import subprocess
import keyboard
from Unidad_Accion.Accion_NoInvasiva.Toasts import ToastifyWindows
from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth
from APIs.autoGPT import AutoGPT  # Importamos la clase AutoGPT desde el archivo que contenga esta clase

# Inicializamos Toastify, reconocimiento de voz y la clase AutoGPT
ToastifyWindows = ToastifyWindows()
r = sr.Recognizer()
AutoGPT = AutoGPT()

# Variables de control
stop_response = False









voice_mode_active = False  # Controla si el modo de voz está activo o en modo silencioso
voice_thread = None  # Variable para el hilo de voz
lista_guardada = None  # Para monitoreo de cambios en la lista de SharePoint
execution_lock = threading.Lock()  # Lock para controlar la ejecución simultánea






# Variable global para la autenticación de SharePoint
sharepoint_auth = None

# Función para inicializar la autenticación de SharePoint
def inicializar_sharepoint():
    global sharepoint_auth
    if sharepoint_auth is None:
        print("Inicializando autenticación interactiva con SharePoint...")
        sharepoint_auth = SharePointInteractiveAuth()
        if not sharepoint_auth.authenticate_interactive():
            raise Exception("No se pudo autenticar con SharePoint")
        print("Autenticación con SharePoint exitosa")
    return sharepoint_auth

# Función para obtener los datos de la lista de SharePoint
def obtener_datos_lista():
    auth = inicializar_sharepoint()
    return auth.get_list_items('Tutorias', ['ID', 'Aula', 'Tipo de Tutoria', 'Contactado','Estado', 'Telefono', 'Nombre Tutor', 'Fecha de Tutoria', 'Hora Tutoria', 'Clases','Temas','Alumnos', 'TutoresRechazaron'])

# Monitorear cambios en la lista de SharePoint en un hilo separado
def detectar_cambios():
    
    global lista_guardada
    lista_guardada = obtener_datos_lista()  # Guardar los datos iniciales de la lista
    while True:
        


        lista_actual = obtener_datos_lista()
        if lista_actual != lista_guardada:
            lista_guardada = lista_actual
            lanzar_notificacion()
        time.sleep(10)  # Verificar cada 10 segundos

# Función que maneja la notificación y escucha si se permite la acción
def lanzar_notificacion():
    ejecutar_algoritmo_tutorias()
    ToastifyWindows.sendMessagetoToast("Se detectó un cambio en la lista de SharePoint. ¿Permitir acción?")
    # escuchar_respuesta()

# Función que maneja la respuesta de la IA
def process_response(prompt):
    global stop_response
    response = "Not Yet Implemented"
    # response = AutoGPT.get_response(prompt)
    for chunk in response.split('. '):  # Dividimos la respuesta en trozos
        if stop_response:
            print("Respuesta interrumpida.")
            ToastifyWindows.sendMessagetoToast("Respuesta interrumpida.")
            break
        print(chunk + '.')
        time.sleep(2)

# Función para escuchar voz y enviar preguntas
def listen_to_voice():
    global stop_response, voice_mode_active
    while True:
        if not voice_mode_active:
            time.sleep(1)
            continue

        with sr.Microphone() as source:
            print("Escuchando... Di algo!")
            audio = r.listen(source)

        try:
            text = r.recognize_google(audio, language='es-ES')
            print("Voz detectada: " + text)

            if text.lower() == "permitir":
                ejecutar_contactar_tutorias()
            elif text.lower() == "stop":
                stop_response = True
                print("Interrupción por voz.")
        except sr.UnknownValueError:
            print("No se pudo entender lo que dijiste.")
        except sr.RequestError as e:
            print(f"Error en el servicio de reconocimiento de voz: {e}")

# Función para manejar entradas por consola
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

# Función para ejecutar AlgoritmoTutorias.py y luego ContactarTutorias.py
def ejecutar_contactar_tutorias():
    ruta_tests = os.path.join(os.path.dirname(__file__), 'Tests')

    with execution_lock:  # Bloquea para que nadie más pueda ejecutar mientras esto está corriendo
        print("Ejecutando ContactarTutorias.py...")
        subprocess.run(["python", os.path.join(ruta_tests, "ContactarTutorias.py")])

def ejecutar_algoritmo_tutorias():
    ruta_tests = os.path.join(os.path.dirname(__file__), 'Tests')

    with execution_lock:  # Bloquea para que nadie más pueda ejecutar mientras esto está corriendo
        print("Ejecutando AlgoritmoTutorias.py...")
        subprocess.run(["python", os.path.join(ruta_tests, "AlgoritmoTutorias.py")], check=True)
     
# Función para manejar interrupciones mediante F12
def on_key_event(event):
    global stop_response
    if event.name == 'f12':  # Usar F12 como trigger
        stop_response = True
        print("Interrupción por tecla F12.")

# Registrar el listener para la tecla F12
keyboard.on_press(on_key_event)



# Función principal que gestiona la ejecución de hilos
def main():

    ejecutar_algoritmo_tutorias()
    global voice_thread

    # Crear hilos para monitoreo de lista, voz y consola
    sharepoint_thread = threading.Thread(target=detectar_cambios)
    voice_thread = threading.Thread(target=listen_to_voice)
    console_thread = threading.Thread(target=listen_to_console)

    # Iniciar todos los hilos
    sharepoint_thread.start()
    voice_thread.start()
    console_thread.start()

    # Esperar a que terminen
    sharepoint_thread.join()
    voice_thread.join()
    console_thread.join()

# Iniciar la ejecución
main()




# import speech_recognition as sr
# import threading
# import time
# import keyboard
# from APIs.autoGPT import AutoGPT  # Importamos la clase AutoGPT desde el archivo que contenga esta clase

# # Inicializamos el reconocimiento de voz y la clase AutoGPT
# r = sr.Recognizer()
# AutoGPT = AutoGPT()

# # Variables de control
# stop_response = False
# voice_mode_active = True  # Controla si el modo de voz está activo o en modo silencioso
# voice_thread = None  # Variable para el hilo de voz

# # Función que maneja la respuesta de la IA
# def process_response(prompt):
#     global stop_response
#     response = AutoGPT.get_response(prompt)
#     for chunk in response.split('. '):  # Dividimos la respuesta en trozos
#         if stop_response:
#             print("Respuesta interrumpida.")
#             break
#         print(chunk + '.')  # Mostramos cada trozo de la respuesta
#         time.sleep(2)  # Simula el tiempo entre partes de la respuesta

# # Función para escuchar voz y enviar preguntas
# def listen_to_voice():
#     global stop_response
#     global voice_mode_active
#     while True:
#         if not voice_mode_active:
#             time.sleep(1)  # Si el modo de voz está desactivado, espera antes de reintentar
#             continue

#         with sr.Microphone() as source:
#             print("Escuchando... Di algo!")
#             audio = r.listen(source)

#         try:
#             text = r.recognize_google(audio, language='es-ES')
#             print("Voz detectada: " + text)

#             if text.lower() == "stop":
#                 stop_response = True
#                 print("Interrupción por voz.")
#                 continue

#             # Reiniciar la señal de interrupción y lanzar la respuesta de la IA en un hilo separado
#             stop_response = False
#             response_thread = threading.Thread(target=process_response, args=(text,))
#             response_thread.start()

#         except sr.UnknownValueError:
#             print("No se pudo entender lo que dijiste.")
#         except sr.RequestError as e:
#             print(f"Error en el servicio de reconocimiento de voz: {e}")

# # Función para manejar entradas por consola
# def listen_to_console():
#     global stop_response
#     global voice_mode_active
#     global voice_thread

#     while True:
#         user_input = input("Tú (Texto): ")

#         if user_input.lower() == "exit":
#             print("¡Adiós!")
#             break

#         if user_input.lower() == "stop":
#             stop_response = True
#             print("Interrupción por consola.")
#             continue

#         if user_input.lower() == "modo silencioso":
#             voice_mode_active = False
#             print("Modo silencioso activado. El reconocimiento de voz está pausado.")
#             continue


#         if user_input.lower() == "modo de voz":
#             if not voice_mode_active:
#                 voice_mode_active = True
#                 print("Modo de voz activado. El reconocimiento de voz se reanuda.")
#             else:
#                 print("El modo de voz ya está activado.")
#             continue

#         # Reiniciar la señal de interrupción y lanzar la respuesta de la IA en un hilo separado
#         stop_response = False
#         response_thread = threading.Thread(target=process_response, args=(user_input,))
#         response_thread.start()

# # Función para manejar interrupciones mediante F12
# def on_key_event(event):
#     global stop_response
#     if event.name == 'f12':  # Usar F12 como trigger
#         stop_response = True
#         print("Interrupción por tecla F12.")

# # Registrar el listener para la tecla F12
# keyboard.on_press(on_key_event)

# # Función principal que gestiona la ejecución de hilos
# def main():
#     global voice_thread

#     # Crear hilos para la voz y la consola
#     voice_thread = threading.Thread(target=listen_to_voice)
#     console_thread = threading.Thread(target=listen_to_console)

#     # Iniciar ambos hilos
#     voice_thread.start()
#     console_thread.start()

#     # Esperar a que ambos terminen
#     voice_thread.join()
#     console_thread.join()

# # Iniciar la ejecución
# main()




# import speech_recognition as sr
# import threading
# import time
# import keyboard

# # Simulación de AutoGPT para responder preguntas
# class AutoGPT:
#     def get_response(self, prompt):
#         # Simulación de respuestas largas
#         return "Esta es una respuesta larga de la IA que se dividirá en varias partes para que puedas interrumpirla si lo deseas. " \
#                "El objetivo es que el sistema funcione de manera más interactiva, simulando un diálogo en tiempo real. " \
#                "Puedes usar la palabra 'stop' para interrumpir cuando quieras."

# AutoGPT = AutoGPT()
# r = sr.Recognizer()

# # Variable de control para interrupciones
# stop_response = False

# # Función que maneja la respuesta de la IA
# def process_response(prompt):
#     global stop_response
#     response = AutoGPT.get_response(prompt)
#     for chunk in response.split('. '):  # Dividimos la respuesta en trozos
#         if stop_response:
#             print("Respuesta interrumpida.")
#             break
#         print(chunk + '.')
#         time.sleep(2)  # Simula el tiempo entre partes de la respuesta

# # Función para escuchar voz y enviar preguntas
# def listen_to_voice():
#     global stop_response
#     while True:
#         with sr.Microphone() as source:
#             print("Escuchando... Di algo!")
#             audio = r.listen(source)

#         try:
#             text = r.recognize_google(audio, language='es-ES')
#             print("Voz detectada: " + text)

#             if text.lower() == "stop":
#                 stop_response = True
#                 print("Interrupción por voz.")
#                 continue

#             # Reiniciar la señal de interrupción y lanzar la respuesta de la IA en un hilo separado
#             stop_response = False
#             response_thread = threading.Thread(target=process_response, args=(text,))
#             response_thread.start()

#         except sr.UnknownValueError:
#             print("No se pudo entender lo que dijiste.")
#         except sr.RequestError as e:
#             print("Error en el servicio de reconocimiento de voz: {0}".format(e))

# # Función para manejar entradas por consola
# def listen_to_console():
#     global stop_response
#     while True:




#         user_input = input("Tú (Texto): ")

#         if user_input.lower() == "exit":
#             print("¡Adiós!")
#             break

#         if user_input.lower() == "stop":
#             stop_response = True
#             print("Interrupción por consola.")
#             continue

#         # Reiniciar la señal de interrupción y lanzar la respuesta de la IA en un hilo separado
#         stop_response = False
#         response_thread = threading.Thread(target=process_response, args=(user_input,))
#         response_thread.start()

# # Función para manejar interrupciones mediante F12
# def on_key_event(event):
#     global stop_response
#     if event.name == 'f12':  # Usar F12 como trigger
#         stop_response = True
#         print("Interrupción por tecla F12.")

# # Registrar el listener para la tecla F12
# keyboard.on_press(on_key_event)

# # Función principal que gestiona la ejecución de hilos
# def main():
#     # Crear hilos para la voz y la consola
#     voice_thread = threading.Thread(target=listen_to_voice)
#     console_thread = threading.Thread(target=listen_to_console)

#     # Iniciar ambos hilos
#     voice_thread.start()
#     console_thread.start()

#     # Esperar a que ambos terminen
#     voice_thread.join()
#     console_thread.join()

# # Iniciar la ejecución
# main()


# -----------------------------------------------------------------------

# import speech_recognition as sr
# import keyboard
# import threading
# import time
# # from APIs.amazonPolly import *
# from APIs.autoGPT import *
# from Unidad_Accion.Accion_NoInvasiva.Toasts import *

# # Inicializar reconocimiento y servicios
# r = sr.Recognizer()
# AutoGPT = AutoGPT()
# ToastifyWindows = ToastifyWindows()

# # Variable de control para interrupciones
# stop_response = False

# # Función para procesar la respuesta de la IA en un hilo separado
# def process_response(text):
#     global stop_response
#     response = AutoGPT.get_response(text)
#     for chunk in response.split('. '):  # Dividir respuesta en trozos
#         if stop_response:  # Verificar si se pidió detener
#             ToastifyWindows.sendMessagetoToast("Respuesta interrumpida")
#             break
#         ToastifyWindows.toastAndReproduceMessage(chunk)
#         time.sleep(1)  # Simular tiempo de procesamiento/habla

# # Función para escuchar el micrófono y manejar preguntas
# def listen_and_respond():
#     global stop_response
#     while True:
#         with sr.Microphone() as source:
#             print("Di algo!")
#             audio = r.listen(source)
        
#         # Transcribir audio a texto
#         try:
#             text = r.recognize_google(audio, language='es-ES')
#             print("Has dicho: " + text)
#             ToastifyWindows.sendMessagetoToast(text)

#             # Si se dice "finalizar prueba", salir
#             if text.lower() == "finalizar prueba":
#                 stop_response = True
#                 break

#             # Si detectamos una pregunta, iniciar la respuesta en otro hilo
#             if any(keyword in text.lower() for keyword in ["pregunta", "cuántos", "quiénes", "cuáles", "por qué"]):
#                 stop_response = False
#                 response_thread = threading.Thread(target=process_response, args=(text,))
#                 response_thread.start()

#         except sr.UnknownValueError:
#             print("No se pudo entender lo que dijiste")
#         except sr.RequestError as e:
#             print("Error de servicio de reconocimiento de voz: {0}".format(e))

# # Función para manejar interrupciones con F12
# def on_key_event(event):
#     global stop_response
#     if event.name == 'f12':  # Usar F12 como trigger
#         stop_response = True  # Señalar que se debe detener la respuesta
#         ToastifyWindows.sendMessagetoToast("Interrupción detectada. Deteniendo respuesta.")

# # Registrar el listener para F12
# keyboard.on_press(on_key_event)

# # Iniciar el ciclo de escucha
# listen_and_respond()





# ========================================================================================



# import speech_recognition as sr
# #from APIs.amazonPolly import *
# from APIs.autoGPT import *
# import keyboard
# from Unidad_Accion.Accion_NoInvasiva.Toasts import *

# # Crear un objeto de reconocimiento
# r = sr.Recognizer()

# #Polly = Polly()
# AutoGPT = AutoGPT()
# ToastifyWindows = ToastifyWindows()



# def on_key_event(event):
#     if event.name == 'f12':  # Usar F12 como tecla trigger
#         Output = ToastifyWindows.inputToast("Se presionó la tecla F12")
        
#         while "close" != str(Output['message']):
#               ToastifyWindows.inputToast(AutoGPT.get_response(Output['message']))
        
        


# # Registrar el keylistener
# keyboard.on_press(on_key_event)

# #Utilizar el micrófono como fuente de audio
# while True:
#     with sr.Microphone() as source:
#         print("Di algo!")
#         audio = r.listen(source)

    
    




#     #Transcribir el audio a texto
#     try:
#         text = r.recognize_google(audio, language='es-ES')
#         r.recognize_google(audio, language='es-ES')
#        # text = "Informacion de Prueba: ¿Cuántos años tiene el presidente de México?"
#         print("Has dicho: " + text)

#        # Polly.myPlayAudio(text)
#         ToastifyWindows.sendMessagetoToast(text)

#         #Salir del bucle si se dice "finalizar prueba"
#         if text.lower() == "finalizar prueba":
#              break


#         if "pregunta" in text.lower() or "cuantos" in text.lower() or "quienes" in text.lower() or "cuales" in text.lower() or "porque" in text.lower() or "por que" in text.lower() :
#         # Polly.myPlayAudio(AutoGPT.get_response(text))
#             ToastifyWindows.toastAndReproduceMessage(AutoGPT.get_response(text))




#     except sr.UnknownValueError:
#         print("No se pudo entender lo que dijiste")
#     except sr.RequestError as e:
#         print("No se pudo conectar con el servicio de reconocimiento de voz: {0}".format(e))















































































