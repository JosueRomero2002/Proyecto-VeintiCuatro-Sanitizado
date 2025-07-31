import speech_recognition as sr
from APIs.amazonPolly import *
from APIs.autoGPT import *


# Crear un objeto de reconocimiento
r = sr.Recognizer()

Polly = Polly()
AutoGPT = AutoGPT()


# Utilizar el micrófono como fuente de audio
while True:
    with sr.Microphone() as source:
        print("Di algo!")
        audio = r.listen(source)

    # Transcribir el audio a texto
    try:
        text = r.recognize_google(audio, language="es-ES")
        #  r.recognize_google(audio, language='es-ES')
        print("Has dicho: " + text)

        Polly.myPlayAudio(text)

        # Salir del bucle si se dice "finalizar prueba"
        if text.lower() == "finalizar prueba":
            break

        if (
            "pregunta" in text.lower()
            or "cuantos" in text.lower()
            or "quienes" in text.lower()
            or "cuales" in text.lower()
            or "porque" in text.lower()
            or "por que" in text.lower()
        ):
            Polly.myPlayAudio(AutoGPT.get_response(text))

    except sr.UnknownValueError:
        print("No se pudo entender lo que dijiste")
    except sr.RequestError as e:
        print(
            "No se pudo conectar con el servicio de reconocimiento de voz: {0}".format(
                e
            )
        )

        # Procesar el input: El procesamiento de input tiene el objetivo de manejar a donde debe dirigirse el programa para considerar la respuesta adecuada.

        # Palabras clave para filtrar en Tipos de Peticion
        # Peticion de Accion: El programa debe hacer algo (mandar whatsapp, enviar un correo)
        # Peticion de Respuesta: El programa debe responder a una pregunta
        # Respuesta Web: La respuesta puede ser encontrada en la web o requiere analisis
        # Respuesta Analitica: La respuesta requiere de un analisis profundo
        # Respuesta en Memoria: La respuesta se encuentra dentro de la memoria del programa
        # Respuesta Combinada: La respuesta puede ser web, analitica o y de memoria.
