
from io import BytesIO
import boto3

import pygame

#TODO: crear el env
class Polly():

   

    def __init__(self):

     # Configurar las credenciales de acceso a AWS
         aws_access_key_id = ''
         aws_secret_access_key = ''

    # Configurar la región de AWS donde está disponible Amazon Polly
         region_name = 'us-west-2'

    # Crear un objeto de cliente de Amazon Polly
         self.polly = boto3.client('polly', 
                     aws_access_key_id=aws_access_key_id, 
                     aws_secret_access_key=aws_secret_access_key,
                     region_name=region_name)
         

         print("Inicializando Funcion Play Audio")
         self.myPlayAudio( "Inicializando Funcion Play Audio")



    def myPlayAudio(self, inputText): 
        # Sintetizar el texto en voz utilizando Amazon Polly
        response = self.polly.synthesize_speech(Text=inputText, 
                                           OutputFormat='mp3', 
                                           VoiceId='Miguel')

        # Convertir la secuencia de bytes en un objeto de archivo que admite búsqueda
        stream = BytesIO(response['AudioStream'].read())
        stream.seek(0)

        # Reproducir la salida de voz en tiempo real utilizando pygame
        self.play_audio(stream)
         
   
        
    # Reproducir la salida de voz en tiempo real utilizando pygame
    def play_audio(self,audio):
        pygame.mixer.init()
        pygame.mixer.music.load(audio)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy() == True:
            continue