import openai
from dotenv import load_dotenv
import os
import threading
import time

load_dotenv()

class OpenAIResponder:
    def __init__(self, api_key):
        self.api_key = api_key
        openai.api_key = api_key
        self.stop_response = False

    def get_response(self, prompt):
        """Obtener una respuesta de OpenAI."""
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message['content']
        except Exception as e:
            return f"Error obteniendo respuesta de OpenAI: {str(e)}"

    def respond(self, prompt, callback):
        """Responder al usuario y permitir interrupciones."""
        self.stop_response = False
        response = self.get_response(prompt)
        for chunk in response.split('. '):
            if self.stop_response:
                print("Respuesta interrumpida.")
                break
            callback(chunk + '.')  # Llamar al callback con cada parte de la respuesta
            time.sleep(2)  # Simular un tiempo entre partes de la respuesta

    def stop(self):
        """Detener la respuesta."""
        self.stop_response = True