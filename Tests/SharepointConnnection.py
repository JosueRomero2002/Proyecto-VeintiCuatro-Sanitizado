import pywhatkit
from datetime import datetime, timedelta
import time
import pyautogui
import keyboard as k

from shareplum import Site
from shareplum import Office365


class SharepointConnection:
    def __init__(self):
        

    def get_response(self, input_message):
        self.chat_history.append({"role": "user", "content": input_message})

        messages = [{"role": "system", "content": "dame una respuesta de al menos 2 lineas"}]
        messages.extend(self.chat_history)

        # response = openai.ChatCompletion.create(
        #     model=self.model,
        #     messages=messages,
        #     temperature=self.temperature,
        #     max_tokens=self.max_tokens,
        # )
        response = openai.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=self.temperature,
            max_tokens=self.max_tokens,
        )


        respuesta = response.choices[0].message.content
        
        self.chat_history.append({"role": "assistant", "content": respuesta})

        return respuesta

    def setautogpt(self, model, temperature, max_tokens):
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        return f"Modelo de GPT actualizado a: {model} con temperatura: {temperature} y max tokens: {max_tokens}"
