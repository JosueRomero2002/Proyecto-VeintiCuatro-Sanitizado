import openai
from dotenv import load_dotenv
import os


class AutoGPT:
    def __init__(self):
        load_dotenv()
        ApiKey = os.getenv("AUTOGPT_API_KEY")
        openai.api_key = ApiKey
        self.model = "gpt-4-turbo-2024-04-09"
        self.temperature = 0.8
        self.max_tokens = 150
        self.chat_history = []

    def get_response(self, input_message):
        self.chat_history.append({"role": "user", "content": input_message})

        messages = [
            {"role": "system", "content": "dame una respuesta de al menos 2 lineas"}
        ]
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
