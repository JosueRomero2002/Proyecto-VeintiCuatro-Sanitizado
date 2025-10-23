import openai
import sounddevice as sd
import wave
import threading
import time
from dotenv import load_dotenv
import os

load_dotenv()

class WhisperListener:
    def __init__(self, api_key):
        self.api_key = api_key
        openai.api_key = api_key
        self.listening = False
        self.stop_listening = False

    def record_audio(self, output_file, duration, sample_rate=44100):
        """Grabar audio desde el micrófono."""
        print("Escuchando... Di algo!")
        audio_data = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype='int16')
        sd.wait()  # Esperar hasta que termine la grabación
        print("Grabación finalizada.")

        # Guardar el audio en un archivo
        with wave.open(output_file, 'wb') as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)  # 2 bytes para 'int16'
            wf.setframerate(sample_rate)
            wf.writeframes(audio_data.tobytes())

    def transcribe_audio(self, audio_file_path):
        """Transcribir el audio usando Whisper."""
        if not os.path.exists(audio_file_path):
            raise FileNotFoundError(f"Archivo de audio no encontrado: {audio_file_path}")

        try:
            with open(audio_file_path, "rb") as audio_file:
                response = openai.Audio.transcribe(
                    model="whisper-1",
                    file=audio_file
                )
                return response.get("text", "No se pudo transcribir el audio.")
        except Exception as e:
            return f"Error transcribiendo audio: {str(e)}"

    def listen(self, callback):
        """Escuchar continuamente y transcribir el audio."""
        self.listening = True
        while not self.stop_listening:
            audio_file = "temp_audio.wav"
            self.record_audio(audio_file, duration=5)  # Grabar 5 segundos de audio
            transcription = self.transcribe_audio(audio_file)
            if transcription and not self.stop_listening:
                callback(transcription)  # Llamar al callback con la transcripción
            time.sleep(1)  # Esperar 1 segundo antes de la siguiente grabación

    def stop(self):
        """Detener la escucha."""
        self.stop_listening = True
        self.listening = False