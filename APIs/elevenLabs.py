from elevenlabs import Voice, VoiceSettings, play
from elevenlabs.client import ElevenLabs

from elevenlabs import stream


# Dotenv
from dotenv import load_dotenv
import os

load_dotenv()

client = ElevenLabs(api_key=os.getenv("ELEVENLABS_API_KEY"))

audio = client.generate(
    text="En el contexto de la librería ElevenLabs, el término stream se refiere al proceso de manejar audio generado en tiempo real, en lugar de esperar a que toda la generación del audio esté completa antes de reproducirlo o procesarlo. ",
    voice=Voice(
        voice_id="XA2bIQ92TabjGbpO2xRr",
        settings=VoiceSettings(
            stability=0.71, similarity_boost=0.5, style=0.0, use_speaker_boost=True
        ),
    ),
    stream=True,
)

play(audio)
