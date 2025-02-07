import openai
import argparse
import os
import sounddevice as sd
import wave
from dotenv import load_dotenv

class WhisperAudioTranscriber:
    def __init__(self, api_key):
        """
        Initialize the WhisperAudioTranscriber with the OpenAI API key.

        :param api_key: OpenAI API key to access the Whisper service.
        """
        openai.api_key = api_key

    def transcribe_audio(self, audio_file_path):
        """
        Transcribe the audio file using Whisper API.

        :param audio_file_path: Path to the audio file.
        :return: Transcription of the audio.
        """
        if not os.path.exists(audio_file_path):
            raise FileNotFoundError(f"Audio file not found: {audio_file_path}")

        try:
            with open(audio_file_path, "rb") as audio_file:
                response = openai.Audio.transcribe(
                    model="whisper-1",
                    file=audio_file
                )
                return response.get("text", "No transcription available.")
        except Exception as e:
            return f"Error transcribing audio: {str(e)}"

    def record_audio(self, output_file, duration, sample_rate=44100):
        """
        Record audio from the microphone and save it to a file.

        :param output_file: Path to save the recorded audio file.
        :param duration: Duration of the recording in seconds.
        :param sample_rate: Sampling rate for the recording.
        """
        print("Recording... Speak now!")
        audio_data = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype='int16')
        sd.wait()  # Wait until the recording is finished
        print("Recording finished.")

        # Save the recorded audio to a file
        with wave.open(output_file, 'wb') as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)  # 2 bytes for 'int16'
            wf.setframerate(sample_rate)
            wf.writeframes(audio_data.tobytes())

if __name__ == "__main__":
    load_dotenv()
    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY not found in .env file.")

    parser = argparse.ArgumentParser(description="Transcribe audio using OpenAI Whisper API.")
    parser.add_argument("--duration", type=int, default=5, help="Duration of the recording in seconds.")
    parser.add_argument("--output_file", default="recorded_audio.wav", help="Path to save the recorded audio file.")

    args = parser.parse_args()

    transcriber = WhisperAudioTranscriber(api_key=api_key)

    # Record audio from the microphone
    transcriber.record_audio(output_file=args.output_file, duration=args.duration)

    # Transcribe the recorded audio
    transcription = transcriber.transcribe_audio(audio_file_path=args.output_file)

    print("\nTranscription:")
    print(transcription)
