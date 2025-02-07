# from win11toast import toast
# from time import sleep
# from win11toast import notify, update_progress

# notify(progress={
#     'title': 'YouTube',
#     'status': 'Downloading...',
#     'value': '0',
#     'valueStringOverride': '0/15 videos'
# })

# for i in range(1, 15+1):
#     sleep(1)
#     update_progress({'value': i/15, 'valueStringOverride': f'{i}/15 videos'})

# update_progress({'status': 'Completed!'})


# import winsdk.windows.media.speechsynthesis as speechsynthesis
# import winsdk.windows.media as media
# import asyncio
# from win11toast import toast

# # Función para listar todas las voces disponibles
# def list_voices():
#     synthesizer = speechsynthesis.SpeechSynthesizer()
#     all_voices = speechsynthesis.SpeechSynthesizer.all_voices
#     for voice in all_voices:
#         print(f"Voice Name: {voice.display_name}, Language: {voice.language}, Gender: {voice.gender}")

# # Función para reproducir texto con una voz específica
# async def speak_text(text, voice_name):
#     synthesizer = speechsynthesis.SpeechSynthesizer()
#     all_voices = speechsynthesis.SpeechSynthesizer.all_voices
#     selected_voice = None

#     for voice in all_voices:
#         if voice.display_name == voice_name:
#             selected_voice = voice
#             break

#     if selected_voice:
#         synthesizer.voice = selected_voice
#         stream = await synthesizer.synthesize_text_to_stream_async(text)
#         player = media.MediaPlayer()
#         player.source = media.MediaSource.create_from_stream(stream, "")
#         player.play()
#         # Esperar a que termine la reproducción
#         await asyncio.sleep(stream.size / 16000)
#     else:
#         print(f"La voz '{voice_name}' no está disponible.")

# # Listar todas las voces disponibles
# list_voices()

# # Mostrar la notificación de toast
# dialogue = 'La operación tardará entre 5 a 10 minutos, por si desea continuar.'
# toast('Hello Python🐍', dialogue)

# # Reproducir el mensaje de voz
# voice_name = "Microsoft Jorge"
# asyncio.run(speak_text(dialogue, voice_name))

from win11toast import toast

text = "Hello, World!"

result = toast('Hello', 'Type anything', input='reply', button='Send')


toast(result['user_input']['reply'])

