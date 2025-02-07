import pywhatkit

class WhatsappSender:
    def __init__(self):
        return ("Inicializacion de Whatsapp Sender")

  

    def sendMessagetoNumber(self, input_message, input_number):
        
        
        pywhatkit.sendwhatmsg_instantly(input_number, input_message, 25, False, 4)
        
        return "mensaje enviado con exito"
