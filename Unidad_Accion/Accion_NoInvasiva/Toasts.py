from win11toast import toast

class ToastifyWindows:
    def __init__(self):
        print ("Inicializacion de ToastifyWindows")

  

    def sendMessagetoToast(self, text):
        toast(text)
        return "Toast Creado Con Exito"
    
    def inputToast(self, text):
        result = toast(text, 'Type anything', input='reply', button='Send')
        return {"status": "Toast Creado Con Exito", "message": result['user_input']['reply']}
    

    def toastAndReproduceMessage(self, text):
        toast(text, dialogue=text)
        return "Toast Creado Con Exito"
    

