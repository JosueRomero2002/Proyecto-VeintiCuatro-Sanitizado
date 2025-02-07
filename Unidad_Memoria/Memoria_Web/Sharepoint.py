from shareplum import Site
from shareplum import Office365

class MySharepoint:
    def __init__(self):
        self.authcookie = Office365('https://unitechn.sharepoint.com', username='josuepinro21@unitec.edu', password='Duque21J').GetCookies()
        self.site = Site('https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/', authcookie=self.authcookie)
        return ("Inicializacion de Sharepoint")

    def getSharepointListData(self, ListName, fields):
        sp_list = self.site.List(ListName)
        data = sp_list.GetListItems(fields=fields)
        return data
       
    
    def setSharepontProperties(self, user, password):

        return "mensaje enviado con exito"
