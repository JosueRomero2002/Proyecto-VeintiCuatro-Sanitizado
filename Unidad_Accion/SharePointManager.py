from shareplum import Site
from shareplum import Office365

class SharePointManager:
    def __init__(self, username, password, url_site, url_base):
        """
        Inicializa la conexión con SharePoint.
        :param username: Usuario de SharePoint.
        :param password: Contraseña de SharePoint.
        :param url_site: URL completa del sitio de SharePoint (ej. 'https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/').
        :param url_base: URL base de SharePoint (ej. 'https://unitechn.sharepoint.com').
        """
        self.username = username
        self.password = password
        self.url_site = url_site
        self.url_base = url_base
        self.authcookie = self._autenticar()
        self.site = Site(self.url_site, authcookie=self.authcookie)

    def _autenticar(self):
        """Autentica con SharePoint y devuelve la cookie de autenticación."""
        authcookie = Office365(self.url_base, username=self.username, password=self.password).GetCookies()
        return authcookie

    def obtener_lista(self, nombre_lista):
        """Obtiene una lista de SharePoint por su nombre."""
        return self.site.List(nombre_lista)

    def obtener_items_lista(self, nombre_lista, campos=None):
        """
        Obtiene los items de una lista de SharePoint.
        :param nombre_lista: Nombre de la lista de SharePoint.
        :param campos: Campos específicos a obtener.
        :return: Items de la lista.
        """
        lista = self.obtener_lista(nombre_lista)
        return lista.GetListItems(fields=campos)

# Ejemplo de uso:
# sp_manager = SharePointManager(username='usuario', password='contraseña', 
#                                url_site='https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/',
#                                url_base='https://unitechn.sharepoint.com')

# Tutoriasdata = sp_manager.obtener_items_lista('Tutorias', campos=['ID', 'Aula', 'Tipo de Tutoria', 'Estado'])
