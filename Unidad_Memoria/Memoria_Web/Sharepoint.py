import os
import sys

# Add parent directory to path to import SharePointInteractiveAuth
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth

class MySharepoint:
    def __init__(self):
        """
        Inicializa SharePoint con autenticación interactiva.
        Se abrirá el navegador para autenticación.
        """
        print("Inicializando SharePoint con autenticación interactiva...")
        self.auth = SharePointInteractiveAuth()
        if not self.auth.authenticate_interactive():
            raise Exception("No se pudo autenticar con SharePoint")
        print("SharePoint inicializado exitosamente")

    def getSharepointListData(self, ListName, fields):
        """
        Obtiene datos de una lista de SharePoint usando autenticación interactiva.
        :param ListName: Nombre de la lista de SharePoint.
        :param fields: Campos específicos a obtener.
        :return: Datos de la lista.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.get_list_items(ListName, fields)
    
    def setSharepontProperties(self, user, password):
        """
        Método de compatibilidad - ya no se usa autenticación básica.
        La autenticación interactiva no requiere credenciales hardcodeadas.
        """
        print("Nota: Se está usando autenticación interactiva, no se requieren credenciales hardcodeadas")
        return "Autenticación interactiva configurada exitosamente"
    
    def createListItem(self, list_name, item_data):
        """
        Crea un nuevo item en una lista de SharePoint.
        :param list_name: Nombre de la lista.
        :param item_data: Datos del item a crear.
        :return: Resultado de la creación.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.create_list_item(list_name, item_data)
    
    def updateListItem(self, list_name, item_id, item_data):
        """
        Actualiza un item existente en una lista de SharePoint.
        :param list_name: Nombre de la lista.
        :param item_id: ID del item a actualizar.
        :param item_data: Nuevos datos del item.
        :return: Resultado de la actualización.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.update_list_item(list_name, item_id, item_data)
