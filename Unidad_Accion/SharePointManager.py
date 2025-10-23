from SharePointInteractiveAuth import SharePointInteractiveAuth
import os
import sys

# Add parent directory to path to import config
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import *

class SharePointManager:
    def __init__(self, use_interactive_auth=True):
        """
        Inicializa la conexión con SharePoint usando autenticación interactiva.
        :param use_interactive_auth: Si usar autenticación interactiva (recomendado).
        """
        self.use_interactive_auth = use_interactive_auth
        self.auth = None
        
        if use_interactive_auth:
            print("Inicializando SharePoint con autenticación interactiva...")
            self.auth = SharePointInteractiveAuth()
            if not self.auth.authenticate_interactive():
                raise Exception("No se pudo autenticar con SharePoint usando autenticación interactiva")
        else:
            raise Exception("Solo se soporta autenticación interactiva. Use use_interactive_auth=True")

    def obtener_lista(self, nombre_lista):
        """
        Obtiene una lista de SharePoint por su nombre.
        Nota: En la nueva implementación, esto se maneja internamente.
        """
        print(f"Accediendo a la lista: {nombre_lista}")
        return self.auth

    def obtener_items_lista(self, nombre_lista, campos=None):
        """
        Obtiene los items de una lista de SharePoint usando autenticación interactiva.
        :param nombre_lista: Nombre de la lista de SharePoint.
        :param campos: Campos específicos a obtener.
        :return: Items de la lista.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.get_list_items(nombre_lista, campos)

    def crear_item_lista(self, nombre_lista, datos_item):
        """
        Crea un nuevo item en una lista de SharePoint.
        :param nombre_lista: Nombre de la lista de SharePoint.
        :param datos_item: Diccionario con los datos del item.
        :return: Resultado de la creación.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.create_list_item(nombre_lista, datos_item)

    def actualizar_item_lista(self, nombre_lista, item_id, datos_item):
        """
        Actualiza un item existente en una lista de SharePoint.
        :param nombre_lista: Nombre de la lista de SharePoint.
        :param item_id: ID del item a actualizar.
        :param datos_item: Diccionario con los nuevos datos del item.
        :return: Resultado de la actualización.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.update_list_item(nombre_lista, item_id, datos_item)

    def obtener_listas_disponibles(self):
        """
        Obtiene las listas disponibles en el sitio de SharePoint.
        :return: Lista de listas disponibles.
        """
        if not self.auth:
            raise Exception("No hay conexión autenticada con SharePoint")
        
        return self.auth.get_available_lists()

# Ejemplo de uso:
# sp_manager = SharePointManager(use_interactive_auth=True)
# Tutoriasdata = sp_manager.obtener_items_lista('Tutorias', campos=['ID', 'Aula', 'Tipo de Tutoria', 'Estado'])
