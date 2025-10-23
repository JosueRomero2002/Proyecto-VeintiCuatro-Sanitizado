"""
Autenticación interactiva con SharePoint usando Microsoft Graph API
Esta es la alternativa más moderna y segura
"""

import requests
from msal import PublicClientApplication
import os
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()


class SharePointInteractiveAuth:
    def __init__(self):
        # Configuración para autenticación interactiva
        self.client_id = (
            "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Microsoft Graph PowerShell
        )
        self.tenant_id = "common"  # Usar 'common' para multi-tenant
        self.scope = ["https://graph.microsoft.com/.default"]
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        # URLs de SharePoint
        self.sharepoint_url = os.getenv(
            "SHAREPOINT_URL", "https://unitechn.sharepoint.com"
        )
        self.site_url = os.getenv(
            "SHAREPOINT_SITE_URL",
            "https://unitechn.sharepoint.com/sites/TutoriasUNITEC2",
        )

        # Variables de estado
        self.app = None
        self.access_token = None
        self.site_id = None

    def authenticate_interactive(self):
        """Autenticación interactiva usando navegador web"""
        try:
            print("Iniciando autenticación interactiva...")
            print("Se abrirá tu navegador para autenticarte...")

            # Crear aplicación MSAL
            self.app = PublicClientApplication(
                client_id=self.client_id, authority=self.authority
            )

            # Obtener token interactivo
            result = self.app.acquire_token_interactive(scopes=self.scope)

            if "access_token" in result:
                self.access_token = result["access_token"]
                print("Autenticación interactiva exitosa")
                return True
            else:
                print(
                    f"Error en autenticación: {result.get('error_description', 'Error desconocido')}"
                )
                return False

        except Exception as e:
            print(f"Error en autenticación interactiva: {str(e)}")
            return False

    def get_site_id(self):
        """Obtener el ID del sitio de SharePoint"""
        if not self.access_token:
            print("No hay token de acceso. Autentica primero.")
            return None

        try:
            # Convertir URL del sitio a formato requerido por Graph API
            site_url = self.site_url.replace("https://", "").replace("/", ":")
            url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            site_data = response.json()
            self.site_id = site_data.get("id")
            print(f"ID del sitio obtenido: {self.site_id}")
            return self.site_id

        except Exception as e:
            print(f"Error obteniendo ID del sitio: {str(e)}")
            return None

    def get_list_items(self, list_name, fields=None):
        """Obtener elementos de una lista usando Graph API"""
        if not self.access_token:
            print("No hay token de acceso. Autentica primero.")
            return None

        if not self.site_id:
            if not self.get_site_id():
                return None

        try:
            # Primero obtener el ID de la lista
            list_id = self._get_list_id(list_name)
            if not list_id:
                print(f"No se pudo encontrar la lista '{list_name}'")
                return None

            # Construir URL con campos específicos si se proporcionan
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{list_id}/items"

            # Agregar parámetros de consulta
            params = []
            # Siempre expandir fields para obtener los datos
            params.append("expand=fields")

            if params:
                url += "?" + "&".join(params)

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            items_data = response.json()
            items = items_data.get("value", [])

            # Procesar los elementos para que tengan el formato esperado
            processed_items = []
            for item in items:
                processed_item = {"ID": item.get("id")}
                if "fields" in item:
                    for field_name, field_value in item["fields"].items():
                        processed_item[field_name] = field_value
                processed_items.append(processed_item)

            print(
                f"Se obtuvieron {len(processed_items)} elementos de la lista '{list_name}'"
            )
            return processed_items

        except Exception as e:
            print(f"Error obteniendo elementos de la lista '{list_name}': {str(e)}")
            return None

    def _get_list_id(self, list_name):
        """Obtener el ID de una lista por su nombre"""
        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists"
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            lists_data = response.json()
            lists = lists_data.get("value", [])

            # Buscar la lista por nombre
            for lista in lists:
                if lista.get("displayName") == list_name:
                    return lista.get("id")

            return None

        except Exception as e:
            print(f"Error obteniendo ID de la lista '{list_name}': {str(e)}")
            return None

    def get_available_lists(self):
        """Obtener listas disponibles usando Graph API"""
        if not self.access_token:
            print("No hay token de acceso. Autentica primero.")
            return None

        if not self.site_id:
            if not self.get_site_id():
                return None

        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists"
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            lists_data = response.json()
            lists = lists_data.get("value", [])

            print(f"Se encontraron {len(lists)} listas disponibles")
            return lists

        except Exception as e:
            print(f"Error obteniendo listas: {str(e)}")
            return None

    def create_list_item(self, list_name, item_data):
        """Crear un nuevo elemento en una lista"""
        if not self.access_token:
            print("No hay token de acceso. Autentica primero.")
            return None

        if not self.site_id:
            if not self.get_site_id():
                return None

        try:
            # Obtener el ID de la lista
            list_id = self._get_list_id(list_name)
            if not list_id:
                print(f"No se pudo encontrar la lista '{list_name}'")
                return None

            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{list_id}/items"

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            # Preparar datos para Graph API
            payload = {"fields": item_data}

            response = requests.post(url, headers=headers, json=payload)
            response.raise_for_status()

            print(f"Elemento creado exitosamente en la lista '{list_name}'")
            return response.json()

        except Exception as e:
            print(f"Error creando elemento en la lista '{list_name}': {str(e)}")
            return None

    def update_list_item(self, list_name, item_id, item_data):
        """Actualizar un elemento existente en una lista"""
        if not self.access_token:
            print("No hay token de acceso. Autentica primero.")
            return None

        if not self.site_id:
            if not self.get_site_id():
                return None

        try:
            # Obtener el ID de la lista
            list_id = self._get_list_id(list_name)
            if not list_id:
                print(f"No se pudo encontrar la lista '{list_name}'")
                return None

            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{list_id}/items/{item_id}"

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            }

            # Preparar datos para Graph API
            payload = {"fields": item_data}

            response = requests.patch(url, headers=headers, json=payload)
            response.raise_for_status()

            print(
                f"Elemento {item_id} actualizado exitosamente en la lista '{list_name}'"
            )
            return response.json()

        except Exception as e:
            print(
                f"Error actualizando elemento {item_id} en la lista '{list_name}': {str(e)}"
            )
            return None


# Función de compatibilidad para reemplazar shareplum
def create_sharepoint_connection_interactive():
    """Crear conexión interactiva con SharePoint"""
    print("Creando conexión interactiva con SharePoint...")
    print("NOTA: Se abrirá tu navegador para autenticarte")

    auth = SharePointInteractiveAuth()
    if auth.authenticate_interactive():
        print("Conexión establecida exitosamente")
        return auth
    else:
        print("No se pudo establecer la conexión")
        return None
