import time
import hashlib
import os
import requests

class Centinela:
    def __init__(self, tipo_elemento, ruta_elemento, intervalo_segundos, callback):
        """
        Inicializa la clase Centinela.
        :param tipo_elemento: Puede ser 'web', 'codigo', 'programa' (archivo ejecutable).
        :param ruta_elemento: La URL (si es web) o ruta del archivo (si es código o programa).
        :param intervalo_segundos: El intervalo de tiempo en segundos para revisar los cambios.
        :param callback: Función que se ejecuta cuando se detecta un cambio.
        """
        self.tipo_elemento = tipo_elemento
        self.ruta_elemento = ruta_elemento
        self.intervalo_segundos = intervalo_segundos
        self.callback = callback
        self.estado_inicial = self.obtener_estado_inicial()

    def obtener_estado_inicial(self):
        """
        Obtiene el estado inicial del elemento (su hash o contenido).
        :return: Hash del contenido inicial.
        """
        if self.tipo_elemento == 'web':
            return self._hash_contenido_web(self.ruta_elemento)
        elif self.tipo_elemento in ['codigo', 'programa']:
            return self._hash_contenido_archivo(self.ruta_elemento)
        else:
            raise ValueError(f"Tipo de elemento {self.tipo_elemento} no soportado.")

    def _hash_contenido_web(self, url):
        """
        Obtiene el hash del contenido de una página web.
        :param url: URL de la página web.
        :return: Hash del contenido.
        """
        try:
            response = requests.get(url)
            contenido = response.content
            return hashlib.md5(contenido).hexdigest()
        except Exception as e:
            print(f"Error al obtener el contenido web: {e}")
            return None

    def _hash_contenido_archivo(self, ruta_archivo):
        """
        Obtiene el hash del contenido de un archivo.
        :param ruta_archivo: Ruta del archivo.
        :return: Hash del contenido.
        """
        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as archivo:
                contenido = archivo.read()
            return hashlib.md5(contenido).hexdigest()
        else:
            print(f"El archivo {ruta_archivo} no existe.")
            return None

    def monitorear(self):
        """
        Monitorea el elemento en busca de cambios según el intervalo de tiempo asignado.
        """
        while True:
            estado_actual = self.obtener_estado_inicial()
            if estado_actual != self.estado_inicial:
                print("¡Cambio detectado!")
                self.callback(self.ruta_elemento)
                self.estado_inicial = estado_actual  # Actualizar el estado inicial tras el cambio
            time.sleep(self.intervalo_segundos)

# Ejemplo de uso:
def notificar_cambio(ruta_elemento):
    print(f"Se ha detectado un cambio en el elemento: {ruta_elemento}")

# Crear una instancia de Centinela para una página web
# centinela_web = Centinela(tipo_elemento='web', ruta_elemento='https://example.com', intervalo_segundos=30, callback=notificar_cambio)

# Iniciar la monitorización (en un hilo o proceso separado para evitar bloquear el programa principal)
# centinela_web.monitorear()
