import time
import threading

class Sentinela:
    def __init__(self, interval, callback):
        self.interval = interval
        self.callback = callback
        self.running = False

    def start(self):
        """Iniciar el monitoreo."""
        self.running = True
        threading.Thread(target=self.monitor).start()

    def monitor(self):
        """Monitorear continuamente y ejecutar el callback."""
        while self.running:
            self.callback()  # Ejecutar el callback
            time.sleep(self.interval)

    def stop(self):
        """Detener el monitoreo."""
        self.running = False