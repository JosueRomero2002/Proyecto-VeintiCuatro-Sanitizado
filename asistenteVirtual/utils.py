import os
from dotenv import load_dotenv

load_dotenv()

def get_env_variable(key):
    """Obtener una variable de entorno."""
    return os.getenv(key)