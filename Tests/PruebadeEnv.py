from dotenv import load_dotenv
import os

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# Ahora puedes acceder a las variables de entorno
secret_key = os.getenv("SECRET_KEY")
database_url = os.getenv("DATABASE_URL")

print(f"Secret Key: {secret_key}")
print(f"Database URL: {database_url}")
