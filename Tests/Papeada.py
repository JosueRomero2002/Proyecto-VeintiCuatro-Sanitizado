import requests
from datetime import date, timedelta

headers = {
  "apikey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZpZ2NvaHFpenpvYXB0YXFnc2tzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MzY3Mzk5MjYsImV4cCI6MjA1MjMxNTkyNn0.TckWcvyJX-vA_F7R3WTczYMc9CQGhN40jwRPbDEmvP0",
  "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZpZ2NvaHFpenpvYXB0YXFnc2tzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MzY3Mzk5MjYsImV4cCI6MjA1MjMxNTkyNn0.TckWcvyJX-vA_F7R3WTczYMc9CQGhN40jwRPbDEmvP0",
  "Content-Type": "application/json",
  "Prefer": "return=representation",
  "Content-Profile": "public"
}

usuario_id = 3220
usuario_id1 = 3281
usuario_id2 = 5395
usuario_id3 = 3283


for i in range(1):  # Lunes a Viernes
    fecha = (date(2025, 10, 21) + timedelta(days=i)).isoformat()
    
    for horario_id in range(1, 4):  # Horarios 1 al 8
        reservacion_body = {
            "motivo_uso": "Proyecto: ADAI GRUPAL",
            "cantidad_usuarios": 4,
            "fecha": fecha,
            "dias_repeticion": "",
            "laboratorio_id": "1"
        }

        # Crear una reservación por cada horario
        r = requests.post( 
            "https://figcohqizzoaptaqgsks.supabase.co/rest/v1/reservaciones?select=*",
            headers=headers,
            json=reservacion_body
        )

        if r.status_code != 201:
            print(f"❌ Error creando reservación para {fecha} (horario {horario_id}):", r.text)
            continue

        reservacion_id = r.json()[0]["id"]
        print(f"✅ Reservación creada: ID {reservacion_id} | Fecha: {fecha} | Horario: {horario_id}")

        # Relacionar con horario
        h_res = requests.post(
            "https://figcohqizzoaptaqgsks.supabase.co/rest/v1/reservaciones_horarios?columns=\"reservacion_id\",\"horario_id\"",
            headers=headers,
            json={"reservacion_id": reservacion_id, "horario_id": horario_id}
        )

        # Relacionar con usuario
        u_res = requests.post(
            "https://figcohqizzoaptaqgsks.supabase.co/rest/v1/reservaciones_usuarios?columns=\"reservacion_id\",\"usuario_id\"",
            headers=headers,
            json=[{"reservacion_id": reservacion_id, "usuario_id": usuario_id}, 
                  {"reservacion_id": reservacion_id, "usuario_id": usuario_id1},
                  {"reservacion_id": reservacion_id, "usuario_id": usuario_id2},
                  {"reservacion_id": reservacion_id, "usuario_id": usuario_id3},
            ]
        )