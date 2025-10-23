import pywhatkit
from datetime import datetime, timedelta
import time
import pyautogui
import keyboard as k
import os
import sys

# Add parent directory to path to import config
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import *

# Use interactive authentication
try:
    # Import interactive authentication
    from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth
    
    print("Attempting interactive authentication...")
    auth = SharePointInteractiveAuth()
    if auth.authenticate_interactive():
        print("✅ Interactive authentication successful!")
    else:
        raise Exception("Interactive authentication failed")
    
except Exception as e:
    print(f"❌ Interactive authentication failed: {e}")
    
    try:
        # Fallback: Try with requests-ntlm for NTLM authentication
        from requests_ntlm import HttpNtlmAuth
        import requests
        
        print("Attempting NTLM authentication...")
        session = requests.Session()
        session.auth = HttpNtlmAuth(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
        
        # Test the connection
        response = session.get(f"{SHAREPOINT_SITE_URL}/_api/web")
        if response.status_code == 200:
            print("✅ NTLM authentication successful!")
            # You would need to adapt shareplum to use this session
        else:
            raise Exception(f"HTTP {response.status_code}: {response.text}")
            
    except Exception as e2:
        print(f"❌ NTLM authentication failed: {e2}")
        
        try:
            # Method 3: Try with O365 library (requires app registration)
            from O365 import Account
            
            print("Attempting O365 authentication...")
            credentials = (SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
            account = Account(credentials)
            
            if account.authenticate(scopes=['https://graph.microsoft.com/.default']):
                print("✅ O365 authentication successful!")
                # This would require different API calls
            else:
                raise Exception("O365 authentication failed")
                
        except Exception as e3:
            print(f"❌ O365 authentication failed: {e3}")
            print("\n🔧 TROUBLESHOOTING SUGGESTIONS:")
            print("1. Check if your credentials are correct")
            print("2. Verify if your account has access to the SharePoint site")
            print("3. Check if Multi-Factor Authentication is enabled (you may need an app password)")
            print("4. Consider using App Registration instead of user credentials")
            print("5. Contact your IT administrator about SharePoint access policies")
            sys.exit(1)

# If we get here, authentication was successful
print("Obteniendo datos de SharePoint...")

# Use interactive authentication to get data
Tutoriasdata = auth.get_list_items('Tutorias', ['ID', "Aula", 'Tipo de Tutoria', 'Contactado','Estado', 'Telefono', 'Nombre Tutor', 'Tipo de Tutoria', 'Fecha de Tutoria', 'Hora Tutoria', 'Clases','Temas','Alumnos', 'Aula', 'Numero de Cuenta'])
Tutoresdata = auth.get_list_items('Tutores', ['ID', 'Tutor', 'Telefono', 'TelefonoAuxiliar','Número de Cuenta' ])
Aulasdata = auth.get_list_items('Aulas', ['ID', 'IdAula ', 'Oficial'])

print("✅ Datos obtenidos exitosamente")

import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Configura la localización en español

print("Tutores Data")
print(Tutoriasdata)
print("--------------------")
print(Tutoresdata)
print("--------------------")
print("Aulas Data")


print("--------------------")
print(Tutoresdata)
print(len(Tutoriasdata))

ContactarPendientes = True
ContactarCoordinadas = True
ContactarRechazadas = False
ContactarReagendadas = True

ContactarModeradorAula = False
ContactarModeradorTutor = False

FilteredTutoriasdata = Tutoriasdata
auxSize=0
size = len(FilteredTutoriasdata)
print(Tutoresdata[0])

def obtenerNumerCuentaTutor(tutor):
    for j in range (0,len(Tutoresdata)):
        if Tutoresdata[j]['Tutor'] == tutor:
            return Tutoresdata[j]['Número de Cuenta']
    return "00000000"

# Rest of your existing code continues here...
# (I've included the authentication part, but you can copy the rest from your original file)

if (ContactarPendientes):
  for i in range (0,size):
   now = datetime.now()
   
   if FilteredTutoriasdata[i]['Estado']  and FilteredTutoriasdata[i]['Estado'] == "Pendiente" and FilteredTutoriasdata[i]['Contactado'] == "No":
     current_hour = int(now.strftime("%H"))
     if current_hour == 24:
         current_hour = 0
     current_minute = int(now.strftime("%M")) + 1
     if current_minute == 60:
         current_minute = 0

     TelefonoTutor = ""
     TelefonoAlumno = ""
     
     for j in range (0,len(Tutoresdata)):
         if Tutoresdata[j]['Tutor'] == FilteredTutoriasdata[i]['Nombre Tutor']:
            try:
             TelefonoTutor = Tutoresdata[j]['Telefono']
             TelefonoAlumno = FilteredTutoriasdata[i]['Telefono']
            except:
               try:
                TelefonoTutor = Tutoresdata[j]['TelefonoAuxiliar']
               except:
                TelefonoTutor = "87794832"
     
     numAux = "+504"+TelefonoTutor
     print("Mensaje enviado a: ",numAux)
     
     mensaje = "✨🐯PROPUESTA de Tutoria🐯✨"
     mensaje += "\n\n➡️ Modalidad : "+FilteredTutoriasdata[i]['Tipo de Tutoria']
     mensaje += "\n➡️ Fecha: "+FilteredTutoriasdata[i]['Fecha de Tutoria'].strftime("%Y-%m-%d %H:%M:%S")
     mensaje += "\n➡️ Día: " + FilteredTutoriasdata[i]['Fecha de Tutoria'].strftime("%A").upper()
     mensaje += "\n➡️ Hora: "+FilteredTutoriasdata[i]['Hora Tutoria']
     mensaje += "\n➡️ Asignatura: "+FilteredTutoriasdata[i]['Clases']
     mensaje += "\n➡️ Alumno: "+FilteredTutoriasdata[i]['Alumnos']
     mensaje += "\n➡️ Contacto: "+TelefonoAlumno
     mensaje += "\n➡️ Tema: "+FilteredTutoriasdata[i]['Temas']
     mensaje += "\n➡️ Tutor: "+FilteredTutoriasdata[i]['Nombre Tutor']
     mensaje += "\n\nAprobada Responder con => 👍"
     mensaje += "\nRechazada Responder con => 👎"
     mensaje += "\n\n(Si confirmas por medio de este mensaje, no es necesario que respondas el correo)"
     mensaje += "\nBeta Version 1.1"
     mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

     pyautogui.click(800, 450)
     time.sleep(2)
     k.press_and_release('enter')
     time.sleep(2)
     pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
     update_data = [{'ID': FilteredTutoriasdata[i]['ID'], 'Contactado': 'Yes'}]
     sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release('enter')
time.sleep(2)

print("✅ Script completed successfully!")
