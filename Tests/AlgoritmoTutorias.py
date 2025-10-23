import pywhatkit
from datetime import datetime, timedelta
import time
import pyautogui
import keyboard as k
import random
import difflib
import os
import sys

# Add parent directory to path to import SharePointInteractiveAuth
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth

# Autenticación interactiva con SharePoint
print("Iniciando autenticación interactiva con SharePoint...")
auth = SharePointInteractiveAuth()
if not auth.authenticate_interactive():
    raise Exception("No se pudo autenticar con SharePoint")

print("Autenticación exitosa, obteniendo datos...")

# Obtener datos usando autenticación interactiva
Tutoriasdata = auth.get_list_items('Tutorias', ['ID', 'Aula', 'Tipo de Tutoria', 'Contactado','Estado', 'Telefono', 'Nombre Tutor', 'Fecha de Tutoria', 'Hora Tutoria', 'Clases','Temas','Alumnos', 'TutoresRechazaron'])
Tutoresdata = auth.get_list_items('Tutores', ['ID', 'Tutor', 'Telefono', 'TelefonoAuxiliar', 'Habilitado/Deshabilitado', 'Clases que Imparte', 'Horario Lunes', 'Horario Martes', 'Horario Miercoles', 'Horario Jueves', 'Horarios Viernes', 'Horario Sabado'])
random.shuffle(Tutoresdata) # Mezclar lista de tutores
Aulasdata = auth.get_list_items('Aulas', ['ID', 'IdAula ', 'Oficial'])

# print(Aulasdata)
# print(Tutoresdata)

AulasHabilitadas = [aula for aula in Aulasdata if 'Oficial' in aula ]

# print(AulasHabilitadas)

TutoriasCoordinadasCONAULA = [tut for tut in Tutoriasdata if tut['Estado'] in ['Coordinada'] and 'Aula' in tut and 'Nombre Tutor' in tut]
TutoriasCoordinadas = [tut for tut in Tutoriasdata if tut['Estado'] in ['Coordinada', 'Confirmada'] and 'Aula' in tut and 'Nombre Tutor' in tut]
TutoriasPendientesRechazadas = [tut for tut in Tutoriasdata if tut['Estado'] in ['Rechazada', "Pendiente"]]

def tutor_disponible(tutor, fecha_tutoria, hora_tutoria):

 
    
    # Validar si es parte de Horario del tutor
    # for tut in Tutoresdata:
    #     if tut['Tutor'] == tutor and hora_tutoria in tut[horario]:
    #         return True
    # return False


    # Validar si el tutor está disponible

    for tut in TutoriasCoordinadas:
        if tut['Nombre Tutor'] == tutor and tut['Fecha de Tutoria'] == fecha_tutoria and HoraChoca(tut['Hora Tutoria'], hora_tutoria):
            return False
    for tut in TutoriasPendientesRechazadas:
        if tut['Nombre Tutor'] == tutor and tut['Fecha de Tutoria'] == fecha_tutoria and HoraChoca(tut['Hora Tutoria'], hora_tutoria):
            return False
    return True

def aula_disponible(aula, fecha_tutoria, hora_tutoria):
    # Validar si el aula está disponible
    for tut in TutoriasCoordinadasCONAULA:
        if tut['Aula'] == aula and tut['Fecha de Tutoria'] == fecha_tutoria and HoraChoca(tut['Hora Tutoria'],hora_tutoria):  #HoraChoca(tut['Hora Tutoria'], hora_tutoria):
            return False
    return True

def asignar_tutor(tutoria):
    fecha_tutoria = tutoria['Fecha de Tutoria']
    hora_tutoria = tutoria['Hora Tutoria']
    clase_tutoria = tutoria['Clases']


    now = datetime.now()
    now = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if not(fecha_tutoria >= (now)): 
            return False


   # Seleccionar Horario segun Fecha
    if fecha_tutoria.weekday() == 0:
        horario = "Horario Lunes"
    elif fecha_tutoria.weekday() == 1:
        horario = "Horario Martes"
    elif fecha_tutoria.weekday() == 2:
        horario = "Horario Miercoles"
    elif fecha_tutoria.weekday() == 3:
        horario = "Horario Jueves"
    elif fecha_tutoria.weekday() == 4:
        horario = "Horarios Viernes"
    elif fecha_tutoria.weekday() == 5:
        horario = "Horario Sabado"
    else:
        return False


    rechazaron_tutoria = []
    if 'TutoresRechazaron' in tutoria and tutoria['TutoresRechazaron']:
        rechazaron_tutoria = tutoria['TutoresRechazaron']


 

    tutor_asignado = None


    # if (tutoria['Nombre Tutor']  and tutoria['Nombre Tutor'] != None and not (tutoria['Nombre Tutor'] in rechazaron_tutoria)) and tutoria['Nombre Tutor'] != "" and tutor_disponible(tutoria['Nombre Tutor'], fecha_tutoria, hora_tutoria):
        # tutor_asignado = tutoria['Nombre Tutor']
    # else:

    # Validar si la fecha de la tutoría es mayor a un día desde hoy
    


    for tutor in Tutoresdata:                                                                                       #Algebra   in ['Algebra Lineal', 'Calculo', 'Fisica'] Error lo toma como true
            if (not (tutor['Tutor'] in rechazaron_tutoria)) and tutor["Habilitado/Deshabilitado"] == 'Yes' and (clase_tutoria in tutor["Clases que Imparte"]) and hora_tutoria in tutor[horario]  and tutor_disponible(tutor['Tutor'], fecha_tutoria, hora_tutoria):
                tutor_asignado = tutor['Tutor']
                break
    # for tutor in Tutoresdata:
    #             if (not (tutor['Tutor'] in rechazaron_tutoria)) and tutor["Habilitado/Deshabilitado"] == 'Yes' and \
    #             (clase_tutoria in tutor["Clases que Imparte"].split(',')) and hora_tutoria in tutor[horario] and \
    #             tutor_disponible(tutor['Tutor'], fecha_tutoria, hora_tutoria):
    #                 tutor_asignado = tutor['Tutor']
    #                 break

    


    if tutor_asignado:
        tutoria['Nombre Tutor'] = tutor_asignado
        return True  
    return False  

def asignar_aula(tutoria):
    fecha_tutoria = tutoria['Fecha de Tutoria']
    hora_tutoria = tutoria['Hora Tutoria']
    

    aula_asignada = None
    for aula in AulasHabilitadas:
        if aula['Oficial'] == 'Yes' and aula_disponible(aula['IdAula '], fecha_tutoria, hora_tutoria):
            aula_asignada = aula['IdAula ']
            break
    

    if aula_asignada:
        tutoria['Aula'] = aula_asignada
        return True 
    return False 

def comparar_textos(texto1, texto2, umbral=0.6):
    # Calcular la similitud
    similaridad = difflib.SequenceMatcher(None, texto1, texto2).ratio()
    return similaridad >= umbral

def posibleUnion(tutoria):
    for tut in TutoriasCoordinadas:
        if comparar_textos(tut['Temas'], "PAE"):
            return False
        elif comparar_textos(tut['Clases'], tutoria['Clases']) and comparar_textos(tut['Temas'], tutoria['Temas']) and tut['Fecha de Tutoria'] == tutoria['Fecha de Tutoria'] and tut['Hora Tutoria'] == tutoria['Hora Tutoria']:
            tutoria['Aula'] = tut['Aula']
            tutoria['Nombre Tutor'] = tut['Nombre Tutor']
            tutoria['Tutor Asignado'] = tut['Nombre Tutor']
            tutoria['Tutor'] = tut['Nombre Tutor']
            return True
    return False

def HoraChoca(hora1, hora2):
    # Formato de hora: "HH:MM xm - HH:MM xm"
    hora1 = hora1.split(" - ")
    hora2 = hora2.split(" - ")

    # Convertir a formato de 24 horas
    hora1[0] = time.strptime(hora1[0], "%I:%M %p") # ejemplo 9:55 am
    hora1[1] = time.strptime(hora1[1], "%I:%M %p") # ejemplo 11:15 am
    hora2[0] = time.strptime(hora2[0], "%I:%M %p") # ejemplo 9:55 am
    hora2[1] = time.strptime(hora2[1], "%I:%M %p") # ejemplo 11:15 am

    # Verificar si las horas chocan

    # 9:55 am <= 11:15 am <= 11:15 am
    # Hora1 empieza antes de Hora2 y termina después de Hora2
    # Hora2 empieza antes de Hora1 y termina después de Hora1
         
    if (hora1[0] <= hora2[0] and hora2[0] <  hora1[1]) or (hora1[0] < hora2[1] and hora2[1] <= hora1[1]) or (hora2[0] <= hora1[0] and hora1[0] < hora2[1]) or (hora2[0] < hora1[1] and hora1[1] <= hora2[1]):
        return True #Si Choca
    


    return False


def SolicitarTutoria(Solicitud):
    # Crear una nueva tutoría
    sp_list_Tutorias.UpdateListItems(data=Solicitud, kind='New')
    print("Tutoría solicitada")


# def ObtenerHorariosDisponiblesParaClase(clase):
#     #obtener tutores que dan esa clase
#     tutores = [tutor for tutor in Tutoresdata if clase in tutor['Clases que Imparte']]
#     horarios = []

#     #filtrar tutores disponibles segun tutorias pendientes confirmadas rechazadas y coordinadas
#     for tutor in tutores:
#         for dia in ['Horario Lunes', 'Horario Martes', 'Horario Miercoles', 'Horario Jueves', 'Horarios Viernes', 'Horario Sabado']:
#             if tutor[dia] != "":
#                 for tut in TutoriasCoordinadas:
#                     if tutor['Tutor'] == tut['Nombre Tutor'] and tutor[dia] == tut['Hora Tutoria']:
#                         break
#                     else:

#                 horarios.append(tutor[dia])

#     return horarios




# Arreglo de Solcitides de Tutorias Ejemplo


# SolicitudesTutoriasTemporales = [
#     {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '98003715',
#   'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
#   'Fecha de Tutoria': '2025-01-21',
#   'Hora Tutoria': '9:55 am - 11:15 am',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ALFREDO ROBLEDA CASCO'},
# ]


SolicitudesTutoriasTemporales =[

 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-01-27'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-01-29'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-03'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-05'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-10'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-12'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-17'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-19'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-24'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-02-26'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-03'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-05'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-10'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-12'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-17'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-19'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-24'},
 {'Tipo de Tutoria': 'Presencial',
  'Contactado': 'No',
  'Estado': 'Confirmada',
  'Telefono': '98003715',
  'Nombre Tutor': 'IAN ROMAN BELTRAND PADILLA',
  'Hora Tutoria': '9:55 am - 11:15 am',
  'Clases': 'Algebra',
  'Temas': 'PAE',
  'Alumnos': 'ALFREDO ROBLEDA CASCO',
  'Fecha de Tutoria': '2025-03-26'},
]
# 
#  = [
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-01-27',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-01-28',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-03',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-04',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-10',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-11',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-17',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-18',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-24',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-02-25',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-03',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-04',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-10',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-11',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-17',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-18',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-24',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  {'Tipo de Tutoria': 'Presencial',
#   'Contactado': 'No',
#   'Estado': 'Confirmada',
#   'Telefono': '31781621',
#   'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Fecha de Tutoria': '2025-03-25',
#   'Hora Tutoria': '4:00 pm - 5:20 pm',
#   'Clases': 'Algebra',
#   'Temas': 'PAE',
#   'Alumnos': 'ROQUE TURCIOS SEVILLA'},
#  ]

# SolicitudesLICINTHIA = [
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-01-21', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-01-23', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-01-28', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-01-30', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-04', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-06', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-11', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-13', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-18', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-20', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-25', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-02-27', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#     {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-03-04', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'},
#  {'Tipo de Tutoria': 'Presencial', 'Contactado': 'No', 'Estado': 'Confirmada', 'Telefono': '98070043', 'Nombre Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA', 'Fecha de Tutoria': '2025-03-06', 'Hora Tutoria': '1:20 pm - 2:40 pm', 'Clases': 'Geometria y Trigonometria', 'Temas': 'PAE', 'Alumnos': 'KEILY YAMILETH FERNANDEZ PALOMO'}
# ]



#  [
   
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-01-31',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Por qué cuidar nuestra salud mental? Secciones 19, 17',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-01-31',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Por qué cuidar nuestra salud mental? Secciones 14, 13',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-01-31',
#         'Hora Tutoria': '9:55 am - 11:15 am',
#         'Clases': 'PAE',
#         'Temas': '¿Por qué cuidar nuestra salud mental? Sección 242',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#      {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-07',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': 'Mi Ser: Mi Autoestima Secciones 19, 17',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-07',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': 'Mi Ser: Mi Autoestima Secciones 14, 13',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-07',
#         'Hora Tutoria': '9:55 am - 11:15 am',
#         'Clases': 'PAE',
#         'Temas': 'Mi Ser: Mi Autoestima Sección 242',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#       {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-14',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Son mis pensamientos negativos? Secciones 19, 17',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-14',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Son mis pensamientos negativos? Secciones 14, 13',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-14',
#         'Hora Tutoria': '9:55 am - 11:15 am',
#         'Clases': 'PAE',
#         'Temas': '¿Son mis pensamientos negativos? Sección 242',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#      {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-28',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Para afrontar: Conozco de Resiliencia? Secciones 19, 17',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-28',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': '¿Para afrontar: Conozco de Resiliencia? Secciones 14, 13',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-02-28',
#         'Hora Tutoria': '9:55 am - 11:15 am',
#         'Clases': 'PAE',
#         'Temas': '¿Para afrontar: Conozco de Resiliencia? Sección 242',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#      {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-03-07',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': 'Proyectando mi futuro Secciones 19, 17',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-03-07',
#         'Hora Tutoria': '8:10 am - 9:30 am',
#         'Clases': 'PAE',
#         'Temas': 'Proyectando mi futuro Secciones 14, 13',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     },
#     {
#         'Tipo de Tutoria': 'Presencial',
#         'Contactado': 'No',
#         'Estado': 'Confirmada',
#         'Telefono': '87657664',
#         'Nombre Tutor': 'CARDENAS DELCID CYNTHIA STEPHANIE',
#         'Fecha de Tutoria': '2025-03-07',
#         'Hora Tutoria': '9:55 am - 11:15 am',
#         'Clases': 'PAE',
#         'Temas': 'Proyectando mi futuro Sección 242',
#         'Alumnos': 'JOSUE ALDAIR ROMERO PINEDA'
#     }

# ]

#  LUNES =================================================
SolicitudesLUNES = [ 
{
    'Tipo de Tutoria': 'Presencial',
    'Contactado': 'No',
    'Estado': 'Confirmada',
    'Telefono': '94454782',
    'Nombre Tutor': 'ERICK EDUARDO ARITA HENRIQUEZ',
    'Fecha de Tutoria': (datetime.now() + timedelta(days=(7 - datetime.now().weekday()))).date(),
    'Hora Tutoria': '9:55 am - 11:15 am',
    'Clases': 'Algebra Lineal',
    'Temas': 'PAE',
    'Alumnos': 'JEREMY JOSE RIVERA LARA'
},
  {
      
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94454782',
        'Nombre Tutor': 'JOSUE ARTURO DEL VALLE ZEPEDA',
        'Fecha de Tutoria':(datetime.now() + timedelta(days=(7 - datetime.now().weekday()))).date(),
        'Hora Tutoria': '11:15 am - 12:35 pm',
        'Clases': 'Calculo 1',
        'Temas': 'PAE',
       'Alumnos': 'JEREMY JOSE RIVERA LARA'
    },
      { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '98983878',
        'Nombre Tutor': 'CESIA ELISABETH ALFARO HERNANDEZ',
        'Fecha de Tutoria': (datetime.now() + timedelta(days=(7 - datetime.now().weekday()))).date(),
        'Hora Tutoria': '8:10 am - 9:30 am',
        'Clases': 'Geometria',
        'Temas': 'PAE',
       'Alumnos': 'RICARDO JARED CALIX HERNANDEZ'
    },
          { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94907224',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': (datetime.now() + timedelta(days=(7 - datetime.now().weekday()))).date(),
        'Hora Tutoria': '1:20 pm - 2:40 pm',
        'Clases': 'Circuitos DC',
        'Temas': 'PAE',
       'Alumnos': 'LUIS FERNANDO VELASQUEZ PONCE'
    },
       {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '93850183',
        'Nombre Tutor': 'EMMANUEL JOSE DISCUA RODRIGUEZ',
        'Fecha de Tutoria': (datetime.now() + timedelta(days=(7 - datetime.now().weekday()))).date(),
        'Hora Tutoria': '11:15 am - 12:15 pm',
        'Clases': 'Comput. Aplic. Al Diseño',
        'Temas': 'PAE',
       'Alumnos': 'RICHARD DANIEL ESTRADA RODRIGUEZ'
    },
    

]
# MARTES =================================================

SolicitudesMARTES = [ 
 
       {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '93850183',
        'Nombre Tutor': 'EMMANUEL JOSE DISCUA RODRIGUEZ',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '11:15 am - 12:15 pm',
        'Clases': 'Comput. Aplic. Al Diseño',
        'Temas': 'PAE',
       'Alumnos': 'RICHARD DANIEL ESTRADA RODRIGUEZ'
    },
       { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94907224',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '1:20 pm - 2:40 pm',
        'Clases': 'Circuitos DC',
        'Temas': 'PAE',
       'Alumnos': 'LUIS FERNANDO VELASQUEZ PONCE'
    },
            { 
                'Tipo de Tutoria': 'Presencial',
                'Contactado': 'Yes',
                'Estado': 'Confirmada',
                'Telefono': '98983878',
                'Nombre Tutor': 'CESIA ELISABETH ALFARO HERNANDEZ',
                'Fecha de Tutoria': datetime.now(),
                'Hora Tutoria': '8:10 am - 9:30 am',
                'Clases': 'Geometria',
                'Temas': 'PAE',
            'Alumnos': 'RICARDO JARED CALIX HERNANDEZ'
            },
            {
            'Tipo de Tutoria': 'Presencial',
            'Contactado': 'No',
            'Estado': 'Confirmada',
            'Telefono': '94454782',
            'Nombre Tutor': 'ERICK EDUARDO ARITA HENRIQUEZ',
            'Fecha de Tutoria': datetime.now(),
            'Hora Tutoria': '9:55 am - 11:15 am',
            'Clases': 'Algebra Lineal',
            'Temas': 'PAE',
            'Alumnos': 'JEREMY JOSE RIVERA LARA'
        },

       

      {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '89732130',
        'Nombre Tutor': 'VALERIA ESTEFANIA SOLORZANO SUAZO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '3:00 pm - 4:00 pm',
        'Clases': 'Intro al Algebra',
        'Temas': 'PAE',
      'Alumnos': 'MANUELA TOBON VÁSQUEZ'
    },
 {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '89732130',
        'Nombre Tutor': 'AXEL JAFET MARTEL RUBIO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '4:00 pm - 5:00 pm',
        'Clases': 'Comput. Aplic. Al Diseño',
        'Temas': 'PAE',
       'Alumnos': 'BYRON ANDRES COLE FAJARDO'
    },
   { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '31868331',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '4:10 pm - 5:10 pm',
        'Clases': 'Circuitos AC',
        'Temas': 'PAE',
       'Alumnos': 'IRVING ANDRES GAMONEDA CERRATO'
    },

]


# MIERCOLES =================================================



SolicitudesMIERCOLES = [ 
        
    
        
         {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '89732130',
        'Nombre Tutor': 'VALERIA ESTEFANIA SOLORZANO SUAZO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '3:00 pm - 4:00 pm',
        'Clases': 'Intro al Algebra',
        'Temas': 'PAE',
       'Alumnos': 'MANUELA TOBON VÁSQUEZ'
    },
       {
      
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '98437598',
        'Nombre Tutor': 'MEYBELIN SARAI BONILLA RIVERA',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '4:00 pm - 5:00 pm',
        'Clases': 'Algebra Lineal',
        'Temas': 'PAE',
       'Alumnos': 'CARLOS ENRIQUE LOPEZ ARITA'
    },
      {
      
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94454782',
        'Nombre Tutor': 'JOSUE ARTURO DEL VALLE ZEPEDA',
        'Fecha de Tutoria':datetime.now(),
        'Hora Tutoria': '11:15 am - 12:35 pm',
        'Clases': 'Calculo 1',
        'Temas': 'PAE',
       'Alumnos': 'JEREMY JOSE RIVERA LARA'
    },
        { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '98983878',
        'Nombre Tutor': 'KENNETH DANIEL REYES REYES',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '12:30 pm - 1:20 pm',
        'Clases': 'Programacion Estructurada',
        'Temas': 'PAE',
       'Alumnos': 'RICARDO JARED CALIX HERNANDEZ'
    },
   { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '94907224',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '1:20 pm - 2:40 pm',
        'Clases': 'Programacion Estructurada',
        'Temas': 'PAE',
       'Alumnos': 'LUIS FERNANDO VELASQUEZ PONCE'
    },
         {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '93850183',
        'Nombre Tutor': 'EMMANUEL JOSE DISCUA RODRIGUEZ',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '11:15 am - 12:15 pm',
        'Clases': 'Comput. Aplic. Al Diseño',
        'Temas': 'PAE',
       'Alumnos': 'RICHARD DANIEL ESTRADA RODRIGUEZ'
    },
            { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94974643',
        'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
        'Fecha de Tutoria':  datetime.now(),
        'Hora Tutoria': '2:40 pm - 4:00 pm',
        'Clases': 'Algebra',
        'Temas': 'PAE',
       'Alumnos': 'JEREMY ETHAN BODDEN MORALES'
    },
    
        {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '88403709',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '12:30 pm - 1:30 pm',
        'Clases': 'Electronica 2',
        'Temas': 'PAE',
       'Alumnos': 'MARIA JOSE GALVEZ GIRON'
    },
            
]

# JUEVES =================================================

SolicitudesJUEVES = [    

   { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '31868331',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '4:10 pm - 5:10 pm',
        'Clases': 'Circuitos AC',
        'Temas': 'PAE',
       'Alumnos': 'IRVING ANDRES GAMONEDA CERRATO'
    },

 {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '99871952',
        'Nombre Tutor':'EMMANUEL JOSE DISCUA RODRIGUEZ',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '5:00 pm - 6:00 pm',
        'Clases': 'Comput. Aplic. Al Diseño',
        'Temas': 'PAE',
       'Alumnos': 'BYRON ANDRES COLE FAJARDO'
 },


      {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '89732130',
        'Nombre Tutor': 'VALERIA ESTEFANIA SOLORZANO SUAZO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '3:00 pm - 4:00 pm',
        'Clases': 'Intro al Algebra',
        'Temas': 'PAE',
       'Alumnos': 'MANUELA TOBON VÁSQUEZ'
    },


{
    'Tipo de Tutoria': 'Presencial',
    'Contactado': 'No',
    'Estado': 'Confirmada',
    'Telefono': '94454782',
    'Nombre Tutor': 'ERICK EDUARDO ARITA HENRIQUEZ',
    'Fecha de Tutoria': datetime.now(),
    'Hora Tutoria': '9:55 am - 11:15 am',
    'Clases': 'Algebra Lineal',
    'Temas': 'PAE',
    'Alumnos': 'JEREMY JOSE RIVERA LARA'
},
        { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '98983878',
        'Nombre Tutor': 'KENNETH DANIEL REYES REYES',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '12:30 pm - 1:20 pm',
        'Clases': 'Programacion Estructurada',
        'Temas': 'PAE',
       'Alumnos': 'RICARDO JARED CALIX HERNANDEZ'
    },
   { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '94907224',
        'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '1:20 pm - 2:40 pm',
        'Clases': 'Circuitos DC',
        'Temas': 'PAE',
       'Alumnos': 'LUIS FERNANDO VELASQUEZ PONCE'
    },
          { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '94974643',
        'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
        'Fecha de Tutoria':  datetime.now(),
        'Hora Tutoria': '2:40 pm - 4:00 pm',
        'Clases': 'Algebra',
        'Temas': 'PAE',
       'Alumnos': 'JEREMY ETHAN BODDEN MORALES'
    },
       


]


# VIERNES =================================================

SolicitudesVIERNES = [



       {
      
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '98079632',
        'Nombre Tutor': 'JUAN SEBASTIAN MELENDEZ CRUZ',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '5:20 pm - 6:40 pm',
        'Clases': 'Circuitos DC',
        'Temas': 'PAE',
       'Alumnos': 'ALEX ROLANDO HERNANDEZ TRIMINIO'
    },
         {
      

        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '98437598',
        'Nombre Tutor': 'MEYBELIN SARAI BONILLA RIVERA',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '8:10 am - 9:30 pm',
        'Clases': 'Algebra Lineal',
        'Temas': 'PAE',
       'Alumnos': 'CARLOS ENRIQUE LOPEZ ARITA'
    },

         { 
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'No',
        'Estado': 'Confirmada',
        'Telefono': '00000000',
        'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '8:30 am - 9:30 am',
        'Clases': 'Intro al Algebra',
        'Temas': 'PAE',
       'Alumnos': 'RAUL FERNANDO PINEDA ACEBEDO'
    },



]




# SABADO =================================================

SolicitudesSABADO = [

      {
        'Tipo de Tutoria': 'Presencial',
        'Contactado': 'Yes',
        'Estado': 'Confirmada',
        'Telefono': '98003715',
        'Nombre Tutor': 'VALERIA ESTEFANIA SOLORZANO SUAZO',
        'Fecha de Tutoria': datetime.now(),
        'Hora Tutoria': '9:30 am - 11:30 am',
        'Clases': 'Algebra',
        'Temas': 'PAE',
       'Alumnos': 'ALFREDO ROBLEDA CASCO'
    },

]

    #    {
      
    #     'Tipo de Tutoria': 'Presencial',
    #     'Contactado': 'No',
    #     'Estado': 'Confirmada',
    #     'Telefono': '98437598',
    #     'Nombre Tutor': 'MEYBELIN SARAI BONILLA RIVERA',
    #     'Fecha de Tutoria': datetime.now(),
    #     'Hora Tutoria': '8:10 am - 9:30 pm',
    #     'Clases': 'Algebra Lineal',
    #     'Temas': 'PAE',
    #    'Alumnos': 'CARLOS ENRIQUE LOPEZ ARITA'
    # },
    # {
      
    #     'Tipo de Tutoria': 'Presencial',
    #     'Contactado': 'No',
    #     'Estado': 'Confirmada',
    #     'Telefono': '94454782',
    #     'Nombre Tutor': 'JOSUE ARTURO DEL VALLE ZEPEDA',
    #     'Fecha de Tutoria': datetime.now(),
    #     'Hora Tutoria': '11:15 am - 12:35 pm',
    #     'Clases': 'Algebra Lineal',
    #     'Temas': 'PAE',
    #    'Alumnos': 'JEREMY JOSE RIVERA LARA'
    # },
    # { 
    #     'Tipo de Tutoria': 'Presencial',
    #     'Contactado': 'No',
    #     'Estado': 'Confirmada',
    #     'Telefono': '98983878',
    #     'Nombre Tutor': 'KENNETH DANIEL REYES REYES',
    #     'Fecha de Tutoria': datetime.now(),
    #     'Hora Tutoria': '12:30 pm - 1:20 pm',
    #     'Clases': 'Programacion Estructurada',
    #     'Temas': 'PAE',
    #    'Alumnos': 'RICARDO JARED CALIX HERNANDEZ'
    # },
    #   { 
    #     'Tipo de Tutoria': 'Presencial',
    #     'Contactado': 'No',
    #     'Estado': 'Confirmada',
    #     'Telefono': '94907224',
    #     'Nombre Tutor': 'FARES ISAAC ENAMORADO GALEANO',
    #     'Fecha de Tutoria': datetime.now(),
    #     'Hora Tutoria': '1:20 pm - 2:40 pm',
    #     'Clases': 'Programacion Estructurada',
    #     'Temas': 'PAE',
    #    'Alumnos': 'LUIS FERNANDO VELASQUEZ PONCE'
    # },
    #      { 
    #     'Tipo de Tutoria': 'Presencial',
    #     'Contactado': 'No',
    #     'Estado': 'Confirmada',
    #     'Telefono': '94974643',
    #     'Nombre Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
    #     'Fecha de Tutoria': datetime.now(),
    #     'Hora Tutoria': '2:40 pm - 4:00 pm',
    #     'Clases': 'Programacion Estructurada',
    #     'Temas': 'PAE',
    #    'Alumnos': 'JEREMY ETHAN BODDEN MORALES'
    # },
   


# Solicitar tutorías
# for solicitud in SolicitudesLUNES:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesMARTES:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesMIERCOLES:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesJUEVES:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesVIERNES:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesSABADO:
    # SolicitarTutoria([solicitud])

# for solicitud in SolicitudesLICINTHIA:
    # SolicitarTutoria([solicitud])



from datetime import datetime, timedelta

def generar_solicitudes(tutor, clase, alumno, telefono, hora, dias, fecha_inicio, fecha_fin, temas='PAE'):
    solicitudes = []
    fecha_actual = datetime.strptime(fecha_inicio, '%Y-%m-%d')
    fecha_fin = datetime.strptime(fecha_fin, '%Y-%m-%d')
    
    dias_numeros = {
        'lunes': 0, 'martes': 1, 'miercoles': 2, 'jueves': 3, 'viernes': 4, 'sabado': 5, 'domingo': 6
    }
    
    dias_seleccionados = {dias_numeros[dia.lower()] for dia in dias}
    
    while fecha_actual <= fecha_fin:
        if fecha_actual.weekday() in dias_seleccionados:
            solicitudes.append({
                'Tipo de Tutoria': 'Presencial',
                'Contactado': 'No',
                'Estado': 'Confirmada',
                'Telefono': telefono,
                'Nombre Tutor': tutor,
                'Fecha de Tutoria': fecha_actual.strftime('%Y-%m-%d'),
                'Hora Tutoria': hora,
                'Clases': clase,
                'Temas': temas,
                'Alumnos': alumno
            })
        fecha_actual += timedelta(days=1)
    
    return solicitudes


# tutor = 'DANIEL ALEXANDER HERNANDEZ ESCOTO'
# clase = 'Calculo 1'
# alumno = 'RICARDO JARED CALIX HERNANDEZ'
# telefono = '98070043'
# hora = '9:30 am - 10:50 am'
# dias = ['miercoles', 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# tutor = 'ANTHONY EMANUEL FUNEZ GUERRA'
# clase = 'Calculo 1'
# alumno = 'VALERIA MICHELL CANTILLANO CARRANZA'
# telefono = '95373977'
# hora = '2:40 pm - 4:00 pm'
# dias = ['martes', 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# tutor = 'YASIN SINOE HERNANDEZ DIAZ'
# clase = 'Dibujo'
# alumno = 'JAHIR SAMUEL FLORES SANTAMARIA'
# telefono = '97977557'
# hora = '1:20 pm - 2:40 pm'
# dias = ['martes', 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# tutor = 'KENNETH DANIEL REYES REYES'
# clase = 'Intro al Algebra'
# alumno = 'MARIA FERNANDA MARTINEZ MURILLO'
# telefono = '97116220'
# hora = '12:35 pm - 1:55 pm'
# dias = [ 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# tutor = 'KENNETH DANIEL REYES REYES'
# clase = 'Intro al Algebra'
# alumno = 'MARIA FERNANDA MARTINEZ MURILLO'
# telefono = '97116220'
# hora = '11:15 am - 12:35 pm'
# dias = [ 'viernes']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'


# tutor = 'JOSUE ARTURO DEL VALLE ZEPEDA'
# clase = 'Circuitos DC'
# alumno = 'ALEX ROLANDO HERNANDEZ TRIMINIO'
# telefono = '98273847'
# hora = '1:20 pm - 2:40 pm'
# dias = [ 'viernes']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# //////////////////////////////////////////////////////////
# tutor = 'JOSUE ARTURO DEL VALLE ZEPEDA'
# clase = 'Calculo 2'
# alumno = 'JEREMY JOSE RIVERA LARA'
# telefono = '32879125'
# hora = '11:15 am - 12:35 pm'
# dias = ['martes', 'miercoles']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# # eroras

# tutor = 'MEYBELIN SARAI BONILLA RIVERA'
# clase = 'Algebra Lineal'
# alumno = 'JEREMY JOSE RIVERA LARA'
# telefono = '32879125'
# hora = '3:00 pm - 4:00 pm'
# dias = ['lunes', 'martes']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'


# tutor = 'KENNETH DANIEL REYES REYES'
# clase = 'Programacion Estructurada',
# alumno = 'CARLOS ENRIQUE LOPEZ ARITA'
# telefono = '98437598'
# hora = '11:15 am - 12:35 pm'
# dias = [ 'lunes']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'




# tutor = 'YASIN SINOE HERNANDEZ DIAZ'
# clase = 'Dibujo'
# alumno = 'JEREMY ETHAN BODDEN MORALES'
# telefono = '94974643'
# hora = '2:40 pm - 4:00 pm'
# dias = ['lunes', 'martes', 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'


# tutor = 'YASIN SINOE HERNANDEZ DIAZ'
# clase = 'Bocetaje y visualización 1'
# alumno = 'RICHARD DANIEL ESTRADA RODRIGUEZ'
# telefono = '93850183'
# hora = '1:00 pm - 2:20 pm'
# dias = ['martes', 'jueves']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'

# tutor = 'ERICK EDUARDO ARITA HENRIQUEZ'
# clase = 'Fisica 3'
# alumno = 'CRISTIAN ALONSO CABRERA CRUZ'
# telefono = '95100560'
# hora = '9:55 am - 11:15 am'
# dias = ['miercoles']
# fecha_inicio = '2025-01-20'
# fecha_fin = '2025-03-09'


tutor ='JOSUE ALDAIR ROMERO PINEDA'
clase = 'Programacion Estructurada'
alumno =  'ANGIE MAYELI GUEVARA PORTILLO'
telefono = '95100560'
hora = '11:00 am - 12:00 pm'
dias = ['lunes', 'martes', 'miercoles', 'jueves','viernes' ]
fecha_inicio = '2025-01-20'
fecha_fin = '2025-03-09'


# SolicitudesLICINTHIA = generar_solicitudes(tutor, clase, alumno, telefono, hora, dias, fecha_inicio, fecha_fin)

# for solicitud in SolicitudesLICINTHIA:
#     SolicitarTutoria([solicitud])


# def ObtenerHorariosDisponiblesParaClaseSegunDia(clase, dia):
#     # Obtener tutores que dan esa clase
#     tutores = [tutor for tutor in Tutoresdata if clase in tutor['Clases que Imparte']]
#     horarios = []

#     # Filtrar tutores disponibles según tutorías pendientes, confirmadas, rechazadas y coordinadas
#     for tutor in tutores:
#         if tutor[dia] != "":
#             horarios.append(tutor[dia])

#     return horarios

# def imprimirHorariosPorDia(clase, dia):
#     horarios = ObtenerHorariosDisponiblesParaClaseSegunDia(clase, dia)
#     if horarios:
#         print(f"Horario {dia} ------")
#         print("\n".join(horarios))
#     else:
#         print(f"Horario {dia} ------ No hay horarios disponibles")

# clase_PPP = 'Programacion 2'

# # Llamada a la función para cada día
# imprimirHorariosPorDia(clase_PPP, 'Horario Lunes')
# imprimirHorariosPorDia(clase_PPP, 'Horario Martes')
# imprimirHorariosPorDia(clase_PPP, 'Horario Miercoles')
# imprimirHorariosPorDia(clase_PPP, 'Horario Jueves')
# imprimirHorariosPorDia(clase_PPP, 'Horarios Viernes')
# imprimirHorariosPorDia(clase_PPP, 'Horario Sabado')

for tutoria in Tutoriasdata:
    if tutoria['Estado'] == "Solicitada" or tutoria['Estado'] == "Rechazada":
        if posibleUnion(tutoria):
            update_data = [{'ID': tutoria['ID'], 'Estado': 'Coordinada', 'Nombre Tutor': tutoria['Nombre Tutor'], 'Tutor Asignado': tutoria['Nombre Tutor'], 'Tutor': tutoria['Nombre Tutor'], 'Aula': tutoria['Aula'] , 'Contactado': "No"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')
            TutoriasPendientesRechazadas.append(tutoria)

            print(f"Tutoria ID {tutoria['ID']} coordinada")
        elif asignar_tutor(tutoria):


          
            update_data = [{'ID': tutoria['ID'], 'Nombre Tutor': tutoria['Nombre Tutor'], 'Tutor Asignado': "", 'Tutor': tutoria['Nombre Tutor'], 'Estado': 'Pendiente', 'Contactado': "No"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')

            TutoriasPendientesRechazadas.append(tutoria)
            print(f"Tutoria ID {tutoria['ID']} asignada con {tutoria['Nombre Tutor']}")
        else:
            update_data = [{'ID': tutoria['ID'],  'Estado': 'Sin Tutor', 'Contactado': "No"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')
            print(f"(-)  No se encontró tutor")
    elif tutoria['Estado'] == "Confirmada": 
        if asignar_aula(tutoria):
            
            update_data = [{'ID': tutoria['ID'], 'Aula': tutoria['Aula'], 'Estado': 'Coordinada', 'Contactado': "No"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')
            print(f"Aula asignada para tutoría ID {tutoria['ID']}: {tutoria['Aula']}")
            # Hacer el append de la tutoría actualizada
            tutoria_actualizada = tutoria.copy()  # Hacer una copia para evitar modificar el original
            tutoria_actualizada['Estado'] = 'Coordinada'  # Actualizar el estado en la copia
            tutoria_actualizada['Aula'] = tutoria['Aula']  # Asegurarse de que el aula esté actualizada

            # Agregar la tutoría actualizada a la lista
            TutoriasCoordinadasCONAULA.append(tutoria_actualizada)




        else:
            update_data = [{'ID': tutoria['ID'],  'Estado': 'CASO', 'Contactado': "No"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')
            print(f"(-)  No se encontró aula")

print("Proceso finalizado")

# DEJE un error a proposito, no todos los nombres son unicos Tutor Asignado podria dar problemas en un futuro, lo mejor era usar el email
# pero no me pagan lo suficiente como para hacerlo bien asi que lo dejo asi


    # fecha_tutoria = tutoria['Fecha de Tutoria']
    # hora_tutoria = tutoria['Hora Tutoria']
    
    # # Buscar tutor disponible
    # tutor_asignado = None
    # for tutor in Tutoresdata:
    #     if tutor_disponible(tutor['Tutor'], fecha_tutoria, hora_tutoria):
    #         tutor_asignado = tutor['Tutor']
    









    #         break
    
    # # Buscar aula disponible (si es presencial)
    # aula_asignada = None
    # if tutoria['Tipo de Tutoria'] == "Presencial":
    #     for aula in Aulasdata:
    #         if aula_disponible(aula['IdAula '], fecha_tutoria, hora_tutoria):
    #             aula_asignada = aula['IdAula ']
    #             break
    
    # # Actualizar datos de tutoría si se encontró un tutor y aula
    # if tutor_asignado:
    #     tutoria['Nombre Tutor'] = tutor_asignado
    #     tutoria['Tutor'] = tutor_asignado
    #     if aula_asignada:
    #         tutoria['Aula'] = aula_asignada
    #     return True  # Se asignó exitosamente
    # return False  # No se pudo asignar