import re
import pywhatkit
from datetime import datetime, timedelta
import time
import pyautogui
import keyboard as k
from shareplum import Site
from shareplum import Office365
import random
import difflib
import pandas
import os
from dotenv import load_dotenv

load_dotenv()  

userSh = os.getenv('SHAREPOINT_USER')
passSh = os.getenv('SHAREPOINT_PASS')

# Autenticación en SharePoint
authcookie = Office365('https://unitechn.sharepoint.com', username=userSh, password=passSh).GetCookies()

site = Site('https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/', authcookie=authcookie)



sp_list_Tutorias = site.List('Tutorias')
sp_list_Tutores = site.List('Tutores')
sp_list_Aulas = site.List('Aulas')

Tutoriasdata = sp_list_Tutorias.GetListItems(fields=['ID', 'Aula', 'Tipo de Tutoria', 'Contactado','Estado', 'Telefono', 'Nombre Tutor', 'Fecha de Tutoria', 'Hora Tutoria', 'Clases','Temas','Alumnos', 'TutoresRechazaron'])
Tutoresdata = sp_list_Tutores.GetListItems(fields=['ID', 'Tutor', 'Telefono', 'TelefonoAuxiliar', 'Habilitado/Deshabilitado', 'Clases que Imparte', 'Horario Lunes', 'Horario Martes', 'Horario Miercoles', 'Horario Jueves', 'Horarios Viernes', 'Horario Sabado'])
random.shuffle(Tutoresdata) # Mezclar lista de tutores
Aulasdata = sp_list_Aulas.GetListItems(fields=['ID', 'IdAula ', 'Oficial'])



def actulizarHorariosTutores(DatosTutores):
    for tutor in DatosTutores:
        for t in Tutoresdata:
            if tutor['Tutor'] == t['Tutor']:
                sp_list_Tutores.UpdateListItems(data=[{
                    'ID': t['ID'],
                    'Horario Lunes': tutor['Horario Lunes'],
                    'Horario Martes': tutor['Horario Martes'],
                    'Horario Miercoles': tutor['Horario Miercoles'],
                    'Horario Jueves': tutor['Horario Jueves'],
                    'Horarios Viernes': tutor['Horarios Viernes'],
                    'Horario Sabado': tutor['Horario Sabado'],
                }])
                break



def obtenerClasesDesdeSharePoint():
    """
    Obtiene los IDs y nombres de clases desde SharePoint.

    :return: Diccionario donde las claves son los nombres de las clases y los valores sus IDs.
    """
    try:
        # Consulta-la lista de clases en SharePoint
        sp_list_Clases = site.List('Clases')  # Asegúrate de que este sea el nombre correcto
        clases_data = sp_list_Clases.GetListItems(fields=['ID', 'Nombre de Clase'])  # Ajusta 'NombreClase' según el campo real
        # Crear un diccionario {NombreClase: ID}
        clases_dict = {clase['Nombre de Clase']: clase['ID'] for clase in clases_data}
        print("Clases obtenidas desde SharePoint:", clases_dict)
        return clases_dict
    except Exception as e:
        print(f"Error al obtener las clases desde SharePoint: {e}")
        return {}


def generarStringClases(clases_tutor, clases_dict):
    """
    Genera el string en formato {id}:#{nombreClase}#; para SharePoint.

    :param clases_tutor: Lista de nombres de clases proporcionados por el tutor.
    :param clases_dict: Diccionario {NombreClase: ID} de clases válidas en SharePoint.
    :return: String en formato válido para SharePoint o una cadena vacía si no hay coincidencias.
    """
    partes = [
        f"{clases_dict[clase]};#{clase};#"
        for clase in clases_tutor if clase in clases_dict
    ]
    return ''.join(partes)  if partes else ''



def crearTutores(DatosTutores):
    """
    Crea tutores en la lista de SharePoint asignándoles clases con el formato {id}:#{nombreClase}#;.

    :param DatosTutores: Lista de diccionarios con datos de los tutores-agregar.
    """
    # Obtener la lista de clases válidas desde SharePoint
    clases_dict = obtenerClasesDesdeSharePoint()
    
    if not clases_dict:
        print("No se encontraron clases válidas, proceso cancelado.")
        return

    for tutor in DatosTutores:
        # Dividir las clases proporcionadas por el tutor
        clases_tutor = tutor.get('Clases', '').split(', ')  # Manejar clases separadas por comas
        # Generar el string en el formato necesario

        print(clases_tutor)
        clases_formato = generarStringClases(clases_tutor, clases_dict)
        
        if not clases_formato:
            print(f"Tutor {tutor['Tutor']} no tiene clases válidas, omitiendo.")
            continue
        
        # Crear o actualizar al tutor en SharePoint
        try:
            sp_list_Tutores.UpdateListItems(data=[{
                'Tutor': tutor['Tutor'],
                'Telefono': tutor['Telefono'],
                'Habilitado/Deshabilitado': tutor['Habilitado/Deshabilitado'],
                'Clases que Imparte': clases_formato,
            }], kind='New')
            print(f"Tutor {tutor['Tutor']} agregado correctamente con clases: {clases_formato}")
        except Exception as e:
            print(f"Error al agregar al tutor {tutor['Tutor']}: {e}")




def obtenerHorariosDesdeSharePoint():
    """
    Obtiene los IDs y horarios desde SharePoint.

    :return: Diccionario donde las claves son los horarios y los valores sus IDs.
    """
    try:
        # Consulta-la lista de horarios en SharePoint
        sp_list_Horarios = site.List('Horarios Contenido')  # Asegúrate de que este sea el nombre correcto
        horarios_data = sp_list_Horarios.GetListItems(fields=['ID', 'Horarios'])  # Ajusta 'Horario' según el campo real
        # Crear un diccionario {Horario: ID}
        horarios_dict = {horario['Horarios']: horario['ID'] for horario in horarios_data}
        print("Horarios obtenidos desde SharePoint:", horarios_dict)
        return horarios_dict
    except Exception as e:
        print(f"Error al obtener los horarios desde SharePoint: {e}")
        return {}


def generarStringHorarios(horarios_tutor, horarios_dict):
    """
    Genera el string en formato {id};#{horario}#; para SharePoint.

    :param horarios_tutor: Lista de horarios proporcionados por el tutor.
    :param horarios_dict: Diccionario {Horario: ID} de horarios válidos en SharePoint.
    :return: String en formato válido para SharePoint o una cadena vacía si no hay coincidencias.
    """
    partes = [
        f"{horarios_dict[horario]};#{horario};#"
        for horario in horarios_tutor if horario in horarios_dict
    ]
    return ''.join(partes) if partes else ''



from datetime import datetime

def procesarRangosHorarios(rangos_horarios):
    """
    Divide un string de rangos de horarios separados por comas en una lista de rangos individuales.

    :param rangos_horarios: String con rangos de horarios, separados por comas.
    :return: Lista de rangos de horarios individuales.
    """
    return [rango.strip() for rango in rangos_horarios.split(',') if rango.strip()]

def filtrarHorariosSinV(rangos_horarios, horarios_dict):
    """
    Filtra los horarios disponibles que caen dentro de uno o más rangos de tiempo,
    ignorando valores especiales y horarios que finalizan con '[V]'.

    :param rangos_horarios: Rango(s) en formato "12:00 pm - 4:00 pm, 6:00 pm - 8:00 pm".
    :param horarios_dict: Diccionario {Horario: ID} con horarios disponibles.
    :return: Lista de horarios en formato {id};#{horario}#;.
    """
    rangos = procesarRangosHorarios(rangos_horarios)
    horarios_en_rango = []

    for rango in rangos:
        try:
            inicio, fin = [datetime.strptime(h.strip(), "%I:%M %p") for h in rango.split('-')]
        except ValueError:
            print(f"Formato de rango inválido: {rango}")
            continue

        for horario, id_horario in horarios_dict.items():
            if horario.endswith("[V]"):
                continue  # Ignorar horarios que finalicen con '[V]'

            try:
                hora_inicio = datetime.strptime(horario.split(' - ')[0], "%I:%M %p")
                if inicio <= hora_inicio < fin:
                    horarios_en_rango.append(f"{id_horario};#{horario};#")
            except ValueError:
                print(f"Ignorando horario no válido: {horario}")

    return horarios_en_rango

def filtrarHorariosConV(rangos_horarios, horarios_dict):
    """
    Filtra los horarios disponibles que caen dentro de uno o más rangos de tiempo,
    aceptando solo los horarios que finalizan con '[V]' como válidos.

    :param rangos_horarios: Rango(s) en formato "12:00 pm - 4:00 pm, 6:00 pm - 8:00 pm".
    :param horarios_dict: Diccionario {Horario: ID} con horarios disponibles.
    :return: Lista de horarios en formato {id};#{horario}#;.
    """
    rangos = procesarRangosHorarios(rangos_horarios)
    horarios_en_rango = []

    for rango in rangos:
        try:
            inicio, fin = [datetime.strptime(h.strip(), "%I:%M %p") for h in rango.split('-')]
        except ValueError:
            print(f"Formato de rango inválido: {rango}")
            continue

        for horario, id_horario in horarios_dict.items():
            if not horario.endswith("[V]"):
                continue  # Ignorar horarios que no terminan con '[V]'

            try:
                hora_inicio = datetime.strptime(horario.split(' - ')[0], "%I:%M %p")
                if inicio <= hora_inicio < fin:
                    horarios_en_rango.append(f"{id_horario};#{horario};#")
            except ValueError:
                print(f"Ignorando horario no válido: {horario}")

    return horarios_en_rango




def crearTutoresConRangosHorarios(DatosTutores):
    """
    Crea tutores en la lista de SharePoint asignándoles clases y horarios filtrados por rangos.

    :param DatosTutores: Lista de diccionarios con datos de los tutores-agregar.
    """
    # Obtener listas de clases y horarios desde SharePoint
    clases_dict = obtenerClasesDesdeSharePoint()
    horarios_dict = obtenerHorariosDesdeSharePoint()
    
    if not clases_dict or not horarios_dict:
        print("No se encontraron clases o horarios válidos, proceso cancelado.")
        return

    for tutor in DatosTutores:
        # Manejar clases
        clases_tutor = tutor.get('Clases', '').split(', ')
        clases_formato = generarStringClases(clases_tutor, clases_dict)
        
        # Manejar horarios por rango
        rango_horario_lunes = tutor.get('Horario Lunes', '')
        rango_horario_martes = tutor.get('Horario Martes', '')
        rango_horario_miercoles = tutor.get('Horario Miercoles', '')
        rango_horario_jueves = tutor.get('Horario Jueves', '')
        rango_horario_viernes = tutor.get('Horarios Viernes', '')
        rango_horario_sabado = tutor.get('Horario Sabado', '')

        # Filtrar horarios virtuales por rango
        rango_horario_virtual_lunes = tutor.get('Horarios Virtuales Lunes', '')
        rango_horario_virtual_martes = tutor.get('Horarios Virtuales Martes', '')
        rango_horario_virtual_miercoles = tutor.get('Horarios Virtuales Miercoles', '')
        rango_horario_virtual_jueves = tutor.get('Horarios Virtuales Jueves', '')
        rango_horario_virtual_viernes = tutor.get('Horarios Virtuales Viernes', '')
        rango_horario_virtual_sabado = tutor.get('Horarios Virtuales Sabado', '')
        
        horarios_en_rango_lunes = filtrarHorariosSinV(rango_horario_lunes, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_lunes, horarios_dict)
        horarios_en_rango_martes = filtrarHorariosSinV(rango_horario_martes, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_martes, horarios_dict)
        horarios_en_rango_miercoles = filtrarHorariosSinV(rango_horario_miercoles, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_miercoles, horarios_dict)
        horarios_en_rango_jueves = filtrarHorariosSinV(rango_horario_jueves, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_jueves, horarios_dict)
        horarios_en_rango_viernes = filtrarHorariosSinV(rango_horario_viernes, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_viernes, horarios_dict)
        horarios_en_rango_sabado = filtrarHorariosSinV(rango_horario_sabado, horarios_dict) + filtrarHorariosConV(rango_horario_virtual_sabado, horarios_dict)
        # + filtrarHorariosConV(rango_horario, horarios_dict) 
        horarios_formato_lunes = ''.join(horarios_en_rango_lunes) if horarios_en_rango_lunes else ''
        horarios_formato_martes = ''.join(horarios_en_rango_martes) if horarios_en_rango_martes else ''
        horarios_formato_miercoles = ''.join(horarios_en_rango_miercoles) if horarios_en_rango_miercoles else ''
        horarios_formato_jueves = ''.join(horarios_en_rango_jueves) if horarios_en_rango_jueves else ''
        horarios_formato_viernes = ''.join(horarios_en_rango_viernes) if horarios_en_rango_viernes else ''
        horarios_formato_sabado = ''.join(horarios_en_rango_sabado) if horarios_en_rango_sabado else ''



        if not clases_formato:
            print(f"Tutor {tutor['Tutor']} no tiene clases u horarios válidos, omitiendo.")
            continue

        # print(horarios_formato)
        
        # Crear o actualizar al tutor en SharePoint
        try:
            sp_list_Tutores.UpdateListItems(data=[{
                'Tutor': tutor['Tutor'],
                'Telefono': tutor['Telefono'],
                'Habilitado/Deshabilitado': tutor['Habilitado/Deshabilitado'],
                'Clases que Imparte': clases_formato,
                'Horarios de Tutor': horarios_formato_lunes,
                'Horario Lunes': horarios_formato_lunes,
                'Horario Martes': horarios_formato_martes,
                'Horario Miercoles': horarios_formato_miercoles,
                'Horario Jueves': horarios_formato_jueves,
                'Horarios Viernes': horarios_formato_viernes,
                'Horario Sabado': horarios_formato_sabado,

            }], kind='New')
            print(f"Tutor {tutor['Tutor']} agregado correctamente con clases y horarios.")
        except Exception as e:
            print(f"Error al agregar al tutor {tutor['Tutor']}: {e}")

# ==================================================================================================

# # Datos Tutores
# DatosTutoresPayload = [{'Tutor': 'ERICK EDUARDO ARITA HENRIQUEZ',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '97002422',
#   'Clases': 'Algebra, Fisica 3, Circuitos DC, Circuitos AC',
#   'Horario Lunes': '9:55 am -  11:15 pm, 11:15 am - 12:35 pm, 2:40 pm - 4:00 pm',
#   'Horario Martes': '9:55 am -  11:15 pm, 11:15 am - 12:35 pm, 2:40 pm - 4:00 pm',
#   'Horario Miercoles': '9:55 am -  11:15 pm, 11:15 am - 12:35 pm, 2:40 pm - 4:00 pm',
#   'Horario Jueves': '9:55 am -  11:15 pm, 11:15 am - 12:35 pm, 2:40 pm - 4:00 pm',
#   'Horarios Viernes': '2:40 pm-5:20 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'RONALD JOSEP PERDOMO MENDOZA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '94435540',
#   'Clases': 'Algebra',
#   'Horario Lunes': '8:10 am - 12:35 pm, 1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Martes': '8:10 am - 12:35 pm, 1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Miercoles': '8:10 am - 12:35 pm, 1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Jueves': '1:20 pm - 2:40 pm, 04:00 pm - 06:40 pm',
#   'Horarios Viernes': '8:10 am - 12:35 pm, 1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'KENNETH DANIEL REYES REYES',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '98918825',
#   'Clases': 'Fisica 1',
#   'Horario Lunes': '11:15 am - 5:00 pm',
#   'Horario Martes': '11:15 am - 5:00 pm',
#   'Horario Miercoles': '11:15 am - 5:00 pm',
#   'Horario Jueves': '11:15 am - 5:00 pm',
#   'Horarios Viernes': '',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'FARES ISAAC ENAMORADO GALEANO',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '31443203',
#   'Clases': 'Circuitos DC, Circuitos AC, Electronica 1, Estatica, Dinamica, Calculo 2',
#   'Horario Lunes': '10:00 am - 1:00 pm, 4:00 pm - 5:00 pm',
#   'Horario Martes': '10:00 am - 1:00 pm, 4:00 pm - 5:00 pm',
#   'Horario Miercoles': '10:00 am - 1:00 pm, 4:00 pm - 5:00 pm',
#   'Horario Jueves': '10:00 am - 1:00 pm, 4:00 pm - 5:00 pm',
#   'Horarios Viernes': '9:00 am - 1:00 pm',
#   'Horario Sabado': '',
#    'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'JOSUE ARTURO DEL VALLE ZEPEDA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '95870132',
#   'Clases': 'Dibujo Técnico (Solidworks), Variable Compleja, Ecuaciones Diferenciales',
#   'Horario Lunes': '11:15 am - 12:35 pm, 6:40 pm - 8:00 pm',
#   'Horario Martes': '11:15 am - 12:35 pm, 6:40 pm - 8:00 pm',
#   'Horario Miercoles': '11:15 am - 12:35 pm, 5:20 pm - 8:00 pm',
#   'Horario Jueves': '11:15 am - 12:35 pm, 5:20 pm - 8:00 pm',
#   'Horarios Viernes': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'VALERIA ESTEFANIA SOLORZANO SUAZO',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '96724847',
#   'Clases': 'Algebra',
#   'Horario Lunes': '',
#   'Horario Martes': '',
#   'Horario Miercoles': '2:40 pm - 4:00 pm',
#   'Horario Jueves': '2:40 pm - 4:00 pm',
#   'Horarios Viernes': '2:40 pm - 4:00 pm',
#   'Horario Sabado': '',
#    'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'MEYBELIN SARAI BONILLA RIVERA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '93664394',
#   'Clases': 'Algebra Lineal, Calculo 1',
#   'Horario Lunes': '11:15 am - 12:15 pm, 3:00 pm - 4:00 pm, 5:30 pm - 6:30 pm',
#   'Horario Martes': '11:15 am - 12:15 pm, 3:00 pm - 4:00 pm, 5:30 pm - 6:30 pm',
#   'Horario Miercoles': '11:15 am - 12:15 pm, 3:00 pm - 4:00 pm, 5:30 pm - 6:30 pm',
#   'Horario Jueves': '11:15am - 12:15pm, 3:00 pm - 4:00 pm, 5:30 pm - 6:30 pm',
#   'Horarios Viernes': '11:15 am - 12:15 pm, 1:00 pm - 3:00 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'FERNANDO DAVID SOSA FLORES',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '96186573',
#   'Clases': '',
#   'Horario Lunes': '1:20 pm - 2:40 pm',
#   'Horario Martes': '1:20 pm - 2:40 pm',
#   'Horario Miercoles': '1:20 pm - 2:40 pm',
#   'Horario Jueves': '1:20 pm - 2:40 pm',
#   'Horarios Viernes': '9:55 am - 11:15 am, 1:20 pm - 2:40 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'ISAAC MAURICIO JUAREZ BACA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '97949526',
#   'Clases': 'Geometria y Trigonometria, Calculo 1, Calculo 2',
#   'Horario Lunes': '5:20 pm - 6:40 pm, 8:00 pm - 9:20 pm',
#   'Horario Martes': '5:20 pm - 6:40 pm, 8:00 pm - 9:20 pm',
#   'Horario Miercoles': '5:20 pm - 6:40 pm, 8:00 pm - 9:20 pm',
#   'Horario Jueves': '5:20 pm - 6:40 pm, 8:00 pm - 9:20 pm',
#   'Horarios Viernes': '5:20 pm - 6:40 pm, 8:00 pm - 9:20 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'ALEX ROBERTO HERNANDEZ GARCIA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '32102440',
#   'Clases': 'Intro al Algebra',
#   'Horario Lunes': '4:30 pm - 8:00 pm',
#   'Horario Martes': '2:30 pm- 8:00 pm',
#   'Horario Miercoles': '4:30 pm - 8:00 pm',
#   'Horario Jueves': '4:30 pm - 8:00 pm',
#   'Horarios Viernes': '4:30 pm - 8:00 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '4:30 pm - 8:00 pm',
#   'Horarios Virtuales Martes': '2:30 pm- 8:00 pm',
#   'Horarios Virtuales Miercoles': '4:30 pm - 8:00 pm',
#   'Horarios Virtuales Jueves': '4:30 pm - 8:00 pm',
#   'Horarios Virtuales Viernes': '4:30 pm - 8:00 pm',
#   'Horarios Virtuales Sabado': ''},
# #  {'Tutor': 'Daniel Hernandez',
# #   'Habilitado/Deshabilitado': 'Yes',
# #   'Email': '',
# #   'Telefono': '',
# #   'Clases': '',
# #   'Horario Lunes': '',
# #   'Horario Martes': '8:00 am-11:30 am',
# #   'Horario Miercoles': '8:00 am-11:30 am',
# #   'Horario Jueves': '8:00 am-11:30 am',
# #   'Horarios Viernes': '8:00 am-11:30 am, 1:00 pm-3:00 pm',
# #   'Horario Sabado': '',
# #   'Horarios Virtuales Lunes': '',
# #   'Horarios Virtuales Martes': '8:00 am-11:30 am',
# #   'Horarios Virtuales Miercoles': '8:00 am-11:30 am',
# #   'Horarios Virtuales Jueves': '8:00 am-11:30 am',
# #   'Horarios Virtuales Viernes': '8:00 am-11:30 am, 1:00 pm-3:00 pm',
# #   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'DIEGO ANDRES RIVERA VALLE',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '94744365',
#   'Clases': 'Calculo 1, Algebra Lineal, Calculo 2, Fisica 1, Fisica 3, Variable Compleja, Circuitos DC',
#   'Horario Lunes': '8:00 am - 2:40 pm, 4:00 pm - 5:20 pm',
#     'Horario Martes': '8:00 am - 2:40 pm, 4:00 pm - 5:20 pm',
#     'Horario Miercoles': '8:00 am - 2:40 pm, 4:00 pm - 5:20 pm',
#     'Horario Jueves': '8:00 am - 2:40 pm, 4:00 pm - 5:20 pm',
#     'Horarios Viernes': '8:00 am - 2:40 pm, 4:00 pm - 5:20 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'ANTHONY EMANUEL FUNEZ GUERRA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '33989502',
#   'Clases': 'Algebra, Quimica',
#   'Horario Lunes': '2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm, 5:20 pm - 6:40 pm',
#   'Horario Martes': '2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm, 5:20 pm - 6:40 pm',
#   'Horario Miercoles': '2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm, 5:20 pm - 6:40 pm',
#   'Horario Jueves': '2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm, 5:20 pm - 6:40 pm',
#   'Horarios Viernes': '11:15 am - 12:35 pm, 2:40 am - 4:00 pm, 4:00 am - 5:20 pm, 5:20 am - 6:40 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'IAN ROMAN BELTRAND PADILLA',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '31693049',
#   'Clases': 'Ecuaciones Diferenciales',
#   'Horario Lunes': '9:55 am - 11:15 am, 4:00 pm - 5:20 pm',
#   'Horario Martes': '9:55 am - 11:15 am, 4:00 pm - 5:20 pm',
#   'Horario Miercoles': '9:55 am - 11:15 am, 4:00 pm - 5:20 pm',
#   'Horario Jueves': '9:55 am - 11:15 am, 4:00 pm - 5:20 pm',
#   'Horarios Viernes': '9:55 am - 11:15 am, 11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
#  {'Tutor': 'JUAN SEBASTIAN MELENDEZ CRUZ',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '98722061',
#   'Clases': 'Circuitos DC',
#   'Horario Lunes': '2:40 pm - 4:00 pm',
#   'Horario Martes': '2:40 pm - 4:00 pm',
#   'Horario Miercoles': '2:40 pm - 4:00 pm',
#   'Horario Jueves': '2:40 pm - 4:00 pm',
#   'Horarios Viernes': '2:40 pm -4:00 pm, 6:40 pm - 8:00 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
# ]








#Plantilla
DatosTutoresPayload = {
    'Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA',
    'Habilitado/Deshabilitado': 'Yes',
    'Email': '',
    'Telefono': '',
    'Clases': 'Algebra',
    'Horario Lunes': '1:20 pm - 2:40 pm',
    'Horario Martes': '1:20 pm - 2:40 pm',
    'Horario Miercoles': '1:20 pm - 2:40 pm',
    'Horario Jueves': '1:20 pm - 2:40 pm',
    'Horarios Viernes': '9:55 am - 11:15 am, 1:20 pm - 2:40 pm',
    'Horario Sabado': '',
    'Horarios Virtuales Lunes': '',
    'Horarios Virtuales Martes': '',
    'Horarios Virtuales Miercoles': '',
    'Horarios Virtuales Jueves': '',
    'Horarios Virtuales Viernes': '',
    'Horarios Virtuales Sabado': ''
}

def parseWhatsappNumber(message):
    pattern = r'https://wa\.me/(\d+)\?text=(.*)'
    match = re.search(pattern, message)
    if match:
        number = match.group(1)
        return number

name = input("Ingrese el nombre del archivo de Excel (ejemplo: 'Tutorias Q1 2023'): ")
hoja = input("Ingrese el nombre de la hoja de Excel (ejemplo: 'Tutores'): ")
excel_data_df = pandas.read_excel(name + '.xlsx', sheet_name=hoja)

structure = {
    'Tutor': excel_data_df.Nombre,
    'Habilitado/Deshabilitado': excel_data_df['Habilitado/Deshabilitado'],
    'Email': len(excel_data_df.Nombre)*" ",
    'Telefono': excel_data_df['Enlace Whatsapp'],
    'Clases': excel_data_df.Clases,
    'Horario Lunes': excel_data_df.Lunes,
    'Horario Martes': excel_data_df.Martes,
    'Horario Miercoles': excel_data_df.Miercoles,
    'Horario Jueves': excel_data_df.Jueves,
    'Horarios Viernes': excel_data_df.Viernes,
    'Horario Sabado': excel_data_df.Sabado,
    'Horarios Virtuales Lunes': len(excel_data_df.Nombre)*" ",
    'Horarios Virtuales Martes': len(excel_data_df.Nombre)*" ",
    'Horarios Virtuales Miercoles': len(excel_data_df.Nombre)*" ",
    'Horarios Virtuales Jueves': len(excel_data_df.Nombre)*" ",
    'Horarios Virtuales Viernes': len(excel_data_df.Nombre)*" ",
    'Horarios Virtuales Sabado': len(excel_data_df.Nombre)*" "
}

tutores = []

for i in range(len(excel_data_df.Nombre)):
    tutor = {
        'Tutor': structure['Tutor'][i],
        'Habilitado/Deshabilitado': structure['Habilitado/Deshabilitado'][i],
        'Email': "",
        'Telefono': parseWhatsappNumber(structure['Telefono'][i]),
        'Clases': structure['Clases'][i],
        'Horario Lunes': structure['Horario Lunes'][i],
        'Horario Martes': structure['Horario Martes'][i],
        'Horario Miercoles': structure['Horario Miercoles'][i],
        'Horario Jueves': structure['Horario Jueves'][i],
        'Horarios Viernes': structure['Horarios Viernes'][i],
        'Horario Sabado': structure['Horario Sabado'][i],
        'Horarios Virtuales Lunes': "",
        'Horarios Virtuales Martes': "",
        'Horarios Virtuales Miercoles': "",
        'Horarios Virtuales Jueves': "",
        'Horarios Virtuales Viernes': "",
        'Horarios Virtuales Sabado': ""
    }
    tutores.append(tutor)

crearTutoresConRangosHorarios(tutores)


# ==================================================================================================