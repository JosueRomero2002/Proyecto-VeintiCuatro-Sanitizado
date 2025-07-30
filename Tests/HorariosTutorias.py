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

# Autenticación en SharePoint
import os
from dotenv import load_dotenv

load_dotenv()  
userSh = os.getenv('SHAREPOINT_USER')
passSh = os.getenv('SHAREPOINT_PASS')

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
# ==================================================================================================

# # Datos Tutores
DatosTutoresPayload = [
# {'Tutor': 'ERICK EDUARDO ARITA HENRIQUEZ',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '97002422',
#   'Clases': 'Intro al algebra, Algebra, Geometria y trigonometria, Calculo 1, Calculo 2, Algebra lineal, Ecuaciones diferenciales , Fisica 2, Fisica 3, Circuitos DC, Circuitos AC, Mecanica de Fluidos, Electronica 1, Electronica 2',
#   'Horario Lunes': '1:20 pm - 5:20 pm',
#   'Horario Martes': '1:20 pm - 5:20 pm',
#   'Horario Miercoles': '1:20 pm - 5:20 pm',
#   'Horario Jueves': '1:20 pm - 5:20 pm',
#   'Horarios Viernes': '1:20 pm - 5:20 pm',
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
#   'Clases': 'Electronica 1, Algebra, Introducción al álgebra, Ofimatica avanzada',
#   'Horario Lunes': '8:10 am - 9:55 am, 11:15 am - 12:35 pm, 01:20 am - 4:00 pm',
#   'Horario Martes': '8:10 am - 9:55 am, 11:15 am - 12:35 pm, 01:20 am - 4:00 pm',
#   'Horario Miercoles': '8:10 am - 9:55 am, 11:15 am - 12:35 pm',
#   'Horario Jueves': '8:10 am - 9:55 am, 11:15 am - 12:35 pm, 01:20 pm - 4:00 pm ',
#   'Horarios Viernes': '8:10 am - 12:35 pm, 1:20 pm - 4:00 pm, 5:20 pm - 8:00 pm',
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
#   'Clases': 'Programación 1, Programación 2, Programación 3, Lab. de Programación 1, Lab. de Programación 2, Lab. de Programación 3, Programación Estructurada, Lógica de Programación, Teoría de Base de Datos 1, Teoría de Base de Datos 2, Intro. a la Computacion y Ciencia de Datos, Fisica 1',
#   'Horario Lunes': '12:00 pm - 4:00 pm, 5:20 pm - 6:40 pm',
#   'Horario Martes': '12:00 pm - 4:00 pm, 5:20 pm - 6:40 pm',
#   'Horario Miercoles': '12:00 pm - 4:00 pm, 5:20 pm - 6:40 pm',
#   'Horario Jueves': '12:00 pm - 4:00 pm, 5:20 pm - 6:40 pm',
#   'Horarios Viernes': '12:00 pm - 4:00 pm, 5:20 pm - 6:40 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},

#  {'Tutor': 'ASHLY GISSELE MARTINEZ GRANADOS',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '97033933',
#   'Clases': 'Intro a Algebra, Algebra, Geometria y Trigonometria',
#   'Horario Lunes': '11:15 am - 5:20 pm',
#   'Horario Martes': '11:15 am - 5:20 pm',
#   'Horario Miercoles': '11:15 am - 5:20 pm',
#   'Horario Jueves': '11:15 am - 5:20 pm',
#   'Horarios Viernes': '8:10 am - 9:30 am',
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
#   'Clases': 'Introducción al álgebra, Álgebra, Geometría y Trigonometría, Álgebra Lineal, Cálculo 1, Cálculo 2, Dibujo Técnico (Solidworks), Variable Compleja, Ecuaciones Diferenciales, Física 2, Física 3, Estática, Dinámica, Resistencia de Materiales 1, Matemáticas Discretas, Programación C++, Análisis de Circuitos en DC, Análisis de Circuitos en AC, Electrónica 1, Electrónica 2, Compuertas Lógicas, Transformadores y motores eléctricos',
#   'Horario Lunes': '1:20 pm - 2:40 pm, 6:40 pm - 8:00 pm',
#   'Horario Martes': '1:20 pm - 2:40 pm, 6:40 pm - 8:00 pm',
#   'Horario Miercoles': '1:20 pm - 2:40 pm, 6:40 pm - 8:00 pm',
#   'Horario Jueves': '1:20 pm - 2:40 pm, 6:40 pm - 8:00 pm',
#   'Horarios Viernes': '1:20 pm - 6:40 pm',
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
#   'Clases': 'Álgebra Lineal, Intro al Algebra',
#   'Horario Lunes': '6:00 pm - 7:00 pm',
#   'Horario Martes': '8:30 am - 10:30 am',
#   'Horario Miercoles': '8:30 am - 10:30 am',
#   'Horario Jueves': '8:30 am - 10:30 am',
#   'Horarios Viernes': '3:00 pm - 5:00 pm',
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

#  {'Tutor': 'DANIEL ALEXANDER HERNANDEZ ESCOTO',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '99825386',
#   'Clases': 'Algebra, Intro al Algebra, Geometria y trigonometria, Calculo 1, Fisica 2, Circuitos DC',
#   'Horario Lunes': '7:00 am - 12:35 pm',
#   'Horario Martes': '7:00 am - 12:35 pm',
#   'Horario Miercoles': '7:00 am - 12:35 pm',
#   'Horario Jueves': '7:00 am - 12:35 pm',
#   'Horarios Viernes': '7:00 am - 11:15 am',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},

#  {'Tutor': 'DIEGO ANDRES RIVERA VALLE',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '94744365',
#   'Clases': 'Fisica 1, Fisica 2, Fisica 3, Dibujo Tecnico, Ecuaciones diferenciales, Variable Compleja, Matematicas Discretas, Estatica, Dinamica, Analisis de Circuitos DC, Analisis de Circuitos AC, Termodinamica, Geometria, Electronica 1, Sensores y actuadores, Maquinas Hidraulicas, Mecanismos, Mecanica de materiales',
#     'Horario Lunes': '9:55 am - 4:00 pm',
#     'Horario Martes': '12:35 pm - 4:00 pm',
#     'Horario Miercoles': '9:55 am - 4:00 pm',
#     'Horario Jueves': '2:40 pm - 4:00 pm',
#     'Horarios Viernes': '12:35 pm - 4:00 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': '9:55 am - 4:00 pm'},

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
#   'Clases': 'Intro al Algebra, Algebra, Geometria y Trigonometria, Calculo 1, Calculo 2, Ecuaciones Diferenciales, Matematicas Discretas, Algebra lineal, Fisica 1, Fisica 2, Estadistica Matematica 1, Estadistica Matematica 2, Investigacion de Operaciones',
#   'Horario Lunes': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horario Martes': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horario Miercoles': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horario Jueves': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horarios Viernes': '11:15 am - 12:35 pm, 1:20 pm - 2:40 pm, 2:40 pm - 4:00 pm, 4:00 pm - 5:20 pm',
#   'Horario Sabado': '',
#  'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},



# {'Tutor': 'DANILO ENRIQUE ALVARADO HERNANDEZ',
#     'Habilitado/Deshabilitado': 'Yes',
#     'Email': '',
#     'Telefono': '31959108',
#     'Clases': 'Ecuaciones diferenciales, Variable compleja, Electronica 1, Electronica 2, Teoria de control, Intro al Algebra',
#     'Horario Lunes': '10:00 am - 11:00 am',
#     'Horario Martes': '10:00 am - 11:00 am',
#     'Horario Miercoles': '10:00 am - 11:00 am',
#     'Horario Jueves': '10:00 am - 11:00 am',
#     'Horarios Viernes': '10:00 am - 11:00 am',
#     'Horario Sabado': '',
#     'Horarios Virtuales Lunes': '',
#     'Horarios Virtuales Martes': '',
#     'Horarios Virtuales Miercoles': '',
#     'Horarios Virtuales Jueves': '',
#     'Horarios Virtuales Viernes': '',
#     'Horarios Virtuales Sabado': ''},

    # {'Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '95889108',
    # 'Clases': 'Introduccion al Algebra, Algebra, Geometria y trigonometria, Calculo 1, Calculo 2',
    # 'Horario Lunes': '11:15 am - 12:35 pm',
    # 'Horario Martes': '11:15 am - 12:35 pm',
    # 'Horario Miercoles': '11:15 am - 12:35 pm',
    # 'Horario Jueves': '11:15 am - 12:35 pm, 4:00 pm - 5:20 pm',
    # 'Horarios Viernes': '',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

    # {'Tutor': 'JABES DANIEL LOPEZ ROMERO',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '33770733',
    # 'Clases': 'Intro al Algebra, Algebra, Geometria y Trigonometria, Calculo 1, Algebra Lineal, Quimica',
    # 'Horario Lunes': '8:00 am - 12:00 pm, 4:00 pm - 5:00 pm',
    # 'Horario Martes': '8:00 am - 12:00 pm, 4:00 pm - 5:00 pm',
    # 'Horario Miercoles': '8:00 am - 12:00 pm, 4:00 pm - 5:00 pm',
    # 'Horario Jueves': '8:00 am - 12:00 pm, 4:00 pm - 5:00 pm',
    # 'Horarios Viernes': '8:00 am - 12:00 pm, 4:00 pm - 5:00 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': '' },

    # {'Tutor': 'JONATHAN ANDRES UMANA LOPEZ',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '96201348',
    # 'Clases': 'Intro al Algebra, Calculo 2, Ecuaciones Diferenciales, Fisica 1, Electronica 1, Electronica 2, Algebra, Calculo 1, Geometria, Dinamica, Estatica',
    # 'Horario Lunes': '8:10 am - 9:30 am, 9:55 am - 12:30 pm, 1:00 pm - 4:00 pm',
    # 'Horario Martes': '8:10 am - 9:30 am, 9:55 am - 12:30 pm, 1:00 pm - 4:00 pm',
    # 'Horario Miercoles': '8:10 am - 9:30 am, 9:55 am - 12:30 pm, 1:00 pm - 4:00 pm',
    # 'Horario Jueves': '8:10 am - 9:30 am, 9:55 am - 12:30 pm, 1:00 pm - 4:00 pm',
    # 'Horarios Viernes': '8:10 am - 12:30 pm, 1:00 pm - 8:00 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': '' },

    # {'Tutor': 'JOSUE MANUEL GAVIDIA GONZALEZ',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '33751005',
    # 'Clases': 'Intro al Algebra, Logica de Programación, Programación 1, Programación para Ingeniería',
    # 'Horario Lunes': '8:10 am - 1:20 pm, 2:40 pm - 5:20 pm, 6:40 pm - 8:00 pm',
    # 'Horario Martes': '9:30 am - 1:20 pm, 2:40 pm - 5:20 pm, 6:40 pm - 8:00 pm',
    # 'Horario Miercoles': '8:10 am - 1:20 pm, 2:40 pm - 5:20 pm, 6:40 pm - 8:00 pm',
    # 'Horario Jueves': '8:10 am - 1:20 pm, 2:40 pm - 5:20 pm, 6:40 pm - 8:00 pm',
    # 'Horarios Viernes': '8:10 am - 1:20 pm, 2:40 pm - 5:20 pm, 6:40 pm - 8:00 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': '' },

    # {'Tutor': 'OSVALDO ANDRES DACOSTA FERNANDEZ',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '95086037',
    # 'Clases': 'Estadistica Matematica 1, Introducción a Finanzas y Economía',
    # 'Horario Lunes': '4:00 pm - 5:20 pm, 6:40 pm - 7:40 pm',
    # 'Horario Martes': '4:00 pm - 5:20 pm, 6:40 pm - 7:40 pm',
    # 'Horario Miercoles': '4:00 pm - 5:20 pm, 6:40 pm - 7:40 pm',
    # 'Horario Jueves': '4:00 pm - 5:0 pm, 6:40 pm - 7:40 pm',
    # 'Horarios Viernes': '',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

    # {'Tutor': 'CARLOS MARCELO ALVARADO RIVAS',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '97897906',
    # 'Clases': 'Quimica, Calculo 1, Calculo 2, Fisica 1, Ecuaciones diferenciales, Variable Compleja, Estatica, Análisis de circuitos en DC, Análisis de circuitos en AC',
    # 'Horario Lunes': '',
    # 'Horario Martes': '10:00 am - 12:00 pm',
    # 'Horario Miercoles': '',
    # 'Horario Jueves': '10:00 am - 12:00 pm',
    # 'Horarios Viernes': '5:00 pm - 7:00 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

    # {'Tutor': 'CARLOS ENRIQUE LOPEZ ARITA',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '98437598',
    # 'Clases': 'Geometria y Trigonometria, Calculo 1',
    # 'Horario Lunes': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Martes': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Miercoles': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Jueves': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horarios Viernes': '',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

    # {'Tutor': 'JENNIFER LAGOS HERNÁNDEZ',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '89206905',
    # 'Telefono': '',
    # 'Clases': 'Intro al algebra, Geometria y Trigonometria',
    # 'Horario Lunes': '6:45 am - 8:10 am, 12:30 pm - 1:20 pm, 2:40 pm - 6:40 pm',
    # 'Horario Martes': '6:45 am - 8:10 am, 12:30 pm - 1:20 pm, 2:40 pm - 6:40 pm',
    # 'Horario Miercoles': '6:45 am - 8:10 am, 12:30 pm - 1:20 pm, 2:40 pm - 6:40 pm',
    # 'Horario Jueves': '6:45 am - 8:10 am, 12:30 pm - 1:20 pm, 2:40 pm - 6:40 pm',
    # 'Horarios Viernes': '6:45 am - 8:10 am, 11:15 am - 5:20 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},
#-----------------------------------------------------------------------------------------------------

#DATOS LISTIS PERO DICE PROVICIONAL
#  {'Tutor': 'JUAN SEBASTIAN MELENDEZ CRUZ',
#   'Habilitado/Deshabilitado': 'Yes',
#   'Email': '',
#   'Telefono': '98722061',
#   'Clases': 'Circuitos DC, Analisis de circuitos electricos 1, Analisis de circuitos electricos 2',
#   'Horario Lunes': '1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Martes': '1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Miercoles': '1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Jueves': '1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horarios Viernes': '1:20 pm - 2:40 pm, 4:00 pm - 6:40 pm',
#   'Horario Sabado': '',
#   'Horarios Virtuales Lunes': '',
#   'Horarios Virtuales Martes': '',
#   'Horarios Virtuales Miercoles': '',
#   'Horarios Virtuales Jueves': '',
#   'Horarios Virtuales Viernes': '',
#   'Horarios Virtuales Sabado': ''},
   

#DATOS LISTOS PERO DICE QUE SOLO PAI - OFIMATICA
    # {'Tutor': 'ALEX REYES GUTIERREZ',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '32407099',
    # 'Clases': 'Solo PAI - Ofimatica',
    # 'Horario Lunes': '1:20 pm - 5:20 pm',
    # 'Horario Martes': '',
    # 'Horario Miercoles': '1:20 pm - 5:20 pm',
    # 'Horario Jueves': '1:20 pm - 5:20 pm',
    # 'Horarios Viernes': '1:20 pm - 5:20 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''}

#DATOS LISTOS PERO EL TELEFONO NO ESTA EN EL LINK
    # {'Tutor': 'ANA VALERIA FLORES GALINDO',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '33118152',
    # 'Clases': 'Etica y Ciudadanía',
    # 'Horario Lunes': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Martes': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Miercoles': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horario Jueves': '8:10 am - 9:30 am, 9:55 am - 11:15 am',
    # 'Horarios Viernes': '',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

#DATOS LISTOS PERO DICE QUE SOLO PAI - DIBUJO
    # {'Tutor': 'Yasin Hernandez',
    # 'Habilitado/Deshabilitado': 'Yes',
    # 'Email': '',
    # 'Telefono': '87668930',
    # 'Clases': 'Dibujo',
    # 'Horario Lunes': '11:15 am - 12:00 pm, 3:00 pm - 4:00 pm',
    # 'Horario Martes': '11:15 am - 12:00 pm, 3:00 pm - 4:00 pm',
    # 'Horario Miercoles': '11:15 am - 12:00 pm, 3:00 pm - 4:00 pm',
    # 'Horario Jueves': '11:15 am - 12:00 pm, 3:00 pm - 4:00 pm',
    # 'Horarios Viernes': '11:15 am - 12:00 pm, 3:00 pm - 4:00 pm',
    # 'Horario Sabado': '',
    # 'Horarios Virtuales Lunes': '',
    # 'Horarios Virtuales Martes': '',
    # 'Horarios Virtuales Miercoles': '',
    # 'Horarios Virtuales Jueves': '',
    # 'Horarios Virtuales Viernes': '',
    # 'Horarios Virtuales Sabado': ''},

#DATOS EL LINK DE WS SU TELEFONO NO FUNCIONA
#     {'Tutor': 'JAVIER MAURICIO SANCHEZ GUTIERREZ',
#     'Habilitado/Deshabilitado': 'Yes',
#     'Email': '',
#     'Telefono': '',
#     'Clases': 'Algebra para Lic, Estadistica 1, Contabilidad',
#     'Horario Lunes': '9:20 am - 11:00 am, 2:40 pm - 4:00 pm',
#     'Horario Martes': '9:20 am - 11:00 am, 2:40 pm - 4:00 pm',
#     'Horario Miercoles': '9:20 am - 11:00 am, 2:40 pm - 4:00 pm',
#     'Horario Jueves': '9:20 am - 11:00 am, 2:40 pm - 4:00 pm',
#     'Horarios Viernes': '',
#     'Horario Sabado': '',
#     'Horarios Virtuales Lunes': '',
#     'Horarios Virtuales Martes': '',
#     'Horarios Virtuales Miercoles': '',
#     'Horarios Virtuales Jueves': '',
#     'Horarios Virtuales Viernes': '',
#     'Horarios Virtuales Sabado': ''
# }

    #DATOS LISTOS PERO DICE PROVICIONAL
 {'Tutor': 'ISAAC MAURICIO JUAREZ BACA',
  'Habilitado/Deshabilitado': 'Yes',
  'Email': '',
  'Telefono': '97949526',
  'Clases': 'Intro al Algebra, Calculo 1, Calculo 2',
  'Horario Lunes': '6:40 pm - 8:00 pm',
  'Horario Martes': '6:40 pm - 8:00 pm',
  'Horario Miercoles': '6:40 pm - 8:00 pm',
  'Horario Jueves': '6:40 pm - 8:00 pm',
  'Horarios Viernes': '2:40 pm - 5:20 pm ',
  'Horario Sabado': '',
  'Horarios Virtuales Lunes': '',
  'Horarios Virtuales Martes': '',
  'Horarios Virtuales Miercoles': '',
  'Horarios Virtuales Jueves': '',
  'Horarios Virtuales Viernes': '',
  'Horarios Virtuales Sabado': ''},


   
 ]



# #Plantilla
# # DatosTutoresPayload = {'Tutor': 'FERNANDO JOSUÉ PORTILLO PEÑA',
#     'Habilitado/Deshabilitado': 'Yes',
#     'Email': '',
#     'Telefono': '',
#     'Clases': 'Algebra',
#     'Horario Lunes': '1:20 pm - 2:40 pm',
#     'Horario Martes': '1:20 pm - 2:40 pm',
#     'Horario Miercoles': '1:20 pm - 2:40 pm',
#     'Horario Jueves': '1:20 pm - 2:40 pm',
#     'Horarios Viernes': '9:55 am - 11:15 am, 1:20 pm - 2:40 pm',
#     'Horario Sabado': '',
#     'Horarios Virtuales Lunes': '',
#     'Horarios Virtuales Martes': '',
#     'Horarios Virtuales Miercoles': '',
#     'Horarios Virtuales Jueves': '',
#     'Horarios Virtuales Viernes': '',
#     'Horarios Virtuales Sabado': ''},

# def parseWhatsappNumber(message):
#     pattern = r'https://wa\.me/(\d+)\?text=(.*)'
#     match = re.search(pattern, message)
#     if match:
#         number = match.group(1)
#         return number

# name = input("Ingrese el nombre del archivo de Excel (ejemplo: 'Tutorias Q1 2023'): ")
# hoja = input("Ingrese el nombre de la hoja de Excel (ejemplo: 'Tutores'): ")
# excel_data_df = pandas.read_excel(name + '.xlsx', sheet_name=hoja)
# print(excel_data_df)
# structure = {
#     'Tutor': excel_data_df.Nombre,
#     'Habilitado/Deshabilitado': excel_data_df['Actualizado en power apps'],
#     'Email': len(excel_data_df.Nombre)*" ",
#     'Telefono': excel_data_df['Enlace Whatsapp'],
#     'Clases': excel_data_df.Clases,
#     'Horario Lunes': excel_data_df.Lunes,
#     'Horario Martes': excel_data_df.Martes,
#     'Horario Miercoles': excel_data_df.Miercoles,
#     'Horario Jueves': excel_data_df.Jueves,
#     'Horarios Viernes': excel_data_df.Viernes,
#     'Horario Sabado': excel_data_df['Sábado'],
#     'Horarios Virtuales Lunes': len(excel_data_df.Nombre)*" ",
#     'Horarios Virtuales Martes': len(excel_data_df.Nombre)*" ",
#     'Horarios Virtuales Miercoles': len(excel_data_df.Nombre)*" ",
#     'Horarios Virtuales Jueves': len(excel_data_df.Nombre)*" ",
#     'Horarios Virtuales Viernes': len(excel_data_df.Nombre)*" ",
#     'Horarios Virtuales Sabado': len(excel_data_df.Nombre)*" "
# }

# tutores = []

# for i in range(len(excel_data_df.Nombre)):
#     # skip if its a nan
#     if structure['Clases'][i] == "" or structure['Clases'][i] == " ":
#         print(f"El tutor {structure['Tutor'][i]} no tiene clases asignadas, omitiendo.")
#         break
#     tutor = {
#         'Tutor': structure['Tutor'][i],
#         'Habilitado/Deshabilitado': structure['Habilitado/Deshabilitado'][i],
#         'Email': "",
#         'Telefono': parseWhatsappNumber(structure['Telefono'][i]),
#         'Clases': structure['Clases'][i],
#         'Horario Lunes': structure['Horario Lunes'][i],
#         'Horario Martes': structure['Horario Martes'][i],
#         'Horario Miercoles': structure['Horario Miercoles'][i],
#         'Horario Jueves': structure['Horario Jueves'][i],
#         'Horarios Viernes': structure['Horarios Viernes'][i],
#         'Horario Sabado': structure['Horario Sabado'][i],
#         'Horarios Virtuales Lunes': "",
#         'Horarios Virtuales Martes': "",
#         'Horarios Virtuales Miercoles': "",
#         'Horarios Virtuales Jueves': "",
#         'Horarios Virtuales Viernes': "",
#         'Horarios Virtuales Sabado': ""
#     }
#     print(f"Agregando tutor: {tutor['Tutor']} con telefono: {tutor['Telefono']}\n")
#     tutores.append(tutor)

# print(tutores)

crearTutoresConRangosHorarios(DatosTutoresPayload)




# ==================================================================================================