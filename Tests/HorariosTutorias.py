import re
import os
from dotenv import load_dotenv
from datetime import datetime
from shareplum import Site
from shareplum import Office365
import random

# Cargar las variables de entorno desde el archivo .env
load_dotenv()
userSh = os.getenv("SHAREPOINT_USER")
passSh = os.getenv("SHAREPOINT_PASS")

# Cargar la información de los tutores desde un archivo JSON
script_dir = os.path.dirname(__file__)
json_file_path = os.path.join(script_dir, "Datos", "HorarioTutores.json")
import json

DatosTutoresPayload = {}
with open(json_file_path, "r", encoding="utf-8") as file:
    DatosTutoresPayload = json.load(file)

authcookie = Office365(
    "https://unitechn.sharepoint.com", username=userSh, password=passSh
).GetCookies()
site = Site(
    "https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/", authcookie=authcookie
)

sp_list_Tutorias = site.List("Tutorias")
sp_list_Tutores = site.List("Tutores")
sp_list_Aulas = site.List("Aulas")

Tutoriasdata = sp_list_Tutorias.GetListItems(
    fields=[
        "ID",
        "Aula",
        "Tipo de Tutoria",
        "Contactado",
        "Estado",
        "Telefono",
        "Nombre Tutor",
        "Fecha de Tutoria",
        "Hora Tutoria",
        "Clases",
        "Temas",
        "Alumnos",
        "TutoresRechazaron",
    ]
)

Tutoresdata = sp_list_Tutores.GetListItems(
    fields=[
        "ID",
        "Tutor",
        "Telefono",
        "TelefonoAuxiliar",
        "Habilitado/Deshabilitado",
        "Clases que Imparte",
        "Horario Lunes",
        "Horario Martes",
        "Horario Miercoles",
        "Horario Jueves",
        "Horarios Viernes",
        "Horario Sabado",
    ]
)

random.shuffle(Tutoresdata)  # Mezclar lista de tutores
Aulasdata = sp_list_Aulas.GetListItems(fields=["ID", "IdAula ", "Oficial"])


def actulizarHorariosTutores(DatosTutores):
    for tutor in DatosTutores:
        for t in Tutoresdata:
            if tutor["Tutor"].upper() == t["Tutor"].upper():
                sp_list_Tutores.UpdateListItems(
                    data=[
                        {
                            "ID": t["ID"],
                            "Horario Lunes": tutor["Horario Lunes"],
                            "Horario Martes": tutor["Horario Martes"],
                            "Horario Miercoles": tutor["Horario Miercoles"],
                            "Horario Jueves": tutor["Horario Jueves"],
                            "Horarios Viernes": tutor["Horarios Viernes"],
                            "Horario Sabado": tutor["Horario Sabado"],
                        }
                    ]
                )
                break


def obtenerClasesDesdeSharePoint():
    """
    Obtiene los IDs y nombres de clases desde SharePoint.

    :return: Diccionario donde las claves son los nombres de las clases y los valores sus IDs.
    """
    try:
        # Consulta-la lista de clases en SharePoint
        sp_list_Clases = site.List(
            "Clases"
        )  # Asegúrate de que este sea el nombre correcto
        clases_data = sp_list_Clases.GetListItems(
            fields=["ID", "Nombre de Clase"]
        )  # Ajusta 'NombreClase' según el campo real
        # Crear un diccionario {NombreClase: ID}
        clases_dict = {clase["Nombre de Clase"]: clase["ID"] for clase in clases_data}
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
        for clase in clases_tutor
        if clase in clases_dict
    ]
    return "".join(partes) if partes else ""


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
        clases_tutor = tutor.get("Clases", "").split(
            ", "
        )  # Manejar clases separadas por comas
        # Generar el string en el formato necesario

        print(clases_tutor)
        clases_formato = generarStringClases(clases_tutor, clases_dict)

        if not clases_formato:
            print(f"Tutor {tutor['Tutor']} no tiene clases válidas, omitiendo.")
            continue

        # Crear o actualizar al tutor en SharePoint
        try:
            sp_list_Tutores.UpdateListItems(
                data=[
                    {
                        "Tutor": tutor["Tutor"].upper(),
                        "Telefono": tutor["Telefono"],
                        "Habilitado/Deshabilitado": tutor["Habilitado/Deshabilitado"],
                        "Clases que Imparte": clases_formato,
                    }
                ],
                kind="New",
            )
            print(
                f"Tutor {tutor['Tutor']} agregado correctamente con clases: {clases_formato}"
            )
        except Exception as e:
            print(f"Error al agregar al tutor {tutor['Tutor']}: {e}")


def obtenerHorariosDesdeSharePoint():
    """
    Obtiene los IDs y horarios desde SharePoint.

    :return: Diccionario donde las claves son los horarios y los valores sus IDs.
    """
    try:
        # Consulta-la lista de horarios en SharePoint
        sp_list_Horarios = site.List(
            "Horarios Contenido"
        )  # Asegúrate de que este sea el nombre correcto
        horarios_data = sp_list_Horarios.GetListItems(
            fields=["ID", "Horarios"]
        )  # Ajusta 'Horario' según el campo real
        # Crear un diccionario {Horario: ID}
        horarios_dict = {
            horario["Horarios"]: horario["ID"] for horario in horarios_data
        }
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
        for horario in horarios_tutor
        if horario in horarios_dict
    ]
    return "".join(partes) if partes else ""


def procesarRangosHorarios(rangos_horarios):
    """
    Divide un string de rangos de horarios separados por comas en una lista de rangos individuales.

    :param rangos_horarios: String con rangos de horarios, separados por comas.
    :return: Lista de rangos de horarios individuales.
    """
    return [rango.strip() for rango in rangos_horarios.split(",") if rango.strip()]


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
            inicio, fin = [
                datetime.strptime(h.strip(), "%I:%M %p") for h in rango.split("-")
            ]
        except ValueError:
            print(f"Formato de rango inválido: {rango}")
            continue

        for horario, id_horario in horarios_dict.items():
            if horario.endswith("[V]"):
                continue  # Ignorar horarios que finalicen con '[V]'

            try:
                hora_inicio = datetime.strptime(horario.split(" - ")[0], "%I:%M %p")
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
            inicio, fin = [
                datetime.strptime(h.strip(), "%I:%M %p") for h in rango.split("-")
            ]
        except ValueError:
            print(f"Formato de rango inválido: {rango}")
            continue

        for horario, id_horario in horarios_dict.items():
            if not horario.endswith("[V]"):
                continue  # Ignorar horarios que no terminan con '[V]'

            try:
                hora_inicio = datetime.strptime(horario.split(" - ")[0], "%I:%M %p")
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
        clases_tutor = tutor.get("Clases", "").split(", ")
        clases_formato = generarStringClases(clases_tutor, clases_dict)

        # Manejar horarios por rango
        rango_horario_lunes = tutor.get("Horario Lunes", "")
        rango_horario_martes = tutor.get("Horario Martes", "")
        rango_horario_miercoles = tutor.get("Horario Miercoles", "")
        rango_horario_jueves = tutor.get("Horario Jueves", "")
        rango_horario_viernes = tutor.get("Horarios Viernes", "")
        rango_horario_sabado = tutor.get("Horario Sabado", "")

        # Filtrar horarios virtuales por rango
        rango_horario_virtual_lunes = tutor.get("Horarios Virtuales Lunes", "")
        rango_horario_virtual_martes = tutor.get("Horarios Virtuales Martes", "")
        rango_horario_virtual_miercoles = tutor.get("Horarios Virtuales Miercoles", "")
        rango_horario_virtual_jueves = tutor.get("Horarios Virtuales Jueves", "")
        rango_horario_virtual_viernes = tutor.get("Horarios Virtuales Viernes", "")
        rango_horario_virtual_sabado = tutor.get("Horarios Virtuales Sabado", "")

        horarios_en_rango_lunes = filtrarHorariosSinV(
            rango_horario_lunes, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_lunes, horarios_dict)
        horarios_en_rango_martes = filtrarHorariosSinV(
            rango_horario_martes, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_martes, horarios_dict)
        horarios_en_rango_miercoles = filtrarHorariosSinV(
            rango_horario_miercoles, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_miercoles, horarios_dict)
        horarios_en_rango_jueves = filtrarHorariosSinV(
            rango_horario_jueves, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_jueves, horarios_dict)
        horarios_en_rango_viernes = filtrarHorariosSinV(
            rango_horario_viernes, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_viernes, horarios_dict)
        horarios_en_rango_sabado = filtrarHorariosSinV(
            rango_horario_sabado, horarios_dict
        ) + filtrarHorariosConV(rango_horario_virtual_sabado, horarios_dict)
        # + filtrarHorariosConV(rango_horario, horarios_dict)
        horarios_formato_lunes = (
            "".join(horarios_en_rango_lunes) if horarios_en_rango_lunes else ""
        )
        horarios_formato_martes = (
            "".join(horarios_en_rango_martes) if horarios_en_rango_martes else ""
        )
        horarios_formato_miercoles = (
            "".join(horarios_en_rango_miercoles) if horarios_en_rango_miercoles else ""
        )
        horarios_formato_jueves = (
            "".join(horarios_en_rango_jueves) if horarios_en_rango_jueves else ""
        )
        horarios_formato_viernes = (
            "".join(horarios_en_rango_viernes) if horarios_en_rango_viernes else ""
        )
        horarios_formato_sabado = (
            "".join(horarios_en_rango_sabado) if horarios_en_rango_sabado else ""
        )

        if not clases_formato:
            print(
                f"Tutor {tutor['Tutor']} no tiene clases u horarios válidos, omitiendo."
            )
            continue

        # Crear o actualizar al tutor en SharePoint
        try:
            sp_list_Tutores.UpdateListItems(
                data=[
                    {
                        "Tutor": tutor["Tutor"].upper(),
                        "Telefono": tutor["Telefono"],
                        "Habilitado/Deshabilitado": tutor["Habilitado/Deshabilitado"],
                        "Clases que Imparte": clases_formato,
                        "Horarios de Tutor": horarios_formato_lunes,
                        "Horario Lunes": horarios_formato_lunes,
                        "Horario Martes": horarios_formato_martes,
                        "Horario Miercoles": horarios_formato_miercoles,
                        "Horario Jueves": horarios_formato_jueves,
                        "Horarios Viernes": horarios_formato_viernes,
                        "Horario Sabado": horarios_formato_sabado,
                    }
                ],
                kind="New",
            )
            print(
                f"Tutor {tutor['Tutor']} agregado correctamente con clases y horarios."
            )
        except Exception as e:
            print(f"Error al agregar al tutor {tutor['Tutor']}: {e}")


def parseWhatsappNumber(message):
    pattern = r"https://wa\.me/(\d+)\?text=(.*)"
    match = re.search(pattern, message)
    if match:
        number = match.group(1)
        return number


#Menu para agregar tutores
#Tipos de tutores: DatosTutoresPayload, DatosTutoresPayloadProvisional, DatosTutoresPayloadPAE, Plantilla
def menuAgregarTutores():
    print("Bienvenido al menú de agregación de tutores.")
    print("Selecciona el tipo de tutores a agregar:")
    print("1. DatosTutoresPayload")
    print("2. DatosTutoresPayloadProvisional")
    print("3. DatosTutoresPayloadPAE")
    print("4. Plantilla")

    opcion = input("Ingresa el número de la opción: ")

    if opcion == "1":
        print("Agregando tutores con DatosTutoresPayload...")
        crearTutoresConRangosHorarios(DatosTutoresPayload.get("DatosTutoresPayload", []))
    elif opcion == "2":
        print("Agregando tutores con DatosTutoresPayloadProvisional...")
        crearTutoresConRangosHorarios(DatosTutoresPayload.get("DatosTutoresPayloadProvisional", []))
    elif opcion == "3":
        print("Agregando tutores con DatosTutoresPayloadPAE...")
        crearTutoresConRangosHorarios(DatosTutoresPayload.get("DatosTutoresPayloadPAE", []))
    elif opcion == "4":
        print("Mostrando plantilla de tutores...")
        print("Plantilla de tutores:")
        print(DatosTutoresPayload.get("Plantilla", []))
    else:
        print("Opción no válida.")

menuAgregarTutores()

# ==================================================================================================
