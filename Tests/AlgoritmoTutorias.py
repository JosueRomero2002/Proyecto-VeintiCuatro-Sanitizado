from datetime import datetime, timedelta
import time
from shareplum import Site
from shareplum import Office365
import random
import difflib
import os
from dotenv import load_dotenv

# Cargar las variables de entorno desde el archivo .env
load_dotenv()
userSh = os.getenv("SHAREPOINT_USER")
passSh = os.getenv("SHAREPOINT_PASS")

# Autenticación en SharePoint
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
        "HoraClasica",
        "ClaseClasica",
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

AulasHabilitadas = [aula for aula in Aulasdata if "Oficial" in aula]

TutoriasCoordinadasCONAULA = [
    tut
    for tut in Tutoriasdata
    if tut["Estado"] in ["Coordinada"] and "Aula" in tut and "Nombre Tutor" in tut
]

TutoriasCoordinadas = [
    tut
    for tut in Tutoriasdata
    if tut["Estado"] in ["Coordinada", "Confirmada"]
    and "Aula" in tut
    and "Nombre Tutor" in tut
]

TutoriasPendientesRechazadas = [
    tut for tut in Tutoriasdata if tut["Estado"] in ["Rechazada", "Pendiente"]
]


def tutor_disponible(tutor, fecha_tutoria, hora_tutoria):
    for tut in TutoriasCoordinadas:
        if (
            tut["Nombre Tutor"] == tutor
            and tut["Fecha de Tutoria"] == fecha_tutoria
            and HoraChoca(tut["Hora Tutoria"], hora_tutoria)
        ):
            return False
    for tut in TutoriasPendientesRechazadas:
        if (
            tut["Nombre Tutor"] == tutor
            and tut["Fecha de Tutoria"] == fecha_tutoria
            and HoraChoca(tut["Hora Tutoria"], hora_tutoria)
        ):
            return False
    return True


def aula_disponible(aula, fecha_tutoria, hora_tutoria):
    # Validar si el aula está disponible
    for tut in TutoriasCoordinadasCONAULA:
        if (
            tut["Aula"] == aula
            and tut["Fecha de Tutoria"] == fecha_tutoria
            and HoraChoca(tut["Hora Tutoria"], hora_tutoria)
        ):  # HoraChoca(tut['Hora Tutoria'], hora_tutoria):
            return False
    return True


def asignar_tutor(tutoria):
    fecha_tutoria = tutoria["Fecha de Tutoria"]
    hora_tutoria = None
    if "Hora Tutoria" in tutoria:
        hora_tutoria = tutoria["Hora Tutoria"]
    else:
        hora_tutoria = tutoria["HoraClasica"]
        tutoria["Hora Tutoria"] = hora_tutoria

    clase_tutoria = None
    if "Clases" in tutoria:
        clase_tutoria = tutoria["Clases"]
    else:
        clase_tutoria = tutoria["ClaseClasica"]
        tutoria["Clases"] = clase_tutoria

    now = datetime.now()
    now = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if not (fecha_tutoria >= (now)):
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
    if "TutoresRechazaron" in tutoria and tutoria["TutoresRechazaron"]:
        rechazaron_tutoria = tutoria["TutoresRechazaron"]

    tutor_asignado = None

    for (
        tutor
    ) in (
        Tutoresdata
    ):  # Algebra   in ['Algebra Lineal', 'Calculo', 'Fisica'] Error lo toma como true
        if (
            (not (tutor["Tutor"] in rechazaron_tutoria))
            and tutor["Habilitado/Deshabilitado"] == "Yes"
            and (clase_tutoria in tutor["Clases que Imparte"])
            and hora_tutoria in tutor[horario]
            and tutor_disponible(tutor["Tutor"], fecha_tutoria, hora_tutoria)
        ):
            tutor_asignado = tutor["Tutor"]
            break

    if tutor_asignado:
        tutoria["Nombre Tutor"] = tutor_asignado
        return True
    return False


def asignar_aula(tutoria):
    fecha_tutoria = tutoria["Fecha de Tutoria"]
    hora_tutoria = tutoria["Hora Tutoria"]

    aula_asignada = None
    for aula in AulasHabilitadas:
        if aula["Oficial"] == "Yes" and aula_disponible(
            aula["IdAula "], fecha_tutoria, hora_tutoria
        ):
            aula_asignada = aula["IdAula "]
            break

    if aula_asignada:
        tutoria["Aula"] = aula_asignada
        return True
    return False


def comparar_textos(texto1, texto2, umbral=0.6):
    # Calcular la similitud
    similaridad = difflib.SequenceMatcher(None, texto1, texto2).ratio()
    return similaridad >= umbral


def posibleUnion(tutoria):
    for tut in TutoriasCoordinadas:
        if comparar_textos(tut["Temas"], "PAE"):
            return False
        elif (
            comparar_textos(tut["Clases"], tutoria["Clases"])
            and comparar_textos(tut["Temas"], tutoria["Temas"])
            and tut["Fecha de Tutoria"] == tutoria["Fecha de Tutoria"]
            and tut["Hora Tutoria"] == tutoria["Hora Tutoria"]
        ):
            tutoria["Aula"] = tut["Aula"]
            tutoria["Nombre Tutor"] = tut["Nombre Tutor"]
            tutoria["Tutor Asignado"] = tut["Nombre Tutor"]
            tutoria["Tutor"] = tut["Nombre Tutor"]
            return True
    return False


def HoraChoca(hora1, hora2):
    # Formato de hora: "HH:MM xm - HH:MM xm"
    hora1 = hora1.split(" - ")
    hora2 = hora2.split(" - ")

    # Convertir a formato de 24 horas
    hora1[0] = time.strptime(hora1[0], "%I:%M %p")  # ejemplo 9:55 am
    hora1[1] = time.strptime(hora1[1], "%I:%M %p")  # ejemplo 11:15 am
    hora2[0] = time.strptime(hora2[0], "%I:%M %p")  # ejemplo 9:55 am
    hora2[1] = time.strptime(hora2[1], "%I:%M %p")  # ejemplo 11:15 am

    # Verificar si las horas chocan
    if (
        (hora1[0] <= hora2[0] and hora2[0] < hora1[1])
        or (hora1[0] < hora2[1] and hora2[1] <= hora1[1])
        or (hora2[0] <= hora1[0] and hora1[0] < hora2[1])
        or (hora2[0] < hora1[1] and hora1[1] <= hora2[1])
    ):
        return True  # Si Choca

    return False


def SolicitarTutoria(Solicitud):
    # Crear una nueva tutoría
    sp_list_Tutorias.UpdateListItems(data=Solicitud, kind="New")
    print("Tutoría solicitada")


def generar_solicitudes(
    tutor, clase, alumno, telefono, hora, dias, fecha_inicio, fecha_fin, temas="PAE"
):
    solicitudes = []
    fecha_actual = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    fecha_fin = datetime.strptime(fecha_fin, "%Y-%m-%d")

    dias_numeros = {
        "lunes": 0,
        "martes": 1,
        "miercoles": 2,
        "jueves": 3,
        "viernes": 4,
        "sabado": 5,
        "domingo": 6,
    }

    dias_seleccionados = {dias_numeros[dia.lower()] for dia in dias}

    while fecha_actual <= fecha_fin:
        if fecha_actual.weekday() in dias_seleccionados:
            solicitudes.append(
                {
                    "Tipo de Tutoria": "Presencial",
                    "Contactado": "No",
                    "Estado": "Confirmada",
                    "Telefono": telefono,
                    "Nombre Tutor": tutor,
                    "Fecha de Tutoria": fecha_actual.strftime("%Y-%m-%d"),
                    "Hora Tutoria": hora,
                    "Clases": clase,
                    "Temas": temas,
                    "Alumnos": alumno,
                }
            )
        fecha_actual += timedelta(days=1)

    return solicitudes


def generarTutoriaPAE():
    for i, tutor_data in enumerate(Tutoresdata):
        print(f"{i} - {tutor_data['Tutor']}\n")

    tutor_index = int(input("Ingrese opcion de tutor (#): "))
    if not (0 <= tutor_index < len(Tutoresdata)):
        print("Opción de tutor no válida.")
        return

    tutor = Tutoresdata[tutor_index]
    raw_clases = tutor["Clases que Imparte"].split(";#")
    available_classes = []
    for i, item in enumerate(raw_clases):
        if i % 2 != 0 and item:
            available_classes.append(item)

    if not available_classes:
        print(f"El tutor {tutor['Tutor']} no tiene clases disponibles para impartir.")
        return

    print("\nClases disponibles:")
    for i, class_name in enumerate(available_classes):
        print(f"{i} - {class_name}")

    clase_index = int(input("Ingrese opcion de clase (#): "))
    if not (0 <= clase_index < len(available_classes)):
        print("Opción de clase no válida.")
        return

    clase_seleccionada = available_classes[clase_index]

    alumno = input("Ingrese nombre del alumno: ").upper()
    telefono = input("Ingrese telefono del alumno (########): ")
    hora = input("Ingrese hora de la tutoria (ejemplo: 11:00 am - 12:00 pm): ")
    dias = (
        input("Ingrese dias de la tutoria (separados por comas, ej: lunes, martes): ")
        .lower()
        .split(",")
    )
    dias = [dia.strip() for dia in dias]

    fecha_inicio = input("Ingrese fecha de inicio (YYYY-MM-DD): ")
    fecha_fin = input("Ingrese fecha de fin (YYYY-MM-DD): ")

    SolicitudesLICINTHIA = generar_solicitudes(
        tutor["Tutor"],
        clase_seleccionada,
        alumno,
        telefono,
        hora,
        dias,
        fecha_inicio,
        fecha_fin,
    )
    for solicitud in SolicitudesLICINTHIA:
        SolicitarTutoria([solicitud])


def generarTutoriaManual():
    for i, tutor_data in enumerate(Tutoresdata):
        if tutor_data["Habilitado/Deshabilitado"] == "No":
            continue
        print(f"{i} - {tutor_data['Tutor']}\n")

    tutor_index = int(input("Ingrese opcion de tutor (#): "))
    if not (0 <= tutor_index < len(Tutoresdata)):
        print("Opción de tutor no válida.")
        return

    tutor = Tutoresdata[tutor_index]
    raw_clases = tutor["Clases que Imparte"].split(";#")

    available_classes = []
    for i, item in enumerate(raw_clases):
        if i % 2 != 0 and item:
            available_classes.append(item)

    if not available_classes:
        print(f"El tutor {tutor['Tutor']} no tiene clases disponibles para impartir.")
        return

    print("\nClases disponibles:")
    for i, class_name in enumerate(available_classes):
        print(f"{i} - {class_name}")

    clase_index = int(input("Ingrese opcion de clase (#): "))
    if not (0 <= clase_index < len(available_classes)):
        print("Opción de clase no válida.")
        return

    clase_seleccionada = available_classes[clase_index]

    alumno = input("Ingrese el numero de cuenta del alumno: ")
    telefono = input("Ingrese telefono del alumno (########): ")
    hora = input("Ingrese hora de la tutoria (ejemplo: 11:00 am - 12:00 pm): ")
    fecha = input("Ingrese fecha de la tutoria (YYYY-MM-DD): ")
    temas = input("Ingrese temas de la tutoria: ").upper()

    tutoria = {
        "Tipo de Tutoria": "Presencial",
        "Contactado": "No",
        "Estado": "Pendiente",
        "Telefono": telefono,
        "Nombre Tutor": tutor["Tutor"],
        "Fecha de Tutoria": fecha,
        "Hora Tutoria": hora,
        "Clases": clase_seleccionada,
        "Temas": temas,
        "Numero de Cuenta": alumno,
    }
    SolicitarTutoria([tutoria])


def atenderTutorias():
    for tutoria in Tutoriasdata:
        if tutoria["Estado"] == "Solicitada" or tutoria["Estado"] == "Rechazada":
            if posibleUnion(tutoria):
                update_data = [
                    {
                        "ID": tutoria["ID"],
                        "Estado": "Coordinada",
                        "Nombre Tutor": tutoria["Nombre Tutor"],
                        "Tutor Asignado": tutoria["Nombre Tutor"],
                        "Tutor": tutoria["Nombre Tutor"],
                        "Aula": tutoria["Aula"],
                        "Contactado": "No",
                    }
                ]
                sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
                TutoriasPendientesRechazadas.append(tutoria)

                print(f"Tutoria ID {tutoria['ID']} coordinada")
            elif asignar_tutor(tutoria):

                update_data = [
                    {
                        "ID": tutoria["ID"],
                        "Nombre Tutor": tutoria["Nombre Tutor"],
                        "Tutor Asignado": "",
                        "Tutor": tutoria["Nombre Tutor"],
                        "Estado": "Pendiente",
                        "Contactado": "No",
                    }
                ]
                sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

                TutoriasPendientesRechazadas.append(tutoria)
                print(
                    f"Tutoria ID {tutoria['ID']} de clase {tutoria['Clases']} asignada con {tutoria['Nombre Tutor']}"
                )
            else:
                update_data = [
                    
                    {"ID": tutoria["ID"], "Estado": "Sin Tutor", "Contactado": "No"}
                ]
                sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
                print(f"(-)  No se encontró tutor")
        elif tutoria["Estado"] == "Confirmada":
            if asignar_aula(tutoria):

                update_data = [
                    {
                        "ID": tutoria["ID"],
                        "Aula": tutoria["Aula"],
                        "Estado": "Coordinada",
                        "Contactado": "No",
                    }
                ]
                sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
                print(
                    f"Aula asignada para tutoría ID {tutoria['ID']}: {tutoria['Aula']}"
                )
                # Hacer el append de la tutoría actualizada
                tutoria_actualizada = (
                    tutoria.copy()
                )  # Hacer una copia para evitar modificar el original
                tutoria_actualizada["Estado"] = (
                    "Coordinada"  # Actualizar el estado en la copia
                )
                tutoria_actualizada["Aula"] = tutoria[
                    "Aula"
                ]  # Asegurarse de que el aula esté actualizada

                # Agregar la tutoría actualizada a la lista
                TutoriasCoordinadasCONAULA.append(tutoria_actualizada)

            else:
                update_data = [
                    {"ID": tutoria["ID"], "Estado": "CASO", "Contactado": "No"}
                ]
                sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
                print(f"(-)  No se encontró aula")


def buscarTutorDisponible(tutor_viejo):
    for tutor in Tutoresdata:
        if (
            tutor["Habilitado/Deshabilitado"] == "Yes"
            and tutor["Telefono"] != ""
            and tutor["TelefonoAuxiliar"] != ""
            and tutor["Nombre Tutor"] != tutor_viejo
        ):
            return tutor        
    return None


def asignarTutorATutoria(tutoriasPorActualizar):
    for tutoria, tutor_asignado in tutoriasPorActualizar:
        update_data = [
            {
                "ID": tutoria["ID"],
                "Nombre Tutor": tutor_asignado["Nombre Tutor"],
                "Tutor": tutor_asignado["Tutor"],
                "Estado": "Pendiente",
                "Contactado": "No",
            }
        ]
        sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")


def agendarConOtroTutor():
    tutoriasPorActualizar = []
    now = datetime.now()

    for tutoria in Tutoriasdata:
        if (tutoria["Estado"] == "Agendar Otro Tutor" and tutoria["Fecha de Tutoria"] < now) :
            print(f"Agendando tutoría ID {tutoria['ID']} con otro tutor...")
            tutor_asignado = buscarTutorDisponible(tutoria['Nombre Tutor'])
            if tutor_asignado:
                tutoriasPorActualizar.append((tutoria, tutor_asignado))
            else:
                print(f"No se encontró un tutor disponible para la tutoría ID {tutoria['ID']}")
    asignarTutorATutoria(tutoriasPorActualizar)


def menuOpciones():
    print("Bienvenido al menú de agregación de tutorías.")
    while True:
        print("\nSeleccione una opción:")
        print("1. Atender solicitudes de tutoría")
        print("2. Generar tutoría PAE")
        print("3. Generar tutoría manual")
        print("4. Salir")
        opcion = input("Ingrese el número de la opción: ")

        if opcion == "1":
            atenderTutorias()
        elif opcion == "2":
            generarTutoriaPAE()
        elif opcion == "3":
            generarTutoriaManual()
        elif opcion == "4":
            print("Saliendo del menú.")
            break
        else:
            print("Opción no válida, por favor intente de nuevo.")


menuOpciones()