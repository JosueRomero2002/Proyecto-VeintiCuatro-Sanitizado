import pywhatkit
from datetime import datetime, timedelta
import time
import pyautogui
import keyboard as k
import os
from dotenv import load_dotenv
from shareplum import Site
from shareplum import Office365

load_dotenv()

userSh = os.getenv("SHAREPOINT_USER")
passSh = os.getenv("SHAREPOINT_PASS")


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
        "Tipo de Tutoria",
        "Fecha de Tutoria",
        "Hora Tutoria",
        "Clases",
        "Temas",
        "Alumnos",
        "Aula",
        "Numero de Cuenta",
        "HoraClasica",
        "ClaseClasica",
    ]
)
Tutoresdata = sp_list_Tutores.GetListItems(
    fields=["ID", "Tutor", "Telefono", "TelefonoAuxiliar", "Número de Cuenta"]
)
Aulasdata = sp_list_Aulas.GetListItems(fields=["ID", "IdAula ", "Oficial"])

import locale


locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")  # Configura la localización en español

print("Tutores Data")


print(Tutoriasdata)

print("--------------------")

print(Tutoresdata)
print("--------------------")
print("Aulas Data")
print(Aulasdata)


print(sp_list_Tutores)
print("--------------------")
print(Tutoresdata)


print(len(Tutoriasdata))


ContactarPendientes = True
ContactarCoordinadas = True
ContactarRechazadas = False
ContactarReagendadas = False
ContactarCanceladas = False

ContactarModeradorAula = False
ContactarModeradorTutor = True  # contactar sin tutor


FilteredTutoriasdata = Tutoriasdata


auxSize = 0


size = len(FilteredTutoriasdata)
print(Tutoresdata[0])


def obtenerNumerCuentaTutor(tutor):
    for j in range(0, len(Tutoresdata)):
        if Tutoresdata[j]["Tutor"] == tutor:
            return Tutoresdata[j]["Número de Cuenta"]
    return "00000000"


# def obtenerHorariosDisponiblesDeClaseEnTutor(clase):


if ContactarPendientes:
    for i in range(0, size):

        now = datetime.now()

        # print(FilteredTutoriasdata[i] )
        # print(FilteredTutoriasdata[i]['Estado'] )
        if (
            FilteredTutoriasdata[i]["Estado"]
            and FilteredTutoriasdata[i]["Estado"] == "Pendiente"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
            # and FilteredTutoriasdata[i]['Fecha de Tutoria'] >= now:
            #  if encuestado[i] != "Aplicada":

            current_hour = int(now.strftime("%H"))

            if current_hour == 24:

                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:

                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""

            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        try:
                            TelefonoTutor = Tutoresdata[j]["TelefonoAuxiliar"]

                        except:
                            TelefonoTutor = "87794832"

            # num = str(FilteredTutoriasdata[i]['Telefono'])
            numAux = "+504" + TelefonoTutor
            # TelefonoTutor

            print("Mensaje enviado a: ", numAux)
            #   print(i," ", nombres[i],"\t",numAux)

            # ✨🐯PROPUESTA de Tutoria🐯✨

            # ➡️ Modalidad : Presencial
            # ➡️ Fecha: 26 enero 2024
            # ➡️ Hora: 9:00 am - 10:00 am
            # ➡️ Asignatura: Ecuaciones Diferenciales
            # ➡️ Alumno: ELIAS JOSUE BONILLA MENDEZ
            # ➡️Tema: Separables y Sustitución
            # ➡️ Tutor: ERICK EDUARDO ARITA HENRIQUEZ

            clase = None
            if "Clases" in FilteredTutoriasdata[i]:
                clase = FilteredTutoriasdata[i]["Clases"]
            else:
                clase = FilteredTutoriasdata[i]["ClaseClasica"]
            hora = None
            if "Hora Tutoria" in FilteredTutoriasdata[i]:
                hora = FilteredTutoriasdata[i]["Hora Tutoria"]
            else:
                hora = FilteredTutoriasdata[i]["HoraClasica"]

            mensaje = "✨🐯PROPUESTA de Tutoria🐯✨"
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + FilteredTutoriasdata[i]["Nombre Tutor"]

            mensaje += "\n\nAprobada Responder con => 👍"
            mensaje += "\nRechazada Responder con => 👎"
            mensaje += "\n\n(Si confirmas por medio de este mensaje, no es necesario que respondas el correo)"
            mensaje += "\nBeta Version 1.1"
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"
            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)

            pyautogui.click(800, 450)
            time.sleep(2)

            k.press_and_release("enter")

            time.sleep(2)

            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, True, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        #   pyautogui.click(1050, 950)

        #   time.sleep(2)
        #   k.press_and_release('enter')
        #   excel_data_df.loc[i,['Encuesta']] = "Aplicada"

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")

time.sleep(2)


if ContactarCoordinadas:
    for i in range(0, size):

        now = datetime.now()
        now = now.replace(hour=0, minute=0, second=0, microsecond=0)
        fecha_str = FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%Y-%m-%d")
        fecha = datetime.strptime(fecha_str, "%Y-%m-%d")

        if (
            FilteredTutoriasdata[i]["Estado"] == "Coordinada"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] >= now
            and FilteredTutoriasdata[i]["Temas"] != "PAE"
        ):

            if (
                FilteredTutoriasdata[i]["Nombre Tutor"]
                == "CARDENAS DELCID CYNTHIA STEPHANIE"
                and FilteredTutoriasdata[i]["Fecha de Tutoria"] != now
            ):
                # Si es Cynthia y no es el dia de la tutoria se omite
                print("Omitiendo Lic Cynthia")
                continue

            current_hour = int(now.strftime("%H"))

            if current_hour == 24:
                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]

                    except:
                        TelefonoTutor = "87794832"
                    try:

                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = "87794832"

            numAux = "+504" + TelefonoTutor

            print("Mensaje enviado a: ", numAux)

            AulaTutoria = "Zoom"

            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]
            if FilteredTutoriasdata[i][
                "Nombre Tutor"
            ] == "DIEGO ANDRES RIVERA VALLE" and (
                FilteredTutoriasdata[i]["Alumnos"] == "CLAUDIA MARYSOL GRADIZ ECHEVERRY"
                or FilteredTutoriasdata[i]["Alumnos"] == "DIEGO ANDRES RIVERA VALLE"
            ):
                NombreTutor = "DANIELA LARRISA PINEDA CASTRO"
                TelefonoTutor = "92064537"

            if FilteredTutoriasdata[i]["Tipo de Tutoria"] == "Presencial":
                AulaTutoria = FilteredTutoriasdata[i]["Aula"]

            mensaje = "✅✅✅🐯RECORDATORIO de Tutoria🐯✅✅✅"
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + FilteredTutoriasdata[i]["Hora Tutoria"]

            mensaje += "\n➡️ Asignatura: " + FilteredTutoriasdata[i]["Clases"]

            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + NombreTutor

            mensaje += "\n➡️ Contacto: " + TelefonoTutor

            mensaje += "\n➡️ Aula: " + AulaTutoria
            #  == "Virtual"+ FilteredTutoriasdata[i]['Aula'] if FilteredTutoriasdata[i]['Aula'] else "Zoom" #FilteredTutoriasdata[i]['Tipo de Tutoria'] == "Virtual" if
            mensaje += "\n\n✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅"
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)

            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)

            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)

            numAux = "+504" + TelefonoAlumno
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")

            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        if (
            FilteredTutoriasdata[i]["Temas"] == "PAE"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and fecha == now
        ):

            print("Tutoria PAE")
            print(FilteredTutoriasdata[i]["Fecha de Tutoria"])
            print(FilteredTutoriasdata[i]["Nombre Tutor"])

            current_hour = int(now.strftime("%H"))

            if current_hour == 24:

                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]

                    except:
                        TelefonoTutor = "87794832"
                    try:

                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = "87794832"

            numAux = "+504" + TelefonoTutor

            print("Mensaje enviado a: ", numAux)

            AulaTutoria = "Zoom"

            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]
            if FilteredTutoriasdata[i][
                "Nombre Tutor"
            ] == "DIEGO ANDRES RIVERA VALLE" and (
                FilteredTutoriasdata[i]["Alumnos"] == "CLAUDIA MARYSOL GRADIZ ECHEVERRY"
                or FilteredTutoriasdata[i]["Alumnos"] == "DIEGO ANDRES RIVERA VALLE"
            ):
                NombreTutor = "DANIELA LARRISA PINEDA CASTRO"
                TelefonoTutor = "92064537"

            if FilteredTutoriasdata[i]["Tipo de Tutoria"] == "Presencial":
                try:
                    AulaTutoria = FilteredTutoriasdata[i]["Aula"]
                except:
                    AulaTutoria = "Zoom"

            mensaje = "✅✅✅🐯RECORDATORIO de Tutoria🐯✅✅✅"
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + FilteredTutoriasdata[i]["Hora Tutoria"]

            mensaje += "\n➡️ Asignatura: " + FilteredTutoriasdata[i]["Clases"]

            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + NombreTutor

            mensaje += "\n➡️ Contacto: " + TelefonoTutor

            mensaje += "\n➡️ Aula: " + AulaTutoria
            #  == "Virtual"+ FilteredTutoriasdata[i]['Aula'] if FilteredTutoriasdata[i]['Aula'] else "Zoom" #FilteredTutoriasdata[i]['Tipo de Tutoria'] == "Virtual" if
            mensaje += "\n\n✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅"
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)

            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)

            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)

            numAux = "+504" + TelefonoAlumno
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")

            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")

time.sleep(2)

if ContactarRechazadas:

    for i in range(0, size):
        now = datetime.now()
        if (
            (FilteredTutoriasdata[i]["Estado"] == "Rechazada")
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] < now
        ):
            current_hour = int(now.strftime("%H"))
            if current_hour == 24:

                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoTutor = "87794832"
            numAux = "+504" + TelefonoAlumno

            print("Mensaje enviado a: ", numAux)

            mensaje = "🐯Tutorias Unitec🐯"
            mensaje += "\n\nSaludos de Tutorias Unitec, espero que se encuentre bien. Le escribo por la tutoria que habia solicitado, la cual no pudimos agendar con un tutor a tiempo por falta de disponibilidad. Si desea reagendar o cancelar la tutoria, por favor responder a este mensaje."
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + FilteredTutoriasdata[i]["Nombre Tutor"]
            mensaje += "\n\nDesea reagendar para esta semana => 👍"
            mensaje += "\nPrefiere ya no recibir la tutoria => 👎"

            mensaje += "\nTambien puedes responder a este mensaje con que otro horario en otro dia podrias recibir la tutoria"

            #  mensaje += "\n\n(Si confirmas por medio de este mensaje, no es necesario que respondas el correo)"
            mensaje += "\n\nBeta Version 1.0"
            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        #   pyautogui.click(1050, 950)

        #   time.sleep(2)
        #   k.press_and_release('enter')
        #   excel_data_df.loc[i,['Encuesta']] = "Aplicada"

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")
time.sleep(2)


if ContactarModeradorAula:

    for i in range(0, size):
        now = datetime.now()
        if (
            FilteredTutoriasdata[i]["Estado"] == "CASO"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
            current_hour = int(now.strftime("%H"))
            if current_hour == 24:
                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoTutor = "87794832"
            numAux = "+504" + TelefonoAlumno

            print("Mensaje enviado a: ", numAux)

            clase = None
            if "Clases" in FilteredTutoriasdata[i]:
                clase = FilteredTutoriasdata[i]["Clases"]
            else:
                clase = FilteredTutoriasdata[i]["ClaseClasica"]
            hora = None
            if "Hora Tutoria" in FilteredTutoriasdata[i]:
                hora = FilteredTutoriasdata[i]["Hora Tutoria"]
            else:
                hora = FilteredTutoriasdata[i]["HoraClasica"]

            mensaje = "✨🐯Solicitud de AULA de Tutorias Unitec🐯✨"
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            #  mensaje += "\n➡️ Numero de Cuenta: "+FilteredTutoriasdata[i]['Numero de Cuenta']
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno

            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + FilteredTutoriasdata[i]["Nombre Tutor"]
            #  mensaje += "\n➡️ Numero de Cuenta Tutor: "+obtenerNumerCuentaTutor(FilteredTutoriasdata[i]['Nombre Tutor'])
            mensaje += "\nBeta Version 1.0"
            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly("+50489886363", mensaje, 25, False, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        #   pyautogui.click(1050, 950)

        #   time.sleep(2)
        #   k.press_and_release('enter')
        #   excel_data_df.loc[i,['Encuesta']] = "Aplicada"


if ContactarReagendadas:
    for i in range(0, size):

        now = datetime.now()
        now = now.replace(hour=0, minute=0, second=0, microsecond=0)
        fecha_str = FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%Y-%m-%d")
        fecha = datetime.strptime(fecha_str, "%Y-%m-%d")

        if (
            FilteredTutoriasdata[i]["Estado"] == "Reagendar"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] >= now
            and FilteredTutoriasdata[i]["Temas"] != "PAE"
        ):

            if (
                FilteredTutoriasdata[i]["Nombre Tutor"]
                == "CARDENAS DELCID CYNTHIA STEPHANIE"
                and FilteredTutoriasdata[i]["Fecha de Tutoria"] != now
            ):
                # Si es Cynthia y no es el dia de la tutoria se omite
                print("Omitiendo Lic Cynthia")
                continue

            current_hour = int(now.strftime("%H"))

            if current_hour == 24:
                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]

                    except:
                        TelefonoTutor = "87794832"
                    try:

                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = "87794832"

            numAux = "+504" + TelefonoTutor

            print("Mensaje enviado a: ", numAux)

            AulaTutoria = "Zoom"

            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]
            if FilteredTutoriasdata[i][
                "Nombre Tutor"
            ] == "DIEGO ANDRES RIVERA VALLE" and (
                FilteredTutoriasdata[i]["Alumnos"] == "CLAUDIA MARYSOL GRADIZ ECHEVERRY"
                or FilteredTutoriasdata[i]["Alumnos"] == "DIEGO ANDRES RIVERA VALLE"
            ):
                NombreTutor = "DANIELA LARRISA PINEDA CASTRO"
                TelefonoTutor = "92064537"

            if FilteredTutoriasdata[i]["Tipo de Tutoria"] == "Presencial":
                AulaTutoria = FilteredTutoriasdata[i]["Aula"]

            clase = None
            if "Clases" in FilteredTutoriasdata[i]:
                clase = FilteredTutoriasdata[i]["Clases"]
            else:
                clase = FilteredTutoriasdata[i]["ClaseClasica"]
            hora = None
            if "Hora Tutoria" in FilteredTutoriasdata[i]:
                hora = FilteredTutoriasdata[i]["Hora Tutoria"]
            else:
                hora = FilteredTutoriasdata[i]["HoraClasica"]

            mensaje = "📅🔄 *Solicitud de Reagendamiento de Tutoría* 🔄📅"
            mensaje += "\n\n➡️ Modalidad: " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + NombreTutor
            mensaje += "\n➡️ Contacto: " + TelefonoTutor
            mensaje += "\n➡️ Aula: " + AulaTutoria

            mensaje += "\n\n⚠️ ¿Confirmas reagendar la tutoría? ⚠️"
            mensaje += "\n\n✅ Aprobada: Responder con => 👍"
            mensaje += "\n❌ Rechazada: Responder con => 👎"

            mensaje += "\n\n🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄🔄"
            mensaje += "\n\n📌 Ver Mis Tutorías: https://heylink.me/josue546/"

            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)

            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)

            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)

            numAux = "+504" + TelefonoAlumno
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")

            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")
time.sleep(2)


if ContactarModeradorTutor:

    for i in range(0, size):
        now = datetime.now()
        if (
            FilteredTutoriasdata[i]["Estado"] == "Sin Tutor"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
            current_hour = int(now.strftime("%H"))
            if current_hour == 24:
                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            #  for j in range (0,len(Tutoresdata)):

            #       if Tutoresdata[j]['Tutor'] == FilteredTutoriasdata[i]['Nombre Tutor']:

            #          try:
            #           TelefonoTutor = Tutoresdata[j]['Telefono']
            TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
            #          except:
            #            TelefonoTutor = "87794832"
            numAux = "+504" + TelefonoAlumno

            print("Mensaje enviado a: ", numAux)

            #  mensaje = "✨🐯Solicitud de TUTOR de Tutorias Unitec🐯✨"
            #  mensaje += "\n\n➡️ Modalidad : "+FilteredTutoriasdata[i]['Tipo de Tutoria']
            #  mensaje += "\n➡️ Fecha: "+FilteredTutoriasdata[i]['Fecha de Tutoria'].strftime("%Y-%m-%d %H:%M:%S")
            #  mensaje += "\n➡️ Hora: "+FilteredTutoriasdata[i]['Hora Tutoria']
            #  mensaje += "\n➡️ Asignatura: "+FilteredTutoriasdata[i]['Clases']
            #  mensaje += "\n➡️ Alumno: "+FilteredTutoriasdata[i]['Alumnos']
            #  mensaje += "\n➡️ Numero de Cuenta: "+FilteredTutoriasdata[i]['Numero de Cuenta']
            #  mensaje += "\n➡️ Contacto: "+TelefonoAlumno
            #  mensaje += "\n➡️ Tema: "+FilteredTutoriasdata[i]['Temas']
            #  mensaje += "\nBeta Version 1.0"
            #  #pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)
            #  pyautogui.click(800, 450)
            #  time.sleep(2)
            #  k.press_and_release('enter')
            #  time.sleep(2)
            #  pywhatkit.sendwhatmsg_instantly("+50489886363", mensaje, 25, False, 4)
            #  update_data = [{'ID': FilteredTutoriasdata[i]['ID'], 'Contactado': 'Yes'}]
            #  sp_list_Tutorias.UpdateListItems(data=update_data, kind='Update')
            clase = None
            if "Clases" in FilteredTutoriasdata[i]:
                clase = FilteredTutoriasdata[i]["Clases"]
            else:
                clase = FilteredTutoriasdata[i]["ClaseClasica"]
            hora = None
            if "Hora Tutoria" in FilteredTutoriasdata[i]:
                hora = FilteredTutoriasdata[i]["Hora Tutoria"]
            else:
                hora = FilteredTutoriasdata[i]["HoraClasica"]

            mensaje = "🐯Tutorias Unitec🐯 \n📕📕📕📕📕📕📕📕📕📕📕📕📕📕"
            mensaje += "\n\nEsperamos que estés bien. Queríamos informarte que, lamentablemente, *no pudimos coordinar tu solicitud de tutoría porque no había disponibilidad para ese día u horario 📅.* Esto sucede debido a la alta demanda de tutorías."
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]

            mensaje += "\n\nTe sugerimos intentar solicitar nuevamente 🔄 y revisar el catálogo de tutores en @tutorias_unitecsps en Instagram 📲, donde puedes ver la oferta disponible. ¡Esperamos poder ayudarte pronto! ✨"
            #  mensaje += "\nPrefiere ya no recibir la tutoria => 👎"

            mensaje += "\nTambien puedes responder a este mensaje con que otro horario en otro dia podrias recibir la tutoria"

            #  mensaje += "\n\n(Si confirmas por medio de este mensaje, no es necesario que respondas el correo)"
            mensaje += "\n\nBeta Version 1.0\n📕📕📕📕📕📕📕📕📕📕📕📕📕📕"
            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        #   pyautogui.click(1050, 950)

        #   time.sleep(2)
        #   k.press_and_release('enter')
        #   excel_data_df.loc[i,['Encuesta']] = "Aplicada"

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")
time.sleep(2)


if ContactarCanceladas:

    for i in range(0, size):
        now = datetime.now()
        if (
            (FilteredTutoriasdata[i]["Estado"] == "Cancelada")
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] < now
        ):
            current_hour = int(now.strftime("%H"))
            if current_hour == 24:

                current_hour = 0

            current_minute = int(now.strftime("%M")) + 1

            if current_minute == 60:
                current_minute = 0

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):

                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:

                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoTutor = "87794832"
            numAux = "+504" + TelefonoAlumno

            print("Mensaje enviado a: ", numAux)
            clase = None
            if "Clases" in FilteredTutoriasdata[i]:
                clase = FilteredTutoriasdata[i]["Clases"]
            else:
                clase = FilteredTutoriasdata[i]["ClaseClasica"]
            hora = None
            if "Hora Tutoria" in FilteredTutoriasdata[i]:
                hora = FilteredTutoriasdata[i]["Hora Tutoria"]
            else:
                hora = FilteredTutoriasdata[i]["HoraClasica"]

            mensaje = "🐯Tutorias Unitec🐯"
            mensaje += "\n\nSaludos de Tutorias Unitec, espero que se encuentre bien. Le escribo por la tutoria que habia solicitado, ."
            mensaje += "\n\n➡️ Modalidad : " + FilteredTutoriasdata[i]["Tipo de Tutoria"]
            mensaje += "\n➡️ Fecha: " + FilteredTutoriasdata[i][
                "Fecha de Tutoria"
            ].strftime("%Y-%m-%d %H:%M:%S")
            mensaje += (
                "\n➡️ Día: "
                + FilteredTutoriasdata[i]["Fecha de Tutoria"].strftime("%A").upper()
            )
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + FilteredTutoriasdata[i]["Nombre Tutor"]
            mensaje += "\n\nDesea reagendar para esta semana => 👍"
            mensaje += "\nPrefiere ya no recibir la tutoria => 👎"

            mensaje += "\nTambien puedes responder a este mensaje con que otro horario en otro dia podrias recibir la tutoria"

            #  mensaje += "\n\n(Si confirmas por medio de este mensaje, no es necesario que respondas el correo)"
            mensaje += "\n\nBeta Version 1.0"
            # pywhatkit.sendwhatmsg(numAux,mensaje,current_hour,current_minute,40,True,50)
            pyautogui.click(800, 450)
            time.sleep(2)
            k.press_and_release("enter")
            time.sleep(2)
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        #   pyautogui.click(1050, 950)

        #   time.sleep(2)
        #   k.press_and_release('enter')
        #   excel_data_df.loc[i,['Encuesta']] = "Aplicada"

pyautogui.click(800, 450)
time.sleep(2)
k.press_and_release("enter")
time.sleep(2)

# telefonoCoordinadora = "+50489886363"

# pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
