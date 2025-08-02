import pywhatkit
from datetime import datetime
import time
import pyautogui
import keyboard as k
import os
from dotenv import load_dotenv
from shareplum import Site
from shareplum import Office365
import locale

load_dotenv()
locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")

userSh = os.getenv("SHAREPOINT_USER")
passSh = os.getenv("SHAREPOINT_PASS")
auxPhone = os.getenv("AUX_PHONE")

# Autenticación en SharePoint
authcookie = Office365(
    "https://unitechn.sharepoint.com", username=userSh, password=passSh
).GetCookies()
site = Site(
    "https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/", authcookie=authcookie
)

# Listas de SharePoint
sp_list_Tutorias = site.List("Tutorias")
sp_list_Tutores = site.List("Tutores")
sp_list_Aulas = site.List("Aulas")

# Obtener datos de las listas
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
contactados = {
"ContactarPendientes": False,
"ContactarCoordinadas": False,
"ContactarRechazadas": False,
"ContactarReagendadas": False,
"ContactarCanceladas": False,
"ContactarModeradorAula": False,
"ContactarSinTutor": False
}

FilteredTutoriasdata = Tutoriasdata
size = len(FilteredTutoriasdata)


def obtenerNumerCuentaTutor(tutor):
    for j in range(0, len(Tutoresdata)):
        if Tutoresdata[j]["Tutor"] == tutor:
            return Tutoresdata[j]["Número de Cuenta"]
    return "00000000"


def autoGuiSkip():
    pyautogui.click(800, 450)
    time.sleep(2)
    k.press_and_release("enter")
    time.sleep(2)


def contactarPendientes():
    for i in range(0, size):
        if (
            FilteredTutoriasdata[i]["Estado"]
            and FilteredTutoriasdata[i]["Estado"] == "Pendiente"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
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
                            TelefonoTutor = auxPhone

            numAux = "+504" + TelefonoTutor
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
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, True, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarPendientes"] = True


def contactarCoordinadas():
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
            if (
                FilteredTutoriasdata[i]["Nombre Tutor"]
                == "CARDENAS DELCID CYNTHIA STEPHANIE"
                and FilteredTutoriasdata[i]["Fecha de Tutoria"] != now
            ):
                # Si es Cynthia y no es el dia de la tutoria se omite
                continue

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                    except:
                        TelefonoTutor = auxPhone
                    try:
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = auxPhone

            numAux = "+504" + TelefonoTutor
            print("Mensaje enviado a: ", numAux)
            AulaTutoria = "Zoom"
            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]

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
            mensaje += "\n➡️ Hora: " + hora
            mensaje += "\n➡️ Asignatura: " + clase
            mensaje += "\n➡️ Alumno: " + FilteredTutoriasdata[i]["Alumnos"]
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + NombreTutor
            mensaje += "\n➡️ Contacto: " + TelefonoTutor
            mensaje += "\n➡️ Aula: " + AulaTutoria
            mensaje += "\n\n✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅"
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)

            autoGuiSkip()
            numAux = "+504" + TelefonoAlumno
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")

        if (
            FilteredTutoriasdata[i]["Temas"] == "PAE"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and fecha == now
        ):
            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                    except:
                        TelefonoTutor = auxPhone
                    try:
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = auxPhone

            numAux = "+504" + TelefonoTutor
            print("Mensaje enviado a: ", numAux)
            AulaTutoria = "Zoom"
            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]

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
            mensaje += "\n\n✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅"
            mensaje += "\n\nVer Mis Tutorias: https://heylink.me/josue546/"

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)

            numAux = "+504" + TelefonoAlumno
            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarCoordinadas"] = True


def contactarRechazadas():
    now = datetime.now()
    for i in range(0, size):
        if (
            (FilteredTutoriasdata[i]["Estado"] == "Rechazada")
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] < now
        ):
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

            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]

            numAux = "+504" + TelefonoAlumno
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

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarRechazadas"] = True


def contactarModeradorAula():
    for i in range(0, size):
        if (
            FilteredTutoriasdata[i]["Estado"] == "CASO"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]

            numAux = "+504" + TelefonoAlumno
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
            mensaje += "\n➡️ Contacto: " + TelefonoAlumno
            mensaje += "\n➡️ Tema: " + FilteredTutoriasdata[i]["Temas"]
            mensaje += "\n➡️ Tutor: " + FilteredTutoriasdata[i]["Nombre Tutor"]

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly("+50489886363", mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarCoordinadas"] = True


def contactarReagendadas():
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
                continue

            TelefonoTutor = ""
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    try:
                        TelefonoTutor = Tutoresdata[j]["Telefono"]
                    except:
                        TelefonoTutor = auxPhone
                    try:
                        TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
                    except:
                        TelefonoAlumno = auxPhone

            numAux = "+504" + TelefonoTutor
            AulaTutoria = "Zoom"
            NombreTutor = FilteredTutoriasdata[i]["Nombre Tutor"]

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

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)

            numAux = "+504" + TelefonoAlumno
            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 15, False, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [{"ID": FilteredTutoriasdata[i]["ID"], "Contactado": "Yes"}]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarReagendadas"] = True


def contactarSinTutor():
    for i in range(0, size):
        if (
            FilteredTutoriasdata[i]["Estado"] == "Sin Tutor"
            and FilteredTutoriasdata[i]["Contactado"] == "No"
        ):
            TelefonoAlumno = ""
            TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
            numAux = "+504" + TelefonoAlumno
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
            mensaje += "\nTambien puedes responder a este mensaje con que otro horario en otro dia podrias recibir la tutoria\n📕📕📕📕📕📕📕📕📕📕📕📕📕📕"

            autoGuiSkip()
            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarReagendadas"] = True


def contactarCanceladas():
    now = datetime.now()
    for i in range(0, size):
        if (
            (FilteredTutoriasdata[i]["Estado"] == "Cancelada")
            and FilteredTutoriasdata[i]["Contactado"] == "No"
            and FilteredTutoriasdata[i]["Fecha de Tutoria"] < now
        ):
            TelefonoAlumno = ""
            for j in range(0, len(Tutoresdata)):
                if Tutoresdata[j]["Tutor"] == FilteredTutoriasdata[i]["Nombre Tutor"]:
                    TelefonoAlumno = FilteredTutoriasdata[i]["Telefono"]
            
            numAux = "+504" + TelefonoAlumno
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

            pywhatkit.sendwhatmsg_instantly(numAux, mensaje, 25, False, 4)
            print("Mensaje enviado a: ", numAux)
            autoGuiSkip()
            update_data = [
                {
                    "ID": FilteredTutoriasdata[i]["ID"],
                    "Contactado": "Yes",
                    "Estado": "No se impartio",
                }
            ]
            sp_list_Tutorias.UpdateListItems(data=update_data, kind="Update")
            contactados["ContactarCanceladas"] = True


def menuContactar():
    while True:
        print("\n🐯 Menú de Contacto de Tutorías 🐯")
        print("1. Contactar Tutorias Pendientes")
        print("2. Contactar Tutorias Coordinadas")
        print("3. Contactar Tutorias Rechazadas")
        print("4. Contactar Moderador de Aula")
        print("5. Contactar Tutorias Reagendadas")
        print("6. Contactar Tutorias Sin Tutor")
        print("7. Contactar Tutorias Canceladas")
        print("8. Salir")

        opcion = input("Seleccione una opción: ")
        if opcion == "1":
            if contactados["ContactarPendientes"]:
                print("Ya se han contactado las tutorias pendientes.")
                return
            print("Contactando Tutorias Pendientes...")
            contactarPendientes()
        elif opcion == "2":
            if contactados["ContactarCoordinadas"]:
                print("Ya se han contactado las tutorias coordinadas.")
                return
            print("Contactando Tutorias Coordinadas...")
            contactarCoordinadas()
        elif opcion == "3":
            if contactados["ContactarRechazadas"]:
                print("Ya se han contactado las tutorias rechazadas.")
                return
            print("Contactando Tutorias Rechazadas...")
            contactarRechazadas()
        elif opcion == "4":
            if contactados["ContactarReagendadas"]:
                print("Ya se han contactado las tutorias reagendadas.")
                return
            print("Contactando Tutorias Reagendadas...")
            contactarReagendadas()
        elif opcion == "5":
            if contactados["ContactarModeradorAula"]:
                print("Ya se ha contactado al moderador de aula.")
                return
            print("Contactando Moderador de Aula...")
            contactarModeradorAula()
        elif opcion == "6":
            if contactados["ContactarSinTutor"]:
                print("Ya se han contactado las tutorias sin tutor.")
                return
            print("Contactando Moderador Tutor...")
            contactarSinTutor()
        elif opcion == "7":
            if contactados["ContactarCanceladas"]:
                print("Ya se han contactado las tutorias canceladas.")
                return
            print("Contactando Tutorias Canceladas...")
            contactarCanceladas()
        elif opcion == "8":
            print("🐯 Saliendo del programa 🐯")
            break
        else:
            print("Opción no válida. Intente nuevamente.")


menuContactar()
