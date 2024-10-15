import wmi
import colorama
from colorama import Fore, Style
from datetime import datetime

colorama.init(autoreset=True)

c = wmi.WMI()
eventos = c.Win32_NTLogEvent(Logfile='Security')
archivo_salida = "conexiones_eventos.txt"

for evento in eventos:
    if evento.EventCode in [4624, 4672]:
        mensaje = evento.InsertionStrings
        fecha_hora = datetime.strptime(evento.TimeGenerated.split('.')[0], "%Y%m%d%H%M%S")

        # Extraer información de manera segura
        usuario = mensaje[5] if len(mensaje) > 5 else "Desconocido"
        dominio = mensaje[6] if len(mensaje) > 6 else "Desconocido"
        ip_remota = mensaje[18] if len(mensaje) > 18 else "-"
        estado = "Éxito" if evento.EventCode == 4624 else "Con Privilegios"

        # Determinar el tipo de inicio de sesión
        tipo_inicio_sesion = mensaje[8] if len(mensaje) > 8 else "Desconocido"

        if tipo_inicio_sesion == "2":
            protocolo = "Consola"
            color = Fore.YELLOW
        elif tipo_inicio_sesion == "3":
            protocolo = "Red"
            color = Fore.CYAN
        elif tipo_inicio_sesion == "10":
            protocolo = "RDP"
            color = Fore.GREEN
        elif tipo_inicio_sesion == "7":
            protocolo = "Desbloqueo"
            color = Fore.MAGENTA
        else:
            protocolo = f"Otro ({tipo_inicio_sesion})"
            color = Fore.WHITE

        # Mostrar resultado en consola
        print(color + f"[{fecha_hora}] Usuario: {usuario} | Dominio: {dominio} | IP: {ip_remota} | Protocolo: {protocolo} | Estado: {estado}")

        # Escribir resultado en el archivo de salida
        with open(archivo_salida, 'a') as archivo:
            archivo.write(f"{fecha_hora} | {usuario} | {dominio} | {ip_remota} | {protocolo} | {estado}\n")

print(Style.BRIGHT + Fore.YELLOW + f"\nResultados guardados en {archivo_salida}")
