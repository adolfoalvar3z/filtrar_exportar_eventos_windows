import win32evtlog
import win32con
import datetime
from win32com.client import Dispatch
from colorama import init, Fore, Style

# Inicializamos colorama para los colores en la terminal
init(autoreset=True)

# Diccionario para interpretar tipos de inicio de sesión
logon_types = {
    2: "Inicio interactivo (Consola)",
    3: "Red (Conexión SMB u otra red)",
    4: "Inicio de sesión por lotes",
    5: "Inicio de sesión de servicio",
    7: "Desbloqueo de la estación de trabajo",
    10: "Inicio remoto (RDP)",
    11: "Inicio en caché (sin acceso a DC)"
}

def export_remote_connections():
    server = None  # None para el equipo local
    log_type = 'Security'  # Consultamos el log de seguridad
    
    hand = win32evtlog.OpenEventLog(server, log_type)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    
    events = []
    
    try:
        while True:
            records = win32evtlog.ReadEventLog(hand, flags, 0)
            if not records:
                break
            
            for record in records:
                if record.EventID in [4624, 4776]:  # Eventos de inicio de sesión remoto y SMB
                    event_data = {
                        "Event ID": record.EventID,
                        "Time Generated": record.TimeGenerated.Format(),
                        "User Name": "",
                        "Domain": "",
                        "Source IP": "",
                        "Workstation Name": "",
                        "Logon Type": "",
                        "Logon Success": "Correcta"
                    }
                    
                    for data in record.StringInserts:
                        if "Account Name:" in data:
                            event_data["User Name"] = data.split(":")[-1].strip()
                        elif "Account Domain:" in data:
                            event_data["Domain"] = data.split(":")[-1].strip()
                        elif "Source Network Address:" in data:
                            event_data["Source IP"] = data.split(":")[-1].strip()
                        elif "Workstation Name:" in data:
                            event_data["Workstation Name"] = data.split(":")[-1].strip()
                        elif "Logon Type:" in data:
                            logon_type_code = int(data.split(":")[-1].strip())
                            event_data["Logon Type"] = logon_types.get(logon_type_code, f"Desconocido ({logon_type_code})")
                    
                    if record.EventID == 4776 and "Failure" in record.StringInserts[0]:
                        event_data["Logon Success"] = "Fallida"
                    
                    events.append(event_data)
    
    finally:
        win32evtlog.CloseEventLog(hand)
    
    return events

# Guardamos los eventos en un archivo y los mostramos por pantalla
events = export_remote_connections()

# Separar los eventos de conexiones remotas y SMB del resto
events_remote_smb = [e for e in events if e["Logon Type"] in ["Red (Conexión SMB u otra red)", "Inicio remoto (RDP)"]]
events_other = [e for e in events if e not in events_remote_smb]

# Mostrar solo las conexiones remotas y SMB en pantalla con colores distintos
for event in events_remote_smb:
    color = Fore.GREEN if event["Logon Type"] == "Inicio remoto (RDP)" else Fore.YELLOW
    formatted_event = (
        f"{color}Fecha: {event['Time Generated']}\n"
        f"Usuario: {event['User Name']}\n"
        f"Dominio: {event['Domain']}\n"
        f"IP de origen: {event['Source IP']}\n"
        f"Nombre de estación de trabajo: {event['Workstation Name']}\n"
        f"Tipo de conexión: {event['Logon Type']}\n"
        f"Resultado de la conexión: {event['Logon Success']}\n"
        f"ID del evento: {event['Event ID']} (Tipo de conexión: {event['Logon Type']})
"
        f"----------------------------------------"
    )
    print(formatted_event)  # Mostrar por pantalla

# Formateo de los eventos para la exportación
formatted_events = []
for event in events_remote_smb + events_other:
    formatted_event = (
        f"Fecha: {event['Time Generated']}\n"
        f"Usuario: {event['User Name']}\n"
        f"Dominio: {event['Domain']}\n"
        f"IP de origen: {event['Source IP']}\n"
        f"Nombre de estación de trabajo: {event['Workstation Name']}\n"
        f"Tipo de conexión: {event['Logon Type']}\n"
        f"Resultado de la conexión: {event['Logon Success']}\n"
        f"ID del evento: {event['Event ID']} (Tipo de conexión: {event['Logon Type']})
"
        "----------------------------------------"
    )
    formatted_events.append(formatted_event)

# Guardar los eventos en un archivo de texto
with open("remote_connections_log.txt", "w", encoding="utf-8") as f:
    f.write("\n\n".join(formatted_events))

print(f"\nSe encontraron {len(events)} eventos. Exportados a 'remote_connections_log.txt'")