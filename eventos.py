import win32evtlog
from win32com.client import Dispatch

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
    
    # Abrimos el log de eventos de seguridad
    hand = win32evtlog.OpenEventLog(server, log_type)
    
    flags = win32evtlog.EVENTLOG_FORWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    
    total_events = win32evtlog.GetNumberOfEventLogRecords(hand)
    
    events = []
    
    while True:
        records = win32evtlog.ReadEventLog(hand, flags, 0)
        
        if not records:
            break
        
        for record in records:
            # Filtramos eventos relacionados con conexión remota o acceso SMB (ID 4624, 4776)
            if record.EventID in [4624, 4776]:
                event_time = record.TimeGenerated.Format()
                user_name = None
                domain_name = None
                source_ip = None
                logon_type = None
                logon_success = "Correcta"
                
                # Extraemos detalles del evento
                for i, string in enumerate(record.StringInserts):
                    if "Account Name" in string:
                        user_name = string.split(":")[-1].strip()
                    elif "Account Domain" in string:
                        domain_name = string.split(":")[-1].strip()
                    elif "Source Network Address" in string:
                        source_ip = string.split(":")[-1].strip()
                    elif "Logon Type" in string:
                        logon_type_code = int(record.StringInserts[8])
                        if logon_type_code in logon_types:
                            logon_type = logon_types[logon_type_code]
                        else:
                            continue  # Saltamos este evento si el tipo de inicio de sesión no está en logon_types
                
                # Detectar si fue un inicio exitoso o no (ID 4624 suele ser exitoso, 4776 depende de otros factores)
                if record.EventID == 4776 and "FAILURE" in record.StringInserts[0]:
                    logon_success = "Fallida"
                
                # Solo añadimos el evento si el tipo de inicio de sesión está en logon_types
                if logon_type:
                    events.append({
                        "Event ID": record.EventID,
                        "Time Generated": event_time,
                        "User Name": user_name,
                        "Domain": domain_name,
                        "Source IP": source_ip,
                        "Logon Type": logon_type,
                        "Logon Success": logon_success
                    })
    
    win32evtlog.CloseEventLog(hand)
    return events

# Guardamos los eventos en un archivo y los mostramos por pantalla
events = export_remote_connections()

# Formateo de los eventos para la exportación
formatted_events = []
for event in events:
    formatted_event = (
        f"Fecha: {event['Time Generated']}\n"
        f"Usuario: {event['User Name']}\n"
        f"Dominio: {event['Domain']}\n"
        f"IP de origen: {event['Source IP']}\n"
        f"Tipo de conexión: {event['Logon Type']}\n"
        f"Resultado de la conexión: {event['Logon Success']}\n"
        f"ID del evento: {event['Event ID']}\n"
        "----------------------------------------"
    )
    formatted_events.append(formatted_event)
    print(formatted_event)  # Mostrar por pantalla

# Guardar los eventos en un archivo de texto
with open("logs_de_conexiones.txt", "w") as f:
    f.write("\n\n".join(formatted_events))

print("\nEventos exportados a 'logs_de_conexiones.txt'")