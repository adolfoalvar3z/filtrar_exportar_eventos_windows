# Registro de Eventos de Inicio de Sesión en Windows

Este proyecto consiste en un script de Python que registra y exporta eventos de inicio de sesión en sistemas Windows, utilizando la biblioteca pywin32.

## Descripción

El script `eventos.py` se conecta al registro de eventos de seguridad de Windows y extrae información sobre los inicios de sesión, incluyendo:

- Fecha y hora del evento
- Nombre de usuario
- Dominio
- Dirección IP de origen
- Tipo de inicio de sesión
- Resultado del inicio de sesión (exitoso o fallido)

Los tipos de inicio de sesión monitoreados incluyen:

- Inicio interactivo (Consola)
- Red (Conexión SMB u otra red)
- Inicio de sesión por lotes
- Inicio de sesión de servicio
- Desbloqueo de la estación de trabajo
- Inicio remoto (RDP)
- Inicio en caché (sin acceso a DC)

## Requisitos

- Windows 7 o superior
- Python 3.6 o superior
- pywin32

## Instalación

1. Asegúrate de tener Python instalado en tu sistema Windows.
2. Instala pywin32 usando pip:
