from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

def dateAndHourPlusTen():
    # Definimos la zona horaria de Venezuela
    tz_venezuela = ZoneInfo("America/Caracas")

    # Obtenemos la fecha actual ajustada a la zona horaria
    now = datetime.now(tz_venezuela)

    # Creamos y sumamos los diez minutos
    future = now + timedelta(minutes=10)

    # ¡AQUÍ ESTÁ EL TRUCO! 
    # .timestamp() genera el número, e int() le quita los decimales
    return int(future.timestamp())

def dateAndHourNow():
    # Definimos la zona horaria de Venezuela
    tz_venezuela = ZoneInfo("America/Caracas")

    # Obtenemos la fecha actual ajustada a la zona horaria
    now = datetime.now(tz_venezuela)

    # La devolvemos formateada
    return now.strftime("%Y-%m-%d %H:%M:%S")
