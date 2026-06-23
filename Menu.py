import os
from EscrituraDatosExcel import reiniciar_archivo_solicitudes
from ConsultaLimitesPagoMovil import ConsultaLimitesPagoMovil
from ConsultaLimitesTransferencia import run as consulta_limites
from ConsultaTransferencia import run as consulta_transferencias
from ListadeMovimiento import listaMovimientos
from PagoMovil import hacerPagoMovil
from Transferencia import transferir
from LeerXlsx import guardar_y_cerrar_excel
from ObtenerRuta import resource_path
from ListasdeBanco import run as lista_banco
from Debitoinmediato import run as debitoinmediato
from Cobro_Comercio_C2P import run as cobrocomercio
from Anulacion_C2P import run as anulacion

def limpiar_terminal():
    if os.name == 'nt':
        os.system('cls')


def main():
    while True: 
        # Colores ANSI para darle vida a la consola (se ven bien en WSL, CMD moderna y Linux)
        CYAN = "\033[36m"
        AMARILLO = "\033[33m"
        ROJO = "\033[31m"
        VERDE = "\033[32m"
        RESET = "\033[0m"
        NEGRITA = "\033[1m"

        # Encabezado principal estilizado
        print(f"\n{CYAN}╔═══════════════════════════════════════════════════════════╗{RESET}")
        print(f"{CYAN}║{RESET}  {NEGRITA}       ¡BIENVENIDO AL SISTEMA DE AUTOMATIZACIÓN!{RESET}         {CYAN}║{RESET}")
        print(f"{CYAN}╚═══════════════════════════════════════════════════════════╝{RESET}")

        # Sección de indicaciones formateada como tarjeta
        print(f"{AMARILLO}┌──────────────── INDICACIONES IMPORTANTES ─────────────────┐{RESET}")
        print(f"{AMARILLO}│{RESET} {NEGRITA}✔ Verificación:{RESET} Guarda los cambios en el archivo de       {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET}   solicitudes y revisa todo antes de ejecutar.            {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET}                                                           {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET} {NEGRITA}⚠ Restricciones:{RESET} No modifiques encabezados ni nombres     {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET}   de archivos; podría romper el sistema.                  {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET}                                                           {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET} {ROJO}⚡ Advertencia:{RESET} El programa procesará TODOS los datos     {AMARILLO}│{RESET}")
        print(f"{AMARILLO}│{RESET}   de la hoja de Excel del proceso elegido.                {AMARILLO}│{RESET}")
        print(f"{AMARILLO}└───────────────────────────────────────────────────────────┘{RESET}")
        
        print(f"¿Sugerencias u observaciones? Escríbenos al: {VERDE}0424-2035554{RESET}\n")

        # Panel de Opciones (Menú)
        print(f"{CYAN}╔═══════════════════════════════════════════════════════════╗{RESET}")
        print(f"{CYAN}║{RESET}                    {NEGRITA}PANEL DE OPCIONES{RESET}                      {CYAN}║{RESET}")
        print(f"{CYAN}╠═══════════════════════════════════════════════════════════╣{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[1]{RESET} Consultar Límites de Pago Móvil                      {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[2]{RESET} Consultar Límites de Transferencia                   {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[3]{RESET} Consultar Transferencia                              {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[4]{RESET} Consultar Lista de Movimientos                       {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[5]{RESET} Realizar Pago Móvil                                  {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[6]{RESET} Realizar Transferencias                              {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[7]{RESET} Listar Bancos                                        {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[8]{RESET} Debito inmediato                                     {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[9]{RESET} Cobro comercio C2P                                   {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[10]{RESET} Anulacion C2P                                       {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {AMARILLO}[11]{RESET} Reiniciar el archivo Solicitudes                    {CYAN}║{RESET}")
        print(f"{CYAN}╠═══════════════════════════════════════════════════════════╣{RESET}")
        print(f"{CYAN}║{RESET}  {ROJO}[0]{RESET} Salir del Menú                                       {CYAN}║{RESET}")
        print(f"{CYAN}╚═══════════════════════════════════════════════════════════╝{RESET}")
        
        opcion = input(f"\n{NEGRITA}👉 Selecciona una opción:{RESET} ").strip()
        
        if opcion == "1":
            guardar_y_cerrar_excel(resource_path("Solicitudes.xlsx"))
            guardar_y_cerrar_excel("Resultados.xlsx")
            ConsultaLimitesPagoMovil()
        elif opcion == "2":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            consulta_limites()
        elif opcion == "3":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            consulta_transferencias()
        elif opcion == "4":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            listaMovimientos()
        elif opcion == "5":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            hacerPagoMovil()
        elif opcion == "6":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            transferir()
        elif opcion == "7":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            lista_banco()
        elif opcion == "8":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            debitoinmediato()
        elif opcion == "9":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            cobrocomercio()
        elif opcion == "10":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            anulacion()
        elif opcion == "11":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            reiniciar_archivo_solicitudes()
        elif opcion == "0":
            guardar_y_cerrar_excel("Solicitudes.xlsx")
            guardar_y_cerrar_excel("Resultados.xlsx")
            print(f"\n{ROJO}👋 Cerrando el menú y finalizando procesos...{RESET}\n")
            break
        else:
            limpiar_terminal()
            continue

if __name__ == "__main__":
    main()