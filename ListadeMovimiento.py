import os
import sys
import json
import pandas as pd
from playwright.sync_api import sync_playwright
from LeerXlsx import leer_datos_lista_movimientos
from LogicaHora import dateAndHourNow
from LogicaDeEntrada import payLoadListaMovimientos

# =========================================================================
#  FUNCIÓN LOCAL PARA RUTAS ABSOLUTAS EXTERNAS (CORREGIDA PARA .EXE)
# =========================================================================
def obtener_ruta_absoluta(nombre_archivo):
    """
    Garantiza que los archivos se busquen SIEMPRE en la carpeta raíz externa
    donde está el archivo ejecutable (.exe), no dentro de '_internal'.
    """
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable comprimido, apunta a la carpeta del .exe
        directorio_ejecutable = os.path.dirname(sys.executable)
    else:
        # Si corre desde VSCode / PyCharm, apunta a la carpeta del script
        directorio_ejecutable = os.path.dirname(os.path.abspath(__file__))
        
    return os.path.join(directorio_ejecutable, nombre_archivo)

# --- VARIABLES GLOBALES CONFIGURABLES (Apuntando siempre afuera) ---
ARCHIVO_ENTRADA = obtener_ruta_absoluta("Solicitudes.xlsx")
ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Lista de Movimientos"
CARPETA_TXT = obtener_ruta_absoluta("Resultados_TXT")
# ----------------------------------------

def aplanar_dict(d, parent_key='', sep='_'):
    """Rompe diccionarios aninados (JSON) y los vuelve columnas planas"""
    items = []
    if not isinstance(d, dict):
        return {}
    for k, v in d.items():
        if k == 'registros' and isinstance(v, list):
            continue
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(aplanar_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def formatear_fecha_ddmmyyyy(valor):
    """Fuerza el formato DD/MM/AAAA sin importar el formato origen"""
    if valor is None:
        return valor
    if hasattr(valor, 'strftime'):
        return valor.strftime('%d/%m/%Y')
    if not isinstance(valor, str):
        valor = str(valor)
    valor = valor.strip()
    if len(valor) >= 10 and valor[4] == "-" and valor[7] == "-":
        fecha_corta = valor[:10]
        año, mes, dia = fecha_corta.split("-")
        return f"{dia}/{mes}/{año}"
    if len(valor) == 8 and valor.isdigit():
        return f"{valor[6:8]}/{valor[4:6]}/{valor[0:4]}"
    return valor


def normalizar_fecha_payload(obj):
    if isinstance(obj, dict):
        return {
            k: normalizar_fecha_payload(
                formatear_fecha_ddmmyyyy(v)
                if "fecha" in k.lower()
                else v
            )
            for k, v in obj.items()
        }
    if isinstance(obj, list):
        return [normalizar_fecha_payload(item) for item in obj]
    return obj


def listaMovimientos():
    # Usamos la ruta global externa para Solicitudes.xlsx
    lista_usuarios = leer_datos_lista_movimientos(ARCHIVO_ENTRADA)
    if not lista_usuarios: 
        print(f"⚠️ No se encontraron datos válidos para procesar. Verifique el archivo: {ARCHIVO_ENTRADA}")
        return

    # --- MENÚ INTERACTIVO EN TERMINAL ---
    while True: 
        CYAN = "\033[36m"
        VERDE = "\033[32m"
        RESET = "\033[0m"
        NEGRITA = "\033[1m"
        
        print(f"{CYAN}╔════════════════════════════════════════════════════════════════════╗{RESET}")
        print(f"{CYAN}║{RESET}                    {NEGRITA}SELECCIONE EL FORMATO DE SALIDA{RESET}                 {CYAN}║{RESET}")
        print(f"{CYAN}╠════════════════════════════════════════════════════════════════════╣{RESET}")  
        print(f"{CYAN}║{RESET}  {VERDE}[1]{RESET} Guardar acumulado en pestaña EXCEL ({PESTANA_DESTINO})     {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[2]{RESET} Guardar archivos individuales en formato TXT                  {CYAN}║{RESET}")
        print(f"{CYAN}╚════════════════════════════════════════════════════════════════════╝{RESET}")

        opcion = input(f"\n{NEGRITA}👉 Selecciona una opción:{RESET} ").strip()
    
        if opcion in ["1", "2"]:
            break
        else:
            print("❌ Opción inválida. Intente de nuevo.")

    if opcion == "2" and not os.path.exists(CARPETA_TXT):
        os.makedirs(CARPETA_TXT)
        print(f"📂 Carpeta externa '{os.path.basename(CARPETA_TXT)}' creada para guardar los archivos de texto.")

    registros_finales = []

    columnas_formato_texto = [
        "Env_cuenta",
        "Res_datos_registros_numero_documento_empresa",
        "Res_datos_registros_cuenta_empresa",
        "Res_datos_registros_referencia",
        "Res_datos_registros_numero_documento_contraparte",
        "Res_datos_registros_telefono_contraparte",
        "Res_datos_registros_monto"
    ]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        print("\n🌐 Conectando a Swagger UI...")
        page.goto("http://172.19.161.137:9001/swagger-ui/index.html")
        
        selector_v4 = "#operations-listar-movimientos-controller-v-4-consultarCuentas"
        selector_btn_descarga = f"{selector_v4} .download-contents"
        
        endpoint_v4 = page.locator(selector_v4)
        endpoint_v4.scroll_into_view_if_needed()
        endpoint_v4.click()
        
        endpoint_v4.get_by_role("button", name="Try it out").click()
        textarea = endpoint_v4.locator("textarea.body-param__text")
        
        for i, datos in enumerate(lista_usuarios):
            print(f"🔄 [{i+1}/{len(lista_usuarios)}] Ejecutando RIF: {datos['rif']} - Cuenta: {datos['cuenta'] if datos['cuenta'] else 'VACÍA'}...")
            
            marca_tiempo_actual = dateAndHourNow()
            
            try:
                payload_dict = payLoadListaMovimientos(datos)
                payload_dict = normalizar_fecha_payload(payload_dict)
                payload_envio_str = json.dumps(payload_dict, indent=2)
                
                textarea.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                textarea.fill(payload_envio_str)
                
                # Clic en Execute
                endpoint_v4.locator(".btn.execute").click()
                
                # Tiempo de espera calibrado a 7 segundos
                print("   ⏳ Esperando 7 segundos para la carga completa de datos...")
                page.wait_for_timeout(7000) 
                
                page.wait_for_selector(selector_btn_descarga, timeout=30000)
                
                # Descarga del archivo individual
                with page.expect_download(timeout=15000) as download_info:
                    page.locator(selector_btn_descarga).click()
                
                download = download_info.value
                path_temp_descarga = download.path()
                
                with open(path_temp_descarga, "r", encoding="utf-8") as f_descarga:
                    contenido_respuesta = f_descarga.read()
                
                download.delete()

                # --- MODO 1: PROCESAR UNO POR UNO PARA EL EXCEL ---
                if opcion == "1":
                    base_fila = {
                        "Fecha, Hora y Día Ejecución": marca_tiempo_actual
                    }
                    datos_enviados_planos = aplanar_dict(payload_dict, parent_key='Env')
                    base_fila.update(datos_enviados_planos)

                    try:    
                        data_json = json.loads(contenido_respuesta)
                        datos_recibidos_base = aplanar_dict(data_json, parent_key='Res')
                        
                        lista_movimientos_data = []
                        if isinstance(data_json, dict):
                            if "datos" in data_json and isinstance(data_json["datos"], dict):
                                lista_movimientos_data = data_json["datos"].get("registros", [])
                            if not lista_movimientos_data:
                                lista_movimientos_data = data_json.get("registros", [])

                        if isinstance(lista_movimientos_data, list) and len(lista_movimientos_data) > 0:
                            print(f"   ↳ ✨ Descargados e identificados {len(lista_movimientos_data)} registros. Insertando uno por uno...")
                            for movimiento in lista_movimientos_data:
                                fila_movimiento = base_fila.copy()
                                fila_movimiento.update(datos_recibidos_base)
                                
                                movimiento_plano = aplanar_dict(movimiento, parent_key='Res_datos_registros')
                                fila_movimiento.update(movimiento_plano)
                                
                                for col in columnas_formato_texto:
                                    if col in fila_movimiento and fila_movimiento[col] is not None:
                                        valor_str = str(fila_movimiento[col]).strip()
                                        if valor_str != "" and not valor_str.startswith("'"):
                                            fila_movimiento[col] = f"'{valor_str}"
                                            
                                registros_finales.append(fila_movimiento)
                        else:
                            fila_vacia = base_fila.copy()
                            fila_vacia.update(datos_recibidos_base)
                            registros_finales.append(fila_vacia)
                            
                    except Exception as e_json:
                        fila_error = base_fila.copy()
                        fila_error["Res_Error_Respuesta"] = f"Error interpretando JSON: {str(e_json)} | Respuesta original: {contenido_respuesta}"
                        registros_finales.append(fila_error)

                # --- MODO 2: GUARDAR DIRECTO A ARCHIVO TXT ---
                elif opcion == "2":
                    cuenta_limpia = datos['cuenta'] if datos['cuenta'] else "SIN_CUENTA"
                    nombre_txt = f"Resultado_{datos['rif']}_{cuenta_limpia}_{i}.txt"
                    ruta_txt_final = os.path.join(CARPETA_TXT, nombre_txt)
                    
                    with open(ruta_txt_final, "w", encoding="utf-8") as f_out:
                        f_out.write(f"EJECUCION: {marca_tiempo_actual}\n")
                        f_out.write(f"--- ENVÍO ---\n{payload_envio_str}\n\n")
                        f_out.write(f"--- RESPUESTA ---\n{contenido_respuesta}")
                    print(f"💾 Guardado individual en texto: {ruta_txt_final}")

                btn_clear = endpoint_v4.locator(".btn-group .clear")
                if btn_clear.is_visible():
                    btn_clear.click()
                   
            except Exception as e:
                print(f"⚠️ Error procesando registro {i}: {e}")
                try:
                    endpoint_v4.locator(".btn-group .clear").click()
                except:
                    pass

        # --- CONSOLIDACIÓN Y ACTUALIZACIÓN EN EL EXCEL ---
        if opcion == "1":
            if registros_finales:
                df_nuevos_datos = pd.DataFrame(registros_finales)
                
                print(f"\n📦 Actualizando base de datos acumulativa en: {ARCHIVO_EDITABLE}...")

                if os.path.exists(ARCHIVO_EDITABLE):
                    try:
                        df_existente = pd.read_excel(ARCHIVO_EDITABLE, sheet_name=PESTANA_DESTINO)
                        df_consolidado = pd.concat([df_existente, df_nuevos_datos], ignore_index=True)
                        print(f"🔄 Archivo existente localizado. Se agregaron {len(df_nuevos_datos)} nuevas filas de transacciones.")
                    except Exception:
                        df_consolidado = df_nuevos_datos
                        print(f"➕ Archivo base existente, pero la pestaña '{PESTANA_DESTINO}' es nueva. Creándola...")
                    
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_consolidado.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                else:
                    print(f"📝 El archivo no existía. Creando '{os.path.basename(ARCHIVO_EDITABLE)}' por primera vez en la raíz...")
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl') as writer:
                        df_nuevos_datos.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                
                print(f"✨ ¡Todos los registros desglosados fueron acumulados con éxito!")
            else:
                print("⚠️ No se pudieron estructurar nuevos registros en esta sesión.")
        else:
            print(f"\n✨ Proceso completado. Todos los archivos de texto están listos en la carpeta '{os.path.basename(CARPETA_TXT)}'.")

        print("🏁 Proceso terminado.")
        browser.close()

if __name__ == "__main__":
    listaMovimientos()