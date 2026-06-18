import datetime
import os
import sys
import json
import time
import pandas as pd
from LogicaHora import dateAndHourNow
from playwright.sync_api import sync_playwright

# =========================================================================
#  FUNCIÓN LOCAL PARA RUTAS ABSOLUTAS EXTERNAS (CORREGIDA PARA .EXE)
# =========================================================================
def obtener_ruta_absoluta(nombre_archivo):
    """
    Garantiza que los archivos se busquen SIEMPRE en la carpeta raíz externa
    donde está el archivo ejecutable (.exe), no dentro de '_internal'.
    """
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable compilado, apunta a la carpeta del .exe
        directorio_ejecutable = os.path.dirname(sys.executable)
    else:
        # Si corre desde VSCode / PyCharm, apunta a la carpeta del script
        directorio_ejecutable = os.path.dirname(os.path.abspath(__file__))
        
    return os.path.join(directorio_ejecutable, nombre_archivo)

# --- VARIABLES GLOBALES CONFIGURABLES (Apuntando siempre afuera) ---
ARCHIVO_ENTRADA = obtener_ruta_absoluta("Solicitudes.xlsx")
PESTANA_ENTRADA = "Limites Transferencias"   # Pestaña de donde se LEYERAN los datos

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Limites Transferencias"   # Pestaña donde se GUARDARÁ el acumulado
CARPETA_TXT = obtener_ruta_absoluta("Resultados_Limites_TXT")
# ----------------------------------------

def leer_datos_entrada_directo(ruta_excel, nombre_pestaña):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=nombre_pestaña, dtype=str)
        print(f"📖 Leyendo datos con éxito desde '{os.path.basename(ARCHIVO_ENTRADA)}' -> Pestaña: '{nombre_pestaña}'")
    except ValueError:
        print(f"❌ Error: No existe la pestaña '{nombre_pestaña}' dentro de {os.path.basename(ARCHIVO_ENTRADA)}")
        return []
    except FileNotFoundError:
        print(f"❌ No se encontró el archivo: {ruta_excel}")
        return []
    except Exception as e:
        print(f"❌ Error leyendo Excel '{ruta_excel}': {e}")
        return []

    df = df.fillna('')
    df.columns = [str(c).strip().lower() for c in df.columns]

    filas = []
    for _, row in df.iterrows():
        fila = {str(k).strip().lower(): str(v).strip() if not pd.isna(v) else '' for k, v in row.items()}
        if any(str(v).strip() for v in fila.values()):
            filas.append(fila)
    return filas

def aplanar_dict(d, parent_key='', sep='_'):
    """Rompe diccionarios anidados (JSON) y los vuelve columnas planas"""
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

def run():
    # Usamos directamente las variables globales que ya traen la ruta absoluta calculada correctamente
    datos_a_procesar = leer_datos_entrada_directo(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    
    if not datos_a_procesar:
        print("⚠️ No se encontraron registros válidos para procesar. Verifica el archivo y la pestaña.")
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
        print(f"{CYAN}║{RESET}  {VERDE}[1]{RESET} Guardar acumulado en pestaña EXCEL ({PESTANA_DESTINO})   {CYAN}║{RESET}")
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
        "Env_datosPeticion_idCliente",
        "Env_datosPeticion_idConsumidor",
        "Res_datos_registros_cuenta", 
        "Res_datos_registros_numero_documento",
        "Res_datos_registros_id_cliente",
        "Res_datos_registros_limiteTransaccion",
        "Res_datos_registros_limiteDiario",
        "Res_datos_registros_limiteMensual",
        "Res_datos_registros_acumuladoDiario",
        "Res_datos_registros_acumuladoMensual",
        "Res_datos_registros_disponibleDiario",
        "Res_datos_registros_disponibleMensual"
    ]

    # --- AUTOMATIZACIÓN PLAYWRIGHT ---
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True) 
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        url = "http://172.19.161.137:7006/swagger-ui/index.html"
        
        try:
            print(">>> Navegando a Swagger...")
            page.goto(url, wait_until="load", timeout=90000)
            
            print("🔍 Buscando el contenedor del endpoint...")
            bloque_api = page.locator(".opblock").first
            bloque_api.scroll_into_view_if_needed()
            
            clase_bloque = bloque_api.get_attribute("class") or ""
            if "is-open" not in clase_bloque:
                print("🔓 Desplegando el acordeón del Swagger...")
                bloque_api.locator(".opblock-summary").click()
                page.wait_for_timeout(1000)

            btn_try = bloque_api.locator("button.try-out__btn")
            btn_try.wait_for(state="visible", timeout=20000)
            
            if "Cancel" not in btn_try.inner_text():
                btn_try.click()
                print("✅ Modo 'Try it out' activado.")
            
            for i, fila in enumerate(datos_a_procesar):
                id_cliente = str(fila.get('id_cliente', '')).strip()
                
                if not id_cliente or id_cliente.lower() == 'nan':
                    print(f"[{i+1}] Saltando registro vacío o sin id_cliente...")
                    continue

                try:
                    val_canal = fila.get('id_canal', 1)
                    id_canal = int(float(val_canal)) if pd.notna(val_canal) else 1
                except:
                    id_canal = 1

                id_usuario = str(fila.get('id_usuario', 'E123456')).strip()
                id_consumidor = str(fila.get('id_consumidor', '')).strip()

                payload = {
                    "datosPeticion": {
                        "idCliente": id_cliente,
                        "idSesion": datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")[:16],
                        "idCanal": id_canal,
                        "idUsuario": id_usuario if id_usuario != "" else "E123456",
                        "idTerminal": "agencia1",
                        "ip": "0.0.0.0",
                        "idConsumidor": id_consumidor if id_consumidor != "" else id_cliente
                    }
                }

                print(f"🔄 [{i+1}/{len(datos_a_procesar)}] Procesando Cliente: {id_cliente}")
                marca_tiempo_actual = dateAndHourNow()
                payload_str = json.dumps(payload, indent=2)

                textarea = bloque_api.locator("textarea.body-param__text")
                textarea.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                textarea.fill(payload_str)
                
                bloque_api.locator("button.execute").click()

                selector_descarga = ".download-contents"
                page.wait_for_selector(selector_descarga, state="visible", timeout=30000)
                
                try:
                    with page.expect_download(timeout=15000) as download_info:
                        bloque_api.locator(selector_descarga).first.click()
                    
                    download = download_info.value
                    temp_path = download.path()
                    
                    with open(temp_path, 'r', encoding='utf-8') as f_temp:
                        contenido_respuesta = f_temp.read()
                    
                    download.delete()

                    if opcion == "1":
                        base_fila = {
                            "Fecha, Hora y Día Ejecución": marca_tiempo_actual
                        }
                        datos_enviados_planos = aplanar_dict(json.loads(payload_str), parent_key='Env')
                        base_fila.update(datos_enviados_planos)

                        try:
                            data_json = json.loads(contenido_respuesta)
                            datos_recibidos_base = aplanar_dict(data_json, parent_key='Res')
                            
                            lista_registros = data_json.get("datos", {}).get("registros", [])
                            
                            if isinstance(lista_registros, list) and len(lista_registros) > 0:
                                for reg in lista_registros:
                                    fila_registro = base_fila.copy()
                                    fila_registro.update(datos_recibidos_base)
                                    
                                    registro_plano = aplanar_dict(reg, parent_key='Res_datos_registros')
                                    fila_registro.update(registro_plano)
                                    
                                    for col in columnas_formato_texto:
                                        if col in fila_registro and fila_registro[col] != "":
                                            valor_str = str(fila_registro[col]).strip()
                                            if not valor_str.startswith("'"):
                                                fila_registro[col] = f"'{valor_str}"
                                    registros_finales.append(fila_registro)
                            else:
                                fila_vacia = base_fila.copy()
                                fila_vacia.update(datos_recibidos_base)
                                for col in columnas_formato_texto:
                                    if col in fila_vacia and fila_vacia[col] != "":
                                        valor_str = str(fila_vacia[col]).strip()
                                        if not valor_str.startswith("'"):
                                            fila_vacia[col] = f"'{valor_str}"
                                registros_finales.append(fila_vacia)
                        except Exception:
                            fila_error = base_fila.copy()
                            fila_error["Res_Error_Respuesta"] = contenido_respuesta
                            registros_finales.append(fila_error)

                    elif opcion == "2":
                        nombre_txt = f"Resultado_{id_cliente}_{i+1}.txt"
                        ruta_txt_final = os.path.join(CARPETA_TXT, nombre_txt)
                        
                        with open(ruta_txt_final, "w", encoding="utf-8") as f_out:
                            f_out.write(f"EJECUCION: {marca_tiempo_actual}\n")
                            f_out.write(f"--- ENVÍO ---\n{payload_str}\n\n")
                            f_out.write(f"--- RESPUESTA ---\n{contenido_respuesta}")
                        print(f"💾 Guardado individual texto: {ruta_txt_final}")

                except Exception as e:
                    print(f"⚠️ Error capturando o guardando respuesta para {id_cliente}: {e}")
                
                time.sleep(0.5)

            # --- ALMACENAMIENTO ACUMULATIVO ---
            if opcion == "1" and registros_finales:
                print(f"\n📊 Abriendo {os.path.basename(ARCHIVO_EDITABLE)} para actualizar la pestaña '{PESTANA_DESTINO}'...")
                df_nuevos_datos = pd.DataFrame(registros_finales)
                
                if os.path.exists(ARCHIVO_EDITABLE):
                    try:
                        df_historico = pd.read_excel(ARCHIVO_EDITABLE, sheet_name=PESTANA_DESTINO, dtype=str).fillna('')
                        df_consolidado_total = pd.concat([df_historico, df_nuevos_datos], ignore_index=True)
                        print(f"📚 Datos previos acoplados. Añadiendo {len(df_nuevos_datos)} líneas al final...")
                    except Exception:
                        print(f"🆕 La pestaña '{PESTANA_DESTINO}' no existía en el archivo. Se creará ahora.")
                        df_consolidado_total = df_nuevos_datos
                    
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_consolidado_total.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                else:
                    print(f"📦 El archivo {os.path.basename(ARCHIVO_EDITABLE)} no existía. Creándolo desde cero en la raíz...")
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl') as writer:
                        df_nuevos_datos.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                
                print(f"✨ ¡Datos acumulados exitosamente en la pestaña '{PESTANA_DESTINO}' de '{os.path.basename(ARCHIVO_EDITABLE)}'!")
            elif opcion == "1":
                print("⚠️ No se generaron registros nuevos para guardar en Excel.")

        except Exception as e:
            print(f"❌ Error general en la automatización: {e}")
        finally:
            print("🏁 Proceso finalizado por completo.")
            browser.close()

if __name__ == "__main__":
    run()