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
        # Si es un ejecutable comprimido, apunta a la carpeta del .exe
        directorio_ejecutable = os.path.dirname(sys.executable)
    else:
        # Si corre desde VSCode / PyCharm, apunta a la carpeta del script
        directorio_ejecutable = os.path.dirname(os.path.abspath(__file__))
        
    return os.path.join(directorio_ejecutable, nombre_archivo)

# --- VARIABLES GLOBALES CONFIGURABLES (Apuntando siempre afuera) ---
ARCHIVO_ENTRADA = obtener_ruta_absoluta("Solicitudes.xlsx")
PESTANA_ENTRADA = "Listar Bancos"   # Pestaña de lectura en Solicitudes

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Listar Bancos"   # Pestaña de guardado en Resultados
CARPETA_TXT = obtener_ruta_absoluta("Resultados_Bancos_TXT")
# ----------------------------------------

def leer_datos_entrada_directo(ruta_excel, nombre_pestaña):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=nombre_pestaña, dtype=str)
        print(f"📖 Leyendo datos con éxito desde '{os.path.basename(ruta_excel)}' -> Pestaña: '{nombre_pestaña}'")
    except Exception as e:
        print(f"❌ Error leyendo Excel '{ruta_excel}' en pestaña '{nombre_pestaña}': {e}")
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
    """Descompone diccionarios aninados (JSON) en formato plano para columnas de Excel"""
    items = []
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
    # Se eliminaron las concatenaciones locales erróneas. Ahora lee directo de la ruta global externa corregida.
    datos_a_procesar = leer_datos_entrada_directo(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    
    if not datos_a_procesar:
        print(f"⚠️ No se encontraron registros válidos en la pestaña 'Listar Bancos'. Verifica: {ARCHIVO_ENTRADA}")
        return

    # --- SELECCIÓN DE FORMATO DE SALIDA ---
    while True: 
        CYAN = "\033[36m"
        VERDE = "\033[32m"
        RESET = "\033[0m"
        NEGRITA = "\033[1m"
        
        print(f"{CYAN}╔═══════════════════════════════════════════════════════════╗{RESET}")
        print(f"{CYAN}║{RESET}                    {NEGRITA}SELECCIONE EL FORMATO DE SALIDA{RESET}        {CYAN}║{RESET}")
        print(f"{CYAN}╠═══════════════════════════════════════════════════════════╣{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[1]{RESET} Guardar acumulado en pestaña EXCEL ({PESTANA_DESTINO})   {CYAN}║{RESET}")
        print(f"{CYAN}║{RESET}  {VERDE}[2]{RESET} Guardar archivos individuales en formato TXT         {CYAN}║{RESET}")
        print(f"{CYAN}╚═══════════════════════════════════════════════════════════╝{RESET}")

        opcion = input(f"\n{NEGRITA}👉 Selecciona una opción:{RESET} ").strip()
        if opcion in ["1", "2"]:
            break
        print("❌ Opción inválida. Intente de nuevo.")

    if opcion == "2" and not os.path.exists(CARPETA_TXT):
        os.makedirs(CARPETA_TXT)
        print(f"📂 Carpeta externa '{os.path.basename(CARPETA_TXT)}' creada para los archivos individuales.")

    registros_finales = []
    
    # Columnas que necesitan conservar ceros a la izquierda o formatos numéricos largos sin convertirse a científico
    columnas_formato_texto = [
        "Env_idCliente", "Env_idConsumidor", "Env_idCanal",
        "Res_datos_bancos_codigo", "Res_datos_bancos_rif", 
        "Res_datos_registros_codigo", "Res_datos_registros_rif"
    ]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True) 
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        try:
            print(">>> Navegando al Swagger de Bancos...")
            page.goto("http://172.19.161.137:9005/swagger-ui/index.html#/listar-bancos-controller-v-3/listarBancosGeneral", wait_until="load", timeout=90000)
            
            print("🔍 Buscando el endpoint 'listarBancosGeneral'...")
            
            bloque_api = page.locator(".opblock", has=page.locator("span.url", has_text="listarBancosGeneral")).first
            
            if not bloque_api.is_visible():
                bloque_api = page.locator("[id^='operations-listar-bancos-controller-v-3-listarBancosGeneral']").first

            bloque_api.wait_for(state="attached", timeout=20000)
            bloque_api.scroll_into_view_if_needed()
            
            clase_bloque = bloque_api.get_attribute("class") or ""
            if "is-open" not in clase_bloque:
                print("🔓 Desplegando el acordeón de Swagger...")
                bloque_api.locator(".opblock-summary").click()
                page.wait_for_timeout(1500)

            btn_try = bloque_api.locator("button.try-out__btn")
            btn_try.wait_for(state="visible", timeout=15000)
            if "Cancel" not in btn_try.inner_text():
                btn_try.click()
                print("✅ Modo 'Try it out' activado con éxito.")
            
            for i, fila in enumerate(datos_a_procesar):
                id_cliente = str(fila.get('id_cliente', '')).strip()
                if not id_cliente or id_cliente.lower() == 'nan':
                    print(f"[{i+1}] Saltando registro con id_cliente inválido.")
                    continue

                id_canal = str(fila.get('id_canal', '01')).strip()
                tipo_persona = str(fila.get('tipo_persona_beneficiario', 'J')).strip()
                id_consumidor = str(fila.get('id_consumidor', id_cliente)).strip()
                tipo_op = str(fila.get('tipo_operacion', 'P')).strip()
                
                try:
                    tipo_listado = int(float(fila.get('tipo_listado', 0)))
                except:
                    tipo_listado = 0

                payload = {
                    "idCliente": id_cliente,
                    "idUsuario": "usuario_beca",
                    "idTerminal": "terminal_beca",
                    "idCanal": id_canal,
                    "tipoPersonaBeneficiario": tipo_persona,
                    "idConsumidor": id_consumidor,
                    "ipOrigen": "111.111.11.1",
                    "tipoOperacion": tipo_op,
                    "tipoListado": tipo_listado
                }

                print(f"🔄 [{i+1}/{len(datos_a_procesar)}] Consultando Bancos para Cliente: {id_cliente}")
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
                        base_fila = {"Fecha, Hora y Día Ejecución": marca_tiempo_actual}
                        datos_enviados_planos = aplanar_dict(json.loads(payload_str), parent_key='Env')
                        base_fila.update(datos_enviados_planos)

                        try:
                            data_json = json.loads(contenido_respuesta)
                            datos_recibidos_base = aplanar_dict(data_json, parent_key='Res')
                            
                            lista_bancos = data_json.get("datos", {}).get("bancos", data_json.get("datos", {}).get("registros", []))
                            
                            if isinstance(lista_bancos, list) and len(lista_bancos) > 0:
                                for banco in lista_bancos:
                                    fila_banco = base_fila.copy()
                                    fila_banco.update(datos_recibidos_base)
                                    banco_plano = aplanar_dict(banco, parent_key='Res_datos_bancos')
                                    fila_banco.update(banco_plano)
                                    
                                    for col in columnas_formato_texto:
                                        if col in fila_banco and fila_banco[col] != "":
                                            v_str = str(fila_banco[col]).strip()
                                            if not v_str.startswith("'"):
                                                fila_banco[col] = f"'{v_str}"
                                    registros_finales.append(fila_banco)
                            else:
                                fila_vacia = base_fila.copy()
                                fila_vacia.update(datos_recibidos_base)
                                registros_finales.append(fila_vacia)
                        except Exception:
                            fila_error = base_fila.copy()
                            fila_error["Res_Error_Respuesta"] = contenido_respuesta
                            registros_finales.append(fila_error)

                    elif opcion == "2":
                        nombre_txt = f"Bancos_{id_cliente}_{i+1}.txt"
                        ruta_txt_final = os.path.join(CARPETA_TXT, nombre_txt)
                        with open(ruta_txt_final, "w", encoding="utf-8") as f_out:
                            f_out.write(f"EJECUCION: {marca_tiempo_actual}\n")
                            f_out.write(f"--- ENVÍO ---\n{payload_str}\n\n")
                            f_out.write(f"--- RESPUESTA ---\n{contenido_respuesta}")
                        print(f"💾 Guardado individual texto: {ruta_txt_final}")

                except Exception as e:
                    print(f"⚠️ Error extrayendo respuesta de {id_cliente}: {e}")
                
                time.sleep(0.5)

            # --- ESCRITURA FINAL DE LOS DATOS CONCATENADOS ---
            if opcion == "1" and registros_finales:
                print(f"\n📊 Abriendo {os.path.basename(ARCHIVO_EDITABLE)} para actualizar la pestaña '{PESTANA_DESTINO}'...")
                df_nuevos = pd.DataFrame(registros_finales)
                
                if os.path.exists(ARCHIVO_EDITABLE):
                    try:
                        df_hist = pd.read_excel(ARCHIVO_EDITABLE, sheet_name=PESTANA_DESTINO, dtype=str).fillna('')
                        df_total = pd.concat([df_hist, df_nuevos], ignore_index=True)
                        print(f"📚 Datos previos acoplados. Añadiendo {len(df_nuevos)} líneas al final...")
                    except Exception:
                        print(f"🆕 La pestaña '{PESTANA_DESTINO}' no existía en el archivo. Se creará ahora.")
                        df_total = df_nuevos
                    
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_total.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                else:
                    print(f"📦 El archivo {os.path.basename(ARCHIVO_EDITABLE)} no existía. Creándolo desde cero en la raíz...")
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl') as writer:
                        df_nuevos.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                print(f"✨ ¡Datos de bancos consolidados exitosamente en la pestaña '{PESTANA_DESTINO}'!")
            elif opcion == "1":
                print("⚠️ No se generaron registros válidos para guardar en Excel.")

        except Exception as e:
            print(f"❌ Error crítico en automatización de Bancos: {e}")
        finally:
            print("🏁 Proceso de bancos concluido.")
            browser.close()

if __name__ == "__main__":
    run()