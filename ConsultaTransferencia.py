import datetime
import os
import sys
import json
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
PESTANA_ENTRADA = "Consulta Transferencias"  # Busca esta pestaña en Solicitudes

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Consulta Transferencias"   # Guarda el acumulado aquí
CARPETA_TXT = obtener_ruta_absoluta("Resultados_Transferencias_TXT")
# ----------------------------------------

def leer_datos_desde_excel(ruta_excel, nombre_pestaña):
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
        fila = {
            str(k).strip().lower(): str(v).strip() if not pd.isna(v) else ''
            for k, v in row.items()
        }
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


def formatear_fecha_yyyymmdd(valor):
    """
    Fuerza estrictamente el formato AAAA-MM-DD requerido por tu sistema,
    soportando objetos datetime, strings de Excel (DD/MM/AAAA) y textos continuos.
    """
    if not valor or str(valor).strip().lower() in ['nan', '']:
        return ""
        
    if hasattr(valor, 'strftime'):
        return valor.strftime('%Y-%m-%d')
        
    if not isinstance(valor, str):
        valor = str(valor)
        
    valor = valor.strip()
    
    if len(valor) >= 10 and valor[4] == "-" and valor[7] == "-":
        return valor[:10]
        
    if len(valor) >= 10 and valor[4] == "/" and valor[7] == "/":
        return valor[:10].replace("/", "-")

    if "/" in valor:
        partes = valor.split(" ")[0].split("/")
        if len(partes) == 3:
            if len(partes[0]) == 4:
                return f"{partes[0]}-{partes[1]}-{partes[2]}"
            return f"{partes[2]}-{partes[1]}-{partes[0]}"
        
    if len(valor) == 8 and valor.isdigit():
        return f"{valor[0:4]}-{valor[4:6]}-{valor[6:8]}"
        
    return valor


def generar_json_payload(fila):
    canal_val = fila.get('id_canal', '')
    if canal_val and canal_val.lower() != 'nan':
        try: canal_val = int(float(canal_val))
        except: pass
        
    monto_val = fila.get('monto', '')
    if monto_val and monto_val.lower() != 'nan':
        try: monto_val = float(monto_val)
        except: monto_val = 0.0
    else:
        monto_val = 0.0

    banco = fila.get('bancocontraparte', '').strip()
    otro_banco_val = fila.get('otrobanco', '')
    
    if otro_banco_val.lower() in ['true', '1']:
        otro_banco = True
    elif otro_banco_val.lower() in ['false', '0']:
        otro_banco = False
    else:
        otro_banco = False if banco == "0115" else True

    pos_val = fila.get('posicioninicial', '')
    if pos_val and pos_val.lower() != 'nan':
        try: pos_val = int(float(pos_val))
        except: pass
    else:
        pos_val = 0

    payload = {
        "datosPeticion": {
            "idCliente": fila.get('id_cliente', ''),
            "idSesion": datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")[:16],
            "idCanal": canal_val if canal_val != "" else 1,
            "idUsuario": fila.get('id_usuario', 'E123456') if fila.get('id_usuario', '') != "" else "E123456",
            "idTerminal": "agencia1",
            "ip": "0.0.0.0",
            "idConsumidor": fila.get('id_consumidor', '') if fila.get('id_consumidor', '') != "" else fila.get('id_cliente', '')
        },
        "fecha": formatear_fecha_yyyymmdd(fila.get('fecha')), 
        "documentoContraparte": fila.get('documentocontraparte', '').replace("-", ""),
        "referencia": fila.get('referencia', ''),
        "cuenta": fila.get('cuenta', ''),
        "bancoContraparte": banco,
        "monto": monto_val,
        "telefonoContraparte": fila.get('telefonocontraparte', '').replace("-", ""),
        "tipoConsulta": fila.get('tipoconsulta', 'D').upper() if fila.get('tipoconsulta') else "D",
        "otroBanco": otro_banco,
        "posicionInitial": pos_val
    }
    return json.dumps(payload, indent=2)

def run():
    lista = leer_datos_desde_excel(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    if not lista: 
        print("⚠️ No se encontraron registros válidos para procesar.")
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
        print(f"{CYAN}║{RESET}  {VERDE}[1]{RESET} Guardar acumulado en pestaña EXCEL ({PESTANA_DESTINO})  {CYAN}║{RESET}")
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
        "Env_documentoContraparte",
        "Env_referencia",
        "Env_cuenta",
        "Env_telefonoContraparte",
        "Res_datos_registros_numero_documento_empresa",
        "Res_datos_registros_cuenta_empresa",
        "Res_datos_registros_referencia",
        "Res_datos_registros_numero_documento_contraparte",
        "Res_datos_registros_telefono_contraparte",
        "Res_datos_registros_monto"
    ]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        url = "http://172.19.161.137:7004/swagger-ui/index.html#/consulta-transferencias-controller-v-4/consultasTransferenciasInmediatasDebitoCredito"
        
        try:
            print("\n🔗 Cargando Swagger...")
            page.goto(url, wait_until="load", timeout=90000)
            
            selector_v4 = "#operations-consulta-transferencias-controller-v-4-consultasTransferenciasInmediatasDebitoCredito"
            
            # LINEA CORREGIDA AQUÍ:
            section = page.locator(selector_v4)
            section.scroll_into_view_if_needed()
            
            if "is-open" not in (section.get_attribute("class") or ""):
                section.click()

            btn_try = section.locator("button.try-out__btn")
            btn_try.wait_for(state="visible")
            if "Cancel" not in btn_try.inner_text():
                btn_try.click()

            for i, fila in enumerate(lista):
                id_cliente = fila.get('id_cliente', 'desconocido')
                cuenta_log = fila.get('cuenta', 'vacia')
                print(f"🔄 [{i+1}/{len(lista)}] Ejecutando Cliente: {id_cliente} - Cuenta: {cuenta_log}")
                
                marca_tiempo_actual = dateAndHourNow()
                payload_str = generar_json_payload(fila)
                
                textarea = section.locator("textarea.body-param__text")
                textarea.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                textarea.fill(payload_str)
                
                section.locator("button.execute").click()
                
                selector_descarga = f"{selector_v4} div.download-contents"
                
                try:
                    # El script pasa directo a esperar de forma dinámica el selector del botón de descarga
                    page.wait_for_selector(selector_descarga, state="visible", timeout=30000)
                    
                    with page.expect_download(timeout=15000) as download_info:
                        page.locator(selector_descarga).click()
                    
                    download = download_info.value
                    temp_path = download.path()
                    
                    with open(temp_path, "r", encoding="utf-8") as f_res:
                        contenido_respuesta = f_res.read()
                    
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
                            
                            lista_movimientos = data_json.get("datos", {}).get("registros", [])
                            
                            if isinstance(lista_movimientos, list) and len(lista_movimientos) > 0:
                                for movimiento in lista_movimientos:
                                    fila_movimiento = base_fila.copy()
                                    fila_movimiento.update(datos_recibidos_base)
                                    
                                    movimiento_plano = aplanar_dict(movimiento, parent_key='Res_datos_registros')
                                    fila_movimiento.update(movimiento_plano)
                                    
                                    for col in columnas_formato_texto:
                                        if col in fila_movimiento and fila_movimiento[col] != "":
                                            valor_str = str(fila_movimiento[col]).strip()
                                            if not valor_str.startswith("'"):
                                                fila_movimiento[col] = f"'{valor_str}"
                                    registros_finales.append(fila_movimiento)
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
                        cuenta_limpia = fila['cuenta'] if fila['cuenta'] else "SIN_CUENTA"
                        nombre_txt = f"Resultado_{id_cliente}_{cuenta_limpia}_{i+1}.txt"
                        ruta_txt_final = os.path.join(CARPETA_TXT, nombre_txt)
                        
                        with open(ruta_txt_final, "w", encoding="utf-8") as f_out:
                            f_out.write(f"EJECUCION: {marca_tiempo_actual}\n")
                            f_out.write(f"--- ENVÍO ---\n{payload_str}\n\n")
                            f_out.write(f"--- RESPUESTA ---\n{contenido_respuesta}")
                        print(f"💾 Guardado individual texto: {ruta_txt_final}")

                except Exception as e:
                    print(f"⚠️ Error capturando respuesta para {id_cliente}: {e}")

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
                    print(f"📦 El archivo no existía. Creando '{os.path.basename(ARCHIVO_EDITABLE)}' desde cero en la raíz...")
                    with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl') as writer:
                        df_nuevos_datos.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
                
                print(f"✨ ¡Datos acumulados exitosamente en la pestaña '{PESTANA_DESTINO}'!")
            elif opcion == "1":
                print("⚠️ No se generaron registros nuevos para guardar en Excel.")

        except Exception as e:
            print(f"❌ Error general en la automatización: {e}")
        finally:
            print("🏁 Proceso finalizado por completo.")
            browser.close()

if __name__ == "__main__":
    run()