import datetime
import os
import sys
import json
import time
import re
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
PESTANA_ENTRADA = "Limites Pago Movil"       # <-- Pestaña de donde se LEYERAN los datos

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Limites Pago Móvil"       # <-- Pestaña donde se GUARDARÁ el acumulado
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

def limpiar_telefono(valor):
    """Limpia y normaliza el teléfono al formato 584XXXXXXXXX"""
    tel = str(valor).strip().split('.')[0]
    tel = re.sub(r'\D', '', tel)
    if not tel: return ""

    if tel.startswith("580"):
        tel = "58" + tel[3:]
    elif tel.startswith("0"):
        tel = "58" + tel[1:]
    elif not tel.startswith("58"):
        tel = "58" + tel
    return tel

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

def ConsultaLimitesPagoMovil():
    # Usamos directamente la variable global con su ruta absoluta calculada correctamente
    lista_datos = leer_datos_entrada_directo(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    
    if not lista_datos:
        print("⚠️ No se encontraron registros válidos para procesar. Verifica el archivo y la pestaña.")
        return

    registros_finales = []
    
    columnas_formato_texto = [
        "env_telefonocliente", 
        "res_datos_limites_nrocuenta",
        "res_datos_limites_montopagosnaturalesdiadisp",
        "res_datos_limites_montopagosnaturalesmesdisp",
        "res_datos_limites_cantpagosnaturalesdia",
        "res_datos_limites_montopagosnaturalesdia",
        "res_datos_limites_cantpagosnaturalesmesdisp",
        "res_datos_limites_cantpagosnaturalesmes",
        "res_datos_limites_montopagosnaturalesmes",
        "res_datos_limites_cantpagosjuridicosdiadisp",
        "res_datos_limites_montopagosjuridicosdiadisp",
        "res_datos_limites_cantpagosjuridicosdia",
        "res_datos_limites_montopagosjuridicosdia",
        "res_datos_limites_cantpagosjuridicosmesdisp",
        "res_datos_limites_montopagosjuridicosmesdisp",
        "res_datos_limites_cantpagosjuridicosmes",
        "res_datos_limites_montopagosjuridicosmes",
        "res_comision_monto"
    ]

    # --- AUTOMATIZACIÓN PLAYWRIGHT ---
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True) 
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        url = "http://172.19.161.113:9009/swagger-ui/index.html?configUrl=/v3/api-docs/swagger-config#/consultas-limites-controller/consultasLimites"
        
        try:
            print("🔗 Navegando a Swagger en segundo plano (Invisible)...")
            page.goto(url, wait_until="load", timeout=90000)
            
            print("🔍 Buscando el contenedor del endpoint...")
            bloque_api = page.locator(".opblock").first
            bloque_api.scroll_into_view_if_needed()
            
            clase_bloque = bloque_api.get_attribute("class") or ""
            if "is-open" not in clase_bloque:
                bloque_api.locator(".opblock-summary").click()
                page.wait_for_timeout(1000)

            btn_try = bloque_api.locator("button.try-out__btn")
            btn_try.wait_for(state="visible", timeout=20000)
            
            if "Cancel" not in btn_try.inner_text():
                btn_try.click()
                print("✅ Modo 'Try it out' activado.")

            for i, fila in enumerate(lista_datos):
                id_cliente_actual = str(fila.get('id_cliente', '')).strip()
                
                if not id_cliente_actual or id_cliente_actual.lower() == 'nan':
                    print(f"[{i+1}] Saltando registro vacío o sin id_Cliente...")
                    continue

                print(f"🔄 [{i+1}/{len(lista_datos)}] Procesando Cliente: {id_cliente_actual}")
                marca_tiempo_actual = dateAndHourNow()

                id_canal_raw = str(fila.get('id_canal', '')).strip()
                id_canal = id_canal_raw.zfill(2) if id_canal_raw and id_canal_raw.lower() != 'nan' else "00"

                payload = {
                    "idCliente": id_cliente_actual,
                    "idUsuario": "usuario_beca",
                    "idTerminal": "terminal_beca",
                    "idCanal": id_canal,
                    "idConsumidor": str(fila.get('id_consumidor', '')).strip(),
                    "telefonoCliente": limpiar_telefono(fila.get('telefono_cliente', '')),
                    "ipOrigen": "111.111.11.1"
                }

                textarea = bloque_api.locator("textarea.body-param__text")
                textarea.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                textarea.fill(json.dumps(payload, indent=2))

                bloque_api.locator("button.execute").click()

                selector_res = ".responses-wrapper .response .highlight-code pre"
                try:
                    page.wait_for_selector(selector_res, timeout=15000)
                    page.wait_for_timeout(500)
                    resultado_texto = page.locator(selector_res).first.inner_text()
                except:
                    resultado_texto = "⚠️ Error: No se encontró la respuesta."

                registro_fila = {
                    "Fecha, Hora y Día Ejecución": marca_tiempo_actual
                }

                datos_enviados_planos = aplanar_dict(payload, parent_key='Env')
                registro_fila.update(datos_enviados_planos)

                try:
                    data_json = json.loads(resultado_texto)
                    datos_recibidos_planos = aplanar_dict(data_json, parent_key='Res')
                    registro_fila.update(datos_recibidos_planos)
                except Exception:
                    registro_fila["Res_Error_Respuesta"] = resultado_texto

                fila_limpia = {}
                for k, v in registro_fila.items():
                    llave_min = str(k).lower()
                    val_str = str(v).strip()
                    
                    if llave_min in columnas_formato_texto and val_str != "":
                        if not val_str.startswith("'"):
                            val_str = f"'{val_str}"
                    fila_limpia[k] = val_str

                registros_finales.append(fila_limpia)

        except Exception as e:
            print(f"❌ Error en el proceso general: {e}")
        finally:
            # --- ALMACENAMIENTO ACUMULATIVO ---
            if registros_finales:
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
            else:
                print("⚠️ No se generaron registros nuevos para guardar en Excel.")

            print("🏁 Proceso finalizado por completo.")
            browser.close()

if __name__ == "__main__":
    ConsultaLimitesPagoMovil()