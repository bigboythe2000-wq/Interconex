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
    if getattr(sys, 'frozen', False):
        directorio_ejecutable = os.path.dirname(sys.executable)
    else:
        directorio_ejecutable = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(directorio_ejecutable, nombre_archivo)


# --- VARIABLES GLOBALES CONFIGURABLES ---
ARCHIVO_ENTRADA = obtener_ruta_absoluta("Solicitudes.xlsx")
PESTANA_ENTRADA = "Anulaciones"  # <-- Cambia aquí el nombre de tu pestaña de Excel si es otra

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Anulaciones_Resultados" 
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

def limpiar_y_formatear_telefono(valor):
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


# =========================================================================
#  PROCESAMIENTO AUTOMÁTICO DE ANULACIONES EN SWAGGER
# =========================================================================
def ejecutar_anulaciones(p, registros):
    print("\n🚀 Iniciando procesamiento de Anulaciones en Swagger...")
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 800})
    page = context.new_page()
    
    url_anulaciones = "http://172.19.161.112:9008/swagger-ui/index.html?configUrl=/v3/api-docs/swagger-config#/anulaciones-controller/anulacionesResponse"
    resultados_finales = []
    
    try:
        page.goto(url_anulaciones, wait_until="load", timeout=90000)
        bloque_api = page.locator(".opblock").first
        bloque_api.scroll_into_view_if_needed()
        
        if "is-open" not in (bloque_api.get_attribute("class") or ""):
            bloque_api.locator(".opblock-summary").click()
            page.wait_for_timeout(1000)
            
        btn_try = bloque_api.locator("button.try-out__btn")
        btn_try.wait_for(state="visible")
        if "Cancel" not in btn_try.inner_text():
            btn_try.click()

        columnas_texto = [
            "env_idcliente", "env_idconsumidor", "env_idpagador", 
            "env_telfpagador", "env_telfreceptor", "env_ctareceptor",
            "env_serialoperacion", "res_nrocomprobante", "res_referencia"
        ]

        for i, fila in enumerate(registros):
            id_cliente = str(fila.get('id_cliente', '')).strip()
            if not id_cliente:
                continue
                
            print(f"🔄 [{i+1}/{len(registros)}] Procesando Anulación para Cliente: {id_cliente}")
            marca_tiempo = dateAndHourNow()
            
            # ARMADO DEL PAYLOAD CON EDITABLES DESDE EXCEL Y CAMPOS FIJOS
            payload_anulacion = {
                "idCliente": id_cliente,                                                            # <-- Editable
                "idUsuario": "usuario_beca",                                                        # <-- Fijo
                "idTerminal": "terminal_beca",                                                      # <-- Fijo
                "idCanal": str(fila.get('id_canal', '01')).strip().zfill(2),                        # <-- Editable
                "idConsumidor": str(fila.get('id_consumidor', id_cliente)).strip(),                 # <-- Editable
                "ipOrigen": "111.111.11.1",                                                         # <-- Fijo
                "idPagador": str(fila.get('id_pagador', '')).strip(),                                # <-- Editable
                "telfPagador": limpiar_y_formatear_telefono(fila.get('telf_pagador', '')),         # <-- Editable
                "telfReceptor": limpiar_y_formatear_telefono(fila.get('telf_receptor', '')),       # <-- Editable
                "ctaReceptor": str(fila.get('cta_receptor', '')).strip(),                           # <-- Editable
                "moneda": str(fila.get('moneda', 'VES')).strip(),                                   # <-- Editable
                "serialOperacion": str(fila.get('serial_operacion', '')).strip(),                   # <-- Editable
                "monto": str(fila.get('monto', '0')).strip(),                                       # <-- Editable
                "codBanco": str(fila.get('cod_banco', '0115')).strip()                              # <-- Editable
            }
            
            textarea = bloque_api.locator("textarea.body-param__text")
            textarea.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            textarea.fill(json.dumps(payload_anulacion, indent=2))
            
            bloque_api.locator("button.execute").click()
            
            selector_res = ".responses-wrapper .response .highlight-code pre"
            
            registro_fila = {
                "Fecha, Hora y Día Ejecución": marca_tiempo
            }
            
            datos_enviados_planos = aplanar_dict(payload_anulacion, parent_key='Env')
            registro_fila.update(datos_enviados_planos)
            
            try:
                page.wait_for_selector(selector_res, timeout=15000)
                page.wait_for_timeout(500)
                resultado_texto = page.locator(selector_res).first.inner_text()
                
                data_json = json.loads(resultado_texto)
                datos_recibidos_planos = aplanar_dict(data_json, parent_key='Res')
                registro_fila.update(datos_recibidos_planos)
            except Exception:
                registro_fila["Res_Error_Respuesta"] = resultado_texto if 'resultado_texto' in locals() else "Timeout/Error"

            fila_limpia = {}
            for k, v in registro_fila.items():
                llave_min = str(k).lower()
                val_str = str(v).strip()
                if llave_min in columnas_texto and val_str != "":
                    if not val_str.startswith("'"):
                        val_str = f"'{val_str}"
                fila_limpia[k] = val_str
                
            resultados_finales.append(fila_limpia)
            time.sleep(0.5)

    finally:
        browser.close()
    return resultados_finales


def main():
    try:
        from LeerXlsx import guardar_y_cerrar_excel
        print(f"🔒 Verificando estados de bloqueo para los archivos Excel...")
        guardar_y_cerrar_excel(ARCHIVO_ENTRADA)
        guardar_y_cerrar_excel(ARCHIVO_EDITABLE)
    except Exception as e:
        print(f"ℹ️ No se pudo ejecutar el cierre automático de Excel: {e}")

    datos_a_procesar = leer_datos_entrada_directo(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    if not datos_a_procesar:
        print("⚠️ No hay registros válidos para trabajar. Proceso detenido.")
        return
        
    with sync_playwright() as p:
        registros_finales = ejecutar_anulaciones(p, datos_a_procesar)

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
            print(f"📦 El archivo {os.path.basename(ARCHIVO_EDITABLE)} no existía. Creándolo desde cero...")
            with pd.ExcelWriter(ARCHIVO_EDITABLE, engine='openpyxl') as writer:
                df_nuevos_datos.to_excel(writer, sheet_name=PESTANA_DESTINO, index=False)
        
        print(f"✨ ¡Datos consolidados exitosamente en '{PESTANA_DESTINO}' de '{os.path.basename(ARCHIVO_EDITABLE)}'!")
    else:
        print("⚠️ No se generaron registros para guardar.")
        
    print("🏁 Proceso de Anulaciones Finalizado.")

if __name__ == "__main__":
    main()