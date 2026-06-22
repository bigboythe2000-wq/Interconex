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
PESTANA_ENTRADA = "Cobro_comercio_C2P"       

ARCHIVO_EDITABLE = obtener_ruta_absoluta("Resultados.xlsx")
PESTANA_DESTINO = "Cobro_Comercio_C2P_Resultados" 
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
    """Limpia y normaliza el teléfono estrictamente al formato 58XXXXXXXXX"""
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
#  FASE 1: CONSULTAR Y EXTRAER LA OTP (CON MONTO DINÁMICO VALIDADO)
# =========================================================================
def obtener_otp_swagger(p, registros):
    print("\n🚀 [FASE 1] Iniciando extracción de OTPs de Swagger...")
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 800})
    page = context.new_page()
    
    url_otp = "http://172.19.160.160:7070/swagger-ui/index.html?urls.primaryName=-%20Version%202%20-#/otp-controller"
    
    try:
        page.goto(url_otp, wait_until="load", timeout=90000)
        bloque_api = page.locator(".opblock").first
        bloque_api.scroll_into_view_if_needed()
        
        if "is-open" not in (bloque_api.get_attribute("class") or ""):
            bloque_api.locator(".opblock-summary").click()
            page.wait_for_timeout(1000)
            
        btn_try = bloque_api.locator("button.try-out__btn")
        btn_try.wait_for(state="visible")
        if "Cancel" not in btn_try.inner_text():
            btn_try.click()

        for i, fila in enumerate(registros):
            id_pagador = str(fila.get('id_pagador', '')).strip()
            telf_pagador = limpiar_y_formatear_telefono(fila.get('telf_pagador', ''))
            
            # Captura y conversión del monto dinámico para la OTP
            monto_raw = str(fila.get('monto', '1.22')).strip()
            try:
                monto_otp = float(monto_raw) if monto_raw else 1.22
            except ValueError:
                monto_otp = 1.22
            
            canal_raw = str(fila.get('id_canal', '1')).strip()
            canal_otp = str(int(canal_raw)) if canal_raw.isdigit() else "1"
            
            if not id_pagador:
                continue
                
            print(f"🔍 [{i+1}] Solicitando OTP para Pagador: {id_pagador} | Teléfono: {telf_pagador} | Monto: {monto_otp}")
            
            payload_otp = {
                "idSesion": "20210101130101",
                "canal": canal_otp,
                "idTerminal": "123BE",
                "idConsumidor": "J1234",
                "idCliente": "J5678",
                "idUsuario": "J876543210",
                "codUsuario": "E7654",
                "ip": "172.1.1.1",
                "datosOTP": {
                    "nroIdCliente": id_pagador,
                    "telefono": telf_pagador, # Ahora va formateado con 58...
                    "tipo": 0,
                    "monto": monto_otp,     # <-- Tomando el monto real del Excel de forma numérica
                    "origen": "03"
                }
            }
            
            textarea = bloque_api.locator("textarea.body-param__text")
            textarea.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            textarea.fill(json.dumps(payload_otp, indent=2))
            
            bloque_api.locator("button.execute").click()
            
            selector_res = ".responses-wrapper .response .highlight-code pre"
            try:
                page.wait_for_selector(selector_res, timeout=15000)
                page.wait_for_timeout(3000) 
                res_txt = page.locator(selector_res).first.inner_text()
                data_json = json.loads(res_txt)
                
                otp_detectada = data_json.get("datos", {}).get("otp", "") or data_json.get("otp", "")
                
                if otp_detectada:
                    fila['otp_encontrada'] = str(otp_detectada)
                    fila['estado_otp'] = "OTP_ENCONTRADA"
                    print(f"🔑 OTP Detectada con éxito: {otp_detectada}")
                else:
                    fila['otp_encontrada'] = ""
                    fila['estado_otp'] = "Sin OTP en JSON"
                    print("⚠️ Respuesta recibida pero sin campo 'otp'")
            except Exception:
                fila['otp_encontrada'] = ""
                fila['estado_otp'] = "Error respuesta OTP"
                print(f"❌ No se pudo extraer la OTP de la respuesta.")
                
            time.sleep(0.5)
            
    finally:
        browser.close()
    return registros


# =========================================================================
#  FASE 2: EJECUTAR EL COBRO COMERCIO C2P (CON TELÉFONOS EN FORMATO 58...)
# =========================================================================
def ejecutar_cobro_c2p(p, registros):
    print("\n🚀 [FASE 2] Iniciando procesamiento de Cobros Comercio C2P...")
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 800})
    page = context.new_page()
    
    url_cobro = "http://172.19.160.141:9007/swagger-ui/index.html?configUrl=/#/cobros-controller/cobrosResponse"
    resultados_finales = []
    
    try:
        page.goto(url_cobro, wait_until="load", timeout=90000)
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
            "env_telfpagador", "env_telfreceptor", "env_ctareceptor", "env_otp",
            "res_nrocomprobante", "res_referencia"
        ]

        for i, fila in enumerate(registros):
            id_cliente = str(fila.get('id_cliente', '')).strip()
            if not id_cliente:
                continue
                
            otp_usar = str(fila.get('otp_encontrada', '')).strip()
            estado_previo = str(fila.get('estado_otp', '')).strip()
            
            print(f"🔄 [{i+1}/{len(registros)}] Procesando Cobro C2P para Comercio: {id_cliente}")
            marca_tiempo = dateAndHourNow()
            
            if not otp_usar or "Error" in estado_previo or "Sin" in estado_previo:
                print(f"⏭️ Saltando Cobro C2P debido a falta de OTP válida.")
                fila_res = {
                    "Fecha, Hora y Día Ejecución": marca_tiempo,
                    "Env_idCliente": id_cliente,
                    "Env_idPagador": str(fila.get('id_pagador', '')).strip(),
                    "Env_otp": "",
                    "Estado_Paso_Previa_OTP": estado_previo,
                    "Res_Resultado_Cobro": "SALTADO_SIN_OTP"
                }
                resultados_finales.append(fila_res)
                continue

            # JSON estructurado para el Cobro C2P - Ambos teléfonos formateados con 58...
            payload_cobro = {
                "idCliente": id_cliente,
                "idUsuario": "usuario_beca",
                "idTerminal": "terminal_beca",
                "idCanal": str(fila.get('id_canal', '01')).strip().zfill(2),
                "idConsumidor": str(fila.get('id_consumidor', id_cliente)).strip(),
                "ipOrigen": "111.111.11.1",
                "idPagador": str(fila.get('id_pagador', '')).strip(),
                "telfPagador": limpiar_y_formatear_telefono(fila.get('telf_pagador', '')),    # <-- Formato 58...
                "telfReceptor": limpiar_y_formatear_telefono(fila.get('telf_receptor', '')),  # <-- Formato 58...
                "ctaReceptor": str(fila.get('cta_receptor', '')).strip(),
                "moneda": str(fila.get('moneda', 'VES')).strip(),
                "concepto": str(fila.get('concepto', 'pago material')).strip(),
                "monto": str(fila.get('monto', '10.5')).strip(),
                "otp": otp_usar, 
                "codBanco": str(fila.get('cod_banco', '0115')).strip(),
                "envioEMailPagador": True,
                "envioEMailReceptor": True
            }
            
            textarea = bloque_api.locator("textarea.body-param__text")
            textarea.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            textarea.fill(json.dumps(payload_cobro, indent=2))
            
            bloque_api.locator("button.execute").click()
            
            selector_res = ".responses-wrapper .response .highlight-code pre"
            
            registro_fila = {
                "Fecha, Hora y Día Ejecución": marca_tiempo,
                "Estado_Paso_Previa_OTP": estado_previo
            }
            
            datos_enviados_planos = aplanar_dict(payload_cobro, parent_key='Env')
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
    datos_a_procesar = leer_datos_entrada_directo(ARCHIVO_ENTRADA, PESTANA_ENTRADA)
    if not datos_a_procesar:
        print("⚠️ No hay registros válidos para trabajar. Proceso detenido.")
        return
        
    with sync_playwright() as p:
        datos_con_otp = obtener_otp_swagger(p, datos_a_procesar)
        registros_finales = ejecutar_cobro_c2p(p, datos_con_otp)

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
        
    print("🏁 Proceso de Cobro Comercio C2P Finalizado.")

if __name__ == "__main__":
    main()