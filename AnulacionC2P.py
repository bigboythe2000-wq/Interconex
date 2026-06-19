from LeerXlsx import leerExcel
from LogicaDeEntrada import escribirPayload
from playwright.sync_api import sync_playwright
from EscrituraDatosExcel import guardar_payload_en_excel
import json
# API http://172.19.161.112:9008/api/qa/v1/anulaciones-cobros
def AnulacionC2P():
    anulaciones = leerExcel("Anulación C2P")
    with sync_playwright() as p: 
        print("🌐 Iniciando navegador...")
        # Lanzar navegador
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
       
        # Ir a la página
        print("🔗 Accediendo a la URL de transferencias...")
        page.goto("http://172.19.161.112:9008/swagger-ui/index.html?configUrl=/v3/api-docs/swagger-config#/anulaciones-controller/anulacionesResponse")

        # Esperar a que el botón de 'Try it out' esté visible
        page.wait_for_selector("css=.btn", timeout=10000)
        page.click("css=.btn")

        for idx, anulacion in enumerate(anulaciones, 1):
            
            print(f"📝 Procesando anulación {idx}/{len(anulacion)}: {anulacion["serialoperacion"]}")
            estructuraPayload = page.locator("css=.body-param__text").input_value()
            estructuraFinal = json.loads(estructuraPayload)
            # 1. Llamas a tu función, que devuelve un diccionario
            payload_dict = escribirPayload(estructuraFinal, anulacion)
            # 2. Conviertes ese diccionario a una cadena (string) legible
            payload_str = json.dumps(payload_dict, indent=4)
            # 2. Conviertes ese diccionario a una cadena (st)
            page.fill("css=.body-param__text", payload_str)


            descripcion = "Error en el proceso"

            try:
                with page.expect_response("http://172.19.161.112:9008/api/qa/v1/anulaciones-cobros", timeout=120000) as response_info:
                    page.click("css=.execute")
                
                response = response_info.value
                json_resp = response.json()
                
                # 1. Obtener estado de la operación
                datos = json_resp.get('datos', {})
                anulacion_ok = datos.get('anulacionRealizada', False)
                
                # 2. Obtener descripción (manejo de caso nulo)
                resultado = json_resp.get('resultado', {})
                descripcion = resultado.get('descripcion')
                
                # Si la descripción es nula, intentamos obtener el código o un mensaje por defecto
                if descripcion is None:
                    codigo = resultado.get('codigo', 'N/A')
                    descripcion = f"Sin descripción (Código: {codigo})"

                print(f"✅ Respuesta recibida. ¿Anulación realizada?: {anulacion_ok}")

                        # --- AQUÍ GUARDAMOS EL RESULTADO EN EXCEL ---
                # Creamos un diccionario con la información que queremos registrar
                payload_registro = {
                    "Estado": "Éxito" if anulacion_ok else "Fallido",
                    "AnulacionRealizada": anulacion_ok,
                    "Descripcion": descripcion,
                    "Codigo": resultado.get('codigo', 'N/A')
                }
                
                # Llamada a tu función (asegúrate de que el archivo exista o la ruta sea correcta)
                guardar_payload_en_excel(payload_registro, "Resultados.xlsx", "Resultados Anulacion")

            except Exception as e:
                print(f"⚠️ Error en la operación: {e}")
                descripcion = f"Falla: {str(e)}"
                
                # Guardamos el error también en el Excel para tener trazabilidad
                payload_error = {
                    "Estado": "Error",
                    "AnulacionRealizada": False,
                    "Descripcion": descripcion,
                    "Codigo": "ERR"
                }
                guardar_payload_en_excel(payload_error, "Resultados.xlsx", "Resultados Anulacion")
                
                try:
                    page.click("css=.execute")
                except:
                    pass
                        
                        
                    # Cerrar navegador
                    print("✅ Proceso finalizado correctamente.")
                    browser.close()
    
   


if __name__ == "__main__": 
    AnulacionC2P()