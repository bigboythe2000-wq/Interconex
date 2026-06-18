from playwright.sync_api import sync_playwright
from LogicaDeEntrada import buildTransfer
from LogicaHora import dateAndHourNow
from EscrituraDatosExcel import escribirTransferencia
from LeerXlsx import leerTransferencia
from ObtenerRuta import resource_path
import os
import sys
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.path.dirname(sys.executable), "ms-playwright")
 
def transferir():
    print("🚀 Iniciando proceso de transferencias...")
    
    # Cargamos la lista de cuentas desde el Excel
    ruta_solicitudes = resource_path('Solicitudes.xlsx')
    Transferencias = leerTransferencia(ruta_solicitudes)
    
    if not Transferencias:
        print("⚠️ No hay transferencias para procesar. Verifica el archivo Excel.")
        return
    
    # Iniciar Playwright
    with sync_playwright() as p: 
        print("🌐 Iniciando navegador...")
        # Lanzar navegador
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
       
        # Ir a la página
        print("🔗 Accediendo a la URL de transferencias...")
        page.goto("http://172.19.161.137:7005/swagger-ui/index.html?urls.primaryName=-%20Version%202%20-#/transferencias-inmediatas-controller-v-2/trasferenciasV2")

        # Esperar a que el botón de 'Try it out' esté visible
        page.wait_for_selector("css=.btn", timeout=10000)
        page.click("css=.btn")

        for idx, transferencia in enumerate(Transferencias, 1):
            print(f"📝 Procesando transferencia {idx}/{len(Transferencias)}: {transferencia.beneficiario} por {transferencia.monto} {transferencia.moneda}")
            
            page.fill("css=.body-param__text", buildTransfer(transferencia))

            reference = "No se obtuvo referencia"
            descripcion = "Error en el proceso"

            try:
                with page.expect_response("http://172.19.161.137:7005/api/qa/v2/transferenciasInmediatas/enviar", timeout=120000) as response_info:
                    page.click("css=.execute")
                response = response_info.value
                json_resp = response.json()
                
                if 'datosTransferenciainmediata' in json_resp:
                    reference = json_resp['datosTransferenciainmediata'].get('referencia', reference)
                
                if 'resultado' in json_resp:
                    descripcion = json_resp['resultado'].get('descripcion', descripcion)
                elif 'descripcion' in json_resp:
                    descripcion = json_resp['descripcion']
                    
            except Exception as e:
                print(f"⚠️ Reintentando transferencia {idx} debido a: {e}")
                try:
                    # Intento de recuperación o segundo click
                    page.click("css=.execute")
                    # No esperamos respuesta aquí por si el error fue de red pero la API procesó
                except:
                    pass
                descripcion = f"Falla: {str(e)}"

            escribirTransferencia(transferencia, dateAndHourNow(), 
                                descripcion if descripcion != "Error en el proceso" else "Parada", 
                                reference, descripcion, resource_path('Resultados.xlsx'))
            
        # Cerrar navegador
        print("✅ Proceso finalizado correctamente.")
        browser.close()
 
if __name__ == "__main__": 
    transferir()
