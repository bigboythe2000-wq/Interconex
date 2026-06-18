from playwright.sync_api import sync_playwright
from LogicaHora import dateAndHourPlusTen, dateAndHourNow
from LogicaDeEntrada import jsonFillDataPay, swaggerFillDataPay
from LeerXlsx import leerPagoMovil
from EscrituraDatosExcel import escribirPagoMovil
from collections import namedtuple
from TokenGenerator import TokenGenerator
from ObtenerRuta import resource_path

def hacerPagoMovil():
    print("🚀 Iniciando proceso de automatización de Pagos Móviles...")
    
    # Rutas de archivos
    nombreArchivo = resource_path("Solicitudes.xlsx")
    archivo_resultados = resource_path("Resultados.xlsx")
    
    # 1. Lectura de datos inicial
    try:
        print(f"📂 Leyendo archivo de origen: {nombreArchivo}")
        pagos = leerPagoMovil(nombreArchivo)
        if not pagos:
            print("⚠️ El archivo está vacío o no contiene solicitudes.")
            return
        print(f"✅ Se cargaron {len(pagos)} solicitudes exitosamente.")
    except Exception as e:
        print(f"❌ Error crítico al leer el Excel: {e}")
        return

    # 2. Generación de Tokens
    try:
        print("🔐 Generando tokens de seguridad...")
        clave_privada = """-----BEGIN PRIVATE KEY-----
MIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQC7VJTUt9Us8cKj
MzEfYyjiWA4R4/M2bS1GB4t7NXp98C3SC6dVMvDuictGeurT8jNbvJZHtCSuYEvu
NMoSfm76oqFvAp8Gy0iz5sxjZmSnXyCdPEovGhLa0VzMaQ8s+CLOyS56YyCFGeJZ
qgtzJ6GR3eqoYSW9b9UMvkBpZODSctWSNGj3P7jRFDO5VoTwCQAWbFnOjDfH5Ulg
p2PKSQnSJP3AJLQNFNe7br1XbrhV//eO+t51mIpGSDCUv3E0DDFcWDTH9cXDTTlR
ZVEiR2BwpZOOkE/Z0/BVnhZYL71oZV34bKfWjQIt6V/isSMahdsAASACp4ZTGtwi
VuNd9tybAgMBAAECggEBAKTmjaS6tkK8BlPXClTQ2vpz/N6uxDeS35mXpqasqskV
laAidgg/sWqpjXDbXr93otIMLlWsM+X0CqMDgSXKejLS2jx4GDjI1ZTXg++0AMJ8
sJ74pWzVDOfmCEQ/7wXs3+cbnXhKriO8Z036q92Qc1+N87SI38nkGa0ABH9CN83H
mQqt4fB7UdHzuIRe/me2PGhIq5ZBzj6h3BpoPGzEP+x3l9YmK8t/1cN0pqI+dQwY
dgfGjackLu/2qH80MCF7IyQaseZUOJyKrCLtSD/Iixv/hzDEUPfOCjFDgTpzf3cw
ta8+oE4wHCo1iI1/4TlPkwmXx4qSXtmw4aQPz7IDQvECgYEA8KNThCO2gsC2I9PQ
DM/8Cw0O983WCDY+oi+7JPiNAJwv5DYBqEZB1QYdj06YD16XlC/HAZMsMku1na2T
N0driwenQQWzoev3g2S7gRDoS/FCJSI3jJ+kjgtaA7Qmzlgk1TxODN+G1H91HW7t
0l7VnL27IWyYo2qRRK3jzxqUiPUCgYEAx0oQs2reBQGMVZnApD1jeq7n4MvNLcPv
t8b/eU9iUv6Y4Mj0Suo/AU8lYZXm8ubbqAlwz2VSVunD2tOplHyMUrtCtObAfVDU
AhCndKaA9gApgfb3xw1IKbuQ1u4IF1FJl3VtumfQn//LiH1B3rXhcdyo3/vIttEk
48RakUKClU8CgYEAzV7W3COOlDDcQd935DdtKBFRAPRPAlspQUnzMi5eSHMD/ISL
DY5IiQHbIH83D4bvXq0X7qQoSBSNP7Dvv3HYuqMhf0DaegrlBuJllFVVq9qPVRnK
xt1Il2HgxOBvbhOT+9in1BzA+YJ99UzC85O0Qz06A+CmtHEy4aZ2kj5hHjECgYEA
mNS4+A8Fkss8Js1RieK2LniBxMgmYml3pfVLKGnzmng7H2+cwPLhPIzIuwytXywh
2bzbsYEfYx3EoEVgMEpPhoarQnYPukrJO4gwE2o5Te6T5mJSZGlQJQj9q4ZB2Dfz
et6INsK0oG8XVGXSpQvQh3RUYekCZQkBBFcpqWpbIEsCgYAnM3DQf3FJoSnXaMhr
VBIovic5l0xFkEHskAjFTevO86Fsz1C2aSeRKSqGFoOQ0tmJzBEs1R6KqnHInicD
TQrKhArgLXX4v3CddjfTRJkFWDbE/CkvKZNOrcf1nhaGCPspRJj2KUkj1Fhl9Cnc
dn/RsYEONbwQSjIfMPkvxF+8HQ==
-----END PRIVATE KEY-----"""
        
        generator = TokenGenerator(clave_privada)
        Transfer = namedtuple('Transfer', ['pago','data'])
        
        jsonDataArray = []
        for pago in pagos:
            token = generator.generate_token(jsonFillDataPay(dateAndHourPlusTen(), pago))
            jsonDataArray.append(Transfer(pago, token))
            
    except Exception as e:
        print(f"❌ Error en la generación de tokens: {e}")
        return

    # 3. Ejecución en Navegador
    try:
        with sync_playwright() as p:
            print("🌐 Iniciando navegador...")
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(ignore_https_errors=True, user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            swaggerPage = context.new_page()
            
            swaggerPage.goto("http://172.19.161.136:9006/swagger-ui/index.html#/realizar-pago-sw-controller-v-2/pagoResponse")
            swaggerPage.click("css=.btn") # Botón Try it out

            for i, jsonFinalData in enumerate(jsonDataArray, 1):
                print(f"\n--- Procesando registro {i}/{len(jsonDataArray)} ---")
                print(f"👤 Ordenante: {jsonFinalData.pago.idOrdenante}")
                
                try:
                    # Llenado de datos
                    swaggerPage.fill("css=.body-param__text", swaggerFillDataPay(jsonFinalData.data, jsonFinalData.pago))
                    
                    # Ejecución y espera de respuesta
                    with swaggerPage.expect_response("**/api/qa/v2/realizar-pago-movil", timeout=120000) as response_info:
                        swaggerPage.locator("css=.execute").click()
                    
                    response = response_info.value
                    json_resp = response.json()
                                    
                    # Log de resultados
                    print(f"✅ Respuesta: {json_resp['resultado']['descripcion']}")
                    print(f"📑 Referencia: {json_resp['datos']['referencia']}")
                    
                    # Escritura en Excel
                    escribirPagoMovil(jsonFinalData.pago, dateAndHourNow(), 
                                      json_resp['resultado']['descripcion'], 
                                      json_resp['datos']['referencia'], 
                                      archivo_resultados)
                    
                    swaggerPage.wait_for_selector("css=.microlight:nth-child(3)", state="attached")
                    swaggerPage.wait_for_load_state("networkidle")
                    
                except Exception as e:
                    print(f"❌ Error procesando el pago {i}: {e}")
                    continue 
                except PermissionError:
                    print(f"❌ ERROR: El archivo {archivo_resultados} está abierto. Cérralo para continuar.")
                    # Puedes poner un input aquí para pausar y permitir que el usuario lo cierre
                    input("Presiona Enter una vez que hayas cerrado el archivo...")
                    # Intentar escribir de nuevo
                    escribirPagoMovil(jsonFinalData.pago, dateAndHourNow(), 
                      json_resp['resultado']['descripcion'], 
                      json_resp['datos']['referencia'], 
                      archivo_resultados)

            browser.close()
            print("\n🏁 Proceso finalizado correctamente.")

    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    hacerPagoMovil()