import os
import win32com.client
import pandas as pd
import traceback
from collections import namedtuple

from ObtenerRuta import resource_path

# Definición de estructuras
PagoMovil = namedtuple('Pago',['idOrdenante','cuentaOrdenante','telefonoOrdenante','idReceptor','telefonoReceptor','banco','nombreBanco','monto','moneda','canal','idConsumidor','concepto'])
Transferencia = namedtuple('Transferencia',['idEmisor','cuentaEmisor','idReceptor','cuentaReceptor','telefonoReceptor','banco','monto','moneda','canal','idConsumidor','beneficiario','concepto'])



import os
import win32com.client

def guardar_y_cerrar_excel(ruta_archivo):
    """
    Busca si el archivo Excel específico está abierto en la instancia de Excel 
    y lo guarda/cierra.
    """
    # 1. Obtener la ruta absoluta real del archivo
    ruta_absoluta = os.path.abspath(ruta_archivo)
    
    if not os.path.exists(ruta_absoluta):
        print(f"⚠️ El archivo no existe: {ruta_absoluta}")
        return False

    try:
        # Intentar conectar con la instancia de Excel en ejecución
        excel_app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("ℹ️ Excel no está abierto. El archivo está libre.")
        return True

    try:
        archivo_encontrado = False
        # Recorrer todos los libros abiertos
        for workbook in excel_app.Workbooks:
            # Comparar rutas (normalizando a minúsculas para Windows)
            if os.path.abspath(workbook.FullName).lower() == ruta_absoluta.lower():
                print(f"💾 Guardando y cerrando '{os.path.basename(ruta_archivo)}'...")
                workbook.Close(SaveChanges=True)
                archivo_encontrado = True
                break # Salir del bucle al encontrarlo

        if archivo_encontrado:
            return True
        else:
            print(f"ℹ️ El archivo '{os.path.basename(ruta_archivo)}' no estaba abierto en Excel.")
            return True

    except Exception as e:
        print(f"❌ Error al interactuar con Excel: {e}")
        return False

def leerTransferencia(nombreArchivo):
    transferencias = []
    
    # Función auxiliar para limpiar strings y dejar vacío si es NaN o nulo
    def limpiar_str(valor):
        if pd.isna(valor):
            return ""
        val_str = str(valor).strip()
        # Evita que el texto 'nan' o 'None' se guarde como string literales
        if val_str.lower() in ['nan', 'none']:
            return ""
        return val_str

    # Función auxiliar para manejar el monto de forma segura
    def limpiar_monto(valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return 0.0  # O puedes poner None si tu clase lo permite
        try:
            return float(valor)
        except ValueError:
            return 0.0

    try:
        # Leemos la hoja 'Transferencias'
        df = pd.read_excel(nombreArchivo, sheet_name='Transferencias', engine='openpyxl', dtype=str)
        
        if df.empty:
            print(f"⚠️ La hoja 'Transferencias' está vacía.")
            return []

        for _, row in df.iterrows():
            t = Transferencia(
                idEmisor = limpiar_str(row.get('Id Emisor')),
                cuentaEmisor = limpiar_str(row.get('Cuenta Emisor')),
                idReceptor = limpiar_str(row.get('Id Receptor')),
                cuentaReceptor = limpiar_str(row.get('Cuenta Receptor')),
                telefonoReceptor = limpiar_str(row.get('Teléfono Receptor')),
                banco = limpiar_str(row.get('banco')),
                monto = limpiar_monto(row.get('monto')),
                moneda = limpiar_str(row.get('moneda')),
                canal = limpiar_str(row.get('canal')),
                idConsumidor = limpiar_str(row.get('id Consumidor')),
                beneficiario = limpiar_str(row.get('Beneficiario')),
                concepto = limpiar_str(row.get('Concepto'))
            )
            transferencias.append(t)
            
        return transferencias

    except Exception as e:
        print(f"❌ Error al leer la hoja 'Transferencias': {e}")
        return []
    

def leerPagoMovil(nombreArchivo):
    pagos = []
    
    # Función auxiliar para limpiar strings y dejar vacío si es NaN o nulo
    def limpiar_str(valor):
        if pd.isna(valor):
            return ""
        val_str = str(valor).strip()
        # Evita que el texto 'nan' o 'None' se guarde como string literal
        if val_str.lower() in ['nan', 'none']:
            return ""
        return val_str

    # Función auxiliar para manejar el monto de forma segura
    def limpiar_monto(valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return 0.0
        try:
            return float(valor)
        except ValueError:
            return 0.0

    try:
        # Leemos específicamente la hoja 'Pago Móvil'
        df = pd.read_excel(nombreArchivo, sheet_name='Pago Móvil', engine='openpyxl', dtype=str)
        
        if df.empty:
            print(f"⚠️ La hoja 'Pago Móvil' está vacía.")
            return []

        for _, row in df.iterrows():
            pm = PagoMovil(
                idOrdenante = limpiar_str(row.get('Id Ordenante')),
                cuentaOrdenante = limpiar_str(row.get('Cuenta Ordenante')),
                telefonoOrdenante = limpiar_str(row.get('Teléfono Ordenante')),
                idReceptor = limpiar_str(row.get('Id Receptor')),
                telefonoReceptor = limpiar_str(row.get('Teléfono Receptor')),
                banco = limpiar_str(row.get('Banco (código)')),
                nombreBanco = limpiar_str(row.get('Banco (nombre)')),
                monto = limpiar_monto(row.get('monto')),
                moneda = limpiar_str(row.get('moneda')),
                canal = limpiar_str(row.get('canal')),
                idConsumidor = limpiar_str(row.get('id Consumidor')),
                concepto = limpiar_str(row.get('Concepto'))
            )
            pagos.append(pm)
            
        return pagos

    except Exception as e:
        print(f"❌ Error al leer la hoja 'Pago Móvil': {e}")
        traceback.print_exc()
        return []

def leer_datos_lista_movimientos(ruta_excel="Solicitudes.xlsx"):
    """
    Lee los datos de los usuarios desde la pestaña 'Lista de Movimientos',
    normaliza las columnas y formatea las cuentas a 20 dígitos (zfill).
    """
    usuarios = []
    try:
        df = pd.read_excel(ruta_excel, sheet_name='Lista De Movimientos', engine='openpyxl', dtype=str).fillna('')
        
        if df.empty:
            print(f"⚠️ La hoja 'Lista de Movimientos' está vacía.")
            return []

        # Normalizar encabezados (quitar espacios invisibles y pasar a minúsculas)
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        for _, fila in df.iterrows():
            rif_cliente = str(fila.get('rifcliente', '')).strip()
            
            # Saltar filas vacías o con errores de lectura de celda
            if not rif_cliente or rif_cliente.lower() == 'nan':
                continue
                
            id_canal = str(fila.get('idcanal', '')).strip()
            rif_consumidor_txt = str(fila.get('rifconsumidor', '')).strip()
            
            # Formatear la cuenta a los 20 dígitos obligatorios del banco
            cuenta_raw = str(fila.get('cuenta', '')).strip()
            cuenta = cuenta_raw.zfill(20) if cuenta_raw and cuenta_raw.lower() != 'nan' else ""
            
            fecha = str(fila.get('fecha', '')).strip()

            usuarios.append({
                "rif": rif_cliente,
                "canal": id_canal,
                "consumidor": rif_consumidor_txt,
                "cuenta": cuenta,
                "fecha": fecha,
                "posicion": str(fila.get('posicion', '')).strip(),
                "listado": str(fila.get('tipolistado', '')).strip()
            })
            
        return usuarios

    except Exception as e:
        print(f"❌ Error leyendo la hoja 'Lista de Movimientos': {e}")
        return []

def leerConsultaTransferencia(ruta_excel="Datos.xlsx"):
    if not os.path.exists(ruta_excel):
        print(f"❌ Error: No existe el archivo {ruta_excel}")
        return []
    try:
        # Forzamos todo a string para mantener intactos RIFs, cuentas y números largos
        df = pd.read_excel(ruta_excel, dtype=str).fillna('')
        df.columns = [c.strip().lower() for c in df.columns]
        
        usuarios = []
        for _, fila in df.iterrows():
            id_cliente = str(fila.get('idclientee', fila.get('idcliente', ''))).strip()
            
            if not id_cliente or id_cliente.lower() == 'nan':
                continue
                
            usuarios.append({
                "idcliente": id_cliente,
                "idcanal": str(fila.get('idcanal', '')).strip(),
                "idconsumidor": str(fila.get('idconsumidor', '')).strip(),
                "fecha": fila.get('fecha', ''),
                "documentocontraparte": str(fila.get('documentocontraparte', '')).strip(),
                "referencia": str(fila.get('referencia', '')).strip(),
                "cuenta": str(fila.get('cuenta', '')).strip(),
                "bancocontraparte": str(fila.get('bancocontraparte', '')).strip(),
                "monto": str(fila.get('monto', '')).strip(),
                "telefonocontraparte": str(fila.get('telefonocontraparte', fila.get('telefono', ''))).strip(),
                "tipoconsulta": str(fila.get('tipoconsulta', '')).strip(),
                "otrobanco": str(fila.get('otrobanco', '')).strip(),
                "posicioninicial": str(fila.get('posicioninicial', fila.get('posicion', ''))).strip()
            })
        return usuarios
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return []
    
def leerConsultaLimitesPagoMovil():
    # Usamos la ruta especificada
    ruta_archivo = resource_path("Solicitudes.xlsx")

    if not os.path.exists(ruta_archivo):
        print(f"❌ No se encontró el archivo de origen: {ruta_archivo}")
        return []

    try:
        # Leer el archivo
        df = pd.read_excel(ruta_archivo, sheet_name='Consulta Límites Pago Móvil', engine='openpyxl', dtype=str).fillna('')
        
        if df.empty:
            print("⚠️ La hoja 'Consulta Límites Pago Móvil' está vacía.")
            return []
            
        # Normalizar nombres de columnas
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        # Convertir a lista de diccionarios
        registros_finales = df.to_dict('records')
        
        return registros_finales
        
    except Exception as e:
        print(f"❌ Error al intentar leer la hoja 'Consulta Límites Pago Móvil': {e}")
        return []

def leerConsultaLimitesTransferencia():
    ruta_archivo = resource_path("Solicitudes.xlsx")

    if not os.path.exists(ruta_archivo):
        print(f"❌ No se encontró el archivo de origen: {ruta_archivo}")
        return []

    try:
        df = pd.read_excel(ruta_archivo, sheet_name="Consulta Límite Transferencias", engine='openpyxl', dtype=str).fillna('')
        
        if df.empty:
            print("⚠️ La hoja de Excel está vacía.")
            return []
            
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        registros_finales = df.to_dict('records')
        
        return registros_finales
        
    except Exception as e:
        print(f"❌ Error al intentar leer el Excel: {e}")
        return []

def leerExcel(nombreHoja, archivo="Solicitudes.xlsx"):
    """
    Función única para leer cualquier hoja de 'Solicitudes.xlsx'.
    """
    if not os.path.exists(archivo):
        print(f"❌ Error: No existe el archivo {archivo}")
        return []

    try:
        # 1. Leer el archivo, forzando todo a string para mantener integridad de datos
        df = pd.read_excel(archivo, sheet_name=nombreHoja, engine='openpyxl', dtype=str).fillna('')
        
        if df.empty:
            print(f"⚠️ La hoja '{nombreHoja}' está vacía.")
            return []

        # 2. Normalizar encabezados (quitar espacios y pasar a minúsculas)
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        # 3. Retornar como lista de diccionarios (formato estándar)
        return df.to_dict('records')

    except Exception as e:
        print(f"❌ Error al leer la hoja '{nombreHoja}': {e}")
        return []


# # Ejemplo de uso
# guardar_y_cerrar_excel("Solicitudes.xlsx")
# guardar_y_cerrar_excel(resource_path("Resultados.xlsx"))