def buildTransfer(Transferencia):

    # Usamos f''' para poder meter variables entre llaves {variable}
    # Convertimos monto_valor a string explícitamente dentro de la f-string
    pagador_rif = Transferencia.idEmisor
    idConsumidor = Transferencia.idConsumidor
    canal = Transferencia.canal
    cuentaPagadora = Transferencia.cuentaEmisor
    numeroCuenta = Transferencia.cuentaReceptor
    codigoBanco = Transferencia.banco   
    tlfReceptor = Transferencia.telefonoReceptor
    rif = Transferencia.idReceptor
    monto_valor = Transferencia.monto
    moneda = Transferencia.moneda
    beneficiario = Transferencia.beneficiario
    concepto = Transferencia.concepto
    
    return f'''{{
    "datosPeticion": {{
        "idCliente": "{pagador_rif}",
        "idSesion": "20240101102545",
        "idCanal": {canal},
        "idUsuario": "E123456",
        "idTerminal": "agencia1",
        "ip": "172.1.1.1",
        "idConsumidor": "{idConsumidor}"
    }},
    "transferenciaInmediata": {{
        "ctaPagadora": "{cuentaPagadora}",
        "ctaReceptora": "{numeroCuenta}",
        "codigobancoReceptor": "{codigoBanco}",
        "telefonoReceptor": "{tlfReceptor}",
        "idReceptor": "{rif}",
        "monto": {monto_valor},
        "moneda": "{moneda}",
        "nombreBeneficiario": "{beneficiario}",
        "concepto": "{concepto}"
    }}
}}'''

def jsonFillDataPay(date,pagoMovil):    
    return {
        "exp": date,
        "idBeneficiario": pagoMovil.idReceptor,
        "telefonoBeneficiario": pagoMovil.telefonoReceptor,
        "telefonoEmisor": pagoMovil.telefonoOrdenante,
        "cuentaEmisor": pagoMovil.cuentaOrdenante,
        "moneda": pagoMovil.moneda,
        "monto": pagoMovil.monto
    }

def swaggerFillDataPay(jsonData, pagoMovil):
  # Usamos {variable} para que Python inserte el valor real
  return f'''{{
  "idCliente": "{pagoMovil.idOrdenante}",
  "idUsuario": "usuario_beca",
  "idTerminal": "terminal_beca",
  "idCanal": "{pagoMovil.canal}",
  "codigoBanco": "{pagoMovil.banco}",
  "nombreBanco": "{pagoMovil.nombreBanco}",
  "concepto": "{pagoMovil.concepto}",
  "idOperacion": "2123455",
  "fechaOperacion": "17/09/2020 16:00:00.111",
  "datosAutorizados": "{jsonData}",     
  "envioEmailEmisor": true,
  "envioEmailReceptor": true,
  "idConsumidor": "{pagoMovil.idConsumidor}",
  "ipOrigen": "111.111.11.1"
}}'''

def payLoadListaMovimientos(datos):
    def entero_o_vacio(valor):
        if valor == "" or str(valor).lower() == 'nan':
            return ""
        try:
            return int(float(valor))
        except:
            return str(valor)

    canal_valor = datos["canal"]
    if canal_valor and canal_valor.lower() != 'nan':
        try:
            canal_valor = str(int(float(canal_valor))).zfill(2)
        except:
            pass

    return {
        "idCliente": datos["rif"],
        "idCanal": canal_valor,
        "idConsumidor": datos["consumidor"],
        "ipOrigen": "111.111.11.1",
        "idUsuario": "usuario_beca",
        "idTerminal": "terminal_beca",
        "cuenta": datos["cuenta"],
        "fecha": datos["fecha"],
        "posicionInicial": entero_o_vacio(datos["posicion"]),
        "tipoListado": entero_o_vacio(datos["listado"])
    }

def escribirPayload(payLoadEstructura, payLoadDatos):
    """
    Actualiza la estructura mapeando claves ignorando mayúsculas/minúsculas.
    """
    nuevo_payload = payLoadEstructura.copy()
    
    # Creamos un diccionario auxiliar de los datos recibidos con las llaves en minúsculas
    # para que la comparación no falle
    datos_normalizados = {str(k).lower(): v for k, v in payLoadDatos.items()}
    
    for clave in nuevo_payload:
        clave_min = str(clave).lower()
        
        # Si la versión en minúscula de la clave existe en los datos
        if clave_min in datos_normalizados:
            # Actualizamos con el valor del Excel
            nuevo_payload[clave] = datos_normalizados[clave_min]
            
    return nuevo_payload

