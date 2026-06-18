import jwt
from cryptography.hazmat.primitives import serialization
from LeerXlsx import leerPagoMovil
from ObtenerRuta import resource_path
from LogicaDeEntrada import jsonFillDataPay
from LogicaHora import dateAndHourPlusTen
class TokenGenerator:
    def __init__(self, private_key_str: str, password: bytes = None):
        """
        Inicializa el generador cargando la clave privada.
        """
        try:
            # Convertir a bytes y cargar la clave
            key_bytes = private_key_str.encode('utf-8')
            self.private_key = serialization.load_pem_private_key(
                key_bytes,
                password=password
            )
        except Exception as e:
            raise ValueError(f"Error al cargar la clave privada: {e}")

    def generate_token(self, payload: dict,) -> str:
        """
        Genera el token agregando automáticamente el tiempo de expiración.
        """
        # Clonamos el payload para no modificar el original
        data = payload.copy()
        
        # Generar el token
        return jwt.encode(data, self.private_key, algorithm="RS256")
