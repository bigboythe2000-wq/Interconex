import sys
import os

def resource_path(relative_path):
    """ Obtiene la ruta absoluta al recurso """
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable, busca en el directorio del ejecutable
        base_path = os.path.dirname(sys.executable)
    else:
        # Si es script, busca en el directorio del script
        base_path = os.path.abspath(os.path.dirname(__file__))
    
    return os.path.join(base_path, relative_path)

