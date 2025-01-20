import os
from PyPDF2 import PdfReader
import logging

# Configurar logging para ignorar warnings
logging.getLogger('PyPDF2').setLevel(logging.ERROR)

def leer_todos_certificates():
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    carpeta_certificates = os.path.join(directorio_actual, 'certificates')
    
    if not os.path.exists(carpeta_certificates):
        print(f"El directorio {carpeta_certificates} no existe")
        return
    
    # Recorrer todas las subcarpetas
    for raiz, dirs, archivos in os.walk(carpeta_certificates):
        certificates = [archivo for archivo in archivos]
        
        if certificates:
            # Mostrar en qué carpeta estamos
            # print(f"\n Leyendo certificates en: {certificates}"
            
            for certificate in certificates:
                ruta_completa = os.path.join(raiz, certificate)
                try:
                    with open(ruta_completa, 'rb') as archivo_pdf:
                        lector = PdfReader(archivo_pdf, strict=False)
                        num_paginas = len(lector.pages)
                        
                        # print(f"\nArchivo: {certificate}")
                        # print(f"Número de páginas: {num_paginas}")
                        
                        if (num_paginas == 1):
                            print('leer certificate')
                            
                except Exception as e:
                    print(f"Error al leer el archivo {certificate}: {str(e)}")

if __name__ == "__main__":
    leer_todos_certificates()