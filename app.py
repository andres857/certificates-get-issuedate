import os
import logging
from enum import Enum
from readDocs import OfficeDocumentExtractor
from utils import extraer_id_archivo, renombrar_archivo_con_fechas
from excel import create_excel_template, insert_certificate_data

# Enum simple para las extensiones prohibidas
class ArchivoProhibido(Enum):
    INI = '.ini'
    EXE = '.exe'
    HTML = '.html'

def leer_todos_certificates():
    extractor = OfficeDocumentExtractor()
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    carpeta_certificates = os.path.join(directorio_actual, 'certificates')
    
    if not os.path.exists(carpeta_certificates):
        print(f"El directorio {carpeta_certificates} no existe")
        return
    
    extensiones_prohibidas = {ext.value for ext in ArchivoProhibido}
    
    for raiz, dirs, archivos in os.walk(carpeta_certificates):
        # Obtenemos el nombre de la carpeta actual
        nombre_carpeta = os.path.basename(raiz)
        path_folder = os.path.join(raiz)

        path_report = create_excel_template(path_folder)
        
        archivos_permitidos = [
            archivo for archivo in archivos 
            if os.path.splitext(archivo)[1].lower() not in extensiones_prohibidas
        ]
        
        print(f"\nLeyendo archivos en carpeta: {nombre_carpeta}")
        
        for archivo in archivos_permitidos:
            path_file = os.path.join(raiz, archivo)
            inference_response = extractor.extract_content(path_file)
            issue_date = inference_response.get('issue_date')
            expiration_date = inference_response.get('expiration_date')
            # Llamamos a la función con las fechas correspondientes
            if issue_date:
                renombrar_archivo_con_fechas(
                    ruta_original=path_file,
                    issue_date=issue_date,
                    expiration_date=expiration_date
                )
            else:
                inference_response['name'] = archivo
                insert_certificate_data(path_report, inference_response)
                
if __name__ == "__main__":
    leer_todos_certificates()