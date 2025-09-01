import os
import shutil
from pathlib import Path
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

def crear_carpeta_duplicados(path_folder):
    """
    Crea la carpeta 'duplicates' dentro del directorio de certificados
    """
    carpeta_duplicados = os.path.join(path_folder, 'duplicates')
    if not os.path.exists(carpeta_duplicados):
        os.makedirs(carpeta_duplicados)
        print(f"📁 Carpeta de duplicados creada: {carpeta_duplicados}")
    return carpeta_duplicados


def generar_clave_certificado(inference_response):
    """
    Genera una clave única basada en identificación, issue_date y expiration_date
    """
    identification = inference_response.get('identification')
    issue_date = inference_response.get('issue_date')
    expiration_date = inference_response.get('expiration_date')
    
    if not identification:
        return None
        
    # Crear clave única - manejar casos donde las fechas pueden ser None
    issue_str = issue_date if issue_date else "sin_fecha_emision"
    expiration_str = expiration_date if expiration_date else "sin_fecha_expiracion"
    
    return f"{identification}|{issue_str}|{expiration_str}"

def archivo_ya_procesado(nombre_archivo):
    """
    Verifica si un archivo ya fue procesado buscando 'issueddate' en el nombre
    """
    partes = nombre_archivo.split('_')
    return any('issueddate' in parte for parte in partes)


def mover_a_duplicados(archivo_path, carpeta_duplicados, razon=""):
    """
    Mueve un archivo a la carpeta de duplicados
    """
    try:
        if not os.path.exists(archivo_path):
            print(f"❌ Archivo no existe: {archivo_path}")
            return False
            
        nombre_archivo = os.path.basename(archivo_path)
        destino = os.path.join(carpeta_duplicados, nombre_archivo)
        
        # Si ya existe en duplicados, agregar sufijo numérico
        contador = 1
        nombre_base, extension = os.path.splitext(nombre_archivo)
        while os.path.exists(destino):
            nuevo_nombre = f"{nombre_base}_dup{contador}{extension}"
            destino = os.path.join(carpeta_duplicados, nuevo_nombre)
            contador += 1
        
        shutil.move(archivo_path, destino)
        print(f"🗂️  ✅ MOVIDO a duplicados: {nombre_archivo} → {os.path.basename(destino)} {razon}")
        return True
        
    except Exception as e:
        print(f"❌ Error moviendo {archivo_path} a duplicados: {e}")
        return False
    
def leer_todos_certificates():
    extractor = OfficeDocumentExtractor()
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    carpeta_certificates = os.path.join(directorio_actual, 'certificates')
    
    if not os.path.exists(carpeta_certificates):
        print(f"El directorio {carpeta_certificates} no existe")
        return
    
    extensiones_prohibidas = {ext.value for ext in ArchivoProhibido}
    
    for raiz, dirs, archivos in os.walk(carpeta_certificates):
        nombre_carpeta = os.path.basename(raiz)
        path_folder = os.path.join(raiz)
        
        # Crear carpeta de duplicados para esta carpeta de certificados
        carpeta_duplicados = crear_carpeta_duplicados(path_folder)
        path_report = create_excel_template(path_folder)
        
        archivos_permitidos = [
            archivo for archivo in archivos 
            if os.path.splitext(archivo)[1].lower() not in extensiones_prohibidas
            and archivo.lower() != 'report.xlsx' 
        ]
        
        print(f"\n📂 Procesando carpeta: {nombre_carpeta}")
        print(f"📄 Total de archivos: {len(archivos_permitidos)}")
        
        # Diccionario para trackear certificados ya procesados
        certificados_procesados = {}
        archivos_a_procesar = []
        
        # Fase 1: Identificar archivos ya procesados y duplicados obvios
        for archivo in archivos_permitidos:
            if archivo_ya_procesado(archivo):
                print(f"⏭️  Ya procesado: {archivo}")
                continue
            archivos_a_procesar.append(archivo)
        
        print(f"🔄 Archivos pendientes por procesar: {len(archivos_a_procesar)}")
        
        # Fase 2: Procesar archivos y detectar duplicados
        for archivo in archivos_a_procesar:
            path_file = os.path.join(raiz, archivo)
            print(f"\n🔍 Procesando: {archivo}")
            
            # Verificar que el archivo aún existe (no fue movido como duplicado)
            if not os.path.exists(path_file):
                print(f"⏭️  Archivo ya movido: {archivo}")
                continue
            
            try:
                inference_response = extractor.extract_content(path_file)
                
                # Generar clave única del certificado
                clave_certificado = generar_clave_certificado(inference_response)
                
                if clave_certificado:
                    # Verificar si ya existe este certificado
                    if clave_certificado in certificados_procesados:
                        archivo_original = certificados_procesados[clave_certificado]
                        print(f"🔄 DUPLICADO DETECTADO:")
                        print(f"   Original: {archivo_original}")
                        print(f"   Duplicado: {archivo}")
                        
                        # Mover el duplicado actual a la carpeta de duplicados
                        mover_a_duplicados(path_file, carpeta_duplicados, 
                                         f"(duplicado de {archivo_original})")
                        continue
                    else:
                        # Registrar este certificado como procesado
                        certificados_procesados[clave_certificado] = archivo
                
                identification = inference_response.get('identification')
                issue_date = inference_response.get('issue_date')
                expiration_date = inference_response.get('expiration_date')
                    
                if identification:
                    # Verificar si el renombrado fue exitoso
                    renombrado_exitoso = renombrar_archivo_con_fechas(
                        ruta_original=path_file,
                        identification=identification,
                        issue_date=issue_date,
                        expiration_date=expiration_date,
                        carpeta_duplicados=carpeta_duplicados
                    )
                    
                    if renombrado_exitoso:
                        print(f"✅ Procesado y renombrado correctamente")
                    else:
                        print(f"🗂️  Procesado y movido a duplicados (archivo ya existía)")
                        # No agregamos al reporte porque se procesó correctamente, solo se movió
                else:
                    inference_response['name'] = archivo
                    insert_certificate_data(path_report, inference_response)
                    print(f"⚠️  Sin identificación - agregado al reporte")
                    
            except Exception as e:
                print(f"💥 Error procesando {archivo}: {str(e)}")
                inference_response = {'name': archivo, 'message_error': str(e)}
                insert_certificate_data(path_report, inference_response)
        
        print(f"\n📊 Resumen de {nombre_carpeta}:")
        print(f"   - Certificados únicos procesados: {len(certificados_procesados)}")
        
        # Contar duplicados movidos
        if os.path.exists(carpeta_duplicados):
            duplicados_count = len([f for f in os.listdir(carpeta_duplicados) 
                                 if os.path.isfile(os.path.join(carpeta_duplicados, f))])
            print(f"   - Duplicados movidos: {duplicados_count}")

if __name__ == "__main__":
    leer_todos_certificates()