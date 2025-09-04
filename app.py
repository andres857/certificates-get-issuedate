import os
import shutil
from pathlib import Path
import logging
from enum import Enum
from readDocs import OfficeDocumentExtractor
from utils import extraer_id_archivo, renombrar_archivo_con_fechas, obtener_identificacion_validada
from excel import (create_excel_template, insert_processed_certificate, 
                   insert_duplicate_certificate, insert_error_certificate, 
                   update_summary_stats)

# Enum simple para las extensiones prohibidas
class ArchivoProhibido(Enum):
    INI = '.ini'
    EXE = '.exe'
    HTML = '.html'

def crear_carpeta_duplicados(path_folder):
    """Crea la carpeta 'duplicates' dentro del directorio de certificados"""
    carpeta_duplicados = os.path.join(path_folder, 'duplicates')
    if not os.path.exists(carpeta_duplicados):
        os.makedirs(carpeta_duplicados)
        print(f"📁 Carpeta de duplicados creada: {carpeta_duplicados}")
    return carpeta_duplicados

def generar_clave_certificado(inference_response):
    """Genera una clave única basada en identificación, issue_date y expiration_date"""
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
    """Verifica si un archivo ya fue procesado buscando 'issueddate' en el nombre"""
    partes = nombre_archivo.split('_')
    return any('issueddate' in parte for parte in partes)

def mover_a_duplicados(archivo_path, carpeta_duplicados, razon=""):
    """Mueve un archivo a la carpeta de duplicados"""
    try:
        if not os.path.exists(archivo_path):
            print(f"❌ Archivo no existe: {archivo_path}")
            return False, ""
            
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
        nombre_final = os.path.basename(destino)
        print(f"🗂️  ✅ MOVIDO a duplicados: {nombre_archivo} → {nombre_final} {razon}")
        return True, nombre_final
        
    except Exception as e:
        print(f"❌ Error moviendo {archivo_path} a duplicados: {e}")
        return False, ""
    
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
        
        # Estadísticas para el reporte
        stats = {
            'total_files': len(archivos_permitidos),
            'processed': 0,
            'duplicates': 0,
            'errors': 0,
            'already_processed': 0
        }
        
        # Diccionario para trackear certificados ya procesados
        certificados_procesados = {}
        archivos_a_procesar = []
        
        # Fase 1: Identificar archivos ya procesados
        for archivo in archivos_permitidos:
            if archivo_ya_procesado(archivo):
                print(f"⏭️  Ya procesado: {archivo}")
                stats['already_processed'] += 1
                continue
            archivos_a_procesar.append(archivo)
        
        print(f"🔄 Archivos pendientes por procesar: {len(archivos_a_procesar)}")
        
        # Fase 2: Procesar archivos y detectar duplicados
        for archivo in archivos_a_procesar:
            path_file = os.path.join(raiz, archivo)
            print(f"\n🔍 Procesando: {archivo}")
            
            # Verificar que el archivo aún existe
            if not os.path.exists(path_file):
                print(f"⏭️  Archivo ya movido: {archivo}")
                continue
            
            try:
                inference_response = extractor.extract_content(path_file)
                
                # Verificar si hay errores en la extracción
                if inference_response.get('message_error'):
                    print(f"⚠️  Error en extracción: {inference_response['message_error']}")
                    insert_error_certificate(
                        path_report, 
                        archivo,
                        "Error de extracción",
                        inference_response['message_error'],
                        inference_response
                    )
                    stats['errors'] += 1
                    continue
                
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
                        movido, nombre_final = mover_a_duplicados(
                            path_file, 
                            carpeta_duplicados, 
                            f"(duplicado de {archivo_original})"
                        )
                        
                        if movido:
                            insert_duplicate_certificate(
                                path_report,
                                inference_response,
                                archivo,
                                archivo_original,
                                f"Duplicado detectado por clave: {clave_certificado}"
                            )
                            stats['duplicates'] += 1
                        continue
                    else:
                        # Registrar este certificado como procesado
                        certificados_procesados[clave_certificado] = archivo
                
                identification, es_validacion_ok, mensaje_validacion = obtener_identificacion_validada(inference_response, archivo)

                print(f"🔍 {mensaje_validacion}")

                # Si no es válido pero tenemos identification del archivo, continuar con advertencia
                if not es_validacion_ok:
                    if identification:  # Tenemos ID del archivo pero hay problemas con la inferencia
                        print(f"⚠️  ADVERTENCIA: Continuando con ID del archivo: {identification}")
                
                    else:  
                        print(f"❌ ERROR: {mensaje_validacion}")
                        insert_error_certificate(
                            path_report,
                            archivo,
                            "Sin identificación válida",
                            mensaje_validacion,
                            inference_response
                        )
                        stats['errors'] += 1
                        continue
                
                issue_date = inference_response.get('issue_date')
                expiration_date = inference_response.get('expiration_date')
                    
                if identification:
                    # Intentar renombrar el archivo
                    renombrado_exitoso = renombrar_archivo_con_fechas(
                        ruta_original=path_file,
                        identification=identification,
                        issue_date=issue_date,
                        expiration_date=expiration_date,
                        carpeta_duplicados=carpeta_duplicados
                    )
                    
                    if renombrado_exitoso:
                        # Generar el nombre del archivo renombrado
                        nombre_sin_extension = os.path.splitext(archivo)[0]
                        extension = os.path.splitext(archivo)[1]
                        nuevo_nombre = nombre_sin_extension
                        
                        if issue_date:
                            fecha_emision = issue_date.split('T')[0] if 'T' in issue_date else issue_date
                            nuevo_nombre += f"_issueddate{fecha_emision}"
                        
                        if expiration_date:
                            fecha_expiracion = expiration_date.split('T')[0] if 'T' in expiration_date else expiration_date
                            nuevo_nombre += f"_expirationdate{fecha_expiracion}"
                        
                        nuevo_nombre += extension
                        
                        insert_processed_certificate(
                            path_report,
                            inference_response,
                            archivo,
                            nuevo_nombre
                        )
                        stats['processed'] += 1
                        print(f"✅ Procesado y renombrado correctamente")
                    else:
                        # Se movió a duplicados porque el archivo ya existía
                        insert_duplicate_certificate(
                            path_report,
                            inference_response,
                            archivo,
                            "Archivo con nombre ya existente",
                            "Archivo renombrado ya existía en el directorio"
                        )
                        stats['duplicates'] += 1
                        print(f"🗂️  Procesado y movido a duplicados (archivo ya existía)")
                else:
                    # Sin identificación - es un error
                    insert_error_certificate(
                        path_report,
                        archivo,
                        "Sin identificación",
                        "No se pudo extraer la identificación del certificado",
                        inference_response
                    )
                    stats['errors'] += 1
                    print(f"⚠️  Sin identificación - agregado al reporte de errores")
                    
            except Exception as e:
                print(f"💥 Error procesando {archivo}: {str(e)}")
                insert_error_certificate(
                    path_report,
                    archivo,
                    "Error de procesamiento",
                    str(e),
                    {'name': archivo}
                )
                stats['errors'] += 1
        
        # Actualizar estadísticas en el reporte
        update_summary_stats(path_report, stats)
        
        print(f"\n📊 Resumen de {nombre_carpeta}:")
        print(f"   - Certificados únicos procesados: {stats['processed']}")
        print(f"   - Duplicados detectados: {stats['duplicates']}")
        print(f"   - Errores encontrados: {stats['errors']}")
        print(f"   - Ya procesados (omitidos): {stats['already_processed']}")
        print(f"   - Total de archivos: {stats['total_files']}")
        
        # Contar archivos físicos en duplicados
        if os.path.exists(carpeta_duplicados):
            duplicados_fisicos = len([f for f in os.listdir(carpeta_duplicados) 
                                   if os.path.isfile(os.path.join(carpeta_duplicados, f))])
            print(f"   - Archivos físicos en duplicados: {duplicados_fisicos}")

if __name__ == "__main__":
    leer_todos_certificates()