import subprocess, shutil
from pathlib import Path
import logging, os

def convert_pptx_to_pdf(pptx_path: str) -> Path:
    """
    Convierte un archivo PPTX a PDF usando LibreOffice y lo guarda en la misma carpeta.
    
    Esta función hace lo siguiente:
    1. Verifica que LibreOffice esté instalado
    2. Convierte el archivo PPTX a PDF
    3. Guarda el PDF en la misma ubicación que el PPTX original
    
    Args:
        pptx_path: Ruta al archivo PPTX que se quiere convertir
        
    Returns:
        Path: Ruta al archivo PDF generado
        
    Raises:
        FileNotFoundError: Si el archivo PPTX no existe o LibreOffice no está instalado
        subprocess.CalledProcessError: Si hay un error durante la conversión
    """
    # Configuramos logging para tener información del proceso
    logging.basicConfig(level=logging.INFO, 
                       format='%(asctime)s - %(message)s')
    
    # Convertimos la ruta a objeto Path para mejor manejo
    input_path = Path(pptx_path).resolve()
    
    # Verificamos que el archivo existe
    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")
    
    # Verificamos que sea un archivo PPTX
    if input_path.suffix.lower() != '.pptx':
        raise ValueError(f"El archivo debe ser una presentación PPTX, no {input_path.suffix}")
    
    try:
        # Ejecutamos la conversión usando LibreOffice
        logging.info(f"Iniciando conversión de {input_path.name}")
        process = subprocess.run(
            [
                'soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(input_path.parent),
                str(input_path)
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Verificamos si la conversión fue exitosa
        if process.returncode != 0:
            raise subprocess.CalledProcessError(
                process.returncode,
                'soffice',
                output=process.stdout,
                stderr=process.stderr
            )
        
        # Construimos la ruta al PDF generado (mismo nombre, extensión .pdf)
        pdf_path = input_path.with_suffix('.pdf')
        
        # Verificamos que el PDF se generó correctamente
        if pdf_path.exists():
            logging.info(f"Conversión exitosa. PDF guardado en: {pdf_path}")
            return pdf_path
        else:
            raise FileNotFoundError("No se generó el archivo PDF")
            
    except FileNotFoundError:
        logging.error(
            "LibreOffice no está instalado. Por favor, instálalo con:\n"
            "sudo apt-get install libreoffice"
        )
        raise
    except Exception as e:
        logging.error(f"Error durante la conversión: {str(e)}")
        raise

def convert_doc_to_pdf(doc_path: str) -> Path:
    """
    Convierte un archivo DOC/DOCX a PDF usando LibreOffice y lo guarda en la misma carpeta.
    
    Esta función realiza el siguiente proceso:
    1. Verifica que LibreOffice esté instalado en el sistema
    2. Convierte el archivo DOC/DOCX a formato PDF
    3. Guarda el PDF resultante en la misma ubicación que el archivo original
    
    Args:
        doc_path: Ruta al archivo DOC/DOCX que se desea convertir
        
    Returns:
        Path: Ruta al archivo PDF generado
        
    Raises:
        FileNotFoundError: Si el archivo DOC/DOCX no existe o LibreOffice no está instalado
        subprocess.CalledProcessError: Si ocurre un error durante el proceso de conversión
        ValueError: Si el archivo no tiene la extensión correcta
    """
    # Configuramos el sistema de logging para tener un registro detallado del proceso
    logging.basicConfig(level=logging.INFO, 
                       format='%(asctime)s - %(message)s')
    
    # Convertimos la ruta de entrada a un objeto Path para mejor manipulación
    input_path = Path(doc_path).resolve()
    
    # Verificamos la existencia del archivo
    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")
    
    # Verificamos que sea un archivo DOC o DOCX
    valid_extensions = ['.doc', '.docx']
    if input_path.suffix.lower() not in valid_extensions:
        raise ValueError(f"El archivo debe ser un documento DOC o DOCX, no {input_path.suffix}")
    
    try:
        # Iniciamos el proceso de conversión usando LibreOffice
        # logging.info(f"Iniciando conversión de {input_path.name}")
        
        # Ejecutamos LibreOffice en modo headless (sin interfaz gráfica)
        process = subprocess.run(
            [
                'soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(input_path.parent),
                str(input_path)
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True  # Esto hará que se lance una excepción si hay error
        )
        
        # Construimos la ruta al archivo PDF resultante
        pdf_path = input_path.with_suffix('.pdf')
        
        # Verificamos que el PDF se haya generado correctamente
        if pdf_path.exists():
            # logging.info(f"Conversión exitosa. PDF guardado en: {pdf_path}")
            return pdf_path
        else:
            raise FileNotFoundError("No se pudo generar el archivo PDF")
            
    except FileNotFoundError:
        logging.error(
            "LibreOffice no está instalado en el sistema. Para instalarlo, ejecuta:\n"
            "sudo apt-get install libreoffice"
        )
        raise
    except subprocess.CalledProcessError as e:
        logging.error(f"Error durante la ejecución de LibreOffice: {e.stderr}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado durante la conversión: {str(e)}")
        raise


def extraer_id_archivo(nombre_archivo: str) -> str | None:
    """
    Extrae el ID numérico del nombre de un archivo, buscando los números
    que aparecen antes del primer guion bajo (_).
    
    Args:
        nombre_archivo: Ruta o nombre del archivo del cual extraer el ID
        
    Returns:
        str: El ID numérico si se encuentra
        None: Si no se encuentra un ID válido
    
    Ejemplos:
        >>> extraer_id_archivo("23496192_luz_marina_bustos_rodriguez.pptx")
        "23496192"
        >>> extraer_id_archivo("documento_sin_id.pdf")
        None
    """
    # Obtenemos solo el nombre del archivo sin la ruta completa
    nombre_base = os.path.basename(nombre_archivo)
    
    # Buscamos la posición del primer guion bajo
    posicion_guion = nombre_base.find('_')
    
    # Si no hay guion bajo o está al inicio, no hay ID válido
    if posicion_guion <= 0:
        return None
        
    # Extraemos la parte antes del primer guion bajo
    posible_id = nombre_base[:posicion_guion]
    
    # Verificamos si lo que obtuvimos es un número
    if posible_id.isdigit():
        return posible_id
    
    return None


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

def renombrar_archivo_con_fechas(ruta_original: str, identification: str | None = None, 
                                issue_date: str | None = None, expiration_date: str | None = None,
                                carpeta_duplicados: str | None = None) -> bool:
    """
    Renombra un archivo agregando las fechas al final. Si ya existe, lo mueve a duplicados.
    
    Args:
        ruta_original: Ruta completa del archivo
        identification: ID de la persona
        issue_date: Fecha de emisión en formato ISO
        expiration_date: Fecha de expiración en formato ISO (opcional)
        carpeta_duplicados: Ruta a la carpeta de duplicados
    
    Returns:
        bool: True si el renombrado fue exitoso, False si se movió a duplicados
    """
    directorio = os.path.dirname(ruta_original)
    nombre_original = os.path.basename(ruta_original)
    nombre_sin_extension, extension = os.path.splitext(nombre_original)
    
    # Si no hay identification, no podemos renombrar
    if not identification:
        print(f"❌ No se puede renombrar {nombre_original}: falta identification")
        return False
    
    # Construir el nuevo nombre base
    nuevo_nombre = f"{nombre_sin_extension}"
    
    # Agregar fecha de emisión si existe
    if issue_date:
        # Convertir de ISO a YYYY-MM-DD si es necesario
        fecha_emision = issue_date.split('T')[0] if 'T' in issue_date else issue_date
        nuevo_nombre += f"_issueddate{fecha_emision}"
    
    # Agregar fecha de expiración si existe
    if expiration_date:
        # Convertir de ISO a YYYY-MM-DD si es necesario
        fecha_expiracion = expiration_date.split('T')[0] if 'T' in expiration_date else expiration_date
        nuevo_nombre += f"_expirationdate{fecha_expiracion}"
    
    # Agregar la extensión
    nuevo_nombre += extension
    
    # Crear la ruta completa nueva
    nueva_ruta = os.path.join(directorio, nuevo_nombre)
    
    # Verificar si el archivo destino ya existe
    if os.path.exists(nueva_ruta):
        print(f"⚠️  El archivo renombrado ya existe: {nuevo_nombre}")
        
        # Si tenemos carpeta de duplicados, mover el archivo actual allí CON LAS FECHAS
        if carpeta_duplicados:
            # Primero renombramos el archivo temporal con las fechas
            try:
                temp_nueva_ruta = os.path.join(carpeta_duplicados, nuevo_nombre)
                
                # Si ya existe en duplicados, agregar sufijo numérico
                contador = 1
                nombre_base_dup, extension_dup = os.path.splitext(nuevo_nombre)
                while os.path.exists(temp_nueva_ruta):
                    nuevo_nombre_dup = f"{nombre_base_dup}_dup{contador}{extension_dup}"
                    temp_nueva_ruta = os.path.join(carpeta_duplicados, nuevo_nombre_dup)
                    contador += 1
                
                # Mover con el nombre que incluye fechas
                shutil.move(ruta_original, temp_nueva_ruta)
                print(f"🗂️  ✅ MOVIDO a duplicados con fechas: {nombre_original} → {os.path.basename(temp_nueva_ruta)} (archivo ya renombrado previamente)")
                return False  # Se movió a duplicados, no se renombró en lugar original
                
            except Exception as e:
                print(f"❌ Error moviendo {ruta_original} a duplicados: {e}")
                return False
        
        print(f"   No se realizó el renombrado")
        return False
    
    try:
        os.rename(ruta_original, nueva_ruta)
        print(f"✅ Archivo renombrado correctamente:")
        print(f"   De: {nombre_original}")
        print(f"   A:  {nuevo_nombre}")
        return True
    except Exception as e:
        print(f"❌ Error al renombrar el archivo: {str(e)}")
        return False

if __name__ == "__main__":
    # Puedes especificar aquí la ruta de tu archivo PPTX
    # pptx_file = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23496192_luz_marina_bustos_rodriguez_certificados_formacion_continua.pptx"
    
    # try:
    #     pdf_file = convert_pptx_to_pdf(pptx_file)
    #     print(f"PDF generado exitosamente: {pdf_file}")
    # except Exception as e:
    #     print(f"Error: {str(e)}")

    # doc_file = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23002896_clarissa_quintana_ramos_certificado_reanimacion.docx"
    
    # try:
    #     pdf_file = convert_doc_to_pdf(doc_file)
    #     print(f"PDF generado exitosamente en: {pdf_file}")
    # except Exception as e:
    #     print(f"Error durante la conversión: {str(e)}")

    # id1 = extraer_id_archivo("23496192_luz_marina_bustos_rodriguez.pptx")
    # print(type(id1))
    pass