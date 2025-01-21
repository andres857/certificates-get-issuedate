import subprocess
from pathlib import Path
import logging

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

# Ejemplo de uso
if __name__ == "__main__":
    # Puedes especificar aquí la ruta de tu archivo PPTX
    # pptx_file = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23496192_luz_marina_bustos_rodriguez_certificados_formacion_continua.pptx"
    
    # try:
    #     pdf_file = convert_pptx_to_pdf(pptx_file)
    #     print(f"PDF generado exitosamente: {pdf_file}")
    # except Exception as e:
    #     print(f"Error: {str(e)}")

    doc_file = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23002896_clarissa_quintana_ramos_certificado_reanimacion.docx"
    
    try:
        pdf_file = convert_doc_to_pdf(doc_file)
        print(f"PDF generado exitosamente en: {pdf_file}")
    except Exception as e:
        print(f"Error durante la conversión: {str(e)}")