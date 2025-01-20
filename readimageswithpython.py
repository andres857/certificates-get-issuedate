import pytesseract
from PIL import Image
import cv2
import numpy as np
from ia import get_inference_for_pdf_open_ai 


prompt_for_images = """Tu tarea es extraer información específica de certificados escaneados, siendo especialmente flexible con variaciones en el texto debido a OCR.

INSTRUCCIONES DE EXTRACCIÓN:

1. NOMBRE DEL CERTIFICADO:
- Busca el nombre completo, considerando posibles errores de OCR
- Puede aparecer después de palabras como: 'CERTIFICA', 'CERTIFICA A', 'CERTIFICA QUE', 'OTORGADO A', 'CERTIEICA'
- Ignora nombres de instituciones y firmantes
- Si hay varios nombres, prioriza el que aparece después de las palabras clave mencionadas

2. IDENTIFICACIÓN:
- Busca números precedidos por: 'C.C.', 'CC', 'Cédula', '(C.C', 'C.C:'
- Extrae solo los dígitos, ignorando caracteres especiales que pueda introducir el OCR
- Si encuentras múltiples números, prioriza el que está cerca del nombre

3. FECHA DE EMISIÓN:
- Busca patrones de fecha considerando:
  * Frases como: 'Realizado', 'El día', 'Fecha', 'Expedido'
  * Meses escritos en texto o números
  * Años en formato completo (YYYY) o corto (YY)
- Para fechas múltiples, prioriza:
  * La fecha que está después de 'realizado el', 'expedido el'
  * La fecha final en caso de rangos
- Acepta variaciones como:
  * "3 de Diciembre de 2021"
  * "3 Diciembre 2021"
  * "Diciembre 3, 2021"
  * "03/12/2021"

4. FECHA DE VENCIMIENTO:
- Busca términos como: 'válido hasta', 'vigencia', 'vence', 'expira'
- Usa el mismo procesamiento flexible que para la fecha de emisión
- Si no se encuentra, devuelve null

FORMATO DE SALIDA:
{
    "name": string | null,
    "identification": string | null,
    "issue_date": "2019-12-03T00:00:00.000Z" | null,
    "expiration_date": "2019-12-03T00:00:00.000Z" | null
}

EJEMPLO DE PROCESAMIENTO:
Entrada confusa por OCR: "CERTIEICA YA ñ JUAN PERÉZ (C.C.S1234567) Realizado; enja ejudad el.día:3 de: Diciembre de 2021"
Debe producir:
{
    "name": "JUAN PEREZ",
    "identification": "1234567",
    "issue_date": 2019-12-03T00:00:00.000Z,
    "expiration_date": null
}

IMPORTANTE:
- Sé tolerante con errores de OCR y caracteres especiales
- Busca patrones aproximados cuando el texto esté distorsionado
- Prioriza la extracción de fechas en formato completo (día, mes y año)
- Solo devuelve el JSON, sin explicaciones adicionales"""

def get_text_from_image(image_path):
    try:
        # Leer la imagen
        image = cv2.imread(image_path)
        
        # Convertir a escala de grises
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Aplicar umbral adaptativo para mejorar el contraste
        thresh = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            cv2.THRESH_BINARY, 11, 2
        )
        
        # Opcional: Aplicar procesamiento adicional para mejorar la calidad
        denoised = cv2.fastNlMeansDenoising(thresh)
        
        # Convertir la imagen procesada a formato PIL
        pil_image = Image.fromarray(denoised)
        
        # Configurar pytesseract para mejor precisión
        custom_config = r'--oem 3 --psm 6'
        
        # Realizar OCR
        text = pytesseract.image_to_string(
            pil_image, 
            config=custom_config, 
            lang='spa'  # Puedes cambiar el idioma según necesites
        )
        print('holaaaaaaaa',text)
        return text.strip()

    except Exception as e:
        print(f"Error al procesar la imagen: {str(e)}")
        return None

def analyze_certificate_image(image_path):
    try:
        # Obtener texto de la imagen
        text = get_text_from_image(image_path)
        if text:
            # Usar tu función existente de inferencia
            result = get_inference_for_pdf_open_ai(text, prompt_for_images)
            return result
        else:
            print("No se pudo extraer texto de la imagen")
            return None
            
    except Exception as e:
        print(f"Error en el procesamiento: {str(e)}")
        return None

# Uso
if __name__ == "__main__":
    image_path = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/1012434734_dayana_aleman_ibarra_certificado_reanimacion2.jpg"
    result = analyze_certificate_image(image_path)
    print(result)