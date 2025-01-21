import google.generativeai as genai
from openai import OpenAI
import os, json, base64
from pathlib import Path
from PIL import Image
from typing import Optional
from dotenv import load_dotenv

load_dotenv()

class InferenceError(Exception):
    """Clase personalizada para errores de inferencia"""
    pass

class ResponseParsingError(Exception):
    """Clase personalizada para errores de parsing de respuesta"""
    pass

class CertificateInfo:
    def __init__(self, name: Optional[str], identification: Optional[str], issue_date: Optional[str], expiration_date: Optional[str]):
        self.name = name
        self.identification = identification
        self.issue_date = issue_date
        self.expiration_date = expiration_date

    @property
    def as_dict(self):
        return self.__dict__

GOOGLE_API_KEY = os.getenv("GEMINI_API") 
genai.configure(api_key=GOOGLE_API_KEY)

model = genai.GenerativeModel('gemini-1.5-flash')
api_key = os.getenv('API_OPENAI')

client_open_ai = OpenAI(api_key=api_key)

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

def encode_image_to_base64(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def parse_inference_response(response_text: str) -> dict:
    try:
        start_idx = response_text.find('{')
        end_idx = response_text.rfind('}')
        
        if start_idx == -1 or end_idx == -1:
            raise ResponseParsingError("No se encontró una estructura JSON válida en la respuesta")
        
        json_str = response_text[start_idx:end_idx + 1]
        
        data = json.loads(json_str)
        
        if not isinstance(data, dict):
            raise ResponseParsingError("La respuesta no es un objeto JSON válido")
        
        if 'name' not in data:
            raise ResponseParsingError("Falta el campo 'nombre' en la respuesta")
        
        return {
            'name': data.get('name'),
            'identification': data.get('identification'),
            'issue_date': data.get('issue_date'),
            'expiration_date': data.get('expiration_date')
        }
    
    except json.JSONDecodeError as e:
        print(f"Texto recibido: {response_text}")
        print(f"Intento de JSON: {json_str if 'json_str' in locals() else 'No se pudo extraer JSON'}")
        raise ResponseParsingError(f"Error al decodificar JSON: {str(e)}")
    except Exception as e:
        raise ResponseParsingError(f"Error inesperado al procesar la respuesta: {str(e)}")

def analyze_image_gemini(image_path):
    """
    Analiza una imagen usando la API de Gemini Pro Vision.

    Args:
        image_path (str): La ruta a la imagen que se va a analizar.
        prompt (str): La pregunta o instrucción para el modelo.

    Returns:
        str: La respuesta del modelo.
    """

    try:
        image = Image.open(image_path)  # Abre la imagen usando PIL
        response = model.generate_content([prompt_for_images, image])
        response.resolve()  # Resuelve la promesa para obtener el contenido

        if response.text:
            res_parse = parse_inference_response(response.text)
            return res_parse
        else:
            return "No se generó respuesta de texto."

    except FileNotFoundError:
        return "Error: La imagen no se encontró en la ruta especificada."
    except Exception as e:
        return f"Ocurrió un error: {e}"

def analyze_image_open_ai(image_path):
    """
    Analiza una imagen usando la API de OpenAI Vision.

    Args:
        image_path (str): La ruta a la imagen que se va a analizar.
        prompt (str): La pregunta o instrucción para el modelo.

    Returns:
        dict: La respuesta procesada del modelo.
    """
    try:
        base64_image = encode_image_to_base64(image_path)

        response = client_open_ai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_for_images},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=300
        )

        if response.choices:
            response_text = response.choices[0].message.content
            res_parse = parse_inference_response(response_text)
            return res_parse
        else:
            return "No se generó respuesta de texto."

    except FileNotFoundError:
        return "Error: La imagen no se encontró en la ruta especificada."
    except Exception as e:
        return f"Ocurrió un error: {e}"
    
if __name__ == "__main__":
    image_path = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23496192_luz_marina_bustos_rodriguez_certificado_reanimacion.jpg"

    result = analyze_image_gemini(image_path)
    print(result)

# png, jpg, jpeg