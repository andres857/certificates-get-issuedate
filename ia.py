from dotenv import load_dotenv
import google.generativeai as genai
from openai import OpenAI
from typing import Optional, Dict, Any
import json, os
from pathlib import Path
from PIL import Image  
load_dotenv()

class InferenceError(Exception):
    """Clase personalizada para errores de inferencia"""
    pass

class ResponseParsingError(Exception):
    """Clase personalizada para errores de parsing de respuesta"""
    pass

class CertificateInfo:
    """Clase para almacenar la información extraída del certificado"""
    def __init__(self, name: Optional[str], identification: Optional[str], issue_date: Optional[str], expiration_date: Optional[str]):
        self.name = name
        self.identification = identification
        self.issue_date = issue_date
        self.expiration_date = expiration_date

    @property
    def as_dict(self):
        return self.__dict__
    
prompt="Tu tarea es extraer información específica de certificados de cursos y capacitaciones. Instrucciones: 1. Lee cuidadosamente el texto del certificado proporcionado 2. Identifica el nombre completo de la persona certificada 3. Identifica el número de identificación (puede estar precedido por C.C., Cédula, DNI, etc.)4.Identifica la fecha de emision y si tiene la fecha de vencimiento del certificado 5. Devuelve la información en formato JSON con las siguientes claves: - 'name': nombre completo de la persona - 'identification': número de identificación (solo números, sin prefijos) - 'issue_date': date en formato nodejs - expiration_date: date en formato nodejs, Si algún campo no se encuentra, devuelve null para ese campo. Ejemplo de entrada: CERTIFICA A: JUAN CARLOS PEREZ MARTINEZ C.C. 79845632 Por su asistencia... Ejemplo de salida: { 'nombre': 'JUAN CARLOS PEREZ MARTINEZ', 'identificacion': '79845632', 'issue_date:en formato new Date()', expiration_date: en formato newDate() } El contendo es diverso si no puedes inferir que se certifica con un nombre como CERTIFICA A: JUAN CARLOS PEREZ MARTINEZ devuelve null ya que no todos los nombres son validos ya que pueden ser el nombre de la intitucion o el nombre del certificado o quien firma pero queremos obtener a quien se otorga y si tiene , el nombre quitala A partir de ahora, procesa el siguiente texto y devuelve solo el JSON:"

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

api_key = os.getenv('API_OPENAI')

client_open_ai = OpenAI(api_key=api_key)

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
    
def get_inference_for_pdf_open_ai(content_certificate, prompt_system):
    try:
        completion = client_open_ai.chat.completions.create(
            model="gpt-3.5-turbo",
            store=True,
            messages=[
                {"role": "system", "content": prompt_system},
                {"role": "user", "content": content_certificate}
            ],
            temperature=0.3
        )
        response_text = completion.choices[0].message.content

        if not response_text:
            raise Exception("La API no devolvió ninguna respuesta")
            
        certificate_info = parse_inference_response(response_text)
        print(f"Información extraída: {certificate_info}")
        return certificate_info
    except Exception as e:
        print(f"Error inesperado en la API de OpenAI: {str(e)}")
        return None

if __name__ == "__main__":
    content_pdf="Texto extraído por OCR:EN ] 0 (aa. ! 'FUNDACIÓN:CARDIOINFANTIL:- INSTITUTO DE'CARDIOLOGIA: : *CENTRO' DE SIMULACIÓN Y. HABILIDADES'CLINICAS “VALENTÍN FUSTER” ¿CERTIEICAYAS ñ ROSALBA, GUERRERO MONTOYA: (C.C.S1904946' o Por participación eme caer 3x0. PEDIATRIC ADVANCED LIFE. SUPPORT (PAIS) 77 4 7. REANIMACIÓN 'AVANZADA PEDIATRICA: a Realizado; enja ejudad de Bogotá! Ps . El.día:3 de: Diciembre de 2021. A ¡Con'unalntensidad de 48 horas IN Encálidád dez - AS. + ¡PROMEEDOR, ' eno IE. 3 7% ¿Curso Oficial. que.sigue los lineamientos establecidos porta: ap CE A ¿American Heart Association... o o pi JAIME FERNÁNDEZ SARMIENTO Mod: 000,0 a ta, e Director HOSP SIMULADO 0 ai "
    inference = get_inference_for_pdf_open_ai(content_pdf, prompt)
    print (inference)