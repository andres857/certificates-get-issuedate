from abc import ABC, abstractmethod
from dotenv import load_dotenv
import google.generativeai as genai
from openai import OpenAI, RateLimitError, APIConnectionError, Timeout, OpenAIError
from typing import Optional
import json, os
from PIL import Image
import anthropic
import base64

load_dotenv()

class InferenceError(Exception):
    """Clase personalizada para errores de inferencia"""
    pass

class ResponseParsingError(Exception):
    """Clase personalizada para errores de parsing de respuesta"""
    pass

class BaseInference(ABC):
    """
    Clase base abstracta que define la estructura común para las clases de inferencia.
    Contiene la lógica compartida para el procesamiento de respuestas y el prompt base.
    """
    def __init__(self):
        self.prompt = """Tu tarea es extraer información específica de certificados escaneados, siendo especialmente flexible con variaciones en el texto debido a OCR.

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
        - Si encuentras una vigencia expresada en años (ej: "Vigencia: 3 años", "vigencia 2 años"):
        * Calcula la fecha de vencimiento sumando los años a la fecha de emisión
        * Si la vigencia está en meses, suma los meses correspondientes
        * Si la vigencia está en días, suma los días correspondientes
        - En caso de certificados con "Fecha Inicio" y "Vigencia X años", usa la fecha de inicio como base
        - Si no se encuentra fecha explícita ni vigencia para calcular, devuelve null

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

    def parse_response(self, response_text: str) -> dict:
        """
        Método común para procesar y validar las respuestas JSON de cualquier modelo.
        Este método se encarga de extraer, validar y formatear la respuesta JSON.
        
        Args:
            response_text: Texto de respuesta del modelo que contiene el JSON
            
        Returns:
            dict: Diccionario con los campos normalizados del certificado
            
        Raises:
            ResponseParsingError: Si hay problemas al procesar el JSON
        """
        try:
            # Encontramos los índices del JSON en la respuesta
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}')
            
            if start_idx == -1 or end_idx == -1:
                raise ResponseParsingError("No se encontró una estructura JSON válida en la respuesta")
            
            # Extraemos la cadena JSON
            json_str = response_text[start_idx:end_idx + 1]
            
            # Parseamos el JSON
            data = json.loads(json_str)
            
            # Validamos la estructura
            if not isinstance(data, dict):
                raise ResponseParsingError("La respuesta no es un objeto JSON válido")
            
            if 'name' not in data:
                raise ResponseParsingError("Falta el campo 'nombre' en la respuesta")
            
            # Retornamos un diccionario normalizado
            return {
                'name': data.get('name'),
                'identification': data.get('identification'),
                'issue_date': data.get('issue_date'),
                'expiration_date': data.get('expiration_date'),
                'message_error': None
            }
        
        except json.JSONDecodeError as e:
            print(f"Texto recibido: {response_text}")
            print(f"Intento de JSON: {json_str if 'json_str' in locals() else 'No se pudo extraer JSON'}")
            raise ResponseParsingError(f"Error al decodificar JSON: {str(e)}")
        except Exception as e:
            raise ResponseParsingError(f"Error inesperado al procesar la respuesta: {str(e)}")

    @abstractmethod
    def get_inference(self, content):
        """
        Método abstracto que debe ser implementado por las clases hijas.
        Define la interfaz común para obtener inferencias de cualquier modelo.
        """
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

class OpenAIInference(BaseInference):
    """Clase para manejar la inferencia usando OpenAI"""
    def __init__(self):
        super().__init__()
        self.api_key = os.getenv('API_OPENAI')
        self.client = OpenAI(api_key=self.api_key)
        self.claude_fallback = AntropicInferenceForPDF()

    def get_inference(self, content_certificate, path_pdf=None):
        try:
            completion = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                store=True,
                messages=[
                    {"role": "system", "content": self.prompt},
                    {"role": "user", "content": content_certificate}
                ],
                temperature=0.3
            )
            response_text = completion.choices[0].message.content

            if not response_text:
                raise Exception("La API no devolvió ninguna respuesta")
                
            certificate_info = self.parse_response(response_text)
            print('EVALUANDOOOOOOOOOOOOOO',certificate_info)
            if certificate_info['issue_date'] is None:
                print('Procesado con IAAAAAAAAA CLAUDE')
                certificate_info = self.claude_fallback.get_inference(path_pdf)
            else:
                print('Procesadooooooooooooooo con OpenAI')

            # print(f"Información extraída: {certificate_info}")
            return certificate_info
        
        except RateLimitError as e:
            print(f"Límite de solicitudes alcanzado: {str(e)}")
            # Tal vez podrías pausar la ejecución aquí y esperar un rato antes de reintentar
            raise
        except APIConnectionError as e:
            print(f"Problema de conexión con la API: {str(e)}")
            raise
        except Timeout as e:
            print(f"Se alcanzó el tiempo de espera al conectar con la API: {str(e)}")
            raise
        except OpenAIError as e:
            print(f"Error general en la API de OpenAI: {str(e)}")
            raise
        except Exception as e:
            print(f"Error inesperado: {str(e)}")
            raise

class GeminiInferenceForImages(BaseInference):
    def __init__(self):
        super().__init__()
        self.api_key = os.getenv('GEMINI_API')
        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel('gemini-1.5-flash')
        
    def get_inference(self, image_path):
        """
        Analiza una imagen usando la API de Gemini.

        Args:
            image_path (str): La ruta a la imagen que se va a analizar.
            prompt (str): La pregunta o instrucción para el modelo.

        Returns:
            str: La respuesta del modelo.
        """
        try:
            image = Image.open(image_path)
            response = self.model.generate_content([self.prompt, image])
            response.resolve()

            if response.text:
                res_parse = self.parse_response(response.text)
                return res_parse
            else:
                return "No se generó respuesta de texto."

        except FileNotFoundError:
            return "Error: La imagen no se encontró en la ruta especificada."
        except Exception as e:
            return f"Ocurrió un error: {e}"

class AntropicInferenceForPDF(BaseInference):
    def __init__(self):
        super().__init__()
        self.api_key = os.getenv('ANTHROPIC_API_KEY')
        self.client = anthropic.Anthropic(api_key=self.api_key)
        self.model = "claude-3-5-sonnet-20241022"
        # claude-3-5-sonnet-20240620
        # "claude-3-5-sonnet-20241022"

    def read_pdf(self, path_pdf):
        try:
            with open(path_pdf, "rb") as f:
                pdf_data = base64.b64encode(f.read()).decode("utf-8")
            return pdf_data
        except FileNotFoundError:
            raise InferenceError("El archivo PDF no se encontró en la ruta especificada")
    
    def get_inference(self, path_pdf):
        try:
            content_pdf = self.read_pdf(path_pdf)
            message = self.client.messages.create(
                model=self.model,
                max_tokens=1024,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "document",
                                "source": {
                                    "type": "base64",
                                    "media_type": "application/pdf",
                                    "data": content_pdf
                                }
                            },
                            {
                                "type": "text",
                                "text": self.prompt
                            }
                        ]
                    }
                ]
            )
            text_response = message.content[0]
            response_dict = json.loads(text_response.text)
            response_dict['message_error'] = None
            print(response_dict)
            return response_dict

        except anthropic.APIError as e:
            error_message = f"Error en la API de Anthropic: {str(e)}"
            print(error_message)
            raise InferenceError(error_message)

        except anthropic.APIConnectionError as e:
            error_message = f"Error de conexión con la API de Anthropic: {str(e)}"
            print(error_message)
            raise InferenceError(error_message)

        except anthropic.RateLimitError as e:
            error_message = f"Se ha excedido el límite de solicitudes a la API: {str(e)}"
            print(error_message)
            raise InferenceError(error_message)

        except Exception as e:
            error_message = f"Error inesperado: {str(e)}"
            print(error_message)
            raise InferenceError(error_message)
        
if __name__ == "__main__":
    content_pdf="Texto extraído por OCR:EN ] 0 (aa. ! 'FUNDACIÓN:CARDIOINFANTIL:- INSTITUTO DE'CARDIOLOGIA: : *CENTRO' DE SIMULACIÓN Y. HABILIDADES'CLINICAS “VALENTÍN FUSTER” ¿CERTIEICAYAS ñ ROSALBA, GUERRERO MONTOYA: (C.C.S1904946' o Por participación eme caer 3x0. PEDIATRIC ADVANCED LIFE. SUPPORT (PAIS) 77 4 7. REANIMACIÓN 'AVANZADA PEDIATRICA: a Realizado; enja ejudad de Bogotá! Ps . El.día:3 de: Diciembre de 2021. A ¡Con'unalntensidad de 48 horas IN Encálidád dez - AS. + ¡PROMEEDOR, ' eno IE. 3 7% ¿Curso Oficial. que.sigue los lineamientos establecidos porta: ap CE A ¿American Heart Association... o o pi JAIME FERNÁNDEZ SARMIENTO Mod: 000,0 a ta, e Director HOSP SIMULADO 0 ai "

    # ia_inference = OpenAIInference()
    # inference = ia_inference.get_inference(content_pdf)
    # print(inference)

    # ia_inference_image = GeminiInferenceForImages()
    # inference_image = ia_inference_image.get_inference('/home/desarrollo/Documents/wc/processing-certificates/certificates/1/23496192_luz_marina_bustos_rodriguez_certificado_reanimacion.jpg')
    # print(inference_image)

    ia_inference_pdf = AntropicInferenceForPDF()
    inference_pdf = ia_inference_pdf.get_inference('/home/desarrollo/Documents/certificados/Certificado La Cardio/Certificado Plan de Entrenamiento/328004_sebastian_kurt_villarroel_hagemann_cargue_plan_de_entrenamiento_1.pdf')