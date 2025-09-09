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
            "name": string | None,
            "identification": string | None,
            "issue_date": "2019-12-03T00:00:00.000Z" | None,
            "expiration_date": "2019-12-03T00:00:00.000Z" | None
        }

        EJEMPLO DE PROCESAMIENTO:
        Entrada confusa por OCR: "CERTIEICA YA ñ JUAN PERÉZ (C.C.S1234567) Realizado; enja ejudad el.día:3 de: Diciembre de 2021"
        Debe producir:
        {
            "name": "JUAN PEREZ",
            "identification": "1234567",
            "issue_date": 2019-12-03T00:00:00.000Z,
            "expiration_date": None
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
        self.gemini_fallback = GeminiInferenceForDocuments()

    def get_inference(self, content_certificate, path_pdf=None, force_model=None):
        """
        Obtiene inferencia de certificados con opción de forzar un modelo específico.
        
        Args:
            content_certificate: Contenido del certificado extraído por OCR
            path_pdf: Ruta al archivo PDF (requerido para fallbacks)
            force_model: Modelo específico a usar ('openai', 'gemini', 'claude'). 
                        Si es None, sigue el flujo normal.
        
        Returns:
            dict: Información extraída del certificado
        """
        
        # Si se fuerza un modelo específico, usar solo ese
        if force_model:
            return self._process_with_specific_model(force_model, content_certificate, path_pdf)
        
        # Flujo normal: OpenAI → Gemini → Claude
        try:
            completion = self.client.chat.completions.create(
                model="gpt-5-mini",
                store=True,
                messages=[
                    {"role": "system", "content": self.prompt},
                    {"role": "user", "content": content_certificate}
                ]
            )
            response_text = completion.choices[0].message.content

            if not response_text:
                raise Exception("La API no devolvió ninguna respuesta")
                
            certificate_info = self.parse_response(response_text)
            print('EVALUANDO OpenAI:', certificate_info)
            
            if certificate_info['issue_date'] is None:
                print('Intentando procesar con GEMINI...')
                try:
                    gemini_result = self.gemini_fallback.get_inference(path_pdf)
                    
                    # Verificar si Gemini extrajo el issue_date
                    if gemini_result.get('issue_date') is not None:
                        certificate_info = gemini_result
                        print('Procesado con GEMINI exitosamente')
                    else:
                        print('GEMINI no extrajo issue_date, intentando con CLAUDE...')
                        try:
                            claude_result = self.claude_fallback.get_inference(path_pdf)
                            if claude_result.get('issue_date') is not None:
                                certificate_info = claude_result
                                print('Procesado con CLAUDE exitosamente')
                            else:
                                print('CLAUDE tampoco extrajo issue_date, usando resultado de GEMINI')
                                certificate_info = gemini_result
                                certificate_info['message_error'] = "Ningún modelo pudo extraer issue_date"
                        except Exception as claude_error:
                            print(f'Error con CLAUDE: {claude_error}')
                            certificate_info = gemini_result
                            certificate_info['message_error'] = f"Gemini no extrajo issue_date. Claude falló: {claude_error}"
                            
                except Exception as gemini_error:
                    print(f'Error con GEMINI: {gemini_error}')
                    print('Intentando procesar con CLAUDE...')
                    try:
                        certificate_info = self.claude_fallback.get_inference(path_pdf)
                        print('Procesado con CLAUDE exitosamente')
                    except Exception as claude_error:
                        print(f'Error con CLAUDE: {claude_error}')
                        print('Todos los fallbacks fallaron, retornando resultado original de OpenAI')
                        certificate_info['message_error'] = f"OpenAI no extrajo issue_date. Gemini falló: {gemini_error}. Claude falló: {claude_error}"
            else:
                print('Procesado con OpenAI exitosamente')

            return certificate_info
        
        except RateLimitError as e:
            print(f"Límite de solicitudes alcanzado: {str(e)}")
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

    def _process_with_specific_model(self, model_name, content_certificate, path_pdf):
        """
        Procesa el certificado usando un modelo específico.
        
        Args:
            model_name: Nombre del modelo ('openai', 'gemini', 'claude')
            content_certificate: Contenido del certificado
            path_pdf: Ruta al archivo PDF
        
        Returns:
            dict: Información extraída del certificado
        """
        model_name = model_name.lower()
        
        try:
            if model_name == 'openai':
                print(f'Forzando procesamiento solo con OpenAI...')
                completion = self.client.chat.completions.create(
                    model="gpt-5-mini",
                    store=True,
                    messages=[
                        {"role": "system", "content": self.prompt},
                        {"role": "user", "content": content_certificate}
                    ]
                )
                response_text = completion.choices[0].message.content
                if not response_text:
                    raise Exception("La API no devolvió ninguna respuesta")
                result = self.parse_response(response_text)
                print('Procesado SOLO con OpenAI')
                return result
                
            elif model_name == 'gemini':
                print(f'Forzando procesamiento solo con Gemini...')
                if not path_pdf:
                    raise ValueError("path_pdf es requerido para usar Gemini")
                result = self.gemini_fallback.get_inference(path_pdf)
                print('Procesado SOLO con Gemini')
                return result
                
            elif model_name == 'claude':
                print(f'Forzando procesamiento solo con Claude...')
                if not path_pdf:
                    raise ValueError("path_pdf es requerido para usar Claude")
                result = self.claude_fallback.get_inference(path_pdf)
                print('Procesado SOLO con Claude')
                return result
                
            else:
                raise ValueError(f"Modelo '{model_name}' no reconocido. Opciones válidas: 'openai', 'gemini', 'claude'")
                
        except Exception as e:
            print(f"Error al procesar con {model_name.upper()}: {e}")
            # Retornar estructura básica con error
            return {
                'name': None,
                'identification': None,
                'issue_date': None,
                'expiration_date': None,
                'message_error': f"Error con {model_name.upper()}: {str(e)}"
            }

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
        # self.model = "claude-sonnet-4-20250514"

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

class GeminiInferenceForDocuments(BaseInference):
    def __init__(self):
        super().__init__()
        self.api_key = os.getenv('GEMINI_API')
        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel('gemini-2.5-flash-lite')
        
    def get_inference(self, content_or_path=None, file_path=None):
        """
        Analiza un documento (imagen o PDF) usando la API de Gemini.
        
        Args:
            content_or_path: Si file_path es None, este es la ruta del archivo.
                           Si file_path no es None, este parámetro se ignora.
            file_path: La ruta al archivo (imagen o PDF) que se va a analizar.
                      Si se proporciona, se usa este en lugar de content_or_path.

        Returns:
            dict: La respuesta parseada del modelo como un diccionario con la estructura:
                  {'name': str|None, 'identification': str|None, 'issue_date': str|None, 
                   'expiration_date': str|None, 'message_error': str|None}
        """
        try:
            # Determinar qué parámetro usar como ruta del archivo
            if file_path is not None:
                # Llamada con dos parámetros: get_inference(content_doc, file_path)
                actual_file_path = file_path
            else:
                # Llamada con un parámetro: get_inference(file_path)
                actual_file_path = content_or_path
                
            # Convertir a string si es un Path object
            file_path_str = str(actual_file_path)
            
            # Validar que el archivo existe
            if not os.path.exists(file_path_str):
                raise InferenceError(f"El archivo no se encontró en la ruta especificada: {file_path_str}")

            mime_type = None
            file_content = None
            file_extension = file_path_str.lower()

            if file_extension.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp")):
                # Es una imagen - usar PIL
                mime_type = self._get_image_mime_type(file_extension)
                file_content = Image.open(file_path_str)
                
            elif file_extension.endswith(".pdf"):
                # Es un PDF - leer como binario
                mime_type = "application/pdf"
                with open(file_path_str, "rb") as f:
                    file_content = f.read()
                    
            else:
                raise InferenceError("Tipo de archivo no soportado. Solo se aceptan imágenes (JPG, PNG, GIF, WEBP, BMP) y PDF.")

            # Preparar el contenido para el modelo
            if isinstance(file_content, Image.Image):
                # Para imágenes: enviar directamente el objeto PIL Image
                contents = [self.prompt, file_content]
            else:
                # Para PDFs: enviar como documento con mime_type y data
                contents = [
                    self.prompt,
                    {
                        "mime_type": mime_type,
                        "data": file_content
                    }
                ]

            # Generar contenido con Gemini
            print('Procesado con Gemini', file_path_str)
            response = self.model.generate_content(contents)
            response.resolve()  # Asegurar que la respuesta está completa

            if not response.text:
                return {
                    "name": None,
                    "identification": None, 
                    "issue_date": None,
                    "expiration_date": None,
                    "message_error": "El modelo no generó respuesta de texto"
                }

            # Parsear la respuesta usando el método heredado de BaseInference
            return self.parse_response(response.text)

        except InferenceError:
            # Re-lanzar errores de inferencia tal como están
            raise
            
        except FileNotFoundError:
            error_msg = f"El archivo no se encontró en la ruta especificada: {file_path_str}"
            raise InferenceError(error_msg)
            
        except Exception as e:
            error_msg = f"Error inesperado durante la inferencia con Gemini: {str(e)}"
            print(f"Error detallado: {e}")
            raise InferenceError(error_msg)
    
    def _get_image_mime_type(self, file_extension):
        """Determinar el mime_type correcto basado en la extensión del archivo"""
        mime_types = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg", 
            ".png": "image/png",
            ".gif": "image/gif",
            ".webp": "image/webp",
            ".bmp": "image/bmp"
        }
        
        for ext, mime in mime_types.items():
            if file_extension.endswith(ext):
                return mime
        
        # Default fallback
        return "image/jpeg"
        
if __name__ == "__main__":
    content_pdf="Texto extraído por OCR:EN ] 0 (aa. ! 'FUNDACIÓN:CARDIOINFANTIL:- INSTITUTO DE'CARDIOLOGIA: : *CENTRO' DE SIMULACIÓN Y. HABILIDADES'CLINICAS “VALENTÍN FUSTER” ¿CERTIEICAYAS ñ ROSALBA, GUERRERO MONTOYA: (C.C.S1904946' o Por participación eme caer 3x0. PEDIATRIC ADVANCED LIFE. SUPPORT (PAIS) 77 4 7. REANIMACIÓN 'AVANZADA PEDIATRICA: a Realizado; enja ejudad de Bogotá! Ps . El.día:3 de: Diciembre de 2021. A ¡Con'unalntensidad de 48 horas IN Encálidád dez - AS. + ¡PROMEEDOR, ' eno IE. 3 7% ¿Curso Oficial. que.sigue los lineamientos establecidos porta: ap CE A ¿American Heart Association... o o pi JAIME FERNÁNDEZ SARMIENTO Mod: 000,0 a ta, e Director HOSP SIMULADO 0 ai "

    # ia_inference = OpenAIInference()
    # inference = ia_inference.get_inference(content_pdf)
    # print(inference)

    ia_inference_image = GeminiInferenceForDocuments()
    inference_image = ia_inference_image.get_inference('/home/desarrollo/Documents/wc/project-certificates-cardio/processing-certificates/certificates/test/6613299_jairo_gallo_diaz_certificados_atencion_a_victimas_de_violencia_sexual.pdf')
    print(inference_image)

    # ia_inference_pdf = AntropicInferenceForPDF()
    # inference_pdf = ia_inference_pdf.get_inference('/home/desarrollo/Documents/certificados/Certificado La Cardio/Certificado Plan de Entrenamiento/328004_sebastian_kurt_villarroel_hagemann_cargue_plan_de_entrenamiento_1.pdf')