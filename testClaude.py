import anthropic
import base64
import os
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("ANTHROPIC_API_KEY")

prompt = """Tu tarea es extraer información específica de certificados escaneados, siendo especialmente flexible con variaciones en el texto debido a OCR.

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

# Lee el PDF local
with open("/home/desarrollo/Downloads/1000698855_isabella_olarte_uruena_cargue_plan_de_entrenamiento.pdf", "rb") as f:
    pdf_data = base64.b64encode(f.read()).decode("utf-8")

client = anthropic.Anthropic(api_key=api_key)

message = client.messages.create(
    model="claude-3-5-sonnet-20241022",
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
                        "data": pdf_data
                    }
                },
                {
                    "type": "text",
                    "text": prompt
                }
            ]
        }
    ]
)

print(message.content)