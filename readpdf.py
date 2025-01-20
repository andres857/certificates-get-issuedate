import pytesseract, time, fitz
from PIL import Image
from ia import  get_inference_for_pdf_open_ai

def get_user_data_by_OCR_METHOD(file_path):
    file = fitz.open(file_path)
    for page in file:
        try:
            # Convertimos la página a imagen
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # Realizamos OCR
            text = pytesseract.image_to_string(img)
            # print (text)
            time.sleep(0.1)
            user_data = get_inference_for_pdf_open_ai(text)
            return user_data
        except Exception as e:
            print(f"Error en el procesamiento del pdf usando OCR: {str(e)}")
            return None
    file.close()