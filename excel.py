from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from pathlib import Path
import logging

def create_excel_template(current_path: str) -> Path:
    """
    Crea un archivo Excel con una estructura predefinida para almacenar información de certificados.
    
    Esta función realiza las siguientes tareas:
    1. Crea un nuevo archivo Excel en la ruta especificada
    2. Define las columnas necesarias para la información de los certificados
    3. Aplica formato al encabezado para mejor visibilidad
    4. Guarda el archivo en la ubicación especificada
    
    Args:
        current_path: Ruta donde se guardará el archivo Excel
        
    Returns:
        Path: Ruta al archivo Excel creado
    """
    # Configuramos logging para seguimiento del proceso
    logging.basicConfig(level=logging.INFO, 
                       format='%(asctime)s - %(message)s')
    
    try:
        # Creamos un nuevo libro de trabajo
        wb = Workbook()
        ws = wb.active
        ws.title = "Certificados"
        
        # Definimos los encabezados de las columnas
        headers = [
            "Identificación",
            "Nombre",
            "Certificado",
            "Fecha de Emisión",
            "Fecha de Expiración"
        ]
        
        # Aplicamos los encabezados y el formato
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            # Aplicamos estilo al encabezado
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="0066CC",
                                  end_color="0066CC",
                                  fill_type="solid")
        
        # Ajustamos el ancho de las columnas
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
        
        # Creamos la ruta completa para el archivo
        excel_path = Path(current_path) / "registro_certificados.xlsx"
        
        # Guardamos el archivo
        wb.save(str(excel_path))
        logging.info(f"Archivo Excel creado exitosamente en: {excel_path}")
        
        return excel_path
    
    except Exception as e:
        logging.error(f"Error al crear el archivo Excel: {str(e)}")
        raise

def insert_certificate_data(excel_path: Path, data: list) -> None:
    """
    Inserta datos de certificados en el archivo Excel existente.
    
    Esta función realiza lo siguiente:
    1. Carga el archivo Excel existente
    2. Encuentra la última fila con datos
    3. Inserta los nuevos datos en las filas siguientes
    4. Guarda los cambios en el archivo
    
    Args:
        excel_path: Ruta al archivo Excel
        data: Lista de diccionarios con la información de los certificados
              Cada diccionario debe contener: identification, name, certificate,
              issue_date, expiration_date
    """
    try:
        # Cargamos el archivo Excel existente
        wb = load_workbook(str(excel_path))
        ws = wb.active
        
        # Encontramos la última fila con datos
        last_row = ws.max_row + 1
        
        # Insertamos los nuevos datos
        for row_data in data:
            # Insertamos cada campo en su columna correspondiente
            ws.cell(row=last_row, column=1, value=row_data.get('identification', ''))
            ws.cell(row=last_row, column=2, value=row_data.get('name', ''))
            ws.cell(row=last_row, column=3, value=row_data.get('certificate', ''))
            ws.cell(row=last_row, column=4, value=row_data.get('issue_date', ''))
            ws.cell(row=last_row, column=5, value=row_data.get('expiration_date', ''))
            
            last_row += 1
        
        # Guardamos los cambios
        wb.save(str(excel_path))
        logging.info(f"Datos insertados exitosamente en {excel_path}")
        
    except Exception as e:
        logging.error(f"Error al insertar datos en el Excel: {str(e)}")
        raise

# Ejemplo de uso con datos de prueba
if __name__ == "__main__":
    # Ruta donde se guardará el Excel
    current_path = "/home/desarrollo/Documents/wc/processing-certificates/certificates/1"
    
    # try:
    #     # Creamos el archivo Excel
    #     excel_file = create_excel_template(current_path)
        
    #     # Datos de prueba
    #     test_data = [
    #         {
    #             "identification": "1234567890",
    #             "name": "Juan Pérez",
    #             "certificate": "Curso de Python",
    #             "issue_date": "2024-01-15",
    #             "expiration_date": "2025-01-15"
    #         },
    #         {
    #             "identification": "0987654321",
    #             "name": "María García",
    #             "certificate": "Curso de JavaScript",
    #             "issue_date": "2024-01-20",
    #             "expiration_date": "2025-01-20"
    #         }
    #     ]
        
    #     # Insertamos los datos de prueba
    #     insert_certificate_data(excel_file, test_data)
    #     print("Proceso completado exitosamente")
        
    # except Exception as e:
    #     print(f"Error en el proceso: {str(e)}")

    # test_data = [
    #     {
    #         "identification": "1234567890",
    #         "name": "Juan Pérez",
    #         "certificate": "Curso de Python",
    #         "issue_date": "2024-01-15",
    #         "expiration_date": "2025-01-15"
    #     },
    #     {
    #         "identification": "0987654321",
    #         "name": "María García",
    #         "certificate": "Curso de JavaScript",
    #         "issue_date": "2024-01-20",
    #         "expiration_date": "2025-01-20"
    #     }
    # ]
    # insert_certificate_data('/home/desarrollo/Documents/wc/processing-certificates/certificates/1/registro_certificados.xlsx', test_data)