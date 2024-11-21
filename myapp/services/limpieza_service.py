import io
import os
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage
import requests
from io import BytesIO
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def get_template_path():
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    return os.path.join(base_dir, 'template', 'LIMPIEZA.xlsx')

def validar_datos_inspeccion(inspeccion_data):
    """Valida la estructura y valores de los datos de inspección"""
    logger.info("Validando datos de inspección")
    
    for elemento, dias in inspeccion_data.items():
        logger.info(f"Validando elemento: {elemento}")
        
        if not isinstance(dias, dict):
            logger.error(f"Error: valores para {elemento} no es un diccionario")
            return False
            
        for dia, valor in dias.items():
            if not isinstance(valor, bool) and valor is not None:
                logger.error(f"Error: valor inválido para {elemento} en {dia}: {valor}")
                return False
                
            logger.info(f"Valor válido para {elemento} en {dia}: {valor}")
    
    return True

def procesar_excel_dinamico(data):
    """
    Procesa la plantilla Excel y llena las celdas según la data recibida.
    """
    logger.info("==================== INICIO PROCESAMIENTO ====================")
    logger.info(f"Data completa recibida: {data}")
    
    wb = load_workbook(get_template_path())
    worksheet = wb.active

    dias_columnas = {
        "lunes": ("E", "G"),
        "martes": ("H", "J"),
        "miercoles": ("K", "M"),
        "jueves": ("N", "P"),
        "viernes": ("Q", "S"),
        "sabado": ("T", "V"),
        "domingo": ("W", "Y")
    }

    # Estilo para las celdas del formulario
    estilo_formulario = {
        'font': Font(
            name='Arial',
            size=16,
            bold=True
        ),
        'alignment': Alignment(
            horizontal='center',
            vertical='center'
        )
    }

    def obtener_celda_principal(hoja, celda):
        """Obtiene la celda principal si está en un rango fusionado"""
        for merged_range in hoja.merged_cells.ranges:
            if celda.coordinate in merged_range:
                return hoja.cell(merged_range.min_row, merged_range.min_col)
        return celda

    # Procesar datos del formulario si existen
    if 'FORMULARIO' in data:
        formulario = data['FORMULARIO']
        
        # Mapeo de campos y sus celdas correspondientes
        campos_formulario = {
            'FECHA': 'F6',
            'AÑO': 'I6',
            'PLACA': 'T7'
        }

        # Aplicar valores y estilos
        for campo, celda in campos_formulario.items():
            if campo in formulario:
                cell = worksheet[celda]
                cell.value = formulario[campo]
                cell.font = estilo_formulario['font']
                cell.alignment = estilo_formulario['alignment']

    # Obtener la data de inspección
    inspeccion = data.get("inspecciones", {})
    logger.info(f"Estructura de inspecciones: {type(inspeccion)}")
    logger.info(f"Contenido de inspecciones: {inspeccion}")

    fila_inicial = 11
    
    # Iterar sobre los elementos en el JSON
    for idx, (nombre_elemento, valores_dias) in enumerate(inspeccion.items()):
        fila_actual = fila_inicial + idx
        logger.info(f"\n=== Procesando elemento: {nombre_elemento} ===")
        logger.info(f"Tipo de valores_dias: {type(valores_dias)}")
        logger.info(f"Contenido de valores_dias: {valores_dias}")
        logger.info(f"Fila actual: {fila_actual}")

        # Iterar por cada día
        for dia, (col_inicio, col_fin) in dias_columnas.items():
            try:
                # Intentar obtener el valor del día
                valor_dia = getattr(valores_dias, dia, None)
                logger.info(f"\nProcesando día: {dia}")
                logger.info(f"Columnas para {dia}: {col_inicio}-{col_fin}")
                logger.info(f"Valor encontrado para {dia}: {valor_dia}")
                
                # Obtener la celda y verificar si está fusionada
                celda_destino = worksheet[f"{col_inicio}{fila_actual}"]
                celda_principal = obtener_celda_principal(worksheet, celda_destino)
                
                logger.info(f"Celda destino: {celda_destino.coordinate}")
                logger.info(f"Celda principal: {celda_principal.coordinate}")

                # Asignar el valor correspondiente
                if valor_dia is True:
                    logger.info(f"Estableciendo ✔ en celda {celda_principal.coordinate}")
                    celda_principal.value = "✔"
                    celda_principal.font = Font(name='Segoe UI Symbol', size=22, bold=True)
                    celda_principal.alignment = Alignment(horizontal='center', vertical='center')
                elif valor_dia is False:
                    logger.info(f"Estableciendo ❌ en celda {celda_principal.coordinate}")
                    celda_principal.value = "❌"
                    celda_principal.font = Font(name='Segoe UI Symbol', size=22, bold=True)
                    celda_principal.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    logger.info(f"Dejando celda {celda_principal.coordinate} vacía")
                    celda_principal.value = ""
                
                # Verificar el valor final
                logger.info(f"Valor final en celda {celda_principal.coordinate}: {celda_principal.value}")
                
            except Exception as e:
                logger.error(f"Error procesando día {dia}: {str(e)}", exc_info=True)

    logger.info("==================== FIN PROCESAMIENTO ====================")

    # Procesar imágenes si existen
    if 'IMAGENES' in data:
        insertar_imagenes(worksheet, data['IMAGENES'])

    # Guardar en buffer de memoria
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    return excel_buffer

def insertar_imagenes(ws, imagenes_data):
    """Inserta las imágenes en el Excel"""
    # Configuración de celdas fijas
    celdas_imagenes = {
        'LOGO': 'B2',
        'FIRMA_USER': 'B25'
    }

    # Tamaños fijos para cada tipo de imagen
    tamanos_fijos = {
        'LOGO': (200, 100),
        'FIRMA_USER': (200, 115)
    }

    # Grupos de celdas para verificar firma por día
    grupos_firma_user = {
        'lunes': ('E', 'G'),
        'martes': ('H', 'J'),
        'miércoles': ('K', 'M'),
        'jueves': ('N', 'P'),
        'viernes': ('Q', 'S'),
        'sábado': ('T', 'V'),
        'domingo': ('W', 'Y')
    }

    def verificar_contenido_columna(columna_inicio, columna_fin, fila_inicial=11, fila_final=20):
        """Verifica si hay contenido en el rango de celdas de una columna"""
        for fila in range(fila_inicial, fila_final + 1):
            for col in range(ord(columna_inicio), ord(columna_fin) + 1):
                celda = ws[f"{chr(col)}{fila}"]
                if celda.value:
                    return True
        return False

    # Insertar logo si existe
    if 'LOGO' in imagenes_data:
        insertar_imagen_en_celda(ws, imagenes_data['LOGO'], 
                               celdas_imagenes['LOGO'], 
                               tamanos_fijos['LOGO'])

    # Insertar firma de usuario donde corresponda
    if 'FIRMA_USER' in imagenes_data:
        for dia, (col_inicio, col_fin) in grupos_firma_user.items():
            if verificar_contenido_columna(col_inicio, col_fin):
                # Calculamos la columna del medio
                col_media = chr(ord(col_inicio) + 1)  # Avanzamos una letra para obtener la columna del medio
                celda_firma = f"{col_media}25"
                insertar_imagen_en_celda(ws, imagenes_data['FIRMA_USER'],
                                       celda_firma,
                                       tamanos_fijos['FIRMA_USER'])

def insertar_imagen_en_celda(ws, url, celda, tamano):
    """Inserta una imagen en una celda específica"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        img_data = BytesIO(response.content)
        
        # Detectar el formato de la imagen
        pil_image = Image.open(img_data)
        formato = pil_image.format.lower()
        
        # Verificar formato soportado
        formatos_soportados = ['png', 'jpeg', 'jpg']
        if formato not in formatos_soportados:
            print(f"Warning: Formato de imagen {formato} no soportado. Convirtiendo a PNG...")
            # Convertir a PNG
            width_px, height_px = tamano
            pil_image = pil_image.convert('RGBA')
            pil_image = pil_image.resize((width_px, height_px), Image.LANCZOS)
            
            img_final = BytesIO()
            pil_image.save(img_final, format='PNG')
            img_final.seek(0)
        else:
            # Procesar normalmente
            width_px, height_px = tamano
            pil_image = pil_image.resize((width_px, height_px), Image.LANCZOS)
            
            img_final = BytesIO()
            pil_image.save(img_final, format='PNG')
            img_final.seek(0)
        
        # Crear y configurar la imagen para Excel
        img = XLImage(img_final)
        img.width = width_px
        img.height = height_px
        
        # Insertar la imagen
        ws.add_image(img, celda)
        print(f"Imagen insertada correctamente en la celda {celda}")
        
    except Exception as e:
        print(f"Error al insertar la imagen en la celda {celda}: {e}")
        # Continuar sin la imagen en caso de error
        pass