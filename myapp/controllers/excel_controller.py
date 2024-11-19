import os
from flask import send_file
from myapp.services.excel_service import procesar_excel
from myapp.services.limpieza_service import procesar_excel_dinamico

OUTPUT_PATH = os.path.join(os.getcwd(), 'plantilla_modificada.xlsx')

def rellenar_excel(request):
    data = request.json
    try:
        procesar_excel(data)
        return send_file(OUTPUT_PATH, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except FileNotFoundError:
        return "El archivo de plantilla de Excel no se encontró. Verifique la ruta. controller", 404
    except Exception as e:
        return str(e), 500
    
def rellenar_excel_limpieza(request):
    data = request.json
    try:
        # Ahora procesar_excel devuelve un buffer en memoria
        excel_buffer = procesar_excel_dinamico(data)
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name='plantilla_modificada.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except FileNotFoundError:
        return "El archivo de plantilla de Excel no se encontró. Verifique la ruta.", 404
    except Exception as e:
        return str(e), 500