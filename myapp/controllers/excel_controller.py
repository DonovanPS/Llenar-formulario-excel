import os
from flask import send_file
from myapp.services.excel_service import procesar_excel
from myapp.services.limpieza_service import procesar_excel_dinamico

OUTPUT_PATH = os.path.join(os.getcwd(), 'plantilla_modificada.xlsx')

def rellenar_excel(request):
    data = request.json
    try:
        excel_buffer = procesar_excel(data)
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='plantilla_modificada.xlsx'
        )
    except FileNotFoundError:
        return "El archivo de plantilla de Excel no se encontró.", 404
    except Exception as e:
        return str(e), 500
    
def rellenar_excel_limpieza(request):
    data = request.json
    try:
        excel_buffer = procesar_excel_dinamico(data)
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='limpieza.xlsx'
        )
    except FileNotFoundError:
        return "El archivo de plantilla de Excel no se encontró.", 404
    except Exception as e:
        return str(e), 500