import os
from flask import send_file
from app.services.excel_service import procesar_excel

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
