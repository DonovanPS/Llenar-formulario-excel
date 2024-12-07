from flask import Blueprint, request
from myapp.controllers.excel_controller import rellenar_excel, rellenar_excel_limpieza, rellenar_excel_salud

excel_blueprint = Blueprint('excel', __name__)

@excel_blueprint.route('/rellenar_excel', methods=['POST'])
def rellenar_excel_route():
    return rellenar_excel(request)

@excel_blueprint.route('/rellenar_excel_limpieza', methods=['POST'])
def rellenar_excel_limpieza_route():
    return rellenar_excel_limpieza(request)

@excel_blueprint.route('/rellenar_excel_salud', methods=['POST'])
def rellenar_excel_salud_route():
    return rellenar_excel_salud(request)