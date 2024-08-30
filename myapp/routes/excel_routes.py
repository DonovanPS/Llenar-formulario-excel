from flask import Blueprint, request, send_file
from myapp.controllers.excel_controller import rellenar_excel

excel_blueprint = Blueprint('excel', __name__)

@excel_blueprint.route('/rellenar_excel', methods=['POST'])
def rellenar_excel_route():
    return rellenar_excel(request)
