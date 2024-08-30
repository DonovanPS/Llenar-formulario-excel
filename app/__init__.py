from flask import Flask
from flask_cors import CORS

def create_app():
    app = Flask(__name__)
    CORS(app)

    # Registrar Blueprints
    from app.routes.excel_routes import excel_blueprint
    app.register_blueprint(excel_blueprint)

    return app
