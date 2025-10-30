from flask import Flask
from flask_login import LoginManager

login_manager = LoginManager()

def create_app():
    """Cria e configura uma instância da aplicação Flask."""
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'uma-chave-secreta-muito-segura'

    login_manager.init_app(app)
    login_manager.login_view = 'login'
    login_manager.login_message = "Por favor, faça o login para acessar esta página."
    login_manager.login_message_category = "info"

    with app.app_context():
        # Importa e registra as rotas
        from . import routes
        return app