from flask_login import UserMixin
from . import login_manager

USERS_DB = {
    "1": {"username": "admin", "password": "password", "role": "administrador"},
    "2": {"username": "gestor.aprovador", "password": "password", "role": "gestor"}
}

class User(UserMixin):
    def __init__(self, user_id, username, role):
        self.id = user_id
        self.username = username
        self.role = role

    @staticmethod
    def get(user_id):
        user_data = USERS_DB.get(user_id)
        if user_data:
            return User(user_id=user_id, username=user_data['username'], role=user_data['role'])
        return None

    @staticmethod
    def find_by_username(username):
        for user_id, user_data in USERS_DB.items():
            if user_data['username'] == username:
                return User(user_id=user_id, username=user_data['username'], role=user_data['role'])
        return None

@login_manager.user_loader
def load_user(user_id):
    return User.get(user_id)
