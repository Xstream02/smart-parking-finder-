import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'A_VERY_SECRET_KEY_FOR_PARKING_APP')
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(BASE_DIR, 'database.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
