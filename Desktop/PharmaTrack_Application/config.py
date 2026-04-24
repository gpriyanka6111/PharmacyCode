# App-level configuration: UPLOAD_FOLDER, PROCESSED_FOLDER, display name, and other constants.
import os

APP_DISPLAY_NAME = "RxInsight"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'processed')
