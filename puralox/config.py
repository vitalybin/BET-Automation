import os

# Configuration constants
UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
DB_NAME = os.getenv('DB_NAME', 'Purlox.db')

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
