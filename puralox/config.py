import os
from dotenv import load_dotenv

load_dotenv()
# Configuration constants
UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
DB_NAME = os.getenv('DB_NAME', 'Purlox.db')
ELABFTW_URL = os.getenv("ELABFTW_URL", "https://localhost/api/v2")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "6-58fefeb5b740b8334164e94dbac6faf1f52f07f10f911f43b2fa2e5f376b1e06b38be33679fbeb3d0cd66")
# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
