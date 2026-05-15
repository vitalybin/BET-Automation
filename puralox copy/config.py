import os
from dotenv import load_dotenv

load_dotenv()

# Puralox paths
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "/usr/src/app/uploads")
DB_NAME = os.getenv("DB_NAME", "/usr/src/app/data/Purlox.db")

# External eLabFTW API configuration
ELABFTW_URL = os.getenv("ELABFTW_URL", "https://dtpa-akg.de/api/v2")
ELABFTW_TOKEN = os.getenv("ELABFTW_TOKEN", "")

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Ensure DB directory exists if DB_NAME points to a nested path
_db_dir = os.path.dirname(DB_NAME)
if _db_dir:
    os.makedirs(_db_dir, exist_ok=True)
