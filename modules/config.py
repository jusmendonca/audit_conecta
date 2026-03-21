import os
from dotenv import load_dotenv

load_dotenv()

BASE_URL: str = os.environ.get("SUPP_BASE_URL", "").rstrip("/")
