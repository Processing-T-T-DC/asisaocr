import os

# =====================
# PATHS
# =====================
BASE_DIR = r"C:\Users\jbertrandelis\OneDrive - Deloitte (O365D)\Desktop\Proyectos\18.Asisa - OCR\RPA\RPA RATs"

DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
EXCEL_INPUT = os.path.join(BASE_DIR, "informe_riesgo.xlsx")
EXCEL_OUTPUT = os.path.join(BASE_DIR, "informe_riesgo_resultado.xlsx")

# Asegurar que existe la carpeta de descargas
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# =====================
# EXCEL
# =====================
SHEET_NAME = "Sheet1"


# =====================
# SELENIUM
# =====================
SELENIUM_TIMEOUT = 30
DOWNLOAD_TIMEOUT = 120


# =====================
# URLs
# =====================
URL_REPORTS = "https://eprivacy.ecix.tech/ereport/?type=14&ancla=14"
