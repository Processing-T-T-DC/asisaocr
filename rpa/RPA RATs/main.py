import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from constants import DOWNLOAD_DIR
from login import login
from etreatment_runner import procesar_etreatment

# =====================
# IMPORT EXCEL LOGGER
# =====================
from excel_runner import (
    cargar_excel,
    escribir_resultado,
    guardar_excel
)

# --------------------------------
# CONFIGURACIÓN GENERAL
# --------------------------------
sys.stdout.reconfigure(encoding="utf-8")

ETREATMENT_URL = "https://eprivacy.ecix.tech/etreatment"

# --------------------------------
# XPATHs NIVEL 2 (CLICK EXPANSIÓN)
# --------------------------------
XPATH_NIVEL_2_CENTROS_MEDICOS = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_CLINICAS = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[3]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_DELEGACIONES = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[4]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_FUNDACIONES = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[5]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_HOSPITALES = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[6]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_NEGOCIO_ASEGURADOR = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[7]/div[1]/div[1]/span/span[1]"
XPATH_NIVEL_2_OTRAS_SOCIEDADES = "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/div/div/div/div/div/div[2]/div/div/div/div[8]/div[1]/div[1]/span/span[1]"

# --------------------------------
# XPATHs BASE RATs POR NIVEL 2
# --------------------------------
XPATH_RATS_CENTROS_MEDICOS = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[2]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_CLINICAS = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[3]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_DELEGACIONES = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[4]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_FUNDACIONES = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[5]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_HOSPITALES = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[6]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_NEGOCIO_ASEGURADOR = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[7]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

XPATH_RATS_OTRAS_SOCIEDADES = (
    "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[4]/"
    "div/div/div/div/div/div[2]/div/div/div/div[8]/div[2]/"
    "div/div/div/div[{i}]/div[1]/div[3]/a"
)

# =====================
# INIT DRIVER
# =====================
def init_driver():
    options = Options()
    options.add_argument("--start-maximized")

    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
    )

    return webdriver.Chrome(options=options)

# =====================
# MAIN
# =====================
def main():

    # =====================
    # INICIALIZAR EXCEL
    # =====================
    wb, ws = cargar_excel()

    # =====================
    # INICIALIZAR SELENIUM
    # =====================
    driver = init_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # =====================
        # LOGIN
        # =====================
        login(driver)

        # =====================
        # IR A ETREATMENT
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        # =====================
        # NIVEL 2 → CENTROS MÉDICOS
        # =====================
        print("🧪 Procesando NIVEL 2: CENTROS MÉDICOS")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_CENTROS_MEDICOS,
            xpath_base_rats=XPATH_RATS_CENTROS_MEDICOS,
            indices=range(2, 36),
            # indices=range(2, 5),
            ws=ws,
            seccion="CENTROS MÉDICOS",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → CLÍNICAS
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: CLÍNICAS")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_CLINICAS,
            xpath_base_rats=XPATH_RATS_CLINICAS,
            indices=range(2, 16),
            # indices=range(2, 5),
            ws=ws,
            seccion="CLÍNICAS",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → DELEGACIONES
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: DELEGACIONES")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_DELEGACIONES,
            xpath_base_rats=XPATH_RATS_DELEGACIONES,
            indices=range(4, 5),
            ws=ws,
            seccion="DELEGACIONES",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → FUNDACIONES
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: FUNDACIONES")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_FUNDACIONES,
            xpath_base_rats=XPATH_RATS_FUNDACIONES,
            indices=range(2, 4),
            ws=ws,
            seccion="FUNDACIONES",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → HOSPITALES
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: HOSPITALES")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_HOSPITALES,
            xpath_base_rats=XPATH_RATS_HOSPITALES,
            indices=range(2, 10),
            ws=ws,
            seccion="HOSPITALES",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → NEGOCIO ASEGURADOR
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: NEGOCIO ASEGURADOR")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_NEGOCIO_ASEGURADOR,
            xpath_base_rats=XPATH_RATS_NEGOCIO_ASEGURADOR,
            indices=range(2, 6),
            ws=ws,
            seccion="NEGOCIO ASEGURADOR",
            expandir_nivel_1=True
        )

        # =====================
        # NIVEL 2 → OTRAS SOCIEDADES
        # =====================
        print("➡️ Accediendo a eTreatment")
        driver.get(ETREATMENT_URL)

        wait.until(EC.presence_of_element_located((By.ID, "etreatment")))
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
            )
        )

        print("🧪 Procesando NIVEL 2: OTRAS SOCIEDADES")
        procesar_etreatment(
            driver=driver,
            xpath_nivel_2=XPATH_NIVEL_2_OTRAS_SOCIEDADES,
            xpath_base_rats=XPATH_RATS_OTRAS_SOCIEDADES,
            indices=range(2, 20),
            # indices=range(2, 5),
            ws=ws,
            seccion="OTRAS SOCIEDADES",
            expandir_nivel_1=True
        )

    except Exception as e:
        error_msg = str(e)
        print(f"❌ Error en ejecución principal: {error_msg}")

        escribir_resultado(
            ws=ws,
            estado="ERROR",
            fichero=None,
            error=error_msg
        )

    finally:
        guardar_excel(wb)
        driver.quit()
        print("🧹 Navegador cerrado")

# =====================
# ENTRY POINT
# =====================
if __name__ == "__main__":
    main()
