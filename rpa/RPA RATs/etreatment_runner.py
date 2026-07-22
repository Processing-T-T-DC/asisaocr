import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException
)

from generar_informe import generar_informe
from excel_runner import escribir_resultado

ETREATMENT_URL = "https://eprivacy.ecix.tech/etreatment"


# ==============================
# CLICK REAL SEGURO
# ==============================
def safe_click(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        element
    )
    driver.execute_script("arguments[0].click();", element)


# ==============================
# EXPANDIR CUALQUIER NIVEL
# ==============================
def expandir_por_xpath(driver, wait, xpath, descripcion="nivel"):
    print(f"➡️ Expandiendo {descripcion}")

    for intento in range(1, 4):
        try:
            elemento = wait.until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                elemento
            )
            time.sleep(1)

            wait.until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )

            driver.execute_script("arguments[0].click();", elemento)

            wait.until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, ".blockUI, .overlay, .modal-backdrop")
                )
            )

            time.sleep(2)
            return

        except (ElementClickInterceptedException, StaleElementReferenceException):
            print(f"⚠️ Reintento {intento}/3 al expandir {descripcion}")
            time.sleep(2)

    raise RuntimeError(f"No se pudo expandir {descripcion}")


# ==============================
# PROCESAR RATs + EXCEL
# ==============================
def procesar_rats_por_indices(
    driver,
    wait,
    indices,
    xpath_base,
    seccion,
    ws
):
    ventana_principal = driver.current_window_handle

    for i in indices:
        xpath = xpath_base.format(i=i)

        try:
            print(f"➡️ Procesando RAT índice {i}")

            link = wait.until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                link
            )
            time.sleep(0.5)

            # abrir en nueva pestaña
            link.send_keys(Keys.CONTROL + Keys.RETURN)
            time.sleep(2)

            nuevas_ventanas = [
                h for h in driver.window_handles
                if h != ventana_principal
            ]

            if not nuevas_ventanas:
                raise RuntimeError("No se abrió nueva pestaña")

            driver.switch_to.window(nuevas_ventanas[0])

            print(f"⬇ Generando informe RAT índice {i}")
            fichero_generado = generar_informe(driver)

            escribir_resultado(
                ws=ws,
                estado="OK",
                fichero=seccion,
                posicion=i,
                error=None
            )

            time.sleep(2)

            driver.close()
            driver.switch_to.window(ventana_principal)
            time.sleep(1)

        except Exception as e:
            error_msg = str(e)
            print(f"❌ Error en RAT índice {i}: {error_msg}")

            escribir_resultado(
                ws=ws,
                estado="ERROR",
                fichero=None,
                error=f"RAT índice {i}: {error_msg}"
            )

            driver.switch_to.window(ventana_principal)
            continue


# ==============================
# FUNCIÓN PRINCIPAL
# ==============================
def procesar_etreatment(
    driver,
    xpath_nivel_2,
    xpath_base_rats,
    indices,
    ws,
    seccion,
    expandir_nivel_1=True
):
    wait = WebDriverWait(driver, 30)

    # ==============================
    # 1️⃣ Expandir NIVEL 1
    # ==============================
    if expandir_nivel_1:
        expandir_por_xpath(
            driver,
            wait,
            "//div[@id='group-category-3110']/div/span/span",
            descripcion="NIVEL 1 (GRUPO ASISA)"
        )

    # ==============================
    # 2️⃣ Expandir NIVEL 2
    # ==============================
    expandir_por_xpath(
        driver,
        wait,
        xpath_nivel_2,
        descripcion="NIVEL 2"
    )

    # ==============================
    # 3️⃣ Procesar RATs
    # ==============================
    print("⬇ Descargando informes RAT")

    procesar_rats_por_indices(
        driver=driver,
        wait=wait,
        indices=indices,
        xpath_base=xpath_base_rats,
        seccion=seccion,
        ws=ws
    )

    print("✅ Nivel 2 procesado")
