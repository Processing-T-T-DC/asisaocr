from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException


# ==============================
# CLICK SEGURO
# ==============================
def safe_click(driver, wait, element):
    try:
        wait.until(EC.element_to_be_clickable(element))
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)


# ==============================
# GENERAR Y DESCARGAR INFORME
# ==============================
def generar_informe(driver):
    wait = WebDriverWait(driver, 40)

    # =====================
    # 1️⃣ ABRIR DROPDOWN "GENERAR INFORME"
    # =====================
    generar_button = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//form//button[contains(@class,'dropdown-toggle')]"
        ))
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        generar_button
    )

    safe_click(driver, wait, generar_button)
    print("✅ Dropdown de generar informe abierto")

    # =====================
    # 2️⃣ SELECCIONAR EXCEL
    # =====================
    excel_option = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//form//ul//a[contains(., 'EXCEL EXTENDIDO')]"
        ))
    )

    safe_click(driver, wait, excel_option)
    print("✅ Formato Excel seleccionado")

    # =====================
    # 3️⃣ IR A LA PÁGINA DE INFORMES
    # =====================
    # driver.get("https://eprivacy.ecix.tech/ereport/?type=14&ancla=14")

    # =====================
    # 4️⃣ ESPERAR A QUE EL INFORME APAREZCA
    # =====================
    try:
        informe_link = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//table//tbody//tr[1]//td[1]//a"
            ))
        )
    except TimeoutException:
        raise Exception("El informe no se generó en el tiempo esperado")

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        informe_link
    )

    safe_click(driver, wait, informe_link)
    print("✅ Informe descargado correctamente")