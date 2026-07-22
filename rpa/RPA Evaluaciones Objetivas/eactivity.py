from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By


def wait_for_table_reload(driver, previous_row_count=None, timeout=30):
    wait = WebDriverWait(driver, timeout)

    def table_ready(d):
        try:
            tbody = d.find_element(By.XPATH, "//table/tbody")
            rows = tbody.find_elements(By.XPATH, "./tr")
            text = tbody.text.strip()

            if "No se encontraron" in text:
                return True

            if previous_row_count is None:
                return len(rows) > 0

            return len(rows) != previous_row_count
        except Exception:
            return False

    wait.until(table_ready)


def go_to_eactivity_and_act(driver, tratamiento, finalidad, nombre, descripcion):
    wait = WebDriverWait(driver, 30)

    # =====================
    # IR A LA PÁGINA
    # =====================
    driver.get("https://eprivacy.ecix.tech/eactivity/global")
    wait.until(lambda d: d.find_element(By.ID, "kt_app_content"))

    # =====================
    # ESPERAR CARGA INICIAL TABLA
    # =====================
    wait_for_table_reload(driver)

    # Guardamos estado inicial
    tbody = driver.find_element(By.XPATH, "//table/tbody")
    initial_row_count = len(tbody.find_elements(By.XPATH, "./tr"))

    # =====================
    # ABRIR FILTROS
    # =====================
    driver.find_element(By.ID, "btn-filter-flotable").click()
    wait.until(lambda d: d.find_element(By.ID, "menu-filter").is_displayed())

    # =====================
    # RELLENAR FILTROS
    # =====================
    driver.find_element(By.ID, "treatment_filter").clear()
    driver.find_element(By.ID, "treatment_filter").send_keys(tratamiento)

    driver.find_element(By.ID, "purpose_filter").clear()
    driver.find_element(By.ID, "purpose_filter").send_keys(finalidad)

    driver.find_element(By.ID, "name_filter").clear()
    driver.find_element(By.ID, "name_filter").send_keys(nombre)

    driver.find_element(By.ID, "description_filter").clear()
    driver.find_element(By.ID, "description_filter").send_keys(descripcion)

    # =====================
    # APLICAR FILTROS
    # =====================
    driver.find_element(By.ID, "id_aplicar_filtros").click()

    # =====================
    # ESPERAR RECARGA TABLA
    # =====================
    wait_for_table_reload(driver, previous_row_count=initial_row_count)

    # =====================
    # VALIDAR RESULTADOS
    # =====================
    tbody = driver.find_element(By.XPATH, "//table/tbody")
    tbody_text = tbody.text

    if "No se encontraron" in tbody_text:
        raise Exception("NO ENCONTRADO")

    if nombre and nombre not in tbody_text:
        raise Exception("RESULTADO NO COINCIDE CON FILTROS")

    # =====================
    # CERRAR FILTROS
    # =====================
    driver.find_element(By.ID, "kt_filter_close").click()
    wait.until(lambda d: not d.find_element(By.ID, "menu-filter").is_displayed())

    # =====================
    # ACCIONES SOBRE LA TABLA
    # =====================
    actions_button = wait.until(
        lambda d: d.find_element(
            By.XPATH,
            "//table/tbody/tr[1]//button[contains(@class,'dropdown-toggle')]"
        )
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        actions_button
    )
    actions_button.click()

    menu_option = wait.until(
        lambda d: d.find_element(
            By.XPATH,
            "//table/tbody/tr[1]/td[9]//ul/li[2]/a"
        )
    )
    menu_option.click()
