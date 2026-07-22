import os
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from constants import (
    DOWNLOAD_DIR,
    SELENIUM_TIMEOUT,
    DOWNLOAD_TIMEOUT
)


# =====================
# ESPERAR DESCARGA REAL
# =====================
def wait_for_download(download_dir, before_files, timeout=DOWNLOAD_TIMEOUT):
    start = time.time()
    last_size = {}

    while time.time() - start < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - before_files

        for filename in new_files:
            if filename.endswith(".crdownload"):
                continue

            path = os.path.join(download_dir, filename)

            if not os.path.exists(path):
                continue

            size = os.path.getsize(path)

            # tamaño estable → descarga completa
            if filename in last_size and last_size[filename] == size:
                return filename

            last_size[filename] = size

        time.sleep(1)

    raise TimeoutError("No se detectó una descarga válida")


# =====================
# ESPERAR TABLA INFORMES
# =====================
def wait_for_reports_table(driver, timeout=SELENIUM_TIMEOUT):
    wait = WebDriverWait(driver, timeout)

    def table_ready(d):
        try:
            tbody = d.find_element(By.XPATH, "//table/tbody")
            rows = tbody.find_elements(By.XPATH, "./tr")
            return len(rows) > 0
        except Exception:
            return False

    wait.until(table_ready)


# =====================
# GENERAR INFORME (BLINDADO)
# =====================
def generar_informe(driver, tratamiento, finalidad, nombre, descripcion):
    wait = WebDriverWait(driver, SELENIUM_TIMEOUT)

    # =====================
    # CONTEXTO INICIAL
    # =====================
    original_window = driver.current_window_handle
    initial_windows = set(driver.window_handles)
    before_files = set(os.listdir(DOWNLOAD_DIR))

    try:
        # =====================
        # ABRIR GENERAR INFORME
        # =====================
        generar_button = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//form//button[contains(text(),'Generar informe')]"
            ))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            generar_button
        )
        generar_button.click()

        # =====================
        # SELECCIONAR EXCEL (ABRE PESTAÑA)
        # =====================
        excel_option = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div/form/div[1]/div/ul/li[3]/a"
            ))
        )
        excel_option.click()

        # =====================
        # ESPERAR NUEVA PESTAÑA
        # =====================
        wait.until(lambda d: len(d.window_handles) > len(initial_windows))
        new_windows = set(driver.window_handles) - initial_windows

        if len(new_windows) != 1:
            raise RuntimeError(f"Número inesperado de pestañas: {len(new_windows)}")

        report_window = new_windows.pop()
        driver.switch_to.window(report_window)

        # =====================
        # VALIDAR CONTEXTO
        # =====================
        wait.until(EC.url_contains("/ereport/"))
        wait_for_reports_table(driver)

        # =====================
        # DESCARGAR INFORME
        # =====================
        download_link = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//table/tbody/tr[1]/td[1]//a"
            ))
        )

        href = download_link.get_attribute("href")
        if not href:
            raise RuntimeError("Link de descarga sin href")

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            download_link
        )
        download_link.click()

        # =====================
        # ESPERAR DESCARGA
        # =====================
        filename = wait_for_download(DOWNLOAD_DIR, before_files)

        return filename

    finally:
        # =====================
        # LIMPIEZA TOTAL
        # =====================
        for handle in driver.window_handles:
            if handle != original_window:
                driver.switch_to.window(handle)
                driver.close()

        driver.switch_to.window(original_window)

        assert driver.current_window_handle == original_window
